require 'date'
require 'fileutils'
require 'rubyXL'
require 'rubyXL/convenience_methods'


# Calendar と RoomInfo から学期別の講義室予約表 Excel を生成するクラス
# @calendar: Calendar オブジェクト
# @room_info: RoomInfo オブジェクト
# @events: イベントの配列
# @lectures: 講義の配列
class XlsxGenerator
	TERM_ORDER = ["1学期", "2学期", "3学期", "4学期", "夏季休業", "春季休業"].freeze
	PERIOD_TO_COL = {
		1 => 2,
		2 => 3,
		3 => 4,
		4 => 5,
		:lunch => 6,
		5 => 7,
		6 => 8,
		7 => 9,
		8 => 10
	}.freeze
	COL_TO_PERIOD = PERIOD_TO_COL.invert.freeze
	EN_WEEK_TO_JA = {
		'mon' => '月',
		'tue' => '火',
		'wed' => '水',
		'thu' => '木',
		'fri' => '金',
		'sat' => '土',
		'sun' => '日'
	}.freeze
	LECTURE_ALLOWED_DAY_TYPES = ['授業日', '試験日', '補講日'].freeze
	TIME_LABELS = [
		'8:40～9:30',
		'9:40～10:30',
		'10:45～11:35',
		'11:45～12:35',
		'12:35～13:25',
		'13:25～14:15',
		'14:25～15:15',
		'15:30～16:20',
		'16:30～17:20'
	].freeze
	DEFAULT_ROW_HEIGHT = 34.0
	HEADER_ROW_HEIGHT = 17.0
	HEADER_NOTE_TEXT = '※重複している予定は赤色で表示されます'.freeze

	# 初期化
	# @param calendar [Calendar] カレンダーオブジェクト
	# @param room_info [RoomInfo] 講義室情報オブジェクト
	# @param events [Array<Event>] イベントの配列
	# @param lectures [Array<Lecture>] 講義の配列
	# @return [void]
	def initialize(calendar:, room_info:, events: [], lectures: [])
		@calendar = calendar
		@room_info = room_info
		@events = Array(events)
		@lectures = Array(lectures)

	end

	# 出力先ディレクトリを作成し，generate_one_term を用いて全学期分の XLSX ファイルを作成する．
	# 予定の重複内容を print_conflict で標準出力に表示し，作成したファイルのパスを返す
	# @param output_dir [String, nil] 作成する講義室予約表 XLSX ファイルを配置するディレクトリのパス
	# @return [Array<String>] 作成した講義室予約表 XLSX ファイルのパスのリスト
	def generate_all(output_dir: nil)
		validate_inputs!

		# (1) Calenderインスタンスから年度を取得し，出力先ディレクトリを決定する．
		year = @calendar.year || infer_year_from_dates
		output_dir ||= "#{year}年度_講義室予約表"

		# (2) 作成する講義室予約表XLSXファイルを配置するディレクトリのパスを確認し，出力先ディレクトリがなければ作成する
		FileUtils.mkdir_p(output_dir)

		generated_files = []
		all_conflicts = []
		TERM_ORDER.each do |term|

			# (3) 1学期，2学期，3学期，4学期，夏季，春季の順に，各学期の日付範囲を取得する
			range = term_range(term)
			next if range.nil?

			# (4) 日付範囲が取得できた学期について，generate_one_term を呼び出し，学期ごとの講義室予約表 XLSX ファイルを作成する
			output_path = File.join(output_dir, "#{year}年度_#{term_file_name(term)}_講義室予約表.xlsx")

			# (5) 作成したXLSXファイルのパスと，学期ごとの予定の重複情報をリストに追加する
			term_conflicts = generate_one_term(term: term, date_range: range, output_path: output_path, year: year)
			generated_files << output_path
			all_conflicts.concat(term_conflicts)
		end

		# (6) print_conflicts を呼び出し，予定の重複情報を標準出力に表示し，作成した講義室予約表の XLSX ファイルのパスのリストを返す
		print_conflicts(all_conflicts)
		generated_files
	end

	private

	# 入力の妥当性を検証する
	# @return [void]
	# @raise [ArgumentError] Calendar または RoomInfo が未設定の場合に発生する
	def validate_inputs!
		raise ArgumentError, 'Calendar が未設定です' if @calendar.nil?
		raise ArgumentError, 'RoomInfo が未設定です' if @room_info.nil?
	end

	# Calendar の日付情報から年度を推測する
	# @return [Integer] 推測された年度
	# @raise [ArgumentError] Calendar に日付情報がない場合に発生する
	def infer_year_from_dates
		first_date = @calendar.date_list.keys.min
		raise ArgumentError, 'Calendar に日付情報がありません' if first_date.nil?

		first_date.year
	end

	# 指定された学期名に対応する日付範囲を取得する
	# @param term [String] 学期名
	# @return [Range<Date>, nil] 学期の期間を表す日付範囲。学期が存在しない場合は nil を返す
	def term_range(term)
		case term
		when '夏季休業'
			@calendar.term_periods['夏季休業']
		when '春季休業'
			@calendar.term_periods['春季休業']
		else
			@calendar.term_periods[term]
		end
	end

	# 学期名をファイル名に適した形式に変換する
	# @param term [String] クラスが保持する学期名
	# @return [String] ファイル名に適した学期名
	def term_file_name(term)
		case term
		when '夏季休業'
			'夏季'
		when '春季休業'
			'春季'
		else
			term
		end
	end

	# 1学期 / 2学期 / 3学期 / 4学期 / 夏季 / 春季 のいずれかの学期について、講義室予約表を生成する
	# @param term [String] 学期名
	# @param date_range [Range<Date>] 生成対象の日付範囲
	# @param output_path [String] 生成する Excel ファイルのパス
	# @param year [Integer] 年度
	# @return [Array<Hash>] 予定の重複情報の配列
	def generate_one_term(term:, date_range:, output_path:, year:)
		workbook = RubyXL::Workbook.new
		sheet = workbook[0]
		setup_header(sheet, year, term)

		styles = default_styles
		cell_entries, conflicts = build_term_cell_entries(term, date_range)
		fill_term_rows(sheet, date_range, styles, cell_entries)

		workbook.write(output_path)
		conflicts
	end

	# 講義室予約表のヘッダ部分を設定する
	# @param sheet [RubyXL::Worksheet] 対象のシート
	# @param year [Integer] 年度
	# @param term [String] 学期名
	# @return [void]
	def setup_header(sheet, year, term)
		term_title = term_title_name(term)
		sheet.sheet_name = "#{year}年度#{term_title}"
		set_default_row_height(sheet)
		set_column_widths(sheet)

		write_cell(sheet, 0, 1, "#{year}年度講義室等使用状況表（#{term_title}）", nil)
		sheet.merge_cells(0, 1, 0, 10)
		sheet[0][1].change_horizontal_alignment('center')
		sheet[0][1].change_vertical_alignment('center')
		sheet[0][1].change_font_bold(true)
		sheet[0][1].change_font_size(14)
		write_cell(sheet, 1, 11, HEADER_NOTE_TEXT, nil)
		sheet.merge_cells(1, 11, 1, 12)
		sheet[1][11].change_vertical_alignment('center')
		sheet[1][11].change_horizontal_alignment('left')
		sheet[1][11].change_font_color('FF0000')

		(1..10).each do |col|
			write_cell(sheet, 1, col, nil, nil)
		end

		write_cell(sheet, 2, 1, '時限', nil)
		[1, 2, 3, 4, '昼休み', 5, 6, 7, 8].each_with_index do |label, index|
			write_cell(sheet, 2, 2 + index, label, nil)
		end

		write_cell(sheet, 3, 1, '時間', nil)
		TIME_LABELS.each_with_index do |label, index|
			write_cell(sheet, 3, 2 + index, label, nil)
		end

		(1..10).each do |col|
			[1, 2, 3].each do |row|
				sheet[row][col].change_horizontal_alignment('center')
				sheet[row][col].change_vertical_alignment('center')
			end
		end

		(1..10).each do |col|
			write_cell(sheet, 2, col, sheet[2][col]&.value, nil)
			write_cell(sheet, 3, col, sheet[3][col]&.value, nil)
			write_cell(sheet, 1, col, sheet[1][col]&.value, nil)
			sheet[1][col].change_border('left', 'thin')
			sheet[1][col].change_border('right', 'thin')
			sheet[1][col].change_border('top', 'thin')
			sheet[1][col].change_border('bottom', 'thin')
			sheet[2][col].change_border('left', 'thin')
			sheet[2][col].change_border('right', 'thin')
			sheet[2][col].change_border('top', 'thin')
			sheet[2][col].change_border('bottom', 'thin')
			sheet[3][col].change_border('left', 'thin')
			sheet[3][col].change_border('right', 'thin')
			sheet[3][col].change_border('top', 'thin')
			sheet[3][col].change_border('bottom', 'thin')
		end

		set_header_row_heights(sheet)

		freeze_top_four_rows(sheet)
	end

	# 学期名をシートタイトル用の表記に変換する
	# @param term [String] 学期名
	# @return [String] シートタイトル用の学期名
	def term_title_name(term)
		case term
		when '夏季休業'
			'夏季休業期間'
		when '春季休業'
			'春季休業期間'
		else
			term
		end
	end

	# 列幅を設定する
	# @param sheet [RubyXL::Worksheet] 対象のシート
	# @return [void]
	def set_column_widths(sheet)
		sheet.change_column_width(0, 14)
		sheet.change_column_width(1, 18)
		(2..10).each { |col| sheet.change_column_width(col, 17) }
		sheet.change_column_width(11, 28)
		sheet.change_column_width(12, 4)
	end

	# シートのデフォルト行高を設定する
	# @param sheet [RubyXL::Worksheet] 対象のシート
	# @return [void]
	def set_default_row_height(sheet)
		sheet.sheet_format_pr ||= RubyXL::WorksheetFormatProperties.new
		sheet.sheet_format_pr.default_row_height = DEFAULT_ROW_HEIGHT
		sheet.sheet_format_pr.custom_height = true
	end

	# ヘッダの2〜4行目の行高を設定する
	# @param sheet [RubyXL::Worksheet] 対象のシート
	# @return [void]
	def set_header_row_heights(sheet)
		(1..3).each do |row_index|
			row = sheet.sheet_data[row_index]
			next if row.nil?

			row.ht = HEADER_ROW_HEIGHT
			row.custom_height = true
		end
	end

	# シートの上部4行を固定する
	# @param sheet [RubyXL::Worksheet] 対象のシート
	# @return [void]
	def freeze_top_four_rows(sheet)
		sheet.sheet_views ||= RubyXL::WorksheetViews.new
		sheet.sheet_views << RubyXL::WorksheetView.new if sheet.sheet_views.empty?

		view = sheet.sheet_views[0]
		view.pane ||= RubyXL::Pane.new

		view.pane.state = 'frozen'
		view.pane.y_split = 4
		view.pane.top_left_cell = 'A5'
	end

	# 講義室予約表のセルに設定するデフォルトのスタイルを返す
	# @return [Hash] スタイル情報のハッシュ
	def default_styles
		{
			left_top: nil,
			left_other: nil,
			room_mesh: Array.new(10)
		}
	end

	# 学期の期間内の各日付について、講義室予約表の行を埋める
	# @param sheet [RubyXL::Worksheet] 対象のシート
	# @param date_range [Range<Date>] 学期の期間内の日付範囲
	# @param styles [Hash] セルのスタイル情報
	# @param cell_entries [Hash] 日付・講義室・時限ごとのセル情報のハッシュ
	# @return [void]
	def fill_term_rows(sheet, date_range, styles, cell_entries)
		rooms = Array(@room_info.output_target_rooms)
		current_row = 4

		date_range.each do |date|
			date_info = @calendar.date_list[date]
			weekday = resolve_weekday(date, date_info)
			remarks = extract_remarks(date_info)

			left_height = 2 + remarks.length
			table_height = [rooms.length, 1].max
			block_height = [left_height, table_height].max

			write_left_block(
				sheet: sheet,
				base_row: current_row,
				date: date,
				weekday: weekday,
				remarks: remarks,
				height: block_height,
				styles: styles
			)

			write_room_table(
				sheet: sheet,
				base_row: current_row,
				date: date,
				rooms: rooms,
				height: block_height,
				styles: styles,
				cell_entries: cell_entries
			)

			current_row += block_height + 1
		end
	end

	# 左側のブロック（日時・曜日・備考）をシートに書き込む
	# @param sheet [RubyXL::Worksheet] 対象のシート
	# @param base_row [Integer] 書き込み開始行
	# @param date [Date] 日付
	# @param weekday [String] 曜日
	# @param remarks [Array<String>] 備考の配列
	# @param height [Integer] ブロックの高さ
	# @param styles [Hash] セルのスタイル情報
	# @return [void]
	def write_left_block(sheet:, base_row:, date:, weekday:, remarks:, height:, styles:)
		(0...height).each do |offset|
			row_index = base_row + offset
			value = if offset.zero?
								"(#{fullwidth_month_day(date)})"
							elsif offset == 1
								weekday
							else
								remarks[offset - 2]
							end

			style_index = offset.zero? ? styles[:left_top] : styles[:left_other]
			write_cell(sheet, row_index, 0, value, style_index)
			sheet[row_index][0].change_vertical_alignment('center')
			apply_weekday_color(sheet[row_index][0], date) if offset == 1
		end
	end

	# 講義室の表をシートに書き込む
	# @param sheet [RubyXL::Worksheet] 対象のシート
	# @param base_row [Integer] 書き込み開始行
	# @param date [Date] 日付
	# @param rooms [Array<String>] 講義室の配列
	# @param height [Integer] ブロックの高さ
	# @param styles [Hash] セルのスタイル情報
	# @param cell_entries [Hash] 日付・講義室・時限ごとのセル情報のハッシュ
	# @return [void]
	def write_room_table(sheet:, base_row:, date:, rooms:, height:, styles:, cell_entries:)
		(0...height).each do |offset|
			row_index = base_row + offset
			row_styles = styles[:room_mesh]
			(1..10).each do |col_index|
				value = nil
				entries = []
				if col_index == 1 && offset < rooms.length
					value = rooms[offset]
				elsif offset < rooms.length
					period = COL_TO_PERIOD[col_index]
					unless period.nil?
						cell_key = [date, normalize_room_name(rooms[offset]), period]
						entries = cell_entries[cell_key] || []
						value = entries.map { |entry| entry[:cell_label] }.uniq.join('，')
					end
				end
				write_cell(sheet, row_index, col_index, value, row_styles[col_index - 1])
				sheet[row_index][col_index].change_vertical_alignment('center')
				sheet[row_index][col_index].change_text_wrap(true)

				if col_index >= 2 && entries.size > 1
					sheet[row_index][col_index].change_font_color('FF0000')
				end
			end
		end

		apply_room_table_mesh_border(sheet: sheet, base_row: base_row, height: height, start_col: 1, end_col: 10)
	end

	# 講義室の表のメッシュ境界を適用する
	# @param sheet [RubyXL::Worksheet] 対象のシート
	# @param base_row [Integer] 書き込み開始行
	# @param height [Integer] ブロックの高さ
	# @param start_col [Integer] 開始列
	# @param end_col [Integer] 終了列
	# @return [void]
	def apply_room_table_mesh_border(sheet:, base_row:, height:, start_col:, end_col:)
		end_row = base_row + height - 1

		(base_row..end_row).each do |row|
			(start_col..end_col).each do |col|
				ensure_cell(sheet, row, col)
				cell = sheet[row][col]

				# 初期データに合わせ、全辺を細線でメッシュ化する
				cell.change_border('left', 'thin')
				cell.change_border('right', 'thin')
				cell.change_border('top', 'thin')
				cell.change_border('bottom', 'thin')
			end
		end
	end

	# 学期の期間内の各(日付, 講義室, 時限)ごとのセル情報を構築する
	# @param term [String] 学期名
	# @param date_range [Range<Date>] 学期の期間内の日付範囲
	# @return [Array<Hash>] 予定の重複情報の配列
	def build_term_cell_entries(term, date_range)
		date_set = date_range.each_with_object({}) { |date, hash| hash[date] = true }
		cell_entries = Hash.new { |hash, key| hash[key] = [] }
		target_rooms = normalize_rooms(@room_info.output_target_rooms).each_with_object({}) { |room, hash| hash[room] = true }

		@events.each do |event|
			next unless event.is_a?(Event)
			event_dates(event).each do |event_date|
				next unless date_set[event_date]

				periods_for_event(event.start_time, event.end_time).each do |period|
					normalize_rooms(event.room).each do |room|
						next unless target_rooms[room]

						cell_entries[[event_date, room, period]] << event_entry_payload(event)
					end
				end
			end
		end

		date_range.each do |date|
			date_info = @calendar.date_list[date]
			@lectures.each do |lecture|
				next unless lecture.is_a?(Lecture)
				next unless lecture_matches_term?(lecture, term)
				next unless lecture_matches_date?(lecture, date, date_info)

				periods_for_lecture(lecture.start_time, lecture.end_time).each do |period|
					normalize_rooms(lecture.room).each do |room|
						next unless target_rooms[room]

						cell_entries[[date, room, period]] << lecture_entry_payload(lecture)
					end
				end
			end
		end

		conflicts = build_conflicts(term, cell_entries)
		[cell_entries, conflicts]
	end

	# イベント情報からセル情報のペイロードを生成する
	# @param event [Event] イベントオブジェクト
	# @return [Hash] セル情報のペイロード
	def event_entry_payload(event)
		{
			type: 'event',
			cell_label: "#{event.title} (利用者: #{event.person})",
			report_line: "イベント名: #{event.title} (利用者：#{event.person})"
		}
	end

	# Event の単日/複数日データを日付の配列に正規化する
	# @param event [Event] イベントオブジェクト
	# @return [Array<Date>] 対象日付の配列
	def event_dates(event)
		raw_dates = if event.respond_to?(:data_list) && !blank?(event.data_list)
				Array(event.data_list)
			elsif event.respond_to?(:date_list) && !blank?(event.date_list)
				Array(event.date_list)
			else
				[event.date]
			end

		raw_dates.filter_map { |value| normalize_event_date(value) }.uniq
	end

	# Event の日付候補を Date に変換する
	# @param value [Object] 日付候補
	# @return [Date, nil] 変換結果
	def normalize_event_date(value)
		return value if value.is_a?(Date)
		return value.to_date if value.respond_to?(:to_date)

		case value
		when Hash
			normalize_event_date(value[:date] || value['date'] || value[:day] || value['day'])
		else
			Date.parse(value.to_s)
		end
	rescue ArgumentError, TypeError
		nil
	end

	# 講義情報からセル情報のペイロードを生成する
	# @param lecture [Lecture] 講義オブジェクト
	# @return [Hash] セル情報のペイロード
	def lecture_entry_payload(lecture)
		{
			type: 'lecture',
			cell_label: "#{lecture.title} (担当者: #{lecture.person})",
			report_line: "講義名: #{lecture.title} (担当者：#{lecture.person})"
		}
	end

	# 指定された開始時刻と終了時刻から、イベントが予約する時限の配列を返す
	# @param start_time [String, Numeric] 開始時刻
	# @param end_time [String, Numeric] 終了時刻
	# @return [Array<Integer>] 対応する時限の配列
	def periods_for_event(start_time, end_time)
		start_period = to_period_int(start_time)
		end_period = to_period_int(end_time)
		return [] if start_period.nil? || end_period.nil?
		return [] if end_period < start_period

		periods = (start_period..end_period).select { |period| PERIOD_TO_COL.key?(period) }
		periods << :lunch if start_period <= 4 && end_period >= 5

		periods
	end

	# 指定された開始時刻と終了時刻から、講義が予約する時限の配列を返す
	# @param start_time [String, Numeric] 開始時刻
	# @param end_time [String, Numeric] 終了時刻
	# @return [Array<Integer>] 対応する時限の配列
	def periods_for_lecture(start_time, end_time)
		start_period = to_period_int(start_time)
		end_period = to_period_int(end_time)
		return [] if start_period.nil? || end_period.nil?
		return [] if end_period < start_period

		(start_period..end_period).select { |period| PERIOD_TO_COL.key?(period) }
	end

	# 指定された値を時限の整数に変換する
	# @param value [String, Numeric] 変換対象の値
	# @return [Integer, nil] 変換結果の時限の整数、変換できない場合は nil
	def to_period_int(value)
		return nil if value.nil?
		return value.to_i if value.is_a?(Numeric)

		str = value.to_s.strip
		return nil if str.empty?

		matched = str.match(/\d+/)
		return nil if matched.nil?

		matched[0].to_i
	end

	# 複数の講義室名を正規化する
	# @param rooms [Array<String>] 講義室名の配列
	# @return [Array<String>] 正規化された講義室名の配列
	def normalize_rooms(rooms)
		Array(rooms).map { |room| normalize_room_name(room) }.reject(&:empty?).uniq
	end

	# 講義室名を正規化する
	# @param room_name [String] 講義室名
	# @return [String] 正規化された講義室名
	def normalize_room_name(room_name)
		value = room_name.to_s.strip
		value.each_char.map do |char|
			codepoint = char.ord
			case codepoint
			when 0x20
				'　'
			when 0x21..0x7E
				(codepoint + 0xFEE0).chr(Encoding::UTF_8)
			else
				char
			end
		end.join
	end

	# 講義が指定された学期に一致するかを判定する
	# @param lecture [Lecture] 講義オブジェクト
	# @param term [String] 学期名
	# @return [Boolean] 一致する場合は true、そうでない場合は false
	def lecture_matches_term?(lecture, term)
		return true if lecture.term.nil?

		term_code = case term
					when '1学期' then '1'
					when '2学期' then '2'
					when '3学期' then '3'
					when '4学期' then '4'
					when '夏季' then 'summer'
					when '春季' then 'spring'
					else term.to_s.downcase
					end

		lecture_term_str = lecture.term.to_s.strip.downcase
		return true if lecture_term_str.empty?

		if %w[summer spring].include?(term_code)
			lecture_term_str.include?(term_code)
		else
			lecture_term_str.split(/[^0-9]+/).reject(&:empty?).include?(term_code)
		end
	end

	# 講義が指定された日付に一致するかを判定する
	# @param lecture [Lecture] 講義オブジェクト
	# @param date [Date] 日付
	# @param date_info [DateInfo, nil] 日付情報（nil の場合は Calendar から取得する）
	# @return [Boolean] 一致する場合は true、そうでない場合は false
	def lecture_matches_date?(lecture, date, date_info)
		lecture_week = lecture.week.to_s.strip
		return false if lecture_week.empty?
		return false unless lecture_allowed_on_date_type?(date_info)

		expected_weekday = EN_WEEK_TO_JA[lecture_week.downcase]
		return false if expected_weekday.nil?

		actual_lect_weekday = if !date_info.nil? && !blank?(date_info.lect_weekday)
						date_info.lect_weekday.to_s[0]
					elsif !date_info.nil? && !blank?(date_info.weekday)
						date_info.weekday.to_s[0]
					else
						%w[日 月 火 水 木 金 土][date.wday]
					end

		expected_weekday == actual_lect_weekday
	end

	# 講義を配置可能な講義日種別かを判定する
	# @param date_info [DateInfo, nil] 日付情報
	# @return [Boolean] 配置可能なら true
	def lecture_allowed_on_date_type?(date_info)
		day_types = extract_date_types(date_info)
		return true if day_types.empty?
		return false if day_types.include?('休講日')

		(day_types & LECTURE_ALLOWED_DAY_TYPES).any?
	end

	# 日付情報から講義日種別の配列を抽出する
	# @param date_info [DateInfo, nil] 日付情報
	# @return [Array<String>] 講義日種別の配列
	def extract_date_types(date_info)
		return [] if date_info.nil? || date_info.type.nil?

		Array(date_info.type)
			.flat_map { |day_type| day_type.to_s.split(/[、,\s]+/) }
			.map(&:strip)
			.reject(&:empty?)
			.uniq
	end

	# 学期の期間内の各(日付, 講義室, 時限)ごとのセル情報から、予定の重複情報を構築する
	# @param term [String] 学期名
	# @param cell_entries [Hash] 日付・講義室・時限ごとのセル情報のハッシュ
	# @return [Array<Hash>] 予定の重複情報の配列
	def build_conflicts(term, cell_entries)
		conflicts = []
		grouped = Hash.new { |hash, key| hash[key] = {} }

		cell_entries.each do |(date, room, period), entries|
			next unless entries.length > 1

			grouped[[date, room]][period] = entries
		end

		grouped.each do |(date, room), period_map|
			sorted_periods = period_map.keys.sort
			next if sorted_periods.empty?

			segment_start = nil
			segment_end = nil
			segment_entries = nil

			sorted_periods.each do |period|
				entries = period_map[period]
				entry_signature = entries.map { |entry| entry[:report_line] }.uniq.sort.join('|')

				if segment_start.nil?
					segment_start = period
					segment_end = period
					segment_entries = entry_signature
				elsif period == segment_end + 1 && segment_entries == entry_signature
					segment_end = period
				else
					conflicts << conflict_record(date, term, room, segment_start, segment_end, period_map[segment_start])
					segment_start = period
					segment_end = period
					segment_entries = entry_signature
				end
			end

			conflicts << conflict_record(date, term, room, segment_start, segment_end, period_map[segment_start]) unless segment_start.nil?
		end

		conflicts.sort_by { |conflict| [conflict[:date], conflict[:room], conflict[:start_period]] }
	end

	# conflict レコード用ハッシュを作成する
	# @param date [Date] 日付
	# @param term [String] 学期名
	# @param room [String] 講義室名
	# @param start_period [Integer] 開始時限
	# @param end_period [Integer] 終了時限
	# @param entries [Array<Hash>] セル情報の配列
	# @return [Hash] 予定の重複情報のハッシュ
	def conflict_record(date, term, room, start_period, end_period, entries)
		weekday = resolve_weekday(date, @calendar.date_list[date])
		{
			date: date,
			term: term,
			weekday: weekday,
			room: room,
			start_period: start_period,
			end_period: end_period,
			entries: entries.map { |entry| entry[:report_line] }.uniq
		}
	end

	# 学期の期間内の各(日付, 講義室, 時限)ごとのセル情報から、予定の重複情報を出力する
	# @param conflicts [Array<Hash>] 予定の重複情報の配列
	# @return [void]
	def print_conflicts(conflicts)
		return if conflicts.empty?

		puts " 予定の重複が #{conflicts.length} 件見つかりました:"
		puts ' ==================================================='

		conflicts.each_with_index do |conflict, idx|
			puts "  #{idx + 1} 件目"
			period_text = if conflict[:start_period] == conflict[:end_period]
								"#{conflict[:start_period]}限"
							else
								"#{conflict[:start_period]}-#{conflict[:end_period]}限"
							end

			puts "  [#{conflict[:date].month}月#{conflict[:date].day}日(#{conflict[:term]}) #{conflict[:weekday]}曜日 #{conflict[:room]} #{period_text}]"
			conflict[:entries].each do |line|
				puts "   #{line}"
			end
			puts ''
		end

		puts ' ==================================================='
	end

	# 指定されたシートのセルに値とスタイルを設定する
	# @param sheet [RubyXL::Worksheet] 対象のシート
	# @param row [Integer] 行番号（0始まり）
	# @param col [Integer] 列番号（0始まり）
	# @param value [String, nil] セルに設定する値
	# @param style_index [Integer, nil] セルに設定するスタイルのインデックス（nil の場合はスタイルを変更しない）
	# @return [void]
	def write_cell(sheet, row, col, value, style_index)
		ensure_cell(sheet, row, col)
		sheet[row][col].change_contents(value)
		sheet[row][col].style_index = style_index unless style_index.nil?
	end

	# 指定されたセルに、日付に応じた曜日の色を適用する
	# @param cell [RubyXL::Cell, nil] 対象のセル
	# @param date [Date] 日付
	# @return [void]
	def apply_weekday_color(cell, date)
		return if cell.nil?

		case date.wday
		when 6
			# 土曜日は水色
			cell.change_font_color('00B0F0')
		when 0
			# 日曜日は赤色
			cell.change_font_color('FF0000')
		end
	end

	# 日付を月/日の全角表記に変換する
	# @param date [Date] 変換対象の日付
	# @return [String] 月/日の全角表記
	def fullwidth_month_day(date)
		month = to_fullwidth_digits(date.month.to_s)
		day = to_fullwidth_digits(date.day.to_s)
		"#{month}/#{day}"
	end

	# 数字を全角に変換する
	# @param text [String] 変換対象の文字列
	# @return [String] 数字が全角に変換された文字列
	def to_fullwidth_digits(text)
		text.tr('0-9', '０-９')
	end

	# 指定されたシートのセルが存在しない場合に、セルを作成する
	# @param sheet [RubyXL::Worksheet] 対象のシート
	# @param row [Integer] 行番号（0始まり）
	# @param col [Integer] 列番号（0始まり）
	# @return [void]
	def ensure_cell(sheet, row, col)
		sheet.add_cell(row, col, nil) if sheet[row].nil? || sheet[row][col].nil?
	end

	# 日付情報から曜日を解決する
	# @param date [Date] 日付
	# @param date_info [DateInfo, nil] 日付情報（nil の場合は Calendar から取得する）
	# @return [String] 曜日
	def resolve_weekday(date, date_info)
		return date_info.weekday unless date_info.nil? || blank?(date_info.weekday)
		return date_info.lect_weekday unless date_info.nil? || blank?(date_info.lect_weekday)

		%w[日 月 火 水 木 金 土][date.wday]
	end

	# 日付情報から備考を抽出する
	# @param date_info [DateInfo, nil] 日付情報（nil の場合は空の配列を返す）
	# @return [Array<String>] 備考の配列
	def extract_remarks(date_info)
		return [] if date_info.nil? || date_info.desc.nil?

		Array(date_info.desc).map { |line| line.to_s.strip }.reject(&:empty?)
	end

	# 指定された値が nil または空文字列の場合に true を返す
	# @param value [Object] 判定対象の値
	# @return [Boolean] 値が nil または空文字列の場合は true、そうでない場合は false
	def blank?(value)
		value.nil? || value.to_s.strip.empty?
	end
end
