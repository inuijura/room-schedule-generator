require 'date'
require 'rubyXL'

# excel 数値→列変換マップ
$EXCEL_COLUMN_MAP = {
  0 => "A",
  1 => "B",
  2 => "C",
  3 => "D",
  4 => "E",
  5 => "F",
  6 => "G",
  7 => "H",
  8 => "I",
  9 => "J",
  10 => "K",
  11 => "L",
  12 => "M", 
  13 => "N",
  14 => "O",
  15 => "P",
  16 => "Q",
  17 => "R",
  18 => "S",
  19 => "T",
  20 => "U",
  21 => "V",
  22 => "W",
  23 => "X",
  24 => "Y",
  25 => "Z",
}

# 曜日マップ
$WEEKDAY_MAP = {
  0 => "日",
  1 => "月",
  2 => "火",
  3 => "水",
  4 => "木",
  5 => "金",
  6 => "土",
}

# 各月の最終日のマップ
$MONTH_END_DAY_MAP = {
  1 => 31,
  2 => 28,
  3 => 31,
  4 => 30,
  5 => 31,
  6 => 30,
  7 => 31,
  8 => 31,
  9 => 30,
  10 => 31,
  11 => 30,
  12 => 31
}



# セル位置を表すクラス
class CellPos
  attr_reader :row, :col

  def initialize(row, col)
    if col.class != Integer
      col = $EXCEL_COLUMN_MAP.key(col)
    else
      col = col
    end

    @row = row
    @col = col
  end
end

# セルの塗りつぶし情報と枠線情報を表すクラス
class CellColor
  attr_reader :fill_info, :border_info

  def initialize(fill_info, border_info)
    @fill_info = fill_info
    @border_info = border_info
  end
end

# 学期期間を表すクラス
class TermPeriod
  attr_reader :term, :start_date, :end_date

  def initialize(term, start_date, end_date)
    @term = term
    @start_date = start_date
    @end_date = end_date
  end

  def to_s
    str = "#{@term}: #{@start_date} 〜 #{@end_date}"
    str
  end
end

# 日付情報を表すクラス
class DateInfo
  # @weekday # 曜日
  # @type # 講義日種別 (例: 休講日、通常講義日など) のリスト
  # @desc # 備考 (例: 休講理由など)
  # @lect_weekday # 講義曜日 (例: 月曜日)

  attr_reader :weekday, :type, :desc, :lect_weekday, :term, :is_changed_weekday

  def initialize(weekday = nil, type = nil)
    @weekday = weekday || @date.strftime("%A")
    @type = type || []
  end

  def add_description(description)
    @desc ||= []

    if description.is_a?(Array)
      @desc.concat(description)
    else
      @desc << description
    end
  end

  def set_lecture_weekday(lect_weekday)
    @lect_weekday = lect_weekday
  end

  def to_s
    str = "曜日: #{@weekday}, 講義曜日: #{@lect_weekday}, 学期: #{@term}, 種別: #{@type}"
    if @desc
      str += ", 備考: [#{@desc.join(', ')}]"
    end
    str
  end
end


class Calendar
  class ParseError < StandardError; end
# 休講日のセル位置
  @workbook # Excel ワークブックオブジェクト
  @sheet # Excel シートオブジェクト

  @@month4_label_cell = CellPos.new(9 - 1, "B") # 4月ラベル
  @@month10_label_cell = CellPos.new(9 - 1, "M") # 10月ラベル

  @@term1_label_cell = CellPos.new(47 - 1, "B") # 1学期ラベル
  @@term1_date_cell = CellPos.new(47 - 1, "E") # 1学期期間
  @@term2_label_cell = CellPos.new(49 - 1, "B") # 2学期ラベル
  @@term2_date_cell = CellPos.new(49 - 1, "E") # 2学期期間
  @@term3_label_cell = CellPos.new(51 - 1, "B") # 3学期ラベル
  @@term3_date_cell = CellPos.new(51 - 1, "E") # 3学期期間
  @@term4_label_cell = CellPos.new(53 - 1, "B") # 4学期ラベル
  @@term4_date_cell = CellPos.new(53 - 1, "E") # 4学期期間
  @@exam_cell = CellPos.new(55 - 1, "B") # 例: 55行目B列
  @@holiday_cell = CellPos.new(47 - 1, "M") # 例: 47行目M列
  @@changed_weekday_cell = CellPos.new(49 - 1, "M") # 例: 49行目M列
  @@extra_lect_cell = CellPos.new(51 - 1, "M") # 例: 51行目M列

  @@year_cell = CellPos.new(2 - 1, "B") # 例: 2行目B列

  @@first_half_start_cell = CellPos.new(7 - 1, "D") # 例: 7行目D列
  @@second_half_start_cell = CellPos.new(7 - 1, "O") # 例: 7行目O列

  @@term1_period_cell = CellPos.new(47 - 1, "E") # 例: 47行目E列
  @@term2_period_cell = CellPos.new(49 - 1, "E") # 例: 49行目E列
  @@term3_period_cell = CellPos.new(51 - 1, "E") # 例: 51行目E列
  @@term4_period_cell = CellPos.new(53 - 1, "E") # 例: 53行目E列

  attr_reader :year
  attr_reader :date_list # 日付情報のハッシュ
  attr_reader :term_periods # 学期 -> 期間(Date Range) のハッシュ
  
  def self.term_periods
    @@term_periods
  end
  
  @@lecture_day_type_info_map = {} # 講義日種別，セルカラーマップ (Symbol -> string(theme))
  @@term_periods = nil

  @@is_loaded = false # 正しく学年暦が読み込まれたかどうかのフラグ

  def initialize(input_files = nil)
    @date_list = {}
    @term_periods = nil
  end

  # セルの塗りつぶし情報を取得
  # @param cell [RubyXL::Cell] セルオブジェクト
  # @param workbook [RubyXL::Workbook] Excel ワークブックオブジェクト
  # @return [String, nil] 塗りつぶし情報の文字列または nil
  def get_fill_info(cell_pos, workbook)
    cell = @sheet[cell_pos.row][cell_pos.col]

    # セルが nil の場合は例外
    if cell.nil?
      raise ParseError, "[Error] 塗りつぶし情報が存在すべきセルが nil です: セル位置 (#{cell_pos.row + 1}, #{$EXCEL_COLUMN_MAP[cell_pos.col]})"
    end

    # セルがシート中で定義されていない場合は例外
    if cell.style_index.nil?
      raise ParseError, "[Error] セルのスタイル情報が取得できません: セル位置 (#{cell_pos.row + 1}, #{$EXCEL_COLUMN_MAP[cell_pos.col]})"
    end

    fill_id = workbook.stylesheet.cell_xfs[cell.style_index]&.fill_id
    fill = workbook.stylesheet.fills[fill_id]

    color = fill&.pattern_fill&.fg_color
    
    if color.nil?
      return nil
    end

    if color.rgb
      info = "fill-rgb:#{color.rgb}"
    elsif color.theme
      # tint は上5桁程度に限定
      info = "fill-theme:#{color.theme}:#{color.tint.to_s[0, 5]}"
    elsif color.indexed
      info = "fill-indexed:#{color.indexed}"
    else
      info = nil
    end

    info
  end

  # セルの枠線情報を取得
  # @param cell [RubyXL::Cell] セルオブジェクト
  # @param workbook [RubyXL::Workbook] Excel ワークブックオブジェクト
  # @return [String, nil] 枠線情報の文字列または nil
  def get_border_info(cell_pos, workbook)
    cell = @sheet[cell_pos.row][cell_pos.col]
    # セルが nil の場合は例外
    if cell.nil?
      raise ParseError, "[Error] 枠線情報が存在すべきセルが nil です: セル位置 (#{cell_pos.row + 1}, #{$EXCEL_COLUMN_MAP[cell_pos.col]})"
    end

    # セルがシート中で定義されていない場合は例外
    if cell.style_index.nil?
      raise ParseError, "[Error] セルのスタイル情報が取得できません: セル位置 (#{cell_pos.row + 1}, #{$EXCEL_COLUMN_MAP[cell_pos.col]})"
    end

    style = workbook.stylesheet.cell_xfs[cell.style_index]
    border = workbook.stylesheet.borders[style.border_id]

    if border.nil?
      return nil
    end

    # 左右上下すべての枠線情報から，多い方の色を採用する
    left_color = border.left&.color
    right_color = border.right&.color
    top_color = border.top&.color
    bottom_color = border.bottom&.color

    border_colors = [left_color, right_color, top_color, bottom_color].compact

    if border_colors.empty?
      return nil
    end

    info = nil

    # rgb， theme， indexed 情報があるかを確認して多数決
    rgb_count = border_colors.count { |color| color.rgb }
    theme_count = border_colors.count { |color| color.theme }
    indexed_count = border_colors.count { |color| color.indexed }

    # 多数決で色を決定
    if rgb_count > theme_count && rgb_count > indexed_count
      # すべての枠線の色のうち，多い情報の色を採用する
      # most_common_color = border_colors.group_by { |color| color.rgb }.max_by { |_, group| group.size }&.first
      color = border_colors.group_by { |color| color.rgb }.max_by { |_, group| group.size }&.last&.first
      info = "border-rgb:#{color}"
    elsif theme_count > indexed_count
      color = border_colors.group_by { |color| color.theme }.max_by { |_, group| group.size }&.last&.first
      info = "border-theme:#{color}:#{color.tint.to_s[0, 5]}"
    else
      color = border_colors.group_by { |color| color.indexed }.max_by { |_, group| group.size }&.last&.first
      info = "border-indexed:#{color}"
    end

    info
  end

  # 学年暦のセルの塗りつぶし情報と枠線情報を解析して，講義日種別マップを作成
  # @return [Hash{String => String}] 講義日種別 -> セルの塗りつぶし情報/枠線情報 のハッシュ
  def make_day_type_map
    lecture_day_type_info_map = {}

    # 種別情報名と対応するセル位置のマップを作成
    day_type_cells = {
      "1学期" => @@term1_label_cell,
      "2学期" => @@term2_label_cell,
      "3学期" => @@term3_label_cell,
      "4学期" => @@term4_label_cell,
      "休講日" => @@holiday_cell,
      "試験日" => @@exam_cell,
      "補講日" => @@extra_lect_cell
    }

    # 各定義のセルが範囲外か確認し，範囲外の場合は例外を投げる
    day_type_cells.each do |day_type, cell|
      if cell.row < 0 || cell.row >= @sheet.sheet_data.rows.size || cell.col < 0 || cell.col >= @sheet.cols.size
        raise ParseError, "[Error] 講義日種別 '#{day_type}' のセル位置が範囲外です: (#{cell.row + 1}, #{$EXCEL_COLUMN_MAP[cell.col]})"
      end
    end

    term1_info = get_fill_info(@@term1_label_cell, @workbook)
    term2_info = get_fill_info(@@term2_label_cell, @workbook)
    term3_info = get_fill_info(@@term3_label_cell, @workbook)
    term4_info = get_fill_info(@@term4_label_cell, @workbook)
    exam_info = get_fill_info(@@exam_cell, @workbook)
    holiday_info = get_fill_info(@@holiday_cell, @workbook)
    
    changed_weekday_info = get_border_info(@@changed_weekday_cell, @workbook)
    extra_lect_info = get_border_info(@@extra_lect_cell, @workbook)


    # 他の講義日種別も同様に処理して、マップを完成させる

    # puts "1学期のセル色情報: #{term1_info}"
    lecture_day_type_info_map["1学期"] = term1_info

    # puts "2学期のセル色情報: #{term2_info}"
    lecture_day_type_info_map["2学期"] = term2_info

    # puts "3学期のセル色情報: #{term3_info}"
    lecture_day_type_info_map["3学期"] = term3_info

    # puts "4学期のセル色情報: #{term4_info}"
    lecture_day_type_info_map["4学期"] = term4_info

    # puts "休講日のセル色情報: #{holiday_info}"
    lecture_day_type_info_map["休講日"] = holiday_info

    # puts "試験日のセル色情報: #{exam_info}"
    lecture_day_type_info_map["試験日"] = exam_info

    lecture_day_type_info_map["変更授業日"] = changed_weekday_info

    # puts "臨時補講日のセル枠線情報: #{extra_lect_info}"
    lecture_day_type_info_map["補講日"] = extra_lect_info

    # 講義日種別のセル情報が nil の場合は，その講義日種別を出力してエラーとして処理
    lecture_day_type_info_map.each do |day_type, info|
      if info.nil?
        raise ParseError, "[Error] 講義日種別 '#{day_type}' のセル情報が取得できません"
      end
    end

    lecture_day_type_info_map
  end

  # 学年暦ファイルから年度を取得
  # @param cell [RubyXL::Cell] 年度を表すセルオブジェクト
  # @return [Integer, nil] 年度の整数値，取得できなかった場合は nils
  def get_calendar_year(cell)
    # セルの年を取得するロジックを実装

    if cell.nil? || cell.value.nil?
      raise ParseError, "[Error] 学年暦の年度を表すセルの値が空です"
    end

    str = cell.value
    # YYYY() ~ の形式を正規表現でマッチさせ，YYYYを抽出する (YYYY は全角数字も許容)
    # exp: ２０２６（令和８）年度　岡山大学授業日程計画
    if str =~ /([0-9０-９]{4})\s*[（(]/
      year_str = $1
      # 全角数字を半角数字に変換
      year_str = year_str.tr('０-９', '0-9')
      return year_str.to_i
    else
      raise ParseError, "[Error] 学年暦の年度を表すセルの値 '#{str}' の形式が不正です"
    end
  end

  # 学期期間を解析して，開始日と終了日の範囲を返す
  # @param text [String] 学期期間を表す文字列 (例: "4月1日〜7月31日")
  # @return [Range<Date>, nil] 学期期間の範囲 (Date Range)，解析に失敗した場合は nil
  def parse_term_period(text)
    return nil if text.nil?

    matched = text.to_s.match(/(\d{1,2})月\s*(\d{1,2})日\s*[〜～\-~]\s*(\d{1,2})月\s*(\d{1,2})日/)
    return nil unless matched

    start_month = matched[1].to_i
    start_day = matched[2].to_i
    end_month = matched[3].to_i
    end_day = matched[4].to_i
    start_year = @year
    end_year = end_month < start_month ? @year + 1 : @year

    Date.new(start_year, start_month, start_day)..Date.new(end_year, end_month, end_day)
  end

  # 学期期間を解析して，1~4，夏季休業，春季休業の学期期間のハッシュを作成
  # @return [Hash{String => Range<Date>}, nil] 学期名 -> 期間(Date Range) のハッシュ，解析に失敗した場合は nil
  def parse_term_periods
    term_cells = [
      ["1学期", @@term1_date_cell],
      ["2学期", @@term2_date_cell],
      ["3学期", @@term3_date_cell],
      ["4学期", @@term4_date_cell],
    ]

    term_periods = term_cells.each_with_object({}) do |(term_name, term_cell), term_periods|
      cell = @sheet[term_cell.row][term_cell.col]
      parsed = parse_term_period(cell&.value)

      # 学期期間が正しく解析できなかった (nilの) 場合は，nil にして終了
      if parsed.nil?
        raise ParseError, "[Error] 学期期間 '#{term_name}' のセルの値 '#{cell&.value}' の形式が不正です"
      end

      term_periods[term_name] = parsed
    end

    # 夏季休業は2学期終了の翌日から開始，3学期開始の前日まで
    if term_periods["2学期"] && term_periods["3学期"]
      summer_start = term_periods["2学期"].end + 1
      summer_end = term_periods["3学期"].begin - 1
      if summer_start <= summer_end
        term_periods["夏季休業"] = summer_start..summer_end
      end
    end

    # 春季休業は4学期終了の翌日から開始，1学期開始の前日まで
    if term_periods["4学期"] && term_periods["1学期"]
      winter_start = term_periods["4学期"].end + 1
      winter_end = term_periods["1学期"].begin - 1
      # winter_end は年を跨いでいるため、年を跨いでいる場合は winter_end を翌年の日付に変換する
      if winter_end < winter_start
        winter_end = Date.new(winter_end.year + 1, winter_end.month, winter_end.day)
      end

      if winter_start <= winter_end
        term_periods["春季休業"] = winter_start..winter_end
      end
    end

    term_periods
  end

  # 学期期間を表すセルの値を解析して、TermPeriod オブジェクトを作成
  # @param [String] term_name 学期名 (例: "1学期")
  # @param [CellPos] term_period_cell 学期期間を表すセルの位置
  # @return [TermPeriod, nil] TermPeriod オブジェクト．解析に失敗した場合は nil
  def make_term_period(term_name, term_period_cell)
    term_period_str = @sheet[term_period_cell.row][term_period_cell.col].value.to_s.strip
    normalized_term_period_str = term_period_str.gsub(/\s+/, "")

    if normalized_term_period_str =~ /(\d+)月(\d+)日?[〜～\-~](\d+)月(\d+)日?/
      start_month = $1.to_i
      start_day = $2.to_i
      end_month = $3.to_i
      end_day = $4.to_i

      begin
        start_date = Date.new(@year, start_month, start_day)
        end_date = Date.new(@year, end_month, end_day)
      rescue ArgumentError => e
        raise ParseError, "[Error] 学期期間を表すセルの値 '#{term_period_str}' の日付が不正です: #{e.message}"
      end

      return TermPeriod.new(term_name, start_date, end_date)
    else
      raise ParseError, "[Error] 学期期間を表すセルの値 '#{term_period_str}' の形式が不正です"
    end
  end

  # セルの塗りつぶし情報から講義日種別を判定する
  # @param color_info [CellColor] セルの塗りつぶし情報と枠線情報を持つ CellColor オブジェクト
  # @return [String, nil] 講義日種別 (例: "休講日", "試験日", "補講日", "授業日")
  def determine_lecture_day_type(color_info)
    @@lecture_day_type_info_map.each do |day_type, info|
      if color_info.fill_info == info || color_info.border_info == info
        if day_type == "休講日" || day_type == "試験日" || day_type == "補講日"
          return day_type
        else
          return "授業日"
        end
      end
    end
    nil
  end 

  # 日付から該当学期の文字列を取得する
  def determine_term_from_date(date)
    @term_periods.each do |term_name, term_range|
      if term_range.include?(date)
        return term_name
      end
    end
    raise ParseError, "[Error] 日付 #{date} の学期判定ができません"
  end

  # セルの塗りつぶし情報から該当学期の文字列を取得する
  def determine_term_from_cell(cell_pos)
    cell_fill_info = get_fill_info(cell_pos, @workbook)
    term_name_by_cell = @@lecture_day_type_info_map.key(cell_fill_info)
    term_name_by_cell
  end

  # セルの塗りつぶし情報と学期期間から，日付の該当学期を判定する
  def determine_term(cell_pos, date, date_info)
    # 日付から，該当学期とその文字列を取得
    term_name_by_date = determine_term_from_date(date)

    # セルの色情報と，該当学期の色情報が一致するかを確認
    term_name_by_cell = determine_term_from_cell(cell_pos)

    if term_name_by_date != term_name_by_cell
      raise ParseError, "[Error] 日付 #{date} の学期判定がセルの色情報と一致しません: 判定結果 '#{term_name_by_date}', セルの学期 '#{term_name_by_cell}'"
    end

    term_name_by_date
  end

  # ひと月分のカレンダーを解析して，日付ごとの情報のハッシュを作成
  # @param start_row [Integer] カレンダーの開始行
  # @param start_col [Integer] カレンダーの開始列
  # @param year [Integer] 年度  
  def parse_month_calendar(start_row, start_col, year, month)
    month_date_info_hash = {}
    desc_col = start_col + 7 # 備考はカレンダーの右隣の列にあると仮定

    # 日付セルの読み取り
    for row in start_row..(start_row + 5) # 例: 6行分を読み取る
      for col in start_col..(start_col + 6) # 例: 7列分を読み取る
        cell = @sheet[row][col]

        fill_info = get_fill_info(CellPos.new(row, col), @workbook)
        border_info = get_border_info(CellPos.new(row, col), @workbook)
        color_info = CellColor.new(fill_info, border_info)
        
        # セルが空の場合はスキップ
        if cell.nil? || cell.value.nil?
          next
        end

        # セルの値が数値でない場合はエラー
        unless cell.value.to_s =~ /^\d+$/
          raise ParseError, "[Error] 日付セルの値が数値ではありません: セル位置 (#{row + 1}, #{$EXCEL_COLUMN_MAP[col]}), 値: #{cell.value}"
        end

        # セルの塗りつぶし情報から講義日種別を判定
        lecture_day_type = determine_lecture_day_type(color_info)
        if lecture_day_type.nil?
          raise ParseError, "[Error] 講義日種別が特定できません: セル位置 (#{row + 1}, #{$EXCEL_COLUMN_MAP[col]})"
        end
 
        # 日付セルの値がその月の日数を超えていないか確認
        if cell.value.to_i <= 0 || cell.value.to_i > $MONTH_END_DAY_MAP[month]
          raise ParseError, "[Error] 日付セルの値が不正です: セル位置 (#{row + 1}, #{$EXCEL_COLUMN_MAP[col]}), 値: #{cell.value}"
        end

        # DateInfo を生成
        date = Date.new(year, month, cell.value.to_i) 
        date_info = DateInfo.new(
          $WEEKDAY_MAP[col - start_col],
          determine_lecture_day_type(color_info)
        )

        # 授業曜日が変更されているかを設定
        if border_info == @@lecture_day_type_info_map["変更授業日"]
          date_info.instance_variable_set(:@is_changed_weekday, true)
        else
          date_info.instance_variable_set(:@is_changed_weekday, false)
        end

        # 授業日の場合は学期を判定して設定
        if date_info.type == "授業日"
          date_term_name = determine_term(CellPos.new(row, col), date, date_info)
          date_info.instance_variable_set(:@term, date_term_name)
        end

        date_info.set_lecture_weekday(date_info.weekday)

        month_date_info_hash[date] = date_info

      end # 列ループ
    end # 行ループ

    # 月の日数を取得 (うるう年の2月は29日)
    num_days_in_month = $MONTH_END_DAY_MAP[month]

    # 要素数が月の日数と一致するか確認
    if month_date_info_hash.size != num_days_in_month
      raise ParseError, "[Error] 月のカレンダーの要素数が月の日数と一致しません: #{month_date_info_hash.size} != #{num_days_in_month}"
    end

    month_date_info_hash
  end


  # 備考情報の日付範囲を展開して，日付の配列を返す
  # @param start_month [Integer] 開始月
  # @param start_day [Integer] 開始日
  # @param end_month [Integer] 終了月
  # @param end_day [Integer] 終了日
  # @return [Array<Date>] 展開された日付の配列
  def expand_desc_date_range(start_month, start_day, end_month, end_day)
    desc_dates = []
    current_year = @year
    current_month = start_month
    current_day = start_day
    end_year = @year

    if end_month < start_month || (end_month == start_month && end_day < start_day)
      end_year += 1
    end

    loop do
      begin
        new_date = Date.new(current_year, current_month, current_day)

        desc_dates << new_date
        break if current_year == end_year && current_month == end_month && current_day == end_day

        current_day += 1
        last_day_of_month = Date.new(current_year, current_month, -1).day

        if current_day > last_day_of_month
          current_day = 1
          current_month += 1
          if current_month > 12
            current_month = 1
            current_year += 1
          end
        end
      rescue ArgumentError
        raise ParseError, "[Error] 備考情報の日付が不正です: #{current_year}-#{current_month}-#{current_day}"
      end
    end

    desc_dates
  end

  # 半年分の備考情報を解析して，日付ごとの備考情報のハッシュを作成
  # @param desc_start_cell [CellPos] 半年分の備考情報の開始セル位置
  # @param year [Integer] 年度
  # @param month [Integer] 開始月
  # @return [Hash{Date => Array<String>}] 半年分の日付ごとの備考情報のハッシュ
  def parse_half_year_desc_info(desc_start_cell, year, month)
    desc_hash = {}

    for i in 0..5
      start_cell = CellPos.new(desc_start_cell.row + (i * 6), desc_start_cell.col)

      for row in start_cell.row..(start_cell.row + 5)
        desc_days_cell = @sheet[row][desc_start_cell.col]
        desc_info_cell = @sheet[row][desc_start_cell.col + 1]

        if desc_days_cell.nil? && desc_info_cell.nil?
          next
        end

        # 日付・備考情報の片方のみが空の場合はエラー
        if (desc_days_cell.nil? || desc_days_cell.value.nil?) && (!desc_info_cell.nil? && !desc_info_cell.value.nil?)
          raise ParseError, "[Error] 備考情報が空です: セル位置 (#{row + 1}, #{$EXCEL_COLUMN_MAP[desc_start_cell.col + 1]})"
        end

        if (desc_info_cell.nil? || desc_info_cell.value.nil?) && (!desc_days_cell.nil? && !desc_days_cell.value.nil?)
          raise ParseError, "[Error] 備考情報の日付が空です: セル位置 (#{row + 1}, #{$EXCEL_COLUMN_MAP[desc_start_cell.col]})"
        end

        desc_days_str = desc_days_cell.value.to_s.strip
        normalized_desc_days_str = desc_days_str.gsub(/\s+/, "")

        begin
          if normalized_desc_days_str =~ /(\d+)月(\d+)日?[〜～\-~](\d+)月(\d+)日?/
            start_month = $1.to_i
            start_day = $2.to_i
            end_month = $3.to_i
            end_day = $4.to_i
            desc_days = expand_desc_date_range(start_month, start_day, end_month, end_day)
          elsif normalized_desc_days_str =~ /(\d+)\s*[〜～\-~]\s*(\d+)日?/
            desc_days = ($1.to_i..$2.to_i).to_a
            desc_days = desc_days.map { |desc_day| Date.new(year, month, desc_day) }
          elsif normalized_desc_days_str =~ /(\d+)日?[〜～\-~](\d+)月(\d+)日?/
            start_day = $1.to_i
            end_month = $2.to_i
            end_day = $3.to_i
            desc_days = expand_desc_date_range(month, start_day, end_month, end_day)
          else
            desc_days = desc_days_str.scan(/\d+/).map(&:to_i)
            desc_days = desc_days.map { |desc_day| Date.new(year, month, desc_day) }
          end
        rescue ArgumentError
          raise ParseError, "[Error] 備考情報の日付が不正です: #{desc_days_str}"
        end

        # 備考情報を日付ごとのハッシュに追加
        desc_days.each do |desc_day|
          # その月に存在しない日付の場合はエラー
          num_days_in_month = $MONTH_END_DAY_MAP[desc_day.month]
          if desc_day.day > num_days_in_month
            raise ParseError, "[Error] 備考情報の日付が不正です: #{desc_day} (月の最終日: #{num_days_in_month})"
          end

          desc_hash[desc_day] ||= []
          desc_hash[desc_day] << (desc_info_cell&.value || "")
        end
      end

      month += 1
      if month > 12
        month = 1
        year += 1
      end
    end

    desc_hash
  end

  # 備考情報を解析して，日付ごとの備考情報のハッシュを作成
  # @param year [Integer] 年度
  # @return [Hash{Date => Array<String>}] 日付ごとの備考情報のハッシュ
  def parse_desc_info(year)
    first_half_desc_start_cell = CellPos.new(@@first_half_start_cell.row, @@first_half_start_cell.col + 7)
    second_half_desc_start_cell = CellPos.new(@@second_half_start_cell.row, @@second_half_start_cell.col + 7)

    desc_info_hash = {}

    # 前期半年分の備考情報を解析
    half_year_desc_info = parse_half_year_desc_info(first_half_desc_start_cell, year, 4)

    desc_info_hash.merge!(half_year_desc_info) do |_date, existing, new_value|
      existing + new_value
    end

    # 後期半年分の備考情報を解析
    half_year_desc_info = parse_half_year_desc_info(second_half_desc_start_cell, year, 10)

    desc_info_hash.merge!(half_year_desc_info) do |_date, existing, new_value|
      existing + new_value
    end

    desc_info_hash
  end
  
  # カレンダーに存在する授業変更日に対応する備考情報があるかを確認
  def check_changed_lecture_desc_info(desc_info_hash)
    changed_lecture_dates = @date_list.select { |date, date_info| date_info.is_changed_weekday == true }.keys
    changed_lecture_dates.each do |changed_date|
      # 備考に「授業」という文言がなければエラー
      if desc_info_hash[changed_date].nil? || !desc_info_hash[changed_date].any? { |info| info.include?("授業") }
        raise ParseError, "[Error] 授業変更日 #{changed_date} に対応する備考情報がありません"
      end
    end
  end

  # 備考情報を該当の日付情報に反映
  # @param desc_info_hash [Hash{Date => Array<String>}]
  def reflect_desc_info(desc_info_hash)
    # date_info の各日付情報に対して、desc_info_hash の備考を反映
    @date_list.each do |date, date_info|
      if desc_info_hash.key?(date)
        date_info.add_description(desc_info_hash[date])

        # 備考情報に講義曜日が含まれている場合は、日付情報に反映
        lect_weekday = desc_info_hash[date].find { |info| info =~ /([月火水木金])曜日/ }&.match(/([月火水木金])曜日/)&.[](1)
        if lect_weekday
          if date_info.is_changed_weekday
            date_info.set_lecture_weekday(lect_weekday)
          else 
            raise ParseError, "[Error] 日付 #{date} は授業変更日のセルと一致していないが，備考情報に講義曜日が含まれています"
          end
        end
      end
    end
  end

  # Excel ファイルを解析して、日付表と学期期間を生成
  # @param [String] file_path Excel ファイルのパス
  def parse(file_path)
    # ファイルの内容を解析して、日付情報のリストを生成する処理を実装
    @date_list = {}
    @term_periods = nil
    @year = nil
    @@lecture_day_type_info_map = {}

    begin
      # 拡張子が .xlsx でない場合は例外
      unless File.extname(file_path) == ".xlsx"
        raise ParseError, "[Error] 学年暦ファイルの拡張子が .xlsx ではありません: #{file_path}"
      end

      xlsx = RubyXL::Parser.parse(file_path)
      @sheet = xlsx[0] # 最初のシートを取得
      @workbook = xlsx # Excel ワークブックオブジェクトを保存

      # 年度を取得
      @year = get_calendar_year(@sheet[@@year_cell.row][@@year_cell.col])
      if @year.nil?
        raise ParseError, "[Error] 学年暦の年度を取得できません. セル位置 (#{@@year_cell.row + 1}, #{$EXCEL_COLUMN_MAP[@@year_cell.col]})"
      end

      # うるう年の判定を行い、2月の日数を調整
      if Date.gregorian_leap?(@year)
        $MONTH_END_DAY_MAP[2] = 29
      end

      # 日付情報の作成に必要な基本情報の解析と取得
      @@lecture_day_type_info_map = make_day_type_map()
      @term_periods = parse_term_periods()
      @@term_periods = @term_periods

      # 日付情報の解析とハッシュの作成
      start_row = @@first_half_start_cell.row
      start_col = @@first_half_start_cell.col

      month = 4 # 例: 4月から開始
      calendar_year = @year

      # 前期半年分のカレンダーを読み取る
      for i in 0..5
        month_label_cell = @sheet[@@month4_label_cell.row + (i * 6)][@@month4_label_cell.col]
        if month_label_cell.nil? || month_label_cell.value.to_i != month
          raise ParseError, "[Error] カレンダーの月ラベルが不正です: セル位置 (#{@@month4_label_cell.row + (i * 6) + 1}, #{$EXCEL_COLUMN_MAP[@@month4_label_cell.col]}), 値: #{month_label_cell&.value}, 期待値: #{month}"
        end

        month_date_info_hash = parse_month_calendar(start_row + (i * 6), start_col, calendar_year, month)
        @date_list.merge!(month_date_info_hash)
        month += 1
      end

      start_row = @@second_half_start_cell.row
      start_col = @@second_half_start_cell.col

      # 後期半年分のカレンダーを読み取る
      for i in 0..5
        month_label_cell = @sheet[@@month10_label_cell.row + (i * 6)][@@month10_label_cell.col]
        if month_label_cell.nil? || month_label_cell.value.to_i != month
          raise ParseError, "[Error] カレンダーの月ラベルが不正です: セル位置 (#{@@month10_label_cell.row + (i * 6) + 1}, #{$EXCEL_COLUMN_MAP[@@month10_label_cell.col]}), 値: #{month_label_cell&.value}, 期待値: #{month}"
        end

        month_date_info_hash = parse_month_calendar(start_row + (i * 6), start_col, calendar_year, month)
        @date_list.merge!(month_date_info_hash)
        month += 1
        if month > 12
          calendar_year += 1
          month = month - 12
        end
      end

      calendar_year = @year

      desc_info = parse_desc_info(calendar_year)
      # desc_info が空の場合の場合は例外
      if desc_info.empty?
        raise ParseError, "[Error] 備考情報が空です"
      end

      # 授業変更日に対応する備考情報があるかを確認
      check_changed_lecture_desc_info(desc_info)

      # date_info の各日付情報に対して、desc_info の備考を反映
      reflect_desc_info(desc_info)
    rescue ParseError => e
      @date_list = nil
      @term_periods = nil
      @year = nil
      @@lecture_day_type_info_map = {}
      # puts " #{e.message}"

      # parse 呼び出し側に例外を伝えるために再度 raise する
      raise e
    end
  end

  def print_calendar()
    @date_list.each do |date, date_info|
      puts "#{date}: #{date_info.to_s}"
    end
  end

  def print_term_periods()
    @term_periods.each do |term, period|
      puts "#{term}: #{period.begin} - #{period.end}"
    end
  end

  # 学年暦が解析でき，日付情報と学期期間の有効なデータが作成されたかを確認
  # @return [Boolean] 有効なデータが作成されていれば true，そうでなければ false
  def is_valid? 
    if @date_list.nil? || @term_periods.nil?
      return false
    end
    true
  end
end