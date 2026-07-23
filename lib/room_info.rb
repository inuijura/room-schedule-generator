
# 講義室情報を管理するクラス
# @output_target_rooms: Excel 出力の対象とする講義室名の配列
# @rooms: 講義室名をキーとし、ユーザ定義フラグ、Excel 定義フラグ、紐づく Reservation 派生オブジェクトの配列を値とするハッシュ
class RoomInfo
	MAX_ROOM_NAME_LENGTH = 64

	attr_reader :output_target_rooms

	# 初期化
	def initialize
		@rooms = {}
		@output_target_rooms = []
	end

	# ユーザ定義の講義室を登録する。
	# @param room_names [String, Array<String>] 講義室名または講義室名の配列

	def add_user_defined_room(room_names)
		Array(room_names).each do |room_name|
			key = normalize_room(room_name)
			next if key.empty?

			entry = (@rooms[key] ||= {
				user_defined: false,
				excel_defined: false,
				reservations: []
			})

			entry[:user_defined] = true
		end
	end

	# Excel 由来の講義室として登録または更新する
	# @param room_names [String, Array<String>] 講義室名または講義室名の配列
	# @return [void]
	def add_excel_defined_room(room_names)
		Array(room_names).each do |room_name|
			key = normalize_room(room_name)
			next if key.empty?

			entry = (@rooms[key] ||= {
				user_defined: false,
				excel_defined: false,
				reservations: []
			})

			entry[:excel_defined] = true
		end
	end
	
	# Event / Lecture を受け取り、含まれる講義室をすべて Excel 由来として反映する
	# @param reservation [Event, Lecture, Array<Event>, Array<Lecture>] 対象の予約オブジェクトまたはその配列
	# @return [void]
	def add_reservation(reservation)
		Array(reservation).each do |single_reservation|
			next if single_reservation.nil?
			normalized_rooms = normalize_reservation_rooms(single_reservation)
			add_excel_defined_room(normalized_rooms)
			append_reservation_to_rooms(single_reservation, normalized_rooms)
		end
	end

	# add_reservation のエイリアス
	alias add_event add_reservation
	alias add_lecture add_reservation
	alias add_events add_reservation
	alias add_lectures add_reservation

	# 既存の講義室定義を使って Reservation 派生オブジェクト一覧を再設定する
	# @param reservations [Array<Event>, Array<Lecture>] 対象の予約オブジェクトの配列
	# @return [void]
	def set_reservations(reservations)
		@rooms.each_value do |entry|
			entry[:reservations] = []
		end

		Array(reservations).each do |reservation|
			next if reservation.nil?
			normalized_rooms = normalize_reservation_rooms(reservation)
			add_excel_defined_room(normalized_rooms)
			append_reservation_to_rooms(reservation, normalized_rooms)
		end
	end

	# 講義室名の一覧を取得する
	# @return [Array<String>] 登録されている講義室名の配列
	def rooms
		@rooms.keys
	end

	# 保存されている講義室情報の一覧を取得する
	# @return [Hash{String => Hash}] 講義室名をキー、講義室情報を値とするハッシュ
	def entries
		@rooms.transform_values do |entry|
			{
				user_defined: entry[:user_defined],
				excel_defined: entry[:excel_defined],
				reservations: entry[:reservations].dup
			}
		end
	end

	# 指定の講義室があるか
	# @param room_name [String] 講義室名
	# @return [Boolean] 講義室が存在する場合は true、そうでない場合は false
	def include?(room_name)
		@rooms.key?(normalize_room(room_name))
	end

	# 指定の講義室情報を取得する
	# @param room_name [String] 講義室名
	# @return [Hash, nil] 講義室情報のハッシュ（存在しない場合は nil）
	def [](room_name)
		@rooms[normalize_room(room_name)]
	end

	# 指定の講義室に紐付く Reservation 派生オブジェクト一覧を取得する
	# @param room_name [String] 講義室名
	# @return [Array] 講義室に紐付く Reservation 派生オブジェクトの配列（存在しない場合は空配列）
	def reservations_for(room_name)
		entry = self[room_name]
		return [] if entry.nil?

		entry[:reservations].dup
	end

	# 出力対象として扱う講義室名リストを設定する
    # @param room_names [String, Array<String>] 講義室名または講義室名の配列
	# @return [void]
	def set_output_target_rooms(room_names)
		@output_target_rooms = normalize_room_names(room_names).select { |room_name| include?(room_name) }
	end

	private

	# 講義室名を正規化する
	# @param room_name [String] 講義室名
	# @return [String] 正規化された講義室名
	def normalize_room(room_name)
		normalized = room_name.to_s.strip.each_char.map do |char|
			case char.ord
			when 0x20
				"　"
			when 0x21..0x7E
				# (char.ord + 0xFEE0).chr(Encoding::UTF_8)
				# ↓ こっちにするとエラーなくなるかも ↓
				(char.ord + 0xFEE0).chr(Encoding::UTF_16BE).encode(Encoding::UTF_8)
			else
				char
			end
		end.join

		normalized.slice(0, MAX_ROOM_NAME_LENGTH)
	end

	# 講義室名のリストを正規化する
	# @param room_names [String, Array<String>] 講義室名または講義室名の配列
	# @return [Array<String>] 正規化された講義室名の配列
	def normalize_room_names(room_names)
		Array(room_names).map { |room_name| normalize_room(room_name) }.reject(&:empty?).uniq
	end

	# Reservation オブジェクトから講義室名のリストを正規化して取得する
	# @param reservation [Event, Lecture] 対象の予約オブジェクト
	# @return [Array<String>] 正規化された講義室名の配列
	def normalize_reservation_rooms(reservation)
		room_values = reservation.respond_to?(:room) ? reservation.room : []
		normalize_room_names(room_values)
	end

	# Reservation オブジェクトを複数の講義室に紐付ける
	# @param reservation [Event, Lecture] 対象の予約オブジェクト
	# @param room_names [Array<String>] 紐付ける講義室名の配列
	# @param only_existing [Boolean] true の場合、既存の講義室にのみ紐付ける（存在しない講義室は無視される）。false の場合、存在しない講義室は新規登録されてから紐付けられる。
	# @return [void]
	def append_reservation_to_rooms(reservation, room_names, only_existing: false)
		room_names.each do |room_name|
			next if only_existing && !@rooms.key?(room_name)

			@rooms[room_name] ||= {
				user_defined: false,
				excel_defined: false,
				reservations: []
			}

			@rooms[room_name][:reservations] << reservation
		end
	end
end

