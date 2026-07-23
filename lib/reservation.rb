require 'tty-prompt'
require 'tty-cursor'

require 'reline'
require 'date'
require 'time'
require 'rubyXL'

# 予定クラス
# 講義・イベント・予約の基底クラス．共通属性と Excel パース処理を定義

# 予約データファイルと時間割データファイルを解析する際の共通処理をまとめたクラス
# @title [String] 講義名／イベント名
# @date [Date] 日付
# @start_time [Integer] 開始時限
# @end_time [Integer] 終了時限
# @room [Array<String>] 講義室名
# @person [Array<String>] 担当者名
class Reservation
  class InputDataError < StandardError
    attr_reader :file_path, :sheet_name, :row_number, :column, :field

    def initialize(message, file_path: nil, sheet_name: nil, row_number: nil, column: nil, field: nil, cause: nil)
      super(message)
      @file_path = file_path
      @sheet_name = sheet_name
      @row_number = row_number
      @column = column
      @field = field
      set_backtrace(cause.backtrace) if cause && cause.backtrace
    end

    def with_context(file_path: nil, sheet_name: nil, row_number: nil, column: nil, field: nil)
      self.class.new(
        message,
        file_path: file_path || @file_path,
        sheet_name: sheet_name || @sheet_name,
        row_number: row_number || @row_number,
        column: column || @column,
        field: field || @field
      )
    end

    def to_s
      context = []
      context << "file=#{@file_path}" if @file_path
      context << "sheet=#{@sheet_name}" if @sheet_name
      context << "row=#{@row_number}" if @row_number
      context << "column=#{@column}" if @column
      context << "field=#{@field}" if @field

      return super if context.empty?

      "#{super} (#{context.join(', ')})"
    end
  end

  MAX_TEXT_LENGTH = 64
  ALL_ROOMS = ["第１講義室", "第２講義室", "第３講義室", "第４講義室", "第５講義室", "第６講義室", "第７講義室", "第８講義室", "第９講義室", "第１０講義室", "第１１講義室", "第１２講義室", "第１３講義室", "第１４講義室", "第１５講義室", "第１６講義室", "第１７講義室"].freeze
  WEEKDAY_LIST = ["月", "火", "水", "木", "金", "土", "日"].freeze
  EXPECTED_HEADERS = [["code", "subject", "grade", "term", "week", "s_period", "e_period", "user", "room"],
                      ["date", "event", "s_period", "e_period", "user", "room"]].freeze

  attr_reader :title, :date, :start_time, :end_time, :room, :person, :term, :week

  # 初期化
  # @param title [String] 講義名／イベント名
  # @param start_time [Integer] 開始時限
  # @param end_time [Integer] 終了時限
  # @param room [Array<String>] 講義室名
  # @param person [Array<String>] 担当者名
  # @param date [Date] 日付
  # ※room・person・date は省略可能で，子クラスの初期化時に指定される
  def initialize(title:, start_time:, end_time:, room: nil, person: nil, date: nil, term: nil, week: nil)
    validated_start_time = self.class.normalize_and_validate_period(start_time, "start_time")
    validated_end_time = self.class.normalize_and_validate_period(end_time, "end_time")
    self.class.validate_period_range(validated_start_time, validated_end_time)

    @title = self.class.truncate_text(title)
    @date = date
    @start_time = validated_start_time
    @end_time = validated_end_time
    @room = self.class.truncate_value(room)
    @person = self.class.truncate_value(person)
    @term = term
    @week = week
  end

  # 与えられたXLSX形式のファイルをRubyXLで読み込み，各行をbuild_excel_rowで解析して配列で返す
  # @param file_path [String] Excelファイルのパス
  # @param start_row [Integer] 読み込み開始行番号（1始まり）
  # @param end_row [Integer] 読み込み終了行番号（1始まり）
  # @return [Array<Hash>] 各行の解析結果を格納した配列
  def self.parse_excel_rows(file_path)
    start_row = 2
    return [] unless File.exist?(file_path)

    xlsx = RubyXL::Parser.parse(file_path)
    sheet = xlsx[0]
    sheet_name = sheet&.sheet_name

    header_row = sheet[0]
    header = header_row.cells.map { |cell| cell&.value.to_s.strip }

    # 末尾の空白セルを無視
    header.pop while header.last == ""

    expected_header = self.expected_header
    unless expected_header == header && header.length == expected_header.length
      raise InputDataError.new(
        "ヘッダー形式が不正です。期待値: #{expected_header.inspect}, 実際: #{header.inspect}",
        file_path: file_path,
        sheet_name: sheet_name,
        row_number: 1
      )
    end

    parser = allocate
    rows = []
    row_number = start_row

    loop do
      row = sheet[row_number - 1]
      break if row.nil? || row.cells.all? { |cell| cell.nil? || cell.value.to_s.strip.empty? }

      begin
        parsed = parser.build_excel_row(row, row_number)
      rescue InputDataError => e
        raise e.with_context(file_path: file_path, sheet_name: sheet_name, row_number: row_number)
      rescue => e
        raise InputDataError.new(
          "行データの解析に失敗しました: #{e.message}",
          file_path: file_path,
          sheet_name: sheet_name,
          row_number: row_number,
          cause: e
        )
      end

      parsed[:__row_number] = row_number unless parsed.nil?
      rows << parsed unless parsed.nil?

      row_number += 1
    end

    rows
  end

  def self.expected_header
    raise NotImplementedError, "subclasses must implement expected_header"
  end

  def self.excel_column_name(index)
    return nil if index.nil? || index < 0

    index += 1
    name = +""
    while index > 0
      index -= 1
      name.prepend((65 + (index % 26)).chr)
      index /= 26
    end
    name
  end

  def self.blank_value?(value)
    return true if value.nil?
    return value.empty? if value.is_a?(Array)

    value.to_s.strip.empty?
  end

  def self.format_value_for_message(value)
    return "(nil)" if value.nil?

    text = value.is_a?(Array) ? value.join(",") : value.to_s
    text = text[0, 64] + "..." if text.length > 64
    text.inspect
  end

  def self.invalid_field_message(input_label:, field_label:, value:, expected:)
    "#{input_label}の#{field_label}が不正です。値: #{format_value_for_message(value)}。期待形式: #{expected}"
  end

  def self.validate_required!(row:, field:, label:, row_number:, column_map:, input_label:, expected:)
    return unless blank_value?(row[field])

    column = excel_column_name(column_map[field])
    raise InputDataError.new(
      "#{input_label}の#{label}は必須です。期待形式: #{expected}",
      row_number: row_number,
      column: column,
      field: field
    )
  end

  def self.parse_period_value!(value:, field:, field_label:, row_number:, column_map:, input_label:)
    period = strict_half_width_period(value)

    return period if period && period >= 1 && period <= 8

    raise InputDataError.new(
      invalid_field_message(
        input_label: input_label,
        field_label: field_label,
        value: value,
        expected: "1〜8 の半角数字"
      ),
      row_number: row_number,
      column: excel_column_name(column_map[field]),
      field: field
    )
  end

  # 1 行の Excel データをハッシュに変換する
  # 各サブクラス (Event, Lecture) で具体的な列の対応を定義する
  # 無視する行は nil を返して、parse_excel_rows の圧縮で落とされる
  # @param row [RubyXL::Row] Excel の 1 行
  # @return [Hash, nil] 解析結果のハッシュ、無視する行は nil
  def build_excel_row(_row)
    raise NotImplementedError, "subclasses must implement build_excel_row"
  end

  # XLSXファイルの項目「room」列の各セルにある文字列をカンマ区切りにして分解し，配列にして返す
  # @param row [Hash, nil] parse_excel_rows の 1 行分の結果
  # @return [Reservation, nil] 生成した Reservation 派生オブジェクト
  def self.parse_room_list(room_value)
    return [] if room_value.nil?

    # (1) room_valueをカンマごとに分割し，カンマごとの文字列（講義室名）を要素に持つ配列として返却
    # 半角数字を含む講義室名があった場合は，全角数字に変換して登録
    # 「全講義室」があった場合は，展開する
    room_value.to_s
              .split(/[,，]/)
              .map(&:strip)
              .reject(&:empty?)
              .map { |room| room.tr("0-9", "０-９") }
              .flat_map { |room| room == "全講義室" ? ALL_ROOMS : [room] }
              .map { |room| truncate_text(room) }
              .uniq
  end

  def self.truncate_text(value)
    return nil if value.nil?

    value.to_s[0, MAX_TEXT_LENGTH]
  end

  def self.truncate_value(value)
    return nil if value.nil?

    return value.map { |item| truncate_text(item) } if value.is_a?(Array)

    truncate_text(value)
  end

  def self.strict_half_width_period(value)
    raw = case value
          when Integer
            value.to_s
          else
            value.to_s.strip
          end

    return nil unless raw.match?(/\A[1-8]\z/)

    raw.to_i
  end

  def self.normalize_and_validate_period(value, field_name)
    period = strict_half_width_period(value)

    raise ArgumentError, "#{field_name} must be an integer between 1 and 8" if period.nil? || period < 1 || period > 8

    period
  end

  def self.validate_period_range(start_time, end_time)
    raise ArgumentError, "start_time must be less than or equal to end_time" if start_time > end_time
  end

  def self.normalize_and_validate_lecture_term(value)
    raw = value.to_s.strip
    raise ArgumentError, "term must be a digit between 1 and 4" unless raw.match?(/\A[1-4]\z/)

    raw
  end

  def self.normalize_and_validate_lecture_week(value)
    raw = value.to_s.strip
    allowed_weeks = %w[Mon Tue Wed Thu Fri]
    raise ArgumentError, "week must be one of Mon, Tue, Wed, Thu, Fri" unless allowed_weeks.include?(raw)

    raw
  end

end


# イベントクラス
# 単一の日付に紐づく一時的な予定 (講演・ワークショップなど)
class Event < Reservation
  INPUT_LABEL = "予約データ".freeze
  HEADER = ["date", "event", "s_period", "e_period", "user", "room"].freeze
  COLUMN_MAP = {
    date: 0,
    event: 1,
    s_period: 2,
    e_period: 3,
    user: 4,
    room: 5,
  }.freeze
  FIELD_LABELS = {
    date: "date",
    event: "event",
    s_period: "s_period",
    e_period: "e_period",
    user: "user",
    room: "room",
  }.freeze
  FIELD_EXPECTATIONS = {
    date: "YYYYMMDD の8桁半角数字",
    event: "1〜64文字",
    s_period: "1〜8 の半角数字",
    e_period: "1〜8 の半角数字",
    user: "1〜64文字",
    room: "1〜64文字（複数はカンマ区切り）",
  }.freeze

  # イベントを初期化する
  # @param title [String] イベント名
  # @param date [Date] 日付
  # @param start_time [Integer] 開始時限
  # @param end_time [Integer] 終了時限
  # @param room [Array<String>] 講義室名
  # @param user [Array<String>] 担当者名
  def initialize(title:, date:, start_time:, end_time:, room:, user: nil)
    inferred_term = self.class.term_from_date(date)
    super(
      title: title,
      start_time: start_time,
      end_time: end_time,
      room: room,
      person: user,
      date: date,
      term: inferred_term,
      week: nil
    )
  end

  # 日付を学期名へ変換する
  # 学年暦の境界に合わせて、1学期から春季休業までの文字列を返す
  # @param date [Date, nil] 判定対象の日付
  # @return [String, nil] 対応する学期名
  def self.term_from_date(date)
    return nil if date.nil?

    term_periods = Calendar.term_periods
    return nil if term_periods.nil?

    term_periods.each do |term_name, term_range|
      return term_name if term_range.include?(date)
    end

    nil
  end

  # person の別名を user として扱う
  alias user person

  # 予定の種別を示す文字列を返す
  def type
    "イベント"
  end

  # def detail
  #   return "" if room_text.empty?
  #   "場所: #{room_text}"
  # end
  
  # 複数の形式から Date オブジェクトへ変換する
  # 変換前の値は YYYYMMDD 形式の8桁半角数字のみ許可する
  # 条件に合わない値、または実在しない日付は ArgumentError を送出する
  # @param value [Object] 変換対象の値
  # @return [Date] 変換結果の日付
  def normalize_date(value)
    raw = if value.is_a?(Numeric)
            if value.respond_to?(:finite?) && !value.finite?
              value.to_s
            elsif value == value.to_i
              value.to_i.to_s
            else
              value.to_s
            end
          else
            value.to_s.strip
          end

    unless raw.match?(/\A[0-9]{8}\z/)
      raise ArgumentError, 'date must be an 8-digit string in YYYYMMDD format'
    end

    begin
      Date.strptime(raw, '%Y%m%d')
    rescue Date::Error
      raise ArgumentError, 'date must be a valid calendar date in YYYYMMDD format'
    end
  end

  # XLSXファイルの1行を受け取り，各項目を変数に格納する
  # 列順序は [日付, イベント名, 開始時限, 終了時限, 担当者, 講義室]
  # @param row [RubyXL::Row] XLSX形式のファイルの1行
  # @return [Hash] 解析結果のハッシュ
  def build_excel_row(row, row_number = nil)
    raw_date = row[0]&.value
    date = begin
      normalize_date(raw_date)
    rescue => e
      raise InputDataError.new(
        self.class.invalid_field_message(
          input_label: INPUT_LABEL,
          field_label: FIELD_LABELS[:date],
          value: raw_date,
          expected: FIELD_EXPECTATIONS[:date]
        ),
        row_number: row_number,
        column: self.class.excel_column_name(COLUMN_MAP[:date]),
        field: :date,
        cause: e
      )
    end

    {
      date: date,
      event: row[1]&.value,
      s_period: row[2]&.value,
      e_period: row[3]&.value,
      user: row[4]&.value,
      room: self.class.parse_room_list(row[5]&.value),
    }
  end

  def self.expected_header
    HEADER
  end

  # build_excel_rowで作成したデータをオブジェクト化し，Event型のインスタンスを1つ作成する
  # @param row [Hash, nil] parse_excel_rows の 1 行分の結果
  # @return [Event, nil] Event型のインスタンス
  def self.from_parsed_row(row)
    return nil if row.nil?

    row_number = row[:__row_number]
    validate_required!(row: row, field: :date, label: FIELD_LABELS[:date], row_number: row_number, column_map: COLUMN_MAP, input_label: INPUT_LABEL, expected: FIELD_EXPECTATIONS[:date])
    validate_required!(row: row, field: :event, label: FIELD_LABELS[:event], row_number: row_number, column_map: COLUMN_MAP, input_label: INPUT_LABEL, expected: FIELD_EXPECTATIONS[:event])
    validate_required!(row: row, field: :s_period, label: FIELD_LABELS[:s_period], row_number: row_number, column_map: COLUMN_MAP, input_label: INPUT_LABEL, expected: FIELD_EXPECTATIONS[:s_period])
    validate_required!(row: row, field: :e_period, label: FIELD_LABELS[:e_period], row_number: row_number, column_map: COLUMN_MAP, input_label: INPUT_LABEL, expected: FIELD_EXPECTATIONS[:e_period])
    validate_required!(row: row, field: :user, label: FIELD_LABELS[:user], row_number: row_number, column_map: COLUMN_MAP, input_label: INPUT_LABEL, expected: FIELD_EXPECTATIONS[:user])
    validate_required!(row: row, field: :room, label: FIELD_LABELS[:room], row_number: row_number, column_map: COLUMN_MAP, input_label: INPUT_LABEL, expected: FIELD_EXPECTATIONS[:room])

    start_period = parse_period_value!(value: row[:s_period], field: :s_period, field_label: FIELD_LABELS[:s_period], row_number: row_number, column_map: COLUMN_MAP, input_label: INPUT_LABEL)
    end_period = parse_period_value!(value: row[:e_period], field: :e_period, field_label: FIELD_LABELS[:e_period], row_number: row_number, column_map: COLUMN_MAP, input_label: INPUT_LABEL)

    if start_period > end_period
      raise InputDataError.new(
        "#{INPUT_LABEL}の時限範囲が不正です。値: s_period=#{start_period}, e_period=#{end_period}。期待形式: s_period <= e_period",
        row_number: row_number,
        column: excel_column_name(COLUMN_MAP[:e_period]),
        field: :e_period
      )
    end

    title = row[:event] || ""
    Event.new(
      title: title,
      date: row[:date],
      start_time: start_period,
      end_time: end_period,
      room: row[:room],
      user: row[:user]
    )
  rescue InputDataError
    raise
  rescue => e
    field = if e.message.include?("start_time")
              :s_period
            elsif e.message.include?("end_time")
              :e_period
            else
              nil
            end
    column = field ? excel_column_name(COLUMN_MAP[field]) : nil
    raise InputDataError.new(
      "イベント行の検証に失敗しました: #{e.message}",
      row_number: row_number,
      column: column,
      field: field,
      cause: e
    )
  end

  # メソッド「from_parsed_row」で作成したインスタンスから配列（イベント情報）を作成する
  # @param rows [Array<Hash>] メソッド「 parse_excel_rows 」で解析したXLSX形式のファイルの各行を要素に持つ配列
  # @return [Array<Event>]  Event型のインスタンスを要素に持つ配列
  def self.from_parsed_rows(rows)
    rows.map { |r| from_parsed_row(r) }.compact
  end

end


# 講義クラス
# 学期と曜日で定義される定期的な予定 (講義・授業など)
class Lecture < Reservation
  INPUT_LABEL = "時間割データ".freeze
  HEADER = ["code", "subject", "grade", "term", "week", "s_period", "e_period", "user", "room"].freeze
  COLUMN_MAP = {
    code: 0,
    subject: 1,
    grade: 2,
    term: 3,
    week: 4,
    s_period: 5,
    e_period: 6,
    user: 7,
    room: 8,
  }.freeze
  VALID_GRADES = %w[B1 B2 B3 B4 M1 M2 D1 D2 D3].freeze
  FIELD_LABELS = {
    code: "code",
    subject: "subject",
    grade: "grade",
    term: "term",
    week: "week",
    s_period: "s_period",
    e_period: "e_period",
    user: "user",
    room: "room",
  }.freeze
  FIELD_EXPECTATIONS = {
    code: "1文字以上の半角英数字",
    subject: "1〜64文字",
    grade: "B1/B2/B3/B4/M1/M2/D1/D2/D3 のいずれか",
    term: "1〜4 の半角数字",
    week: "Mon/Tue/Wed/Thu/Fri のいずれか",
    s_period: "1〜8 の半角数字",
    e_period: "1〜8 の半角数字",
    user: "65文字以上は切り捨て",
    room: "各講義室名65文字以上は切り捨て（複数はカンマ区切り）",
  }.freeze

  # 講義を初期化する
  # term と week で繰り返し回数を管理する
  # @param title [String] 講義名
  # @param start_time [Integer] 開始時限
  # @param end_time [Integer] 終了時限
  # @param room [Array<String>] 講義室名
  # @param teacher [Array<String>] 担当者名
  # @param term [String, nil] 学期名 (1学期, 2学期, 3学期, 4学期, 夏季休業, 春季休業)
  # @param week [String] 曜日名 (Mon, Tue, Wed, Thu, Fri)
  def initialize(title:, start_time:, end_time:, teacher:, room:, date:, term: nil, week: nil)
    validated_term = self.class.normalize_and_validate_lecture_term(term)
    validated_week = self.class.normalize_and_validate_lecture_week(week)

    super(
      title: title,
      start_time: start_time,
      end_time: end_time,
      room: room,
      person: teacher,
      date: nil,
      term: validated_term,
      week: validated_week
    )
  end

  # person の別名を teacher として扱う
  alias teacher person

  # 予定の種別を示す文字列を返す
  def type
    "講義"
  end

  # XLSXファイルの1行を受け取り，各項目を変数に格納する
  # 列順序は [講義コード, 科目名, 学年, 学期, 曜日, 開始時限, 終了時限, 担当者, 講義室]
  # @param row [RubyXL::Row] XLSX形式のファイルの1行
  # @return [Hash] 解析結果のハッシュ
  def build_excel_row(row, _row_number = nil)
    {
      code: row[0]&.value,
      subject: row[1]&.value,
      grade: row[2]&.value,
      term: row[3]&.value,
      week: row[4]&.value,
      s_period: row[5]&.value,
      e_period: row[6]&.value,
      user: row[7]&.value,
      room: self.class.parse_room_list(row[8]&.value),
    }
  end

  def self.expected_header
    HEADER
  end

  # build_excel_row で解析されたハッシュから Lecture オブジェクトを生成する
  # nil や不正な行はスキップし、パース失敗時は警告を出して nil を返す
  # @param row [Hash, nil] メソッド「build_excel_row」でまとめたデータ
  # @return [Lecture, nil] Lecture型のインスタンス
  def self.from_parsed_row(row)
    return nil if row.nil?

    row_number = row[:__row_number]
    validate_required!(row: row, field: :code, label: FIELD_LABELS[:code], row_number: row_number, column_map: COLUMN_MAP, input_label: INPUT_LABEL, expected: FIELD_EXPECTATIONS[:code])
    validate_required!(row: row, field: :subject, label: FIELD_LABELS[:subject], row_number: row_number, column_map: COLUMN_MAP, input_label: INPUT_LABEL, expected: FIELD_EXPECTATIONS[:subject])
    validate_required!(row: row, field: :grade, label: FIELD_LABELS[:grade], row_number: row_number, column_map: COLUMN_MAP, input_label: INPUT_LABEL, expected: FIELD_EXPECTATIONS[:grade])
    validate_required!(row: row, field: :term, label: FIELD_LABELS[:term], row_number: row_number, column_map: COLUMN_MAP, input_label: INPUT_LABEL, expected: FIELD_EXPECTATIONS[:term])
    validate_required!(row: row, field: :week, label: FIELD_LABELS[:week], row_number: row_number, column_map: COLUMN_MAP, input_label: INPUT_LABEL, expected: FIELD_EXPECTATIONS[:week])
    validate_required!(row: row, field: :s_period, label: FIELD_LABELS[:s_period], row_number: row_number, column_map: COLUMN_MAP, input_label: INPUT_LABEL, expected: FIELD_EXPECTATIONS[:s_period])
    validate_required!(row: row, field: :e_period, label: FIELD_LABELS[:e_period], row_number: row_number, column_map: COLUMN_MAP, input_label: INPUT_LABEL, expected: FIELD_EXPECTATIONS[:e_period])
    validate_required!(row: row, field: :user, label: FIELD_LABELS[:user], row_number: row_number, column_map: COLUMN_MAP, input_label: INPUT_LABEL, expected: FIELD_EXPECTATIONS[:user])
    validate_required!(row: row, field: :room, label: FIELD_LABELS[:room], row_number: row_number, column_map: COLUMN_MAP, input_label: INPUT_LABEL, expected: FIELD_EXPECTATIONS[:room])

    start_period = parse_period_value!(value: row[:s_period], field: :s_period, field_label: FIELD_LABELS[:s_period], row_number: row_number, column_map: COLUMN_MAP, input_label: INPUT_LABEL)
    end_period = parse_period_value!(value: row[:e_period], field: :e_period, field_label: FIELD_LABELS[:e_period], row_number: row_number, column_map: COLUMN_MAP, input_label: INPUT_LABEL)

    if start_period > end_period
      raise InputDataError.new(
        "#{INPUT_LABEL}の時限範囲が不正です。値: s_period=#{start_period}, e_period=#{end_period}。期待形式: s_period <= e_period",
        row_number: row_number,
        column: excel_column_name(COLUMN_MAP[:e_period]),
        field: :e_period
      )
    end

    grade = row[:grade].to_s.strip
    if !grade.empty? && !VALID_GRADES.include?(grade)
      raise InputDataError.new(
        invalid_field_message(
          input_label: INPUT_LABEL,
          field_label: FIELD_LABELS[:grade],
          value: row[:grade],
          expected: FIELD_EXPECTATIONS[:grade]
        ),
        row_number: row_number,
        column: excel_column_name(COLUMN_MAP[:grade]),
        field: :grade
      )
    end

    code = row[:code].to_s.strip
    if !code.empty? && !code.match?(/\A[A-Za-z0-9]+\z/)
      raise InputDataError.new(
        invalid_field_message(
          input_label: INPUT_LABEL,
          field_label: FIELD_LABELS[:code],
          value: row[:code],
          expected: FIELD_EXPECTATIONS[:code]
        ),
        row_number: row_number,
        column: excel_column_name(COLUMN_MAP[:code]),
        field: :code
      )
    end

    term = row[:term].to_s.strip
    unless term.match?(/\A[1-4]\z/)
      raise InputDataError.new(
        invalid_field_message(
          input_label: INPUT_LABEL,
          field_label: FIELD_LABELS[:term],
          value: row[:term],
          expected: FIELD_EXPECTATIONS[:term]
        ),
        row_number: row_number,
        column: excel_column_name(COLUMN_MAP[:term]),
        field: :term
      )
    end

    week = row[:week].to_s.strip
    unless %w[Mon Tue Wed Thu Fri].include?(week)
      raise InputDataError.new(
        invalid_field_message(
          input_label: INPUT_LABEL,
          field_label: FIELD_LABELS[:week],
          value: row[:week],
          expected: FIELD_EXPECTATIONS[:week]
        ),
        row_number: row_number,
        column: excel_column_name(COLUMN_MAP[:week]),
        field: :week
      )
    end

    title = row[:subject].to_s.strip
    Lecture.new(
      title: title,
      start_time: start_period,
      end_time: end_period,
      teacher: row[:user],
      room: row[:room],
      date: nil,
      term: term,
      week: week
    )
  rescue InputDataError
    raise
  rescue => e
    field = if e.message.include?("start_time")
              :s_period
            elsif e.message.include?("end_time")
              :e_period
            elsif e.message.include?("term")
              :term
            elsif e.message.include?("week")
              :week
            else
              nil
            end
    column = field ? excel_column_name(COLUMN_MAP[field]) : nil
    raise InputDataError.new(
      "時間割行の検証に失敗しました: #{e.message}",
      row_number: row_number,
      column: column,
      field: field,
      cause: e
    )
  end

  # 解析された複数の行から Lecture オブジェクトの配列を生成する
  # @param rows [Array<Hash>] parse_excel_rows の結果の配列
  # @return [Array<Lecture>] Lecture型のインスタンスを要素に持つ配列
  def self.from_parsed_rows(rows)
    rows.map { |r| from_parsed_row(r) }.compact
  end

end
