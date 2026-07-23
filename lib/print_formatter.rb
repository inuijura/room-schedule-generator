require 'date'

# ターミナルに講義室情報を整形して出力するクラス
class PrintFormatter
  ROOM_HEADER_PREFIX = "================== "
  ROOM_HEADER_SUFFIX = " =================="
  ENTRY_INDENT = "    "
  DETAIL_INDENT = "      "
  DATE_SEPARATOR = " "
  TITLE_LABEL = "講義名/イベント名"
  PERSON_LABEL = "担当者"
  PLACE_LABEL = "日時・場所"
  NOTE_LABEL = "備考"
  PLACE_LIST_SEPARATOR = "、"
  ROOM_GROUP_SEPARATOR = "\n"
  EMPTY_TEXT = ""

  TERM_ORDER = ["1学期", "2学期", "3学期", "4学期", "夏季休業", "春季休業"].freeze
  WEEK_ORDER = ["月", "火", "水", "木", "金", "土", "日"].freeze
  WEEK_DISPLAY_LABELS = {
    "月" => "月曜日",
    "火" => "火曜日",
    "水" => "水曜日",
    "木" => "木曜日",
    "金" => "金曜日",
    "土" => "土曜日",
    "日" => "日曜日",
  }.freeze
  WEEK_LABELS = {
    "Mon" => "月",
    "Tue" => "火",
    "Wed" => "水",
    "Thu" => "木",
    "Fri" => "金",
    "Sat" => "土",
    "Sun" => "日",
    "mon" => "月",
    "tue" => "火",
    "wed" => "水",
    "thu" => "木",
    "fri" => "金",
    "sat" => "土",
    "sun" => "日",
  }.freeze
  TERM_LABELS = {
    1 => "1学期",
    2 => "2学期",
    3 => "3学期",
    4 => "4学期",
    "1" => "1学期",
    "2" => "2学期",
    "3" => "3学期",
    "4" => "4学期",
  }.freeze

  def initialize(calendar:, room_info:, reservations:, terms: nil)
    # @calendar: Calendar オブジェクト（学年暦・日付情報）
    # @room_info: RoomInfo オブジェクト（講義室メタ情報）
    # @reservations: Event/Lecture オブジェクト配列
    # @terms: 表示対象の学期（文字列配列）
    @calendar = calendar
    @room_info = room_info
    @reservations = Array(reservations)
    @terms = Array(terms).compact
  end

  # 講義室ごとに予約をまとめ，各講義室の見出し -> 予約ブロックの順で標準出力する
  # @param なし
  # @return [Array<String>] 標準出力用の文字列配列
  def format_room_info
    validate_inputs!

    # 表示したい文字列一式を格納する配列，Commandクラス（呼び出し）側でまとめて出力する
    result = []

    # (1) 出力対象講義室を取得
    # RoomInfo から出力対象の講義室を引く
    rooms = @room_info.output_target_rooms
    
    rooms.each do |room|

      # 出力対象講義室の講義室名で見出しを作成
      result << room_header(room)

      # (2) 講義室管理情報から出力対象講義室の講義室管理情報を抽出
      room_reservations = @room_info.reservations_for(room)
      next if room_reservations.empty?

      # (3) (2)で抽出した講義室管理情報からreservation_visible?を用いて，表示対象期間内の講義室管理情報を抽出
      visible_reservations = room_reservations.select { |reservation| reservation_visible?(reservation) }
      next if visible_reservations.empty?

      # (4) reservationをformat_reservation_blockで整形し，文字列をまとめる配列に追加
      visible_reservations.each do |reservation|
        # 呼出側で表示するように変更
        # puts format_reservation_block(reservation)
        result << format_reservation_block(reservation)
      end

      result << ROOM_GROUP_SEPARATOR
    end
    
    # (5) (2)～(4)を繰り返し，すべてのreservationを整形し，文字列を配列に追加後，配列を返す
    result << " 講義室管理情報は正しく表示されました"
  end

  private

  # 入力オブジェクトの基本チェックを行う，Calendar と RoomInfo は必須
  # @raise [ArgumentError] Calendar または RoomInfo が未設定の場合に発生
  def validate_inputs!
    raise ArgumentError, 'Calendar が未設定です' if @calendar.nil?
    raise ArgumentError, 'RoomInfo が未設定です' if @room_info.nil?
  end

  # Calendar に登録されている学期境界を引き，指定学期の DateRange を返す
  # 期間が存在しない学期は nil のまま返して，呼び出し側で除外できるようにする
  # @param term [String] 学期ラベル
  # @return [DateRange, nil] 指定学期の期間範囲
  def term_range(term)
    @calendar.term_periods[term]
  end

  # ドキュメントにあわせて作成
  # 見出しを作成し，返す
  # @param room [String] 講義室名
  # @return [String] 見出し文字列
  def room_header(room)
    "#{ROOM_HEADER_PREFIX}#{room}#{ROOM_HEADER_SUFFIX}"
  end

  # 現在の学期フィルタに照らして，その予約を出力するかを判定する
  # 講義は学期指定で判定し，イベントは実際の日付が学期範囲に入るかで判定する
  # @param reservation [Reservation] 判定対象の予約オブジェクト
  # @return [Boolean] 出力対象なら true，そうでなければ false
  def reservation_visible?(reservation)
    if lecture_reservation?(reservation)
      # 講義 (term/weekday) は学期ラベルで判定
      @terms.include?(lecture_term_label(reservation.term))
    else
      # イベントは日付がその学期範囲に含まれるかで判定
      @terms.any? { |term| reservation.date && term_range(term)&.include?(reservation.date) }
    end
  end

  # 1 件の予約を，タイトル・日時/場所・備考の 3 ブロックに分けた複数行文字列へ整形する
  # ターミナル表示時に見出しを揃えるため，ここでラベルとインデントを組み立てる
  # @param reservation [Reservation] 整形対象の予約オブジェクト
  # @return [String] 整形済みの複数行文字列
  def format_reservation_block(reservation)
    date_label = reservation_date_label(reservation)
    time_label = reservation_time_label(reservation)
    place_label = reservation_place_label(reservation)
    title_label = title_label(reservation)
    title_text = reservation_title_for(reservation)
    person_label = reservation_person_label(reservation)
    note_label = reservation_note_label(reservation)

    [
      "#{ENTRY_INDENT}[#{title_label}] #{title_text} #{person_label}",
      "#{ENTRY_INDENT}[#{PLACE_LABEL}]",
      "#{DETAIL_INDENT}#{date_label}#{DATE_SEPARATOR}#{time_label}#{DATE_SEPARATOR}#{place_label}",
      "#{ENTRY_INDENT}[#{NOTE_LABEL}]",
      "#{DETAIL_INDENT}#{note_label}",
      EMPTY_TEXT,
    ].join("\n")
  end

  # 開始時刻と終了時刻を結合し，表示用の時刻範囲文字列にする
  # @param start_time [Time, String, nil] 開始時刻
  # @param end_time [Time, String, nil] 終了時刻
  # @return [String] 整形済みの時刻範囲文字列
  def format_time_range(start_time, end_time)
    start_str = time_to_s(start_time)
    end_str = time_to_s(end_time)
    "#{start_str}-#{end_str}"
  end

  # start_time と end_time を文字列化する
  # @param time [Time, String, nil] 変換対象の時刻
  # @return [String] 表示用の文字列
  def time_to_s(time)
    time.to_s
  end

  # date を持たず term を持つ予約を講義扱いにする
  # イベント系の予約と分岐するための判定で，以降の整形・備考取得を切り替える
  # @param reservation [Reservation] 判定対象の予約オブジェクト
  # @return [Boolean] 講義扱いなら true，そうでなければ false
  def lecture_reservation?(reservation)
    reservation.respond_to?(:term) && reservation.term && reservation.date.nil?
  end

  # 数値または文字列で渡された学期番号を，表示用の学期ラベルに揃える
  # 予約データの揺れを吸収して，学期表記を統一するための補助関数
  # @param term [Integer, String] 学期番号または学期ラベル
  # @return [String] 表示用の学期ラベル
  def lecture_term_label(term)
    TERM_LABELS[term] || TERM_LABELS[term.to_s] || term.to_s
  end

  # 曜日の入力揺れを吸収して，比較しやすい 1 文字ラベルへ寄せる
  # 英語表記や余計な空白があっても，表示・比較で扱いやすい形式に統一する
  # @param week [String] 曜日ラベル
  # @return [String] 正規化された 1 文字の曜日ラベル
  def normalize_week_label(week)
    value = week.to_s.strip
    return "" if value.empty?

    WEEK_LABELS[value] || WEEK_LABELS[value[0, 3]] || value[0]
  end

  # 予約の種類に応じて，一覧の 2 行目に出す日時ラベルを作る
  # 講義は「学期 + 曜日」，イベントは「日付 + 曜日」を表示して区別する
  # @param reservation [Reservation] 日時ラベルを作る対象の予約オブジェクト
  # @return [String] 整形済みの日時ラベル文字列
  def reservation_date_label(reservation)
    if lecture_reservation?(reservation)
      week = normalize_week_label(reservation.week)
      term = lecture_term_label(reservation.term)
      [term, lecture_weekday_label(week)].reject(&:empty?).join(DATE_SEPARATOR)
    else
      reservation.date ? reservation.date.strftime("%Y-%m-%d（%a）") : "不明"
    end
  end

  # 予約の時間表示を整形する
  # @param reservation [Reservation] 時間ラベルを作る対象の予約オブジェクト
  # @return [String] 整形済みの時間ラベル文字列
  def reservation_time_label(reservation)
    format_time_range(reservation.start_time, reservation.end_time)
  end

  # 1 文字の曜日ラベルを，表示用の完全な曜日名に展開する
  # @param week_label [String] 1 文字の曜日ラベル
  # @return [String] 表示用の完全な曜日名
  def lecture_weekday_label(week_label)
    WEEK_DISPLAY_LABELS[week_label] || week_label
  end

  # 予約に紐づく講義室名をまとめ，見出しの下に置ける丸括弧付き文字列へ整形する
  # 複数講義室の共同利用でも，一覧上は 1 行に収めるための表示処理
  # @param reservation [Reservation] 講義室名をまとめる対象の予約オブジェクト
  # @return [String] 丸括弧付きの講義室名文字
  def reservation_place_label(reservation)
    rooms = Array(reservation.room)
    return "" if rooms.empty?

    "（#{rooms.join(PLACE_LIST_SEPARATOR)}）"
  end

  # 予約タイトルをそのまま表示文字列として返す
  # @param reservation [Reservation] タイトルを返す対象の予約オブジェクト
  # @return [String] 予約タイトル文字列
  def reservation_title_for(reservation)
    reservation.title.to_s
  end

  # 予約の種類に応じてタイトルラベルを返す
  # @param reservation [Reservation] タイトルラベルを決める対象の予約オブジェクト
  # @return [String] 講義なら「講義名」，イベントなら「イベント名」
  def title_label(reservation)
    lecture_reservation?(reservation) ? "講義名" : "イベント名"
  end

  # 担当者名をラベル付きで表示する
  # @param reservation [Reservation] 担当者名を返す対象の予約オブジェクト
  # @return [String] ラベル付きの担当者名文字列
  def reservation_person_label(reservation)
    "（#{PERSON_LABEL}：#{reservation.person.to_s}）"
  end

  # Calendar から引いた備考を，表示用の 1 行文字列にまとめる
  # 備考が空なら「なし」を返し，複数ある場合は区切り文字で連結する
  # @param reservation [Reservation] 備考をまとめる対象の予約オブジェクト
  # @return [String] 備考をまとめた 1 行文字列
  def reservation_note_label(reservation)
    notes = reservation_calendar_notes(reservation)
    notes.empty? ? "なし" : notes.join(PLACE_LIST_SEPARATOR)
  end

  # 予約の種類に応じて，Calendar から拾う備考の取り方を切り替える
  # 講義は期間全体を走査して該当曜日の備考を集め，イベントは日付単位で取る
  # @param reservation [Reservation] 備考を集める対象の予約オブジェクト
  # @return [Array<String>] 備考の文字列配列
  def reservation_calendar_notes(reservation)
    if lecture_reservation?(reservation)
      lecture_calendar_notes(reservation)
    else
      event_calendar_notes(reservation)
    end
  end

  # イベント予約は日付を基準に，該当日の Calendar 備考だけを返す
  # @param reservation [Reservation] 備考を集める対象の予約オブジェクト
  # @return [Array<String>] 備考の文字列配列
  def event_calendar_notes(reservation)
    return [] if reservation.date.nil?

    date_info = @calendar.date_list[reservation.date]
    calendar_notes(date_info)
  end

  # 講義予約は学期範囲を総当たりし，曜日が一致する日だけ備考を集める
  # 学期内に複数回出る講義を，実際に該当する全日分まとめて表示するための処理
  # @param reservation [Reservation] 備考を集める対象の予約オブジェクト
  # @return [Array<String>] 備考の文字列配列
  def lecture_calendar_notes(reservation)
    range = term_range(lecture_term_label(reservation.term))
    return [] if range.nil?

    normalized_week = normalize_week_label(reservation.week)
    notes = []

    range.each do |date|
      date_info = @calendar.date_list[date]
      next if date_info.nil?
      next unless lecture_weekday_matches?(date_info, normalized_week, date)

      notes.concat(calendar_notes(date_info))
    end

    notes.compact.map(&:to_s).reject(&:empty?).uniq
  end

  # Calendar の date_info から desc を抜き出し，配列で返す
  # @param date_info [DateInfo, nil] Calendar の日付情報オブジェクト
  # @return [Array<String>] 備考の文字列配列
  def calendar_notes(date_info)
    return [] if date_info.nil?
    Array(date_info.desc).map(&:to_s).reject(&:empty?)
  end

  # 学期内の各日付が，予約に指定された曜日と一致するかを判定する
  # @param date_info [DateInfo] Calendar の日付情報オブジェクト
  # @param normalized_week [String] 正規化された 1 文字の曜日ラベル
  # @param date [Date] 判定対象の日付
  # @return [Boolean] 曜日が一致するなら true，そうでなければ false
  def lecture_weekday_matches?(date_info, normalized_week, date)
    expected_weekday = WEEK_LABELS[normalized_week] || normalized_week
    actual_weekday = lecture_weekday_label_from_date_info(date_info, date)
    expected_weekday == actual_weekday
  end

  # date_infoの持つ曜日情報を，1文字ラベルへ正規化して返す
  # @param date_info [DateInfo] Calendar の日付情報オブジェクト
  # @param date [Date] 判定対象の日付
  # @return [String] 正規化された 1 文字の曜日ラベル
  def lecture_weekday_label_from_date_info(date_info, date)
    if date_info.respond_to?(:lect_weekday) && !date_info.lect_weekday.to_s.empty?
      # 講義曜日（lect_weekday）があるならそれを使う
      date_info.lect_weekday.to_s[0]
    elsif date_info.respond_to?(:weekday) && !date_info.weekday.to_s.empty?
      # それ以外は，Calendar の weekday を使う
      date_info.weekday.to_s[0]
    else
      # これは実行されない想定
      %w[日 月 火 水 木 金 土][date.wday]
    end
  end
end
