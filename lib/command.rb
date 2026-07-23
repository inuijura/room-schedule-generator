require 'tty-prompt'
require 'tty-cursor'
require 'reline'
require 'csv'

require_relative "calendar"
require_relative "reservation"
require_relative "room_info"
require_relative "xlsx_generator"
require_relative "print_formatter"

# コマンドラインインターフェースセッションを管理するクラス
class Command
  SELECTABLE_TERMS = ["1学期", "2学期", "3学期", "4学期", "夏季", "春季"].freeze
  COMMANDS = ["print", "read", "write", "register", "select", "quit"].freeze
  # select コマンドでの講義室名の定義ラベルのソート順を定義
  SORT_ORDER = {
    "(ALL)"  => 0,
    "(USER)" => 1,
    "(UNIV)" => 2,
  }.freeze

  REGISTERED_ROOM_EXPLANATION = <<~MSG
   - USER: 管理対象講義室ファイルにのみ登録されている講義室名
   - UNIV: 大学から提供される予約・時間割データにのみ登録されている講義室名
   - ALL:  管理対象講義室ファイルと予約・時間割データの両方に登録されている講義室名
  MSG

  attr_reader :calendar, :events, :lectures, :reservations, :room_info

  def initialize
    @command = nil
    @command_name = nil
    @command_args = []
    @calendar = nil
    @events = []                # 予約データから抽出したイベント情報をまとめたリスト
    @lectures = []              # 時間割データから抽出した講義情報をまとめたリスト
    @reservations = []          # 予約データと時間割データから抽出したイベント情報と講義情報をまとめたリスト
    @room_info = RoomInfo.new   # 講義室管理情報（[[講義室名] => [イベント情報または講義情報をまとめたリスト]]）
    @scripted_inputs = nil      # テスト時に対話入力の代わりに使う入力キュー
  end

  # テスト時に対話入力を自動化するための入力キューを設定する
  # @param inputs [Array<String>] read_command / read内で順に消費する入力値
  def set_scripted_inputs(inputs)
    @scripted_inputs = Array(inputs).map(&:to_s)
  end

  def read_command
    # コマンド入力受付
    Reline.completion_proc = proc do |word|
      words = Reline.line_buffer.split

      # 1つ目の単語（コマンド名）の場合はコマンド名の補完を行う
      # 2つ目以降の単語（引数）の場合はファイルパスの補完を行う
      if words.length <= 1
        COMMANDS.grep(/^#{Regexp.escape(word)}/)
      else
        normalized = word.tr("\\", "/")

        Dir.glob("#{normalized}*").map do |path|
          path = "#{path}/" if File.directory?(path)
          Gem.win_platform? ? path.tr("/", "\\") : path
        end
      end
    end

    @command = read_line(" > ")
  end

  def parse_command
    # 入力された文字列をコマンド名と引数に分ける
    # @command を解析して
    # @command_name と @command_args に格納
    tokens = @command.split

    @command_name = tokens[0]
    @command_args = tokens[1..] || []
  end

  def execute_command
    # コマンド名に応じて各処理を呼び出す
    case @command_name
    when "print"
      if @command_args.empty?
        exec_print
      else
        puts " [Error] print コマンドには引数が不要です"
      end
    when "read"
      if @command_args.empty?
        exec_read
      else
        puts " [Error] read コマンドには引数が不要です"
      end
    when "write"
      if @command_args.empty?
      exec_write
      else
        puts " [Error] write コマンドには引数が不要です"
      end
    when "register"
      if @command_args.empty?
        puts " [Error] register コマンドにはファイルパスが必要です"
      elsif @command_args.length > 1
        puts " [Error] register コマンドには1つのファイルパスのみ指定してください"
      else
        exec_register(@command_args[0])
      end
    when "select"
      if @command_args.empty?
        exec_select
      else
        puts " [Error] select コマンドには引数が不要です"
      end
    when "quit"
      if @command_args.empty?
        return :quit
      else
        puts " [Error] quit コマンドには引数が不要です"
      end
    else
      puts " [Error] コマンド '#{@command_name}' は存在しません"
    end

    :continue
  end

  private

  def exec_print # Print コマンドの処理
    if @calendar.nil? || @calendar.date_list.empty?
      puts " [Error] 先に read コマンドで学年暦データを読み込んでください"
      return
    end

    if @room_info.output_target_rooms.empty?
      puts " [Error] 出力対象講義室がありません"
      puts "         register/select コマンドで設定してください"
      return
    end

    begin

      # 前回の選択状態をデフォルトとして，再選択を促す
      terms = select_term(@last_selected_terms)
      @last_selected_terms = terms

      # 学期選択のプロンプトを消す
      cursor = TTY::Cursor
      print cursor.up(1)
      print cursor.clear_line

      puts " 以下の学期中の講義室管理情報を表示します:"
      for term in terms
        puts "  - #{term}"
      end
      puts ""

      formatter = PrintFormatter.new(
        calendar: calendar,
        room_info: room_info,
        reservations: @reservations,
        terms: terms
      )

      # 表示（format_room_infoでは整形のみ．こっちで表示する）
      format_room_info = formatter.format_room_info
      format_room_info.each { |line| puts line }
    
    rescue => e
      puts " [Error] 講義室管理情報の表示に失敗しました"
    end  
  end


  # ファイルの読み込み権限，ディレクトリの読み込み権限の確認とエラーメッセージの返却
  def check_file_readable(file_path, file_type_str)
    msg = ""

    begin 
      File.open(file_path, "r") {}
    rescue Errno::ENOENT # ファイルが存在しない場合
      msg = " [Error] #{file_type_str}データファイル '#{file_path}' がありません\n"
      if file_type_str == "学年暦"
        msg += "         学年暦データファイルは必須です"
      end
      return false, msg
    rescue Errno::EACCES # ファイルの読み込み権限がない場合
      msg = " [Error] #{file_type_str}データファイル '#{file_path}' の読み取り権限がありません\n"
      msg+= "         ファイルのアクセス権限を確認してください"
      return false, msg
    end

    # 拡張子が .xlsx でなければエラー
    unless File.extname(file_path).downcase == ".xlsx"
      msg = " [Error] #{file_type_str}データファイルの読み込みに失敗しました\n"
      msg+= "         期待する拡張子は '.xlsx' ですが，'#{File.extname(file_path)}' が入力されました"
      return false, msg
    end

    # ファイルが存在し，読み込み権限があり，拡張子が .xlsx であれば true を返す
    return true, ""
  end

  def exec_read # Read コマンドの処理
    # ファイルパスの補完機能を有効化
    Reline.completion_proc = proc do |word|
      normalized = word.tr("\\", "/")

      Dir.glob("#{normalized}*").map do |path|
        path = "#{path}/" if File.directory?(path)

        if Gem.win_platform?
          path.tr("/", "\\")
        else
          path
        end
      end
    end

    calendar_result = nil
    event_result = nil
    lecture_result = nil
    loaded_calendar = nil

    begin
      input_file_calendar = read_line(" 1. 学年暦データファイル: ").to_s.chomp
      
      if input_file_calendar.empty?
        err_msg = <<~MSG
         [Error] 学年暦データファイルのパスが入力されていません
                 学年暦データファイルは必須です
        MSG
        raise err_msg
      end

      # ファイルの読み取り権限，存在確認，拡張子確認をまとめて行う
      is_readable, err_msg = check_file_readable(input_file_calendar, "学年暦")
      unless is_readable
        raise err_msg
      end

      begin
        loaded_calendar = Calendar.new
        loaded_calendar.parse(input_file_calendar)
      rescue => e
        # Calendar 側で raise された具体的なメッセージを表示
        puts " [Error] '#{input_file_calendar}' の学年暦データ形式が不正です"
        return
      end

      calendar_result = " [Info] 学年暦データファイルは正常に読み込まれました"
    rescue => e
      loaded_calendar = nil
      puts "#{e.message}"
      return
    end

    input_file_event = read_line(" 2. 予約データファイル: ").to_s.chomp
    if input_file_event.empty?
      event_result = " [Info] 予約データファイルは未入力です"
      loaded_events = []
    else
      begin
        # ファイルの読み取り権限，存在確認，拡張子確認をまとめて行う
        is_readable, err_msg = check_file_readable(input_file_event, "予約")
        if !is_readable
          event_result = err_msg
          loaded_events = []
        else
          rows_event = Event.parse_excel_rows(input_file_event)
          loaded_events = Event.from_parsed_rows(rows_event)
          event_result = " [Info] 予約データファイルは正常に読み込まれました"
        end
      rescue => e
        loaded_events = []
        event_result = " [Error] 予約データファイル '#{input_file_event}' のデータ形式が不正です"
      end
    end

    input_file_lecture = read_line(" 3. 時間割データファイル: ").to_s.chomp
    if input_file_lecture.empty?
      lecture_result = " [Info] 時間割データファイルは未入力です"
      loaded_lectures = []
    else
      begin 
        # ファイルの読み取り権限，存在確認，拡張子確認をまとめて行う
        is_readable, err_msg = check_file_readable(input_file_lecture, "時間割")
        if !is_readable
          lecture_result = err_msg
          loaded_lectures = []
        else
          rows_lecture = Lecture.parse_excel_rows(input_file_lecture)
          loaded_lectures = Lecture.from_parsed_rows(rows_lecture)
          lecture_result = " [Info] 時間割データファイルは正常に読み込まれました"
        end
      rescue => e
        loaded_lectures = []
        lecture_result = " [Error] 時間割データファイル '#{input_file_lecture}' のデータ形式が不正です"
      end
    end

    # これは，予約データと時間割データから抽出したデータをまとめたもの
    loaded_reservations = (loaded_events + loaded_lectures).compact

    # 既存の講義室管理情報を使い回しつつ，予約情報だけを更新する
    @room_info.set_reservations(loaded_reservations)

    @calendar = loaded_calendar
    @events = loaded_events || []
    @lectures = loaded_lectures || []
    @reservations = loaded_reservations

    puts calendar_result if calendar_result
    puts event_result if event_result
    puts lecture_result if lecture_result
  end

  def exec_write # Write コマンドの処理
    if @calendar.nil? || @calendar.date_list.empty?
      puts " [Error] 先に read コマンドで学年暦データを読み込んでください"
      return
    end

    if @room_info.output_target_rooms.empty?
      puts " [Error] 出力対象講義室がありません"
      puts "         register/select コマンドで設定してください"
      return
    end

    begin
      year = @calendar.year || @calendar.date_list.keys.min&.year
      output_dir = "#{year}年度講義室予約表"

      # Ctrl + C（SIGINT）を検知した場合
      Signal.trap(:INT) do

        # writeコマンド実行中である場合は，以下の処理を行い，作成中のディレクトリ，ファイルを削除
        if File.directory?(output_dir)
          FileUtils.rm_rf(output_dir)
        end

        # ステータス0で正常終了
        exit(0)
      end

      generator = XlsxGenerator.new(
        calendar: @calendar,
        room_info: @room_info,
        events: @events,
        lectures: @lectures
      )
      generated_files = generator.generate_all(output_dir: output_dir)

      if generated_files.empty?
        puts " [Error] 出力対象学期が見つからなかったため、ファイルは生成されませんでした"
        return
      end

      output_dir = File.dirname(generated_files.first)
      puts " ディレクトリ '#{output_dir}' に予約表を作成しました:"
      generated_files.each do |path|
        puts "  - '#{File.basename(path)}'"
      end
    rescue => e
      puts " [Error] 予約表の生成に失敗しました"
    end
  
  end

  def exec_register(room_list_file) # Register コマンドの処理
    if room_list_file.nil?
      puts " [Error] 管理対象講義室ファイルのパスを引数として入力してください"
      return
    end

    unless File.exist?(room_list_file)
      puts " [Error] 管理対象講義室ファイル '#{room_list_file}' がありません"
      return
    end

    begin
      registered_rooms = parse_room_list_csv(room_list_file)

      if registered_rooms.empty?
        puts " [Error] 管理対象講義室ファイル '#{room_list_file}' から講義室名を取得できませんでした"
        return
      end

      @room_info.add_user_defined_room(registered_rooms)
      @room_info.set_output_target_rooms(registered_rooms)

      puts ""
      puts " 以下の講義室を管理対象として登録しました"
      @room_info.output_target_rooms.each { |room| puts "  - #{room}" }
    rescue => e
      puts " [Error] 管理対象講義室ファイル '#{room_list_file}' の解析に失敗しました"
    end
  end

  def exec_select # Select コマンドの処理

    output_room = []

    if @room_info.rooms.empty?
      err_msg = <<~MSG
       [Error] 管理対象講義室がありません
               register コマンドで管理対象講義室を登録してください
      MSG
      puts err_msg
      return
    end

    prompt = TTY::Prompt.new(
      interrupt: :signal,
      symbols: {
        marker: " >",
        radio_on:  "[*]",
        radio_off: "[ ]", 
      }
    )

    # room_info から，SORT_ORDER に従って講義室名をソートしたリストを作成する
    room_choices = @room_info.rooms.sort_by do |room_name|
      sort_key = room_definition_label(room_name)
      [SORT_ORDER.fetch(sort_key, 99), room_name]
    end.map do |room_name|
      definition_label = room_definition_label(room_name)
      {
        name: "#{room_name} #{definition_label}",
        value: room_name,
        definition: definition_label
      }
    end

    default_indices = @room_info.output_target_rooms.filter_map do |room|
      idx = room_choices.index { |c| c[:value] == room }
      idx + 1 if idx
    end

    # ラベル付きの講義室名を表示するために、room_choices から print_name を抽出
    room_choices_display = room_choices.map { |c| c[:name] }

    # ラベルの説明を表示
    puts REGISTERED_ROOM_EXPLANATION
    puts ""

    selected_rooms = prompt.multi_select(
      " 出力対象の講義室を選択してください:",
      room_choices_display,
      default: default_indices,
      show_help: :never,
      echo: false
    )

    # 選択された講義室の print_name から name を取得
    selected_rooms_name_list = selected_rooms.map do |selected_print_name|
      room_choices.find { |c| c[:name] == selected_print_name }[:value]
    end

    @room_info.set_output_target_rooms(selected_rooms_name_list)

    cursor = TTY::Cursor
    print cursor.up(1)
    print cursor.clear_line

    puts " 以下の講義室を出力対象として選択しました:"
    selected_rooms.each { |room| puts "  - #{room_display_name(room)}" }
  end

  def select_term(terms = nil) # 学期選択の処理

    begin
      prompt = TTY::Prompt.new(
        interrupt: :signal,
        symbols: {
          marker: " >",
          radio_on:  "[*]",
          radio_off: "[ ]",
        }
      )

      options = { show_help: :never, echo: false }
      options[:default] = terms if terms && !terms.empty?

      selected_terms = prompt.multi_select(" 表示対象の学期を選択してください:", SELECTABLE_TERMS, options)
      selected_terms
    rescue TTY::Reader::InputInterrupt, Interrupt
      raise Interrupt, "学期選択が中断されました"
    end
  end

  def parse_room_list_csv(file_path)
    rows = CSV.read(file_path, encoding: "bom|utf-8")
    return [] if rows.empty?

    rows.each_with_object([]) do |row, room_names|
      next if row.nil?

      row.each do |cell|
        room_name = cell.to_s.strip
        next if room_name.empty?

        room_names << room_name
      end
    end.uniq
  end

  def room_display_name(room_name)
    "#{room_name} #{room_definition_label(room_name)}"
  end

  def room_definition_label(room_name)
    room_entry = @room_info[room_name]
    return "" if room_entry.nil?

    user_defined = room_entry[:user_defined]
    excel_defined = room_entry[:excel_defined]

    if user_defined && excel_defined
      "(ALL)"
    elsif excel_defined
      "(UNIV)"
    elsif user_defined
      "(USER)"
    else
      ""
    end
  end

  def format_exception_message(exception, default_label: nil)
    return exception.to_s unless exception.is_a?(Reservation::InputDataError)

    label = default_label || "入力データ"
    base_message = exception.message.to_s.strip
    location = format_input_error_location(exception)
    return "#{label}の形式が不正です" if base_message.empty? && location.empty?

    [base_message, location].reject(&:empty?).join(" ")
  end

  def format_input_error_location(error)
    parts = []
    parts << "シート: #{error.sheet_name}" if error.sheet_name && !error.sheet_name.to_s.empty?
    parts << "行: #{error.row_number}" if error.row_number
    parts << "列: #{error.column}" if error.column && !error.column.to_s.empty?

    return "" if parts.empty?

    "(#{parts.join(', ')})"
  end

  def read_line(prompt)
    if @scripted_inputs && !@scripted_inputs.empty?
      value = @scripted_inputs.shift
      puts "#{prompt}#{value}"
      return value
    end

    Reline.readline(prompt, true)
  end

  def readable_data_file?(path)
    return false if path.nil? || path.to_s.empty?
    return false unless File.file?(path)

    File.readable?(path)
  end

end
