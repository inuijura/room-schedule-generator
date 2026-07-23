require 'pathname'
require 'fileutils'
require 'time'
require 'stringio'
require 'rubyXL'

require_relative '../../lib/command'

ROOT_DIR = File.expand_path('../..', __dir__)
WRITE_DIR = File.join(ROOT_DIR, 'test', 'write')
INPUT_DIR = File.join(WRITE_DIR, 'input_excels')
LOG_DIR = File.join(WRITE_DIR, 'Write_test_log')

FileUtils.mkdir_p(LOG_DIR)
timestamp = Time.now.strftime('%Y%m%d_%H%M%S')
log_path = File.join(LOG_DIR, "Write_test_#{timestamp}.log")
log_lines = []

def write_log_line(log_lines, message)
  puts message
  log_lines << message
end

def capture_stdout
  original_stdout = $stdout
  output = StringIO.new
  $stdout = output
  yield
  output.string
ensure
  $stdout = original_stdout
end

def run_one_command(command, command_line)
  command.set_scripted_inputs([command_line])

  output = capture_stdout do
    command.read_command
    command.parse_command
    command.execute_command
  end

  output
end

def run_read(command, calendar_path:, event_path:, lecture_path:)
  command.set_scripted_inputs([
    'read',
    calendar_path.to_s,
    event_path.to_s,
    lecture_path.to_s
  ])

  capture_stdout do
    command.read_command
    command.parse_command
    command.execute_command
  end
end

def expected_file_names(year)
  [
    "#{year}年度_1学期_講義室予約表.xlsx",
    "#{year}年度_2学期_講義室予約表.xlsx",
    "#{year}年度_3学期_講義室予約表.xlsx",
    "#{year}年度_4学期_講義室予約表.xlsx",
    "#{year}年度_夏季_講義室予約表.xlsx",
    "#{year}年度_春季_講義室予約表.xlsx"
  ]
end

def ensure_output_target_room(command)
  rooms = command.room_info.rooms

  if rooms.empty?
    command.room_info.add_user_defined_room(['第１１講義室'])
    rooms = command.room_info.rooms
  end

  command.room_info.set_output_target_rooms([rooms.first])
end

def with_clean_output_dir(path)
  FileUtils.rm_rf(path)
  yield
ensure
  FileUtils.rm_rf(path)
end

def include_all?(text, needles)
  needles.all? { |needle| text.include?(needle) }
end

def sanitize_for_dirname(text)
  text.to_s.gsub(/[^0-9A-Za-z_-]+/, '_').gsub(/_+/, '_').sub(/^_/, '').sub(/_$/, '')
end

calendar_path = File.join(INPUT_DIR, 'Calender.xlsx')
event_path = File.join(INPUT_DIR, 'event.xlsx')
lecture_path = File.join(INPUT_DIR, 'lecture.xlsx')

required_paths = [calendar_path, event_path, lecture_path]
missing_paths = required_paths.reject { |path| File.file?(path) }
unless missing_paths.empty?
  write_log_line(log_lines, '[Error] 入力Excelが見つかりません')
  missing_paths.each { |path| write_log_line(log_lines, " - #{path}") }
  File.write(log_path, log_lines.join("\n") + "\n")
  exit 1
end

year = 2026
output_dir_name = "#{year}年度講義室予約表"
output_dir_path = File.join(ROOT_DIR, output_dir_name)
expected_files = expected_file_names(year)
artifact_root = File.join(WRITE_DIR, 'output_artifacts', timestamp)
FileUtils.mkdir_p(artifact_root)

started_at = Time.now.iso8601
intro_lines = []
intro_lines << "Write システムテスト開始: #{started_at}"
intro_lines << "Calendar source: #{Pathname.new(calendar_path).relative_path_from(Pathname.new(ROOT_DIR))}"
intro_lines << "Event source: #{Pathname.new(event_path).relative_path_from(Pathname.new(ROOT_DIR))}"
intro_lines << "Lecture source: #{Pathname.new(lecture_path).relative_path_from(Pathname.new(ROOT_DIR))}"

cases = [
  {
    id: '3-001',
    title: '正常: 学年暦/予約/時間割を読み込んで予約表を作成',
    run: lambda do
      command = Command.new
      read_output = run_read(
        command,
        calendar_path: calendar_path,
        event_path: event_path,
        lecture_path: lecture_path
      )
      ensure_output_target_room(command)
      write_output = run_one_command(command, 'write')
      [read_output, write_output]
    end,
    verify: lambda do |read_output, write_output|
      checks = []
      checks << include_all?(read_output, [
        '学年暦データファイルは正常に読み込まれました',
        '予約データファイルは正常に読み込まれました',
        '時間割データファイルは正常に読み込まれました'
      ])
      checks << write_output.include?("ディレクトリ '#{output_dir_name}' に予約表を作成しました:")
      checks << expected_files.all? { |name| write_output.include?("'#{name}'") }
      checks << write_output.include?('予定の重複が 5 件見つかりました:')
      checks << expected_files.all? { |name| File.file?(File.join(output_dir_path, name)) }

      workbook = RubyXL::Parser.parse(File.join(output_dir_path, expected_files.first))
      note_text = workbook[0][1][11]&.value.to_s
      checks << note_text.include?('重複している予定は赤色で表示されます')

      checks.all?
    end
  },
  {
    id: '3-002',
    title: '正常: 予約データなしで予約表を作成',
    run: lambda do
      command = Command.new
      read_output = run_read(
        command,
        calendar_path: calendar_path,
        event_path: '',
        lecture_path: lecture_path
      )
      ensure_output_target_room(command)
      write_output = run_one_command(command, 'write')
      [read_output, write_output]
    end,
    verify: lambda do |read_output, write_output|
      include_all?(read_output, [
        '学年暦データファイルは正常に読み込まれました',
        '予約データファイルは未入力です',
        '時間割データファイルは正常に読み込まれました'
      ]) &&
        write_output.include?("ディレクトリ '#{output_dir_name}' に予約表を作成しました:") &&
        expected_files.all? { |name| File.file?(File.join(output_dir_path, name)) } &&
        !write_output.include?('[Error]')
    end
  },
  {
    id: '3-003',
    title: '正常: 時間割データなしで予約表を作成',
    run: lambda do
      command = Command.new
      read_output = run_read(
        command,
        calendar_path: calendar_path,
        event_path: event_path,
        lecture_path: ''
      )
      ensure_output_target_room(command)
      write_output = run_one_command(command, 'write')
      [read_output, write_output]
    end,
    verify: lambda do |read_output, write_output|
      include_all?(read_output, [
        '学年暦データファイルは正常に読み込まれました',
        '予約データファイルは正常に読み込まれました',
        '時間割データファイルは未入力です'
      ]) &&
        write_output.include?("ディレクトリ '#{output_dir_name}' に予約表を作成しました:") &&
        expected_files.all? { |name| File.file?(File.join(output_dir_path, name)) } &&
        !write_output.include?('[Error]')
    end
  },
  {
    id: '3-004',
    title: '正常: 予約/時間割データなしで予約表を作成',
    run: lambda do
      command = Command.new
      read_output = run_read(
        command,
        calendar_path: calendar_path,
        event_path: '',
        lecture_path: ''
      )
      ensure_output_target_room(command)
      write_output = run_one_command(command, 'write')
      [read_output, write_output]
    end,
    verify: lambda do |read_output, write_output|
      include_all?(read_output, [
        '学年暦データファイルは正常に読み込まれました',
        '予約データファイルは未入力です',
        '時間割データファイルは未入力です'
      ]) &&
        write_output.include?("ディレクトリ '#{output_dir_name}' に予約表を作成しました:") &&
        expected_files.all? { |name| File.file?(File.join(output_dir_path, name)) } &&
        !write_output.include?('[Error]')
    end
  },
  {
    id: '3-005',
    title: '異常: writeコマンドに不要引数を指定',
    run: lambda do
      command = Command.new
      output = run_one_command(command, 'write arg')
      ['', output]
    end,
    verify: lambda do |_read_output, write_output|
      write_output.include?(' [Error] write コマンドには引数が不要です')
    end
  },
  {
    id: '3-006',
    title: '異常: 学年暦未読込でwriteを実行',
    run: lambda do
      command = Command.new
      output = run_one_command(command, 'write')
      ['', output]
    end,
    verify: lambda do |_read_output, write_output|
      write_output.include?(' [Error] 先に read コマンドで学年暦データを読み込んでください')
    end
  },
  {
    id: '3-007',
    title: '異常: 出力対象講義室なしでwriteを実行',
    run: lambda do
      command = Command.new
      read_output = run_read(
        command,
        calendar_path: calendar_path,
        event_path: '',
        lecture_path: ''
      )
      write_output = run_one_command(command, 'write')
      [read_output, write_output]
    end,
    verify: lambda do |read_output, write_output|
      read_output.include?('学年暦データファイルは正常に読み込まれました') &&
        write_output.include?(' [Error] 出力対象講義室がありません') &&
        write_output.include?('register/select コマンドで設定してください')
    end
  }
]

results = []

cases.each do |test_case|
  FileUtils.rm_rf(output_dir_path)

  artifact_dir = nil

  begin
    read_output, write_output = test_case[:run].call
    passed = test_case[:verify].call(read_output, write_output)

    if File.directory?(output_dir_path)
      case_label = sanitize_for_dirname("#{test_case[:id]}_#{test_case[:title]}")
      artifact_dir = File.join(artifact_root, case_label)
      FileUtils.rm_rf(artifact_dir)
      FileUtils.mv(output_dir_path, artifact_dir)
    end

    results << {
      id: test_case[:id],
      title: test_case[:title],
      passed: passed,
      read_output: read_output,
      write_output: write_output,
      artifact_dir: artifact_dir
    }
  rescue StandardError => e
    if File.directory?(output_dir_path)
      case_label = sanitize_for_dirname("#{test_case[:id]}_#{test_case[:title]}_exception")
      artifact_dir = File.join(artifact_root, case_label)
      FileUtils.rm_rf(artifact_dir)
      FileUtils.mv(output_dir_path, artifact_dir)
    end

    results << {
      id: test_case[:id],
      title: test_case[:title],
      passed: false,
      read_output: '',
      write_output: "[Exception] #{e.class}: #{e.message}\n#{e.backtrace&.first}",
      artifact_dir: artifact_dir
    }
  ensure
    FileUtils.rm_rf(output_dir_path)
  end
end

pass_count = results.count { |result| result[:passed] }
fail_count = results.count { |result| !result[:passed] }
failed_ids = results
             .reject { |result| result[:passed] }
             .map { |result| result[:id] }
             .sort_by { |id| id.split('-').last.to_i }

summary_lines = []
summary_lines << '結果サマリ'
summary_lines << "実行ケース数: #{cases.size}"
summary_lines << "成功: #{pass_count}"
summary_lines << "失敗: #{fail_count}"
summary_lines << "失敗項番: #{failed_ids.join(', ')}"
summary_lines << "生成ファイル保存先: #{Pathname.new(artifact_root).relative_path_from(Pathname.new(ROOT_DIR))}"

detail_lines = []
detail_lines << '各項目の詳細'

sorted_results = results.sort_by { |result| result[:id].split('-').last.to_i }
sorted_results.each do |result|
  status_text = result[:passed] ? '成功' : '失敗'
  detail_lines << "#{result[:id]} #{result[:title]}: [#{status_text}]"
  if result[:artifact_dir]
    relative_artifact_dir = Pathname.new(result[:artifact_dir]).relative_path_from(Pathname.new(ROOT_DIR))
    detail_lines << "  - 生成ファイル: #{relative_artifact_dir}"
  else
    detail_lines << '  - 生成ファイル: なし'
  end

  merged_output = [result[:read_output], result[:write_output]].join("\n")
  output_preview = merged_output.lines.map(&:chomp).last(8)
  output_preview.each do |line|
    detail_lines << "    #{line}"
  end
end

# ログファイルは開始情報→サマリ→詳細の順に配置
final_log_lines = intro_lines + [''] + summary_lines + [''] + detail_lines

# コンソールは詳細を先に出して、サマリを最後に配置
console_lines = intro_lines + detail_lines + summary_lines
console_lines.each { |line| puts line }

File.write(log_path, final_log_lines.join("\n") + "\n")
puts "ログ出力: #{log_path}"

exit 1 if fail_count.positive?
