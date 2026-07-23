require 'fileutils'
require 'stringio'
require 'time'

require_relative '../../lib/command'

ROOT_DIR = File.expand_path('../..', __dir__)
READ_DIR = File.join(ROOT_DIR, 'test', 'read')
LOG_DIR = File.join(READ_DIR, 'Read_test_log')

CALENDAR_DIR = File.join(READ_DIR, 'Calendar_data')
EVENT_DIR = File.join(READ_DIR, 'Event_data')
LECTURE_DIR = File.join(READ_DIR, 'Lecture_data')

FileUtils.mkdir_p(LOG_DIR)

def fixture(dir, *parts)
  path = File.join(dir, *parts)
  return path if File.file?(path)

  raise "テストデータが見つかりません: #{path}"
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

def execute_read(calendar:, event: '', lecture: '')
  command = Command.new
  command.set_scripted_inputs(['read', calendar, event, lecture])

  output = capture_stdout do
    command.read_command
    command.parse_command
    command.execute_command
  end

  [command, output]
end

def make_file_inaccessible(source, destination)
  FileUtils.cp(source, destination)
  File.chmod(0o000, destination)
  destination
end

def make_directory_inaccessible(source, parent, filename)
  directory = File.join(parent, "inaccessible_#{File.basename(filename, '.*')}")
  FileUtils.mkdir_p(directory)
  destination = File.join(directory, filename)
  FileUtils.cp(source, destination)
  File.chmod(0o000, directory)
  destination
end

def assert_case(test_case)
  command, output = execute_read(
    calendar: test_case.fetch(:calendar),
    event: test_case.fetch(:event, ''),
    lecture: test_case.fetch(:lecture, '')
  )

  failures = []
  search_offset = 0
  test_case.fetch(:expected).each do |pattern|
    match = pattern.match(output, search_offset)
    if match
      search_offset = match.end(0)
    else
      failures << "出力なし、または順序不正: #{pattern.inspect}"
    end
  end
  test_case.fetch(:unexpected, []).each do |pattern|
    failures << "想定外の出力: #{pattern.inspect}" if pattern.match?(output)
  end

  if test_case.fetch(:calendar_loaded, true) != !command.calendar.nil?
    failures << '学年暦の読込状態が想定と異なる'
  end

  [failures, output]
end

calendar_normal = fixture(CALENDAR_DIR, '正常_学年暦データ_通常.xlsx')
calendar_invalid = fixture(
  CALENDAR_DIR,
  '特定のカレンダ情報',
  '異常_特定のカレンダ情報_欠落_月・日.xlsx'
)
calendar_csv = fixture(CALENDAR_DIR, '異常_学年暦データ_不適_拡張子.csv')

event_normal = fixture(EVENT_DIR, '正常_予約データ_通常.xlsx')
event_header_only = fixture(EVENT_DIR, '正常_カラム有_データ空.xlsx')
event_no_header = fixture(EVENT_DIR, '異常_カラム無_空ファイル.xlsx')
event_invalid = fixture(EVENT_DIR, '異常_予約データ_不適_項目名ミス(data).xlsx')
event_csv = fixture(EVENT_DIR, '異常_予約データ_不適な拡張子.csv')

lecture_normal = fixture(LECTURE_DIR, '正常_時間割データ_通常.xlsx')
lecture_header_only = fixture(LECTURE_DIR, '正常_カラム有_データ空.xlsx')
lecture_no_header = fixture(LECTURE_DIR, '異常_カラム無_空ファイル.xlsx')
lecture_invalid = fixture(LECTURE_DIR, '異常_時間割データ_不適_項目名ミス(weekday).xlsx')
lecture_csv = fixture(LECTURE_DIR, '異常_時間割データ_不適切な拡張子.csv')

timestamp = Time.now.strftime('%Y%m%d_%H%M%S')
temporary_dir = File.join(ROOT_DIR, 'tmp', "read_integration_test_#{timestamp}")
FileUtils.mkdir_p(temporary_dir)

event_inaccessible_directory = make_directory_inaccessible(event_normal, temporary_dir, 'event.xlsx')
lecture_inaccessible_directory = make_directory_inaccessible(lecture_normal, temporary_dir, 'lecture.xlsx')
calendar_inaccessible_directory = make_directory_inaccessible(calendar_normal, temporary_dir, 'calendar.xlsx')
event_inaccessible_file = make_file_inaccessible(event_normal, File.join(temporary_dir, 'event_file.xlsx'))
lecture_inaccessible_file = make_file_inaccessible(lecture_normal, File.join(temporary_dir, 'lecture_file.xlsx'))
calendar_inaccessible_file = make_file_inaccessible(calendar_normal, File.join(temporary_dir, 'calendar_file.xlsx'))

at_exit do
  Dir.glob(File.join(temporary_dir, 'inaccessible_*')).each do |directory|
    File.chmod(0o755, directory) if File.directory?(directory)
  rescue StandardError
    nil
  end
  [event_inaccessible_file, lecture_inaccessible_file, calendar_inaccessible_file].each do |path|
    File.chmod(0o644, path) if path && File.file?(path)
  rescue StandardError
    nil
  end
  FileUtils.rm_rf(temporary_dir) if Dir.exist?(temporary_dir)
end

permission_fixtures = [
  event_inaccessible_directory, lecture_inaccessible_directory, calendar_inaccessible_directory,
  event_inaccessible_file, lecture_inaccessible_file, calendar_inaccessible_file
]
if permission_fixtures.any? { |path| File.readable?(path) }
  raise '読み取り権限テスト用ファイルへのアクセスを遮断できませんでした'
end

info_calendar = /\[Info\] 学年暦データファイルは正常に読み込まれました/
info_event = /\[Info\] 予約データファイルは正常に読み込まれました/
info_lecture = /\[Info\] 時間割データファイルは正常に読み込まれました/

cases = [
  {
    id: '5-001', title: '3ファイルを正常に読み込む',
    calendar: calendar_normal, event: event_normal, lecture: lecture_normal,
    expected: [info_calendar, info_event, info_lecture], unexpected: [/\[Error\]/]
  },
  {
    id: '5-002', title: '予約データを未入力にする',
    calendar: calendar_normal, lecture: lecture_normal,
    expected: [info_calendar, /予約データファイルは未入力です/, info_lecture]
  },
  {
    id: '5-003', title: '予約データに存在しないパスを指定する',
    calendar: calendar_normal, event: File.join(temporary_dir, 'not_found.xlsx'), lecture: lecture_normal,
    expected: [info_calendar, /\[Error\] 予約データファイル '.+' がありません/, info_lecture]
  },
  {
    id: '5-004', title: '見出しのみの予約データを読み込む',
    calendar: calendar_normal, event: event_header_only, lecture: lecture_normal,
    expected: [info_calendar, info_event, info_lecture]
  },
  {
    id: '5-005', title: '見出しのない予約データを読み込む',
    calendar: calendar_normal, event: event_no_header, lecture: lecture_normal,
    expected: [info_calendar, /\[Error\] 予約データファイル '.+' のデータ形式が不正です/, info_lecture]
  },
  {
    id: '5-006', title: '形式不正の予約データを読み込む',
    calendar: calendar_normal, event: event_invalid, lecture: lecture_normal,
    expected: [info_calendar, /\[Error\] 予約データファイル '.+' のデータ形式が不正です/, info_lecture]
  },
  {
    id: '5-007', title: 'xlsx以外の予約データを読み込む',
    calendar: calendar_normal, event: event_csv, lecture: lecture_normal,
    expected: [%r{\[Error\] 予約データファイルの読み込みに失敗しました.*期待する拡張子は '.xlsx' ですが，'.csv' が入力されました}m]
  },
  {
    id: '5-008', title: '権限のないディレクトリ内の予約データを読み込む',
    calendar: calendar_normal, event: event_inaccessible_directory, lecture: lecture_normal,
    expected: [info_calendar, /\[Error\] 予約データファイル '.+' の読み取り権限がありません/, info_lecture]
  },
  {
    id: '5-009', title: '読み取り権限のない予約データを読み込む',
    calendar: calendar_normal, event: event_inaccessible_file, lecture: lecture_normal,
    expected: [info_calendar, /\[Error\] 予約データファイル '.+' の読み取り権限がありません/, info_lecture]
  },
  {
    id: '5-010', title: '時間割データを未入力にする',
    calendar: calendar_normal, event: event_normal,
    expected: [info_calendar, info_event, /時間割データファイルは未入力です/]
  },
  {
    id: '5-011', title: '時間割データに存在しないパスを指定する',
    calendar: calendar_normal, event: event_normal, lecture: File.join(temporary_dir, 'not_found.xlsx'),
    expected: [info_calendar, info_event, /\[Error\] 時間割データファイル '.+' がありません/]
  },
  {
    id: '5-012', title: '見出しのみの時間割データを読み込む',
    calendar: calendar_normal, event: event_normal, lecture: lecture_header_only,
    expected: [info_calendar, info_event, info_lecture]
  },
  {
    id: '5-013', title: '見出しのない時間割データを読み込む',
    calendar: calendar_normal, event: event_normal, lecture: lecture_no_header,
    expected: [info_calendar, info_event, /\[Error\] 時間割データファイル '.+' のデータ形式が不正です/]
  },
  {
    id: '5-014', title: '形式不正の時間割データを読み込む',
    calendar: calendar_normal, event: event_normal, lecture: lecture_invalid,
    expected: [info_calendar, info_event, /\[Error\] 時間割データファイル '.+' のデータ形式が不正です/]
  },
  {
    id: '5-015', title: 'xlsx以外の時間割データを読み込む',
    calendar: calendar_normal, event: event_normal, lecture: lecture_csv,
    expected: [%r{\[Error\] 時間割データファイルの読み込みに失敗しました.*期待する拡張子は '.xlsx' ですが，'.csv' が入力されました}m]
  },
  {
    id: '5-016', title: '権限のないディレクトリ内の時間割データを読み込む',
    calendar: calendar_normal, event: event_normal, lecture: lecture_inaccessible_directory,
    expected: [info_calendar, info_event, /\[Error\] 時間割データファイル '.+' の読み取り権限がありません/]
  },
  {
    id: '5-017', title: '読み取り権限のない時間割データを読み込む',
    calendar: calendar_normal, event: event_normal, lecture: lecture_inaccessible_file,
    expected: [info_calendar, info_event, /\[Error\] 時間割データファイル '.+' の読み取り権限がありません/]
  },
  {
    id: '5-018', title: '学年暦に存在しないパスを指定する',
    calendar: File.join(temporary_dir, 'not_found.xlsx'), calendar_loaded: false,
    expected: [/\[Error\] 学年暦データファイル '.+' がありません/, /学年暦データファイルは必須です/]
  },
  {
    id: '5-019', title: '形式不正の学年暦データを読み込む',
    calendar: calendar_invalid, calendar_loaded: false,
    expected: [/\[Error\] '.+' の学年暦データ形式が不正です/]
  },
  {
    id: '5-020', title: 'xlsx以外の学年暦データを読み込む',
    calendar: calendar_csv, calendar_loaded: false,
    expected: [%r{\[Error\] 学年暦データファイルの読み込みに失敗しました.*期待する拡張子は '.xlsx' ですが，'.csv' が入力されました}m]
  },
  {
    id: '5-021', title: '権限のないディレクトリ内の学年暦データを読み込む',
    calendar: calendar_inaccessible_directory, calendar_loaded: false,
    expected: [
      /\[Error\] 学年暦データファイル '.+' の読み取り権限がありません/,
      /ファイルのアクセス権限を確認してください/
    ]
  },
  {
    id: '5-022', title: '読み取り権限のない学年暦データを読み込む',
    calendar: calendar_inaccessible_file, calendar_loaded: false,
    expected: [
      /\[Error\] 学年暦データファイル '.+' の読み取り権限がありません/,
      /ファイルのアクセス権限を確認してください/
    ]
  }
]

results = cases.map do |test_case|
  failures, output = assert_case(test_case)
  status = failures.empty? ? '成功' : '失敗'
  puts "#{test_case[:id]} #{test_case[:title]}: [#{status}]"
  failures.each { |failure| puts "  - #{failure}" }
  { test_case: test_case, failures: failures, output: output }
end

failed_results = results.reject { |result| result[:failures].empty? }
summary = [
  "実行ケース数: #{results.size}",
  "成功: #{results.size - failed_results.size}",
  "失敗: #{failed_results.size}",
  "失敗項番: #{failed_results.map { |result| result[:test_case][:id] }.join(', ')}"
]
puts ['', '結果サマリ', *summary]

log_path = File.join(LOG_DIR, "Read_test_#{timestamp}.log")
detail = []
results.each do |result|
  test_case = result[:test_case]
  status = result[:failures].empty? ? '成功' : '失敗'
  detail << "#{test_case[:id]} #{test_case[:title]}: [#{status}]"
  result[:failures].each { |failure| detail << "  - #{failure}" }
  detail << result[:output].rstrip
  detail << ''
end

log = [
  '1. 結合テスト',
  "Read 結合テスト実行日時: #{Time.now.iso8601}",
  '',
  '2. 結果サマリ',
  *summary,
  '',
  '3. 結果詳細',
  *detail
]
File.write(log_path, log.join("\n") + "\n")
puts "ログ出力: #{log_path}"

exit(failed_results.empty? ? 0 : 1)
