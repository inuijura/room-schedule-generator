require 'date'

require_relative '../../lib/calendar'

# 引数で指定された学年暦ファイルのパスを取得
CALENDAR_PATH = ARGV[0]
TEST_TYPE = ARGV[1]

if CALENDAR_PATH.nil? || CALENDAR_PATH.empty?
  puts '学年暦ファイルのパスを指定してください。'
  exit 1
end

if TEST_TYPE.nil? || TEST_TYPE.empty?
  puts 'テストタイプを指定してください。'
  exit 1
end

def assert_true!(condition, message)
  raise message unless condition
end

def assert_file_exists!(path, label)
  raise "#{label} が見つかりません: #{path}" unless File.exist?(path)
end

assert_file_exists!(CALENDAR_PATH, '学年暦ファイル')

completed = false
msg = nil
calendar = Calendar.new
begin
  calendar.parse(CALENDAR_PATH)
  completed = true

  calendar.print_term_periods
  calendar.print_calendar
rescue => e
  completed = false
  msg = "エラー (#{e.full_message})"
  puts msg
end

# if calendar.is_valid?
#   puts "[OK] 異常が検出できませんでした．"
# else 
#   puts msg
# end

# if TEST_TYPE == '正常'
#   if completed && calendar.is_valid?
#     puts "[RESULT] 結果整合 passed"
#   else 
#     puts "[RESULT] 結果不整合 failed"
#   end
# end

# if TEST_TYPE == '異常'
#   if completed && calendar.is_valid?
#     puts "[RESULT] 結果不整合 failed"
#   else 
#     puts "[RESULT] 結果整合 passed"
#   end
# end
