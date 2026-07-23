require 'pathname'
require 'fileutils'
require 'time'

require_relative '../../lib/calendar'
require_relative '../../lib/reservation'

ROOT_DIR = File.expand_path('../..', __dir__)
READ_DIR = File.join(ROOT_DIR, 'test', 'read')
LOG_DIR = File.join(READ_DIR, 'Lecture_test_log')

FileUtils.mkdir_p(LOG_DIR)
timestamp = Time.now.strftime('%Y%m%d_%H%M%S')
log_path = File.join(LOG_DIR, "Lecture_test_#{timestamp}.log")
log_lines = []

def write_log_line(log_lines, message)
  puts message
  log_lines << message
end

def sanitize_error_message(message)
  text = message.to_s.dup
  text.gsub!(/file=([^,\)]+)/) { "file=#{File.basename($1)}" }
  text
end

write_log_line(log_lines, "Lecture テストログ開始: #{Time.now.iso8601}")

# Calendar/Calender の両表記を許容して候補ディレクトリを解決
calendar_source_dir = begin
  candidates = [
    File.join(READ_DIR, 'Calendar_date'),
    File.join(READ_DIR, 'Calender_date'),
    File.join(READ_DIR, 'Calendar_data'),
    File.join(READ_DIR, 'Calender_data')
  ]
  candidates.find { |path| File.directory?(path) }
end

if calendar_source_dir.nil?
  write_log_line(log_lines, "Calendar ディレクトリが見つかりません")
  File.write(log_path, log_lines.join("\n") + "\n")
  exit 1
end

calendar_patterns = ['**/*.xlsx', '**/*.xlsm', '**/*.xls']
calendar_candidates = calendar_patterns.flat_map { |pattern| Dir.glob(File.join(calendar_source_dir, pattern)) }
                                      .sort
                                      .uniq

calendar_file = calendar_candidates.find do |path|
  name = File.basename(path)
  name.include?('正常') && name.include?('通常') && name.include?('学年暦')
end

calendar_file ||= calendar_candidates.find do |path|
  name = File.basename(path)
  name.include?('正常') && name.include?('通常')
end

if calendar_file.nil?
  write_log_line(log_lines, "Calendar 正常系通常ファイルが見つかりません")
  File.write(log_path, log_lines.join("\n") + "\n")
  exit 1
end

begin
  calendar = Calendar.new
  calendar.parse(calendar_file)

  unless calendar.is_valid?
    raise "Calendar の解析結果が無効です"
  end

  write_log_line(log_lines, "Calendar 読み込み完了: #{File.basename(calendar_file)}")
rescue => e
  write_log_line(log_lines, "Calendar 読み込みエラー: #{e.class}: #{sanitize_error_message(e.message)}")
  File.write(log_path, log_lines.join("\n") + "\n")
  exit 1
end

# 指示の Lecture_date を優先し、存在しなければ既存の Lecture_data を使う
LECTURE_SOURCE_DIR = begin
  preferred = File.join(READ_DIR, 'Lecture_date')
  fallback = File.join(READ_DIR, 'Lecture_data')
  File.directory?(preferred) ? preferred : fallback
end

unless File.directory?(LECTURE_SOURCE_DIR)
  write_log_line(log_lines, "対象ディレクトリが見つかりません")
  File.write(log_path, log_lines.join("\n") + "\n")
  exit 1
end

excel_files = Dir.glob(File.join(LECTURE_SOURCE_DIR, '**', '*'))
                 .select { |path| File.file?(path) }
                 .sort
                 .uniq

if excel_files.empty?
  write_log_line(log_lines, "対象ファイルが見つかりませんでした")
  File.write(log_path, log_lines.join("\n") + "\n")
  exit 1
end

normal_total = 0
abnormal_total = 0
pass_or_semi_count = 0
failure_count = 0
normal_success = 0
normal_failure = 0
abnormal_semi_success = 0
abnormal_failure = 0
unknown_total = 0

write_log_line(log_lines, "Lecture テスト開始: #{File.basename(LECTURE_SOURCE_DIR)}")

excel_files.each do |file_path|
  file_name = File.basename(file_path)
  completed = false
  message = nil

  begin
    parsed_rows = Lecture.parse_excel_rows(file_path)
    events = Lecture.from_parsed_rows(parsed_rows)
    completed = true
    message = "完了 (rows=#{parsed_rows.size}, events=#{events.size})"
  rescue => e
    completed = false
    message = "エラー (#{e.class}: #{sanitize_error_message(e.message)})"
  end

  if file_name.start_with?('正常')
    normal_total += 1
    if completed
      normal_success += 1
      pass_or_semi_count += 1
      write_log_line(log_lines, "#{file_name}: [成功] #{message}")
    else
      normal_failure += 1
      failure_count += 1
      write_log_line(log_lines, "#{file_name}: [失敗] #{message}")
    end
  elsif file_name.start_with?('異常')
    abnormal_total += 1
    if completed
      abnormal_failure += 1
      failure_count += 1
      write_log_line(log_lines, "#{file_name}: [失敗] #{message}")
    else
      abnormal_semi_success += 1
      pass_or_semi_count += 1
      write_log_line(log_lines, "#{file_name}: [準成功] #{message}")
    end
  else
    unknown_total += 1
    failure_count += 1
    write_log_line(log_lines, "#{file_name}: [失敗] #{message} (ファイル名先頭が 正常/異常 ではありません)")
  end
end

write_log_line(log_lines, "結果サマリ")
write_log_line(log_lines, "総ファイル数: #{excel_files.size}, 正常系ファイル数: #{normal_total}, 異常系ファイル数: #{abnormal_total}")
write_log_line(log_lines, "成功・準成功数: #{pass_or_semi_count}, 失敗数: #{failure_count}")
write_log_line(log_lines, "正常系成功数: #{normal_success}, 正常系失敗数: #{normal_failure}")
write_log_line(log_lines, "異常系準成功数: #{abnormal_semi_success}, 異常系失敗数: #{abnormal_failure}")
write_log_line(log_lines, "未分類ファイル数: #{unknown_total}") if unknown_total.positive?
File.write(log_path, log_lines.join("\n") + "\n")
puts "ログ出力: #{log_path}"
exit 1 if failure_count.positive?
