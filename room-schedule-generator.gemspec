Gem::Specification.new do |spec|
  spec.name          = "room-schedule-generator"
  spec.version       = "1.0"
  spec.summary       = "Room management CLI App"

  spec.authors       = ["Hisaki Teraoka"]

  spec.files = Dir[
    "lib/**/*.rb",
    "exe/*",
    "README.md"
  ]

  spec.bindir = "exe"
  spec.require_paths = ["lib"]

  spec.executables = ["room-schedule-generator"]

  # プログラムが使用する gemの依存関係を追加
  spec.add_dependency "fiddle", '~> 1.1'
  spec.add_dependency "tty-prompt", '~> 0.23'
  spec.add_dependency "tty-cursor", '~> 0.5'
  spec.add_dependency "reline", '~> 0.6'
  spec.add_dependency "date", '~> 3.5'
  spec.add_dependency "rubyXL", '~> 3.4'
  spec.add_dependency 'csv', '~> 3.3'
end