# 講義室管理情報システム

## 1. システムの概要

本システムは、大学の学年暦、講義室の予約データ、時間割データを読み込み、講義室ごとの利用予定を確認・出力するためのコマンドラインツールです。

主に次のことができます。

- 学年暦（Excel）から、授業日や各学期の期間を読み込む
- 予約データおよび時間割データ（Excel）から、講義室の利用予定を読み込む
- 管理対象とする講義室をCSVから登録する
- 読み込んだデータに含まれる講義室から、出力対象を選択する
- 指定した学期の講義室利用状況をターミナル上で確認する
- 学期ごとの講義室予約表をExcel形式で出力する
- 同じ講義室・日時に予定が重複している箇所を確認する

本システムは対話形式で動作します。起動後にコマンドを入力し、表示される案内に従ってファイルや対象講義室、学期を指定します。

## 2. システムのインストール方法

### 2.1 動作環境

以下の環境で動作を確認しています。

| OS | Ruby | RubyGems |
| --- | --- | --- |
| Windows 11 Pro | 4.0.5 | 4.0.10 |
| Ubuntu 24.04 LTS | 3.2.3 | 3.4.20 |

### 2.2 インストール

#### Windows（Gemパッケージ版）

[RubyInstaller](https://rubyinstaller.org/)から、MSYS2環境を含めてRubyをインストールします。その後、ターミナルで次のコマンドを実行します。

```sh
curl -fLO https://github.com/inuijura/room-schedule-generator/releases/download/v1.0/room-schedule-generator-1.0.gem
gem install room-schedule-generator-1.0.gem
del room-schedule-generator-1.0.gem
```

Rubyをインストールせずに利用する場合は、実行ファイル版をダウンロードします。

```sh
curl -fLO https://github.com/inuijura/room-schedule-generator/releases/download/v1.0/room-schedule-generator.exe
```

#### Ubuntu

必要なソフトウェアをインストールした後、インストールスクリプトを実行します。

```sh
sudo apt install ruby ruby-dev build-essential
curl -fsSL https://raw.githubusercontent.com/inuijura/room-schedule-generator/v1.0/script/install.sh | bash
```

インストール後はターミナルを再起動してください。

### 2.3 起動確認

Gemパッケージ版またはUbuntu版では、次のコマンドを実行します。

```sh
room-schedule-generator
```

Windowsの実行ファイル版では、ダウンロード先に合わせて実行ファイルのパスを指定します。

```powershell
.\<path>\room-schedule-generator.exe
```

`>` が表示されれば起動完了です。

## 3. システムの簡単な使い方

### 3.1 起動

Gemパッケージ版またはUbuntu版では、任意のディレクトリで次のコマンドを実行します。

```sh
room-schedule-generator
```

Windowsの実行ファイル版では、実行ファイルのパスを指定して起動します。起動すると `>` が表示され、コマンドを入力できる状態になります。利用できるコマンドは次のとおりです。

### 3.2 コマンド一覧

| コマンド | 説明 |
| --- | --- |
| `read` | 学年暦、予約、時間割の各データファイルを読み込みます。 |
| `register <CSVファイル>` | CSVに記載された講義室を管理対象として登録します。 |
| `select` | 登録済みの講義室から出力対象を選択します。 |
| `print` | 選択した学期の講義室利用状況をターミナルに表示します。 |
| `write` | 学期ごとの講義室予約表をExcel形式で出力します。 |
| `quit` | システムを終了します。 |

コマンド名とファイルパスは、Tabキーで補完できます。

### 3.3 基本的な操作の流れ

ここでは、データの読み込みから講義室予約表の出力までの基本的な操作を順に説明します。画面に `>` が表示されている状態で、以下のコマンドを入力してください。

#### 1. データを読み込む

`read` コマンドを入力し、画面の案内に従って次のファイルを指定します。

1. 学年暦データファイル（必須）
2. 予約データファイル（任意）
3. 時間割データファイル（任意）

```text
> read
 1. 学年暦データファイル: excel_example/2026年度岡山大学授業日程計画.xlsx
 2. 予約データファイル: excel_example/予約データ案_SDM.xlsx
 3. 時間割データファイル: excel_example/時間割データ案_SDM.xlsx
```

予約データまたは時間割データを使用しない場合は、ファイルパスを入力せずに Enter キーを押します。

各ファイルは対象年度に対応したものを用意し、見出し名やセル配置を変更しないでください。ファイルパスには絶対パスと相対パスのどちらも使用できます。

#### 2. 管理対象の講義室を設定する

CSVに記載された講義室を管理対象として登録する場合は、`register` コマンドにCSVファイルのパスを指定します。

```text
> register excel_example/ユーザ定義講義室.csv
```

CSV内の空でない各セルが講義室名として読み込まれ、登録した講義室がそのまま出力対象になります。

また、`read` または `register` で登録された講義室の一覧から出力対象を選び直す場合は、`select` コマンドを使用します。

```text
> select
```

表示される一覧で、出力したい講義室を選択してください。講義室には登録元に応じて次のラベルが表示されます。

- `ALL`: 管理対象講義室CSVと、予約・時間割データの両方に含まれる講義室
- `USER`: 管理対象講義室CSVにのみ含まれる講義室
- `UNIV`: 予約・時間割データにのみ含まれる講義室

#### 3. 講義室の利用状況を確認する

`print` コマンドを入力し、表示対象の学期を選択します。選択できる期間は、1〜4学期、夏季、春季です。

```text
> print
```

選択した学期について、対象講義室の利用状況がターミナル上に表示されます。

#### 4. 講義室予約表を出力する

`write` コマンドを入力すると、読み込んだ年度の学期ごとに講義室予約表がExcel形式で生成されます。

```text
> write
```

たとえば2026年度のデータを読み込んだ場合、カレントディレクトリに `2026年度講義室予約表` ディレクトリが作成され、その中に次のようなファイルが出力されます。

```text
2026年度_1学期_講義室予約表.xlsx
2026年度_2学期_講義室予約表.xlsx
...
```

予定が重複している箇所は、出力ファイル内で赤色で表示されます。

## ライセンスと著作権
このプロジェクトは [MIT License](LICENSE) のもとで公開されています。

Copyright (c) 2026 Software Development Methodology Group1

 Authors:
 - Jura Inui (@inuijura)
 - Tatsuya Kaji (@tattya-hue), 
 - Hiroto Ohtsuki (@o-brothers-hiroto) 
 - Hisaki Teraoka (@teraoka-h)
 - Shuichi Mizoguchi (@MS1208)
 - Kota Takahashi (@Kota-Takahashi7)
