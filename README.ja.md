PyOfficeGrep
===

![Software Version](http://img.shields.io/badge/Version-v0.0.2-green.svg?style=flat)
![Python Version](http://img.shields.io/badge/Python-3.x-blue.svg?style=flat)
[![MIT License](http://img.shields.io/badge/license-MIT-blue.svg?style=flat)](LICENSE)

[English Page](./README.md)

## 概要
Office ファイルを対象に grep 検索を実行する。

## バージョン
v0.0.2

## 要件
### `office_grep.py` を使用する場合
- Python 3.x
- PyWin32
- colorama

[Anaconda3](https://www.anaconda.com/) に全て含まれる。

### `office_grep.exe` を使用する場合
特になし

※ `to_exe.bat` を使用して exe ファイルを生成する場合は `PyInstaller` が必要。

## ライセンス
[MIT License](./LICENSE)

## 仕様・制限事項
- Windows 上でのみ動作可能。
- 現在のところ、Excel と Word ファイル (xls, xlsx, xlsm, doc, docx, docm) にのみ対応。
    - いずれ PowerPoint ファイル (ppt, pptx, pptm) にも対応する予定だが、必要に迫られていないため後回しとした。
- Excel ファイルの使用セルの検索を高速化するために, 少し工夫を加えている。

## 使用方法
`office_grep.py` もしくは `office_grep.exe` を引数付きで実行する。

```
$ python office_grep.py -h
usage: office_grep.py [-h] [-v] [--type TYPE] [--word WORD]
                      [--recursive RECURSIVE] [--ignorecase IGNORECASE]
                      [--regex REGEX] [--parallel PARALLEL]
                      query dirpath

Run a grep search on the Office files.

positional arguments:
  query                 Search query
  dirpath               Target directory path

optional arguments:
  -h, --help            show this help message and exit
  -v, --version         show program's version number and exit
  --type TYPE           Target file types (The following letter combinations:
                        "E":Excel, "W":Word, "P":PowerPoint)
  --word WORD           Match whole words only
  --recursive RECURSIVE
                        Search the directory recursively
  --ignorecase IGNORECASE
                        Case insensitive
  --regex REGEX         Regex
  --parallel PARALLEL   Maximum number of parallel processing
```

### 必須パラメータ

| パラメータ | 説明                   | 値の書式   |
|------------|------------------------|------------|
| query      | 検索クエリ             | 文字列     |
| dirpath    | 対象のディレクトリパス | パス文字列 |

### Optional Parameters

| パラメータ | 説明                       | 値の書式                                                  | 初期値 |
|------------|----------------------------|-----------------------------------------------------------|--------|
| type       | 対象とするファイル種類     | 以下の文字の組み合わせ:<br/>E:Excel, W:Word, P:PowerPoint | EWP    |
| word       | 単語単位で検索             | 真偽値 (True or False)                                    | False  |
| recursive  | ディレクトリを再帰的に検索 | 真偽値 (True or False)                                    | True   |
| ignorecase | 大文字小文字を区別しない   | 真偽値 (True or False)                                    | False  |
| regex      | 正規表現                   | 真偽値 (True or False)                                    | True   |
| parallel   | 並列処理の最大数           | 整数値                                                    | 1      |

注意:  
- parallel
    - 1 を指定すると並列処理が無効となる。
    - 並列処理を有効にすると、Ctrl+C で中断することができなくなる。
    - 現状、並列処理を有効にしても特に速度アップしないため、1 を推奨する。

#### 値の採用方法
以下の優先順位で値が採用される。
1. コマンドライン引数
2. 設定ファイル (カレントディレクトリに置かれた `setting.ini`)
3. 初期値


## 検索結果の表示内容

![result](https://user-images.githubusercontent.com/64964079/86204256-0028ee80-bba2-11ea-8093-1cb48b20acb9.png)

### ファイルパスの行
現在のファイル番号と総ファイル数とともに、ファイルパスが表示される。  
先頭の 1 文字はファイル種類を表す。 (E:Excel, W:Word, P:PowerPoint)

### 検索結果の行
そのファイル中に検索結果が存在する場合に表示される。

前半は情報ブロックで、検索結果の位置を表す。  
後半はテキストブロックで、前後の文章とともに検索結果が表示される。検索により見つかったトークンはハイライトされる。

以下に、ファイル種類別の情報ブロックの内容を記す。

#### Excel
種別 (括弧の前):

| 種別    |
|---------|
| Cell    |
| Shape   |
| Comment |

詳細 (括弧の中):

| 詳細    | 説明               |
|---------|--------------------|
| Sheet   | シート名           |
| Address | セルのアドレス     |
| Name    | オブジェクトの名称 |

#### Word
種別 (括弧の前):

| 種別    |
|---------|
| Text    |
| Table   |
| Shape   |
| Comment |

詳細 (括弧の中):

| 詳細 | 説明                              |
|------|-----------------------------------|
| Page | ページ番号                        |
| Line | ページ内の行番号                  |
| Cell | 表内のセルの位置 (行番号, 列番号) |
| Name | オブジェクトの名称                |

#### PowerPoint
WIP
