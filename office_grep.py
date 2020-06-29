
import os
import sys
import datetime
import pathlib
import glob
import argparse
import configparser
import re

import multiprocessing
from concurrent.futures import ProcessPoolExecutor, ThreadPoolExecutor

import win32com.client
import pythoncom
import colorama
from colorama import Fore, Back, Style


#--------------------------
# クラス
#--------------------------
# 設定値
class Setting:
    def __init__(self):
        # 対象のファイルタイプ ("E":Excel / "W":Word / "P":PowerPoint)
        self.type = 'EWP'
        # 単位検索検索
        self.word = False
        # フォルダを再帰的に検索
        self.recursive = True
        # 英大文字小文字を区別しない
        self.ignorecase = False
        # 正規表現
        self.regex = True
        # 並列処理の最大数
        self.parallel = 1

# Office種類
class OfficeType:
    def __init__(self, char, exts, grep_func):
        # ファイルタイプの文字
        self.char = char
        # 拡張子リスト
        self.exts = exts
        # 検索処理関数
        self.grep_func = grep_func


#--------------------------
# 定数
#--------------------------
# デバッグモード
DEBUG = False

# ログの色設定
COLOR_FILE = Fore.LIGHTGREEN_EX
COLOR_HIT_INFO = Fore.CYAN
COLOR_HIT_POS = Fore.YELLOW

#--------------------------
# グローバル
#--------------------------
# 設定値 (load_setting 関数で読み込まれる)
_setting = None

# Office種類リスト
_office_types = []


def main():
    global _setting, _office_types

    # 開始時刻
    debug_print('Start time: {}'.format(datetime.datetime.now()))
    debug_print('')

    # Office種類リスト
    _office_types = [
        OfficeType(char='E', exts=['xls', 'xlsx', 'xlsm'], grep_func=grep_excel),
        OfficeType(char='W', exts=['doc', 'docx', 'docm'], grep_func=grep_word),
        OfficeType(char='P', exts=['ppt', 'pptx', 'pptm'], grep_func=grep_ppoint),
    ]

    # 設定情報を読み込む
    query, dirpath, _setting = load_setting()

    # カレントディレクトリを変更
    os.chdir(dirpath)

    # 対象とするファイルの拡張子リストを作成
    exts = []
    for office_type in _office_types:
        if office_type.char in _setting.type:
            exts += office_type.exts

    # 対象ファイルパスのリストを取得
    target_fpaths = create_fpaths(dirpath, exts)

    # 検索処理を振り分けながら実行
    grep_while_destribute(query, target_fpaths)

    # 終了時刻
    debug_print('End time: {}'.format(datetime.datetime.now()))
    debug_print('')


# 設定情報の読み込み
def load_setting():

    # bool値変換用クロージャ
    def str2bool(s):
        ret = False
        s = s.lower()
        if s == 'true' or s == 'on':
            ret = True
        return ret

    # 生成
    setting = Setting()

    # 設定ファイルから読み込み
    try:
        config = configparser.ConfigParser()
        config.read('setting.ini', encoding='utf_8')
        if 'grep' in config:
            if 'type' in config['grep']:
                setting.type = config['grep']['type']
            if 'word' in config['grep']:
                setting.recursive = str2bool(config['grep']['word'])
            if 'recursive' in config['grep']:
                setting.recursive = str2bool(config['grep']['recursive'])
            if 'ignorecase' in config['grep']:
                setting.ignorecase = str2bool(config['grep']['ignorecase'])
            if 'regex' in config['grep']:
                setting.regex = str2bool(config['grep']['regex'])
            if 'parallel' in config['grep']:
                setting.parallel = int(config['grep']['parallel'])
    except:
        # 読み込み失敗しても処理を継続
        pass

    # コマンドラインオプション
    parser = argparse.ArgumentParser(description='Run a grep search on the Office files.')
    parser.add_argument('query', help='Search query')
    parser.add_argument('dirpath', help='Target directory path')
    parser.add_argument('--type', type=str, help='Target file types (The following letter combinations: "E":Excel, "W":Word, "P":PowerPoint)')
    parser.add_argument('--word', type=str2bool, help='Match whole words only')
    parser.add_argument('--recursive', type=str2bool, help='Search the directory recursively')
    parser.add_argument('--ignorecase', type=str2bool, help='Case insensitive')
    parser.add_argument('--regex', type=str2bool, help='Regex')
    parser.add_argument('--parallel', type=int, help='Maximum number of parallel processing')
    args = parser.parse_args()

    # コマンドラインの設定を反映
    query = args.query
    dirpath = args.dirpath
    if args.type is not None:
        setting.type = args.type
    if args.word is not None:
        setting.word = args.word
    if args.recursive is not None:
        setting.recursive = args.recursive
    if args.ignorecase is not None:
        setting.ignorecase = args.ignorecase
    if args.regex is not None:
        setting.regex = args.regex
    if args.regex is not None:
        setting.parallel = args.parallel

    # 単語単位検索
    if setting.word:
        # 正規表現
        if setting.regex:
            print('Error: You cannot set up "Match whole words only" and "Regex" at the same time.')
            sys.exit(1)
        else:
            # スペース区切りで複数単語を検索できるようにしておく
            query = '|'.join([r'\b' + q + r'\b' for q in query.split()])
    else:
        # 正規表現
        if setting.regex:
            pass
        else:
            # 検索クエリをエスケープすることで事実上の正規表現無効状態とする
            query = re.escape(query)

    return query, dirpath, setting


# ファイルパスリストを生成
def create_fpaths(dirpath, exts):

    # glob でファイルリストを作成
    glob_path = dirpath + (r'/**/*.*' if _setting.recursive else r'/*.*')
    fpaths = glob.glob(glob_path, recursive=_setting.recursive)

    # ファイル絞り込み用の正規表現パターン
    fname_pattern = r'.*\.(' + r'|'.join(exts) + r')$'
    # 対象ファイルのリストとなるように絞り込む
    return [f for f in fpaths if re.search(fname_pattern, f, flags=re.IGNORECASE)]


# 検索処理を振り分けながら実行
def grep_while_destribute(query, target_fpaths):

    # 並列処理で実行する関数 run_grep の引数リストを準備
    n = len(target_fpaths)

    # # 検索クエリ
    querys = [query] * n
    # Office種類
    office_types = [destribute_by_ext(fpath, _office_types) for fpath in target_fpaths]
    # ファイル番号
    fnums = [n + 1 for n in range(n)]
    # ファイル総数
    fcnts = [n] * n
    # 正規表現フラグ
    re_flagss = [re.IGNORECASE if _setting.ignorecase else 0] * n
    # ロックオブジェクト
    locks = [multiprocessing.Manager().Lock()] * n

    if _setting.parallel > 1:
        # 並列処理
        ## 並列化しても速くならないため max_workers=1 とする（Win32comとの相性？）

        # ThreadPoolExecutor では chunksize は無効。ProcessPoolExecutor では適切な範囲で大きくするとパフォーマンス向上が見込める。
        with ThreadPoolExecutor(max_workers=_setting.parallel) as executor:
            executor.map(run_grep, querys, target_fpaths, office_types, fnums, fcnts, re_flagss, locks, chunksize=128)

    else:
        # 逐次処理
        for args in zip(querys, target_fpaths, office_types, fnums, fcnts, re_flagss, locks):
            run_grep(*args)


# 検索処理の実行
def run_grep(query, fpath, office_type, fnum, fcnt, re_flags, lock):
    if office_type is None:
        return

    # 標準出力カラーの初期化
    colorama.init(autoreset=True)

    # Office別の検索処理を実行してヒットログを受け取る
    hlogs = office_type.grep_func(query, fpath, re_flags)

    # 排他ロック
    with lock:
        # ファイルログを作成
        flog = make_log_file(fnum, fcnt, pathlib.Path(fpath).relative_to(os.getcwd()), office_type)

        # ログ出力
        print(flog)
        if hlogs:
            print('\n'.join(hlogs))


# grep: Excel
def grep_excel(query, fpath, re_flags):
    hlogs = []

    app = None

    try:
        # COM初期化  ※スレッドごとに必要
        pythoncom.CoInitialize()

        # Excel起動
        app = win32com.client.DispatchEx('Excel.Application')
        app.Visible = False
        app.DisplayAlerts = False

        # ブックオープン
        wb = app.Workbooks.Open(fpath)

        # 全シートを処理
        for ws in wb.Worksheets:
            debug_print(ws.Name)
            debug_print(ws.UsedRange.Address)

            # 2: xlSheetVeryHidden (再表示不可) に設定されているシートは対象外
            if ws.Visible == 2:
                continue

            # セル
            #for cell in ws.UsedRange.Cells:
            for cell in [ws.Cells(r, c) for r, c in get_used_range_strict(ws)]:  # データが疎なシートの高速化
                # 検索
                text = str(cell.Value)
                if re.search(query, text, flags=re_flags):
                    # ヒットログを作成
                    l = make_log_hit({'Sheet': ws.Name, 'Cell': cell.Address.replace('$', '')}, query, text, re_flags)
                    hlogs.append(l)

            # 図形
            for shape in ws.Shapes:
                # 図形種類がオートシェイプ(1:msoAutoShape)またはテキストボックス(17:msoTextBox)であり、
                # 文字列を持っている場合(線などではない場合)
                if (shape.Type == 1 or shape.Type == 17) and shape.TextFrame2.HasText:
                    # 検索
                    text = str(shape.TextFrame2.TextRange.Text)
                    if re.search(query, text, flags=re_flags):
                        # ヒットログを作成
                        l = make_log_hit({'Sheet': ws.Name, 'Shape': shape.Name}, query, text, re_flags)
                        hlogs.append(l)

            # コメント
            for comment in ws.Comments:
                # 検索
                text = str(comment.Text())
                if re.search(query, text, flags=re_flags):
                    # ヒットログを作成
                    l = make_log_hit({'Sheet': ws.Name, 'Comment': comment.Parent.Address.replace('$', '')}, query, text, re_flags)
                    hlogs.append(l)

        if wb is not None: wb.Close()

    except Exception as e:
        print('Error: Failer in excel operation.')
        raise

    finally:
        # アプリケーションを終了
        if app is not None: app.Quit()

    return hlogs


# Excelワークシートの使用されているセルを厳密に取得（空のセルを除去）
def get_used_range_strict(ws):

    # 全てのグループ化を解除(グループ化されていると End プロパティが想定通り動作しないため)
    ws.Cells.ClearOutline()

    # シート内の使用されているセル範囲
    row_s = ws.UsedRange(1, 1).Row
    col_s = ws.UsedRange(1, 1).Column
    row_e = row_s + ws.UsedRange.Rows.Count - 1
    col_e = col_s + ws.UsedRange.Columns.Count - 1

    # 対象セルのリスト
    #   [(行番号, 列番号), ...]
    target_cells = []

    # 範囲内の行をループ
    for row in range(row_s, row_e + 1):
        # 範囲内の列をループ
        col = col_s
        while col <= col_e:
            # カレントセル
            cell = ws.Cells(row, col)
            debug_print(cell.Address, cell.GetValue())

            # カレントから右へジャンプしたセル
            cell_rjump = cell.End(-4161)  # -4161: xlToRight
            debug_print(cell_rjump.Address, cell_rjump.GetValue())

            # カレントセルに値がある
            if cell.GetValue() is not None:
                # 右隣のセル
                #cell_rnext = cell.GetOffset(0, 1)
                cell_rnext = ws.Cells(row, col + 1)  # 結合されている場合を考慮して Offset プロパティは使用しない
                debug_print(cell_rnext.Address, cell_rnext.GetValue())

                # 右隣のセルに値がある
                if cell_rnext.GetValue() is not None:
                    # カレントからジャンプ先のセルまでを対象とする
                    target_cells += [(row, c) for c in range(cell.Column, cell_rjump.Column + 1)]

                    # カレントをジャンプ先のさらに右隣に移動
                    col = cell_rjump.Column + 1

                # 右隣のセルに値がない
                else:
                    # カレントを対象とする
                    target_cells.append((row, cell.Column))

                    # カレントをジャンプ先に移動
                    col = cell_rjump.Column

            # カレントセルに値がない
            else:
                # カレントをジャンプ先に移動
                col = cell_rjump.Column

    debug_print(target_cells)
    return target_cells


# grep: Word
def grep_word(query, fpath, re_flags):
    return []


# grep: PowerPoint
def grep_ppoint(query, fpath, re_flags):
    return []


# 拡張子による振り分け
def destribute_by_ext(fpath, office_types):

    # ファイルパスから拡張子を取得
    _, ext = os.path.splitext(fpath)
    ext = ext[1:].lower()   # 先頭の '.' を削除、小文字化

    # 拡張子に対応するものを探す
    for office_type in office_types:
        if ext in office_type.exts:
            return office_type

    return None


# ファイルログを出力
def make_log_file(n, nmax, fpath, office_type):
    n = str(n)
    nmax = str(nmax)
    kind = office_type.char if office_type else '?'
    header = COLOR_FILE + '{}({}/{})'.format(kind, n.rjust(len(nmax)), nmax) + Fore.RESET + Back.RESET

    log = '{}: {}'.format(header, fpath)
    return log


# 検索ヒットログを出力
def make_log_hit(info, query, text, re_flags):

    # ヒット箇所の情報を作成
    info = ['{}="{}"'.format(k, v) for k, v in info.items()]
    info = ', '.join(info)
    info = COLOR_HIT_INFO + '{}'.format(info) + Fore.RESET + Back.RESET

    # 検索にヒットする位置を順に辿りログ出力用の文字列を生成する
    text = text.replace('\n', '')
    log_text = ''
    before = 0
    for m in re.finditer(query, text, flags=re_flags):
        # 前回のヒット位置の直後から今回のヒット位置末尾までを生成
        log_text += text[before:m.start()] + COLOR_HIT_POS + text[m.start():m.end()] + Fore.RESET + Back.RESET
        before = m.end()
    log_text += text[before:]

    log = '  {}: {}'.format(info, log_text)
    return log


def debug_print(*args):
    if DEBUG:
        print(*args)


if __name__ == '__main__':
    main()
