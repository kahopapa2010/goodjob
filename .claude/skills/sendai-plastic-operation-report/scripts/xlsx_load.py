"""日報Excelの読み込み（extract.py・build_masters.py 共通）"""
import openpyxl
import openpyxl.reader.excel as _rx

# 画像・グラフは変換に使わないので読まない。
# 実ファイルに実体のない画像参照（xl/drawings/NULL）があり、そのままだと読み込みで止まるため。
_rx.find_images = lambda archive, path: ([], [])

def load(path, data_only=False):
    return openpyxl.load_workbook(path, data_only=data_only)
