"""月次の操業日報Excel → JUST.DB移行用CSV・検証結果・分析Excel をまとめて作る
使い方: python run_all.py [-j 並列数] <出力フォルダ> <月次日報.xlsx> [<月次日報.xlsx> ...]
出力:
  <出力フォルダ>/03_移行用CSV/   マスタ・T01〜T12・T90・検証CSV・00_移行手順と検証結果_YYYYMM.xlsx
  <出力フォルダ>/04_操業分析（CSV差し替え式）.xlsx   全月分のデータ入り
複数月は並列で変換する（既定はCPU数まで）。全月共通のマスタ M01〜M08 は、
これまでどおり最後に指定した月の内容になる。
"""
import sys, os, shutil, subprocess, tempfile, argparse
from concurrent.futures import ThreadPoolExecutor
HERE = os.path.dirname(os.path.abspath(__file__))

ap = argparse.ArgumentParser()
ap.add_argument("-j", "--jobs", type=int, default=os.cpu_count() or 1)
ap.add_argument("out"); ap.add_argument("srcs", nargs="+")
a = ap.parse_args()
csvd = os.path.join(a.out, "03_移行用CSV")
os.makedirs(csvd, exist_ok=True)

def convert(src, work):
    """1か月分を作業フォルダに変換する。戻り値: (検証OKか, 表示用ログ)"""
    log = []
    for script in ("extract.py", "build_masters.py"):
        p = subprocess.run([sys.executable, os.path.join(HERE, script), src, work],
                           capture_output=True, text=True)
        log.append(p.stdout + p.stderr)
        if script == "extract.py" and p.returncode != 0:
            return None, "".join(log)
    return p.returncode == 0, "".join(log)

with tempfile.TemporaryDirectory(dir=a.out) as tmp:
    works = [os.path.join(tmp, str(i)) for i in range(len(a.srcs))]
    with ThreadPoolExecutor(max_workers=max(1, a.jobs)) as ex:
        results = list(ex.map(convert, a.srcs, works))
    ok = True
    for src, work, (res, log) in zip(a.srcs, works, results):   # 指定順に出力をまとめる
        print(f"=== {os.path.basename(src)} ===")
        print(log, end="")
        if res is None:
            sys.exit(f"変換に失敗しました: {src}")
        ok &= res
        for f in os.listdir(work):
            shutil.copy2(os.path.join(work, f), os.path.join(csvd, f))

subprocess.run([sys.executable, os.path.join(HERE, "build_analysis.py"), csvd,
                os.path.join(a.out, "04_操業分析（CSV差し替え式）.xlsx")], check=True)
from make_zip import make_zip
print("渡すzip:", make_zip(a.out))
print("検証: すべて一致" if ok else "検証: 差がある月があります（上のNG行を確認）")
sys.exit(0 if ok else 1)
