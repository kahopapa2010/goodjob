"""月次の操業日報Excel → JUST.DB移行用CSV・検証結果・分析Excel をまとめて作る
使い方: python run_all.py <出力フォルダ> <月次日報.xlsx> [<月次日報.xlsx> ...]
出力:
  <出力フォルダ>/03_移行用CSV/   マスタ・T01〜T12・T90・検証CSV・00_移行手順と検証結果_YYYYMM.xlsx
  <出力フォルダ>/04_操業分析（CSV差し替え式）.xlsx   全月分のデータ入り
"""
import sys, os, subprocess
HERE = os.path.dirname(os.path.abspath(__file__))
out, srcs = sys.argv[1], sys.argv[2:]
csvd = os.path.join(out, "03_移行用CSV")
os.makedirs(csvd, exist_ok=True)
ok = True
for s in srcs:
    print(f"=== {os.path.basename(s)} ===")
    subprocess.run([sys.executable, os.path.join(HERE, "extract.py"), s, csvd], check=True)
    ok &= subprocess.run([sys.executable, os.path.join(HERE, "build_masters.py"), s, csvd]).returncode == 0
subprocess.run([sys.executable, os.path.join(HERE, "build_analysis.py"), csvd,
                os.path.join(out, "04_操業分析（CSV差し替え式）.xlsx")], check=True)
print("検証: すべて一致" if ok else "検証: 差がある月があります（上のNG行を確認）")
sys.exit(0 if ok else 1)
