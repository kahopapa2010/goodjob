"""変換結果をWindowsで文字化けしないzipにまとめる
使い方: python make_zip.py <出力フォルダ>
<出力フォルダ>/操業日報_JUSTDB移行_YYYYMM.zip（複数月は YYYYMM-YYYYMM）を作る。
ファイル名にUTF-8の目印（汎用フラグ bit 11）を付ける。Linuxの zip コマンドは付けないことがあり、
Windowsのエクスプローラーで開くとファイル名がShift_JISとして読まれて文字化けする。
"""
import sys, os, re, glob, zipfile

def make_zip(out):
    csvd = os.path.join(out, "03_移行用CSV")
    months = sorted({m for p in glob.glob(os.path.join(csvd, "T01_日報_*.csv"))
                     for m in re.findall(r"_(\d{6})\.csv$", p)})
    tag = months[0] if len(months) == 1 else f"{months[0]}-{months[-1]}"
    dst = os.path.join(out, f"操業日報_JUSTDB移行_{tag}.zip")
    if os.path.exists(dst): os.remove(dst)
    files = [os.path.join(csvd, f) for f in sorted(os.listdir(csvd))]
    files += glob.glob(os.path.join(out, "04_*.xlsx"))
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        for p in files:
            name = os.path.relpath(p, out).replace(os.sep, "/")
            zi = zipfile.ZipInfo.from_file(p, name)
            zi.flag_bits |= 0x800          # ファイル名はUTF-8
            zi.compress_type = zipfile.ZIP_DEFLATED
            with open(p, "rb") as f: z.writestr(zi, f.read())
    return dst

if __name__ == "__main__":
    print(make_zip(sys.argv[1]))
