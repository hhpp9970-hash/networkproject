"""src/NetworkSchedule.vba（UTF-8）から、Excel に取り込む NetworkSchedule.bas（Shift_JIS・改行 CRLF）を作る。

使い方
  python tools/build_bas.py           チェックしてから NetworkSchedule.bas を書く
  python tools/build_bas.py --verify  書かずに、今の NetworkSchedule.bas が src と同じかだけ確かめる

チェック（どれかで引っかかったら書かない）
  check_vba.py   If/For/With などの閉じ忘れ、宣言のない代入、定義のない呼び出し、Shift_JIS にない文字
  check_vba2.py  式の中の、宣言も定義もない名前（下の KNOWN にある Excel/VBA の組み込みは除く）
  check_vba3.py  予約語と同じ名前（eNum ＝ Enum など）、同じ範囲での重複、ByRef 引数の型違い
"""
import ast
import os
import pathlib
import subprocess
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
SRC = ROOT / "src" / "NetworkSchedule.vba"
DST = ROOT / "NetworkSchedule.bas"
TOOLS = ROOT / "tools"

# check_vba2.py が「宣言のない名前」として出すが、Excel/VBA の組み込み（またはラベル）で問題ないもの。
# 新しい定数などを使ったら、組み込みであることを確かめてからここに足す
KNOWN = {
    "DoEvents", "Timer", "Like", "Name", "FreeformBuilder", "ListObject", "Done", "Fail", "DateSerial",
    "vbCr", "vbLf", "vbTab", "vbOKCancel", "vbYesNo",
    "msoAlignRight", "msoEditingAuto", "msoSegmentLine", "msoShapeDiamond", "msoShapeRoundedRectangle",
    "xlCalculationManual", "xlDefault", "xlDot", "xlDouble", "xlErrorHandler", "xlInsideHorizontal", "xlPaperA4", "xlInsideVertical", "xlInterrupt", "xlPatternLightUp", "xlThick", "xlWait",
}


def run(script):
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    r = subprocess.run([sys.executable, str(TOOLS / script), str(SRC)], capture_output=True, text=True,
                       encoding="utf-8", env=env)
    if r.returncode != 0:
        sys.exit(f"{script} が止まりました:\n{r.stderr}")
    return r.stdout


def check():
    bad = []
    out = run("check_vba.py")
    for want in ("| unclosed: [] | errors: none", "undeclared assignments: none",
                 "statement calls not defined here: none", "cp932-unencodable: none"):
        if want not in out:
            bad.append(out.strip())
            break
    unknown = {}
    for line in run("check_vba2.py").splitlines():
        if "[" in line and not line.startswith("done"):
            proc = line[:line.index("[")].strip()
            names = set(ast.literal_eval(line[line.index("["):].strip())) - KNOWN
            if names:
                unknown[proc] = sorted(names)
    if unknown:
        bad.append("宣言も定義もない名前（打ち間違いか、Excel/VBA の組み込みなら KNOWN に足す）:\n" +
                   "\n".join(f"  {p}: {', '.join(n)}" for p, n in unknown.items()))
    out3 = run("check_vba3.py")
    if " 0 件" not in out3:
        bad.append(out3.strip())
    return bad


def build_bytes():
    text = SRC.read_text(encoding="utf-8").replace("\r\n", "\n")
    if not text.startswith('Attribute VB_Name = "NetworkSchedule"\n'):
        sys.exit("src/NetworkSchedule.vba の1行目は Attribute VB_Name = \"NetworkSchedule\" にしてください")
    try:
        return text.replace("\n", "\r\n").encode("cp932")
    except UnicodeEncodeError as e:
        line = text[:e.start].count("\n") + 1
        sys.exit(f"Shift_JIS にない文字があります（{line}行目）: {text[e.start:e.end]!r}")


if __name__ == "__main__":
    data = build_bytes()
    if "--verify" in sys.argv:
        same = DST.exists() and DST.read_bytes() == data
        print("NetworkSchedule.bas は src と同じです" if same else "NetworkSchedule.bas が src と違います（build してください）")
        sys.exit(0 if same else 1)
    problems = check()
    if problems:
        print("\n\n".join(problems))
        sys.exit("チェックで止めました。NetworkSchedule.bas は書いていません。")
    DST.write_bytes(data)
    n_lines = data.count(b"\r\n")
    print(f"NetworkSchedule.bas を書きました（{len(data):,} バイト、{n_lines:,} 行、Shift_JIS・CRLF）")
