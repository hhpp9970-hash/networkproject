"""VBA の予約語を名前に使っていないか（例：eNum ＝ Enum）、同じ範囲で名前が重なっていないか、
ByRef の引数に型の違う変数を渡していないかを調べる。check_vba.py / check_vba2.py の補い。"""
import re
import sys

# MS-VBAL 3.3.5.2 の予約語と、識別子に使うと構文エラーになる語
RESERVED = """
Call Case Close Const Declare DefBool DefByte DefCur DefDate DefDbl DefInt DefLng DefLngLng DefLngPtr DefObj DefSng
DefStr DefVar Dim Do Else ElseIf End EndIf Enum Erase Event Exit For Friend Function Get Global GoSub GoTo If
Implements Input Let Lock Loop LSet Next On Open Option Print Private Public Put RaiseEvent ReDim Resume Return RSet
Seek Select Set Static Stop Sub Type Unlock Wend While With Write
Any As ByRef ByVal Each In New Shared Until WithEvents Optional ParamArray Preserve Spc Tab Then To
AddressOf And Eqv Imp Is Like Mod Not Or TypeOf Xor
Abs CBool CByte CCur CDate CDbl CDec CInt CLng CLngLng CLngPtr CSng CStr CVar CVErr Date Debug DoEvents Fix Int Len
LenB Me PSet Scale Sgn String Array Circle InputB LBound UBound
Boolean Byte Currency Double Integer Long LongLong LongPtr Single Variant
True False Nothing Empty Null Rem Attribute LineInput CDecl Decimal DefDec
Error Property Step Object Alias Lib Line Name Width Kill Load Unload
""".split()
RES = {w.lower() for w in RESERVED}


def strip(line):
    """文字列とコメントを消す（"" のエスケープに対応）"""
    out, q, i = [], False, 0
    while i < len(line):
        ch = line[i]
        if q:
            if ch == '"':
                if i + 1 < len(line) and line[i + 1] == '"':
                    i += 2
                    continue
                q = False
            out.append(" ")
        else:
            if ch == '"':
                q = True
                out.append(" ")
            elif ch == "'":
                break
            else:
                out.append(ch)
        i += 1
    return "".join(out)


def logical_lines(text):
    buf, start = "", 0
    for n, raw in enumerate(text.split("\n"), 1):
        code = strip(raw.rstrip("\r"))
        if not buf:
            start = n
        if code.rstrip().endswith(" _"):
            buf += code.rstrip()[:-1] + " "
            continue
        buf += code
        for part in split_colons(buf):
            yield start, part
        buf = ""


def split_colons(code):
    # 文字列は消してあるので、: で文を分ける（ラベル行 "Name:" はそのまま残す）
    if re.match(r"^\s*\w+:\s*$", code):
        return [code]
    return [p for p in code.split(":")]


DECL_SPLIT = re.compile(r",(?![^(]*\))")


def names_in_decl(rest):
    out = []
    for piece in DECL_SPLIT.split(rest):
        m = re.match(r"\s*(?:WithEvents\s+)?(\w+)\s*(\([^)]*\))?\s*(?:As\s+(?:New\s+)?([\w.]+))?", piece, re.I)
        if m:
            out.append((m.group(1), (m.group(3) or "Variant").lower(), bool(m.group(2))))
    return out


def params(sig):
    out = []
    for piece in DECL_SPLIT.split(sig):
        piece = piece.strip()
        if not piece:
            continue
        m = re.match(r"(?:Optional\s+)?(?:(ByVal|ByRef)\s+)?(?:ParamArray\s+)?(\w+)\s*(\(\s*\))?\s*(?:As\s+([\w.]+))?",
                     piece, re.I)
        if m:
            out.append(dict(name=m.group(2), byval=(m.group(1) or "").lower() == "byval",
                            type=(m.group(4) or "Variant").lower(), arr=bool(m.group(3))))
    return out


def check(path):
    text = open(path, encoding="utf-8").read().replace("\r\n", "\n")
    problems = []
    module_names = {}
    procs = {}
    cur = None
    local = {}
    for n, code in logical_lines(text):
        s = code.strip()
        if not s:
            continue
        m = re.match(r"(?:(Private|Public|Friend)\s+)?(?:Static\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)"
                     r"\s*\((.*)\)\s*(?:As\s+([\w.]+))?", s, re.I)
        if m and not re.match(r"(?i)(declare)", s):
            name = m.group(3)
            ps = params(m.group(4))
            procs[name.lower()] = dict(name=name, params=ps, line=n, type=(m.group(5) or "variant").lower())
            if name.lower() in RES:
                problems.append((n, f"プロシージャ名「{name}」が予約語"))
            if name.lower() in module_names and not m.group(2).lower().startswith("property"):
                problems.append((n, f"プロシージャ名「{name}」がモジュールの変数・定数と重なる（{module_names[name.lower()]}行）"))
            module_names.setdefault(name.lower(), n)
            cur = name
            local = {}
            for p in ps:
                if p["name"].lower() in RES:
                    problems.append((n, f"{name} の引数「{p['name']}」が予約語"))
                if p["name"].lower() in local:
                    problems.append((n, f"{name} の引数「{p['name']}」が重複"))
                local[p["name"].lower()] = (n, p["type"])
            continue
        if re.match(r"(?i)End\s+(Sub|Function|Property)\b", s):
            procs[cur.lower()]["locals"] = local
            cur = None
            continue
        m = re.match(r"(?i)(Dim|Static|Private|Public|Global|Const|Private\s+Const|Public\s+Const)\s+(.*)", s)
        if m and not re.match(r"(?i)(Private|Public)\s+(Declare|Enum|Type|Sub|Function|Property)\b", s):
            rest = m.group(2)
            if re.match(r"(?i)Const\s+", rest):
                rest = rest[6:]
            for nm, tp, arr in names_in_decl(rest):
                if nm.lower() in RES:
                    problems.append((n, f"{'モジュール' if cur is None else cur} の変数・定数「{nm}」が予約語"))
                if cur is None:
                    if nm.lower() in module_names:
                        problems.append((n, f"モジュールの「{nm}」が重複（{module_names[nm.lower()]}行）"))
                    module_names[nm.lower()] = n
                    procs.setdefault("@module", dict(name="@module", params=[], locals={}))["locals"][nm.lower()] = (n, tp)
                else:
                    if nm.lower() in local:
                        problems.append((n, f"{cur} の「{nm}」が重複（{local[nm.lower()][0]}行）"))
                    local[nm.lower()] = (n, tp)
            continue
        if cur and re.match(r"^\w+:$", s):
            lab = s[:-1]
            if lab.lower() in RES:
                problems.append((n, f"{cur} のラベル「{lab}」が予約語"))

    # ByRef の引数に、型の違う変数をそのまま渡していないか（式や定数は一時変数になるので問題ない）
    modvars = procs.get("@module", {}).get("locals", {})
    cur = None
    for n, code in logical_lines(text):
        s = code.strip()
        m = re.match(r"(?:(?:Private|Public|Friend)\s+)?(?:Sub|Function|Property\s+\w+)\s+(\w+)", s, re.I)
        if m:
            cur = procs.get(m.group(1).lower())
            continue
        if cur is None:
            continue
        for pname, pr in procs.items():
            if pname == "@module" or not pr["params"]:
                continue
            for cm in re.finditer(r"(?<![\w.])" + re.escape(pr["name"]) + r"\b\s*(\(|\s+(?=[\w(\-]))", s, re.I):
                after = s[cm.end():] if cm.group(1) != "(" else s[cm.end():]
                args = split_args(after, paren=(cm.group(1) == "("))
                for a, p in zip(args, pr["params"]):
                    if p["byval"] or p["type"] == "variant":
                        continue
                    a = a.strip()
                    if not re.fullmatch(r"\w+", a):
                        continue
                    info = cur["locals"].get(a.lower()) if "locals" in cur else None
                    if info is None:
                        info = modvars.get(a.lower())
                    if info is None:
                        continue
                    if info[1] != p["type"]:
                        problems.append((n, f"{pr['name']} の ByRef 引数「{p['name']}」（{p['type']}）に {a}（{info[1]}）を渡している"))
    return problems


def split_args(s, paren):
    args, depth, cur = [], 0, ""
    for ch in s:
        if ch == "(":
            depth += 1
        elif ch == ")":
            if depth == 0 and paren:
                break
            depth -= 1
        if ch == "," and depth == 0:
            args.append(cur)
            cur = ""
            continue
        cur += ch
    args.append(cur)
    return args


if __name__ == "__main__":
    for path in sys.argv[1:]:
        ps = check(path)
        print(f"== {path}: {len(ps)} 件")
        for n, msg in ps:
            print(f"  {n}: {msg}")
