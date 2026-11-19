"""式の中で使っている識別子のうち、宣言・定義・組み込みのどれでもないものを手続きごとに出す（Option Explicit 対策）"""
import re, sys
src = open(sys.argv[1], encoding="utf-8").read().replace("\r\n", "\n")
def strip(line):
    out, q = [], False
    for ch in line:
        if ch == '"':
            q = not q; out.append(" "); continue
        if q: out.append(" "); continue
        if ch == "'": break
        out.append(ch)
    return "".join(out)
module = set()
for m in re.finditer(r"^(?:Private|Public)\s+(?!Sub|Function)(?:Const\s+)?([^\n']+)", src, re.M):
    for part in re.split(r",(?![^(]*\))", m.group(1)):
        nm = re.match(r"\s*(\w+)", part)
        if nm: module.add(nm.group(1).lower())
procs = {m.group(1).lower() for m in re.finditer(r"(?:Sub|Function) (\w+)\(", src)}
builtin = set("""and or not mod is new nothing true false empty null me byval byref optional as long double string boolean
variant object integer single date currency byte collection dictionary worksheet workbook range shape shapes
dim redim preserve set let call if then else elseif end for each to step next do while until loop with select case
exit sub function goto on error resume option explicit private public const attribute vb_name in
len trim mid left right ucase lcase replace split join instr strconv cstr clng cdbl cdate cint int abs sqr format
isdate isnumeric isempty iserror isnull isobject vartype typename month day year weekday date now
msgbox createobject rgb iif lbound ubound array asc ascw chr chrw val dir
thisworkbook application activesheet activewindow activeworkbook worksheets sheets cells columns rows
vbexclamation vbinformation vbcritical vbquestion vbyesnocancel vbyes vbno vbcancel vbcrlf vbnarrow vbsunday vbsaturday vbmonday
vbboolean vbstring vbmsgboxresult msoshaperectangle msoshapeoval msotextorientationhorizontal msotrue msofalse
msolinesolid msolinedash msolinerounddot msolinesquaredot msoarrowheadtriangle msoarrowheadshort msoarrowheadnarrow
msoanchormiddle msoanchortop msoanchorbottom msoalignleft msoaligncenter msoautosizenone msofiledialogfolderpicker
xlup xltoleft xlcenter xlleft xlcontinuous xledgebottom xledgeleft xledgeright xledgetop xlmedium xlthin xllandscape xlpapera3
xltypepdf xlqualitystandard xlfreefloating xlcolorindexnone xlnone
err description number""".split())
issues = {}
for blk in re.split(r"\n(?=(?:Public |Private )?(?:Sub|Function) )", src)[1:]:
    flat = blk.replace(" _\n", " ")
    head = re.match(r"(?:Public |Private )?(?:Sub|Function) (\w+)\((.*?)\)", flat)
    name = head.group(1)
    local = {name.lower()}
    for part in head.group(2).split(","):
        nm = re.match(r"\s*(?:Optional\s+)?(?:ByVal|ByRef)?\s*(\w+)", part)
        if nm: local.add(nm.group(1).lower())
    for m in re.finditer(r"\bDim\s+([^\n']+)", flat):
        for part in re.split(r",(?![^(]*\))", m.group(1)):
            nm = re.match(r"\s*(\w+)", part)
            if nm: local.add(nm.group(1).lower())
    for line in flat.split("\n")[1:]:
        st = strip(line)
        st = re.sub(r"\w+:=", " ", st)                    # 名前付き引数
        st = re.sub(r"^\s*\w+:\s*$", " ", st)             # 行ラベル
        st = re.sub(r"&H[0-9A-Fa-f]+&?", " ", st)         # 16進
        for m in re.finditer(r"(?<![\.\w])([A-Za-z_]\w*)", st):
            v = m.group(1).lower()
            if v in local or v in module or v in procs or v in builtin:
                continue
            if re.match(r"^(errhandler|cleanup)$", v):
                continue
            issues.setdefault(name, set()).add(m.group(1))
for k, v in issues.items():
    print(k, sorted(v))
print("done:", len(issues), "procedures with unknown identifiers")
