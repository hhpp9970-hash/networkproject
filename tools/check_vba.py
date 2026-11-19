import re, sys
src = open(sys.argv[1], encoding="utf-8").read()
lines, buf = [], ""
for raw in src.split("\n"):
    code = raw.rstrip()
    if code.endswith(" _"):
        buf += code[:-2] + " "; continue
    lines.append(buf + code); buf = ""
def strip_comment(s):
    out, q = [], False
    for ch in s:
        if ch == '"': q = not q
        if ch == "'" and not q: break
        out.append(ch)
    return "".join(out).strip()
stack, errors, procs = [], [], 0
for no, line in enumerate(lines, 1):
    s = strip_comment(line); low = s.lower()
    if not s: continue
    def push(k): stack.append((k, no))
    def pop(k):
        if not stack or stack[-1][0] != k: errors.append(f"line {no}: end of {k}; stack={stack[-3:]}")
        else: stack.pop()
    if re.match(r"^(public |private )?(sub|function) ", low): push("proc"); procs += 1
    elif re.match(r"^end (sub|function)$", low): pop("proc")
    elif re.match(r"^if .* then$", low): push("if")
    elif low == "end if": pop("if")
    elif re.match(r"^for ", low): push("for")
    elif re.match(r"^next( \w+)?$", low): pop("for")
    elif re.match(r"^do( while .*| until .*)?$", low): push("do")
    elif re.match(r"^loop", low): pop("do")
    elif re.match(r"^with ", low): push("with")
    elif low == "end with": pop("with")
    elif re.match(r"^select case ", low): push("select")
    elif low == "end select": pop("select")
print("procedures:", procs, "| unclosed:", stack, "| errors:", errors or "none")

module_decl = set()
for m in re.finditer(r"^(?:Private|Public)\s+(?!Sub|Function|Const)([^\n']+)", src, re.M):
    for part in m.group(1).split(","):
        nm = re.match(r"\s*(\w+)", part)
        if nm: module_decl.add(nm.group(1).lower())
for m in re.finditer(r"^Private Const (\w+)", src, re.M):
    module_decl.add(m.group(1).lower())
blocks = re.split(r"\n(?=(?:Public |Private )?(?:Sub|Function) )", src)
issues = []
for block in blocks[1:]:
    flat = block.replace(" _\n", " ")
    head = re.match(r"(?:Public |Private )?(?:Sub|Function) (\w+)\((.*?)\)(?: As [\w()]+)?\s*\n", flat)
    if not head: issues.append(("?", "header", flat[:60])); continue
    name = head.group(1).lower()
    local = {name}
    for part in head.group(2).split(","):
        nm = re.match(r"\s*(?:ByVal|ByRef)?\s*(\w+)", part)
        if nm: local.add(nm.group(1).lower())
    for m in re.finditer(r"\bDim\s+([^\n']+)", flat):
        for part in m.group(1).split(","):
            nm = re.match(r"\s*(\w+)", part)
            if nm: local.add(nm.group(1).lower())
    for line in flat.split("\n")[1:]:
        st = strip_comment(line)
        for m in re.finditer(r"(?:^|:\s*|\bThen\s+|\bElse\s+)(?:Set\s+)?(\w+)(?:\([^)]*\))?\s*=(?!=)", st):
            v = m.group(1).lower()
            if v in ("if", "elseif", "case", "for", "do", "while", "set"): continue
            if v not in local and v not in module_decl:
                issues.append((name, v, st[:70]))
        # 使われている識別子のうち、宣言も定義もないもの（ざっくり）
print("undeclared assignments:", issues or "none")

defined = {m.group(1).lower() for m in re.finditer(r"(?:Sub|Function) (\w+)\(", src)}
called = set()
skip = {"dim","if","for","next","do","loop","with","end","select","case","set","exit","redim",
        "private","public","option","attribute","on","goto","msgbox","else","elseif","call"}
for line in lines:
    st = strip_comment(line)
    m = re.match(r"^(?:If .*? Then )?(\w+)\s+[^=:]", st)
    if m and m.group(1)[0].isupper() and m.group(1).lower() not in skip:
        called.add(m.group(1).lower())
print("statement calls not defined here:", sorted(c for c in called if c not in defined) or "none")
bad = sorted({ch for ch in src if not ch.encode("cp932", errors="ignore")})
print("cp932-unencodable:", bad or "none")
