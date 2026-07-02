# -*- coding: utf-8 -*-
"""VBA整合性検査: 運用中VBA のコンパイル前静的チェック。

使い方:
    python _verify_vba_integrity.py [運用中VBAディレクトリ]
    (省略時はカレントディレクトリ)

検査項目:
  [形式]   CP932厳格デコード / BOM / CRLF・LF混在 / .bas 先頭 Option Explicit
  [戻り値] 関数名と戻り値代入先の不一致 (Impl改名漏れ等)
  [修飾]   多重モジュール修飾 (mod_X.mod_Y.Proc) / 修飾参照先の欠落・Private
  [可視性] 非修飾での他モジュールPrivateプロシージャ呼び出し
  [型]     ユーザー定義型 (As 型名) の未定義使用
  [変数]   m接頭辞モジュール変数の未宣言使用
  [欠落]   *.pre_split 併存時: プロシージャ / モジュールレベル宣言の消失
  [重複]   委譲スタブでない Public 実体の重複定義

終了コード: 要修正ありなら 1、なしなら 0。名前照合は大文字小文字を無視。
"""
import os
import re
import sys

PROC_RE = re.compile(
    r'^\s*(Public|Private|Friend)?\s*(Static\s+)?'
    r'(Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_]\w*)')
END_RE = re.compile(r'^\s*End (Sub|Function|Property)\b')
TYPE_RE = re.compile(r'^\s*(Public|Private)?\s*Type\s+([A-Za-z_]\w*)')
ENUM_RE = re.compile(r'^\s*(Public|Private)?\s*Enum\s+([A-Za-z_]\w*)')
QUAL_RE = re.compile(r'\b(mod_[A-Za-z0-9_]+)\s*\.\s*([A-Za-z_]\w*)')
MULTIQ_RE = re.compile(r'\b(mod_\w+)\s*\.\s*(mod_\w+)\s*\.')
VAR_DECL_RE = re.compile(
    r'^(Public|Private|Dim|Global)\s+(?:WithEvents\s+)?(?:Const\s+)?(.+)$')

VBA_KNOWN_TYPES = {t.lower() for t in [
    'Long', 'LongLong', 'LongPtr', 'String', 'Boolean', 'Variant', 'Object',
    'Integer', 'Double', 'Date', 'Byte', 'Single', 'Currency', 'Worksheet',
    'Workbook', 'Range', 'Collection', 'OLEObject', 'XlCalculation', 'Shape',
    'Comment', 'Name', 'Chart', 'Window', 'Control', 'Application', 'Names',
    'CommandBar', 'FileDialog', 'Validation', 'Font', 'Interior', 'Borders',
    'Border', 'Areas', 'PageSetup', 'Hyperlink', 'ListObject', 'Err',
    'CheckBox', 'OptionButton', 'CommandButton', 'TextBox', 'Label', 'Frame',
    'ListBox', 'ComboBox', 'MultiPage', 'ListView', 'ListItem',
    'MSComctlLib', 'MSForms', 'stdole', 'UserForm', 'ReturnInteger',
    'ReturnBoolean', 'XlBordersIndex', 'XlLineStyle', 'XlBorderWeight',
    'VbMsgBoxResult', 'VbMsgBoxStyle', 'FormatCondition',
    'XlAutoFilterOperator', 'VbCallType', 'VbFileAttribute',
    'XlSheetVisibility', 'XlPasteType', 'AutoFilter', 'Sort',
]}


def strip_code(line):
    s = re.sub(r'"[^"]*"', '""', line)
    ci = s.find("'")
    return s if ci < 0 else s[:ci]


def load(path):
    raw = open(path, 'rb').read()
    problems = []
    if raw[:3] == b'\xef\xbb\xbf':
        problems.append('BOMあり')
    try:
        txt = raw.decode('cp932')
    except UnicodeDecodeError as e:
        problems.append('CP932厳格デコード不可 (byte offset %d)' % e.start)
        txt = raw.decode('cp932', errors='replace')
    lf, crlf = raw.count(b'\n'), raw.count(b'\r\n')
    if 0 < crlf < lf:
        problems.append('CRLF/LF混在')
    return txt.replace('\r\n', '\n'), problems


def decl_names(decl_line):
    """宣言行から変数名リストを抽出 (複数宣言・WithEvents対応)。"""
    m = VAR_DECL_RE.match(decl_line)
    if not m:
        return []
    rest = m.group(2)
    names = []
    for part in rest.split(','):
        nm = re.match(r'\s*(?:WithEvents\s+)?([A-Za-z_]\w*)', part)
        if nm:
            names.append(nm.group(1))
    return names


def module_level_decls(txt):
    lines = txt.split('\n')
    out = []
    i = 0
    while i < len(lines):
        if PROC_RE.match(lines[i]):
            while i < len(lines) and not END_RE.match(lines[i]):
                i += 1
        else:
            s = lines[i].strip()
            if s and not s.startswith("'"):
                out.append(s)
        i += 1
    return out


def proc_bodies(txt):
    lines = txt.split('\n')
    i = 0
    while i < len(lines):
        m = PROC_RE.match(lines[i])
        if m:
            start = i
            body = []
            # 1行完結 (Sub x(): ...: End Sub) 対応
            if re.search(r':\s*End (Sub|Function|Property)\b',
                         strip_code(lines[i])):
                yield (m.group(4), m.group(3), m.group(1) or 'Public',
                       start + 1, body)
                i += 1
                continue
            i += 1
            while i < len(lines) and not END_RE.match(lines[i]):
                body.append(lines[i])
                i += 1
            yield (m.group(4), m.group(3), m.group(1) or 'Public',
                   start + 1, body)
        i += 1


def collect_sig_params(lines, start_idx):
    sig = lines[start_idx]
    j = start_idx + 1
    while sig.rstrip().endswith('_') and j < len(lines):
        sig = sig.rstrip()[:-1] + ' ' + lines[j]
        j += 1
    return {p.lower() for p in re.findall(
        r'(?:ByVal|ByRef|Optional|ParamArray)\s+([A-Za-z_]\w*)', sig)} | \
        {p.lower() for p in re.findall(r'[,(]\s*([A-Za-z_]\w*)\s+As', sig)}


def main(root):
    files = {}
    issues = []
    for f in sorted(os.listdir(root)):
        if not f.endswith(('.bas', '.cls', '.frm')):
            continue
        if f.endswith(('.bak', '.pre_split')):
            continue
        txt, probs = load(os.path.join(root, f))
        files[f] = txt
        for p in probs:
            issues.append((f, 0, '形式', p))
        if f.endswith('.bas'):
            first = txt.split('\n')[0].strip()
            if first != 'Option Explicit':
                issues.append((f, 1, '形式',
                               '先頭が Option Explicit でない: ' + first[:40]))

    # 定義テーブル (キーは小文字)
    defs = {}
    types = {}
    enums = {}
    mod_priv_vars = {}
    pub_vars = set()
    for f, txt in files.items():
        for name, kind, scope, ln, body in proc_bodies(txt):
            defs.setdefault(name.lower(), []).append((f, scope, kind, name))
        mv = set()
        for d in module_level_decls(txt):
            tm = TYPE_RE.match(d)
            if tm:
                types.setdefault(tm.group(2).lower(), []).append(
                    (f, tm.group(1) or 'Private'))
                continue
            em = ENUM_RE.match(d)
            if em:
                enums.setdefault(em.group(2).lower(), []).append(f)
                continue
            for n in decl_names(d):
                mv.add(n.lower())
                if d.lstrip().startswith(('Public', 'Global')):
                    pub_vars.add(n.lower())
        mod_priv_vars[f] = mv

    mod_names = {os.path.splitext(f)[0].lower() for f in files}
    mod_file = {os.path.splitext(f)[0].lower(): f for f in files}

    for f, txt in files.items():
        lines = txt.split('\n')
        local_procs = {k for k, v in defs.items()
                       if any(mm == f for mm, s, kd, orig in v)}
        for name, kind, scope, ln, body in proc_bodies(txt):
            params = collect_sig_params(lines, ln - 1)
            local_vars = set(params)
            for bl in body:
                c = strip_code(bl)
                for dm in re.finditer(
                        r'\b(?:Dim|Const|Static)\s+(.+)$', c):
                    for part in dm.group(1).split(','):
                        nm = re.match(r'\s*([A-Za-z_]\w*)', part)
                        if nm:
                            local_vars.add(nm.group(1).lower())
            is_func = kind == 'Function' or kind.startswith('Property Get')
            for k, bl in enumerate(body):
                c = strip_code(bl)
                am = re.match(r'\s*(?:Set\s+)?([A-Za-z_]\w*)\s*=(?!=)', c)
                if am and is_func:
                    lhs = am.group(1)
                    ll = lhs.lower()
                    if (ll != name.lower() and ll in defs
                            and ll not in local_vars
                            and ll not in mod_priv_vars.get(f, set())
                            and ll not in pub_vars
                            and any(kd == 'Function'
                                    for _, _, kd, _ in defs[ll])):
                        issues.append((f, ln + k + 1, '戻り値',
                                       '関数 %s 内で別関数名 %s へ代入'
                                       % (name, lhs)))
                if MULTIQ_RE.search(c):
                    issues.append((f, ln + k + 1, '修飾',
                                   '多重モジュール修飾: ' + bl.strip()[:80]))
                for qm, pn in QUAL_RE.findall(c):
                    qml, pnl = qm.lower(), pn.lower()
                    if qml not in mod_names:
                        issues.append((f, ln + k + 1, '修飾',
                                       '%s.%s: モジュールが存在しない'
                                       % (qm, pn)))
                        continue
                    ent = [e for e in defs.get(pnl, [])
                           if os.path.splitext(e[0])[0].lower() == qml]
                    tgt = mod_file[qml]
                    if not ent and pnl not in mod_priv_vars.get(tgt, set()):
                        issues.append((f, ln + k + 1, '修飾',
                                       '%s.%s: 定義が見つからない'
                                       % (qm, pn)))
                    elif ent and not any(s == 'Public'
                                         for _, s, _, _ in ent):
                        issues.append((f, ln + k + 1, '修飾',
                                       '%s.%s: Private のため参照不可'
                                       % (qm, pn)))
                for w in set(re.findall(r'(?<![\w.])([A-Za-z_]\w{3,})\s*\(',
                                        c)):
                    wl = w.lower()
                    if (wl in local_procs or wl in local_vars
                            or wl in mod_priv_vars.get(f, set())
                            or wl in pub_vars):
                        continue
                    ent = defs.get(wl)
                    if ent and not any(s == 'Public'
                                       for _, s, _, _ in ent):
                        issues.append((f, ln + k + 1, '可視性',
                                       '%s は他モジュールで Private のみ' % w))
                for tname in re.findall(r'\bAs\s+(?:New\s+)?([A-Za-z_]\w*)',
                                        c):
                    tl = tname.lower()
                    if (tl in VBA_KNOWN_TYPES or tl in types or tl in enums
                            or tl in mod_names or tl in defs):
                        continue
                    issues.append((f, ln + k + 1, '型',
                                   '未定義の型: As %s' % tname))
                for w in set(re.findall(r'\b(m[A-Z]\w+)\b', c)):
                    wl = w.lower()
                    if (wl in mod_priv_vars.get(f, set()) or wl in pub_vars
                            or wl in local_vars or wl in defs
                            or wl in types or wl in enums):
                        continue
                    issues.append((f, ln + k + 1, '変数',
                                   'モジュール変数 %s が未宣言' % w))

    # pre_split 比較
    for f in sorted(os.listdir(root)):
        if not f.endswith('.pre_split'):
            continue
        base = f[:-len('.pre_split')]
        pre_txt, _ = load(os.path.join(root, f))
        pre_procs = {n.lower() for n, k, s, ln, b in proc_bodies(pre_txt)}
        lost = pre_procs - set(defs.keys())
        if lost:
            issues.append((base, 0, '欠落',
                           'pre_split のプロシージャが全モジュールに無い: '
                           + ', '.join(sorted(lost))))
        all_decls = set()
        for tf, ttxt in files.items():
            all_decls |= {d.lower() for d in module_level_decls(ttxt)}
        for d in module_level_decls(pre_txt):
            if d.lower() in all_decls or d.startswith('Option'):
                continue
            # Type/Enum のメンバー行や End Type は本体行なので Type名で照合
            tm = TYPE_RE.match(d)
            if tm and tm.group(2).lower() in types:
                continue
            if re.match(r'(End Type|End Enum)\b', d):
                continue
            if not VAR_DECL_RE.match(d) and not tm and not ENUM_RE.match(d):
                continue
            issues.append((base, 0, '欠落',
                           'モジュールレベル宣言が消失: ' + d[:80]))

    # 重複Public実体
    for nl, v in defs.items():
        pubs = [(m, orig) for m, s, kd, orig in v
                if s == 'Public' and m.endswith('.bas')]
        if len(pubs) > 1:
            real = []
            for m, orig in pubs:
                for n2, k2, s2, ln2, b2 in proc_bodies(files[m]):
                    if n2.lower() == nl and s2 == 'Public':
                        code = [strip_code(x).strip() for x in b2
                                if strip_code(x).strip()]
                        deleg = (len(code) <= 2 and code
                                 and 'mod_' in code[0])
                        if not deleg:
                            real.append(m)
            if len(set(real)) > 1:
                issues.append((','.join(sorted(set(real))), 0, '重複',
                               'Public %s が複数モジュールに実体定義'
                               % pubs[0][1]))

    seen = set()
    fatal_kinds = {'戻り値', '修飾', '可視性', '欠落', '変数', '重複',
                   '形式', '型'}
    fatal = 0
    for f, ln, kind, msg in sorted(set(issues)):
        mark = '[!]' if kind in fatal_kinds else '[?]'
        if kind in fatal_kinds:
            fatal += 1
        print('%s %-6s %s:%s  %s' % (mark, kind, f, ln or '-', msg))
    print()
    print('検出: 要修正 %d 件' % fatal)
    return 1 if fatal else 0


if __name__ == '__main__':
    sys.exit(main(sys.argv[1] if len(sys.argv) > 1 else '.'))
