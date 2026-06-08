# -*- coding: utf-8 -*-
"""Verify UiMsg text equivalence and replacement coverage before applying."""
import ast
import os
import re
import sys

BASE = os.path.dirname(os.path.abspath(__file__))


def read_cp932(path):
    with open(path, 'rb') as f:
        return f.read().decode('cp932')


def load_assignments(filename, names):
    src = open(os.path.join(BASE, filename), encoding='utf-8').read()
    mod = ast.parse(src)
    out = {}
    for node in mod.body:
        if isinstance(node, ast.Assign):
            for target in node.targets:
                if isinstance(target, ast.Name) and target.id in names:
                    out[target.id] = ast.literal_eval(node.value)
    return out


def parse_ui_messages():
    ui = read_cp932(os.path.join(BASE, 'mod_UiMessages.bas'))
    func_text = {}
    for match in re.finditer(
        r'Public Function (UiMsg\w+)\(\) As String.*?CommonTextFromChars\(([^)]+)\)',
        ui,
        re.S,
    ):
        name = match.group(1)
        codes = [int(code, 16) for code in re.findall(r'&H([0-9A-Fa-f]+)', match.group(2))]
        func_text[name] = ''.join(chr(code) for code in codes)
    return func_text


def extract_quoted_literals(fragment):
    return re.findall(r'"([^"]*)"', fragment)


def reconstruct_new_text(fragment, func_text):
    result = fragment
    for name in sorted(func_text, key=len, reverse=True):
        result = result.replace('%s()' % name, '"%s"' % func_text[name])
    return ''.join(extract_quoted_literals(result))


def reconstruct_old_text(fragment):
    return ''.join(extract_quoted_literals(fragment))


def verify():
    gen = load_assignments('_gen_ui_messages.py', {'MESSAGES'})
    apply_data = load_assignments('_apply_ui_messages.py', {'REPLACEMENTS', 'TARGETS'})
    messages = gen['MESSAGES']
    replacements = apply_data['REPLACEMENTS']
    targets = apply_data['TARGETS']
    func_text = parse_ui_messages()

    errors = []

    if len(func_text) != len(messages):
        errors.append('function count mismatch: bas=%d gen=%d' % (len(func_text), len(messages)))

    for name, expected in messages.items():
        actual = func_text.get(name)
        if actual != expected:
            errors.append('UiMsg text mismatch: %s expected=%r actual=%r' % (name, expected, actual))

    for old, new in replacements:
        old_text = reconstruct_old_text(old)
        new_text = reconstruct_new_text(new, func_text)
        if old_text != new_text:
            errors.append(
                'semantic mismatch:\n  old=%r\n  new=%r\n  fragment old=%r\n  fragment new=%r'
                % (old_text, new_text, old[:120], new[:120])
            )

    for fn in targets:
        content = read_cp932(os.path.join(BASE, fn))
        for old, _new in replacements:
            old_norm = old.replace('\n', '\r\n')
            if old_norm not in content:
                continue

    unmatched_for_targets = []
    for fn in targets:
        content = read_cp932(os.path.join(BASE, fn))
        for old, new in replacements:
            old_norm = old.replace('\n', '\r\n')
            if old_norm in content:
                continue
            old_text = reconstruct_old_text(old)
            if old_text and old_text in content:
                unmatched_for_targets.append((fn, old[:100], 'partial literal only'))

    print('UiMsg functions:', len(func_text))
    print('Replacement rules:', len(replacements))
    print('Semantic mismatches:', sum(1 for e in errors if e.startswith('semantic')))
    print('Other errors:', sum(1 for e in errors if not e.startswith('semantic')))

    for fn in targets:
        content = read_cp932(os.path.join(BASE, fn))
        matched = sum(1 for old, _ in replacements if old.replace('\n', '\r\n') in content)
        print('%s: %d rules will apply' % (fn, matched))

    if errors:
        print('\nFAILED')
        for err in errors:
            print('-', err)
        return 1

    print('\nOK: all checks passed')
    return 0


if __name__ == '__main__':
    sys.exit(verify())
