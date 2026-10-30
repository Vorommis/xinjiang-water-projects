# -*- coding: utf-8 -*-
"""
从数据库重建加密看板 index.html。
保持原有的页面骨架、样式、排序/筛选脚本不变，只替换：
  · 表头统计数字
  · 筛选下拉（地州，必要时补新地州）
  · 表格 tbody 全部数据行
最后用原密码重新做 AES-256-GCM 加密。
"""
import base64, hashlib, html, os, re, sqlite3, sys
from cryptography.hazmat.primitives.ciphers.aead import AESGCM

sys.stdout.reconfigure(encoding='utf-8')
PASSWORD = 'sl888888'
INDEX = 'index.html'
DB = 'xinjiang_water_projects.db'

# 分类号 -> (span class, 显示文字)
CAT_STYLE = {
    '可研/设计阶段': ('tag td', '可研/设计'),
    '初步阶段': ('tag te', '初步阶段'),
    '纯新建': ('tag tn', '纯新建'),
}
# 表格列顺序：(数据库字段, 是否用 title 原值, 显示处理)
COLS = [
    ('__seq__', ''), ('地州', ''), ('项目类型', ''), ('分类号', 'cat'),
    ('项目名称', ''), ('建设内容及规模', 'trunc38'), ('总投资数值', 'money'),
    ('资金来源', ''), ('计划开工时间', ''), ('建设地点', ''), ('项目法人', ''),
    ('项目单位性质', ''), ('联系人及联系方式', ''), ('设计单位', ''), ('招标代理', ''),
    ('项目进展情况', ''), ('关键关系资源', ''), ('关系资源级别', ''), ('关系资源备注', ''),
    ('业主决策层联系人', ''), ('业主决策层职务', ''), ('业主参与层联系人', ''),
    ('业主参与层职务', ''), ('业主操作层联系人', ''), ('业主操作层职务', ''),
    ('公共资源中心联系人', ''), ('主要竞争对手', ''), ('评标专家', ''),
    ('投标策略要点', ''), ('策划责任人', ''), ('市场经理', ''), ('攻关状态', ''),
    ('下一步行动', ''), ('预计招标时间', ''), ('中标概率(%)', ''), ('备注', ''),
    ('__blank__', ''),
]


def dec(html_text, password):
    m = re.search(r'var ED="([A-Za-z0-9+/=]+)"', html_text)
    raw = base64.b64decode(m.group(1))
    key = hashlib.sha256(password.encode()).digest()
    return AESGCM(key).decrypt(raw[:12], raw[12:-16] + raw[-16:], None).decode('utf-8')


def enc(plain, password):
    key = hashlib.sha256(password.encode()).digest()
    nonce = os.urandom(12)
    ct = AESGCM(key).encrypt(nonce, plain.encode('utf-8'), None)
    return base64.b64encode(nonce + ct).decode('ascii')


def money(v):
    try:
        n = float(v)
    except (TypeError, ValueError):
        return '-'
    if n >= 10000:
        return f'{n / 10000:.1f}亿'
    return f'{n:.0f}万'


def esc(s):
    return html.escape(str(s), quote=True)


def cell_text(field, rec):
    if field == '__seq__':
        return str(rec['_seq'])
    if field == '__blank__':
        return '-'
    v = rec.get(field)
    if v is None:
        return '-'
    v = str(v).strip()
    return v if v and v != '-' else '-'


def build_row(rec):
    parts = []
    for field, mode in COLS:
        if field == '分类号':
            cat = str(rec.get('分类号') or '可研/设计阶段')
            cls, disp = CAT_STYLE.get(cat, ('tag td', cat.replace('阶段', '')))
            parts.append(f'<td><span class="{cls}">{esc(disp)}</span></td>')
            continue
        txt = cell_text(field, rec)
        if mode == 'money':
            display = money(rec.get('总投资数值'))
            raw = cell_text('总投资_万元', rec)
            parts.append(f'<td title="{esc(raw)}">{esc(display)}</td>')
        elif mode == 'trunc38':
            display = txt if len(txt) <= 38 else txt[:38] + '...'
            parts.append(f'<td title="{esc(txt)}">{esc(display)}</td>')
        else:
            parts.append(f'<td title="{esc(txt)}">{esc(txt)}</td>')
    region = esc(str(rec.get('地州') or ''))
    cat = esc(str(rec.get('分类号') or ''))
    ptype = esc(str(rec.get('项目类型') or ''))
    return (f'<tr data-region="{region}" data-cat="{cat}" data-ptype="{ptype}">'
            + ''.join(parts) + '</tr>')


def main():
    src = open(INDEX, encoding='utf-8').read()
    plain = dec(src, PASSWORD)
    print(f'解密后明文 {len(plain):,} 字符')

    con = sqlite3.connect(DB)
    con.row_factory = sqlite3.Row
    recs = [dict(r) for r in con.execute('SELECT * FROM projects WHERE 当前看板=1 ORDER BY row_id')]
    con.close()
    for i, r in enumerate(recs, 1):
        r['_seq'] = i
    print(f'当前有效记录 {len(recs)} 条')

    # ---- 1) 表格 ----
    body = ''.join(build_row(r) for r in recs)
    s = plain.index('<tbody id="tb">') + len('<tbody id="tb">')
    e = plain.index('</tbody>', s)
    plain = plain[:s] + body + plain[e:]

    # ---- 2) 统计数字 ----
    n_all = len(recs)
    n_pure = sum(1 for r in recs if str(r.get('分类号')) == '纯新建')
    n_ky = sum(1 for r in recs if str(r.get('分类号')) == '可研/设计阶段')
    n_cb = sum(1 for r in recs if str(r.get('分类号')) == '初步阶段')
    n_region = len({str(r.get('地州')) for r in recs if r.get('地州')})

    def set_stat(label, value):
        nonlocal plain
        plain = re.sub(r'<div class="n">\d+</div><div class="l">' + label + '</div>',
                       f'<div class="n">{value}</div><div class="l">{label}</div>', plain)

    set_stat('总项目', n_all)
    set_stat('纯新建', n_pure)
    set_stat('可研/设计', n_ky)
    set_stat('初步阶段', n_cb)
    set_stat('地州/市', n_region)
    plain = re.sub(r'(\|\s*)\d+(\s*个项目)', lambda m: f'{m.group(1)}{n_all}{m.group(2)}', plain)
    plain = re.sub(r'<div class="info" id="cnt">\d+\s*/\s*\d+\s*条</div>',
                   f'<div class="info" id="cnt">{n_all} / {n_all} 条</div>', plain)

    # ---- 3) 下拉框：保留原选项文字，只补新出现的值 ----
    def rebuild_select(select_id, values, text_key=None):
        nonlocal plain
        mm = re.search(r'(<select id="' + select_id + r'"[^>]*>)(.*?)(</select>)', plain, re.S)
        pairs = re.findall(r'<option value="([^"]*)">([^<]*)</option>', mm.group(2))
        label = {v: lab for v, lab in pairs}
        keep = [v for v, _ in pairs]
        for v in values:
            if v and v not in keep:
                keep.append(v)
                label.setdefault(v, v)
        html_opts = ''.join(
            f'<option value="{esc(v)}">{esc(label.get(v) or v)}</option>' for v in keep)
        plain = plain[:mm.start(2)] + html_opts + plain[mm.end(2):]

    regions, types = [], []
    for r in recs:
        g = str(r.get('地州') or '').strip()
        if g and g not in regions:
            regions.append(g)
        p_ = str(r.get('项目类型') or '').strip()
        if p_ and p_ not in types:
            types.append(p_)

    rebuild_select('rg', regions)
    rebuild_select('tp', types)

    # ---- 4) 重新加密写回 ----
    new_ed = enc(plain, PASSWORD)
    out = re.sub(r'var ED="[A-Za-z0-9+/=]+"', f'var ED="{new_ed}"', src)
    open(INDEX, 'w', encoding='utf-8').write(out)
    print(f'已写回 {INDEX}  ({os.path.getsize(INDEX):,} bytes)')
    print(f'统计: 总{n_all} 纯新建{n_pure} 可研/设计{n_ky} 初步{n_cb} 地州{n_region}')

    # ---- 5) 自检 ----
    chk = dec(open(INDEX, encoding='utf-8').read(), PASSWORD)
    from bs4 import BeautifulSoup
    soup = BeautifulSoup(chk, 'lxml')
    rows = soup.select('#tb tr')
    print(f'自检: 解密后行数 = {len(rows)} (应为 {n_all})')
    names = [tr.find_all('td')[4].get('title') for tr in rows]
    dbnames = [str(r['项目名称']) for r in recs]
    print(f'自检: 项目名称逐条一致 = {names == dbnames}')


if __name__ == '__main__':
    main()
