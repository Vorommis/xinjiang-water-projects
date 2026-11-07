# -*- coding: utf-8 -*-
"""
把《奇旺咨询2025年水利项目招标情况.xlsx》中【标绿】的 12 个项目并入信息库。
- 已在库：补全空缺字段 + 追记招标核查结论（不覆盖已有内容）
- 不在库：新建记录
- 县名未带地州的，已通过网络检索确认归属
"""
import json, re, sqlite3, sys
sys.stdout.reconfigure(encoding='utf-8')
TAG = '【奇旺咨询2025招标核查】'

# 网络检索确认的 县/市 -> 地州
GEO = {
    '和田地区洛浦县': ('和田地区', '洛浦县'),
    '策勒县': ('和田地区', '策勒县'),
    '乌什县': ('阿克苏地区', '乌什县'),
    '青河县': ('阿勒泰地区', '青河县'),
    '疏附县': ('喀什地区', '疏附县'),
    '阿克苏市': ('阿克苏地区', '阿克苏市'),
}
# 源项目名 -> 库内项目名（人工核对）
MATCH = {
    '阿克苏地区乌什县永宁干渠建设项目': '阿克苏地区乌什县永宁干渠建设项目',
    '新疆和田地区洛浦县防沙治沙水源工程（一期）': '和田地区洛浦县防沙治沙水源工程（一期）',
    '策勒县民战渠中型引调水工程项目': '策勒县民战渠中型引调水工程',
    '阿克苏地区乌什县北山四条水系引调水工程': '阿克苏地区乌什县北山四条水系引调水工程',
    '阿克苏地区乌什县红山调蓄水池工程': '阿克苏地区乌什县红山调蓄水池工程',
}


def v(x):
    s = '' if x is None else str(x).strip()
    return '' if s.lower() == 'nan' or s == '/' else s


con = sqlite3.connect('xinjiang_water_projects.db'); con.row_factory = sqlite3.Row
rows = [dict(r) for r in con.execute('SELECT * FROM projects')]
byname = {}
for r in rows:
    byname.setdefault(re.sub(r'\s+', '', r['项目名称']), r)

data = json.load(open('_green12.json', encoding='utf-8'))
upd, new = [], []

for d in data:
    name = v(d['项目名称'])
    target = MATCH.get(name)
    old = byname.get(re.sub(r'\s+', '', target)) if target else None

    bond = v(d['债券资金(万元)'])
    cat = v(d['项目类别'])
    bid = v(d['招标情况'])
    unit = v(d['项目单位'])
    note_bits = [f'{TAG}招标情况：{bid or "-"}']
    if cat:      note_bits.append(f'项目类别：{cat}')
    if bond:     note_bits.append(f'债券资金：{bond}万元')
    if unit:     note_bits.append(f'项目单位：{unit}')
    if v(d['备注']):  note_bits.append(v(d['备注']))
    note = '；'.join(note_bits)

    if old:
        sets, vals = [], []
        if unit and v(old.get('项目法人')) in ('', '-'):
            sets.append('"项目法人"=?'); vals.append(unit)
        if cat and v(old.get('建设内容及规模')) in ('', '-'):
            pass
        old_note = v(old.get('备注'))
        if old_note in ('-', '—'):
            old_note = ''
        sets.append('"备注"=?')
        vals.append((old_note + ' | ' + note) if old_note else note)
        vals.append(old['row_id'])
        con.execute(f'UPDATE projects SET {",".join(sets)} WHERE row_id=?', vals)
        upd.append((name, old['项目名称'], bid))
        continue

    region, site = GEO.get(v(d['所在地区']), ('', ''))
    rec = {k: '' for k in (
        'id', '地州', '项目类型', '分类号', '项目名称', '建设内容及规模', '总投资_万元',
        '政府投资_万元', '其他投资_万元', '年2026预计完成投资_万元', '年2026政府投资_万元',
        '年2026其他投资_万元', '计划开工时间', '计划竣工时间', '建设地点', '项目法人',
        '项目单位性质', '联系人及联系方式', 'created_at', '资金来源', '设计单位', '招标代理',
        '项目进展情况', '关键关系资源', '关系资源级别', '关系资源备注', '业主决策层联系人',
        '业主决策层职务', '业主参与层联系人', '业主参与层职务', '业主操作层联系人',
        '业主操作层职务', '政府实施单位', '公共资源中心联系人', '主要竞争对手', '评标专家',
        '投标策略要点', '策划责任人', '市场经理', '攻关状态', '下一步行动', '预计招标时间',
        '中标概率(%)', '备注')}
    rec['项目名称'] = name
    rec['地州'] = region
    rec['建设地点'] = site
    rec['项目法人'] = unit
    rec['总投资_万元'] = v(d['总投资(万元)'])
    rec['资金来源'] = '地方政府专项债券' if bond else ''
    rec['项目类型'] = '储备'
    rec['分类号'] = '早期' if bid == '未招标' else '可研/设计阶段'
    rec['项目进展情况'] = ('未招标（前期储备）' if bid == '未招标'
                       else '已完成设计/可研招标，施工标未释放' if bid else '')
    rec['备注'] = note
    try:
        rec['总投资数值'] = float(rec['总投资_万元']) if rec['总投资_万元'] else None
    except ValueError:
        rec['总投资数值'] = None
    rec['原始序号'] = None
    rec['当前看板'] = 1
    rec['重点跟踪'] = 0
    rec['数据来源'] = '奇旺咨询2025招标情况(标绿)'
    rec['数据质量'] = 'ok'
    cols = list(rec.keys())
    ph = ','.join('?' * len(cols))
    con.execute(f'INSERT INTO projects ({",".join(chr(34)+c+chr(34) for c in cols)}) VALUES ({ph})',
                [rec[c] for c in cols])
    new.append((name, region, site, rec['总投资_万元'], bid))

con.commit()
tot = con.execute('SELECT COUNT(*) FROM projects').fetchone()[0]
cur = con.execute('SELECT COUNT(*) FROM projects WHERE 当前看板=1').fetchone()[0]
con.close()

print(f'更新已有 {len(upd)} 条:')
for a, b, bid in upd:
    print(f'   ~ {a[:32]:34} -> {b[:32]:34} [{bid}]')
print(f'\n新增 {len(new)} 条:')
for n, g, s, inv, bid in new:
    print(f'   + {n[:40]:42} {g:<6} {s:<8} {str(inv):>10}万  {bid}')
print(f'\n库内合计 {tot} 条（当前有效 {cur}）')
