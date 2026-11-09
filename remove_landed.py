# -*- coding: utf-8 -*-
"""
清除"施工/EPC 已落地"的项目（移出储备库，保留记录备查）
- 依据：自治区公共资源交易平台全文检索 + 《奇旺咨询2025年水利项目招标情况》
- 克州乔诺水库工程：仅进场道路及跨河桥中标，主体未招标 -> 按要求保留
"""
import json, sqlite3, sys
sys.stdout.reconfigure(encoding='utf-8')

# 来自 API 复核的命中（保留克州乔诺水库）
API_REMOVE = [
    '新疆昌吉市三屯河灌区现代化改造工程',
    '柯坪县水质净化厂建设项目',
    '新疆和田地区皮山县2025年木奎拉乡沉沙池工程',
    '莎车县勿甫干渠改建项目',
    '阿图什市城市供水水源地工程',
    '沙湾市窝巴特渠首引水能力提升工程',
    '伊宁市伊犁河湿地水生态修复建设工程',
    '伊犁州伊宁市东城区地表水厂及管网工程建设项目',
    '新疆伊犁州巩留县东风中型灌区续建配套与现代化改造项目',
    '和田地区洛浦县英兰干片区供水工程',
]
# 来自《奇旺咨询》表「已中标」且仍在库中的
XLSX_REMOVE = ['阿克苏市萨克沙克等三座水厂提升改造工程']

# 收集证据
ev = {}
for line in open('_api_out.jsonl', encoding='utf-8'):
    d = json.loads(line)
    if d['命中']:
        h = d['命中'][0]
        ev[d['项目名称']] = f"{h['日期']}《{h['标题'][:70]}》"
qw = {x['项目名称']: x for x in json.load(open('_qw_all.json', encoding='utf-8'))}
qw_ev = {}
for name in XLSX_REMOVE:
    for k, v in qw.items():
        if name[:10] in k:
            qw_ev[name] = f"中标单位：{v.get('中标单位','')}；中标金额：{v.get('中标金额(万元)','')}万元"
            break

con = sqlite3.connect('xinjiang_water_projects.db'); con.row_factory = sqlite3.Row
done, skipped = [], []
for name in API_REMOVE + XLSX_REMOVE:
    r = con.execute('SELECT row_id,项目名称,当前看板,备注,数据质量 FROM projects WHERE 项目名称=?', (name,)).fetchone()
    if not r:
        skipped.append((name, '库中未找到'))
        continue
    if r['当前看板'] != 1:
        skipped.append((name, '已不在看板'))
        continue
    note = ev.get(name) or qw_ev.get(name) or ''
    old = str(r['备注'] or '').strip()
    if old in ('-', '—'):
        old = ''
    add = f'【招标复核2026-10】施工/EPC 已落地，移出储备库' + (f'；证据：{note}' if note else '')
    newnote = (old + ' | ' + add) if old else add
    con.execute('UPDATE projects SET 当前看板=0, 数据质量=?, 备注=? WHERE row_id=?',
                ('已移出：施工/EPC已落地', newnote, r['row_id']))
    done.append((name, note))

con.commit()
cur = con.execute('SELECT COUNT(*) FROM projects WHERE 当前看板=1').fetchone()[0]
tot = con.execute('SELECT COUNT(*) FROM projects').fetchone()[0]
con.close()

print(f'已移出 {len(done)} 条:')
for n, e in done:
    print(f'   × {n[:40]:42} {e[:56]}')
if skipped:
    print(f'\n跳过 {len(skipped)} 条:')
    for n, why in skipped:
        print(f'   - {n[:40]:42} ({why})')
print(f'\n库内合计 {tot} 条，当前有效 {cur} 条')
