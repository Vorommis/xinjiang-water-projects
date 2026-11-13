# -*- coding: utf-8 -*-
"""
把《水利工程项目汇总表_未招标清单》中未入库的项目录入信息库。
- 金额取「总投资(万元)」
- 「十五五投资」「建设类别」写入备注
- 建设内容/工程任务取自原始表《水利工程项目汇总表.xlsx》
"""
import json, re, sqlite3, sys
sys.stdout.reconfigure(encoding='utf-8')
REGION_MAP = {
    '伊犁哈萨克自治州': '伊犁州', '巴音郭楞蒙古自治州': '巴州',
    '博尔塔拉蒙古自治州': '博州', '克孜勒苏柯尔克孜自治州': '克州',
    '昌吉回族自治州': '昌吉州',
}
CAT_MAP = {
    '未检索到公告': '早期',
    '仅设计/可研招标': '可研/设计阶段',
    '仅服务类招标': '可研/设计阶段',
}


def v(x):
    s = '' if x is None else str(x).strip()
    return '' if s in ('None', 'nan', '—', '-', '/') else s


new = json.load(open('_new_projects.json', encoding='utf-8'))
# 0) 修正最后一条的行政归属（网络检索确认：新疆伊犁州新源县吐尔根沟水库工程）
for r in new:
    if '吐尔根沟' in r['项目名称']:
        r['所属地州'] = '伊犁哈萨克自治州'
        r['建设地点'] = '新源县'

con = sqlite3.connect('xinjiang_water_projects.db'); con.row_factory = sqlite3.Row
inserted = 0
for r in new:
    d = r.get('_detail') or {}
    region = REGION_MAP.get(v(r.get('所属地州')), v(r.get('所属地州')))
    scene = v(r.get('建设状态')) or v(d.get('建设状态')) or '储备'
    concl = v(r.get('核查结论'))
    cat = CAT_MAP.get(concl, '早期')
    inv = v(r.get('总投资(万元)')) or v(d.get('总投资(万元)'))
    inv15 = v(r.get('十五五投资(万元)')) or v(d.get('十五五投资(万元)'))
    cls = v(r.get('类别')) or v(d.get('建设类别'))

    bits = ['【汇总表录入】']
    if inv15:
        bits.append(f'十五五投资：{inv15}万元')
    if cls:
        bits.append(f'建设类别：{cls}')
    if v(d.get('工程任务')):
        bits.append(f"工程任务：{v(d['工程任务'])}")
    if concl:
        bits.append(f'核查结论：{concl}')
    if v(d.get('十五五其中中央投资(万元)')):
        bits.append(f"十五五中央投资：{v(d['十五五其中中央投资(万元)'])}万元")
    if v(d.get('其中中央投资(万元)')):
        bits.append(f"总投资中中央投资：{v(d['其中中央投资(万元)'])}万元")
    note = '；'.join(bits)

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
    rec['项目名称'] = v(r['项目名称'])
    rec['地州'] = region
    rec['建设地点'] = v(r.get('建设地点')) or '-'
    rec['项目类型'] = scene
    rec['分类号'] = cat
    rec['建设内容及规模'] = v(d.get('建设内容及规模')) or '-'
    rec['总投资_万元'] = inv
    rec['项目进展情况'] = concl or '未检索到公告'
    rec['备注'] = note
    for f in ('项目法人', '项目单位性质', '联系人及联系方式', '资金来源', '设计单位',
              '招标代理', '计划开工时间', '计划竣工时间'):
        rec[f] = '-'
    try:
        rec['总投资数值'] = float(inv) if inv else None
    except ValueError:
        rec['总投资数值'] = None
    rec['id'] = None
    rec['原始序号'] = None
    rec['当前看板'] = 1
    rec['重点跟踪'] = 0
    rec['数据来源'] = '水利工程项目汇总表_未招标清单'
    rec['数据质量'] = 'ok' if inv else 'ok(总投资待补)'
    cols = list(rec.keys())
    con.execute(f'INSERT INTO projects ({",".join(chr(34)+c+chr(34) for c in cols)}) '
                f'VALUES ({",".join("?"*len(cols))})', [rec[c] for c in cols])
    inserted += 1

con.commit()
tot = con.execute('SELECT COUNT(*) FROM projects').fetchone()[0]
cur = con.execute('SELECT COUNT(*) FROM projects WHERE 当前看板=1').fetchone()[0]
con.close()
print(f'已录入 {inserted} 条')
print(f'库内合计 {tot} 条，当前有效 {cur} 条')
