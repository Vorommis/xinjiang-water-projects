# -*- coding: utf-8 -*-
"""
把三份"项目进展/招标核查"文件整合进信息库。

排除规则：已进入施工阶段的（施工已定标 / 施工招标中 / 主体在建）一律排除。
其余：已存在于库中的 -> 更新备注（记录核查结论）；新项目 -> 入库为当前有效记录。

来源：
  1. 和田水利项目进展核查.xlsx        (和田地区 27 -> 排除1 保留26)
  2. 阿克苏2027储备清单-水利5000万以上-项目进展核查.xlsx (25 -> 排除12 保留13)
  3. 新疆老旧水库改造提升先期实施项目招标情况核查报告.docx (21 -> 排除2 保留19)
"""
import pandas as pd, docx, sqlite3, json, re, sys, os
sys.stdout.reconfigure(encoding='utf-8')
D = r'C:\Users\Admin（无密码）\Desktop'
DB = 'xinjiang_water_projects.db'
CHECK_TAG = '【招标核查2026-10】'

REGION_MAP = {'昌吉回族自治州': '昌吉州', '巴音郭楞蒙古自治州': '巴州',
              '兵团第八师石河子市': '兵团八师', '待核实': ''}


def norm(s):
    s = re.sub(r'\s+', '', str(s or ''))
    for a, b in [('（', '('), ('）', ')'), ('－', '-'), ('—', '-'), ('―', '-')]:
        s = s.replace(a, b)
    return s.strip()


def v(x):
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ''
    s = str(x).strip()
    return '' if s.lower() == 'nan' else s


def clean_cols(df):
    """表头去掉换行/空格，避免取列名失败"""
    df.columns = [re.sub(r'\s+', '', str(c)) for c in df.columns]
    return df


def is_awarded(stage):
    s = str(stage or '')
    return ('施工已定标' in s) or ('施工招标中' in s)


def stage_to_cat(s):
    s = str(s or '')
    if '初步' in s:
        return '初步阶段'
    if '早期' in s or '纯新建' in s:
        return '早期'
    return '可研/设计阶段'


# ---------------------------------------------------------------- 读三份来源
def load_hetian():
    prog = clean_cols(pd.read_excel(D + r'\和田水利项目进展核查.xlsx', sheet_name='项目进展总表'))
    det = clean_cols(pd.read_excel(D + r'\和田(1).xls', sheet_name='Sheet1'))
    det = det[det['所属行业'].astype(str).str.strip() == '水利']
    detm = {norm(r['项目名称']): r for _, r in det.iterrows()}
    out = []
    for _, r in prog.iterrows():
        st = v(r['进展阶段'])
        if is_awarded(st):
            continue
        nm = v(r['项目名称'])
        d = detm.get(norm(nm))
        out.append({'来源': '和田进展核查', '地州': '和田地区', '项目名称': nm,
                    '核查阶段': st, '公告数': v(r['平台公告数']),
                    '最新日期': v(r['最新公告日期']), '最新标题': v(r['最新公告标题']),
                    '总投资': v(d['总投资']) if d is not None else v(r['总投资(万元)']),
                    '建设内容': v(d['主要建设内容与规模']) if d is not None else '',
                    '资金来源': v(d['资金来源']) if d is not None else '',
                    '开工': v(d['预计开工时间']) if d is not None else '',
                    '竣工': v(d['预计完成时间']) if d is not None else '',
                    '地点': v(d['建设地点']) if d is not None else v(r.get('县市')),
                    '原进展': v(d['目前进度']) if d is not None else '',
                    '原备注': v(d['备注']) if d is not None else ''})
    return out


def load_akesu():
    prog = clean_cols(pd.read_excel(D + r'\阿克苏2027储备清单-水利5000万以上-项目进展核查.xlsx', sheet_name='项目进展总表'))
    det = clean_cols(pd.read_excel(D + r'\（地区水利局梳理）附件2：阿克苏地区2027年项目储备清单9.3.xlsx', header=1))
    det = det.iloc[5:46]
    det = det[det['项目名称'].notna()]
    detm = {norm(v(r['项目名称'])): r for _, r in det.iterrows()}
    out = []
    for _, r in prog.iterrows():
        st = v(r['进展阶段'])
        if is_awarded(st):
            continue
        nm = v(r['项目名称'])
        d = detm.get(norm(nm))
        yrs = v(d['建设起止年限']) if d is not None else ''
        ys = re.split(r'[-—~至]', yrs) if yrs else []
        out.append({'来源': '阿克苏进展核查', '地州': '阿克苏地区', '项目名称': nm,
                    '核查阶段': st, '公告数': v(r['平台公告数']),
                    '最新日期': v(r['最新公告日期']), '最新标题': v(r['最新公告标题']),
                    '总投资': v(d['总投资（万元）']) if d is not None else v(r['总投资(万元)']),
                    '建设内容': v(d['建设内容及规模']) if d is not None else '',
                    '资金来源': v(d['主要资金来源']) if d is not None else '',
                    '建设性质': v(d['建设性质']) if d is not None else '',
                    '开工': ys[0] if ys else '', '竣工': ys[-1] if len(ys) > 1 else '',
                    '年2026预计': v(d['其中：2026年预计完成投资（万元）']) if d is not None else '',
                    '地点': v(d['县市']) if d is not None else v(r.get('县市')),
                    '原进展': '', '原备注': v(d['备注']) if d is not None else ''})
    return out


def load_reservoir():
    doc = docx.Document(D + r'\新疆老旧水库改造提升先期实施项目招标情况核查报告.docx')
    t1, t2 = doc.tables[1], doc.tables[2]
    info = {}
    for row in t1.rows[1:]:
        c = [x.text.strip() for x in row.cells]
        info[norm(c[1])] = {'库容': c[4], '规模': c[3], '建成': c[5], '初设': c[7],
                            '理由': c[8], '措施': c[9], '前期': c[10], '备注': c[11]}
    out = []
    for row in t2.rows[1:]:
        c = [x.text.strip() for x in row.cells]
        nm, region, measure, status, scope = c[1], c[2], c[3], c[4], c[5]
        if norm(scope).startswith('施工') or 'EPC' in scope.upper():
            continue                      # 施工已定标/招标中 -> 排除
        if nm == '塔里木水库':
            continue                      # 与"尉犁县塔里木水库提升工程"疑重复，见报告表3
        i = info.get(norm(nm), {})
        out.append({'来源': '老旧水库招标核查', '地州': REGION_MAP.get(region, region),
                    '项目名称': nm, '核查阶段': f'{status}｜{scope}',
                    '公告数': '', '最新日期': c[6], '最新标题': '',
                    '总投资': '', '建设内容': f"{measure}；总库容{i.get('库容','')}；{i.get('规模','')}",
                    '资金来源': '', '开工': '', '竣工': '',
                    '地点': REGION_MAP.get(region, region),
                    '原进展': '', '原备注': i.get('备注', ''),
                    '措施': measure, '初设': i.get('初设', ''), '前期': i.get('前期', ''),
                    '中标人': c[7], '理由': i.get('理由', '')})
    return out


# ---------------------------------------------------------------- 手工核对后的匹配表
MATCH = {
    # 和田
    '和康县吉良镇供水保障工程': '和田地区和康县吉良镇供水保障工程',
    '和田地区和康县山洪沟综合治理项目': '和田地区和康县山洪沟综合治理',
    '墨玉县喀尔赛水库清淤工程': '新疆和田地区墨玉县喀尔赛水库清淤增蓄工程',
    '墨玉县东风水库清淤工程': '新疆和田地区墨玉县东风水库清淤增蓄工程',
    # 水库
    '福海县团结水库': '福海县团结水库除险加固工程',
    '哈密市榆树沟水库': '哈密市榆树沟水库除险加固提升工程',
    '和田市尕宗水库': '和田市尕宗水库清淤工程',
    '和田市昆仑水库': '和田市昆仑水库清淤工程',
    '墨玉县东风水库': '墨玉县东风水库除险加固工程',
    '疏勒县艾尔木东水库': '疏勒县艾尔木东水库改扩建项目',
    '吐鲁番市鄯善县柯柯亚水库': '吐鲁番市鄯善县柯柯亚水库除险加固工程',
    '新疆策勒县胜利水库': '新疆策勒县胜利水库除险加固工程',
    '新疆富蕴县喀拉通克水库': '新疆富蕴县喀拉通克水库除险加固工程',
    '新疆和田地区墨玉县喀尔赛水库清淤增蓄': '新疆和田地区墨玉县喀尔赛水库清淤增蓄工程',
}


# ---------------------------------------------------------------- 组装记录
def note_of(c):
    bits = [f"阶段：{c['核查阶段']}"]
    if c.get('公告数'):
        bits.append(f"公告数：{c['公告数']}")
    if c.get('最新日期'):
        bits.append(f"最新：{c['最新日期']}" + (f"《{c['最新标题']}》" if c.get('最新标题') else ''))
    if c.get('中标人') and c['中标人'] not in ('—', '-'):
        bits.append(f"中标人：{c['中标人']}")
    return f"{CHECK_TAG}" + '；'.join(bits)


def build_new(c):
    """把一份候选组装成数据库记录"""
    src = c['来源']
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
    rec['地州'] = c['地州']
    rec['项目名称'] = c['项目名称']
    rec['建设内容及规模'] = c.get('建设内容', '')
    rec['总投资_万元'] = c.get('总投资', '')
    rec['资金来源'] = c.get('资金来源', '')
    rec['计划开工时间'] = c.get('开工', '')
    rec['计划竣工时间'] = c.get('竣工', '')
    rec['建设地点'] = c.get('地点', '')
    rec['年2026预计完成投资_万元'] = c.get('年2026预计', '')

    if src == '和田进展核查':
        rec['项目类型'] = '续建' if '已开工' in c.get('原进展', '') else '储备'
        rec['分类号'] = '初步阶段' if '初步设计' in c.get('原进展', '') else '可研/设计阶段'
        rec['项目进展情况'] = f"{c.get('原进展','')}；招标核查：{c['核查阶段']}".strip('；')
    elif src == '阿克苏进展核查':
        rec['项目类型'] = c.get('建设性质') or '新建'
        rec['分类号'] = '可研/设计阶段'
        rec['项目进展情况'] = f"储备清单项目；招标核查：{c['核查阶段']}"
    else:                                                   # 老旧水库
        rec['项目类型'] = '改扩建'
        init = c.get('初设', '')
        rec['分类号'] = '初步阶段' if (init == '是' or (init and init not in ('否', '／', '-'))) else '可研/设计阶段'
        if init and init not in ('是', '否', '／', '-', ''):
            rec['设计单位'] = init                          # 该栏填的是设计单位名称
        prog = f"改造提升（{c.get('措施','')}）；招标核查：{c['核查阶段']}"
        if c.get('前期'):
            prog += f"；计划{c['前期']}完成前期"
        rec['项目进展情况'] = prog

    extra = []
    if c.get('理由'):
        extra.append(f"改造理由：{c['理由']}")
    if c.get('原备注'):
        extra.append(c['原备注'])
    rec['备注'] = '；'.join([note_of(c)] + extra)
    rec['原始序号'] = None
    rec['id'] = None
    rec['当前看板'] = 1
    rec['数据来源'] = f"{src}(2026-10)"
    rec['数据质量'] = 'ok' if rec['总投资_万元'] else 'ok(总投资待补)'
    rec['总投资数值'] = None
    try:
        rec['总投资数值'] = float(rec['总投资_万元']) if rec['总投资_万元'] else None
    except ValueError:
        rec['总投资数值'] = None
    return rec


def main():
    con = sqlite3.connect(DB)
    con.row_factory = sqlite3.Row
    rows = [dict(r) for r in con.execute('SELECT * FROM projects ORDER BY row_id')]
    by_name = {norm(r['项目名称']): r for r in rows}

    cands = load_hetian() + load_akesu() + load_reservoir()
    print(f'候选（已排除施工阶段）: {len(cands)} 条'
          f'  [和田 {len(load_hetian())} / 阿克苏 {len(load_akesu())} / 水库 {len(load_reservoir())}]')

    updated, inserted, seen = [], [], set()
    for c in cands:
        key = norm(c['项目名称'])
        target = MATCH.get(c['项目名称'])
        old = by_name.get(norm(target)) if target else by_name.get(key)

        if old:
            note = note_of(c)
            old_note = v(old.get('备注'))
            if old_note in ('-', '—', '－'):      # 库内用 '-' 表示空
                old_note = ''
            if norm(note) in norm(old_note):
                continue
            new_note = (old_note + ' | ' + note) if old_note else note
            con.execute('UPDATE projects SET 备注=? WHERE row_id=?', (new_note, old['row_id']))
            old['备注'] = new_note
            updated.append((old['项目名称'], c['来源'], c['核查阶段']))
        else:
            if key in seen:
                continue
            seen.add(key)
            rec = build_new(c)
            cols = [k for k in rec if k not in ('row_id',)]
            ph = ','.join('?' * len(cols))
            con.execute(f'INSERT INTO projects ({",".join(chr(34)+c2+chr(34) for c2 in cols)}) VALUES ({ph})',
                        [rec[k] for k in cols])
            inserted.append((rec['项目名称'], rec['地州'], rec['总投资_万元'], c['核查阶段'], c['来源']))

    con.commit()
    total = con.execute('SELECT COUNT(*) FROM projects').fetchone()[0]
    cur = con.execute('SELECT COUNT(*) FROM projects WHERE 当前看板=1').fetchone()[0]
    con.close()

    print(f'\n更新已有记录备注: {len(updated)} 条')
    for n, s, st in updated:
        print(f'   ~ {n[:40]:42} [{s}] {st}')
    print(f'\n新增记录: {len(inserted)} 条')
    for n, r, inv, st, s in inserted:
        print(f'   + {n[:40]:42} {r:<6} {str(inv):>10}万  {st}')
    print(f'\n库内合计: {total} 条（当前有效 {cur} 条）')


if __name__ == '__main__':
    main()


