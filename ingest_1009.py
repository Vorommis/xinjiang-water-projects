# -*- coding: utf-8 -*-
"""
把《项目信息录入10.9.xlsx》并入信息库。
- 已存在的项目：补全空缺字段 + 追加入库说明，不覆盖已有内容
- 新项目：新建记录
- 已放弃(项目状态=已放弃)的项目：不录入
"""
import pandas as pd, sqlite3, re, sys, difflib
sys.stdout.reconfigure(encoding='utf-8')
P = r'C:\Users\Admin（无密码）\Desktop\项目信息录入10.9.xlsx'
DB = 'xinjiang_water_projects.db'
TAG = '【10.9录入】'


def norm(s):
    s = re.sub(r'\s+', '', str(s or ''))
    for a, b in [('（', '('), ('）', ')'), ('－', '-'), ('—', '-'), ('―', '-'), ('·', '')]:
        s = s.replace(a, b)
    return s.strip()


def core(s):
    s = norm(s)
    s = re.sub(r'^(新疆|喀什地区|阿克苏地区|阿勒泰地区|昌吉州|巴州|和田地区|塔城地区|乌鲁木齐市)', '', s)
    s = re.sub(r'(建设项目|工程|项目)$', '', s)
    return s


def v(x):
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ''
    s = str(x).strip()
    return '' if s.lower() == 'nan' else s


REGION_PREFIX = [
    ('喀什地区', '喀什地区'), ('阿克苏地区', '阿克苏地区'), ('阿勒泰地区', '阿勒泰地区'),
    ('阿勒泰市', '阿勒泰地区'), ('和田地区', '和田地区'), ('昌吉州', '昌吉州'),
    ('乌鲁木齐市', '乌鲁木齐市'), ('乌鲁木齐县', '乌鲁木齐市'), ('巴州', '巴州'),
    ('塔城新区', '塔城地区'), ('塔城地区', '塔城地区'), ('伊犁', '伊犁州'),
    ('哈密', '哈密市'), ('吐鲁番', '吐鲁番市'), ('博州', '博州'), ('克州', '克州'),
    ('克拉玛依', '克拉玛依市'), ('兵团', '兵团'),
]


def region_of(loc, name):
    for k, v_ in REGION_PREFIX:
        if k in loc:
            return v_
    for k, v_ in REGION_PREFIX:
        if k in name:
            return v_
    return ''


# 手工核对过的匹配表（源项目名 -> 库内项目名）
MATCH = {
    '塔里木水库提升工程': '尉犁县塔里木水库提升工程',
    '老大河灌区续建配套与现代化改造工程': '阿克苏地区老大河灌区续建配套与现代化改造项目',
    '马圈沟水库工程': '乌鲁木齐县马圈沟水库工程',
    '和静县哈哈沟水库工程': '和静县哈哈沟水库工程',
    '喀什疏勒县艾尔木东水库改扩建工程': '疏勒县艾尔木东水库改扩建项目',
    '卡布曲尔河向加依落萨伊河引调水工程': '新疆阿克苏地区柯坪县卡布曲尔河向加依洛萨依河调水工程',
    '富蕴县杜热灌区续建配套与节水改造项目': '新疆阿勒泰地区富蕴县杜热灌区续建配套与节水改造项目',
}
# 上面那条写错时回退用
MATCH['卡布曲尔河向加依落萨伊河引调水工程'] = '新疆阿克苏地区柯坪县卡布曲尔河向加依洛萨依河引调水工程'

# 源字段 -> 库字段
FIELD_MAP = {
    '建设地点': '项目地点', '项目法人': '建设单位', '设计单位': '设计单位',
    '资金来源': '资金来源', '投标策略要点': '投标策略要点',
    '关键关系资源': '关键关系资源', '关系资源级别': '关系资源级别', '关系资源备注': '关系资源备注',
    '业主决策层联系人': '业主决策层联系人', '业主决策层职务': '业主决策层职务',
    '业主参与层联系人': '业主参与层联系人', '业主参与层职务': '业主参与层职务',
    '业主操作层联系人': '业主操作层联系人', '业主操作层职务': '业主操作层职务',
    '公共资源中心联系人': '公共资源中心联系人', '主要竞争对手': '主要竞争对手',
    '评标专家': '评标专家', '策划责任人': '策划责任人', '市场经理': '主跟踪人',
    '攻关状态': '攻关状态', '下一步行动': '下一步行动', '预计招标时间': '预计招标时间',
    '中标概率(%)': '中标概率',
}


def load():
    df = pd.read_excel(P, sheet_name='项目录入')
    df.columns = [str(c).strip() for c in df.columns]
    df = df[df['项目名称'].notna()]
    df = df[~df['项目名称'].astype(str).str.startswith('【示例】')]
    total = len(df)
    df = df[df['项目状态'].astype(str).str.strip() != '已放弃']
    print(f'表内 {total} 条，已放弃 {total - len(df)} 条不录入，实际处理 {len(df)} 条')
    return df


def main():
    df = load()
    con = sqlite3.connect(DB)
    con.row_factory = sqlite3.Row
    rows = [dict(r) for r in con.execute('SELECT * FROM projects ORDER BY row_id')]
    by_name = {norm(r['项目名称']): r for r in rows}

    upd, new = [], []
    for _, r in df.iterrows():
        name = v(r['项目名称'])
        target = MATCH.get(name)
        old = by_name.get(norm(target)) if target else by_name.get(norm(name))

        note_bits = [f"{TAG}状态：{v(r['项目状态']) or '-'}"]
        if v(r['项目类别']):        note_bits.append(f"类别：{v(r['项目类别'])}")
        if v(r['项目进展']):        note_bits.append(f"进展：{v(r['项目进展'])}")
        if v(r['承包模式']):        note_bits.append(f"承包模式：{v(r['承包模式'])}")
        if v(r['招标方式']):        note_bits.append(f"招标方式：{v(r['招标方式'])}")
        if v(r['管理单位']):        note_bits.append(f"管理单位：{v(r['管理单位'])}")
        if v(r['已对接单位']):      note_bits.append(f"已对接：{v(r['已对接单位'])}")
        if v(r['联合体伙伴']):      note_bits.append(f"联合体：{v(r['联合体伙伴'])}")
        if v(r['推荐招标代理']):    note_bits.append(f"推荐代理：{v(r['推荐招标代理'])}")
        if v(r['策略说明']):        note_bits.append(f"策略：{v(r['策略说明'])}")
        note = '；'.join(note_bits)

        if old:
            sets, vals = [], []
            for dbf, srcf in FIELD_MAP.items():
                cur = v(old.get(dbf))
                nv = v(r[srcf])
                if nv and (not cur or cur == '-'):
                    sets.append(f'"{dbf}"=?'); vals.append(nv)
            # 总投资：仅在库内为空时补
            inv = v(r['合同额(万元)'])
            if inv and not v(old.get('总投资_万元')):
                sets.append('"总投资_万元"=?'); vals.append(inv)
                try:
                    sets.append('"总投资数值"=?'); vals.append(float(inv))
                except ValueError:
                    pass
            # 计划开工
            if v(r['启动时间']) and not v(old.get('计划开工时间')):
                sets.append('"计划开工时间"=?'); vals.append(v(r['启动时间']))
            # 建设内容
            if v(r['项目简介']) and not v(old.get('建设内容及规模')):
                sets.append('"建设内容及规模"=?'); vals.append(v(r['项目简介']))
            # 备注追加
            old_note = v(old.get('备注'))
            if old_note in ('-', '—'):
                old_note = ''
            newnote = (old_note + ' | ' + note) if old_note else note
            sets.append('"备注"=?'); vals.append(newnote)
            vals.append(old['row_id'])
            con.execute(f'UPDATE projects SET {",".join(sets)} WHERE row_id=?', vals)
            upd.append((name, old['项目名称'], len(sets) - 1))
        else:
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
            rec['地州'] = region_of(v(r['项目地点']), name)
            rec['项目类型'] = '储备' if v(r['项目状态']) == '储备' else '新建'
            prog = v(r['项目进展'])
            rec['分类号'] = '初步阶段' if re.search(r'初设|招标', prog) else '可研/设计阶段'
            rec['建设内容及规模'] = v(r['项目简介'])
            rec['总投资_万元'] = v(r['合同额(万元)'])
            rec['建设地点'] = v(r['项目地点'])
            rec['资金来源'] = v(r['资金来源'])
            rec['计划开工时间'] = v(r['启动时间'])
            rec['设计单位'] = v(r['设计单位'])
            rec['项目进展情况'] = prog or '前期'
            rec['招标代理'] = v(r['推荐招标代理'])
            rec['投标策略要点'] = v(r['投标策略要点'])
            rec['关键关系资源'] = v(r['关键关系资源'])
            rec['关系资源级别'] = v(r['关系资源级别'])
            rec['关系资源备注'] = v(r['关系资源备注'])
            rec['业主决策层联系人'] = v(r['业主决策层联系人'])
            rec['业主决策层职务'] = v(r['业主决策层职务'])
            rec['业主参与层联系人'] = v(r['业主参与层联系人'])
            rec['业主参与层职务'] = v(r['业主参与层职务'])
            rec['业主操作层联系人'] = v(r['业主操作层联系人'])
            rec['业主操作层职务'] = v(r['业主操作层职务'])
            rec['公共资源中心联系人'] = v(r['公共资源中心联系人'])
            rec['主要竞争对手'] = v(r['主要竞争对手'])
            rec['评标专家'] = v(r['评标专家'])
            rec['策划责任人'] = v(r['策划责任人'])
            rec['市场经理'] = v(r['主跟踪人'])
            rec['攻关状态'] = v(r['攻关状态'])
            rec['下一步行动'] = v(r['下一步行动'])
            rec['预计招标时间'] = v(r['预计招标时间'])
            rec['中标概率(%)'] = v(r['中标概率'])
            rec['备注'] = note
            try:
                rec['总投资数值'] = float(rec['总投资_万元']) if rec['总投资_万元'] else None
            except ValueError:
                rec['总投资数值'] = None
            rec['原始序号'] = None
            rec['当前看板'] = 1
            rec['数据来源'] = '项目信息录入10.9'
            rec['数据质量'] = 'ok' if rec['招标代理'] or rec['关键关系资源'] else 'ok'
            cols = list(rec.keys())
            ph = ','.join('?' * len(cols))
            con.execute(f'INSERT INTO projects ({",".join(chr(34)+c+chr(34) for c in cols)}) VALUES ({ph})',
                        [rec[c] for c in cols])
            new.append((name, rec['地州'], rec['总投资_万元'], rec['攻关状态']))

    con.commit()
    tot = con.execute('SELECT COUNT(*) FROM projects').fetchone()[0]
    cur = con.execute('SELECT COUNT(*) FROM projects WHERE 当前看板=1').fetchone()[0]
    con.close()

    print(f'\n更新已有 {len(upd)} 条:')
    for a, b, n in upd:
        print(f'   ~ {a[:34]:36} -> {b[:34]:36} 补{n}个字段')
    print(f'\n新增 {len(new)} 条:')
    for n, g, inv, st in new:
        print(f'   + {n[:38]:40} {g:<6} {str(inv):>10}万  {st}')
    print(f'\n库内合计 {tot} 条（当前有效 {cur}）')


if __name__ == '__main__':
    main()
