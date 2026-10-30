# -*- coding: utf-8 -*-
"""
新疆水利项目信息库 —— 逆向重建脚本
================================================
从 GitHub 仓库 Vorommis/xinjiang-water-projects 中的加密看板 HTML，
反向重建出原始数据库（SQLite + JSON + CSV）。

数据来源（两级）：
  A. 最新看板 index.html（AES-256-GCM 加密，密码在内）：
     - 224 行 × 36 个展示字段  = 当前数据库状态（含最新修正）
  B. Git 历史快照 commit 2d0d510 中内嵌的 JSON：
     - 343 行 × 44 个字段 = 更早期的完整数据库导出（含 id / created_at /
       政府投资 / 其他投资 / 2026 投资 / 计划竣工时间 / 政府实施单位）
     - 且包含后来被 "总投资≥5000万" 规则筛掉的 137 条记录

合并规则：
  - 看板为准（它经过多次人工修正，如"伊犁州34行投资数据修复"）
  - 快照用于补全看板未展示的 8 个字段
  - 已知缺陷：伊犁州 34 行在快照中列错位（总投资 与 项目进展情况 互换，
    且"政府投资"列实际存放的是总投资），脚本按实测规律修复并标注

用法: python rebuild_db.py [--password sl888888] [--outdir .]
"""
import argparse, base64, hashlib, json, os, re, sqlite3, subprocess, sys
from collections import Counter, OrderedDict

sys.stdout.reconfigure(encoding='utf-8')

# ---------------------------------------------------------------- 配置
DEFAULT_PASSWORD = 'sl888888'
DEFAULT_ENDPOINT = 'index.html'
DUMP_COMMIT = '2d0d510'          # 最后一个内嵌 44 字段 JSON 快照的提交
DUMP_COMMIT_DESC = '2d0d510 (2026-07-21 37列全字段看板)'

# 原始数据库的 44 个字段（顺序与快照一致）
SCHEMA = [
    'id', '地州', '项目类型', '分类号', '项目名称', '建设内容及规模',
    '总投资_万元', '政府投资_万元', '其他投资_万元',
    '年2026预计完成投资_万元', '年2026政府投资_万元', '年2026其他投资_万元',
    '计划开工时间', '计划竣工时间', '建设地点', '项目法人', '项目单位性质',
    '联系人及联系方式', 'created_at', '资金来源', '设计单位', '招标代理',
    '项目进展情况', '关键关系资源', '关系资源级别', '关系资源备注',
    '业主决策层联系人', '业主决策层职务', '业主参与层联系人', '业主参与层职务',
    '业主操作层联系人', '业主操作层职务', '政府实施单位', '公共资源中心联系人',
    '主要竞争对手', '评标专家', '投标策略要点', '策划责任人', '市场经理',
    '攻关状态', '下一步行动', '预计招标时间', '中标概率(%)', '备注',
]
# 重建时附加的溯源列（原库没有）
PROV = ['原始序号', '当前看板', '数据来源', '数据质量', '总投资数值', '重点跟踪']
OUT_COLS = SCHEMA + PROV

# 看板展示列名 -> 数据库字段名
DASH_MAP = {
    '地州': '地州', '项目类型': '项目类型', '分类号': '分类号',
    '项目名称': '项目名称', '建设内容及规模': '建设内容及规模',
    '总投资': '总投资_万元', '资金来源': '资金来源',
    '计划开工时间': '计划开工时间', '建设地点': '建设地点',
    '项目法人': '项目法人', '项目单位性质': '项目单位性质',
    '联系人及联系方式': '联系人及联系方式', '设计单位': '设计单位',
    '招标代理': '招标代理', '项目进展情况': '项目进展情况',
    '关键关系资源': '关键关系资源', '关系资源级别': '关系资源级别',
    '关系资源备注': '关系资源备注',
    '业主决策层联系人': '业主决策层联系人', '业主决策层职务': '业主决策层职务',
    '业主参与层联系人': '业主参与层联系人', '业主参与层职务': '业主参与层职务',
    '业主操作层联系人': '业主操作层联系人', '业主操作层职务': '业主操作层职务',
    '公共资源中心联系人': '公共资源中心联系人',
    '主要竞争对手': '主要竞争对手', '评标专家': '评标专家',
    '投标策略要点': '投标策略要点', '策划责任人': '策划责任人',
    '市场经理': '市场经理', '攻关状态': '攻关状态',
    '下一步行动': '下一步行动', '预计招标时间': '预计招标时间',
    '中标概率': '中标概率(%)', '备注': '备注',
}
# 快照独有的、可用于补全的字段
EXTRA_FIELDS = ['政府投资_万元', '其他投资_万元',
                '年2026预计完成投资_万元', '年2026政府投资_万元',
                '年2026其他投资_万元', '计划竣工时间', '政府实施单位']
# 投资细分字段（参考 Excel 也可提供）
INV_FIELDS = ['政府投资_万元', '其他投资_万元',
              '年2026预计完成投资_万元', '年2026政府投资_万元',
              '年2026其他投资_万元', '计划竣工时间']
# 参考 Excel 列名 -> 数据库字段名
XLSX_MAP = {
    '政府投资(万元)': '政府投资_万元',
    '其他投资(万元)': '其他投资_万元',
    '2026预计完成投资(万元)': '年2026预计完成投资_万元',
    '2026政府投资(万元)': '年2026政府投资_万元',
    '2026其他投资(万元)': '年2026其他投资_万元',
    '计划竣工时间': '计划竣工时间',
}
# 快照与看板共有的字段（用于仅快照记录）
COMMON_FIELDS = ['id', '地州', '项目类型', '分类号', '项目名称', '建设内容及规模',
                 '计划开工时间', '建设地点', '项目法人', '项目单位性质',
                 '联系人及联系方式', 'created_at', '资金来源', '设计单位',
                 '招标代理', '备注']


def norm(s):
    return re.sub(r'\s+', '', str(s or '')).strip()


def is_num(s):
    return bool(re.fullmatch(r'-?\d+(\.\d+)?', str(s).strip()))


# ---------------------------------------------------------------- 解密
def decrypt_index(path, password):
    """解开 index.html 里的 AES-256-GCM 数据块，返回明文 HTML。"""
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    src = open(path, encoding='utf-8').read()
    m = re.search(r'var ED="([A-Za-z0-9+/=]+)"', src)
    if not m:
        return src                      # 未加密（老版本），原样返回
    raw = base64.b64decode(m.group(1))
    key = hashlib.sha256(password.encode('utf-8')).digest()
    return AESGCM(key).decrypt(raw[:12], raw[12:-16] + raw[-16:], None).decode('utf-8')


def parse_dashboard(html):
    """从看板 HTML 解析出全部数据行。"""
    from bs4 import BeautifulSoup
    soup = BeautifulSoup(html, 'lxml')
    ths = [t.get_text(strip=True) for t in soup.select('thead th')]
    out = []
    for tr in soup.select('#tb tr'):
        tds = tr.find_all('td')
        if not tds:
            continue
        rec = {'_region': tr.get('data-region'),
               '_cat': tr.get('data-cat'),
               '_ptype': tr.get('data-ptype'),
               '原始序号': tr.find_all('td')[0].get_text(strip=True)}
        for i, td in enumerate(tds):
            name = ths[i] if i < len(ths) else f'col{i}'
            rec[name] = (td.get('title') if td.has_attr('title')
                         else td.get_text(strip=True)) or ''
        out.append(rec)
    return out, ths


def load_snapshot(commit=DUMP_COMMIT):
    """从 git 历史里取出内嵌 JSON 的数据库快照。"""
    try:
        blob = subprocess.run(['git', 'show', f'{commit}:index.html'],
                              capture_output=True, check=True).stdout
    except Exception as e:
        print(f'!! 无法从 git 读取 {commit}: {e}')
        return []
    src = blob.decode('utf-8', errors='replace')
    m = re.search(r'var ED="([A-Za-z0-9+/=]+)"', src)
    if m:                                   # 该快照是加密的
        from cryptography.hazmat.primitives.ciphers.aead import AESGCM
        raw = base64.b64decode(m.group(1))
        key = hashlib.sha256(DEFAULT_PASSWORD.encode()).digest()
        src = AESGCM(key).decrypt(raw[:12], raw[12:-16] + raw[-16:], None).decode('utf-8')
    m = re.search(r'<script[^>]*type="application/json"[^>]*>(.*?)</script>', src, re.S) \
        or re.search(r'var DATA\s*=\s*(\[.*?\]);', src, re.S)
    if not m:
        print(f'!! {commit} 中未找到内嵌 JSON')
        return []
    return json.loads(m.group(1))


# ---------------------------------------------------------------- 参考 Excel
DATE_RE = re.compile(r'^(19|20)\d{2}')
KEEP_AS_IS = {'待定', '-'}


def load_reference_excel(path):
    """
    读取用户提供的旧库导出 Excel（多工作表，取 '水利项目总表'）。
    返回 {归一化项目名称: {字段: 值}}，并清洗明显错位的数据。
    """
    try:
        import openpyxl
    except ImportError:
        print('!! 未安装 openpyxl，跳过参考 Excel')
        return {}, []
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    sheet = '水利项目总表' if '水利项目总表' in wb.sheetnames else wb.sheetnames[0]
    ws = wb[sheet]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if not rows:
        return {}, []
    head = rows[0]
    hdr = [(str(head[i]).strip() if head[i] is not None else '') for i in range(len(head))]

    def cell(v):
        if v is None:
            return ''
        if isinstance(v, float) and v.is_integer():
            v = int(v)
        return str(v).strip()

    out, anomalies = {}, []
    for r in rows[1:]:
        rec = {hdr[i]: cell(v) for i, v in enumerate(r) if i < len(hdr) and hdr[i]}
        name = rec.get('项目名称', '')
        if not name:
            continue
        d = {}
        for xc, mc in XLSX_MAP.items():
            v = rec.get(xc, '')
            if v in ('', '-', 'None'):
                d[mc] = ''
                continue
            if mc == '计划竣工时间':
                # 该列存在串列污染（把"建设地点"写进来了）
                if not DATE_RE.match(v) and v not in KEEP_AS_IS:
                    anomalies.append((name, mc, v, '非日期格式，疑似串列，已置空'))
                    d[mc] = ''
                    continue
            elif not is_num(v):
                anomalies.append((name, mc, v, '非数值，已置空'))
                d[mc] = ''
                continue
            d[mc] = v
        out[norm(name)] = d
    return out, anomalies



# ---------------------------------------------------------------- 合并
def build(cur_rows, snap, ref=None):
    ref = ref or {}
    snap_by = {}
    for r in snap:
        snap_by.setdefault(norm(r.get('项目名称')), r)

    records, stats = [], Counter()
    used = set()

    # ---- 1) 当前看板的 224 条 ----
    for r in cur_rows:
        rec = {k: '' for k in SCHEMA}
        for dash_name, db_name in DASH_MAP.items():
            rec[db_name] = r.get(dash_name, '') or ''
        # 分类号：看板展示的是缩写，data-cat 才是完整值
        if r.get('_cat'):
            rec['分类号'] = r['_cat']
        rec['原始序号'] = r.get('原始序号', '')
        rec['当前看板'] = 1

        key = norm(rec['项目名称'])
        d = snap_by.get(key)
        rf = ref.get(key)
        if d:
            used.add(key)
            rec['id'] = d.get('id', '')
            rec['created_at'] = d.get('created_at', '')
        if rec['地州'] == '伊犁州':
            # 伊犁州 34 行的投资细分字段存在"政府投资 = 总投资"的重复填写，
            # 且快照中"总投资"列被项目进展情况文本污染。此处保留原值并标注存疑。
            for f in INV_FIELDS:
                v = (rf or {}).get(f, '')
                if not v and d:
                    v = d.get(f, '') or ''
                rec[f] = v
            if d:
                rec['政府实施单位'] = d.get('政府实施单位', '')
            if rf or d:
                rec['数据质量'] = '伊犁州：投资细分疑似重复填写(政府投资=总投资)，请复核'
            else:
                rec['数据质量'] = 'ok(无早期来源，投资细分缺失)'
            src = ['看板']
            if rf:
                src.append('Excel')
            if d:
                src.append('DB快照')
            rec['数据来源'] = '+'.join(src)
            stats['伊犁_需复核'] += 1
        else:
            # 投资细分：Excel 优先，其次 git 快照
            for f in INV_FIELDS:
                v = (rf or {}).get(f, '')
                if not v and d:
                    v = d.get(f, '') or ''
                rec[f] = v
            if d:
                rec['政府实施单位'] = d.get('政府实施单位', '')
            src = ['看板']
            if rf:
                src.append('Excel')
            if d:
                src.append('DB快照')
            rec['数据来源'] = '+'.join(src)
            rec['数据质量'] = 'ok'
            stats['已补全投资细分'] += 1
        if not d and not rf:
            rec['数据来源'] = '仅看板(快照之后新增)'
            rec['数据质量'] = 'ok(无早期来源，投资细分缺失)'
            stats['仅在看板'] += 1
        records.append(rec)

    # ---- 2) 快照里有、看板没有的（被 ≥5000万 筛掉的）----
    for d in snap:
        key = norm(d.get('项目名称'))
        if key in used:
            continue
        rec = {k: '' for k in SCHEMA}
        for f in COMMON_FIELDS:
            rec[f] = d.get(f, '')
        rec['当前看板'] = 0
        rec['数据来源'] = '仅DB快照(被≥5000万规则筛除)'
        if d.get('地州') == '伊犁州':
            # 错位修复：总投资 列里是项目进展情况文本，政府投资 列才是真总投资
            rec['总投资_万元'] = d.get('政府投资_万元', '')
            rec['项目进展情况'] = d.get('总投资_万元', '')
            # 投资细分同样存在重复填写问题，保留原值并标注
            for f in INV_FIELDS:
                if f != '计划竣工时间':
                    rec[f] = d.get(f, '')
            rec['计划竣工时间'] = d.get('计划竣工时间', '')
            rec['数据质量'] = '伊犁州列错位(总投资已修复；投资细分疑似重复填写)'
            stats['筛除_伊犁_已修复'] += 1
        else:
            rec['总投资_万元'] = d.get('总投资_万元', '')
            rec['项目进展情况'] = d.get('项目进展情况', '')
            for f in EXTRA_FIELDS:
                rec[f] = d.get(f, '')
            rec['数据质量'] = 'ok'
            stats['筛除_正常'] += 1
        records.append(rec)

    # ---- 3) 排序：看板记录按原始序号在前，筛除记录按 id 在后 ----
    def sortkey(x):
        if x['当前看板']:
            try:
                return (0, int(x['原始序号'] or 0))
            except ValueError:
                return (0, 9999)
        try:
            return (1, int(x['id'] or 0))
        except (ValueError, TypeError):
            return (1, 9999)
    records.sort(key=sortkey)

    # ---- 4) 数值列 ----
    for i, rec in enumerate(records, 1):
        rec['总投资数值'] = float(rec['总投资_万元']) if is_num(rec['总投资_万元']) else None
        rec['_row'] = i
    return records, stats


# ---------------------------------------------------------------- 输出
def write_sqlite(records, path):
    if os.path.exists(path):
        os.remove(path)
    con = sqlite3.connect(path)
    cols = ['row_id'] + OUT_COLS
    INT_COLS = {'row_id', '当前看板', '原始序号', 'id'}
    defs = ', '.join(
        f'"{c}" ' + ('INTEGER PRIMARY KEY AUTOINCREMENT' if c == 'row_id'
                     else 'REAL' if c == '总投资数值'
                     else 'INTEGER' if c in INT_COLS else 'TEXT')
        for c in cols)
    con.execute(f'CREATE TABLE projects ({defs})')

    def conv(c, v):
        if c in INT_COLS or c == '总投资数值':
            if v is None or str(v).strip() in ('', '-'):
                return None
            try:
                return float(v) if c == '总投资数值' else int(float(v))
            except (ValueError, TypeError):
                return None
        return v

    ph = ', '.join('?' * len(cols))
    for rec in records:
        con.execute(f'INSERT INTO projects VALUES ({ph})',
                    [conv(c, rec['_row'] if c == 'row_id' else rec.get(c)) for c in cols])

    con.execute('CREATE INDEX idx_name ON projects("项目名称")')
    con.execute('CREATE INDEX idx_region ON projects("地州")')
    con.execute('CREATE INDEX idx_inv ON projects("总投资数值")')
    con.execute('CREATE INDEX idx_current ON projects("当前看板")')

    # 便捷视图
    con.execute('CREATE VIEW v_current AS SELECT * FROM projects WHERE 当前看板=1')
    con.execute('CREATE VIEW v_filtered_out AS SELECT * FROM projects WHERE 当前看板=0')
    con.commit()
    con.close()


def write_json(records, path):
    out = [{c: r.get(c) for c in OUT_COLS} for r in records]
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(out, f, ensure_ascii=False, indent=1)


def write_csv(records, path):
    import csv
    with open(path, 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.DictWriter(f, fieldnames=OUT_COLS)
        w.writeheader()
        for r in records:
            w.writerow({c: r.get(c) for c in OUT_COLS})


# ---------------------------------------------------------------- 主流程
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--password', default=DEFAULT_PASSWORD)
    ap.add_argument('--index', default=DEFAULT_ENDPOINT)
    ap.add_argument('--outdir', default='.')
    ap.add_argument('--dump-commit', default=DUMP_COMMIT)
    ap.add_argument('--ref-excel', default='新疆水利项目统一信息库.xlsx',
                    help='参考用的旧库导出 Excel（不存在则忽略）')
    ap.add_argument('--no-ref', action='store_true', help='忽略参考 Excel')
    a = ap.parse_args()

    print('=' * 66)
    print('新疆水利项目信息库 —— 数据库逆向重建')
    print('=' * 66)

    print(f'\n[1/5] 解密并解析最新看板 {a.index} ...')
    html = decrypt_index(a.index, a.password)
    cur_rows, ths = parse_dashboard(html)
    print(f'      看板列数 {len(ths)}，数据行 {len(cur_rows)}')

    print(f'\n[2/5] 从 git 历史读取数据库快照 {a.dump_commit} ...')
    snap = load_snapshot(a.dump_commit)
    print(f'      快照记录 {len(snap)} 条，字段 {len(snap[0]) if snap else 0} 个')

    ref = {}
    print(f'\n[3/5] 读取参考 Excel {a.ref_excel} ...')
    if a.no_ref or not os.path.exists(a.ref_excel):
        print('      跳过（未提供）')
    else:
        ref, anomalies = load_reference_excel(a.ref_excel)
        print(f'      参考记录 {len(ref)} 条')
        if anomalies:
            print(f'      清洗掉 {len(anomalies)} 个异常单元格：')
            for name, f, v, why in anomalies[:10]:
                print(f'        - {name[:26]:28} {f:14} = {v!r}  ({why})')
            if len(anomalies) > 10:
                print(f'        ... 另有 {len(anomalies)-10} 个')

    print('\n[4/5] 合并 ...')
    records, stats = build(cur_rows, snap, ref)
    for k, v in stats.items():
        print(f'      {k}: {v}')
    print(f'      重建总记录数: {len(records)}')

    print(f'\n[5/5] 写出到 {a.outdir} ...')
    db = os.path.join(a.outdir, 'xinjiang_water_projects.db')
    js = os.path.join(a.outdir, 'xinjiang_water_projects.json')
    cv = os.path.join(a.outdir, 'xinjiang_water_projects.csv')
    write_sqlite(records, db)
    write_json(records, js)
    write_csv(records, cv)
    for p in (db, js, cv):
        print(f'      {p}  ({os.path.getsize(p):,} bytes)')

    # 汇总
    cur_only = [r for r in records if r['当前看板']]
    print('\n--- 汇总 ---')
    print(f'  当前看板记录        : {len(cur_only)}')
    print(f'  被筛选掉的历史记录  : {len(records) - len(cur_only)}')
    print(f'  地州分布            : {dict(Counter(r["地州"] for r in records).most_common())}')
    inv = [r['总投资数值'] for r in records if r['总投资数值']]
    print(f'  总投资已知          : {len(inv)} 条, 合计 {sum(inv):,.0f} 万元')
    return records


if __name__ == '__main__':
    main()
