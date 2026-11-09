# -*- coding: utf-8 -*-
"""
全库施工/EPC 招标中标复核 —— 走新疆公共资源交易平台全文检索 API
接口: https://ggzy.xinjiang.gov.cn/inteligentsearchnew/rest/esinteligentsearch/getFullTextDataNew
该接口索引覆盖全疆各地州发布的公告。

判定：
  · 公告标题需包含项目名称（全名，或去掉"新疆/地州"前缀后的名称）
  · 且含 施工/EPC/总承包
  · 含 中标/成交/结果/候选人/定标 -> 已中标；仅含 招标 -> 招标中
  · 标题在项目名之外带 （XX片区）/（XX灌片）/（一期）/第N标段 等限定语 -> 判为切片，不处理父项目
  · 模糊的大区域总体项目（大型灌区/水资源配置等）跳过
用法: python verify_api.py [--limit N] [--delay 0.6] [--redo]
"""
import argparse, json, os, re, ssl, sqlite3, sys, time, urllib.request
sys.stdout.reconfigure(encoding='utf-8')

ctx = ssl.create_default_context(); ctx.check_hostname = False; ctx.verify_mode = ssl.CERT_NONE
API = 'https://ggzy.xinjiang.gov.cn/inteligentsearchnew/rest/esinteligentsearch/getFullTextDataNew'
H = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
                   '(KHTML, like Gecko) Chrome/120.0 Safari/537.36',
     'Content-Type': 'application/json;charset=UTF-8',
     'Referer': 'https://ggzy.xinjiang.gov.cn/search.html',
     'Origin': 'https://ggzy.xinjiang.gov.cn'}
OUT = '_api_out.jsonl'

WORK = re.compile(r'施工|EPC|总承包|工程总承包', re.I)
WIN = re.compile(r'中标|成交|结果|候选人|定标')
OPEN = re.compile(r'招标公告|招标|竞争性磋商|资格预审|比选')
SKIP_PAT = re.compile(r'大型灌区|灌区续建配套与现代化改造$|水资源配置工程|东水西调|西水东引')
SLICE_HINT = re.compile(r'片区|灌片|县|乡|镇|团场')
# 名称后面紧跟这些词说明是另一个项目（而非本项目的公告类型后缀）
BAD_NEXT = re.compile(r'^(灌区|水库|枢纽|治理|供水|防洪|灌片|渠道|水厂|泵站|水电站|调蓄|沉沙|沉砂|引水|灌溉)')
PREFIX = ['新疆维吾尔自治区', '新疆', '巴音郭楞蒙古自治州', '昌吉回族自治州', '昌吉州',
          '阿勒泰地区', '和田地区', '喀什地区', '阿克苏地区', '塔城地区', '伊犁州',
          '巴州', '博州', '克州', '乌鲁木齐市', '哈密市', '吐鲁番市', '克拉玛依市']


def norm(s):
    s = re.sub(r'\s+', '', str(s or ''))
    for a, b in [('（', '('), ('）', ')'), ('—', '-'), ('－', '-'), ('―', '-'),
                 ('·', ''), ('“', ''), ('”', ''), ('"', '')]:
        s = s.replace(a, b)
    return s


def name_forms(name):
    n = norm(name)
    out = [n]
    for p in PREFIX:
        pn = norm(p)
        if n.startswith(pn) and len(n) > len(pn) + 5:
            out.append(n[len(pn):])
            break
    return out


def parens(s):
    return re.findall(r'[（(]([^）)]{1,24})[）)]', s)


def query(wd, rn=30):
    payload = {"token": "", "pn": 0, "rn": rn, "sdt": "", "edt": "", "wd": wd, "inc_wd": "",
               "exc_wd": "", "fields": "title;content", "cnum": "001",
               "sort": "{\"webdate\":\"0\"}", "ssort": "title", "cl": 200, "terminal": "",
               "condition": [], "time": "", "highlights": "title", "statistics": None,
               "unionCondition": None, "accuracy": "", "noParticiple": "0",
               "searchRange": None, "isBusiness": "1"}
    req = urllib.request.Request(API, data=json.dumps(payload, ensure_ascii=False).encode('utf-8'),
                                 headers=H, method='POST')
    with urllib.request.urlopen(req, timeout=30, context=ctx) as r:
        raw = json.loads(r.read().decode('utf-8', 'replace'))
    res = raw.get('result')
    if res is None and 'content' in raw:
        res = (json.loads(raw['content']) or {}).get('result') or {}
    return (res or {}).get('records') or []



SERVICE = re.compile(r'施工图|监理|检测|审计|造价|勘察|测绘|咨询|招标代理|项目管理|设计|可研|环评|水保|专题|论证|评估')
REALWORK = re.compile(r'施工总承包|工程总承包|EPC|主体工程|施工[一二三四五六七八九十\d]*标段|施工标段|施工标', re.I)


def work_kind(title):
    """返回 '已中标' / '招标中' / None（服务类或无关）"""
    if SERVICE.search(title) and not REALWORK.search(title):
        return None                      # 设计/监理/检测等前期服务
    if not re.search(r'施工|EPC|总承包', title, re.I):
        return None
    if WIN.search(title):
        return '已中标'
    if OPEN.search(title):
        return '招标中'
    return None


def evaluate(name, records):
    forms = name_forms(name)
    have = set(parens(norm(name)))
    hits, slices, service = [], [], []
    for r in records:
        title = re.sub(r'<[^>]+>', '', str(r.get('title') or r.get('title2') or ''))
        tn = norm(title)
        hit_form = next((f for f in forms if f in tn), None)
        if not hit_form:
            continue
        # 防前缀误判：名称后面直接接另一个工程名词，说明是不同项目
        _rest = tn[tn.find(hit_form) + len(hit_form):]
        if BAD_NEXT.match(_rest):
            continue
        kind = work_kind(title)
        if not kind:
            if WIN.search(title) or OPEN.search(title):
                service.append({'标题': title, '日期': str(r.get('webdate'))[:10]})
            continue
        extra = [g for g in parens(norm(title)) if g not in have]
        if any(SLICE_HINT.search(g) for g in extra):
            slices.append({'标题': title, '日期': str(r.get('webdate'))[:10]})
            continue
        hits.append({'标题': title, '判定': kind, '日期': str(r.get('webdate'))[:10],
                     'infoid': r.get('infoid', '')})
    return hits, slices, service


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--limit', type=int, default=0)
    ap.add_argument('--delay', type=float, default=0.6)
    ap.add_argument('--redo', action='store_true')
    a = ap.parse_args()

    con = sqlite3.connect('xinjiang_water_projects.db'); con.row_factory = sqlite3.Row
    projs = [dict(r) for r in con.execute(
        'SELECT row_id,项目名称,地州,总投资_万元 FROM projects WHERE 当前看板=1 ORDER BY row_id')]
    con.close()

    skip = [p for p in projs if SKIP_PAT.search(p['项目名称'])]
    pool = [p for p in projs if not SKIP_PAT.search(p['项目名称'])]
    if a.redo and os.path.exists(OUT):
        os.remove(OUT)
    done = set()
    if os.path.exists(OUT):
        for line in open(OUT, encoding='utf-8'):
            try:
                done.add(json.loads(line)['row_id'])
            except Exception:
                pass
    todo = [p for p in pool if p['row_id'] not in done]
    if a.limit:
        todo = todo[:a.limit]
    print(f'跳过模糊大区域 {len(skip)} 条 | 池 {len(pool)} | 已完成 {len(done)} | 待做 {len(todo)}', flush=True)

    f = open(OUT, 'a', encoding='utf-8')
    n_hit = 0
    for i, p in enumerate(todo, 1):
        name = p['项目名称']
        try:
            recs = query(name)
            if not recs:
                for alt in name_forms(name)[1:]:
                    recs = query(alt)
                    if recs:
                        break
            hits, slices, service = evaluate(name, recs)
        except Exception as e:
            print(f'[{i}/{len(todo)}] {name[:26]:28} 接口失败 {e}', flush=True)
            time.sleep(5); continue
        f.write(json.dumps({'row_id': p['row_id'], '项目名称': name, '地州': p['地州'],
                            '总投资': p['总投资_万元'], '公告数': len(recs),
                            '命中': hits, '切片': slices, '服务类': service[:5]},
                           ensure_ascii=False) + '\n')
        f.flush()
        if hits:
            n_hit += 1
            print(f'[{i}/{len(todo)}] ★ {name[:26]:28} {hits[0]["判定"]} | {hits[0]["标题"][:52]}', flush=True)
        elif slices:
            print(f'[{i}/{len(todo)}]   {name[:26]:28} 仅切片公告', flush=True)
        else:
            print(f'[{i}/{len(todo)}]   {name[:26]:28} 无', flush=True)
        time.sleep(a.delay)
    f.close()
    print(f'\n完成 {len(todo)} 条，判定已落地 {n_hit} 条', flush=True)


if __name__ == '__main__':
    main()
