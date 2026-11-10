# -*- coding: utf-8 -*-
"""
和田地区项目 施工/EPC 中标复核 —— 走和田公共资源交易平台官方接口
接口: https://ggzy.ht.gov.cn/EpointWebBuilder/rest/GgSearchAction/getInfoMationList
判 定 规则与 verify_api.py 一致
用法: python verify_hetian.py [--limit N] [--delay 0.5] [--redo]
"""
import argparse, json, os, re, sqlite3, ssl, sys, time, urllib.parse, urllib.request
sys.stdout.reconfigure(encoding='utf-8')

ctx = ssl.create_default_context(); ctx.check_hostname = False; ctx.verify_mode = ssl.CERT_NONE
API = 'https://ggzy.ht.gov.cn/EpointWebBuilder/rest/GgSearchAction/getInfoMationList'
GUID = '7eb5f7f1-9041-43ad-8e13-8fcb82ea831a'
H = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
                   '(KHTML, like Gecko) Chrome/120.0 Safari/537.36',
     'X-Requested-With': 'XMLHttpRequest',
     'Referer': 'https://ggzy.ht.gov.cn/search.html',
     'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8'}
OUT = '_ht_out.jsonl'

WIN = re.compile(r'中标|成交|结果|候选人|定标')
OPEN = re.compile(r'招标公告|招标|竞争性磋商|资格预审|比选')
SERVICE = re.compile(r'施工图|监理|检测|审计|造价|勘察|测绘|咨询|招标代理|项目管理|设计|可研|环评|水保|专题|论证|评估|质量检测')
REALWORK = re.compile(r'施工总承包|工程总承包|EPC|主体工程|施工[一二三四五六七八九十\d]*标段|施工标段|施工标', re.I)
SLICE_HINT = re.compile(r'片区|灌片|县|乡|镇|团场')
BAD_NEXT = re.compile(r'^(灌区|水库|枢纽|治理|供水|防洪|灌片|渠道|水厂|泵站|水电站|调蓄|沉沙|沉砂|引水|灌溉)')
PREFIX = ['新疆维吾尔自治区', '新疆', '和田地区']


def norm(s):
    s = re.sub(r'<[^>]+>', '', str(s or ''))
    s = re.sub(r'\s+', '', s)
    for a, b in [('（', '('), ('）', ')'), ('—', '-'), ('－', '-'), ('·', ''), ('“', ''), ('”', ''), ('"', '')]:
        s = s.replace(a, b)
    return s.strip()


def name_forms(name):
    n = norm(name)
    out = [n]
    for p in PREFIX:
        pn = norm(p)
        if n.startswith(pn) and len(n) > len(pn) + 4:
            out.append(n[len(pn):]); break
    return out


def work_kind(title):
    if SERVICE.search(title) and not REALWORK.search(title):
        return None
    if not re.search(r'施工|EPC|总承包', title, re.I):
        return None
    if WIN.search(title):
        return '已中标'
    if OPEN.search(title):
        return '招标中'
    return None


def parens(s):
    return re.findall(r'[（(]([^）)]{1,24})[）)]', s)


def search(keyword, size=30):
    data = urllib.parse.urlencode({'siteGuid': GUID, 'categoryNum': '', 'address': '',
                                   'projectNumber': '', 'keyword': keyword, 'title': '',
                                   'pageIndex': 0, 'pageSize': size, 'YZM': '', 'ImgGuid': ''}).encode()
    req = urllib.request.Request(API, data=data, headers=H, method='POST')
    with urllib.request.urlopen(req, timeout=30, context=ctx) as r:
        res = json.loads(r.read().decode('utf-8', 'replace'))
    return res.get('custom') or []


def evaluate(name, records):
    forms = name_forms(name)
    have = set(parens(norm(name)))
    hits, slices, service = [], [], []
    for r in records:
        title = re.sub(r'<[^>]+>', '', str(r.get('title') or ''))
        tn = norm(title)
        hit_form = next((f for f in forms if f in tn), None)
        if not hit_form:
            continue
        rest = tn[tn.find(hit_form) + len(hit_form):]
        if BAD_NEXT.match(rest):
            continue
        kind = work_kind(title)
        if not kind:
            if WIN.search(title) or OPEN.search(title):
                service.append({'标题': title, '日期': str(r.get('infodate'))[:10]})
            continue
        extra = [g for g in parens(tn) if g not in have]
        if any(SLICE_HINT.search(g) for g in extra):
            slices.append({'标题': title, '日期': str(r.get('infodate'))[:10]})
            continue
        hits.append({'标题': title, '判定': kind, '日期': str(r.get('infodate'))[:10],
                     'infourl': r.get('infourl', '')})
    return hits, slices, service


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--limit', type=int, default=0)
    ap.add_argument('--delay', type=float, default=0.5)
    ap.add_argument('--redo', action='store_true')
    a = ap.parse_args()

    con = sqlite3.connect('xinjiang_water_projects.db'); con.row_factory = sqlite3.Row
    projs = [dict(r) for r in con.execute(
        "SELECT row_id,项目名称,地州,总投资_万元 FROM projects "
        "WHERE 当前看板=1 AND 地州='和田地区' ORDER BY row_id")]
    con.close()

    if a.redo and os.path.exists(OUT):
        os.remove(OUT)
    done = set()
    if os.path.exists(OUT):
        for line in open(OUT, encoding='utf-8'):
            try:
                done.add(json.loads(line)['row_id'])
            except Exception:
                pass
    todo = [p for p in projs if p['row_id'] not in done]
    if a.limit:
        todo = todo[:a.limit]
    print(f'和田当前有效 {len(projs)} 条 | 已完成 {len(done)} | 待做 {len(todo)}', flush=True)

    f = open(OUT, 'a', encoding='utf-8')
    n_hit = 0
    for i, p in enumerate(todo, 1):
        name = p['项目名称']
        try:
            recs = search(name)
            if not recs:
                for alt in name_forms(name)[1:]:
                    recs = search(alt)
                    if recs:
                        break
            hits, slices, service = evaluate(name, recs)
        except Exception as e:
            print(f'[{i}/{len(todo)}] {name[:26]:28} 接口失败 {e}', flush=True)
            time.sleep(3); continue
        f.write(json.dumps({'row_id': p['row_id'], '项目名称': name, '总投资': p['总投资_万元'],
                            '公告数': len(recs), '命中': hits, '切片': slices,
                            '服务类': service[:5]}, ensure_ascii=False) + '\n')
        f.flush()
        if hits:
            n_hit += 1
            print(f'[{i}/{len(todo)}] ★ {name[:26]:28} {hits[0]["判定"]} | {hits[0]["标题"][:50]}', flush=True)
        else:
            print(f'[{i}/{len(todo)}]   {name[:26]:28} 无', flush=True)
        time.sleep(a.delay)
    f.close()
    print(f'\n完成 {len(todo)} 条，判定已落地 {n_hit} 条', flush=True)


if __name__ == '__main__':
    main()
