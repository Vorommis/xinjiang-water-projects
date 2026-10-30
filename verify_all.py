# -*- coding: utf-8 -*-
"""施工/EPC 招标中标核查（全库，搜狗检索）"""
import argparse, json, os, re, ssl, sqlite3, sys, time, urllib.parse, urllib.request
sys.stdout.reconfigure(encoding="utf-8")

ctx = ssl.create_default_context(); ctx.check_hostname = False; ctx.verify_mode = ssl.CERT_NONE
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36",
      "Accept": "text/html,application/xhtml+xml", "Accept-Language": "zh-CN,zh;q=0.9"}
OUT = "_verify_out.jsonl"
WORK = re.compile(r"施工|EPC|总承包", re.I)
BID = re.compile(r"招标|中标|候选人|结果|成交|定标|公示")


def normtxt(s):
    s = re.sub(r"\s+", "", str(s or ""))
    for a, b in [("（", "("), ("）", ")"), ("—", "-"), ("－", "-"), ("―", "-"), ("·", ""), ("“", ""), ("”", "")]:
        s = s.replace(a, b)
    return s


def sogou(q, tries=3):
    url = "https://www.sogou.com/web?query=" + urllib.parse.quote(q)
    last = None
    for k in range(tries):
        try:
            req = urllib.request.Request(url, headers=UA)
            with urllib.request.urlopen(req, timeout=30, context=ctx) as r:
                html = r.read().decode("utf-8", "replace")
            if "验证码" in html or "antispider" in html:
                last = "captcha"; time.sleep(15 * (k + 1)); continue
            out = []
            for m in re.finditer(r"<h3[^>]*>(.*?)</h3>", html, re.S):
                a = re.search(r'<a[^>]*href="([^"]*)"[^>]*>(.*?)</a>', m.group(1), re.S)
                if not a: continue
                title = re.sub(r"<[^>]+>", "", a.group(2)).strip()
                if title: out.append({"title": title, "url": a.group(1)})
            return out
        except Exception as e:
            last = type(e).__name__ + ": " + str(e); time.sleep(6 * (k + 1))
    raise RuntimeError(last)


def judge(name, items):
    n = normtxt(name)
    strict, near = [], []
    core = re.sub(r"(新疆|维吾尔自治区|阿克苏地区|和田地区|喀什地区|阿勒泰地区|昌吉州|巴州|塔城地区|乌鲁木齐市|地区|工程|项目|建设|续建|配套|与|及)", "", n)
    core = core[:10]
    for it in items:
        t = it["title"]; tn = normtxt(t)
        if not WORK.search(t) or not BID.search(t): continue
        if n and n in tn:
            strict.append({"标题": t, "URL": it["url"],
                           "判定": "已中标/定标" if re.search(r"中标|成交|结果|候选人|定标", t) else "招标中"})
        elif len(core) >= 5 and core in tn:
            near.append({"标题": t, "URL": it["url"]})
    return strict, near


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--delay", type=float, default=3.0)
    a = ap.parse_args()
    con = sqlite3.connect("xinjiang_water_projects.db"); con.row_factory = sqlite3.Row
    projs = [dict(r) for r in con.execute("SELECT row_id,项目名称,地州,总投资_万元 FROM projects WHERE 当前看板=1 ORDER BY row_id")]
    con.close()
    done = set()
    if os.path.exists(OUT):
        for line in open(OUT, encoding="utf-8"):
            try: done.add(json.loads(line)["row_id"])
            except Exception: pass
    todo = [p for p in projs if p["row_id"] not in done]
    if a.limit: todo = todo[:a.limit]
    print("待核查 %d 条（已完成 %d）" % (len(todo), len(done)), flush=True)
    f = open(OUT, "a", encoding="utf-8")
    ns = nn = 0
    for i, p in enumerate(todo, 1):
        q = '"' + p["项目名称"] + '" 施工 中标'
        try:
            items = sogou(q)
        except Exception as e:
            print("[%d/%d]  %-32s 检索失败 %s" % (i, len(todo), p["项目名称"][:30], e), flush=True)
            time.sleep(10); continue
        strict, near = judge(p["项目名称"], items)
        f.write(json.dumps({"row_id": p["row_id"], "项目名称": p["项目名称"], "地州": p["地州"],
                            "总投资": p["总投资_万元"], "结果数": len(items),
                            "严格命中": strict, "近似命中": near}, ensure_ascii=False) + "\n")
        f.flush()
        if strict:
            ns += 1
            print("[%d/%d] ★ %-30s" % (i, len(todo), p["项目名称"][:30]), flush=True)
            for h in strict[:2]:
                print("        [%s] %s" % (h["判定"], h["标题"][:74]), flush=True)
        elif near:
            nn += 1
            print("[%d/%d] ? %-30s %s" % (i, len(todo), p["项目名称"][:30], near[0]["标题"][:50]), flush=True)
        else:
            print("[%d/%d]   %-30s 无" % (i, len(todo), p["项目名称"][:30]), flush=True)
        time.sleep(a.delay)
    f.close()
    print("\n本轮 %d 条，严格命中 %d，近似 %d" % (len(todo), ns, nn))


if __name__ == "__main__":
    main()