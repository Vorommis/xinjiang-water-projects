# -*- coding: utf-8 -*-
"""
从 SQLite 导出 JSON / CSV（与 rebuild_db.py 使用同一套字段顺序）。
用于「在库内整合了外部数据之后」重新生成导出文件——
不要直接重跑 rebuild_db.py，那会按原始三个来源重建库、覆盖整合结果。
"""
import argparse, csv, json, os, sqlite3, sys
sys.stdout.reconfigure(encoding='utf-8')

from rebuild_db import OUT_COLS, SCHEMA, PROV   # 复用字段定义


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--db', default='xinjiang_water_projects.db')
    ap.add_argument('--outdir', default='.')
    a = ap.parse_args()

    con = sqlite3.connect(a.db)
    con.row_factory = sqlite3.Row
    rows = [{k: r[k] for k in r.keys()} for r in con.execute('SELECT * FROM projects ORDER BY row_id')]
    con.close()

    js = os.path.join(a.outdir, 'xinjiang_water_projects.json')
    cv = os.path.join(a.outdir, 'xinjiang_water_projects.csv')

    with open(js, 'w', encoding='utf-8') as f:
        json.dump([{c: r.get(c) for c in OUT_COLS} for r in rows], f, ensure_ascii=False, indent=1)
    with open(cv, 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.DictWriter(f, fieldnames=OUT_COLS)
        w.writeheader()
        for r in rows:
            w.writerow({c: r.get(c) for c in OUT_COLS})

    cur = sum(1 for r in rows if r.get('当前看板') == 1)
    print(f'导出 {len(rows)} 条（当前有效 {cur}）')
    for p in (js, cv):
        print(f'  {p}  ({os.path.getsize(p):,} bytes)')


if __name__ == '__main__':
    main()
