# -*- coding: utf-8 -*-
"""
수집된 병원 뉴스(news.csv)에서 서울아산병원 기사만 추려 대시보드용 JSON 을 만듭니다.

수집본 저장소(hospital-news-dashboard)는 운영본이라 건드리지 않습니다.
여기서 읽기만 하고, 결과는 이 저장소의 data/articles.json 으로 커밋합니다.

    python scripts/build_news.py
    python scripts/build_news.py --src ../hospital-news-dashboard/docs/news.csv

새 기사가 쌓이면 다시 돌려서 커밋하세요. 대시보드는 이 파일만 읽습니다.
"""
import argparse
import csv
import io
import json
import os
import re
import sys
from collections import Counter

DEFAULT_SRC = os.path.join('..', 'hospital-news-dashboard', 'docs', 'news.csv')
OUT = os.path.join('data', 'articles.json')
GROUP = '아산'


def norm_title(t):
    """광고 '내용'과 기사 제목을 같은 모양으로 만들어 맞춰보기 위한 키."""
    t = re.sub(r'^\[[^\]]*\]\s*', '', str(t or ''))          # [지면] [잡지] [온라인]
    t = re.sub(r'[‘’“”\'"`]', '', t)      # 따옴표 종류 통일
    # 브라우저 쪽 normTitle() 과 규칙이 같아야 합니다.
    # JS 의 \w 는 ASCII 만 담아서, 한자까지 남기면 양쪽 키가 어긋납니다(年, 韓 등).
    t = re.sub(r'[^0-9A-Za-z가-힣]', '', t)
    return t.lower()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--src', default=DEFAULT_SRC)
    args = ap.parse_args()

    if not os.path.exists(args.src):
        sys.exit('수집본을 찾을 수 없습니다: %s\n--src 로 news.csv 경로를 지정하세요.' % args.src)

    with io.open(args.src, encoding='utf-8-sig', newline='') as f:
        rows = [r for r in csv.DictReader(f) if r.get('병원그룹') == GROUP]

    seen = set()
    out = []
    for r in rows:
        url = (r.get('언론사원문') or r.get('네이버링크') or '').strip()
        title = (r.get('제목') or '').strip()
        if not url or not title:
            continue
        key = norm_title(title)
        if key in seen:        # 같은 기사가 여러 검색어로 중복 수집됩니다
            continue
        seen.add(key)
        out.append({
            # 매칭 키(k)는 브라우저에서 같은 규칙으로 만듭니다. 파일에 넣으면 제목이 두 번 들어가 커집니다.
            't': title,
            'u': url,
            'd': (r.get('날짜') or '').strip(),
            'm': (r.get('매체') or '').strip(),
            'r': (r.get('기자명') or '').strip(),
            'p': (r.get('교수명') or '').strip(),
        })

    out.sort(key=lambda x: x['d'], reverse=True)

    if not os.path.isdir('data'):
        os.makedirs('data')
    with io.open(OUT, 'w', encoding='utf-8') as f:
        json.dump({'group': GROUP, 'count': len(out), 'articles': out}, f, ensure_ascii=False)

    size = os.path.getsize(OUT) / 1024.0 / 1024.0
    years = Counter(a['d'][:4] for a in out if a['d'])
    print('%s 기사 %d건 (중복 제거 전 %d건) → %s  %.2fMB'
          % (GROUP, len(out), len(rows), OUT, size))
    print('연도별:', dict(sorted(years.items())))
    print('매체 상위:', Counter(a['m'] for a in out).most_common(8))


if __name__ == '__main__':
    main()
