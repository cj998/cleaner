# 정리_상담원장.py — 상담원장_2025.xlsx 정리
#
# 따르는 결정 기록: my-wiki/decisions/d-005-상담원장-정리-다섯-가지.md
#   d-005 #1  처리상태 «종결» → «완료», «대기» → «보류» 로 합친다
#   d-005 #2  처리시간 음수는 지우지 않는다. 값은 그대로, `_표시` 칸에 이유
#   d-005 #3  만족도 6~10점은 지우지 않는다. `_표시` 칸에 이유. 평균 만족도에서는 표시한 줄을 뺀다
#   d-005 #4  칸 이름은 원본 그대로 둔다
#   d-005 #5  고객번호 앞뒤 공백을 없앤다
#
# 나머지 처리(테스트 접수 제외, 표기 통일, 상담시간 초→분, 중복·상태 변경 전 줄 제거)는
# 이전 정리 스크립트(clean_counsel.js)에 있던 것을 그대로 옮겼다. 새로 더한 처리는 없다.
#
# 원본은 읽기만 한다. 쓰는 파일:
#   정리본/상담원장_정리.csv · 정리본/지운줄.csv · 정리본/변경요약.md
#
# 돌리기:  python 정리/정리_상담원장.py

import math
import re
import sys
from collections import Counter
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent          # cleaner/
SRC = ROOT.parent / "messy" / "상담원장_2025.xlsx"
OUT_DIR = ROOT / "정리본"
OUT_CLEAN = OUT_DIR / "상담원장_정리.csv"
OUT_DROP = OUT_DIR / "지운줄.csv"
OUT_SUMMARY = OUT_DIR / "변경요약.md"

CH = {"유선": "전화", "전화상담": "전화", "tel": "전화", "전화": "전화",
      "chatbot": "챗봇", "챗봇": "챗봇",
      "app": "앱", "모바일앱": "앱", "앱": "앱",
      "e-mail": "이메일", "메일": "이메일", "이메일": "이메일"}
REG = {"서울시": "서울", "서울특별시": "서울", "경기도": "경기"}
YN = {"Y": "Y", "y": "Y", "예": "Y", "TRUE": "Y",
      "N": "N", "n": "N", "아니오": "N", "FALSE": "N"}
STATUS = {"종결": "완료", "대기": "보류"}                 # d-005 #1
SAT_FLAG = "만족도 1~5점 척도 밖"                         # d-005 #3


def is_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def js_str(v):
    """숫자를 이전 스크립트(JS String())와 같은 모양으로 적는다: 12.0 → 12"""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    if isinstance(v, float):
        return repr(v)
    return str(v)


def esc(v):
    s = js_str(v)
    return '"' + s.replace('"', '""') + '"' if re.search(r'[",\n]', s) else s


def write_csv(path, header, rows):
    text = "﻿" + "\r\n".join([",".join(header)] + [",".join(r) for r in rows]) + "\r\n"
    path.write_bytes(text.encode("utf-8"))


def read_source():
    wb = openpyxl.load_workbook(SRC, read_only=True, data_only=True)
    try:
        it = wb["원장"].iter_rows(values_only=True)
        cols = [str(c) for c in next(it)]
        raw = []
        for r in it:
            if all(v is None for v in r):
                continue
            raw.append({c: ("" if v is None else v) for c, v in zip(cols, r)})
    finally:
        wb.close()
    return cols, raw


def clean_row(r0, cols):
    """한 줄을 정리한다. (정리한 줄, 바꾼 것 목록)"""
    tags = []
    r = {}
    for c in cols:                                      # 앞뒤 공백 (d-005 #5 의 고객번호 포함)
        v = r0[c]
        if isinstance(v, str) and v.strip() != v:
            tags.append(f"앞뒤 공백 제거: {c}")
            v = v.strip()
        r[c] = v

    d = str(r["접수일"]).replace("/", "-").replace(".", "-")
    if d != r["접수일"]:
        tags.append("접수일 표기 통일"); r["접수일"] = d

    ch = CH.get(re.sub(r"\s", "", str(r["접수채널"]).lower()))
    if not ch:
        raise ValueError(f"채널 {r['접수채널']!r}")
    if ch != r["접수채널"]:
        tags.append("접수채널 표기 통일"); r["접수채널"] = ch

    tm = re.sub(r"\s", "", str(r["상담원팀"]))
    if tm != r["상담원팀"]:
        tags.append("상담원팀 띄어쓰기 제거"); r["상담원팀"] = tm

    if r["지역"] in REG:
        tags.append("지역 표기 통일"); r["지역"] = REG[r["지역"]]

    if isinstance(r["월정액"], str) and r["월정액"] != "":
        tags.append("월정액 «원»·쉼표 제거")
        r["월정액"] = int(re.sub(r"[,원]", "", r["월정액"]))

    if r["재문의"] != "":
        y = YN.get(r["재문의"])
        if not y:
            raise ValueError(f"재문의 {r['재문의']!r}")
        if y != r["재문의"]:
            tags.append("재문의 Y/N 통일")
        r["재문의"] = y

    if r["처리상태"] in STATUS:                          # d-005 #1
        tags.append(f"처리상태 «{r['처리상태']}» → «{STATUS[r['처리상태']]}» (d-005 #1)")
        r["처리상태"] = STATUS[r["처리상태"]]

    if r["접수채널"] == "전화" and r["접수일"] >= "2025-07-01":
        r["상담시간"] = math.floor(r["상담시간"] / 60 * 100 + 0.5) / 100
        tags.append("상담시간 초→분 (전화·2025-07-01 이후)")

    flag = []
    m = r["만족도"]
    if isinstance(m, str) and re.fullmatch(r"\d점", m):
        m = int(m[0]); tags.append("만족도 «N점» → 숫자")
    elif m == "-" or m == "미응답" or (is_num(m) and m == 0):
        m = ""; tags.append("만족도 미응답(0·«-»·«미응답») → 빈칸")
    if m != "" and not (is_num(m) and 1 <= m <= 10):
        raise ValueError(f"만족도 {m!r}")
    if is_num(m) and m > 5:                             # d-005 #3: 지우지 않고 표시
        flag.append(f"{SAT_FLAG}({js_str(m)})")
        tags.append("_표시: 만족도 6~10점 (d-005 #3)")
    r["만족도"] = m

    if is_num(r["처리시간"]) and r["처리시간"] < 0:        # d-005 #2: 지우지 않고 표시
        flag.append(f"처리시간 음수({js_str(r['처리시간'])})")
        tags.append("_표시: 처리시간 음수 (d-005 #2)")

    r["_표시"] = " · ".join(flag)
    return r, tags


def main():
    cols, raw = read_source()
    for need in ("접수번호", "고객번호", "상담원ID", "처리상태", "만족도", "처리시간"):
        if need not in cols:
            raise SystemExit(f"원본에 «{need}» 칸이 없다")

    dropped = []                                        # (원본 번호, 이유)
    # 1. 테스트 접수 제외
    rows = []
    for i, r0 in enumerate(raw):
        if str(r0["상담원ID"]).strip() == "AG999":
            dropped.append((i, "테스트 접수(상담원ID AG999)"))
        else:
            r, tags = clean_row(r0, cols)              # 2. 칸별 정리
            rows.append((i, r, tags))

    # 3. 완전 중복 제거 (표기 통일 뒤)
    seen, kept = set(), []
    for i, r, tags in rows:
        k = tuple(r[c] for c in cols + ["_표시"])
        if k in seen:
            dropped.append((i, "완전 중복(모든 칸이 같은 줄이 앞에 있음)"))
            continue
        seen.add(k); kept.append((i, r, tags))

    # 4. 같은 접수번호 → 최종수정일이 늦은 줄만
    latest = {}
    for i, r, tags in kept:
        p = latest.get(r["접수번호"])
        if p is None or r["최종수정일"] > p[1]["최종수정일"]:
            latest[r["접수번호"]] = (i, r)
    out = []
    for i, r, tags in kept:
        if latest[r["접수번호"]][0] == i:
            out.append((i, r, tags))
        else:
            dropped.append((i, "상태 변경 전 줄(같은 접수번호에 최종수정일이 더 늦은 줄 있음)"))
    dropped.sort()

    # 저장 (d-005 #4: 칸 이름은 원본 그대로, 맨 끝에 _표시 만 붙는다)
    OUT_DIR.mkdir(exist_ok=True)
    hdr = cols + ["_표시"]
    write_csv(OUT_CLEAN, hdr, [[esc(r[c]) for c in hdr] for _, r, _ in out])
    write_csv(OUT_DROP, cols + ["원본행", "지운이유"],
              [[esc(raw[i][c]) for c in cols] + [str(i + 2), esc(why)] for i, why in dropped])

    # 숫자
    rs = [r for _, r, _ in out]
    n = len(rs)
    avg_t = sum(r["상담시간"] for r in rs) / n
    rq = [r for r in rs if r["접수일"] >= "2025-03-01"]
    rq_y = sum(r["재문의"] == "Y" for r in rq)
    sat = [r for r in rs if r["만족도"] != "" and SAT_FLAG not in r["_표시"]]   # d-005 #3
    avg_s = sum(r["만족도"] for r in sat) / len(sat)
    done = sum(r["처리상태"] in ("완료", "처리완료") for r in rs)
    cust = len({r["고객번호"] for r in rs})
    cust_raw = len({raw[i]["고객번호"] for i, _, _ in out})

    change = Counter(t for _, _, tags in out for t in tags)
    drop_why = Counter(why for _, why in dropped)
    status = Counter(r["처리상태"] for r in rs)
    flags = Counter(f.split("(")[0] for r in rs for f in r["_표시"].split(" · ") if f)

    def f(x):
        return f"{x:,}"

    L = []
    L.append("# 상담원장 정리 — 변경 요약\n")
    L.append("- 스크립트: `정리/정리_상담원장.py` (따르는 결정 기록: `my-wiki/decisions/d-005-상담원장-정리-다섯-가지.md`)")
    L.append("- 원본: `../messy/상담원장_2025.xlsx` 시트 `원장` (읽기만 함)")
    L.append(f"- 원본 {f(len(raw))}줄 → 정리본 {f(n)}줄 · 지운 줄 {f(len(dropped))}줄 (`정리본/지운줄.csv`, 원본행·지운이유 붙음)")
    L.append("- 줄 수는 **정리본에 남은 줄** 기준이다 (지운 줄에서 바뀐 것은 세지 않음)\n")

    L.append("## 결정 기록 d-005 다섯 가지\n")
    L.append("| # | 정한 것 | 정리본에서 바뀐 줄 |")
    L.append("|---|---|---|")
    st = [k for k in change if k.startswith("처리상태 «")]
    L.append(f"| 1 | 종결 → 완료, 대기 → 보류 | " + " · ".join(f"{k.split(' (')[0].replace('처리상태 ', '')} {f(change[k])}" for k in sorted(st)) + f" (합계 {f(sum(change[k] for k in st))}) |")
    L.append(f"| 2 | 처리시간 음수는 그대로, `_표시` | 값이 바뀐 줄 0 · `_표시` 붙은 줄 {f(change['_표시: 처리시간 음수 (d-005 #2)'])} |")
    L.append(f"| 3 | 만족도 6~10점은 그대로, `_표시`, 평균에서 뺌 | 값이 바뀐 줄 0 · `_표시` 붙은 줄 {f(change['_표시: 만족도 6~10점 (d-005 #3)'])} |")
    L.append("| 4 | 칸 이름 그대로 | 0 (맨 끝에 `_표시` 칸 하나만 붙음) |")
    L.append(f"| 5 | 고객번호 앞뒤 공백 제거 | {f(change['앞뒤 공백 제거: 고객번호'])} · 고객 수 {f(cust_raw)} → {f(cust)} |")
    L.append("")

    L.append("## 상담시간 단위\n")
    L.append(f"- **2025-07-01 이후 접수한 전화 줄 {f(change['상담시간 초→분 (전화·2025-07-01 이후)'])}줄의 상담시간을 초에서 분으로 바꿨다** (÷60, 소수 둘째 자리 반올림).")
    L.append("- 원본 `작성기준` 시트: «2025-07-01 전화 시스템 교체 뒤로 전화 채널은 초 단위로 기록됨». 다른 줄은 원래 분 단위라 그대로 두었다.")
    L.append("- 그래서 정리본의 상담시간은 모든 줄이 **분** 단위다. 원본과 비교할 때 이 줄들은 값이 60분의 1로 보인다.\n")

    L.append("## 그 밖에 바꾼 것 (이전 스크립트에 있던 처리)\n")
    L.append("| 바꾼 것 | 정리본에서 바뀐 줄 |")
    L.append("|---|---|")
    for k, v in sorted(change.items()):
        if "d-005" in k or k.startswith("상담시간") or k == "앞뒤 공백 제거: 고객번호":
            continue
        L.append(f"| {k} | {f(v)} |")
    L.append("")

    L.append("## 지운 줄\n")
    L.append("| 이유 | 줄 |")
    L.append("|---|---|")
    for k, v in sorted(drop_why.items()):
        L.append(f"| {k} | {f(v)} |")
    L.append("")

    L.append("## 정리본 숫자\n")
    L.append("| 숫자 | 값 | 기준 |")
    L.append("|---|---|---|")
    L.append(f"| 줄 수 | {f(n)} | |")
    L.append(f"| 평균 상담시간 | {avg_t:.2f}분 | 모든 줄 |")
    L.append(f"| 재문의율 | {rq_y / len(rq) * 100:.2f}% | 2025-03-01 이후 접수 {f(len(rq))}줄 중 Y {f(rq_y)} |")
    L.append(f"| 평균 만족도 | {avg_s:.2f}점 | 응답 {f(len(sat))}줄 (6~10점 표시 줄 뺌, d-005 #3) |")
    L.append(f"| 완료율 | {done / n * 100:.2f}% | 완료·처리완료 {f(done)}줄 |")
    L.append(f"| 고객 수 | {f(cust)} | 고객번호 종류 (공백 제거 전 {f(cust_raw)}) |")
    L.append("")
    L.append("처리상태: " + " · ".join(f"{k} {f(v)}" for k, v in status.most_common()))
    L.append("")
    L.append("`_표시`: " + " · ".join(f"{k} {f(v)}" for k, v in flags.most_common()))
    L.append("")
    L.append("## 하지 않은 것\n")
    L.append("- «처리완료»·«처리중»은 합치지 않았다 (d-005 «아직 모르는 것»: 정하지 않음).")
    OUT_SUMMARY.write_bytes(("\n".join(L) + "\n").encode("utf-8"))

    print(f"원본 {len(raw)} → 정리본 {n} · 지운 줄 {len(dropped)}")
    print(f"평균 상담시간 {avg_t:.2f} · 재문의율 {rq_y / len(rq) * 100:.2f} · 평균 만족도 {avg_s:.2f} · 완료율 {done / n * 100:.2f}")
    print(f"고객 수 {cust} (공백 제거 전 {cust_raw})")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
