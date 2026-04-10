"""
OCR 금액 인식 진단 도구
사용법: python ocr_debug.py <이미지파일> [이미지파일2 ...]
또는:   python ocr_debug.py <폴더경로>

실패한 영수증 이미지를 넣으면 각 OCR 시도의 원문과
추출된 금액 후보를 상세히 출력합니다.
"""
import re
import sys
from pathlib import Path

from PIL import Image, ImageFilter, ImageOps
import pytesseract

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# ── 상수 ──────────────────────────────────────────────────────────────────────
RESAMPLE_LANCZOS = Image.Resampling.LANCZOS if hasattr(Image, "Resampling") else Image.LANCZOS
SEARCH_TOP_RATIO = 0.01
SEARCH_BOTTOM_RATIO_DARK_BG = 0.40
SEARCH_BOTTOM_RATIO_LIGHT_BG = 0.58
BACKGROUND_DARK_THRESHOLD = 132

FULLWIDTH_TRANS = str.maketrans({
    "０": "0", "１": "1", "２": "2", "３": "3", "４": "4",
    "５": "5", "６": "6", "７": "7", "８": "8", "９": "9",
    "．": ".", "，": ",", "￥": "¥",
})

TOTAL_CONTEXT_AMOUNT_RE = re.compile(
    r"(?:总金额|总计|合计|应付|实付|支付金额|결제금액|총금액|합계|total)\D{0,8}-?\s*([0-9][0-9,]*(?:\.[0-9]{1,2})?)",
    re.IGNORECASE,
)
NEGATIVE_OR_PLAIN_AMOUNT_RE = re.compile(r"-\s*([0-9][0-9,]*(?:\.[0-9]{1,2})?)")
CURRENCY_AMOUNT_RE = re.compile(r"(?:[¥￥]|RMB|CNY)\s*([0-9][0-9,]*(?:\.[0-9]{1,2})?)", re.IGNORECASE)
CONTEXT_AMOUNT_RE = re.compile(
    r"(?:합계|총계|총액|금액|결제금액|실결제|실지불|实付|应付|合计|总计|总额|金额)\D{0,6}([0-9][0-9,]*(?:\.[0-9]{1,2})?)",
    re.IGNORECASE,
)
GENERIC_NUMBER_RE = re.compile(r"([0-9]{1,3}(?:,[0-9]{3})*(?:\.[0-9]{1,2})?|[0-9]+(?:\.[0-9]{1,2})?)")
DASH_AMOUNT_RE = re.compile(r"-\s*([0-9][0-9,]*(?:\.[0-9]{1,2})?)")


# ── 유틸 함수 (메인 앱에서 복사) ──────────────────────────────────────────────
def _otsu_threshold(gray_image):
    hist = gray_image.histogram()
    total = sum(hist)
    if total == 0:
        return 128
    sum_all = sum(i * v for i, v in enumerate(hist))
    w_bg = sum_bg = 0
    max_var = best = 0
    for t, freq in enumerate(hist):
        w_bg += freq
        if w_bg == 0:
            continue
        w_fg = total - w_bg
        if w_fg == 0:
            break
        sum_bg += t * freq
        mean_bg = sum_bg / w_bg
        mean_fg = (sum_all - sum_bg) / w_fg
        var = w_bg * w_fg * (mean_bg - mean_fg) ** 2
        if var > max_var:
            max_var, best = var, t
    return best


def parse_amount_token(token):
    token = token.replace(",", "").strip()
    if token.count(".") > 1:
        return None
    try:
        value = abs(float(token))
    except (TypeError, ValueError):
        return None
    if value == 0 or value > 1_000_000:
        return None
    return round(value, 2)


def extract_amount_from_text(raw_text):
    if not raw_text:
        return None
    text = str(raw_text).translate(FULLWIDTH_TRANS)
    candidates = []
    for match in re.compile(r"-\s*([0-9][0-9,]*\.[0-9]{2})").finditer(text):
        v = parse_amount_token(match.group(1))
        if v: candidates.append((4, v))
    for match in CURRENCY_AMOUNT_RE.finditer(text):
        v = parse_amount_token(match.group(1))
        if v: candidates.append((3, v))
    for match in CONTEXT_AMOUNT_RE.finditer(text):
        v = parse_amount_token(match.group(1))
        if v: candidates.append((2, v))
    for match in GENERIC_NUMBER_RE.finditer(text):
        v = parse_amount_token(match.group(1))
        if v is None: continue
        if float(v).is_integer() and 1900 <= int(v) <= 2100: continue
        candidates.append((1, v))
    if not candidates:
        return None
    top_score = max(s for s, _ in candidates)
    return round(max(v for s, v in candidates if s == top_score), 2)


def extract_total_amount_from_text(raw_text, return_score=False):
    if not raw_text:
        return (None, None) if return_score else None
    text = str(raw_text).translate(FULLWIDTH_TRANS)
    candidates = []
    for match in TOTAL_CONTEXT_AMOUNT_RE.finditer(text):
        v = parse_amount_token(match.group(1))
        if v: candidates.append((6 if "." in match.group(1) else 5, v))
    for match in NEGATIVE_OR_PLAIN_AMOUNT_RE.finditer(text):
        v = parse_amount_token(match.group(1))
        if v: candidates.append((4 if "." in match.group(1) else 2, v))
    fallback = extract_amount_from_text(text)
    if fallback: candidates.append((3, fallback))
    if not candidates:
        return (None, None) if return_score else None
    top_score = max(s for s, _ in candidates)
    selected = round(max(v for s, v in candidates if s == top_score), 2)
    return (selected, top_score) if return_score else selected


def extract_dash_amount_from_text(raw_text, return_score=False):
    if not raw_text:
        return (None, None) if return_score else None
    text = str(raw_text).translate(FULLWIDTH_TRANS)
    text = text.replace("—", "-").replace("–", "-").replace("−", "-")
    candidates = []
    for match in DASH_AMOUNT_RE.finditer(text):
        v = parse_amount_token(match.group(1))
        if v: candidates.append((4 if "." in match.group(1) else 2, v))
    fallback = extract_amount_from_text(text)
    if fallback: candidates.append((3, fallback))
    if not candidates:
        return (None, None) if return_score else None
    top_score = max(s for s, _ in candidates)
    selected = round(max(v for s, v in candidates if s == top_score), 2)
    return (selected, top_score) if return_score else selected


def is_dark_background_image(image):
    gray = ImageOps.grayscale(ImageOps.exif_transpose(image).convert("RGB"))
    w, h = gray.size
    ew, th = max(2, int(w * 0.08)), max(2, int(h * 0.22))
    regions = [gray.crop((0, 0, w, th)), gray.crop((0, 0, ew, h)), gray.crop((max(0, w - ew), 0, w, h))]
    samples = []
    for r in regions:
        samples.extend(r.getdata())
    if not samples:
        return False
    samples.sort()
    return samples[len(samples) // 2] < BACKGROUND_DARK_THRESHOLD


# ── 진단 함수 ─────────────────────────────────────────────────────────────────
def _sep(label=""):
    print(f"\n{'─'*64}")
    if label:
        print(f"  {label}")
        print(f"{'─'*64}")


def debug_image(image_path: Path):
    print(f"\n{'='*64}")
    print(f"  이미지: {image_path.name}")
    print(f"{'='*64}")

    with Image.open(image_path) as img:
        base = ImageOps.exif_transpose(img).convert("RGB")

    w, h = base.size
    is_dark = is_dark_background_image(base)
    print(f"  크기: {w} x {h}px  |  배경: {'어두움' if is_dark else '밝음'}")

    # ── 방식 1: 상단 dash 탐지 ────────────────────────────────────────
    _sep("방식1 — 상단 dash 영역 탐지")
    bottom_ratio = SEARCH_BOTTOM_RATIO_DARK_BG if is_dark else SEARCH_BOTTOM_RATIO_LIGHT_BG
    top_px = max(0, int(h * SEARCH_TOP_RATIO))
    bot_px = min(h, max(top_px + 1, int(h * bottom_ratio)))
    print(f"  탐색범위: {top_px}px ~ {bot_px}px  (상위 {bottom_ratio*100:.0f}%)")

    band = base.crop((0, top_px, w, bot_px))
    enl = band.resize((max(1, band.width * 3), max(1, band.height * 4)), RESAMPLE_LANCZOS)
    gray = ImageOps.grayscale(enl)
    sharp = gray.filter(ImageFilter.SHARPEN)
    auto = ImageOps.autocontrast(sharp)
    inv = ImageOps.invert(auto)
    ot = _otsu_threshold(auto)
    ot_inv = _otsu_threshold(inv)
    print(f"  Otsu 임계값: {ot} (일반) / {ot_inv} (반전)")

    if is_dark:
        targets1 = [
            (inv.point(lambda x: 255 if x > 150 else 0, mode="1"), "6", "dark_fixed150"),
            (inv.point(lambda x, t=ot_inv: 255 if x > t else 0, mode="1"), "6", f"dark_otsu{ot_inv}"),
            (inv, "6", "inv_raw"), (inv, "11", "inv_raw"),
        ]
    else:
        targets1 = [
            (auto.point(lambda x: 255 if x > 165 else 0, mode="1"), "6", "light_fixed165"),
            (auto.point(lambda x, t=ot: 255 if x > t else 0, mode="1"), "6", f"light_otsu{ot}"),
            (auto, "6", "auto_raw"), (auto, "11", "auto_raw"),
        ]

    dash_candidates = []
    for ocr_img, psm, label in targets1:
        cfg = f"--oem 1 --psm {psm} -c tessedit_char_whitelist=0123456789.,-"
        try:
            text = pytesseract.image_to_string(ocr_img, lang="chi_sim+eng", config=cfg).strip()
        except Exception as e:
            print(f"  [{label:<20s} psm{psm}] 오류: {e}")
            continue
        amount, score = extract_dash_amount_from_text(text, return_score=True)
        result = f"→ {amount:.2f} (score={score})" if amount else "→ 없음"
        print(f"  [{label:<20s} psm{psm}] {repr(text[:50]):<55s}  {result}")
        if amount is not None:
            dash_candidates.append((amount, score))

    # ── 방식 2: 전체 영역 폴백 ────────────────────────────────────────
    _sep("방식2 — 전체 영역 탐지 (폴백)")
    regions = [
        ("전체", base),
        ("하단45%", base.crop((0, int(h * 0.45), w, h))),
        ("중앙하단", base.crop((int(w * 0.2), int(h * 0.38), int(w * 0.8), h))),
    ]
    full_candidates = []
    for rname, region in regions:
        gray_r = ImageOps.grayscale(region)
        hc = ImageOps.autocontrast(gray_r.filter(ImageFilter.SHARPEN))
        enl_r = hc.resize((max(1, hc.width * 3), max(1, hc.height * 3)), RESAMPLE_LANCZOS)
        ot_r = _otsu_threshold(enl_r)
        variants = [
            (enl_r, "확대3x"),
            (enl_r.point(lambda x, t=ot_r: 255 if x > t else 0, mode="1"), f"Otsu({ot_r})"),
            (enl_r.point(lambda x: 255 if x > 160 else 0, mode="1"), "fixed160"),
        ]
        for v_img, vname in variants:
            for psm in ("6", "7", "11"):
                try:
                    text = pytesseract.image_to_string(v_img, lang="chi_sim+eng", config=f"--oem 1 --psm {psm}").strip()
                except Exception:
                    continue
                amount, score = extract_total_amount_from_text(text, return_score=True)
                if amount is not None:
                    label2 = f"{rname}/{vname}"
                    print(f"  [{label2:<22s} psm{psm}] {repr(text[:45]):<50s}  → {amount:.2f} (score={score})")
                    full_candidates.append((amount, score))

    # ── 최종 결과 ─────────────────────────────────────────────────────
    _sep("결과 요약")
    all_c = dash_candidates if dash_candidates else full_candidates
    method = "dash(상단)" if dash_candidates else "전체영역(폴백)"

    if not all_c:
        print("  ❌ 금액 추출 실패 — 모든 방법에서 후보 없음")
        return None

    stats = {}
    for val, sc in all_c:
        k = round(val, 2)
        if k not in stats:
            stats[k] = {"count": 0, "score_sum": 0.0, "score_max": 0.0}
        stats[k]["count"] += 1
        stats[k]["score_sum"] += float(sc)
        stats[k]["score_max"] = max(stats[k]["score_max"], float(sc))

    print(f"  사용 방법: {method}")
    print(f"  후보 금액 목록 (출현횟수 높은 순):")
    for val, info in sorted(stats.items(), key=lambda x: -x[1]["count"]):
        bar = "█" * info["count"]
        print(f"    {val:>10.2f}  {bar:<10s} 횟수={info['count']}  max_score={info['score_max']:.0f}")

    best = max(stats.items(), key=lambda kv: (kv[1]["count"], kv[1]["score_max"], kv[1]["score_sum"], kv[0]))[0]
    print(f"\n  ✅ 최종 선택: {best:.2f}")
    return best


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    paths = []
    for arg in sys.argv[1:]:
        p = Path(arg)
        if p.is_dir():
            exts = (".jpg", ".jpeg", ".png")
            paths.extend(sorted(f for f in p.iterdir() if f.suffix.lower() in exts))
        elif p.exists():
            paths.append(p)
        else:
            print(f"파일 없음: {arg}")

    if not paths:
        print("처리할 이미지가 없습니다.")
        sys.exit(1)

    results = {}
    for path in paths:
        amount = debug_image(path)
        results[path.name] = amount

    print(f"\n{'='*64}")
    print("  전체 결과 요약")
    print(f"{'='*64}")
    for name, amount in results.items():
        status = f"{amount:.2f}" if amount else "❌ 실패"
        print(f"  {name:<45s}  {status}")


if __name__ == "__main__":
    main()
