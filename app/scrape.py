import re
import httpx
from bs4 import BeautifulSoup
from urllib.parse import urlparse

def _extract_product_code(url: str) -> str:
    m = re.search(r"/goods/detail/(\d+)", url)
    if m:
        return m.group(1)
    m2 = re.findall(r"(\d{6,})", url)
    return m2[-1] if m2 else ""

def _clean_price_to_int(s: str) -> int:
    s = s.replace(",", "").strip()
    return int(s) if s.isdigit() else 0

def _pick_price_from_text(text: str, anchor: str) -> int:
    idx = text.find(anchor) if anchor else -1
    snippet = text[idx:idx+900] if idx >= 0 else text[:900]
    prices = re.findall(r"([0-9]{1,3}(?:,[0-9]{3})*)\s*원", snippet)
    ints = [_clean_price_to_int(p) for p in prices if _clean_price_to_int(p) > 0]
    if len(ints) >= 2:
        return ints[1]  # 정가/판매가 패턴
    if len(ints) == 1:
        return ints[0]
    return 0

def _extract_name(soup: BeautifulSoup) -> str:
    og = soup.find("meta", attrs={"property": "og:title"})
    if og and og.get("content"):
        return og["content"].strip()
    for tag in ["h1", "h2", "h3"]:
        el = soup.find(tag)
        if el and el.get_text(strip=True):
            return el.get_text(" ", strip=True)
    if soup.title and soup.title.get_text(strip=True):
        return soup.title.get_text(strip=True)
    return ""

def _fetch_html(url: str) -> str:
    # Cloudflare/봇차단 대응: 브라우저 유사 헤더 + 재시도
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/122.0.0.0 Safari/537.36"
        ),
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
        "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
        "Cache-Control": "no-cache",
        "Pragma": "no-cache",
        "Referer": "https://i-screammall.co.kr/",
        "Connection": "keep-alive",
    }

    timeout = httpx.Timeout(30.0, connect=20.0)
    limits = httpx.Limits(max_connections=10, max_keepalive_connections=5)

    last_exc = None
    with httpx.Client(timeout=timeout, follow_redirects=True, headers=headers, limits=limits) as client:
        for attempt in range(3):
            try:
                r = client.get(url)
                # 일부 보호 페이지는 200이어도 본문이 비정상일 수 있어 최소 체크
                if r.status_code >= 400:
                    raise httpx.HTTPStatusError(f"HTTP {r.status_code}", request=r.request, response=r)
                if not r.text or len(r.text) < 500:
                    raise RuntimeError("응답 본문이 비정상적으로 짧습니다(차단/리다이렉트 가능).")
                return r.text
            except Exception as e:
                last_exc = e
    # 마지막 예외 메시지를 최대한 자세히
    if isinstance(last_exc, httpx.HTTPStatusError):
        resp = last_exc.response
        raise RuntimeError(f"요청 실패: HTTP {resp.status_code} (일부 차단/권한/봇방지 가능)")
    raise RuntimeError(f"요청 실패: {type(last_exc).__name__}: {last_exc}")

def scrape_iscreammall(url: str) -> dict:
    parsed = urlparse(url)
    if "i-screammall.co.kr" not in parsed.netloc:
        raise ValueError("현재는 i-screammall.co.kr 상품 URL만 지원합니다.")

    html = _fetch_html(url)
    soup = BeautifulSoup(html, "html.parser")

    name = _extract_name(soup)
    text = soup.get_text("\n", strip=True)

    product_code = _extract_product_code(url)

    unit_price = _pick_price_from_text(text, name)
    if unit_price == 0:
        unit_price = _pick_price_from_text(text, "할인혜택받기")
    if unit_price == 0:
        unit_price = _pick_price_from_text(text, "상품번호")

    # 보호페이지/오류페이지 감지(품명이 비어있으면 실패로 처리)
    if not name:
        raise RuntimeError("품명을 찾지 못했습니다(차단/페이지 구조 변경 가능).")

    return {
        "name": name,
        "unit_price": unit_price,
        "product_code": product_code,
        "site": "아이스크림몰",
    }
