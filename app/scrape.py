import re
import httpx
from bs4 import BeautifulSoup
from urllib.parse import urlparse
from playwright.sync_api import sync_playwright

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
    snippet = text[idx:idx+1200] if idx >= 0 else text[:1200]
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

def _fetch_html_httpx(url: str) -> str:
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

    timeout = httpx.Timeout(35.0, connect=25.0)
    limits = httpx.Limits(max_connections=10, max_keepalive_connections=5)

    last_exc = None
    with httpx.Client(timeout=timeout, follow_redirects=True, headers=headers, limits=limits) as client:
        for _ in range(2):
            try:
                r = client.get(url)
                if r.status_code >= 400:
                    raise httpx.HTTPStatusError(f"HTTP {r.status_code}", request=r.request, response=r)
                if not r.text or len(r.text) < 500:
                    raise RuntimeError("응답 본문이 비정상적으로 짧습니다(차단/리다이렉트 가능).")
                return r.text
            except Exception as e:
                last_exc = e

    raise RuntimeError(f"httpx 실패: {type(last_exc).__name__}: {last_exc}")

def _fetch_html_playwright(url: str) -> str:
    # 브라우저 자동화(봇차단 우회 가능성이 높음)
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            locale="ko-KR",
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/122.0.0.0 Safari/537.36"
            ),
        )
        page = context.new_page()
        page.goto(url, wait_until="domcontentloaded", timeout=60000)
        # 필요한 경우 네트워크 안정화 대기
        page.wait_for_timeout(1500)
        html = page.content()
        context.close()
        browser.close()
        if not html or len(html) < 500:
            raise RuntimeError("Playwright 응답이 비정상적으로 짧습니다(차단/권한 가능).")
        return html

def scrape_iscreammall(url: str) -> dict:
    parsed = urlparse(url)
    if "i-screammall.co.kr" not in parsed.netloc:
        raise ValueError("현재는 i-screammall.co.kr 상품 URL만 지원합니다.")

    # 1) 먼저 가벼운 httpx 시도
    try:
        html = _fetch_html_httpx(url)
    except Exception:
        # 2) 실패하면 Playwright로 재시도(완전 자동 목표)
        html = _fetch_html_playwright(url)

    soup = BeautifulSoup(html, "html.parser")
    name = _extract_name(soup)
    text = soup.get_text("\n", strip=True)

    product_code = _extract_product_code(url)

    unit_price = _pick_price_from_text(text, name)
    if unit_price == 0:
        unit_price = _pick_price_from_text(text, "할인혜택받기")
    if unit_price == 0:
        unit_price = _pick_price_from_text(text, "상품번호")

    if not name:
        raise RuntimeError("품명을 찾지 못했습니다(차단/페이지 구조 변경 가능).")

    return {
        "name": name,
        "unit_price": unit_price,
        "product_code": product_code,
        "site": "아이스크림몰",
    }
