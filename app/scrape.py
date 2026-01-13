import re
import httpx
from bs4 import BeautifulSoup
from urllib.parse import urlparse

def _extract_product_code(url: str) -> str:
    # works for /goods/detail/10970545
    m = re.search(r"/goods/detail/(\d+)", url)
    if m:
        return m.group(1)
    # fallback: last long digit sequence
    m2 = re.findall(r"(\d{6,})", url)
    return m2[-1] if m2 else ""

def _clean_price_to_int(s: str) -> int:
    s = s.replace(",", "").strip()
    return int(s) if s.isdigit() else 0

def _pick_price_from_text(text: str, anchor: str) -> int:
    """
    Heuristic:
    - Find the first occurrence of anchor (usually product name),
    - Take the next ~500 chars and look for '원' prices,
    - If 2+ found, return the second one (판매가가 두 번째로 나오는 패턴 대응).
    - Else return the first found.
    """
    idx = text.find(anchor) if anchor else -1
    snippet = text[idx:idx+700] if idx >= 0 else text[:700]
    prices = re.findall(r"([0-9]{1,3}(?:,[0-9]{3})*)\s*원", snippet)
    ints = [_clean_price_to_int(p) for p in prices if _clean_price_to_int(p) > 0]
    # Remove very small points like 20P mistakenly captured (won't match '원' anyway)
    if len(ints) >= 2:
        # common pattern: 정가, 판매가
        return ints[1]
    if len(ints) == 1:
        return ints[0]
    return 0

def _extract_name(soup: BeautifulSoup) -> str:
    # Try og:title first
    og = soup.find("meta", attrs={"property": "og:title"})
    if og and og.get("content"):
        return og["content"].strip()
    # Fallback: first strong/h1/h2 like elements
    for tag in ["h1", "h2", "h3"]:
        el = soup.find(tag)
        if el and el.get_text(strip=True):
            return el.get_text(" ", strip=True)
    # Last resort: title
    if soup.title and soup.title.get_text(strip=True):
        return soup.title.get_text(strip=True)
    return ""

def scrape_iscreammall(url: str) -> dict:
    """
    Returns minimal fields for 공모전 2차 작품 MVP:
    - 품명(name)
    - 단가(unit_price)
    - 상품코드(product_code)
    - 사이트(site)
    """
    parsed = urlparse(url)
    if "i-screammall.co.kr" not in parsed.netloc:
        raise ValueError("현재는 i-screammall.co.kr 상품 URL만 지원합니다.")

    headers = {
        "User-Agent": "Mozilla/5.0 (compatible; SchoolSuppliesBot/0.1)"
    }
    with httpx.Client(timeout=20, follow_redirects=True, headers=headers) as client:
        r = client.get(url)
        r.raise_for_status()
        html = r.text

    soup = BeautifulSoup(html, "html.parser")
    name = _extract_name(soup)
    text = soup.get_text("\n", strip=True)

    product_code = _extract_product_code(url)
    # price heuristic around name, then around '할인혜택받기', then before '상품번호'
    unit_price = _pick_price_from_text(text, name)
    if unit_price == 0:
        unit_price = _pick_price_from_text(text, "할인혜택받기")
    if unit_price == 0:
        unit_price = _pick_price_from_text(text, "상품번호")

    return {
        "name": name,
        "unit_price": unit_price,
        "product_code": product_code,
        "site": "아이스크림몰",
    }
