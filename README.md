# School Supplies Auto-Fill API (Backend)

## What it does
POST a list of i-screammall product URLs and get back:
- name (품명)
- unit_price (단가)
- product_code (상품코드)
- site (아이스크림몰)

## Local run
```bash
pip install -r requirements.txt
uvicorn app.main:app --reload
```
Test:
```bash
curl -X POST http://127.0.0.1:8000/api/scrape -H "Content-Type: application/json" -d '{"urls":["https://i-screammall.co.kr/goods/detail/10970545"]}'
```

## Deploy on Render (Docker)
- Create a new **Web Service**
- Select **Docker**
- Connect this repo
- Deploy
