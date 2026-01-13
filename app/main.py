from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, HttpUrl
from typing import List, Optional, Dict, Any
from .scrape import scrape_iscreammall

app = FastAPI(title="School Supplies Auto-Fill API", version="0.1.0")

# NOTE: For demo simplicity we allow all origins.
# In production, set this to your GitHub Pages domain only.
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

class ScrapeRequest(BaseModel):
    urls: List[str]

@app.get("/health")
def health():
    return {"ok": True}

@app.post("/api/scrape")
def api_scrape(req: ScrapeRequest):
    results = []
    for url in req.urls:
        url = url.strip()
        if not url:
            continue
        try:
            item = scrape_iscreammall(url)
            results.append({"url": url, "ok": True, "data": item})
        except Exception as e:
            results.append({"url": url, "ok": False, "error": str(e)})
    return {"count": len(results), "results": results}
