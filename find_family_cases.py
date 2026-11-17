"""
find_family_cases.py
Finds family-law judgments on caselaw.shc.gov.pk (Sindh High Court judgment portal)
and writes a spreadsheet with the search criteria to find each case manually.

Run on your own laptop (the court site blocks cloud servers):
    python -m pip install requests beautifulsoup4 pandas openpyxl pypdf
    python find_family_cases.py --test          # 2 judges only, quick check
    python find_family_cases.py                 # full run, list pages only (fast)
    python find_family_cases.py --text          # also download each judgment and confirm by keywords (slow, best)

Output: family_cases.xlsx and family_cases.csv in the same folder.
Safe to stop and restart: every page is cached in ./shc_cache.
"""
import argparse, base64, csv, hashlib, io, os, re, sys, time
import requests
from bs4 import BeautifulSoup

BASE = "https://caselaw.shc.gov.pk/caselaw/public/"
CACHE = "shc_cache"
HEADERS = {"User-Agent": "Mozilla/5.0 (academic research; Habib University FYP)"}

# Category ids seen on the site: 15 = Const. P., 73 = Family Appeal, 6 = M.A.
DEFAULT_CATS = {"Const. P.", "Family Appeal", "M.A."}

FAMILY_TERMS = {
    "family court": "family court", "family judge": "family court",
    "family courts act": "family court", "family suit": "family court",
    "west pakistan family courts": "family court",
    "khula": "khula / dissolution", "khulla": "khula / dissolution",
    "dissolution of marriage": "khula / dissolution", "dissolution of muslim marriages": "khula / dissolution",
    "talaq": "talaq", "divorce": "talaq",
    "maintenance": "maintenance", "iddat": "maintenance",
    "custody": "custody / guardianship", "guardian": "custody / guardianship",
    "hizanat": "custody / guardianship", "minor": "custody / guardianship", "visitation": "custody / guardianship",
    "dower": "dower", "haq mehr": "dower", "mehr": "dower",
    "dowry": "dowry articles", "jahez": "dowry articles",
    "restitution of conjugal": "restitution",
    "muslim family laws ordinance": "MFLO", "nikahnama": "MFLO",
}
WOMAN_MARKERS = re.compile(r"\b(mst|mst\.|mrs|mrs\.|bibi|begum|ms\.)\b", re.I)

session = requests.Session()
session.headers.update(HEADERS)
DELAY = 1.0


def get(url, binary=False):
    """GET with cache, polite delay and retries."""
    os.makedirs(CACHE, exist_ok=True)
    key = hashlib.md5(url.encode()).hexdigest()
    path = os.path.join(CACHE, key + (".bin" if binary else ".html"))
    if os.path.exists(path):
        with open(path, "rb") as f:
            data = f.read()
        return data if binary else data.decode("utf-8", "ignore")
    for attempt in range(4):
        try:
            time.sleep(DELAY)
            r = session.get(url, timeout=60)
            if r.status_code == 200:
                with open(path, "wb") as f:
                    f.write(r.content)
                return r.content if binary else r.text
            print(f"  HTTP {r.status_code} for {url}")
        except requests.RequestException as e:
            print(f"  retry {attempt + 1}: {e}")
            time.sleep(5 * (attempt + 1))
    return None


def judge_ids():
    """Collect judge ids from the report pages (sitting and, where listed, former judges)."""
    ids = {}
    for page in ["rpt-sbdb-cons"]:
        html = get(BASE + page)
        if not html:
            continue
        soup = BeautifulSoup(html, "html.parser")
        for a in soup.find_all("a", href=True):
            m = re.search(r"rpt-[a-z-]+/(\d+)/", a["href"])
            if m and "sbdb-cons-total" in a["href"]:
                ids.setdefault(m.group(1), a.get_text(" ", strip=True))
        # judge name is usually in the same row as the link
        for tr in soup.find_all("tr"):
            a = tr.find("a", href=re.compile(r"rpt-sbdb-cons-total/\d+/"))
            if a:
                jid = re.search(r"/(\d+)/", a["href"]).group(1)
                name = " ".join(td.get_text(" ", strip=True) for td in tr.find_all("td") if "Justice" in td.get_text())
                if name:
                    ids[jid] = name
    return ids


def category_links(jid):
    """From a judge's totals page, return [(category, bench 'S.B'/'D.B', url)]."""
    html = get(f"{BASE}rpt-sbdb-cons-total/{jid}/CAT/-1")
    out = []
    if not html:
        return out
    soup = BeautifulSoup(html, "html.parser")
    for tr in soup.find_all("tr"):
        links = tr.find_all("a", href=re.compile(r"rpt-cat-judgement-list/"))
        if not links:
            continue
        cells = [td.get_text(" ", strip=True) for td in tr.find_all("td")]
        cat = next((c for c in cells if c and not c.isdigit()), "?")
        for a in links:
            m = re.search(r"rpt-cat-judgement-list/\d+/(\d)/", a["href"])
            bench = {"1": "S.B", "2": "D.B"}.get(m.group(1) if m else "", "?")
            href = a["href"] if a["href"].startswith("http") else BASE + a["href"].split("public/")[-1]
            out.append((cat, bench, href))
    return out


def parse_list(url):
    """Parse a judgment list page into dict rows."""
    html = get(url)
    rows = []
    if not html:
        return rows
    soup = BeautifulSoup(html, "html.parser")
    for table in soup.find_all("table"):
        headers = [th.get_text(" ", strip=True).lower() for th in table.find_all("th")]
        for tr in table.find_all("tr"):
            link = tr.find("a", href=re.compile(r"view-file/"))
            if not link:
                continue
            tds = [td.get_text(" ", strip=True) for td in tr.find_all("td")]
            row = dict(zip(headers, tds)) if len(headers) == len(tds) else {}
            row["_cells"] = tds
            row["_view"] = link["href"] if link["href"].startswith("http") else "https://caselaw.shc.gov.pk/caselaw/" + link["href"].split("caselaw/")[-1]
            rows.append(row)
    return rows


def pick(row, *names):
    for k, v in row.items():
        if not k.startswith("_") and any(n in k for n in names):
            return v
    return ""


CASE_RE = re.compile(r"^(?P<cat>.+?)\s+(?P<no>[A-Z]?-?\d+)\s*/\s*(?P<year>\d{4})\s*(?:\((?P<bench>[SD]\.B\.)\))?\s*(?P<court>.*)$")


def split_case(caseno):
    m = CASE_RE.match(caseno.strip())
    if not m:
        return {"category": "", "number": "", "year": "", "seat": ""}
    seat = m.group("court").split(",")[-1].strip() if "," in m.group("court") else m.group("court").strip()
    return {"category": m.group("cat").strip(), "number": m.group("no"), "year": m.group("year"), "seat": seat}


def judgment_text(url):
    data = get(url, binary=True)
    if not data:
        return ""
    if data[:4] == b"%PDF":
        try:
            from pypdf import PdfReader
            return "\n".join((p.extract_text() or "") for p in PdfReader(io.BytesIO(data)).pages)
        except Exception as e:
            print(f"  pdf read failed: {e}")
            return ""
    return BeautifulSoup(data.decode("utf-8", "ignore"), "html.parser").get_text(" ")


def family_hits(text):
    t = text.lower()
    found, topics = [], set()
    for term, topic in FAMILY_TERMS.items():
        if term in t:
            found.append(term)
            if topic != "family court":
                topics.add(topic)
    return found, topics


def main():
    global DELAY
    ap = argparse.ArgumentParser()
    ap.add_argument("--test", action="store_true", help="only 2 judges")
    ap.add_argument("--text", action="store_true", help="download each judgment and confirm with keywords")
    ap.add_argument("--all-categories", action="store_true", help="scan every category, not just Const. P., Family Appeal, M.A.")
    ap.add_argument("--delay", type=float, default=1.0, help="seconds between requests (be polite)")
    args = ap.parse_args()
    DELAY = args.delay

    judges = judge_ids()
    print(f"Found {len(judges)} judges")
    if not judges:
        sys.exit("No judges found. Open https://caselaw.shc.gov.pk/caselaw/public/rpt-sbdb-cons in a browser to check the site is up.")
    items = list(judges.items())[:2] if args.test else list(judges.items())

    seen, results = set(), []
    for n, (jid, jname) in enumerate(items, 1):
        cats = category_links(jid)
        print(f"[{n}/{len(items)}] judge {jid} {jname[:40]}: {len(cats)} category links")
        for cat, bench, url in cats:
            family_named = re.search(r"divorce|family|guardian|khula|dower|maintenance|matrimon|custody|marriage", cat, re.I)
            if family_named:
                print(f"    family-named category found: {cat} ({bench})")
            if not args.all_categories and not family_named and not any(cat.startswith(c) for c in DEFAULT_CATS):
                continue
            for row in parse_list(url):
                caseno = pick(row, "case no") or (row["_cells"][3] if len(row["_cells"]) > 3 else "")
                key = row["_view"]
                if key in seen:
                    continue
                seen.add(key)
                parties = pick(row, "parties")
                parts = split_case(caseno)
                rec = {
                    "case_no_as_listed": caseno,
                    "category": parts["category"] or cat,
                    "number": parts["number"],
                    "year": parts["year"] or pick(row, "year"),
                    "bench": pick(row, "bench") or bench,
                    "seat": parts["seat"],
                    "parties": parties,
                    "order_date": pick(row, "order"),
                    "judge": jname,
                    "woman_party": "yes" if WOMAN_MARKERS.search(parties) else "",
                    "afr": pick(row, "a.f.r", "afr"),
                    "words": "", "type": "",
                    "judgment_url": row["_view"],
                    "family_terms": "", "topics": "", "verdict": "",
                }
                if family_named and not args.text:
                    rec["verdict"] = "maybe (family category)"
                if args.text and not rec["bench"].upper().startswith("S") and not family_named:
                    rec["verdict"] = "skipped (double bench)"
                elif args.text:
                    txt = judgment_text(row["_view"])
                    found, topics = family_hits(txt)
                    wc = len(txt.split())
                    rec["words"] = wc
                    rec["type"] = "full judgment" if (wc >= 1200 or re.search(r"\bJ\s*U\s*D\s*G\s*M\s*E\s*N\s*T\b", txt)) else "short order"
                    rec["family_terms"] = ", ".join(found)
                    rec["topics"] = ", ".join(sorted(topics))
                    strong = any(t in found for t in ("family court", "family judge", "family courts act", "family suit", "khula", "dissolution of marriage", "haq mehr", "nikahnama", "restitution of conjugal"))
                    rec["verdict"] = ("family" if strong else ("maybe" if found else "not family"))
                else:
                    rec["verdict"] = rec["verdict"] or ("maybe (check text)" if rec["woman_party"] and rec["bench"].startswith("S") else "unknown")
                rec["search_on_cases.shc.gov.pk"] = f"Seat: {rec['seat'] or '?'} | Case Category: {rec['category']} | Case No: {rec['number']} | Year: {rec['year']}"
                results.append(rec)
        save(results)  # save after every judge so a crash loses nothing

    fam = [r for r in results if r["verdict"].startswith("family")]
    print(f"\nDone. {len(results)} judgments scanned, {len(fam)} confirmed family, "
          f"{sum(r['verdict'].startswith('maybe') for r in results)} maybe.")
    print("Open family_cases.xlsx and filter the 'verdict' column.")


def save(results):
    if not results:
        return
    cols = list(results[0].keys())
    with open("family_cases.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        w.writerows(results)
    try:
        import pandas as pd
        df = pd.DataFrame(results)
        order = {"family (category)": 0, "family": 1, "maybe": 2, "maybe (check text)": 3, "unknown": 4, "not family": 5}
        df = df.sort_values(by="verdict", key=lambda s: s.map(order).fillna(9))
        with pd.ExcelWriter("family_cases.xlsx") as xw:
            df.to_excel(xw, sheet_name="All scanned", index=False)
            df[df["verdict"].str.startswith(("family", "maybe"))].to_excel(xw, sheet_name="Family + maybe", index=False)
            summary = df.groupby(["verdict", "category"] + (["type"] if "type" in df and df["type"].astype(str).str.len().gt(0).any() else [])).size().reset_index(name="count")
            fam = df[df["verdict"].str.startswith("family") & (df.get("type", "") == "full judgment")]
            fam.to_excel(xw, sheet_name="Family full judgments", index=False)
            summary.to_excel(xw, sheet_name="Summary", index=False)
    except Exception as e:
        print(f"  (xlsx not written: {e}; the CSV is fine)")


if __name__ == "__main__":
    main()