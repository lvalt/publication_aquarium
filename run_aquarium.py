#!/usr/bin/env python3
"""
AI-DOC Publication Aquarium — pipeline
======================================
1. Reads the author list (Excel: name, author_id, status [, orcid, gs_user])
2. Fetches papers (Semantic Scholar; fallback sources are added in a later step)
3. Optionally enriches papers via OpenAlex by DOI (venue name, ISSNs, OpenAlex source ID)
4. Matches each venue to SCImago, JUFO, JIF (InCites/JCR) and Google Scholar Metrics
   in a fixed order:  alias file -> ISSN -> OpenAlex source ID -> exact normalised
   name -> unique abbreviation.  No substring / fuzzy matching.
5. Writes papers_with_rankings_<date>.xlsx, match_report.csv, history snapshot, index.html

Usage:
  python run_aquarium.py authors.xlsx --scimago scimago.csv --jufo jufo.csv \
      --jcr jif.csv --gs gs_metrics.csv --output index.html [--dry-run]

  --dry-run         write preview.html + papers_preview.xlsx; no history snapshot
  --aliases FILE    venue alias file (default: venue_aliases.csv if present)
  --no-openalex     skip the OpenAlex DOI enrichment
  --openalex-key K  optional OpenAlex API key (higher daily budget)
  --contact EMAIL   e-mail sent to OpenAlex (polite use)
  --carry-forward   fill venues that got no ranking this run from the previous Excel
                    (off by default; the match report lists what it would fill)
  --min-year N      default 2023

Requirements: pip install requests pandas openpyxl pillow
"""

import csv, html, io, json, re, sys, time, unicodedata, zipfile
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict
from datetime import date
from pathlib import Path

try:
    import requests
except ImportError:
    print("ERROR: pip install requests"); sys.exit(1)
try:
    import pandas as pd
except ImportError:
    print("ERROR: pip install pandas openpyxl"); sys.exit(1)

HISTORY_FILE = "ai_doc_history.json"
SS_API = "https://api.semanticscholar.org/graph/v1"
OA_API = "https://api.openalex.org"
MIN_YEAR = 2023
API_DELAY = 3.0
SOURCES = ("scimago", "jufo", "jif", "gs")

PAPER_COLS = ["author_name", "author_id", "title", "year", "publicationDate", "venue", "venue_issn", "doi",
              "scimago_quartile", "scimago_sjr", "jufo_level", "gs_h5_index", "gs_h5_median",
              "jif", "jif_quartile", "paper_id",
              "source", "venue_issns", "venue_alt_names", "openalex_source_id",
              "match_scimago", "match_jufo", "match_jif", "match_gs", "jufo_id"]


# ═══════════════════════════════════════════════════════
# Text helpers
# ═══════════════════════════════════════════════════════
_LEAD = re.compile(r"^(?:(?:\d{4}|\d+(?:st|nd|rd|th))\s+)+")          # "2025 ", "14th "
_PROC = re.compile(r"^(?:the\s+)?(?:proceedings|proc)\.?\s+(?:of\s+)?(?:the\s+)?")

def clean(s):
    if s is None: return ""
    s = html.unescape(str(s))
    s = unicodedata.normalize("NFKD", s)
    return "".join(c for c in s if not unicodedata.combining(c)).strip()

def norm(s):
    """Normalised venue name for exact comparison (no fuzzy logic)."""
    s = clean(s).lower()
    s = re.sub(r"\([^)]*\)", " ", s)              # parenthetical abbreviations
    s = s.replace("&", " and ")
    s = re.sub(r"[^a-z0-9]+", " ", s).strip()
    prev = None
    while prev != s:                               # strip generic leading words in any order
        prev = s
        s = _PROC.sub("", s)
        s = _LEAD.sub("", s)
        s = re.sub(r"^(?:the|annual)\s+", "", s)
    return re.sub(r"\s+", " ", s).strip()

def abbrevs(s):
    """Abbreviation candidates in a string: '(ICASSP)' parts, or the whole string if it looks like one."""
    s = clean(s)
    out = {a.strip() for a in re.findall(r"\(([^)]{2,20})\)", s)}
    if s and len(s) <= 16 and " " not in s.strip() and any(c.isupper() for c in s[1:]):
        out.add(s.strip())
    return {a.upper() for a in out if re.search(r"[A-Za-z]", a)}

def norm_issn(s):
    s = re.sub(r"[^0-9Xx]", "", str(s or "")).upper()
    return s if len(s) == 8 else None

def split_issns(s):
    return [i for i in (norm_issn(x) for x in re.split(r"[,;\s|]+", str(s or ""))) if i]

def to_float(s):
    try: return float(str(s).replace(",", "."))
    except (TypeError, ValueError): return None


# ═══════════════════════════════════════════════════════
# File readers
# ═══════════════════════════════════════════════════════
def read_table(path):
    """Read a CSV/TSV export robustly: BOM, quoted fields, ; , or tab, footer lines."""
    raw = Path(path).read_bytes()
    try: text = raw.decode("utf-8-sig")
    except UnicodeDecodeError: text = raw.decode("cp1252")
    first = text.split("\n", 1)[0]
    delim = max(["\t", ";", ","], key=first.count)
    rows = list(csv.reader(io.StringIO(text), delimiter=delim))
    header = [h.strip() for h in rows[0]]
    keys = [h.lower() for h in header]
    out = []
    for r in rows[1:]:
        if len(r) < len(header) - 1 or not any(c.strip() for c in r):
            continue                                  # footer / blank lines
        out.append({k: (v.strip() if v else "") for k, v in zip(keys, r)})
    return out


def read_xlsx(path):
    df = pd.read_excel(path).fillna("")
    return df.to_dict("records")


class Index:
    """name/issn/abbr/openalex/id -> set of record ids; matches only when unique."""
    def __init__(self, label):
        self.label, self.rec = label, {}
        self.by = {k: defaultdict(set) for k in ("issn", "name", "abbr", "oa", "id")}

    def add(self, rid, rec, issns=(), names=(), abbrs=(), oa=None, ids=()):
        self.rec[rid] = rec
        for i in issns:
            if norm_issn(i): self.by["issn"][norm_issn(i)].add(rid)
        for n in names:
            if norm(n): self.by["name"][norm(n)].add(rid)
        for a in abbrs:
            if a: self.by["abbr"][a.upper()].add(rid)
        if oa: self.by["oa"][oa].add(rid)
        for i in ids: self.by["id"][str(i)].add(rid)

    def get(self, kind, key):
        hits = self.by[kind].get(key, set())
        if len(hits) == 1: return next(iter(hits)), None
        if len(hits) > 1: return None, f"ambiguous {kind} ({len(hits)} entries)"
        return None, None

    def names(self):
        return self.by["name"]


def build_scimago(path):
    ix = Index("SCImago")
    for r in read_table(path):
        q = r.get("sjr best quartile", "")
        rec = {"quartile": q if q in ("Q1", "Q2", "Q3", "Q4") else None,
               "sjr": to_float(r.get("sjr")), "title": r.get("title", "")}
        ix.add(r.get("sourceid") or r.get("title"), rec, issns=split_issns(r.get("issn")),
               names=[r.get("title")], ids=[r.get("sourceid", "")])
    return ix


def jufo_level(v):
    v = (v or "").strip()
    if v in ("1", "2", "3"): return int(v)
    if v == "0" or v.lower().startswith("other identified"): return 0
    return None                                   # not (yet) rated


def build_jufo(path):
    ix = Index("JUFO")
    for r in read_table(path):
        if (r.get("type_en") or "").lower() == "book publisher":
            continue                              # publisher levels apply to books, not venues
        jid = r.get("jufo_id")
        others = [o for o in re.split(r";", r.get("other_title", "")) if o.strip()]
        names = [r.get("name")] + [o for o in others if len(o.strip()) > 12]
        abbr = {r["abbreviation"].strip().upper()} if r.get("abbreviation") else set()
        abbr |= {o.strip().upper() for o in others if len(o.strip()) <= 12 and " " not in o.strip()}
        oa = (r.get("openalex_id") or "").rsplit("/", 1)[-1] or None
        rec = {"level": jufo_level(r.get("level")), "jufo_id": jid, "title": r.get("name", ""),
               "raw_level": r.get("level", "")}
        ix.add(jid, rec, issns=[r.get("issnl"), r.get("issn1"), r.get("issn2")],
               names=names, abbrs=abbr, oa=oa, ids=[jid])
    return ix


def build_jif(path):
    ix = Index("JIF")
    for n, r in enumerate(read_table(path)):
        jif = to_float(r.get("journal impact factor"))
        q = r.get("jif quartile", "")
        if jif is None and q not in ("Q1", "Q2", "Q3", "Q4"):
            continue
        rec = {"jif": round(jif, 1) if jif is not None else None,
               "jif_quartile": q if q in ("Q1", "Q2", "Q3", "Q4") else None, "title": r.get("name", "")}
        issns = [r.get("issn"), r.get("eissn")]
        ix.add(f"jif{n}", rec, issns=issns, names=[r.get("name")],
               ids=[i for i in (norm_issn(x) for x in issns) if i])
    return ix


def build_gs(path):
    ix = Index("GS")
    for n, r in enumerate(read_table(path)):
        h5 = to_float(r.get("h5_index"))
        if h5 is None: continue
        med = to_float(r.get("h5_median"))
        name = r.get("venue_name", "")
        rec = {"h5_index": int(h5), "h5_median": int(med) if med is not None else None, "title": name}
        rid = r.get("gs_venue_id") or f"gs{n}"
        ix.add(rid, rec, names=[name], abbrs=abbrevs(name), ids=[rid, name])
    return ix


def load_aliases(path):
    """venue_aliases.csv: venue, source (scimago|jufo|jif|gs), target, note.
    target = SCImago Sourceid | Jufo_ID | ISSN | GS venue id or exact GS name | 'none'."""
    al = {}
    if not path or not Path(path).exists(): return al
    for r in read_table(path):
        v, s, t = norm(r.get("venue")), (r.get("source") or "").lower().strip(), (r.get("target") or "").strip()
        if v and s in SOURCES and t:
            al[(v, s)] = t
    return al


# ═══════════════════════════════════════════════════════
# Matching
# ═══════════════════════════════════════════════════════
def match_venue(ix, src, aliases, names, issns, oa_id):
    """Return (record or None, method string)."""
    if ix is None: return None, "not loaded"
    for n in names:                                         # 1. alias
        t = aliases.get((norm(n), src))
        if t:
            if t.lower() == "none": return None, "alias: none"
            key = norm_issn(t) or t
            rid, _ = ix.get("id", key)
            if rid is None and src == "gs": rid, _ = ix.get("name", norm(t))
            if rid is not None: return ix.rec[rid], f"alias -> {t}"
            return None, f"alias target not found: {t}"
    notes = []
    for i in issns:                                         # 2. ISSN
        rid, why = ix.get("issn", i)
        if rid is not None: return ix.rec[rid], f"issn {i}"
        if why: notes.append(why)
    if oa_id:                                               # 3. OpenAlex source
        rid, why = ix.get("oa", oa_id)
        if rid is not None: return ix.rec[rid], f"openalex {oa_id}"
    for n in names:                                         # 4. exact normalised name
        rid, why = ix.get("name", norm(n))
        if rid is not None: return ix.rec[rid], "name"
        if why: notes.append(why)
    for a in sorted({a for n in names for a in abbrevs(n)}):  # 5. unique abbreviation
        rid, why = ix.get("abbr", a)
        if rid is not None:
            if shares_word(ix.rec[rid].get("title", ""), names):
                return ix.rec[rid], f"abbr {a}"
            notes.append(f"abbr {a} -> '{ix.rec[rid].get('title', '')[:50]}' rejected (no word in common)")
        if why: notes.append(why)
    return None, "; ".join(dict.fromkeys(notes)) or "no match"


_GENERIC = {"international", "conference", "proceedings", "journal", "annual", "workshop", "symposium",
            "transactions", "letters", "society", "association", "advances", "research", "science", "sciences",
            "studies", "review", "applications", "systems", "technology", "engineering", "ieee", "acm", "ieee/cvf"}

def _words(s):
    return {w for w in norm(s).split() if len(w) >= 4 and w not in _GENERIC}

def shares_word(title, names):
    """Guard for abbreviation matches: the target's full name must share a real word with the venue
    (prefix match handles abbreviated forms like 'Comput' / 'Computing')."""
    t = _words(title)
    for n in names:
        for w in _words(n):
            if any(x.startswith(w[:5]) or w.startswith(x[:5]) for x in t):
                return True
    return False


def suggest(ix, names, k=2):
    """Closest names by word overlap - for the report only, never used as a match."""
    if ix is None or not names: return ""
    q = set(norm(names[0]).split()) - {"of", "on", "the", "and", "in", "for", "international", "conference", "journal"}
    if not q: return ""
    scored = []
    for nm, ids in ix.names().items():
        w = set(nm.split())
        s = len(q & w) / len(q | w)
        if s >= 0.35: scored.append((s, nm, ids))
    scored.sort(key=lambda x: -x[0])
    out = []
    for s, nm, ids in scored[:k]:
        rid = next(iter(ids)); rec = ix.rec[rid]
        out.append(f"{rec.get('title') or nm} [{rid}] ({s:.2f})")
    return " | ".join(out)


# ═══════════════════════════════════════════════════════
# Semantic Scholar
# ═══════════════════════════════════════════════════════
def http_get(url, params=None, retries=3):
    for attempt in range(retries):
        try:
            r = requests.get(url, params=params, timeout=30)
            if r.status_code == 429:
                wait = int(r.headers.get("Retry-After", 10))
                print(f"    ⏳ rate limited, waiting {wait}s"); time.sleep(wait); continue
            if r.status_code != 200:
                print(f"    ⚠ HTTP {r.status_code} {url}"); return None
            return r.json()
        except Exception as e:
            if attempt == retries - 1: print(f"    ⚠ {e}")
            else: time.sleep(3)
    return None


def fetch_s2_papers(author_id, min_year):
    data = http_get(f"{SS_API}/author/{author_id}/papers",
                    {"fields": "title,year,publicationDate,venue,externalIds,publicationVenue", "limit": "500"})
    out = []
    for p in (data or {}).get("data", []):
        year = p.get("year")
        if year is None and p.get("publicationDate"):
            m = re.match(r"(\d{4})", str(p["publicationDate"])); year = int(m.group(1)) if m else None
        if year is None or year < min_year: continue
        pv = p.get("publicationVenue") or {}
        issns = [pv.get("issn")] + list(pv.get("alternate_issns") or [])
        out.append({
            "title": p.get("title", ""), "year": year, "publicationDate": p.get("publicationDate") or "",
            "venue": clean(pv.get("name") or p.get("venue") or ""),
            "venue_alt": [clean(x) for x in (pv.get("alternate_names") or []) + [p.get("venue") or ""] if x],
            "issns": [i for i in (norm_issn(x) for x in issns) if i],
            "doi": ((p.get("externalIds") or {}).get("DOI") or "").lower(),
            "paper_id": p.get("paperId", ""), "source": "semanticscholar",
        })
    return out


# ═══════════════════════════════════════════════════════
# Fallback sources for authors without a Semantic Scholar ID
# ═══════════════════════════════════════════════════════
def norm_orcid(s):
    m = re.search(r"(\d{4}-\d{4}-\d{4}-\d{3}[\dXx])", str(s or ""))
    return m.group(1).upper() if m else None


def _oa_work_to_paper(w):
    loc = w.get("primary_location") or {}
    src = loc.get("source") or {}
    issns = [norm_issn(i) for i in (src.get("issn") or []) + [src.get("issn_l")]]
    return {"title": clean(w.get("display_name") or w.get("title") or ""), "year": w.get("publication_year"),
            "publicationDate": w.get("publication_date") or "",
            "venue": clean(src.get("display_name") or ""), "venue_alt": [],
            "issns": list(dict.fromkeys(i for i in issns if i)),
            "doi": (w.get("doi") or "").lower().replace("https://doi.org/", ""),
            "paper_id": (w.get("id") or "").rsplit("/", 1)[-1],
            "openalex_source_id": (src.get("id") or "").rsplit("/", 1)[-1],
            "openalex_type": w.get("type") or ""}


def fetch_openalex_by_orcid(orcid, min_year, key=None, contact=None):
    """Works for an ORCID iD from OpenAlex. Returns None if OpenAlex could not be reached."""
    out, cursor = [], "*"
    while cursor:
        params = {"filter": f"author.orcid:{orcid},publication_year:>{min_year - 1}", "per-page": "200",
                  "cursor": cursor, "select": "id,doi,display_name,publication_year,publication_date,primary_location,type"}
        if key: params["api_key"] = key
        if contact: params["mailto"] = contact
        data = http_get(f"{OA_API}/works", params)
        if data is None: return None if not out else out
        for w in data.get("results", []):
            p = _oa_work_to_paper(w)
            if p["year"] and p["year"] >= min_year:
                p["source"] = "orcid+openalex"; out.append(p)
        cursor = (data.get("meta") or {}).get("next_cursor")
        if not data.get("results"): break
    return out


def fetch_orcid_api(orcid, min_year):
    """Public ORCID API (no login). Fewer details: journal title and DOI, rarely ISSN."""
    try:
        r = requests.get(f"https://pub.orcid.org/v3.0/{orcid}/works",
                         headers={"Accept": "application/json"}, timeout=30)
        if r.status_code != 200:
            print(f"    ⚠ ORCID HTTP {r.status_code}"); return []
        data = r.json()
    except Exception as e:
        print(f"    ⚠ ORCID: {e}"); return []
    out = []
    for g in data.get("group", []):
        s = (g.get("work-summary") or [{}])[0]
        year = (((s.get("publication-date") or {}).get("year") or {}).get("value"))
        if not year or int(year) < min_year: continue
        doi = ""
        for e in ((s.get("external-ids") or {}).get("external-id") or []):
            if (e.get("external-id-type") or "").lower() == "doi":
                doi = (e.get("external-id-value") or "").lower(); break
        out.append({"title": clean(((s.get("title") or {}).get("title") or {}).get("value", "")),
                    "year": int(year), "publicationDate": str(year),
                    "venue": clean((s.get("journal-title") or {}).get("value", "")), "venue_alt": [],
                    "issns": [], "doi": doi.replace("https://doi.org/", ""),
                    "paper_id": f"orcid:{s.get('put-code', '')}", "source": "orcid"})
    return out


GS_USER_RE = re.compile(r"^[A-Za-z0-9_-]{12}$")

def gs_venue_from_line(s):
    """'Proceedings of the AAAI Conference on Artificial Intelligence 38 (16), 1-9, 2024' -> name part."""
    s = clean(s)
    if s.lower().startswith("arxiv"): return "arXiv"
    s = re.split(r",\s*\d|\s+\d+(?![A-Za-z])", s, maxsplit=1)[0]
    return s.strip(" ,.")


class _GSProfile(__import__("html.parser").parser.HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.rows, self.row, self.field, self.gray_n, self.buf, self.skip = [], None, None, 0, "", 0
    def handle_starttag(self, tag, attrs):
        a = dict(attrs); cls = a.get("class", "") or ""
        if tag == "tr" and "gsc_a_tr" in cls:
            self.row, self.gray_n = {"title": "", "venue": "", "year": ""}, 0
        if self.row is None: return
        if tag == "span" and "gs_oph" in cls: self.skip += 1       # duplicated ", 2024"
        elif tag == "a" and "gsc_a_at" in cls: self.field, self.buf = "title", ""
        elif tag == "div" and "gs_gray" in cls:
            self.gray_n += 1
            if self.gray_n == 2: self.field, self.buf = "venue", ""
        elif tag == "span" and "gsc_a_h" in cls: self.field, self.buf = "year", ""
    def handle_endtag(self, tag):
        if self.row is None: return
        if tag == "span" and self.skip: self.skip -= 1; return
        if self.field and tag in ("a", "div", "span"):
            self.row[self.field] = self.buf.strip(); self.field = None
        if tag == "tr":
            if self.row.get("title"): self.rows.append(self.row)
            self.row = None
    def handle_data(self, d):
        if self.field and not self.skip: self.buf += d


def parse_gs_profile(text):
    p = _GSProfile(); p.feed(text); return p.rows


def fetch_gs_profile(user, min_year, delay=20.0):
    """First page of a Google Scholar profile, newest first (robots.txt allows /citations?user=...)."""
    if not GS_USER_RE.match(user or ""):
        print(f"    ⚠ '{user}' doesn't look like a Google Scholar user ID (12 characters)"); return []
    try:
        import scrape_gs_metrics as g
    except ImportError:
        print("    ⚠ scrape_gs_metrics.py not found next to this script - Google Scholar skipped"); return []
    path = f"/citations?user={user}&hl=en&sortby=pubdate&pagesize=100"
    s = requests.Session()
    s.headers.update({"User-Agent": "AI-DOC-publication-aquarium/1.0 (Google Scholar profile reader)",
                      "Accept-Language": "en"})
    global _GS_RULES
    if "_GS_RULES" not in globals():
        _GS_RULES = g.parse_robots(s.get(g.BASE + "/robots.txt", timeout=30).text)
    if not path.startswith("/citations?user=") or not g.robots_allows(_GS_RULES, path):
        print("    ⚠ robots.txt does not allow this profile URL - skipped"); return []
    time.sleep(delay)
    r = s.get(g.BASE + path, timeout=30)
    if g.looks_blocked(r.status_code, r.url, r.text):
        print("    ⚠ Google Scholar returned a block/CAPTCHA - profile skipped this run"); return []
    out = []
    for row in parse_gs_profile(r.text):
        y = int(row["year"]) if row["year"].isdigit() else None
        if y is None or y < min_year: continue
        out.append({"title": clean(row["title"]), "year": y, "publicationDate": str(y),
                    "venue": gs_venue_from_line(row["venue"]), "venue_alt": [clean(row["venue"])],
                    "issns": [], "doi": "", "paper_id": f"gs:{user}:{norm(row['title'])[:60]}",
                    "source": "googlescholar"})
    return out


def openalex_by_title(papers, key=None, contact=None):
    """For papers without a DOI (ORCID / Google Scholar): exact normalised-title match in OpenAlex."""
    n = 0
    for p in papers:
        if p.get("doi") or not p.get("title"): continue
        q = re.sub(r"[^\w\s]", " ", p["title"])[:200]
        params = {"filter": f"title.search:{q},publication_year:{p['year']}", "per-page": "5",
                  "select": "id,doi,display_name,publication_year,publication_date,primary_location,type"}
        if key: params["api_key"] = key
        if contact: params["mailto"] = contact
        data = http_get(f"{OA_API}/works", params)
        if data is None: break
        for w in data.get("results", []):
            if norm(w.get("display_name")) == norm(p["title"]):
                hit = _oa_work_to_paper(w)
                p["doi"] = hit["doi"]; p["openalex_source_id"] = hit["openalex_source_id"]
                if hit["venue"]:
                    if p["venue"] and p["venue"] != hit["venue"]: p["venue_alt"].insert(0, p["venue"])
                    p["venue"] = hit["venue"]
                p["issns"] = list(dict.fromkeys(p["issns"] + hit["issns"])); n += 1; break
        time.sleep(0.2)
    return n


def dedupe(papers):
    seen, out = set(), []
    for p in papers:
        k = p.get("doi") or norm(p.get("title"))
        if k and k in seen: continue
        seen.add(k); out.append(p)
    return out


# ═══════════════════════════════════════════════════════
# OpenAlex enrichment (by DOI, 50 per request)
# ═══════════════════════════════════════════════════════
PREPRINT = re.compile(r"arxiv|biorxiv|medrxiv|chemrxiv|ssrn|social science research network|research square|"
                      r"preprints\.org|techrxiv|psyarxiv|osf preprints|zenodo|lirias", re.I)

def openalex_enrich(papers, key=None, contact=None):
    dois = sorted({p["doi"] for p in papers if p.get("doi")})
    if not dois: return 0
    found = {}
    for i in range(0, len(dois), 50):
        chunk = dois[i:i + 50]
        params = {"filter": "doi:" + "|".join(chunk), "per-page": "50",
                  "select": "doi,primary_location"}
        if key: params["api_key"] = key
        if contact: params["mailto"] = contact
        data = http_get(f"{OA_API}/works", params)
        if data is None:
            print("    ⚠ OpenAlex unavailable - continuing without enrichment"); break
        for w in data.get("results", []):
            src = ((w.get("primary_location") or {}).get("source") or {})
            d = (w.get("doi") or "").lower().replace("https://doi.org/", "")
            if d and src: found[d] = src
        time.sleep(0.2)
    n = 0
    for p in papers:
        src = found.get(p.get("doi"))
        if not src: continue
        n += 1
        p["openalex_source_id"] = (src.get("id") or "").rsplit("/", 1)[-1]
        name = clean(src.get("display_name"))
        if name:
            if not p["venue"]: p["venue"] = name
            elif PREPRINT.search(p["venue"]) and (src.get("issn") or src.get("issn_l")) and not PREPRINT.search(name):
                # Semantic Scholar still shows the preprint server, but the DOI belongs to a published journal
                p["venue_alt"].insert(0, p["venue"]); p["venue"] = name
            elif name not in p["venue_alt"]: p["venue_alt"].append(name)
        for i in (src.get("issn") or []) + [src.get("issn_l")]:
            i = norm_issn(i)
            if i and i not in p["issns"]: p["issns"].append(i)
    return n


# ═══════════════════════════════════════════════════════
# History snapshot (unchanged logic)
# ═══════════════════════════════════════════════════════
def build_snapshot(papers, snapshot_date):
    authors = {p["author_name"] for p in papers if p.get("author_name")}
    qc, jc, jifq = Counter(), Counter(), Counter()
    h5 = {"200+": 0, "100-199": 0, "50-99": 0, "1-49": 0, "N/A": 0}
    for p in papers:
        q = p.get("scimago_quartile", ""); qc[q if q in ("Q1", "Q2", "Q3", "Q4") else "N/A"] += 1
        j = str(p.get("jufo_level", "")); jc[j if j in ("0", "1", "2", "3") else "N/A"] += 1
        jq = p.get("jif_quartile", ""); jifq[jq if jq in ("Q1", "Q2", "Q3", "Q4") else "N/A"] += 1
        v = p.get("gs_h5_index", "")
        if v in ("", None): h5["N/A"] += 1
        elif int(v) >= 200: h5["200+"] += 1
        elif int(v) >= 100: h5["100-199"] += 1
        elif int(v) >= 50: h5["50-99"] += 1
        else: h5["1-49"] += 1
    sjrs = sorted(float(p["scimago_sjr"]) for p in papers if p.get("scimago_sjr") not in ("", None))
    jifs = [float(p["jif"]) for p in papers if p.get("jif") not in ("", None)]
    m = len(sjrs) // 2
    return {
        "date": snapshot_date, "total": len(papers), "authors": len(authors),
        "q1": qc["Q1"], "q2": qc["Q2"], "q3": qc["Q3"], "q4": qc["Q4"], "q_na": qc["N/A"],
        "jufo3": jc["3"], "jufo2": jc["2"], "jufo1": jc["1"], "jufo0": jc["0"], "jufo_na": jc["N/A"],
        "h5_200plus": h5["200+"], "h5_100_199": h5["100-199"], "h5_50_99": h5["50-99"],
        "h5_1_49": h5["1-49"], "h5_na": h5["N/A"],
        "jif_q1": jifq["Q1"], "jif_q2": jifq["Q2"], "jif_q3": jifq["Q3"], "jif_q4": jifq["Q4"], "jif_na": jifq["N/A"],
        "avg_sjr": round(sum(sjrs) / len(sjrs), 2) if sjrs else None,
        "med_sjr": round(sjrs[m] if len(sjrs) % 2 else (sjrs[m - 1] + sjrs[m]) / 2, 2) if sjrs else None,
        "avg_jif": round(sum(jifs) / len(jifs), 2) if jifs else None,
    }


# ═══════════════════════════════════════════════════════
# Previous run (comparison + optional carry-forward)
# ═══════════════════════════════════════════════════════
FIELDS = {"scimago": ("scimago_quartile", "scimago_sjr"), "jufo": ("jufo_level",),
          "jif": ("jif", "jif_quartile"), "gs": ("gs_h5_index", "gs_h5_median")}

def load_previous():
    files = sorted(Path(".").glob("papers_with_rankings_*.xlsx"), reverse=True)
    if not files: return None, {}
    prev = {}
    for p in read_xlsx(files[0]):
        v = norm(p.get("venue"))
        if v and v not in prev: prev[v] = p
    return files[0].name, prev


def blank(v):
    return v is None or (isinstance(v, float) and v != v) or str(v).strip() in ("", "nan", "NaN")


# ═══════════════════════════════════════════════════════
# Main
# ═══════════════════════════════════════════════════════
def load_authors(path):
    """Included: status CONFIRMED (Semantic Scholar ID), or any row with an ORCID iD / Google Scholar user ID.
    Optional columns: orcid, gs_user."""
    authors = []
    for r in read_xlsx(path):
        status = str(r.get("status", "")).upper().strip()
        aid = str(r.get("author_id", "")).replace(".0", "").strip()
        orcid = norm_orcid(r.get("orcid", ""))
        gs = str(r.get("gs_user", "") or "").strip()
        m = re.search(r"user=([A-Za-z0-9_-]{12})", gs)          # accept a full profile URL too
        gs = m.group(1) if m else gs
        if status != "CONFIRMED": aid = ""                       # unconfirmed S2 IDs are not used
        if not (aid or orcid or gs): continue
        authors.append({"name": r.get("name", ""), "author_id": aid, "orcid": orcid or "", "gs_user": gs})
    return authors


def run(author_xlsx, scimago_csv=None, jufo_csv=None, gs_csv=None, jcr_csv=None,
        template_path=None, output_path="index.html", history_path=HISTORY_FILE,
        min_year=MIN_YEAR, aliases_path="venue_aliases.csv", openalex=True,
        openalex_key=None, contact=None, carry_forward=False, dry_run=False,
        fetch=None, fetchers=None, gs_delay=20.0):
    today = date.today().isoformat()
    fetch = fetch or fetch_s2_papers
    fx = {"orcid_oa": lambda o: fetch_openalex_by_orcid(o, min_year, openalex_key, contact),
          "orcid_api": lambda o: fetch_orcid_api(o, min_year),
          "gs": lambda u: fetch_gs_profile(u, min_year, gs_delay)}
    fx.update(fetchers or {})

    print(f"[1/7] Authors from {author_xlsx}")
    authors = load_authors(author_xlsx)
    n_s2 = sum(1 for a in authors if a["author_id"])
    print(f"  {len(authors)} authors: {n_s2} via Semantic Scholar, {len(authors) - n_s2} via ORCID / Google Scholar")

    print("[2/7] Ranking files")
    ix = {"scimago": build_scimago(scimago_csv) if scimago_csv else None,
          "jufo": build_jufo(jufo_csv) if jufo_csv else None,
          "jif": build_jif(jcr_csv) if jcr_csv else None,
          "gs": build_gs(gs_csv) if gs_csv else None}
    for s, i in ix.items():
        print(f"  {s:8s} " + (f"{len(i.rec)} entries" if i else "not loaded"))
    aliases = load_aliases(aliases_path)
    print(f"  aliases  {len(aliases)} entries" + ("" if aliases else " (no venue_aliases.csv)"))

    print(f"[3/7] Fetching papers (min year {min_year})")
    papers, by_source = [], Counter()
    for n, a in enumerate(authors, 1):
        got, via = [], "-"
        if a["author_id"]:
            got, via = fetch(a["author_id"], min_year), "Semantic Scholar"
            if fetch is fetch_s2_papers: time.sleep(API_DELAY)
        else:
            if a["orcid"]:
                got, via = fx["orcid_oa"](a["orcid"]), "ORCID via OpenAlex"
                if not got:                                  # None = OpenAlex unreachable, [] = nothing found
                    got, via = fx["orcid_api"](a["orcid"]), "ORCID public API"
            if not got and a["gs_user"]:
                got, via = fx["gs"](a["gs_user"]), "Google Scholar profile"
        got = dedupe(got or [])
        print(f"  [{n}/{len(authors)}] {a['name']}: {len(got)} paper(s) via {via}")
        by_source[via] += len(got)
        for p in got:
            p.setdefault("venue_alt", []); p.setdefault("issns", [])
            p.update(author_name=a["name"], author_id=a["author_id"] or a["orcid"] or a["gs_user"]); papers.append(p)
    print("  papers by source: " + ", ".join(f"{k}={v}" for k, v in by_source.items()))
    no_doi = [p for p in papers if not p.get("doi") and p.get("source") in ("orcid", "googlescholar")]
    if openalex and no_doi:
        print(f"  OpenAlex title lookup for {len(no_doi)} papers without DOI: "
              f"{openalex_by_title(no_doi, openalex_key, contact)} found")

    if openalex:
        print("[4/7] OpenAlex enrichment by DOI")
        print(f"  enriched {openalex_enrich(papers, openalex_key, contact)} of {len(papers)} papers")
    else:
        print("[4/7] OpenAlex enrichment skipped")

    print("[5/7] Matching venues")
    prev_name, prev = load_previous()
    cache, report = {}, {}
    for p in papers:
        names = [n for n in [p["venue"]] + p.get("venue_alt", []) if n]
        key = (norm(p["venue"]) or p.get("doi") or p["title"], tuple(sorted(p["issns"])))
        if key not in cache:
            res = {s: match_venue(ix[s], s, aliases, names, p["issns"], p.get("openalex_source_id")) for s in SOURCES}
            cache[key] = res
        res = cache[key]
        sci, jufo, jif, gs = (res[s][0] or {} for s in SOURCES)
        p.update(
            venue_issn=p["issns"][0][:4] + "-" + p["issns"][0][4:] if p["issns"] else "",
            venue_issns=", ".join(p["issns"]), venue_alt_names="; ".join(p.get("venue_alt", [])),
            openalex_source_id=p.get("openalex_source_id", ""),
            scimago_quartile=sci.get("quartile") or "", scimago_sjr=sci.get("sjr") if sci.get("sjr") is not None else "",
            jufo_level=jufo.get("level") if jufo.get("level") is not None else "", jufo_id=jufo.get("jufo_id", ""),
            jif=jif.get("jif") if jif.get("jif") is not None else "", jif_quartile=jif.get("jif_quartile") or "",
            gs_h5_index=gs.get("h5_index") if gs.get("h5_index") is not None else "",
            gs_h5_median=gs.get("h5_median") if gs.get("h5_median") is not None else "",
            **{f"match_{s}": res[s][1] for s in SOURCES})
        # report row per venue
        r = report.setdefault(key, {"venue": p["venue"] or "(no venue)", "papers": 0, "issns": p["venue_issns"],
                                   "alt_names": p["venue_alt_names"], "openalex_source": p["openalex_source_id"],
                                   **{f"{s}_value": "" for s in SOURCES}, **{f"{s}_how": res[s][1] for s in SOURCES},
                                   **{f"{s}_suggest": "" for s in SOURCES}, "changed_vs_previous": ""})
        r["papers"] += 1
        r["scimago_value"] = p["scimago_quartile"] or (f"SJR {p['scimago_sjr']} (no quartile)" if p["scimago_sjr"] != "" else "")
        r["jufo_value"] = p["jufo_level"] if p["jufo_level"] != "" else ("in JUFO, no level" if p["jufo_id"] else "")
        r["jif_value"] = p["jif"]; r["gs_value"] = p["gs_h5_index"]

    # comparison with previous run / optional carry-forward
    carried = Counter()
    for key, r in report.items():
        old = prev.get(norm(r["venue"]))
        if not old: continue
        diffs = []
        for s, flds in FIELDS.items():
            o, nw = old.get(flds[0]), r[f"{s}_value"]
            if s == "scimago":
                o = o if str(o) in ("Q1", "Q2", "Q3", "Q4") else ""
                nw = nw if str(nw) in ("Q1", "Q2", "Q3", "Q4") else ""
            if s == "jufo" and not str(nw).replace(".0", "").isdigit(): nw = ""
            if blank(o) and blank(nw): continue
            if blank(nw) and not blank(o):
                diffs.append(f"{s}: {o} -> none")
            elif not blank(o):
                a, b = to_float(o), to_float(nw)
                same = (a == b) if (a is not None and b is not None) else str(o).strip() == str(nw).strip()
                if not same: diffs.append(f"{s}: {o} -> {nw}")
        r["changed_vs_previous"] = "; ".join(diffs)
    if carry_forward and prev:
        for p in papers:
            old = prev.get(norm(p["venue"]))
            if not old: continue
            for s, flds in FIELDS.items():
                if all(blank(p.get(f)) for f in flds) and not all(blank(old.get(f)) for f in flds):
                    for f in flds: p[f] = "" if blank(old.get(f)) else old.get(f)
                    p[f"match_{s}"] = f"carried forward from {prev_name}"; carried[s] += 1

    for (key, r) in report.items():
        names = [r["venue"]] + [n for n in r["alt_names"].split("; ") if n]
        for s in SOURCES:
            if r[f"{s}_value"] in ("", None) and ix[s] is not None and r["venue"] != "(no venue)":
                r[f"{s}_suggest"] = suggest(ix[s], names)

    total = len(papers)
    print(f"  {total} papers, {len(report)} distinct venues")
    for s, col in (("scimago", "scimago_quartile"), ("jufo", "jufo_level"), ("jif", "jif_quartile"), ("gs", "gs_h5_index")):
        how = Counter(p[f"match_{s}"].split(" ")[0] for p in papers
                      if p[f"match_{s}"].split(" ")[0] in ("alias", "issn", "openalex", "name", "abbr", "carried"))
        print(f"  {s:8s} {sum(how.values()):4d}/{total} papers matched  " + ", ".join(f"{k}={v}" for k, v in how.most_common()))
    if carried: print(f"  carried forward: {dict(carried)}")

    rep_path = "match_report.csv"
    cols = ["venue", "papers", "issns", "alt_names", "openalex_source"] + \
           [f"{s}_{x}" for s in SOURCES for x in ("value", "how", "suggest")] + ["changed_vs_previous"]
    with open(rep_path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=cols, extrasaction="ignore"); w.writeheader()
        for r in sorted(report.values(), key=lambda r: (-r["papers"], r["venue"])): w.writerow(r)
    print(f"  match report: {rep_path}" + (f" (compared with {prev_name})" if prev_name else ""))

    print("[6/7] Excel + history")
    excel_path = "papers_preview.xlsx" if dry_run else f"papers_with_rankings_{today}.xlsx"
    df = pd.DataFrame(papers)
    df = df[[c for c in PAPER_COLS if c in df.columns]]
    df.to_excel(excel_path, index=False)
    print(f"  {excel_path}")
    hp = Path(history_path)
    history = json.loads(hp.read_text()) if hp.exists() else []
    if not dry_run:
        history = [h for h in history if h["date"] != today] + [build_snapshot(papers, today)]
        history.sort(key=lambda h: h["date"])
        hp.write_text(json.dumps(history, indent=2))
        print(f"  history: {len(history)} snapshots")
    else:
        print("  dry run: history not changed")

    print("[7/7] HTML")
    bake_html(papers, authors, history, author_xlsx, gs_csv, template_path,
              "preview.html" if dry_run else output_path, today)


def bake_html(papers, authors, history, author_xlsx, gs_csv, template_path, output_path, today):
    if template_path is None:
        for c in [Path("observatory_aquarium.html"), Path(__file__).parent / "observatory_aquarium.html"]:
            if c.exists(): template_path = str(c); break
    if not template_path:
        print("  ⚠ observatory_aquarium.html not found - HTML skipped"); return
    page = Path(template_path).read_text(encoding="utf-8")
    if "/*LOGO_BASE64*/" in page:
        import base64
        for lp in [Path("ai-doc-no-text.png"), Path(__file__).parent / "ai-doc-no-text.png"]:
            if not lp.exists(): continue
            try:
                from PIL import Image
                img = Image.open(lp); img.thumbnail((300, 120))
                buf = io.BytesIO(); img.save(buf, format="PNG", optimize=True); data = buf.getvalue()
            except ImportError:
                print("  ⚠ Pillow not installed - logo embedded at full size (pip install pillow)")
                data = lp.read_bytes()
            page = page.replace("/*LOGO_BASE64*/", base64.b64encode(data).decode()); break
    keep = [c for c in PAPER_COLS if c not in ("venue_alt_names",)]
    slim = [{k: p.get(k, "") for k in keep} for p in papers]
    gs_data = []
    embed = (
        f"const EMBEDDED_PAPERS={json.dumps(slim, separators=(',', ':'), default=str)};\n"
        f"const EMBEDDED_AUTHORS={json.dumps([{'name': a['name'], 'author_id': a['author_id']} for a in authors], separators=(',', ':'))};\n"
        f"const EMBEDDED_GS={json.dumps(gs_data)};\n"
        f"const EMBEDDED_HISTORY={json.dumps(history, separators=(',', ':'))};\n"
        f'const EMBEDDED_DATE="{today}";\n'
        f'const EMBEDDED_SOURCE="{Path(author_xlsx).name}";\n')
    page = page.replace("/*EMBEDDED_DATA_MARKER*/", embed)
    Path(output_path).write_text(page, encoding="utf-8")
    print(f"  wrote {output_path} ({Path(output_path).stat().st_size // 1024} KB)")


# ═══════════════════════════════════════════════════════
# CLI
# ═══════════════════════════════════════════════════════
if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("authors")
    ap.add_argument("--scimago"); ap.add_argument("--jufo"); ap.add_argument("--jcr"); ap.add_argument("--gs")
    ap.add_argument("--template"); ap.add_argument("--output", default="index.html")
    ap.add_argument("--history", default=HISTORY_FILE)
    ap.add_argument("--min-year", type=int, default=MIN_YEAR)
    ap.add_argument("--delay", type=float, default=API_DELAY)
    ap.add_argument("--aliases", default="venue_aliases.csv")
    ap.add_argument("--no-openalex", action="store_true")
    ap.add_argument("--openalex-key"); ap.add_argument("--contact")
    ap.add_argument("--carry-forward", action="store_true")
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()
    API_DELAY = a.delay
    run(a.authors, a.scimago, a.jufo, a.gs, a.jcr, a.template, a.output, a.history, a.min_year,
        a.aliases, not a.no_openalex, a.openalex_key, a.contact, a.carry_forward, a.dry_run)
