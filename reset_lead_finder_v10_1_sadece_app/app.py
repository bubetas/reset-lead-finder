from __future__ import annotations

import io
import json
import hmac
import os
import re
import socket
import time
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable
from urllib.parse import quote_plus, urljoin, urlparse

import dns.resolver
import pandas as pd
import requests
import streamlit as st

APP_DIR = Path(__file__).resolve().parent
PILOT_FILE = APP_DIR / "data" / "reset_fmcg_pilot.csv"
SAMPLE_EMAIL_FILE = APP_DIR / "data" / "oncelikli_25_ornek.xlsx"
SAMPLE_VERIFIED_FILE = APP_DIR / "data" / "dogrulanmis_5_ornek.xlsx"
TAVILY_ENDPOINT = "https://api.tavily.com/search"
ABSTRACT_ENDPOINT = "https://emailreputation.abstractapi.com/v1/"
LUSHA_ENDPOINT = "https://api.lusha.com/v3/contacts/search-and-enrich"

ROLE_GROUPS: dict[str, dict[str, list[str]]] = {
    "Pazarlama & Marka": {
        "quick": ["Marketing Director", "Brand Manager"],
        "deep": ["Marketing Director", "Head of Marketing", "Marketing Manager", "Brand Manager", "Pazarlama Müdürü", "Marka Müdürü"],
    },
    "Satın Alma": {
        "quick": ["Procurement Manager", "Indirect Procurement"],
        "deep": ["Procurement Director", "Procurement Manager", "Purchasing Manager", "Indirect Procurement", "Marketing Procurement", "Satın Alma Müdürü", "Dolaylı Satın Alma"],
    },
    "Ticari Pazarlama & Shopper": {
        "quick": ["Trade Marketing", "Shopper Marketing"],
        "deep": ["Trade Marketing Director", "Trade Marketing Manager", "Shopper Marketing", "Customer Marketing", "Ticari Pazarlama"],
    },
    "Kurumsal İletişim & Etkinlik": {
        "quick": ["Corporate Communications Manager", "Brand Experience"],
        "deep": ["Corporate Communications Director", "Corporate Communications Manager", "Event Manager", "Sponsorship Manager", "Brand Experience", "Kurumsal İletişim Müdürü"],
    },
}

ROLE_KEYWORDS = {
    "marketing", "brand", "pazarlama", "marka", "procurement", "purchasing",
    "satın alma", "indirect", "sourcing", "trade marketing", "shopper",
    "customer marketing", "corporate communications", "kurumsal iletişim",
    "event", "sponsorship", "brand experience",
}
FORMER_MARKERS = {"former", "previous", "ex-", "ex ", "eski", "önceki", "until ", "was ", "formerly", "past role"}
BLOCKED_DOMAINS = {
    "linkedin.com", "facebook.com", "instagram.com", "x.com", "twitter.com", "youtube.com",
    "wikipedia.org", "crunchbase.com", "bloomberg.com", "glassdoor.com", "indeed.com",
    "rocketreach.co", "contactout.com", "apollo.io", "lusha.com", "zoominfo.com",
}
GENERIC_LOCAL_PARTS = {
    "info", "contact", "iletisim", "hello", "merhaba", "support", "destek", "sales", "satis",
    "marketing", "pazarlama", "hr", "ik", "career", "kariyer", "privacy", "kvkk", "webmaster",
    "admin", "office", "reception", "press", "basin", "media", "musteri", "customer",
    "investor", "investors", "relations", "investorrelations", "finance", "ir",
    "corporate", "communications", "communication", "compliance", "legal",
    "procurement", "purchasing", "accounts", "billing", "security", "ethics",
}
KNOWN_DOMAIN_HINTS = {
    "unilever turkiye": "unilever.com",
    "procter gamble turkiye": "pg.com",
    "procter & gamble turkiye": "pg.com",
    "nestle turkiye": "nestle.com.tr",
    "danone turkiye": "danone.com",
    "ulker": "ulker.com.tr",
    "yildiz holding": "yildizholding.com.tr",
    "dimes": "dimes.com.tr",
    "eti gida": "etietieti.com",
    "solen": "solen.com.tr",
    "pinar": "pinar.com.tr",
}
EMAIL_RE = re.compile(r"(?i)\b[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,}\b")


@dataclass
class SearchHit:
    company: str
    role_group: str
    query_title: str
    query: str
    person_name: str
    current_title: str
    linkedin_url: str
    result_title: str
    snippet: str
    score: int
    confidence: str
    source_domain: str
    manual_search_url: str


def normalize_text(value: str) -> str:
    value = unicodedata.normalize("NFKD", str(value or ""))
    value = "".join(ch for ch in value if not unicodedata.combining(ch))
    value = value.casefold()
    value = re.sub(r"[^a-z0-9&]+", " ", value)
    return re.sub(r"\s+", " ", value).strip()


def ascii_slug(value: str) -> str:
    value = unicodedata.normalize("NFKD", str(value or ""))
    value = "".join(ch for ch in value if not unicodedata.combining(ch))
    value = value.casefold().replace("ı", "i")
    return re.sub(r"[^a-z0-9]+", "", value)


NAME_SUFFIXES = {"msc", "mba", "phd", "pmp", "cpa", "cfa", "md", "dr", "prof", "ma", "ba", "bsc"}

def clean_name_parts(name: str) -> list[str]:
    raw = re.sub(r"\([^)]*\)", " ", str(name or ""))
    raw = raw.replace("–", " ").replace("—", " ")
    tokens = [ascii_slug(x) for x in re.split(r"[\s,;/]+", raw) if ascii_slug(x)]
    return [x for x in tokens if x not in NAME_SUFFIXES and len(x) > 1]

def record_key(company: str, name: str) -> str:
    return f"{normalize_text(company)}||{' '.join(clean_name_parts(name))}"


def clean_domain(value: str) -> str:
    raw = str(value or "").strip().lower()
    if not raw:
        return ""
    if "://" not in raw:
        raw = "https://" + raw
    domain = urlparse(raw).netloc.lower().split(":")[0]
    if domain.startswith("www."):
        domain = domain[4:]
    return domain


def domain_is_blocked(domain: str) -> bool:
    return any(domain == d or domain.endswith("." + d) for d in BLOCKED_DOMAINS)


def company_tokens(company: str) -> list[str]:
    stop = {"turkiye", "a", "as", "holding", "group", "international", "gida", "sanayi", "ticaret"}
    return [t for t in normalize_text(company).split() if len(t) >= 3 and t not in stop]


def plausible_name(text: str) -> bool:
    text = re.sub(r"\s+", " ", text).strip(" -|–—")
    if not text or len(text) > 70:
        return False
    words = text.split()
    if not (2 <= len(words) <= 6):
        return False
    lower = normalize_text(text)
    bad = ["linkedin", "jobs", "company", "people", "manager", "director", "marketing", "procurement", "satın"]
    return not any(item in lower for item in bad)


def parse_linkedin_title(title: str) -> tuple[str, str]:
    clean = re.sub(r"\s*[|·]\s*LinkedIn.*$", "", title, flags=re.I).strip()
    parts = [p.strip() for p in re.split(r"\s+[\-–—|]\s+", clean) if p.strip()]
    if parts and plausible_name(parts[0]):
        return parts[0], " — ".join(parts[1:])[:240]
    return "", clean[:240]


def score_hit(company: str, role_title: str, result_title: str, snippet: str, url: str, name: str) -> int:
    combined = normalize_text(f"{result_title} {snippet}")
    score = 0
    if "linkedin.com/in/" in url.lower(): score += 30
    if name: score += 15
    tokens = company_tokens(company)
    if tokens:
        matched = sum(1 for token in tokens if token in combined)
        score += min(30, matched * 12)
    role_norm = normalize_text(role_title)
    if role_norm and role_norm in combined: score += 20
    elif any(normalize_text(k) in combined for k in ROLE_KEYWORDS): score += 12
    if "turkiye" in combined or "turkey" in combined or "istanbul" in combined: score += 5
    if any(marker in combined for marker in FORMER_MARKERS): score -= 25
    return max(0, min(100, score))


def confidence_label(score: int) -> str:
    if score >= 75: return "Yüksek"
    if score >= 55: return "Orta"
    return "Düşük"


def tavily_search(api_key: str, query: str, count: int = 5, include_domains: list[str] | None = None, timeout: int = 45) -> list[dict]:
    payload = {
        "api_key": api_key.strip(),
        "query": query,
        "topic": "general",
        "search_depth": "basic",
        "max_results": min(max(count, 1), 20),
        "include_answer": False,
        "include_raw_content": False,
        "include_images": False,
        "country": "turkey",
    }
    if include_domains:
        payload["include_domains"] = include_domains
    response = requests.post(
        TAVILY_ENDPOINT,
        headers={"Accept": "application/json", "Content-Type": "application/json", "User-Agent": "ResetLeadFinder/10.0"},
        json=payload,
        timeout=timeout,
    )
    if response.status_code in (401, 403):
        raise RuntimeError("Tavily API anahtarı geçersiz veya yetkilendirme başarısız.")
    if response.status_code == 429:
        raise RuntimeError("Tavily hız sınırına ulaşıldı.")
    if response.status_code == 432:
        raise RuntimeError("Tavily aylık kredisi tükendi.")
    response.raise_for_status()
    return response.json().get("results", []) or []


def build_query(company: str, title: str) -> str:
    return f'site:linkedin.com/in "{company}" "{title}" Türkiye'


def extract_hits(company: str, role_group: str, query_title: str, query: str, results: Iterable[dict]) -> list[SearchHit]:
    hits: list[SearchHit] = []
    for item in results:
        url = str(item.get("url") or "").strip()
        if "linkedin.com/in/" not in url.lower():
            continue
        result_title = re.sub(r"<[^>]+>", "", str(item.get("title") or "")).strip()
        snippet = re.sub(r"<[^>]+>", "", str(item.get("content") or item.get("description") or "")).strip()
        name, current_title = parse_linkedin_title(result_title)
        score = score_hit(company, query_title, result_title, snippet, url, name)
        hits.append(SearchHit(company, role_group, query_title, query, name, current_title, url, result_title, snippet, score, confidence_label(score), urlparse(url).netloc, f"https://search.brave.com/search?q={quote_plus(query)}"))
    return hits


def dedupe_hits(hits: list[SearchHit], max_per_company_role: int) -> list[SearchHit]:
    best: dict[tuple[str, str, str], SearchHit] = {}
    for hit in hits:
        key_url = hit.linkedin_url.lower().split("?")[0].rstrip("/")
        key = (normalize_text(hit.company), normalize_text(hit.role_group), key_url)
        if key not in best or hit.score > best[key].score:
            best[key] = hit
    grouped: dict[tuple[str, str], list[SearchHit]] = {}
    for hit in best.values(): grouped.setdefault((hit.company, hit.role_group), []).append(hit)
    final: list[SearchHit] = []
    for group_hits in grouped.values(): final.extend(sorted(group_hits, key=lambda x: x.score, reverse=True)[:max_per_company_role])
    return sorted(final, key=lambda x: (x.company, x.role_group, -x.score))


def hits_to_frame(hits: list[SearchHit]) -> pd.DataFrame:
    columns = ["Firma", "Rol Grubu", "Ad Soyad", "Unvan / Profil Başlığı", "LinkedIn", "Güven Skoru", "Güven", "Kontrol Durumu", "Sonuç Başlığı", "Açıklama", "Aranan Unvan", "Tavily Sorgusu", "Manuel Arama Linki", "Kaynak Domain"]
    return pd.DataFrame([{
        "Firma": h.company, "Rol Grubu": h.role_group, "Ad Soyad": h.person_name,
        "Unvan / Profil Başlığı": h.current_title, "LinkedIn": h.linkedin_url,
        "Güven Skoru": h.score, "Güven": h.confidence, "Kontrol Durumu": "Manuel kontrol",
        "Sonuç Başlığı": h.result_title, "Açıklama": h.snippet, "Aranan Unvan": h.query_title,
        "Tavily Sorgusu": h.query, "Manuel Arama Linki": h.manual_search_url, "Kaynak Domain": h.source_domain,
    } for h in hits], columns=columns)


def read_input_file(uploaded_file) -> tuple[pd.DataFrame, str]:
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(uploaded_file), "CSV"
    xls = pd.ExcelFile(uploaded_file)
    preferred = "E-posta Aşaması" if "E-posta Aşaması" in xls.sheet_names else xls.sheet_names[0]
    sheet_name = st.selectbox("Excel sayfası", xls.sheet_names, index=xls.sheet_names.index(preferred))
    return pd.read_excel(xls, sheet_name=sheet_name), sheet_name


def load_pilot() -> pd.DataFrame:
    return pd.read_csv(PILOT_FILE)


def excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            safe_name = sheet_name[:31]
            df.to_excel(writer, index=False, sheet_name=safe_name)
            ws = writer.book[safe_name]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for cell in ws[1]:
                cell.font = cell.font.copy(bold=True, color="FFFFFF")
                cell.fill = cell.fill.copy(fill_type="solid", fgColor="111827")
            for col_cells in ws.columns:
                width = min(55, max(12, max(len(str(c.value or "")) for c in list(col_cells)[:200]) + 2))
                ws.column_dimensions[col_cells[0].column_letter].width = width
            for row in ws.iter_rows():
                for cell in row:
                    cell.alignment = cell.alignment.copy(vertical="top", wrap_text=True)
    output.seek(0)
    return output.getvalue()


def extract_emails(text: str) -> set[str]:
    emails = {e.strip(".,;:()[]{}<>\"'").lower() for e in EMAIL_RE.findall(str(text or ""))}
    return {e for e in emails if len(e) <= 150 and ".." not in e}


def is_generic_email(email: str) -> bool:
    local = email.split("@", 1)[0].lower()
    compact = ascii_slug(local)
    tokens = {ascii_slug(t) for t in re.split(r"[._+\-]+", local) if t}
    generic_compact = {ascii_slug(x) for x in GENERIC_LOCAL_PARTS}
    return (
        compact in generic_compact
        or bool(tokens & generic_compact)
        or any(local.startswith(prefix + sep) for prefix in GENERIC_LOCAL_PARTS for sep in (".", "_", "-", "+"))
    )


def person_email_match_score(email: str, name: str) -> int:
    """Kişi adı ile e-posta kullanıcı adının gerçekten eşleşip eşleşmediğini puanlar.

    Sadece aynı domaine ait olmak yeterli değildir. Genel/departman adresleri 0 puan alır.
    """
    if not email or is_generic_email(email):
        return 0
    local_raw = email.split("@", 1)[0]
    local = ascii_slug(local_raw)
    parts = clean_name_parts(name)
    if len(parts) < 2:
        return 0
    first, last = parts[0], parts[-1]
    all_last = "".join(parts[1:])
    exact = {
        first + last,
        first + all_last,
        last + first,
        first[0] + last,
        first[0] + all_last,
    }
    if local in exact:
        return 100
    # Nokta/alt çizgi gibi ayraçlar ascii_slug ile kalktığı için ad+soyad eşleşmesi yakalanır.
    if first in local and last in local:
        return 90
    if first in local and all_last and all_last in local:
        return 90
    if last in local and local.startswith(first[:1]):
        return 75
    return 0


def get_mx_status(domain: str) -> tuple[bool, str]:
    if not domain:
        return False, "Domain yok"
    try:
        answers = dns.resolver.resolve(domain, "MX", lifetime=5)
        hosts = sorted(str(r.exchange).rstrip(".") for r in answers)
        return bool(hosts), ", ".join(hosts[:3]) if hosts else "MX bulunamadı"
    except Exception as exc:
        try:
            socket.getaddrinfo(domain, 443)
            return False, f"MX doğrulanamadı; domain açılıyor ({type(exc).__name__})"
        except Exception:
            return False, "Domain/MX doğrulanamadı"


def candidate_domain_score(company: str, item: dict) -> tuple[int, str]:
    url = str(item.get("url") or "")
    domain = clean_domain(url)
    if not domain or domain_is_blocked(domain):
        return -100, domain
    combined = normalize_text(f"{item.get('title', '')} {item.get('content', '')} {domain}")
    score = 0
    tokens = company_tokens(company)
    score += sum(15 for token in tokens if token in combined)
    if any(x in combined for x in ["official", "resmi", "kurumsal"]): score += 10
    if domain.endswith(".com.tr") or domain.endswith(".tr"): score += 4
    path = urlparse(url).path.strip("/")
    if not path: score += 4
    return score, domain


def discover_domain(api_key: str, company: str) -> tuple[str, str, str]:
    hint = KNOWN_DOMAIN_HINTS.get(normalize_text(company))
    if hint:
        return hint, "Hazır güvenilir domain eşlemesi", f"https://{hint}"
    results = tavily_search(api_key, f'"{company}" resmi web sitesi Türkiye', count=7)
    ranked = sorted((candidate_domain_score(company, item)[0], candidate_domain_score(company, item)[1], str(item.get("url") or "")) for item in results)
    for score, domain, url in reversed(ranked):
        if score >= 10 and domain:
            return domain, f"Tavily resmi site araması (skor {score})", url
    return "", "Domain bulunamadı", ""


def fetch_public_site_emails(domain: str) -> tuple[set[str], list[str]]:
    if not domain:
        return set(), []
    emails: set[str] = set()
    sources: list[str] = []
    paths = ["/", "/iletisim", "/contact", "/contact-us", "/kurumsal/iletisim", "/tr/iletisim"]
    headers = {"User-Agent": "Mozilla/5.0 (compatible; ResetLeadFinder/10.0; public-contact-research)"}
    for path in paths:
        url = urljoin(f"https://{domain}", path)
        try:
            response = requests.get(url, headers=headers, timeout=8, allow_redirects=True)
            if response.status_code >= 400 or "text/html" not in response.headers.get("content-type", ""):
                continue
            found = {e for e in extract_emails(response.text) if clean_domain(e.split("@", 1)[1]) == domain or e.endswith("@" + domain)}
            if found:
                emails.update(found)
                sources.append(response.url)
        except Exception:
            continue
    return emails, list(dict.fromkeys(sources))


def search_person_email(api_key: str, name: str, company: str, domain: str) -> tuple[str, str, str]:
    domain_part = f' "@{domain}"' if domain else ""
    query = f'"{name}" "{company}" email OR e-mail{domain_part}'
    results = tavily_search(api_key, query, count=6)
    all_emails: list[tuple[str, str]] = []
    for item in results:
        text = f"{item.get('title', '')}\n{item.get('content', '')}\n{item.get('url', '')}"
        for email in extract_emails(text):
            if domain and not email.endswith("@" + domain):
                continue
            all_emails.append((email, str(item.get("url") or "")))
    if not all_emails:
        return "", "", query
    name_parts = clean_name_parts(name)
    first = name_parts[0] if name_parts else ""
    last = name_parts[-1] if name_parts else ""
    def email_score(pair: tuple[str, str]) -> int:
        email, _ = pair
        score = person_email_match_score(email, name)
        if domain and email.endswith("@" + domain):
            score += 10
        return score
    email, source = max(all_emails, key=email_score)
    # Aynı domaine ait genel adresleri kişi adresi sayma. Ad/soyad eşleşmesi zorunludur.
    if person_email_match_score(email, name) < 75:
        return "", "", query
    return email, source, query


def infer_pattern(site_emails: set[str]) -> str:
    personal = [e for e in site_emails if not is_generic_email(e)]
    locals_ = [e.split("@", 1)[0] for e in personal]
    if any("." in local for local in locals_):
        return "first.last"
    if locals_:
        return "firstlast"
    return "first.last"


def generate_candidates(name: str, domain: str, pattern_hint: str = "first.last") -> list[str]:
    parts = clean_name_parts(name)
    if len(parts) < 2 or not domain:
        return []
    first = parts[0]
    last = parts[-1]
    middle_last = "".join(parts[1:])
    patterns = {
        "first.last": f"{first}.{last}",
        "firstlast": f"{first}{last}",
        "f.last": f"{first[0]}.{last}",
        "flast": f"{first[0]}{last}",
        "last.first": f"{last}.{first}",
        "first": first,
        "first.alllast": f"{first}.{middle_last}",
        "firstalllast": f"{first}{middle_last}",
    }
    order = [pattern_hint, "first.last", "firstlast", "f.last", "flast", "first.alllast", "last.first", "first"]
    result: list[str] = []
    for key in order:
        local = patterns.get(key)
        if local:
            email = f"{local}@{domain}"
            if email not in result:
                result.append(email)
    return result[:6]


def _bool_value(value) -> bool | None:
    if isinstance(value, dict):
        value = value.get("value")
    if isinstance(value, bool):
        return value
    if value is None:
        return None
    text = str(value).strip().lower()
    if text in {"true", "1", "yes"}: return True
    if text in {"false", "0", "no"}: return False
    return None


def abstract_verify(api_key: str, email: str, timeout: int = 30) -> dict:
    """Abstract Email Reputation API sonucunu normalize eder."""
    if not api_key.strip() or not email:
        return {"status": "not_checked", "email": email}
    response = requests.get(
        ABSTRACT_ENDPOINT,
        params={"api_key": api_key.strip(), "email": email},
        headers={"Accept": "application/json", "User-Agent": "ResetLeadFinder/10.0"},
        timeout=timeout,
    )
    if response.status_code in (401, 403):
        raise RuntimeError("Abstract API anahtarı geçersiz veya bu API için yetkisiz.")
    if response.status_code == 429:
        raise RuntimeError("Abstract doğrulama limiti/kredisi aşıldı.")
    response.raise_for_status()
    data = response.json()

    # Yeni cevap yapısı
    deliverability_obj = data.get("email_deliverability") or {}
    quality_obj = data.get("email_quality") or {}
    if deliverability_obj:
        status = str(deliverability_obj.get("status") or "unknown").lower()
        smtp = _bool_value(deliverability_obj.get("is_smtp_valid"))
        mx = _bool_value(deliverability_obj.get("is_mx_valid"))
        catchall = _bool_value(quality_obj.get("is_catchall"))
        role = _bool_value(quality_obj.get("is_role"))
        quality = quality_obj.get("score")
        detail = str(deliverability_obj.get("status_detail") or "")
    else:
        # Klasik cevap yapısı
        status = str(data.get("deliverability") or "unknown").lower()
        smtp = _bool_value(data.get("is_smtp_valid"))
        mx = _bool_value(data.get("is_mx_found"))
        catchall = _bool_value(data.get("is_catchall_email"))
        role = _bool_value(data.get("is_role_email"))
        quality = data.get("quality_score")
        detail = str(data.get("autocorrect") or "")

    deliverable = status in {"deliverable", "valid"} and smtp is not False
    if catchall is True:
        normalized = "catch_all"
    elif deliverable and smtp is True:
        normalized = "deliverable"
    elif status in {"undeliverable", "invalid"} or smtp is False:
        normalized = "undeliverable"
    else:
        normalized = "unknown"
    return {
        "email": email,
        "status": normalized,
        "provider_status": status,
        "smtp": smtp,
        "mx": mx,
        "catchall": catchall,
        "role": role,
        "quality": quality,
        "detail": detail,
    }


def verify_candidates(abstract_key: str, candidates: list[str], max_checks: int = 4, delay: float = 0.15) -> tuple[str, str, list[dict]]:
    """Adayları sırayla doğrular. Kesin teslim edilebilir adresi, yoksa catch-all/unknown adayı döndürür."""
    checks: list[dict] = []
    if not abstract_key.strip():
        return "", "not_checked", checks
    fallback = ""
    fallback_status = ""
    for email in candidates[:max_checks]:
        result = abstract_verify(abstract_key, email)
        checks.append(result)
        status = result.get("status")
        if status == "deliverable":
            return email, "deliverable", checks
        if not fallback and status in {"catch_all", "unknown"}:
            fallback, fallback_status = email, status
        if delay: time.sleep(delay)
    return fallback, fallback_status or "undeliverable", checks



def split_name_for_lusha(name: str) -> tuple[str, str]:
    """Lusha için adı ilk ad + son soyad şeklinde ayırır."""
    raw = re.sub(r"\([^)]*\)", " ", str(name or ""))
    raw = raw.replace("–", " ").replace("—", " ")
    tokens = [x.strip(".,;:/") for x in raw.split() if x.strip(".,;:/")]
    tokens = [x for x in tokens if ascii_slug(x) not in NAME_SUFFIXES]
    if len(tokens) < 2:
        return (tokens[0], "") if tokens else ("", "")
    return tokens[0], tokens[-1]


def lusha_match_score(original_name: str, original_company: str, original_linkedin: str, result: dict) -> tuple[int, str]:
    """Yanlış kişi eşleşmesini azaltmak için Lusha sonucunu giriş kaydıyla kıyaslar."""
    score = 0
    returned_name = str(result.get("fullName") or f"{result.get('firstName', '')} {result.get('lastName', '')}").strip()
    returned_company = str((result.get("company") or {}).get("name") or "")
    returned_domain = str((result.get("company") or {}).get("domain") or "")
    social = result.get("socialLinks") or {}
    returned_linkedin = str(social.get("linkedin") or social.get("linkedinUrl") or "").strip()

    if original_linkedin and returned_linkedin:
        a = original_linkedin.lower().split("?")[0].rstrip("/")
        b = returned_linkedin.lower().split("?")[0].rstrip("/")
        if a == b:
            score += 100

    original_parts = clean_name_parts(original_name)
    returned_parts = clean_name_parts(returned_name)
    if original_parts and returned_parts:
        if original_parts[0] == returned_parts[0]:
            score += 28
        if original_parts[-1] == returned_parts[-1]:
            score += 42
        if " ".join(original_parts) == " ".join(returned_parts):
            score += 20

    comp_text = normalize_text(f"{returned_company} {returned_domain}")
    tokens = company_tokens(original_company)
    if tokens and any(token in comp_text for token in tokens):
        score += 25

    score = min(100, score)
    if score >= 75:
        return score, "Yüksek"
    if score >= 50:
        return score, "Orta"
    return score, "Düşük"


def lusha_search_and_enrich(api_key: str, contacts: list[dict], reveal_email: bool = True,
                            reveal_phone: bool = False, timeout: int = 90) -> dict:
    """Lusha V3 search-and-enrich. Bir istekte en fazla 100 kişi gönderilir."""
    if not api_key.strip() or not contacts:
        return {"results": [], "billing": {}}
    if len(contacts) > 100:
        raise ValueError("Lusha tek istekte en fazla 100 kişi kabul eder.")
    reveal: list[str] = []
    if reveal_email:
        reveal.append("emails")
    if reveal_phone:
        reveal.append("phones")
    payload: dict = {"contacts": contacts}
    if reveal:
        payload["reveal"] = reveal
    response = requests.post(
        LUSHA_ENDPOINT,
        headers={
            "api_key": api_key.strip(),
            "Accept": "application/json",
            "Content-Type": "application/json",
            "User-Agent": "ResetLeadFinder/10.0",
        },
        json=payload,
        timeout=timeout,
    )
    if response.status_code in (401, 403):
        raise RuntimeError("Lusha API anahtarı geçersiz veya hesabında API erişimi yok.")
    if response.status_code == 402:
        raise RuntimeError("Lusha kredisi veya API anahtarı kredi limiti tükendi.")
    if response.status_code == 429:
        raise RuntimeError("Lusha hız sınırına ulaşıldı. Bir süre sonra tekrar dene.")
    if response.status_code >= 400:
        try:
            detail = response.json().get("message") or response.text[:500]
        except Exception:
            detail = response.text[:500]
        raise RuntimeError(f"Lusha API hatası ({response.status_code}): {detail}")
    data = response.json()
    if not isinstance(data.get("results"), list):
        raise RuntimeError("Lusha yanıtında results listesi bulunamadı.")
    return data


def parse_lusha_result(result: dict) -> dict:
    error = result.get("error") or {}
    if error:
        code = str(error.get("code") or "ERROR")
        message = str(error.get("message") or "")
        return {
            "status": "not_found" if code == "NOT_FOUND" else "error",
            "error_code": code,
            "error_message": message,
            "raw": result,
        }
    emails = result.get("emails") or []
    phones = result.get("phones") or []
    job = result.get("jobTitle") or {}
    company = result.get("company") or {}
    social = result.get("socialLinks") or {}
    first_email = emails[0] if emails else {}
    first_phone = phones[0] if phones else {}
    return {
        "status": "success",
        "id": str(result.get("id") or ""),
        "full_name": str(result.get("fullName") or f"{result.get('firstName', '')} {result.get('lastName', '')}").strip(),
        "email": str(first_email.get("email") or "").strip().lower(),
        "email_type": str(first_email.get("type") or ""),
        "email_confidence": str(first_email.get("confidence") or ""),
        "email_updated": str(first_email.get("updateDate") or ""),
        "phone": str(first_phone.get("number") or ""),
        "do_not_call": bool(first_phone.get("doNotCall")) if first_phone else False,
        "job_title": str(job.get("title") or ""),
        "job_seniority": str(job.get("seniority") or ""),
        "company_name": str(company.get("name") or ""),
        "company_domain": str(company.get("domain") or ""),
        "linkedin": str(social.get("linkedin") or social.get("linkedinUrl") or ""),
        "raw": result,
    }


def enrich_with_lusha(rows: list[dict], lusha_key: str, abstract_key: str, mode: str,
                      reveal_phone: bool, cross_verify_email: bool, delay: float,
                      errors: list[dict]) -> tuple[list[dict], list[dict]]:
    """Abstract sonrası gereken kayıtları Lusha ile toplu zenginleştirir."""
    if not lusha_key.strip() or mode == "off" or not rows:
        for row in rows:
            row.setdefault("Lusha Durumu", "Kullanılmadı")
        return rows, []

    payload_items: list[dict] = []
    ref_to_index: dict[str, int] = {}
    for index, row in enumerate(rows):
        abstract_status = str(row.get("Doğrulama Durumu") or "").lower()
        should_query = mode == "all" or abstract_status != "deliverable"
        if not should_query:
            row["Lusha Durumu"] = "Atlandı — Abstract doğrulandı"
            continue
        ref = f"reset-{index}"
        first, last = split_name_for_lusha(str(row.get("Ad Soyad") or ""))
        contact: dict = {"clientReferenceId": ref}
        linkedin = str(row.get("LinkedIn") or "").strip()
        current_email = str(row.get("Önerilen E-posta") or "").strip()
        domain = str(row.get("Şirket Domaini") or "").strip()
        company = str(row.get("Firma") or "").strip()
        if linkedin:
            contact["linkedinUrl"] = linkedin
        elif current_email and abstract_status in {"deliverable", "catch_all", "unknown"}:
            contact["email"] = current_email
        else:
            if first:
                contact["firstName"] = first
            if last:
                contact["lastName"] = last
            if domain:
                contact["companyDomain"] = domain
            elif company:
                contact["companyName"] = company
        if len(contact) <= 1:
            row["Lusha Durumu"] = "Atlandı — yetersiz kimlik bilgisi"
            continue
        payload_items.append(contact)
        ref_to_index[ref] = index
        row["Lusha Durumu"] = "Sorgulanıyor"

    billing_rows: list[dict] = []
    for start in range(0, len(payload_items), 100):
        chunk = payload_items[start:start + 100]
        try:
            response = lusha_search_and_enrich(
                lusha_key, chunk, reveal_email=True, reveal_phone=reveal_phone
            )
        except Exception as exc:
            for item in chunk:
                idx = ref_to_index.get(str(item.get("clientReferenceId")))
                if idx is not None:
                    rows[idx]["Lusha Durumu"] = "API hatası"
                    rows[idx]["Lusha Hata"] = str(exc)
                    errors.append({
                        "Firma": rows[idx].get("Firma", ""),
                        "Ad Soyad": rows[idx].get("Ad Soyad", ""),
                        "Aşama": "Lusha",
                        "Hata": str(exc),
                    })
            continue

        billing = response.get("billing") or {}
        if billing:
            billing_rows.append({"Batch": start // 100 + 1, "Billing": json.dumps(billing, ensure_ascii=False)})

        seen: set[str] = set()
        for raw_result in response.get("results", []):
            ref = str(raw_result.get("clientReferenceId") or "")
            seen.add(ref)
            idx = ref_to_index.get(ref)
            if idx is None:
                continue
            row = rows[idx]
            parsed = parse_lusha_result(raw_result)
            if parsed.get("status") == "not_found":
                row["Lusha Durumu"] = "Bulunamadı"
                row["Lusha Hata Kodu"] = parsed.get("error_code", "")
                continue
            if parsed.get("status") == "error":
                row["Lusha Durumu"] = "Hata"
                row["Lusha Hata Kodu"] = parsed.get("error_code", "")
                row["Lusha Hata"] = parsed.get("error_message", "")
                continue

            match_score, match_label = lusha_match_score(
                str(row.get("Ad Soyad") or ""),
                str(row.get("Firma") or ""),
                str(row.get("LinkedIn") or ""),
                raw_result,
            )
            row.update({
                "Lusha Durumu": "Başarılı",
                "Lusha ID": parsed.get("id", ""),
                "Lusha Eşleşme": match_label,
                "Lusha Eşleşme Skoru": match_score,
                "Lusha E-posta": parsed.get("email", ""),
                "Lusha E-posta Tipi": parsed.get("email_type", ""),
                "Lusha E-posta Güveni": parsed.get("email_confidence", ""),
                "Lusha E-posta Güncelleme": parsed.get("email_updated", ""),
                "Lusha Telefon": parsed.get("phone", ""),
                "Lusha Do Not Call": parsed.get("do_not_call", False),
                "Lusha Güncel Unvan": parsed.get("job_title", ""),
                "Lusha Kıdem": parsed.get("job_seniority", ""),
                "Lusha Şirket": parsed.get("company_name", ""),
                "Lusha Şirket Domaini": parsed.get("company_domain", ""),
                "Lusha LinkedIn": parsed.get("linkedin", ""),
            })

            if match_label == "Düşük":
                row["Önerilen Aksiyon"] = "Lusha eşleşmesi düşük; kişiyi manuel kontrol et"
                continue

            lusha_email = str(parsed.get("email") or "").strip().lower()
            if lusha_email:
                lusha_check: dict = {}
                if cross_verify_email and abstract_key.strip():
                    try:
                        lusha_check = abstract_verify(abstract_key, lusha_email)
                    except Exception as exc:
                        errors.append({
                            "Firma": row.get("Firma", ""),
                            "Ad Soyad": row.get("Ad Soyad", ""),
                            "Aşama": "Lusha e-posta çapraz doğrulama",
                            "Hata": str(exc),
                        })
                row["Lusha SMTP Durumu"] = lusha_check.get("status", "not_checked") if lusha_check else "not_checked"
                row["Lusha SMTP Geçerli"] = lusha_check.get("smtp", "") if lusha_check else ""
                row["Lusha Catch-all"] = lusha_check.get("catchall", "") if lusha_check else ""

                if lusha_check.get("status") == "deliverable":
                    row["Önerilen E-posta"] = lusha_email
                    row["E-posta Durumu"] = "Lusha + SMTP doğrulandı"
                    row["Güven"] = "Yüksek"
                    row["Doğrulama Servisi"] = "Lusha + Abstract API"
                    row["Doğrulama Durumu"] = "deliverable"
                    row["SMTP Geçerli"] = lusha_check.get("smtp", "")
                    row["Catch-all"] = lusha_check.get("catchall", "")
                    row["Kalite Skoru"] = lusha_check.get("quality", "")
                    row["Önerilen Aksiyon"] = "Kişiselleştirilmiş e-posta ve LinkedIn teması için kullanılabilir"
                elif not cross_verify_email or not abstract_key.strip():
                    row["Önerilen E-posta"] = lusha_email
                    row["E-posta Durumu"] = "Lusha kayıtlı iş e-postası"
                    row["Güven"] = "Yüksek" if match_label == "Yüksek" else "Orta"
                    row["Doğrulama Servisi"] = "Lusha"
                    row["Doğrulama Durumu"] = "lusha_found"
                    row["Önerilen Aksiyon"] = "Düşük hacimde kişiselleştirilmiş iletişim için kullanılabilir"
                elif lusha_check.get("status") in {"catch_all", "unknown"}:
                    row["Önerilen E-posta"] = lusha_email
                    row["E-posta Durumu"] = "Lusha bulundu — SMTP belirsiz"
                    row["Güven"] = "Orta"
                    row["Doğrulama Servisi"] = "Lusha + Abstract API"
                    row["Doğrulama Durumu"] = "lusha_found"
                    row["Önerilen Aksiyon"] = "Toplu gönderme; düşük hacimde kontrollü dene"

            phone = str(parsed.get("phone") or "").strip()
            if phone:
                if parsed.get("do_not_call"):
                    row["Önerilen Telefon"] = ""
                    row["Telefon Durumu"] = "Lusha Do Not Call — kullanma"
                else:
                    row["Önerilen Telefon"] = phone
                    row["Telefon Durumu"] = "Lusha kaynağı"
                    if not row.get("Önerilen E-posta"):
                        row["Önerilen Aksiyon"] = "Telefon ve LinkedIn üzerinden kişiselleştirilmiş temas"

            if delay:
                time.sleep(min(delay, 0.25))

        for item in chunk:
            ref = str(item.get("clientReferenceId") or "")
            if ref not in seen:
                idx = ref_to_index.get(ref)
                if idx is not None:
                    rows[idx]["Lusha Durumu"] = "Yanıtta kayıt yok"

    return rows, billing_rows


def enrich_emails(api_key: str, abstract_key: str, source_df: pd.DataFrame, company_col: str, name_col: str,
                  role_col: str | None, priority_col: str | None, linkedin_col: str | None,
                  provided_domain_col: str | None, delay: float, max_verifications: int = 4,
                  lusha_key: str = "", lusha_mode: str = "off", lusha_reveal_phone: bool = False,
                  lusha_cross_verify: bool = True) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    work = source_df.copy()
    domain_cache: dict[str, tuple[str, str, str, bool, str, set[str], list[str]]] = {}
    rows: list[dict] = []
    errors: list[dict] = []
    companies = [str(x).strip() for x in work[company_col].dropna().unique() if str(x).strip()]
    total_steps = len(companies) + len(work)
    progress = st.progress(0)
    status = st.empty()
    step = 0

    for company in companies:
        status.write(f"**{company}** · şirket domaini ve e-posta formatı araştırılıyor")
        subset = work[work[company_col].astype(str).str.strip() == company]
        provided = ""
        if provided_domain_col:
            vals = subset[provided_domain_col].dropna().astype(str).map(clean_domain)
            provided = next((v for v in vals if v), "")
        try:
            if provided:
                domain, domain_note, domain_source = provided, "Dosyadaki domain kullanıldı", f"https://{provided}"
            else:
                domain, domain_note, domain_source = discover_domain(api_key, company)
            mx_ok, mx_detail = get_mx_status(domain)
            site_emails, site_sources = fetch_public_site_emails(domain)
            domain_cache[company] = (domain, domain_note, domain_source, mx_ok, mx_detail, site_emails, site_sources)
        except Exception as exc:
            errors.append({"Firma": company, "Aşama": "Domain", "Hata": str(exc)})
            domain_cache[company] = ("", "Hata", "", False, "", set(), [])
        step += 1
        progress.progress(min(1.0, step / max(total_steps, 1)))
        if delay:
            time.sleep(delay)

    for _, original in work.iterrows():
        company = str(original.get(company_col, "") or "").strip()
        name = str(original.get(name_col, "") or "").strip()
        role = str(original.get(role_col, "") or "").strip() if role_col else ""
        priority = str(original.get(priority_col, "") or "").strip() if priority_col else ""
        linkedin = str(original.get(linkedin_col, "") or "").strip() if linkedin_col else ""
        domain, domain_note, domain_source, mx_ok, mx_detail, site_emails, site_sources = domain_cache.get(company, ("", "", "", False, "", set(), []))
        status.write(f"**{company}** · {name} için e-posta aranıyor")
        public_email = ""
        email_source = ""
        search_query = ""
        try:
            public_email, email_source, search_query = search_person_email(api_key, name, company, domain)
        except Exception as exc:
            errors.append({"Firma": company, "Ad Soyad": name, "Aşama": "Kişi e-postası", "Hata": str(exc)})
        hint = infer_pattern(site_emails)
        candidates = generate_candidates(name, domain, hint)
        verify_pool = ([public_email] if public_email else []) + [c for c in candidates if c != public_email]
        verified_email = ""
        verification_status = "not_checked"
        verification_checks: list[dict] = []
        try:
            verified_email, verification_status, verification_checks = verify_candidates(
                abstract_key, verify_pool, max_checks=max_verifications, delay=min(delay, 0.25)
            )
        except Exception as exc:
            errors.append({"Firma": company, "Ad Soyad": name, "Aşama": "E-posta doğrulama", "Hata": str(exc)})

        selected = verified_email or public_email or (candidates[0] if candidates else "")
        if verification_status == "deliverable":
            status_label = "SMTP doğrulandı"
            confidence = "Yüksek"
            action = "Tekil ve kişiselleştirilmiş iletişim için kullanılabilir"
        elif verification_status == "catch_all":
            status_label = "Catch-all — posta kutusu kesin değil"
            confidence = "Orta"
            action = "Toplu gönderme; kişiselleştirilmiş düşük hacimde dene"
        elif verification_status == "unknown":
            status_label = "Doğrulama belirsiz"
            confidence = "Düşük"
            action = "Lusha veya ikinci bir kaynakla doğrula"
        elif abstract_key.strip() and verification_status == "undeliverable":
            status_label = "Adaylar doğrulanamadı"
            confidence = "Düşük"
            selected = ""
            action = "Lusha fallback veya LinkedIn teması kullan"
        elif public_email:
            status_label = "Açık kaynakta bulundu — doğrulanmadı"
            confidence = "Orta"
            action = "Gönderimden önce doğrulama servisi kullan"
        elif selected and mx_ok:
            status_label = "Tahmini kurumsal format"
            confidence = "Orta"
            action = "Toplu gönderme; önce doğrula"
        elif selected:
            status_label = "Tahmini — MX doğrulanamadı"
            confidence = "Düşük"
            action = "Kullanma; manuel doğrulama gerekli"
        else:
            status_label = "Bulunamadı"
            confidence = "Düşük"
            action = "Lusha fallback veya LinkedIn teması kullan"

        chosen_check = next((x for x in verification_checks if x.get("email") == selected), {})
        check_summary = " | ".join(f"{x.get('email')}: {x.get('status')}" for x in verification_checks)
        row = {
            "Firma": company,
            "Ad Soyad": name,
            "Rol": role,
            "Öncelik": priority,
            "LinkedIn": linkedin,
            "Şirket Domaini": domain,
            "Önerilen E-posta": selected,
            "E-posta Durumu": status_label,
            "Güven": confidence,
            "Doğrulama Servisi": "Abstract API" if abstract_key.strip() else "Kullanılmadı",
            "Doğrulama Durumu": verification_status,
            "SMTP Geçerli": chosen_check.get("smtp", ""),
            "Catch-all": chosen_check.get("catchall", ""),
            "Kalite Skoru": chosen_check.get("quality", ""),
            "Doğrulanan Adaylar": check_summary,
            "Alternatif 1": candidates[1] if len(candidates) > 1 else "",
            "Alternatif 2": candidates[2] if len(candidates) > 2 else "",
            "Alternatif 3": candidates[3] if len(candidates) > 3 else "",
            "MX Durumu": "Var" if mx_ok else "Doğrulanamadı",
            "MX Detayı": mx_detail,
            "Bulunan Genel E-postalar": ", ".join(sorted(site_emails)[:8]),
            "E-posta Kaynağı": email_source,
            "Domain Kaynağı": domain_source,
            "Domain Notu": domain_note,
            "Site Kaynakları": " | ".join(site_sources[:5]),
            "Tavily Arama Sorgusu": search_query,
            "Önerilen Aksiyon": action,
            "İletişim Durumu": "Araştırıldı",
            "Önerilen Telefon": "",
            "Telefon Durumu": "",
            "Lusha Durumu": "Bekliyor" if lusha_mode != "off" else "Kullanılmadı",
        }
        rows.append(row)
        step += 1
        progress.progress(min(1.0, step / max(total_steps, 1)))
        if delay:
            time.sleep(delay)

    if lusha_mode != "off" and lusha_key.strip():
        status.write("**Lusha** · başarısız veya seçili kayıtlar toplu zenginleştiriliyor")
    rows, billing_rows = enrich_with_lusha(
        rows, lusha_key, abstract_key, lusha_mode, lusha_reveal_phone,
        lusha_cross_verify, delay, errors,
    )
    status.success(f"Hibrit araştırma tamamlandı: {len(rows)} kişi işlendi.")
    return pd.DataFrame(rows), pd.DataFrame(errors), pd.DataFrame(billing_rows)



def read_uploaded_sheet(uploaded_file, widget_key: str, preferred: list[str]) -> tuple[pd.DataFrame, str]:
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(uploaded_file), "CSV"
    xls = pd.ExcelFile(uploaded_file)
    selected = next((x for x in preferred if x in xls.sheet_names), xls.sheet_names[0])
    sheet_name = st.selectbox(
        "Excel sayfası",
        xls.sheet_names,
        index=xls.sheet_names.index(selected),
        key=widget_key,
    )
    return pd.read_excel(xls, sheet_name=sheet_name), sheet_name



def standardize_previous_results(previous_df: pd.DataFrame, lusha_policy: str = "off") -> pd.DataFrame:
    if previous_df is None or previous_df.empty:
        return pd.DataFrame()
    company_col = next((c for c in ["Firma", "Şirket"] if c in previous_df.columns), None)
    name_col = next((c for c in ["Ad Soyad", "Kişi"] if c in previous_df.columns), None)
    if not company_col or not name_col:
        return pd.DataFrame()
    out = previous_df.copy()
    out["__key"] = [record_key(c, n) for c, n in zip(out[company_col], out[name_col])]
    status = out.get("Doğrulama Durumu", pd.Series("", index=out.index)).astype(str).str.lower()
    email_col = "Önerilen E-posta" if "Önerilen E-posta" in out.columns else ("E-posta" if "E-posta" in out.columns else None)
    has_email = out[email_col].fillna("").astype(str).str.contains("@") if email_col else pd.Series(False, index=out.index)
    lusha_status = out.get("Lusha Durumu", pd.Series("", index=out.index)).astype(str).str.lower()
    lusha_attempted = lusha_status.isin(["başarılı", "bulunamadı", "hata", "yanıtta kayıt yok", "atlandı — yetersiz kimlik bilgisi"])

    if lusha_policy == "all":
        # Önceki Abstract sonuçlarını Lusha ile ayrıca çapraz kontrol etmek için yeniden aç.
        out["__done"] = lusha_attempted
    elif lusha_policy == "fallback":
        # Abstract deliverable kayıtları tamam; diğerleri Lusha sorgulanmışsa tamam.
        out["__done"] = status.eq("deliverable") | lusha_attempted
    else:
        out["__done"] = status.isin(["deliverable", "catch_all", "unknown", "undeliverable", "lusha_found"]) | has_email
    return out



def merge_result_frames(previous: pd.DataFrame, new: pd.DataFrame) -> pd.DataFrame:
    frames = []
    for df in [previous, new]:
        if df is None or df.empty:
            continue
        x = df.copy()
        if "__key" not in x.columns and {"Firma", "Ad Soyad"}.issubset(x.columns):
            x["__key"] = [record_key(c, n) for c, n in zip(x["Firma"], x["Ad Soyad"])]
        frames.append(x)
    if not frames:
        return pd.DataFrame()
    combined = pd.concat(frames, ignore_index=True, sort=False)
    if "__key" in combined.columns:
        combined = combined.drop_duplicates("__key", keep="last")
    return combined



def build_crm_sheet(target_df: pd.DataFrame, company_col: str, name_col: str, role_col: str | None,
                    priority_col: str | None, linkedin_col: str | None, combined_results: pd.DataFrame) -> pd.DataFrame:
    base = pd.DataFrame({
        "Firma": target_df[company_col].fillna("").astype(str).str.strip(),
        "Ad Soyad": target_df[name_col].fillna("").astype(str).str.strip(),
        "Rol": target_df[role_col].fillna("").astype(str).str.strip() if role_col else "",
        "Öncelik": target_df[priority_col].fillna("").astype(str).str.strip() if priority_col else "",
        "LinkedIn": target_df[linkedin_col].fillna("").astype(str).str.strip() if linkedin_col else "",
    })
    base["__key"] = [record_key(c, n) for c, n in zip(base["Firma"], base["Ad Soyad"])]
    merge_columns = [
        "__key", "Şirket Domaini", "Önerilen E-posta", "E-posta Durumu", "Güven",
        "Doğrulama Durumu", "SMTP Geçerli", "Catch-all", "Kalite Skoru",
        "Önerilen Aksiyon", "İletişim Durumu", "MX Durumu", "Domain Kaynağı",
        "Önerilen Telefon", "Telefon Durumu", "Lusha Durumu", "Lusha ID",
        "Lusha Eşleşme", "Lusha Eşleşme Skoru", "Lusha E-posta", "Lusha E-posta Güveni",
        "Lusha Telefon", "Lusha Do Not Call", "Lusha Güncel Unvan", "Lusha Şirket",
        "Lusha LinkedIn", "Lusha SMTP Durumu",
    ]
    if combined_results is not None and not combined_results.empty and "__key" in combined_results.columns:
        available = [c for c in merge_columns if c in combined_results.columns]
        lookup = combined_results[available].drop_duplicates("__key", keep="last")
        base = base.merge(lookup, on="__key", how="left")
    for col in merge_columns[1:]:
        if col not in base.columns:
            base[col] = ""

    def crm_status(row) -> str:
        status = str(row.get("Doğrulama Durumu", "") or "").lower()
        phone = str(row.get("Önerilen Telefon", "") or "").strip()
        if status in {"deliverable", "lusha_found"}:
            return "İletişime hazır"
        if phone:
            return "Telefon hazır"
        if status in {"catch_all", "unknown"}:
            return "Manuel kontrol"
        if status == "undeliverable" or str(row.get("Lusha Durumu", "")).lower() == "bulunamadı":
            return "LinkedIn teması"
        return "Araştırılacak"

    base["CRM Durumu"] = base.apply(crm_status, axis=1)

    def channel(row) -> str:
        email = str(row.get("Önerilen E-posta", "") or "").strip()
        phone = str(row.get("Önerilen Telefon", "") or "").strip()
        if email and phone:
            return "E-posta + Telefon + LinkedIn"
        if email:
            return "E-posta + LinkedIn"
        if phone:
            return "Telefon + LinkedIn"
        if row.get("CRM Durumu") == "Manuel kontrol":
            return "LinkedIn + kontrollü e-posta"
        if row.get("CRM Durumu") == "LinkedIn teması":
            return "LinkedIn"
        return "Araştırılacak"

    base["Kullanılacak Kanal"] = base.apply(channel, axis=1)
    base["İlk Temas Tarihi"] = ""
    base["Yanıt Durumu"] = "Bekliyor"
    base["Takip Tarihi"] = ""
    base["Kişiselleştirme Notu"] = ""
    priority_order = {"A": 0, "B": 1, "C": 2, "": 9}
    base["__p"] = base["Öncelik"].map(priority_order).fillna(8)
    base = base.sort_values(["__p", "Firma", "Ad Soyad"]).drop(columns=["__key", "__p"])
    return base


def build_summary(crm: pd.DataFrame, combined: pd.DataFrame) -> pd.DataFrame:
    statuses = crm.get("CRM Durumu", pd.Series(dtype=str)).astype(str)
    ready = int(statuses.isin(["İletişime hazır", "Telefon hazır"]).sum())
    manual = int((statuses == "Manuel kontrol").sum())
    linkedin = int((statuses == "LinkedIn teması").sum())
    waiting = int((statuses == "Araştırılacak").sum())
    lusha_success = int((crm.get("Lusha Durumu", pd.Series(dtype=str)).astype(str) == "Başarılı").sum())
    return pd.DataFrame({
        "Metrik": ["Toplam hedef kişi", "İletişime hazır", "Manuel kontrol", "LinkedIn teması", "Araştırılmayı bekleyen", "Lusha başarılı"],
        "Değer": [len(crm), ready, manual, linkedin, waiting, lusha_success],
    })


st.set_page_config(page_title="RE:SET Lead Intelligence", page_icon="R", layout="wide", initial_sidebar_state="expanded")


def apply_reset_brand() -> None:
    st.markdown(
        """
        <style>
        :root {
          --reset-ink:#0B0B0C;
          --reset-paper:#F5F4F0;
          --reset-white:#FFFFFF;
          --reset-muted:#77746D;
          --reset-line:rgba(11,11,12,.12);
          --reset-accent:#FF4D36;
          --reset-accent-soft:rgba(255,77,54,.10);
          --reset-radius:18px;
        }
        html, body, [class*="css"] { font-family: Inter, ui-sans-serif, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; }
        .stApp { background:var(--reset-paper); color:var(--reset-ink); }
        [data-testid="stHeader"] { background:transparent; height:0; }
        #MainMenu, footer, [data-testid="stToolbar"], [data-testid="stDecoration"] { visibility:hidden; height:0; }
        .block-container { max-width:1440px; padding-top:2.4rem; padding-bottom:4rem; }
        [data-testid="stSidebar"] { background:var(--reset-ink); border-right:0; }
        [data-testid="stSidebar"] * { color:#F5F4F0; }
        [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] a { color:#FF8A78; }
        [data-testid="stSidebar"] .stButton button { background:#F5F4F0; color:#0B0B0C; border:0; }
        .reset-side-brand { padding:8px 2px 24px; border-bottom:1px solid rgba(255,255,255,.15); margin-bottom:22px; }
        .reset-side-mark { font-size:30px; line-height:1; letter-spacing:-.07em; font-weight:900; color:#fff; }
        .reset-side-mark span { color:var(--reset-accent); }
        .reset-side-caption { margin-top:8px; font-size:11px; letter-spacing:.16em; text-transform:uppercase; color:rgba(255,255,255,.56); }
        .reset-hero { position:relative; overflow:hidden; background:var(--reset-ink); color:#fff; border-radius:28px; padding:34px 38px 36px; margin-bottom:24px; box-shadow:0 24px 70px rgba(11,11,12,.12); }
        .reset-hero:after { content:""; position:absolute; width:320px; height:320px; right:-120px; top:-150px; border:1px solid rgba(255,255,255,.18); border-radius:50%; box-shadow:0 0 0 54px rgba(255,255,255,.035),0 0 0 110px rgba(255,255,255,.025); }
        .reset-kicker { position:relative; z-index:1; font-size:11px; letter-spacing:.18em; text-transform:uppercase; color:#B9B6AF; font-weight:700; }
        .reset-title { position:relative; z-index:1; margin:14px 0 12px; font-size:clamp(38px,5.5vw,76px); line-height:.92; letter-spacing:-.065em; font-weight:900; }
        .reset-title .accent { color:var(--reset-accent); }
        .reset-subtitle { position:relative; z-index:1; max-width:760px; margin:0; color:#CAC7C1; font-size:16px; line-height:1.6; }
        .reset-status { position:relative; z-index:1; display:flex; gap:8px; flex-wrap:wrap; margin-top:20px; }
        .reset-chip { display:inline-flex; align-items:center; gap:7px; padding:7px 10px; border-radius:999px; border:1px solid rgba(255,255,255,.16); font-size:12px; color:#E9E6E0; background:rgba(255,255,255,.045); }
        .reset-dot { width:7px; height:7px; border-radius:50%; background:#65D58B; box-shadow:0 0 0 4px rgba(101,213,139,.12); }
        [data-testid="stMetric"] { background:var(--reset-white); border:1px solid var(--reset-line); padding:17px 18px; border-radius:var(--reset-radius); box-shadow:0 8px 30px rgba(11,11,12,.035); }
        [data-testid="stMetricLabel"] { color:var(--reset-muted); }
        [data-testid="stMetricValue"] { letter-spacing:-.04em; }
        [data-testid="stExpander"], [data-testid="stForm"], [data-testid="stFileUploader"] { background:rgba(255,255,255,.68); border:1px solid var(--reset-line); border-radius:var(--reset-radius); }
        [data-testid="stTabs"] [data-baseweb="tab-list"] { gap:8px; background:transparent; }
        [data-testid="stTabs"] button[role="tab"] { border-radius:999px; padding:10px 18px; border:1px solid var(--reset-line); background:rgba(255,255,255,.72); }
        [data-testid="stTabs"] button[aria-selected="true"] { color:#fff; background:var(--reset-ink); border-color:var(--reset-ink); }
        [data-testid="stTabs"] [data-baseweb="tab-highlight"] { display:none; }
        .stButton > button, .stDownloadButton > button { border-radius:14px; min-height:46px; font-weight:750; border:1px solid var(--reset-ink); transition:transform .15s ease, box-shadow .15s ease; }
        .stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"] { background:var(--reset-accent); border-color:var(--reset-accent); color:#fff; }
        .stButton > button:hover, .stDownloadButton > button:hover { transform:translateY(-1px); box-shadow:0 10px 25px rgba(11,11,12,.12); }
        [data-baseweb="input"] > div, [data-baseweb="select"] > div, [data-baseweb="textarea"] > div { border-radius:13px !important; border-color:var(--reset-line) !important; background:#fff !important; }
        [data-testid="stDataFrame"] { border:1px solid var(--reset-line); border-radius:16px; overflow:hidden; background:#fff; }
        [data-testid="stAlert"] { border-radius:14px; }
        .reset-login-wrap { max-width:480px; margin:9vh auto 0; background:#fff; border:1px solid var(--reset-line); border-radius:28px; padding:32px; box-shadow:0 24px 80px rgba(11,11,12,.10); }
        .reset-login-logo { font-size:38px; font-weight:900; letter-spacing:-.07em; margin-bottom:22px; }
        .reset-login-logo span { color:var(--reset-accent); }
        .reset-login-title { font-size:26px; font-weight:800; letter-spacing:-.035em; }
        .reset-login-text { color:var(--reset-muted); margin:8px 0 18px; }
        @media (max-width:700px){ .block-container{padding:1rem .85rem 2rem}.reset-hero{padding:26px 22px;border-radius:22px}.reset-title{font-size:42px}.reset-subtitle{font-size:14px} }
        </style>
        """,
        unsafe_allow_html=True,
    )

apply_reset_brand()

def get_secret(name: str, default: str = "") -> str:
    """Streamlit Cloud secret, environment variable, or empty fallback."""
    try:
        value = st.secrets.get(name, default)
    except Exception:
        value = os.environ.get(name, default)
    return str(value or "").strip()

def get_first_secret(*names: str) -> tuple[str, str]:
    """Return the first configured secret and the name that matched, without exposing its value."""
    for name in names:
        value = get_secret(name)
        if value:
            return value, name
    return "", ""

def require_login() -> None:
    expected = get_secret("APP_PASSWORD")
    if not expected:
        return
    if st.session_state.get("reset_authenticated"):
        return
    st.markdown("""<div class="reset-login-wrap"><div class="reset-login-logo">RE<span>:</span>SET</div><div class="reset-login-title">Lead Intelligence</div><div class="reset-login-text">Reset İletişim özel çalışma alanı. Devam etmek için erişim şifresini gir.</div></div>""", unsafe_allow_html=True)
    password = st.text_input("Uygulama şifresi", type="password", key="login_password", placeholder="••••••••")
    if st.button("Giriş yap", type="primary", use_container_width=True):
        if hmac.compare_digest(password, expected):
            st.session_state["reset_authenticated"] = True
            st.rerun()
        else:
            st.error("Şifre yanlış.")
    st.stop()

require_login()

st.markdown(
    """
    <section class="reset-hero">
      <div class="reset-kicker">Reset İletişim · Private Intelligence Tool</div>
      <h1 class="reset-title">LEAD<span class="accent">:</span>FINDER</h1>
      <p class="reset-subtitle">Hedef şirketlerdeki karar vericileri bulur, kurumsal e-posta adaylarını doğrular ve kullanıma hazır CRM çıktısına dönüştürür.</p>
      <div class="reset-status">
        <span class="reset-chip"><span class="reset-dot"></span> Cloud online</span>
        <span class="reset-chip">People discovery</span>
        <span class="reset-chip">Email verification</span>
        <span class="reset-chip">Lusha fallback</span>
        <span class="reset-chip">CRM export</span>
      </div>
    </section>
    """,
    unsafe_allow_html=True,
)

stored_tavily_key, tavily_secret_name = get_first_secret("TAVILY_API_KEY", "TAVILY_KEY", "tavily_api_key")
stored_abstract_key, abstract_secret_name = get_first_secret("ABSTRACT_API_KEY", "ABSTRACT_KEY", "abstract_api_key")
stored_lusha_key, lusha_secret_name = get_first_secret("LUSHA_API_KEY", "LUSHA_KEY", "lusha_api_key")

# Always-visible connection status — independent of uploaded files and selected tab.
st.markdown("### API bağlantıları")
status_col1, status_col2, status_col3 = st.columns(3)
with status_col1:
    if stored_tavily_key:
        st.success("✓ Tavily API bağlı")
        st.caption(f"Secrets: `{tavily_secret_name}`")
    else:
        st.error("Tavily API bağlı değil")
        st.caption("Secrets içine `TAVILY_API_KEY` ekle.")
with status_col2:
    if stored_abstract_key:
        st.success("✓ Abstract API bağlı")
        st.caption(f"Secrets: `{abstract_secret_name}`")
    else:
        st.error("Abstract API bağlı değil")
        st.caption("Secrets içine `ABSTRACT_API_KEY` ekle.")
with status_col3:
    if stored_lusha_key:
        st.success("✓ Lusha API bağlı")
        st.caption(f"Secrets: `{lusha_secret_name}`")
    else:
        st.error("Lusha API bağlı değil")
        st.caption("Secrets içine `LUSHA_API_KEY` ekle ve uygulamayı Reboot et.")

with st.expander("Bağlantı tanılama", expanded=not bool(stored_lusha_key)):
    st.write({
        "Uygulama sürümü": "v10.1",
        "Tavily secret bulundu": bool(stored_tavily_key),
        "Abstract secret bulundu": bool(stored_abstract_key),
        "Lusha secret bulundu": bool(stored_lusha_key),
    })
    if not stored_lusha_key:
        st.warning("Lusha anahtarı okunamadı. Secret adı tam olarak `LUSHA_API_KEY` olmalı. Save sonrası Streamlit Cloud'da Reboot app yap.")
    else:
        st.info("Lusha anahtarı sunucuda bulundu. Anahtar değeri güvenlik nedeniyle ekranda gösterilmez.")

with st.sidebar:
    st.markdown("""<div class="reset-side-brand"><div class="reset-side-mark">RE<span>:</span>SET</div><div class="reset-side-caption">Lead Intelligence · v10.1</div></div>""", unsafe_allow_html=True)
    st.header("Bağlantılar")
    if stored_tavily_key:
        api_key = stored_tavily_key
        st.success("Tavily bağlı")
    else:
        api_key = st.text_input("Tavily API anahtarı", type="password", help="Sunucu secret tanımlı değilse bu oturum için girilir.")
        st.markdown("[Tavily paneli](https://app.tavily.com/)")
    if stored_abstract_key:
        abstract_key = stored_abstract_key
        st.success("Abstract Email Reputation bağlı")
    else:
        abstract_key = st.text_input("Abstract Email Reputation API anahtarı", type="password", help="Sunucu secret tanımlı değilse bu oturum için girilir.")
        st.markdown("[Abstract paneli](https://app.abstractapi.com/)")
    if stored_lusha_key:
        lusha_key = stored_lusha_key
        st.success("✓ Lusha API bağlı")
    else:
        lusha_key = st.text_input("Lusha API anahtarı — isteğe bağlı", type="password", help="Lusha API erişimin varsa gir. Abstract başarısız olduğunda devreye girer.")
        st.markdown("[Lusha API ayarları](https://dashboard.lusha.com/)")
    st.divider()
    st.caption("API anahtarları kaynak koduna yazılmaz. Bulut dağıtımında Secrets/Environment Variables bölümünde saklanır.")
    if get_secret("APP_PASSWORD") and st.button("Oturumu kapat", use_container_width=True):
        st.session_state.clear()
        st.rerun()

people_tab, email_tab = st.tabs(["1 · Kişileri Bul", "2 · E-postaları Bul"])

with people_tab:
    with st.expander("Tarama ayarları", expanded=True):
        selected_roles = st.multiselect("Hedef ekipler", list(ROLE_GROUPS.keys()), default=["Pazarlama & Marka", "Satın Alma", "Ticari Pazarlama & Shopper"])
        c1, c2, c3, c4 = st.columns(4)
        with c1: scan_mode = st.radio("Tarama modu", ["Hızlı", "Derin"], horizontal=True)
        with c2: results_per_query = st.slider("Sorgu başına sonuç", 3, 10, 5)
        with c3: max_per_company_role = st.slider("Firma/rol başına kişi", 1, 8, 4)
        with c4: request_delay = st.slider("Bekleme (sn)", 0.0, 1.5, 0.25, 0.05, key="people_delay")

    source_mode = st.radio("Şirket kaynağı", ["Hazır FMCG pilotu", "Kendi Excel/CSV dosyam"], horizontal=True)
    if source_mode == "Hazır FMCG pilotu":
        source_df = load_pilot(); company_column = "Firma"
    else:
        upload = st.file_uploader("Şirket Excel/CSV dosyası", type=["xlsx", "xls", "csv"], key="people_file")
        if upload is None:
            st.info("Şirket listesini yükle.")
            source_df = pd.DataFrame({"Firma": []}); company_column = "Firma"
        else:
            source_df, _ = read_input_file(upload)
            company_column = st.selectbox("Şirket sütunu", list(source_df.columns), key="people_company_col")
    companies = [c for c in dict.fromkeys(source_df.get(company_column, pd.Series(dtype=str)).dropna().astype(str).str.strip()) if c]
    if companies:
        c1, c2, c3 = st.columns(3)
        with c1: max_companies = st.number_input("Taranacak şirket", 1, len(companies), min(5, len(companies)), key="max_companies")
        mode_key = "quick" if scan_mode == "Hızlı" else "deep"
        estimated_requests = int(max_companies) * sum(len(ROLE_GROUPS[r][mode_key]) for r in selected_roles)
        with c2: st.metric("Tahmini Tavily kredisi", estimated_requests)
        with c3: st.metric("Listedeki şirket", len(companies))
        selected_companies = companies[:int(max_companies)]
        st.dataframe(pd.DataFrame({"Taranacak Firma": selected_companies}), use_container_width=True, hide_index=True)
        if st.button("Kişileri Araştır", type="primary", use_container_width=True):
            if not api_key.strip(): st.error("Tavily API anahtarını gir."); st.stop()
            all_hits: list[SearchHit] = []; errors: list[dict] = []
            progress = st.progress(0); status = st.empty(); completed = 0
            for company in selected_companies:
                for role_group in selected_roles:
                    for query_title in ROLE_GROUPS[role_group][mode_key]:
                        query = build_query(company, query_title); status.write(f"**{company}** · {role_group} · {query_title}")
                        try: all_hits.extend(extract_hits(company, role_group, query_title, query, tavily_search(api_key, query, results_per_query, ["linkedin.com"])))
                        except Exception as exc: errors.append({"Firma": company, "Rol Grubu": role_group, "Sorgu": query, "Hata": str(exc)})
                        completed += 1; progress.progress(min(1.0, completed / max(estimated_requests, 1)))
                        if request_delay: time.sleep(request_delay)
            results_df = hits_to_frame(dedupe_hits(all_hits, max_per_company_role))
            st.session_state["people_results"] = results_df; st.session_state["people_errors"] = pd.DataFrame(errors)
            status.success(f"Tarama tamamlandı: {len(results_df)} aday kişi bulundu.")
    if "people_results" in st.session_state:
        results_df = st.session_state["people_results"]
        st.dataframe(results_df[["Firma", "Rol Grubu", "Ad Soyad", "Unvan / Profil Başlığı", "Güven", "Güven Skoru", "LinkedIn"]], use_container_width=True, hide_index=True, column_config={"LinkedIn": st.column_config.LinkColumn("LinkedIn")})
        st.download_button("Kişi Excel’ini indir", data=excel_bytes({"Bulunan Kişiler": results_df}), file_name="reset_bulunan_kisiler.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

with email_tab:
    st.success("Hibrit mod: Tavily + Abstract ana akışıdır. İstersen Abstract başarısız olduğunda Lusha e-posta/telefon verisiyle tamamlar.")
    target_upload = st.file_uploader("1) Öncelikli kişiler dosyası", type=["xlsx", "xls", "csv"], key="target_email_file")
    previous_upload = st.file_uploader("2) Önceki e-posta sonuçları — isteğe bağlı", type=["xlsx", "xls", "csv"], key="previous_email_file")

    d1, d2 = st.columns(2)
    with d1:
        with open(SAMPLE_EMAIL_FILE, "rb") as f:
            st.download_button("Örnek 25 kişilik hedef dosya", data=f.read(), file_name="reset_oncelikli_25.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with d2:
        with open(SAMPLE_VERIFIED_FILE, "rb") as f:
            st.download_button("Örnek doğrulanmış 5 kişi", data=f.read(), file_name="reset_dogrulanmis_5.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    if target_upload is not None:
        email_df, _ = read_uploaded_sheet(target_upload, "target_sheet", ["E-posta Aşaması", "Öncelikli 25"])
        previous_df = pd.DataFrame()
        if previous_upload is not None:
            previous_df, _ = read_uploaded_sheet(previous_upload, "previous_sheet", ["E-posta Sonuçları", "CRM"])

        st.subheader("Lusha hibrit katmanı")
        lusha_options = ["Kapalı", "Sadece Abstract başarısız veya belirsizse", "Tüm seçili kişilerde çapraz kontrol"]
        default_lusha_index = 1 if lusha_key.strip() else 0
        lusha_choice = st.radio("Lusha kullanım biçimi", lusha_options, index=default_lusha_index, horizontal=True, key="lusha_mode")
        lusha_policy = {
            "Kapalı": "off",
            "Sadece Abstract başarısız veya belirsizse": "fallback",
            "Tüm seçili kişilerde çapraz kontrol": "all",
        }[lusha_choice]
        lc1, lc2 = st.columns(2)
        with lc1:
            lusha_reveal_phone = st.toggle("Lusha'dan telefon verisini de getir", value=False, disabled=lusha_policy == "off", key="lusha_phone")
        with lc2:
            lusha_cross_verify = st.toggle("Lusha e-postasını Abstract ile çapraz doğrula", value=True, disabled=lusha_policy == "off", key="lusha_cross")
        if lusha_policy != "off" and not lusha_key.strip():
            st.warning("Lusha modu seçili ancak LUSHA_API_KEY tanımlı değil. Streamlit Secrets'a ekle veya soldaki alana gir.")
        if lusha_reveal_phone:
            st.warning("Telefon açma daha fazla Lusha kredisi tüketebilir. Do Not Call işaretli numaralar önerilen telefon alanına alınmaz.")

        previous_std = standardize_previous_results(previous_df, lusha_policy)

        columns = list(email_df.columns)
        def idx(names: list[str], default: int = 0) -> int:
            normalized = [normalize_text(c) for c in columns]
            for name in names:
                nn = normalize_text(name)
                if nn in normalized: return normalized.index(nn)
            return min(default, len(columns) - 1)

        c1, c2, c3 = st.columns(3)
        with c1:
            company_col = st.selectbox("Firma sütunu", columns, index=idx(["Firma", "Şirket"]), key="resume_company")
            name_col = st.selectbox("Ad soyad sütunu", columns, index=idx(["Ad Soyad", "Kişi"]), key="resume_name")
        with c2:
            role_col = st.selectbox("Rol sütunu", ["— Yok —"] + columns, index=(idx(["Önerilen Rol", "Rol"]) + 1), key="resume_role")
            priority_col = st.selectbox("Öncelik sütunu", ["— Yok —"] + columns, index=(idx(["Öncelik"]) + 1), key="resume_priority")
        with c3:
            linkedin_col = st.selectbox("LinkedIn sütunu", ["— Yok —"] + columns, index=(idx(["LinkedIn"]) + 1), key="resume_linkedin")
            domain_col = st.selectbox("Hazır domain sütunu", ["— Yok —"] + columns, index=0, key="resume_domain")

        work = email_df.copy()
        work["__key"] = [record_key(c, n) for c, n in zip(work[company_col], work[name_col])]
        done_keys = set(previous_std.loc[previous_std.get("__done", False) == True, "__key"].astype(str)) if not previous_std.empty else set()
        work["__processed"] = work["__key"].isin(done_keys)

        if priority_col != "— Yok —":
            priorities = [x for x in work[priority_col].dropna().astype(str).unique() if x]
            selected_priorities = st.multiselect("İşlenecek öncelikler", sorted(priorities), default=sorted(priorities), key="resume_priorities")
            work = work[work[priority_col].astype(str).isin(selected_priorities)]
            order = {"A": 0, "B": 1, "C": 2}
            work["__p"] = work[priority_col].astype(str).map(order).fillna(9)
            work = work.sort_values(["__processed", "__p", company_col, name_col])
        else:
            work = work.sort_values(["__processed", company_col, name_col])

        remaining = work[~work["__processed"]].copy()
        already = int(work["__processed"].sum())
        m1, m2, m3 = st.columns(3)
        m1.metric("Hedef kişi", len(work))
        m2.metric("Önceden işlendi", already)
        m3.metric("Kalan", len(remaining))

        if remaining.empty:
            st.info("Seçili kişiler daha önce işlenmiş. Nihai CRM Excel’ini aşağıdan oluşturabilirsin.")
            batch = remaining
        else:
            batch_size = st.number_input("Bu turda işlenecek kalan kişi", 1, len(remaining), min(10, len(remaining)), key="resume_batch")
            max_verifications = st.slider("Kişi başına doğrulanacak aday", 1, 4, 2, key="resume_checks")
            delay = st.slider("Sorgular arası bekleme (sn)", 0.0, 1.5, 0.25, 0.05, key="resume_delay")
            batch = remaining.head(int(batch_size)).copy()
            unique_companies = batch[company_col].dropna().astype(str).nunique()
            e1, e2, e3 = st.columns(3)
            e1.metric("Tahmini Tavily kredisi", len(batch) + unique_companies)
            e2.metric("Azami Abstract doğrulaması", len(batch) * int(max_verifications))
            e3.metric("Azami Lusha kişisi", len(batch) if lusha_policy != "off" else 0)
            st.dataframe(batch.drop(columns=[c for c in ["__key", "__processed", "__p"] if c in batch.columns]), use_container_width=True, hide_index=True)

            if st.button("Hibrit araştırmayı başlat", type="primary", use_container_width=True):
                if not api_key.strip(): st.error("Tavily API anahtarını gir."); st.stop()
                if not abstract_key.strip(): st.error("Abstract Email Reputation API anahtarını gir."); st.stop()
                if lusha_policy != "off" and not lusha_key.strip(): st.error("Lusha modu için LUSHA_API_KEY gerekli."); st.stop()

                # Önceki sonuçlardaki şirket domainlerini yeni batch'e taşı; gereksiz Tavily domain sorgusunu azaltır.
                run_batch = batch.copy()
                resume_domain_col = None
                if not previous_std.empty and "Şirket Domaini" in previous_std.columns:
                    domain_map = previous_std.dropna(subset=["Şirket Domaini"]).drop_duplicates("Firma", keep="last").set_index("Firma")["Şirket Domaini"].to_dict()
                    run_batch["__Resume Domain"] = run_batch[company_col].map(domain_map).fillna("")
                    if run_batch["__Resume Domain"].astype(str).str.len().gt(0).any():
                        resume_domain_col = "__Resume Domain"

                results, errors, lusha_billing = enrich_emails(
                    api_key, abstract_key, run_batch, company_col, name_col,
                    None if role_col == "— Yok —" else role_col,
                    None if priority_col == "— Yok —" else priority_col,
                    None if linkedin_col == "— Yok —" else linkedin_col,
                    resume_domain_col if resume_domain_col else (None if domain_col == "— Yok —" else domain_col),
                    delay,
                    int(max_verifications),
                    lusha_key=lusha_key,
                    lusha_mode=lusha_policy,
                    lusha_reveal_phone=lusha_reveal_phone,
                    lusha_cross_verify=lusha_cross_verify,
                )
                st.session_state["resume_new_results"] = results
                st.session_state["resume_errors"] = errors
                st.session_state["lusha_billing"] = lusha_billing

        new_results = st.session_state.get("resume_new_results", pd.DataFrame())
        combined = merge_result_frames(previous_std, new_results)
        crm = build_crm_sheet(
            work, company_col, name_col,
            None if role_col == "— Yok —" else role_col,
            None if priority_col == "— Yok —" else priority_col,
            None if linkedin_col == "— Yok —" else linkedin_col,
            combined,
        )
        summary_df = build_summary(crm, combined)
        remaining_crm = crm[crm["CRM Durumu"] == "Araştırılacak"].copy()
        verified_crm = crm[crm["CRM Durumu"].isin(["İletişime hazır", "Telefon hazır"])].copy()

        if not new_results.empty:
            st.subheader("Bu turdaki sonuçlar")
            cols = [c for c in ["Firma", "Ad Soyad", "Rol", "Önerilen E-posta", "E-posta Durumu", "Önerilen Telefon", "Telefon Durumu", "Lusha Durumu", "Lusha Eşleşme", "Lusha Güncel Unvan", "Doğrulama Durumu", "SMTP Geçerli", "Güven", "LinkedIn"] if c in new_results.columns]
            st.dataframe(new_results[cols], use_container_width=True, hide_index=True, column_config={"LinkedIn": st.column_config.LinkColumn("LinkedIn")})

        st.subheader("Nihai CRM özeti")
        s1, s2, s3, s4 = st.columns(4)
        s1.metric("İletişime hazır", len(verified_crm))
        s2.metric("Araştırılacak", len(remaining_crm))
        s3.metric("Lusha başarılı", int((crm.get("Lusha Durumu", pd.Series(dtype=str)) == "Başarılı").sum()))
        s4.metric("Toplam hedef", len(crm))

        sheets = {
            "CRM": crm,
            "İletişime Hazır": verified_crm,
            "Kalanlar": remaining_crm,
            "Tüm Araştırma Sonuçları": combined.drop(columns=[c for c in ["__key", "__done"] if c in combined.columns], errors="ignore"),
            "Özet": summary_df,
        }
        errors = st.session_state.get("resume_errors", pd.DataFrame())
        if not errors.empty:
            sheets["Hatalar"] = errors
        lusha_billing = st.session_state.get("lusha_billing", pd.DataFrame())
        if not lusha_billing.empty:
            sheets["Lusha Billing"] = lusha_billing
        st.download_button(
            "Nihai CRM Excel’ini indir",
            data=excel_bytes(sheets),
            file_name="reset_lead_finder_nihai_crm.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )

st.divider()
st.caption("Araç LinkedIn hesabına giriş yapmaz. Açık web kaynakları, kurumsal e-posta doğrulaması ve isteğe bağlı Lusha API zenginleştirmesi kullanır. Do Not Call işaretli telefonları kullanmayın.")
