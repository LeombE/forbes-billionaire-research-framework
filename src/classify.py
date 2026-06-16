"""Rule-based wealth-engine classification."""

from __future__ import annotations

from dataclasses import dataclass

import pandas as pd

from .config import WEALTH_ENGINE_CATEGORIES


@dataclass(frozen=True)
class WealthEngineResult:
    category: str
    secondary_wealth_engines: str
    classification_confidence: str
    public_equity_dependency_flag: bool
    key_asset_or_company: str
    evidence_summary: str


PUBLIC_COMPANY_KEYWORDS = {
    "tesla",
    "meta",
    "facebook",
    "amazon",
    "oracle",
    "lvmh",
    "berkshire",
    "alphabet",
    "google",
    "microsoft",
    "nvidia",
    "dell",
    "walmart",
    "l'oreal",
    "loreal",
    "inditex",
    "zara",
    "reliance",
    "adani",
    "byd",
    "xiaomi",
    "pinduoduo",
    "pdd",
    "alibaba",
    "tencent",
    "netflix",
    "nike",
    "kweichow moutai",
    "equities",
    "stock",
    "shares",
}

PRIVATE_COMPANY_KEYWORDS = {
    "spacex",
    "xai",
    "bytedance",
    "tiktok",
    "chanel",
    "mars",
    "cargill",
    "bloomberg",
    "aldi",
    "quiktrip",
    "abc supply",
    "fidelity",
}

EARLY_EXECUTIVE_NAMES = {
    "steve ballmer",
    "eric schmidt",
    "macKenzie scott".lower(),
}

PLATFORM_KEYWORDS = {
    "facebook",
    "meta",
    "instagram",
    "whatsapp",
    "google",
    "alphabet",
    "amazon",
    "microsoft",
    "oracle",
    "tencent",
    "alibaba",
    "bytedance",
    "tiktok",
    "pinduoduo",
    "pdd",
    "booking",
    "mercadolibre",
}

LUXURY_RETAIL_KEYWORDS = {
    "fashion",
    "retail",
    "luxury",
    "lvmh",
    "louis vuitton",
    "sephora",
    "zara",
    "inditex",
    "chanel",
    "hermes",
    "l'oreal",
    "loreal",
    "cosmetics",
    "walmart",
    "nike",
    "aldi",
    "supermarkets",
}

REAL_ASSET_KEYWORDS = {
    "real estate",
    "land",
    "property",
    "infrastructure",
    "construction",
    "logistics",
    "ports",
    "rail",
    "telecom infrastructure",
}

RESOURCE_KEYWORDS = {
    "oil",
    "gas",
    "energy",
    "coal",
    "mining",
    "metals",
    "steel",
    "aluminum",
    "aluminium",
    "copper",
    "petrochemical",
    "petrochemicals",
    "commodities",
    "battery",
    "batteries",
    "resources",
    "cement",
}

INVESTOR_KEYWORDS = {
    "finance",
    "investments",
    "investment",
    "hedge fund",
    "private equity",
    "asset management",
    "banking",
    "capital",
    "berkshire",
    "trading",
}

DIVERSIFIED_KEYWORDS = {
    "diversified",
    "holding",
    "holdings",
    "conglomerate",
    "group",
}

INHERITED_KEYWORDS = {
    "inherited",
    "heir",
    "heiress",
    "family",
    "widow",
    "son of",
    "daughter",
    "grandchild",
}


def _text(row: pd.Series) -> str:
    parts = [
        row.get("name", ""),
        row.get("source_of_wealth", ""),
        row.get("industry", ""),
        row.get("primary_company_or_asset", ""),
        row.get("self_made_or_inherited_if_available", ""),
        row.get("_classification_text", ""),
    ]
    return " ".join(str(p).lower() for p in parts if pd.notna(p))


def _has_any(text: str, keywords: set[str]) -> bool:
    return any(keyword in text for keyword in keywords)


def _matched_keywords(text: str, keywords: set[str], limit: int = 4) -> list[str]:
    return [keyword for keyword in sorted(keywords) if keyword in text][:limit]


def public_equity_dependency(row: pd.Series) -> bool:
    """Infer whether the fortune is heavily exposed to public equity markets."""
    text = _text(row)
    if _has_any(text, PUBLIC_COMPANY_KEYWORDS):
        return True
    if _has_any(text, PRIVATE_COMPANY_KEYWORDS) and not _has_any(text, PUBLIC_COMPANY_KEYWORDS):
        return False
    industry = str(row.get("industry", "")).lower()
    if industry in {"finance & investments", "technology", "automotive"}:
        return True
    return False


def classify_wealth_engine(row: pd.Series) -> WealthEngineResult:
    """Classify a billionaire into exactly one wealth-engine category."""
    text = _text(row)
    name = str(row.get("name", "")).lower()
    asset = str(row.get("primary_company_or_asset") or row.get("source_of_wealth") or "").strip()
    self_made_label = str(row.get("self_made_or_inherited_if_available", "")).lower()
    public_flag = public_equity_dependency(row)

    matches: list[str] = []
    secondary: list[str] = []
    if "inherited" in self_made_label and _has_any(text, INHERITED_KEYWORDS | LUXURY_RETAIL_KEYWORDS | DIVERSIFIED_KEYWORDS):
        category = "Inherited/family-controlled business"
        matches = _matched_keywords(text, INHERITED_KEYWORDS | LUXURY_RETAIL_KEYWORDS | DIVERSIFIED_KEYWORDS)
    elif name in EARLY_EXECUTIVE_NAMES or "former ceo" in text or "early employee" in text:
        category = "Early employee/executive equity"
        matches = ["executive equity"]
    elif _has_any(text, INVESTOR_KEYWORDS):
        category = "Investor/capital allocator"
        matches = _matched_keywords(text, INVESTOR_KEYWORDS)
    elif _has_any(text, PLATFORM_KEYWORDS):
        category = "Technology/platform monopoly/network effects"
        matches = _matched_keywords(text, PLATFORM_KEYWORDS)
    elif public_flag:
        category = "Founder/operator public equity"
        matches = _matched_keywords(text, PUBLIC_COMPANY_KEYWORDS) or ["public-market-linked company"]
    elif _has_any(text, REAL_ASSET_KEYWORDS):
        category = "Real estate/land/infrastructure"
        matches = _matched_keywords(text, REAL_ASSET_KEYWORDS)
    elif _has_any(text, RESOURCE_KEYWORDS):
        category = "Commodities/energy/resources"
        matches = _matched_keywords(text, RESOURCE_KEYWORDS)
    elif _has_any(text, LUXURY_RETAIL_KEYWORDS):
        category = "Luxury/retail brand ownership"
        matches = _matched_keywords(text, LUXURY_RETAIL_KEYWORDS)
    elif _has_any(text, DIVERSIFIED_KEYWORDS):
        category = "Diversified holding company"
        matches = _matched_keywords(text, DIVERSIFIED_KEYWORDS)
    elif _has_any(text, PRIVATE_COMPANY_KEYWORDS):
        category = "Founder/operator private company"
        matches = _matched_keywords(text, PRIVATE_COMPANY_KEYWORDS)
    else:
        category = "Other/unclear"
        matches = []

    if category not in WEALTH_ENGINE_CATEGORIES:
        category = "Other/unclear"

    category_checks = [
        ("Investor/capital allocator", INVESTOR_KEYWORDS),
        ("Technology/platform monopoly/network effects", PLATFORM_KEYWORDS),
        ("Founder/operator public equity", PUBLIC_COMPANY_KEYWORDS),
        ("Founder/operator private company", PRIVATE_COMPANY_KEYWORDS),
        ("Inherited/family-controlled business", INHERITED_KEYWORDS),
        ("Luxury/retail brand ownership", LUXURY_RETAIL_KEYWORDS),
        ("Real estate/land/infrastructure", REAL_ASSET_KEYWORDS),
        ("Commodities/energy/resources", RESOURCE_KEYWORDS),
        ("Diversified holding company", DIVERSIFIED_KEYWORDS),
    ]
    for possible_category, keywords in category_checks:
        if possible_category != category and _has_any(text, keywords):
            secondary.append(possible_category)
    secondary_text = "; ".join(dict.fromkeys(secondary[:3]))

    confidence = "Low"
    if matches and asset and category != "Other/unclear":
        confidence = "Medium"
    if public_flag and matches and category in {
        "Founder/operator public equity",
        "Technology/platform monopoly/network effects",
        "Investor/capital allocator",
    }:
        confidence = "Medium"
    if "inherited" in self_made_label and category == "Inherited/family-controlled business":
        confidence = "Medium"

    year_value = row.get("year", 2025)
    try:
        source_year = int(year_value) if str(year_value).strip() and str(year_value).lower() != "nan" else 2025
    except (TypeError, ValueError):
        source_year = 2025
    evidence = (
        f"Forbes {source_year} source '{row.get('source_of_wealth', '')}', industry "
        f"'{row.get('industry', '')}', and key asset '{asset}'. "
        f"Primary rule signals: {', '.join(matches) if matches else 'no strong keyword match'}. "
        "Classification is a transparent heuristic pending primary ownership/filing review."
    )
    return WealthEngineResult(
        category=category,
        secondary_wealth_engines=secondary_text,
        classification_confidence=confidence,
        public_equity_dependency_flag=public_flag,
        key_asset_or_company=asset,
        evidence_summary=evidence,
    )
