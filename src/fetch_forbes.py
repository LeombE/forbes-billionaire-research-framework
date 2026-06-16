"""Fetch official Forbes annual billionaire-list data with caching."""

from __future__ import annotations

import json
import time
import urllib.robotparser
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import requests

from .config import (
    DEFAULT_TARGET_YEAR,
    FORBES_API_LIMIT,
    FORBES_ROBOTS_URL,
    RAW_DIR,
    YearConfig,
    REQUEST_DELAY_SECONDS,
    REQUEST_TIMEOUT_SECONDS,
    USER_AGENT,
    annual_api_url,
    ensure_project_dirs,
    get_year_config,
)
from .manual_templates import write_manual_import_templates


class ForbesFetchError(RuntimeError):
    """Raised when canonical Forbes source data cannot be fetched."""


@dataclass(frozen=True)
class RobotsCheck:
    """Result of a robots.txt check."""

    robots_url: str
    annual_api_allowed: bool
    checked_url: str
    robots_path: Path
    note: str


def raw_json_path(year: int, raw_dir: Path = RAW_DIR) -> Path:
    """Return the raw JSON cache path for a Forbes annual-list year."""
    return raw_dir / f"forbes_billionaires_{year}.json"


def raw_error_path(year: int, raw_dir: Path = RAW_DIR) -> Path:
    """Return the error-log path for a failed Forbes annual-list year."""
    return raw_dir / f"forbes_billionaires_{year}_error.txt"


def year_raw_dir(year: int) -> Path:
    """Return the preferred year-specific raw directory for one annual source year."""
    return RAW_DIR / "forbes" / str(year)


def extract_records(payload: dict[str, Any]) -> list[dict[str, Any]]:
    """Extract the person records list from a Forbes annual-list payload."""
    try:
        records = payload["personList"]["personsLists"]
    except KeyError as exc:
        raise ForbesFetchError("Forbes payload does not contain personList.personsLists") from exc
    if not isinstance(records, list):
        raise ForbesFetchError("Forbes payload personList.personsLists is not a list")
    return records


class ForbesAnnualFetcher:
    """Small, conservative Forbes annual-list fetcher."""

    def __init__(
        self,
        user_agent: str = USER_AGENT,
        timeout_seconds: int = REQUEST_TIMEOUT_SECONDS,
        request_delay_seconds: float = REQUEST_DELAY_SECONDS,
        config: YearConfig | None = None,
    ) -> None:
        self.config = config or get_year_config(DEFAULT_TARGET_YEAR)
        self.raw_dir = self.config.raw_forbes_dir
        self.user_agent = user_agent
        self.timeout_seconds = timeout_seconds
        self.request_delay_seconds = request_delay_seconds
        self.session = requests.Session()
        self.session.headers.update({"User-Agent": self.user_agent, "Accept": "application/json,text/plain,*/*"})
        self._last_request_at = 0.0

    def _respect_rate_limit(self) -> None:
        elapsed = time.monotonic() - self._last_request_at
        if elapsed < self.request_delay_seconds:
            time.sleep(self.request_delay_seconds - elapsed)

    def check_robots(self, year: int | None = None) -> RobotsCheck:
        """Fetch robots.txt and check whether the annual API path is allowed."""
        ensure_project_dirs()
        self.raw_dir.mkdir(parents=True, exist_ok=True)
        year = int(year or self.config.year)
        response = self.session.get(FORBES_ROBOTS_URL, timeout=self.timeout_seconds)
        response.raise_for_status()
        robots_text = response.text
        robots_path = self.raw_dir / "forbes_robots.txt"
        robots_path.write_text(robots_text, encoding="utf-8")

        parser = urllib.robotparser.RobotFileParser()
        parser.set_url(FORBES_ROBOTS_URL)
        parser.parse(robots_text.splitlines())
        checked_url = annual_api_url(year, limit=100)
        allowed = parser.can_fetch(self.user_agent, checked_url)
        note = (
            "Forbes robots.txt was checked before fetching. The annual API path "
            "is not listed under disallowed paths for the generic user-agent rules."
            if allowed
            else "Forbes robots.txt does not allow this user agent to fetch the annual API path."
        )
        return RobotsCheck(
            robots_url=FORBES_ROBOTS_URL,
            annual_api_allowed=allowed,
            checked_url=checked_url,
            robots_path=robots_path,
            note=note,
        )

    def fetch_year(self, year: int, *, force: bool = False, limit: int = FORBES_API_LIMIT) -> dict[str, Any]:
        """Fetch one annual-list year, writing and reading a raw JSON cache."""
        ensure_project_dirs()
        raw_dir = year_raw_dir(year) if not self.config.legacy_layout else self.raw_dir
        raw_dir.mkdir(parents=True, exist_ok=True)
        cache_path = raw_json_path(year, raw_dir)
        if cache_path.exists() and not force:
            return json.loads(cache_path.read_text(encoding="utf-8"))

        url = annual_api_url(year, limit=limit)
        self._respect_rate_limit()
        self._last_request_at = time.monotonic()
        try:
            response = self.session.get(url, timeout=self.timeout_seconds)
            response.raise_for_status()
            payload = response.json()
            records = extract_records(payload)
            if not records:
                raise ForbesFetchError(f"No records returned for Forbes annual list {year}.")
        except Exception as exc:  # noqa: BLE001 - persisted diagnostics help reproducibility.
            raw_error_path(year, raw_dir).write_text(f"{type(exc).__name__}: {exc}\nURL: {url}\n", encoding="utf-8")
            raise ForbesFetchError(f"Could not fetch Forbes annual list {year}: {exc}") from exc

        cache_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        return payload

    def load_cached_year(self, year: int) -> dict[str, Any]:
        """Load one cached annual-list payload."""
        candidates = []
        if not self.config.legacy_layout:
            candidates.append(raw_json_path(year, year_raw_dir(year)))
        candidates.append(raw_json_path(year, self.raw_dir))
        candidates.append(raw_json_path(year, RAW_DIR))
        path = next((candidate for candidate in candidates if candidate.exists()), candidates[0])
        if not path.exists():
            raise FileNotFoundError(f"Missing cached Forbes annual data: {path}")
        return json.loads(path.read_text(encoding="utf-8"))

    def fetch_years(
        self,
        years: list[int],
        *,
        force: bool = False,
        limit: int = FORBES_API_LIMIT,
    ) -> tuple[dict[int, dict[str, Any]], list[str], RobotsCheck]:
        """Fetch many annual-list years, continuing past non-canonical failures."""
        write_manual_import_templates(self.config.year)
        robots_check = self.check_robots(self.config.year)
        if not robots_check.annual_api_allowed:
            raise ForbesFetchError(
                "Forbes robots.txt does not allow the annual API path for this user agent. "
                "Use the manual-import templates in data/raw/ instead."
            )

        payloads: dict[int, dict[str, Any]] = {}
        failures: list[str] = []
        for year in years:
            try:
                payloads[year] = self.fetch_year(year, force=force, limit=limit)
            except ForbesFetchError as exc:
                failures.append(str(exc))
        return payloads, failures, robots_check
