#!/usr/bin/env python3
"""
Account Guardian — Automated Ads Account Audit & Change Forensics.

Phase 1: Google Ads audit modules 1-4 + health score + Excel workbook,
with the Module-5 change_event archive switched ON from the first run so
history accrues while Phase 2 (the before/after join) is built.

Every figure comes from the Google Ads API or from this tool's own archive
of previous API pulls. No manual input, ever. Read-only: never writes to
the ad account.

This single file is the whole project: the audit, the API probe, the unit
tests, the .env template, and a demo report generator.

Setup (once):
    pip install google-ads openpyxl requests
    python guardian.py --init             # writes a .env template next to this file
    python guardian.py --list-accounts    # confirm account config
    python guardian.py --probe Store1     # check the four API items flagged "probe first"

Daily:
    python guardian.py                    # audit all configured accounts
    python guardian.py --accounts Store1  # subset by name
    python guardian.py --skip-url-checks  # skip Module 3 HTTP checking
    python guardian.py --no-llm           # deterministic narrative only

Checks:
    python guardian.py --self-test        # built-in unit tests, no API needed
    python guardian.py --demo             # sample workbook from made-up data, no API needed
"""

import argparse
import concurrent.futures
import csv
import dataclasses
import datetime as dt
import json
import os
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

# ============================================================================
# SECTION 1 — CONFIG, ACCOUNTS, THRESHOLDS
# ============================================================================

ROOT = Path(__file__).resolve().parent
HISTORY_DIR = ROOT / "history"
REPORTS_DIR = ROOT / "reports"
ENV_FILE = ROOT / ".env"


def load_env_file(path):
    """Minimal .env reader (KEY=VALUE, # comments). Real environment wins."""
    if not path.exists():
        return
    for line in path.read_text(encoding="utf-8-sig").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        value = value.strip().strip('"').strip("'")
        os.environ.setdefault(key.strip(), value)


load_env_file(ENV_FILE)


def _env(name, default=None, cast=str):
    raw = os.getenv(name)
    if raw is None or raw == "":
        return default
    return cast(raw)


THRESHOLDS = {
    # Module 1
    "QS_MAX": _env("GUARDIAN_QS_MAX", 3, int),               # low-QS ceiling
    # Module 2 / 4
    "LOST_IS_MIN": _env("GUARDIAN_LOST_IS_MIN", 0.10, float),  # budget-limited floor
    # Module 3
    "SOFT404_MIN_BODY": _env("GUARDIAN_SOFT404_MIN_BODY", 512, int),
    "MAX_REDIRECTS": _env("GUARDIAN_MAX_REDIRECTS", 5, int),
    "URL_TIMEOUT": _env("GUARDIAN_URL_TIMEOUT", 10, int),
    "MAX_URLS": _env("GUARDIAN_MAX_URLS", 500, int),
    "URL_WORKERS": _env("GUARDIAN_URL_WORKERS", 16, int),
    # Windows
    "WINDOW_DAYS": _env("GUARDIAN_WINDOW_DAYS", 30, int),
    # Module 4
    "PACE_HIGH": _env("GUARDIAN_PACE_HIGH", 1.15, float),
    "PACE_LOW": _env("GUARDIAN_PACE_LOW", 0.70, float),
    "OVERFUNDED_ROAS_FRACTION": _env("GUARDIAN_OVERFUNDED_ROAS_FRACTION", 0.5, float),
    "MIN_SPEND_FOR_FINDING": _env("GUARDIAN_MIN_SPEND_FOR_FINDING", 5.0, float),
}

SEVERITY_ORDER = {"CRITICAL": 0, "HIGH": 1, "MEDIUM": 2, "LOW": 3}
SEVERITY_PENALTY = {"CRITICAL": 8.0, "HIGH": 5.0, "MEDIUM": 2.0, "LOW": 1.0}
MODULE_SCORE_CAP = {
    "1. Structural waste": 25.0,
    "2. Policy & delivery": 25.0,
    "3. Destination health": 30.0,
    "4. Budget & pacing": 20.0,
}


@dataclasses.dataclass
class Account:
    name: str
    customer_id: str            # digits only
    login_customer_id: str      # digits only ("" = use default MCC)


def load_accounts():
    """Parse ACCOUNTS=Name:customer_id[:login_customer_id],... from .env.

    Accounts under a different manager than the default MCC carry a third
    field — the per-account login-customer-id override.
    """
    raw = _env("ACCOUNTS", "")
    default_login = re.sub(r"\D", "", _env("GOOGLE_ADS_LOGIN_CUSTOMER_ID", "") or "")
    accounts = []
    for chunk in [c.strip() for c in raw.split(",") if c.strip()]:
        parts = chunk.split(":")
        if len(parts) not in (2, 3):
            raise SystemExit(f".env ACCOUNTS entry not Name:id[:login_id] — {chunk!r}")
        name = parts[0].strip()
        cid = re.sub(r"\D", "", parts[1])
        login = re.sub(r"\D", "", parts[2]) if len(parts) == 3 else default_login
        accounts.append(Account(name, cid, login))
    return accounts


@dataclasses.dataclass
class Finding:
    account: str
    module: str
    check: str
    severity: str          # CRITICAL / HIGH / MEDIUM / LOW
    entity: str            # campaign / ad group / keyword / URL affected
    detail: str
    monthly_cost: float | None   # priced waste in account currency, if computable
    fix: str


@dataclasses.dataclass
class CoverageRow:
    account: str
    source: str
    status: str            # "ok" | "gap"
    rows: int
    note: str


def micros(v):
    return (v or 0) / 1_000_000.0


# ============================================================================
# SECTION 2 — GOOGLE ADS CLIENT & QUERY HELPERS
# ============================================================================

def build_client(account: Account):
    """One client per account so login-customer-id overrides apply cleanly."""
    from google.ads.googleads.client import GoogleAdsClient  # lazy: tests run without it
    cfg = {
        "developer_token": _env("GOOGLE_ADS_DEVELOPER_TOKEN"),
        "client_id": _env("GOOGLE_ADS_CLIENT_ID"),
        "client_secret": _env("GOOGLE_ADS_CLIENT_SECRET"),
        "refresh_token": _env("GOOGLE_ADS_REFRESH_TOKEN"),
        "use_proto_plus": True,
    }
    if account.login_customer_id:
        cfg["login_customer_id"] = account.login_customer_id
    missing = [k for k, v in cfg.items() if v in (None, "") and k != "login_customer_id"]
    if missing:
        raise SystemExit(f"Missing Google Ads credentials in .env: {', '.join(missing)}")
    return GoogleAdsClient.load_from_dict(cfg)


def stream(client, customer_id, query):
    svc = client.get_service("GoogleAdsService")
    for batch in svc.search_stream(customer_id=customer_id, query=query):
        for row in batch.results:
            yield row


class Puller:
    """Runs GAQL pulls for one account, recording coverage (ok vs gap)."""

    def __init__(self, account: Account, coverage: list):
        self.account = account
        self.coverage = coverage
        self.client = build_client(account)

    def pull(self, source_name, query, mapper, gap_note=""):
        rows = []
        try:
            for row in stream(self.client, self.account.customer_id, query):
                rows.append(mapper(row))
            self.coverage.append(CoverageRow(self.account.name, source_name, "ok", len(rows), gap_note))
        except Exception as exc:  # any API failure is a stated gap, never zeros
            self.coverage.append(CoverageRow(
                self.account.name, source_name, "gap", 0,
                f"{type(exc).__name__}: {str(exc)[:300]} — fix credentials/access and re-run"))
            return None
        return rows


# ============================================================================
# SECTION 3 — DATA PULLS (GAQL)
# ============================================================================

def pull_account_data(account: Account, coverage: list, window_days: int):
    p = Puller(account, coverage)
    d = {"account": account}
    during = f"segments.date DURING LAST_{window_days}_DAYS" if window_days == 30 \
        else _between_clause(window_days)

    d["customer"] = p.pull("customer", """
        SELECT customer.id, customer.descriptive_name, customer.currency_code,
               customer.time_zone
        FROM customer""", lambda r: {
        "name": r.customer.descriptive_name,
        "currency": r.customer.currency_code,
        "tz": r.customer.time_zone})

    d["campaigns"] = p.pull("campaign + campaign_budget", """
        SELECT campaign.id, campaign.name, campaign.status,
               campaign.advertising_channel_type, campaign.bidding_strategy_type,
               campaign_budget.id, campaign_budget.amount_micros,
               campaign_budget.delivery_method, campaign_budget.explicitly_shared,
               campaign_budget.recommended_budget_amount_micros
        FROM campaign
        WHERE campaign.status != 'REMOVED'""", lambda r: {
        "id": r.campaign.id, "name": r.campaign.name,
        "status": r.campaign.status.name,
        "channel": r.campaign.advertising_channel_type.name,
        "bidding": r.campaign.bidding_strategy_type.name,
        "budget_id": r.campaign_budget.id,
        "budget": micros(r.campaign_budget.amount_micros),
        "budget_delivery": r.campaign_budget.delivery_method.name,
        "budget_shared": r.campaign_budget.explicitly_shared,
        "budget_recommended": micros(r.campaign_budget.recommended_budget_amount_micros)})

    d["campaign_daily"] = p.pull("campaign daily metrics", f"""
        SELECT campaign.id, campaign.name, campaign.status,
               campaign.advertising_channel_type, segments.date,
               metrics.cost_micros, metrics.clicks, metrics.impressions,
               metrics.conversions, metrics.conversions_value,
               metrics.search_budget_lost_impression_share
        FROM campaign
        WHERE {during} AND campaign.status != 'REMOVED'""", lambda r: {
        "id": r.campaign.id, "name": r.campaign.name,
        "channel": r.campaign.advertising_channel_type.name,
        "date": r.segments.date,
        "cost": micros(r.metrics.cost_micros),
        "clicks": r.metrics.clicks, "impressions": r.metrics.impressions,
        "conv": r.metrics.conversions, "conv_value": r.metrics.conversions_value,
        "lost_is": r.metrics.search_budget_lost_impression_share},
        gap_note="PMAX lost-IS 0.0 sentinel guarded downstream")

    d["keywords"] = p.pull("ad_group_criterion (keywords + ad-group negatives)", """
        SELECT campaign.id, campaign.name, ad_group.id, ad_group.name,
               ad_group.status, ad_group_criterion.criterion_id,
               ad_group_criterion.keyword.text, ad_group_criterion.keyword.match_type,
               ad_group_criterion.status, ad_group_criterion.negative,
               ad_group_criterion.quality_info.quality_score,
               ad_group_criterion.final_urls
        FROM ad_group_criterion
        WHERE ad_group_criterion.type = 'KEYWORD'""", lambda r: {
        "campaign_id": r.campaign.id, "campaign": r.campaign.name,
        "ad_group_id": r.ad_group.id, "ad_group": r.ad_group.name,
        "ad_group_status": r.ad_group.status.name,
        "criterion_id": r.ad_group_criterion.criterion_id,
        "text": r.ad_group_criterion.keyword.text,
        "match": r.ad_group_criterion.keyword.match_type.name,
        "status": r.ad_group_criterion.status.name,
        "negative": r.ad_group_criterion.negative,
        "qs": r.ad_group_criterion.quality_info.quality_score or None,
        "urls": list(r.ad_group_criterion.final_urls)})

    d["campaign_negatives"] = p.pull("campaign_criterion negatives", """
        SELECT campaign.id, campaign_criterion.keyword.text,
               campaign_criterion.keyword.match_type
        FROM campaign_criterion
        WHERE campaign_criterion.type = 'KEYWORD'
          AND campaign_criterion.negative = TRUE""", lambda r: {
        "campaign_id": r.campaign.id,
        "text": r.campaign_criterion.keyword.text,
        "match": r.campaign_criterion.keyword.match_type.name})

    d["shared_negatives"] = p.pull("shared_criterion (negative lists)", """
        SELECT shared_set.id, shared_set.name,
               shared_criterion.keyword.text, shared_criterion.keyword.match_type
        FROM shared_criterion
        WHERE shared_criterion.type = 'KEYWORD'""", lambda r: {
        "set_id": r.shared_set.id, "set_name": r.shared_set.name,
        "text": r.shared_criterion.keyword.text,
        "match": r.shared_criterion.keyword.match_type.name})

    d["campaign_shared_sets"] = p.pull("campaign_shared_set", """
        SELECT campaign.id, shared_set.id
        FROM campaign_shared_set
        WHERE campaign_shared_set.status = 'ENABLED'""", lambda r: {
        "campaign_id": r.campaign.id, "set_id": r.shared_set.id})

    d["keyword_metrics"] = p.pull("keyword_view metrics", f"""
        SELECT campaign.id, ad_group.id, ad_group_criterion.criterion_id,
               metrics.cost_micros, metrics.clicks,
               metrics.conversions, metrics.conversions_value
        FROM keyword_view
        WHERE {during}""", lambda r: {
        "ad_group_id": r.ad_group.id,
        "criterion_id": r.ad_group_criterion.criterion_id,
        "cost": micros(r.metrics.cost_micros), "clicks": r.metrics.clicks,
        "conv": r.metrics.conversions, "conv_value": r.metrics.conversions_value})

    d["ad_group_metrics"] = p.pull("ad_group metrics", f"""
        SELECT campaign.id, ad_group.id,
               metrics.cost_micros, metrics.conversions, metrics.conversions_value
        FROM ad_group
        WHERE {during}""", lambda r: {
        "ad_group_id": r.ad_group.id,
        "cost": micros(r.metrics.cost_micros),
        "conv": r.metrics.conversions, "conv_value": r.metrics.conversions_value})

    d["ads"] = p.pull("ad_group_ad (policy, strength, urls)", """
        SELECT campaign.id, campaign.name, ad_group.id, ad_group.name,
               ad_group.status, ad_group_ad.ad.id, ad_group_ad.ad.type,
               ad_group_ad.status, ad_group_ad.policy_summary.approval_status,
               ad_group_ad.policy_summary.review_status, ad_group_ad.ad_strength,
               ad_group_ad.ad.final_urls
        FROM ad_group_ad
        WHERE ad_group_ad.status != 'REMOVED'""", lambda r: {
        "campaign_id": r.campaign.id, "campaign": r.campaign.name,
        "ad_group_id": r.ad_group.id, "ad_group": r.ad_group.name,
        "ad_group_status": r.ad_group.status.name,
        "ad_id": r.ad_group_ad.ad.id, "ad_type": r.ad_group_ad.ad.type_.name,
        "status": r.ad_group_ad.status.name,
        "approval": r.ad_group_ad.policy_summary.approval_status.name,
        "review": r.ad_group_ad.policy_summary.review_status.name,
        "strength": r.ad_group_ad.ad_strength.name,
        "urls": list(r.ad_group_ad.ad.final_urls)})

    d["asset_groups"] = p.pull("asset_group (PMAX)", """
        SELECT campaign.id, campaign.name, asset_group.id, asset_group.name,
               asset_group.status, asset_group.final_urls, asset_group.ad_strength
        FROM asset_group""", lambda r: {
        "campaign_id": r.campaign.id, "campaign": r.campaign.name,
        "id": r.asset_group.id, "name": r.asset_group.name,
        "status": r.asset_group.status.name,
        "urls": list(r.asset_group.final_urls),
        "strength": r.asset_group.ad_strength.name},
        gap_note="probe: ad_strength coverage on current API version")

    d["asset_group_assets"] = p.pull("asset_group_asset policy", """
        SELECT campaign.id, asset_group.id, asset_group.name,
               asset_group_asset.field_type,
               asset_group_asset.policy_summary.approval_status,
               asset.id, asset.type
        FROM asset_group_asset""", lambda r: {
        "campaign_id": r.campaign.id,
        "asset_group_id": r.asset_group.id, "asset_group": r.asset_group.name,
        "field_type": r.asset_group_asset.field_type.name,
        "approval": r.asset_group_asset.policy_summary.approval_status.name,
        "asset_id": r.asset.id, "asset_type": r.asset.type_.name},
        gap_note="probe: policy field coverage varies by asset type")

    d["sitelinks"] = p.pull("asset (sitelinks)", """
        SELECT asset.id, asset.sitelink_asset.link_text, asset.final_urls,
               asset.policy_summary.approval_status
        FROM asset
        WHERE asset.type = 'SITELINK'""", lambda r: {
        "asset_id": r.asset.id, "text": r.asset.sitelink_asset.link_text,
        "urls": list(r.asset.final_urls),
        "approval": r.asset.policy_summary.approval_status.name})

    d["landing_pages"] = p.pull("landing_page_view (unexpanded)", f"""
        SELECT landing_page_view.unexpanded_final_url,
               campaign.id, campaign.status,
               metrics.cost_micros, metrics.clicks,
               metrics.conversions, metrics.conversions_value
        FROM landing_page_view
        WHERE {during} AND campaign.status IN ('ENABLED', 'PAUSED')""", lambda r: {
        "url": r.landing_page_view.unexpanded_final_url,
        "campaign_id": r.campaign.id,
        "cost": micros(r.metrics.cost_micros), "clicks": r.metrics.clicks,
        "conv": r.metrics.conversions, "conv_value": r.metrics.conversions_value},
        gap_note="sums reconciled against campaign totals (over-attribution guard)")

    d["recommendations"] = p.pull("recommendation", """
        SELECT recommendation.type, recommendation.resource_name,
               recommendation.campaign,
               recommendation.impact.base_metrics.cost_micros,
               recommendation.impact.potential_metrics.cost_micros,
               recommendation.impact.base_metrics.conversions,
               recommendation.impact.potential_metrics.conversions
        FROM recommendation""", lambda r: {
        "type": r.recommendation.type_.name,
        "campaign": r.recommendation.campaign,
        "base_cost": micros(r.recommendation.impact.base_metrics.cost_micros),
        "potential_cost": micros(r.recommendation.impact.potential_metrics.cost_micros),
        "base_conv": r.recommendation.impact.base_metrics.conversions,
        "potential_conv": r.recommendation.impact.potential_metrics.conversions},
        gap_note="probe: which types these accounts actually return")

    d["puller"] = p
    return d


def _between_clause(window_days):
    end = dt.date.today() - dt.timedelta(days=1)
    start = end - dt.timedelta(days=window_days - 1)
    return f"segments.date BETWEEN '{start}' AND '{end}'"


# ============================================================================
# SECTION 4 — VARIANT-MATCH / NEGATIVE-COLLISION ENGINE  (pure, tested)
# ============================================================================

_PUNCT_RE = re.compile(r"[^\w\s]", re.UNICODE)
_WS_RE = re.compile(r"\s+")


def normalize_kw(text):
    """Lowercase, strip broad-match '+' prefixes and punctuation, collapse spaces."""
    t = (text or "").lower()
    t = re.sub(r"(^|\s)\+", r"\1", t)          # legacy broad-match-modifier '+'
    t = t.replace("[", " ").replace("]", " ").replace('"', " ")
    t = _PUNCT_RE.sub(" ", t)
    return _WS_RE.sub(" ", t).strip()


def _stem(token):
    """Light close-variant stem: trailing 's' (not 'ss')."""
    if len(token) > 3 and token.endswith("s") and not token.endswith("ss"):
        return token[:-1]
    return token


def _tokens(text, close=False):
    toks = normalize_kw(text).split()
    return [_stem(t) for t in toks] if close else toks


def _contiguous(needle, haystack):
    n, h = len(needle), len(haystack)
    return any(haystack[i:i + n] == needle for i in range(h - n + 1)) if n else False


def negative_blocks(neg_text, neg_match, pos_text, close=False):
    """Does a negative keyword block a positive keyword's own text?

    Negative semantics (not positive close-variant semantics):
      EXACT  — blocks the identical query only
      PHRASE — blocks queries containing the phrase in order
      BROAD  — blocks queries containing every word, any order
    close=True additionally applies light singular/plural stemming, reported
    as a lower-confidence "close variant" collision.
    """
    n = _tokens(neg_text, close)
    p = _tokens(pos_text, close)
    if not n or not p:
        return False
    m = (neg_match or "BROAD").upper()
    if m == "EXACT":
        return n == p
    if m == "PHRASE":
        return _contiguous(n, p)
    return set(n) <= set(p)   # BROAD


def find_collisions(positives, negatives):
    """positives: [{text, match, key,...}]; negatives: [{text, match, scope,...}].
    Returns [(positive, negative, confidence)] with confidence 'literal'|'close'.
    """
    out = []
    for neg in negatives:
        for pos in positives:
            if negative_blocks(neg["text"], neg["match"], pos["text"], close=False):
                out.append((pos, neg, "literal"))
            elif negative_blocks(neg["text"], neg["match"], pos["text"], close=True):
                out.append((pos, neg, "close"))
    return out


# ============================================================================
# SECTION 5 — MODULE 1: STRUCTURAL WASTE
# ============================================================================

def run_module1(d, findings):
    acct = d["account"].name
    kws = d.get("keywords") or []
    if not kws:
        return
    kmetrics = {(m["ad_group_id"], m["criterion_id"]): m
                for m in (d.get("keyword_metrics") or [])}

    positives = [k for k in kws if not k["negative"] and k["status"] == "ENABLED"
                 and k["ad_group_status"] == "ENABLED"]
    ag_negatives = [dict(k, scope=f"ad group '{k['ad_group']}'")
                    for k in kws if k["negative"]]

    camp_names = {c["id"]: c["name"] for c in (d.get("campaigns") or [])}
    camp_negatives = [dict(n, scope=f"campaign '{camp_names.get(n['campaign_id'], n['campaign_id'])}'")
                      for n in (d.get("campaign_negatives") or [])]

    shared = d.get("shared_negatives") or []
    links = d.get("campaign_shared_sets") or []
    sets_by_campaign = defaultdict(set)
    for l in links:
        sets_by_campaign[l["campaign_id"]].add(l["set_id"])
    shared_by_set = defaultdict(list)
    for s in shared:
        shared_by_set[s["set_id"]].append(s)

    def kw_spend(k):
        m = kmetrics.get((k["ad_group_id"], k["criterion_id"]))
        return (m or {}).get("cost", 0.0), (m or {}).get("conv", 0.0), (m or {}).get("conv_value", 0.0)

    # --- Negative-vs-positive collisions -----------------------------------
    # Scope-aware: ad-group negatives against that ad group's positives,
    # campaign + attached shared-list negatives against that campaign's.
    pos_by_ag = defaultdict(list)
    pos_by_campaign = defaultdict(list)
    for k in positives:
        pos_by_ag[k["ad_group_id"]].append(k)
        pos_by_campaign[k["campaign_id"]].append(k)

    def emit_collision(pos, neg, confidence):
        cost, conv, value = kw_spend(pos)
        if cost < THRESHOLDS["MIN_SPEND_FOR_FINDING"] and conv == 0:
            return
        sev = "HIGH" if confidence == "literal" else "MEDIUM"
        findings.append(Finding(
            acct, "1. Structural waste", "Negative blocks active keyword", sev,
            f"{pos['campaign']} › {pos['ad_group']} › \"{pos['text']}\" ({pos['match']})",
            f"Blocked by negative \"{neg['text']}\" ({neg['match']}) in {neg['scope']}"
            f" [{confidence} match]. Trailing {THRESHOLDS['WINDOW_DAYS']}d:"
            f" {conv:.1f} conv, {value:,.0f} value.",
            cost,
            f"Remove or narrow the negative in {neg['scope']}, or exclude this "
            f"ad group from the list."))

    for ag_id, plist in pos_by_ag.items():
        negs = [n for n in ag_negatives if n["ad_group_id"] == ag_id]
        for pos, neg, conf in find_collisions(plist, negs):
            emit_collision(pos, neg, conf)

    for camp_id, plist in pos_by_campaign.items():
        negs = [n for n in camp_negatives if n["campaign_id"] == camp_id]
        for set_id in sets_by_campaign.get(camp_id, ()):
            negs += [dict(s, scope=f"shared list '{s['set_name']}'")
                     for s in shared_by_set.get(set_id, ())]
        for pos, neg, conf in find_collisions(plist, negs):
            emit_collision(pos, neg, conf)

    # --- Ad groups with zero enabled ads / zero enabled keywords ------------
    ads = d.get("ads") or []
    ag_spend = {m["ad_group_id"]: m for m in (d.get("ad_group_metrics") or [])}
    enabled_ads_by_ag = defaultdict(int)
    ag_info = {}
    for a in ads:
        ag_info[a["ad_group_id"]] = a
        if a["status"] == "ENABLED":
            enabled_ads_by_ag[a["ad_group_id"]] += 1
    enabled_kws_by_ag = defaultdict(int)
    search_ags = set()
    for k in kws:
        search_ags.add(k["ad_group_id"])
        if not k["negative"] and k["status"] == "ENABLED":
            enabled_kws_by_ag[k["ad_group_id"]] += 1

    seen_ags = {a["ad_group_id"] for a in ads if a["ad_group_status"] == "ENABLED"} | \
               {k["ad_group_id"] for k in kws if k["ad_group_status"] == "ENABLED"}
    for ag_id in seen_ags:
        info = ag_info.get(ag_id) or next(
            (k for k in kws if k["ad_group_id"] == ag_id), None)
        if not info:
            continue
        name = f"{info['campaign']} › {info['ad_group']}"
        spend = ag_spend.get(ag_id, {}).get("cost", 0.0)
        if enabled_ads_by_ag.get(ag_id, 0) == 0:
            findings.append(Finding(
                acct, "1. Structural waste", "Ad group has zero enabled ads",
                "HIGH" if spend > 0 else "MEDIUM", name,
                f"Enabled ad group with no enabled ads — it has silently stopped. "
                f"Trailing spend now orphaned.",
                spend or None, "Add or re-enable an ad, or pause the ad group."))
        if ag_id in search_ags and enabled_kws_by_ag.get(ag_id, 0) == 0:
            findings.append(Finding(
                acct, "1. Structural waste", "Ad group has zero enabled keywords",
                "MEDIUM", name,
                "Enabled search ad group with no enabled keywords.",
                spend or None, "Add keywords or pause the ad group."))

    # --- Low quality score still spending ------------------------------------
    total_cost = sum(m["cost"] for m in kmetrics.values())
    total_clicks = sum(m["clicks"] for m in kmetrics.values())
    acct_cpc = total_cost / total_clicks if total_clicks else 0.0
    qs_null = 0
    for k in positives:
        if k["qs"] is None:
            qs_null += 1
            continue
        if k["qs"] <= THRESHOLDS["QS_MAX"]:
            cost, conv, _ = kw_spend(k)
            m = kmetrics.get((k["ad_group_id"], k["criterion_id"])) or {}
            clicks = m.get("clicks", 0)
            kw_cpc = cost / clicks if clicks else 0.0
            if cost < THRESHOLDS["MIN_SPEND_FOR_FINDING"]:
                continue
            premium = max(0.0, cost - acct_cpc * clicks) if kw_cpc > acct_cpc else 0.0
            findings.append(Finding(
                acct, "1. Structural waste", f"QS ≤ {THRESHOLDS['QS_MAX']} still spending",
                "MEDIUM",
                f"{k['campaign']} › {k['ad_group']} › \"{k['text']}\"",
                f"QS {k['qs']}, {cost:,.2f} spend at CPC {kw_cpc:,.2f} vs account {acct_cpc:,.2f}.",
                premium if premium > 0 else None,
                "Improve ad relevance / landing page, or pause and reroute the spend."))
    if qs_null:
        d["puller"].coverage.append(CoverageRow(
            acct, "quality_score", "gap", qs_null,
            f"{qs_null} keywords return NULL quality score (normal for low-volume/"
            f"non-search) — stated as gap, not scored as zero"))

    # --- Duplicate keywords across ad groups ---------------------------------
    groups = defaultdict(list)
    for k in positives:
        groups[(normalize_kw(k["text"]), k["match"])].append(k)
    for (text, match), members in groups.items():
        ags = {m["ad_group_id"] for m in members}
        if len(ags) < 2:
            continue
        scored = sorted(members, key=lambda k: kw_spend(k)[1], reverse=True)
        winner, losers = scored[0], scored[1:]
        loser_cost = sum(kw_spend(l)[0] for l in losers)
        if loser_cost < THRESHOLDS["MIN_SPEND_FOR_FINDING"]:
            continue
        findings.append(Finding(
            acct, "1. Structural waste", "Duplicate keyword competes across ad groups",
            "MEDIUM", f"\"{text}\" ({match}) in {len(ags)} ad groups",
            f"Best performer: {winner['campaign']} › {winner['ad_group']}. "
            f"{len(losers)} duplicates splitting traffic.",
            loser_cost, "Keep the best-performing instance; pause the duplicates."))


# ============================================================================
# SECTION 6 — MODULE 2: POLICY & DELIVERY HEALTH
# ============================================================================

def run_module2(d, findings):
    acct = d["account"].name
    ads = d.get("ads") or []
    ag_spend = {m["ad_group_id"]: m for m in (d.get("ad_group_metrics") or [])}

    # Disapproved / limited ads; and ad groups whose LAST live ad just died.
    by_ag = defaultdict(list)
    for a in ads:
        by_ag[a["ad_group_id"]].append(a)
    for ag_id, group in by_ag.items():
        sample = group[0]
        if sample["ad_group_status"] != "ENABLED":
            continue
        name = f"{sample['campaign']} › {sample['ad_group']}"
        enabled = [a for a in group if a["status"] == "ENABLED"]
        bad = [a for a in enabled if a["approval"] in ("DISAPPROVED", "AREA_OF_INTEREST_ONLY")]
        limited = [a for a in enabled if a["approval"] == "APPROVED_LIMITED"]
        servable = [a for a in enabled if a["approval"] not in
                    ("DISAPPROVED", "AREA_OF_INTEREST_ONLY")]
        spend = ag_spend.get(ag_id, {}).get("cost", 0.0)
        if enabled and not servable:
            findings.append(Finding(
                acct, "2. Policy & delivery", "Last live ad in ad group disapproved",
                "CRITICAL", name,
                f"All {len(enabled)} enabled ads disapproved — the ad group has "
                f"stopped with no alarm.",
                spend or None, "Fix the policy issue or add a compliant ad now."))
        else:
            for a in bad:
                findings.append(Finding(
                    acct, "2. Policy & delivery", "Disapproved ad", "HIGH",
                    f"{name} › ad {a['ad_id']}",
                    f"{a['ad_type']} approval={a['approval']} review={a['review']}.",
                    None, "Fix or replace the ad; appeal if wrongly flagged."))
            for a in limited:
                findings.append(Finding(
                    acct, "2. Policy & delivery", "Ad approved (limited)", "MEDIUM",
                    f"{name} › ad {a['ad_id']}",
                    f"{a['ad_type']} is serving restricted (APPROVED_LIMITED).",
                    None, "Review the policy topic; adjust creative or targeting."))

        # Ad strength Poor/Average on spending ad groups
        weak = [a for a in servable if a["strength"] in ("POOR", "AVERAGE")]
        if weak and spend > THRESHOLDS["MIN_SPEND_FOR_FINDING"]:
            worst = "POOR" if any(a["strength"] == "POOR" for a in weak) else "AVERAGE"
            findings.append(Finding(
                acct, "2. Policy & delivery", "Weak ad strength on spending ad group",
                "MEDIUM" if worst == "POOR" else "LOW", name,
                f"{len(weak)} enabled ad(s) rated {worst} while spending "
                f"{spend:,.2f} in the window.",
                None, "Add headlines/descriptions & assets to lift ad strength."))

    # PMAX asset groups: policy + missing/limited assets + ad strength
    ag_assets = d.get("asset_group_assets")
    if ag_assets is not None:
        by_group = defaultdict(list)
        for a in ag_assets:
            by_group[a["asset_group_id"]].append(a)
        for gid, assets in by_group.items():
            sample = assets[0]
            bad = [a for a in assets if a["approval"] in ("DISAPPROVED",)]
            if bad:
                findings.append(Finding(
                    acct, "2. Policy & delivery", "Disapproved PMAX asset(s)", "HIGH",
                    f"asset group '{sample['asset_group']}'",
                    f"{len(bad)} asset(s) disapproved "
                    f"({', '.join(sorted({a['field_type'] for a in bad}))}).",
                    None, "Replace the disapproved assets."))
    for g in (d.get("asset_groups") or []):
        if g["status"] == "ENABLED" and g["strength"] in ("POOR", "AVERAGE"):
            findings.append(Finding(
                acct, "2. Policy & delivery", "PMAX asset group strength weak",
                "MEDIUM" if g["strength"] == "POOR" else "LOW",
                f"{g['campaign']} › {g['name']}",
                f"asset_group.ad_strength = {g['strength']}.",
                None, "Fill missing asset types; add video/image variety."))

    for s in (d.get("sitelinks") or []):
        if s["approval"] == "DISAPPROVED":
            findings.append(Finding(
                acct, "2. Policy & delivery", "Disapproved sitelink", "MEDIUM",
                f"sitelink \"{s['text']}\"", "Sitelink asset disapproved.",
                None, "Fix destination/text and resubmit."))


# ============================================================================
# SECTION 7 — MODULE 3: DESTINATION HEALTH (the dead-click detector)
# ============================================================================

SOFT404_PATTERNS = [
    "page not found", "page cannot be found", "404", "no longer available",
    "nothing was found", "we couldn't find", "we could not find",
    "this product is no longer", "does not exist",
]
OOS_PATTERNS = [
    "out of stock", "sold out", "currently unavailable", "outofstock",
    "no longer carried", "discontinued item",
]


def classify_page(status, requested_url, final_url, body, title=""):
    """Pure classifier → (verdict, reason). verdict: ok|broken|soft404|oos|suspect."""
    if status is None:
        return "broken", "request failed"
    if status >= 400:
        return "broken", f"HTTP {status}"
    text = (body or "")[:200_000].lower()
    ttl = (title or "").lower()

    def _path(u):
        m = re.match(r"https?://[^/]+(/[^?#]*)?", u or "")
        return (m.group(1) or "/") if m else "/"
    if _path(requested_url) not in ("/", "") and _path(final_url) in ("/", ""):
        return "soft404", "redirects to homepage"
    for pat in SOFT404_PATTERNS:
        if pat in ttl or (pat != "404" and pat in text):
            return "soft404", f"page says '{pat}'"
    for pat in OOS_PATTERNS:
        if pat in text:
            return "oos", f"page says '{pat}'"
    if body is not None and len(body) < THRESHOLDS["SOFT404_MIN_BODY"]:
        return "suspect", f"thin page ({len(body)} bytes)"
    return "ok", ""


def check_url(url):
    """Follow redirects manually to record the chain; GET body for heuristics."""
    import requests
    session = requests.Session()
    session.headers["User-Agent"] = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                                     "AccountGuardian/1.0 destination-health-check")
    chain = []
    current = url
    try:
        for _ in range(THRESHOLDS["MAX_REDIRECTS"] + 1):
            resp = session.get(current, timeout=THRESHOLDS["URL_TIMEOUT"],
                               allow_redirects=False, stream=True)
            chain.append((current, resp.status_code))
            if resp.status_code in (301, 302, 303, 307, 308):
                nxt = resp.headers.get("Location")
                if not nxt:
                    break
                current = requests.compat.urljoin(current, nxt)
                continue
            body = resp.raw.read(200_000, decode_content=True) if resp.ok else b""
            body_text = body.decode("utf-8", "ignore")
            m = re.search(r"<title[^>]*>(.*?)</title>", body_text, re.I | re.S)
            title = m.group(1).strip() if m else ""
            verdict, reason = classify_page(resp.status_code, url, current,
                                            body_text, title)
            if len(chain) > 3 and verdict == "ok":
                verdict, reason = "suspect", f"redirect chain of {len(chain) - 1}"
            return {"url": url, "final_url": current, "status": resp.status_code,
                    "chain": len(chain) - 1, "verdict": verdict, "reason": reason}
        return {"url": url, "final_url": current, "status": None,
                "chain": len(chain) - 1, "verdict": "broken",
                "reason": f"redirect chain exceeded {THRESHOLDS['MAX_REDIRECTS']}"}
    except requests.RequestException as exc:
        return {"url": url, "final_url": current, "status": None, "chain": len(chain),
                "verdict": "broken", "reason": type(exc).__name__}


def reconcile_landing_spend(lp_rows, campaign_totals):
    """Guard against landing_page_view over-attribution:
    scale each campaign's URL costs down so they never exceed the campaign total."""
    by_campaign = defaultdict(float)
    for r in lp_rows:
        by_campaign[r["campaign_id"]] += r["cost"]
    scale = {}
    for cid, lp_sum in by_campaign.items():
        total = campaign_totals.get(cid)
        scale[cid] = min(1.0, total / lp_sum) if total is not None and lp_sum > 0 else 1.0
    out = defaultdict(float)
    for r in lp_rows:
        out[_url_key(r["url"])] += r["cost"] * scale[r["campaign_id"]]
    return dict(out), scale


def _url_key(url):
    return re.sub(r"[?#].*$", "", (url or "").strip().lower()).rstrip("/")


def run_module3(d, findings, skip_checks, max_urls, url_results_out, checker=None):
    checker = checker or check_url
    acct = d["account"].name

    urls = set()
    for a in (d.get("ads") or []):
        if a["status"] == "ENABLED":
            urls.update(a["urls"])
    for k in (d.get("keywords") or []):
        if not k["negative"] and k["status"] == "ENABLED":
            urls.update(k["urls"])
    for s in (d.get("sitelinks") or []):
        urls.update(s["urls"])
    for g in (d.get("asset_groups") or []):
        if g["status"] == "ENABLED":
            urls.update(g["urls"])
    urls = {u for u in urls if u.startswith("http")}

    lp_rows = d.get("landing_pages") or []
    campaign_totals = defaultdict(float)
    for r in (d.get("campaign_daily") or []):
        campaign_totals[r["id"]] += r["cost"]
    spend_by_url, _ = reconcile_landing_spend(lp_rows, dict(campaign_totals))

    ranked = sorted(urls, key=lambda u: spend_by_url.get(_url_key(u), 0.0),
                    reverse=True)[:max_urls]
    if skip_checks or not ranked:
        if ranked:
            d["puller"].coverage.append(CoverageRow(
                acct, "destination checks", "gap", 0,
                f"--skip-url-checks: {len(ranked)} URLs not probed this run"))
        return

    with concurrent.futures.ThreadPoolExecutor(THRESHOLDS["URL_WORKERS"]) as pool:
        results = list(pool.map(checker, ranked))

    sev_map = {"broken": "CRITICAL", "soft404": "HIGH", "oos": "HIGH", "suspect": "LOW"}
    for r in results:
        r["account"] = acct
        r["spend"] = spend_by_url.get(_url_key(r["url"]), 0.0)
        url_results_out.append(r)
        if r["verdict"] == "ok":
            continue
        spend = r["spend"]
        sev = sev_map[r["verdict"]]
        if r["verdict"] == "suspect" and spend < THRESHOLDS["MIN_SPEND_FOR_FINDING"]:
            continue
        findings.append(Finding(
            acct, "3. Destination health",
            {"broken": "Paid clicks sent to broken URL",
             "soft404": "Paid clicks sent to soft-404",
             "oos": "Paid clicks sent to out-of-stock page",
             "suspect": "Suspect destination"}[r["verdict"]],
            sev, r["url"],
            f"{r['reason']} (final: {r['final_url']}, status {r['status']}, "
            f"{r['chain']} redirects). Spend last {THRESHOLDS['WINDOW_DAYS']}d shown.",
            spend if spend > 0 else None,
            "Fix/redirect the page, or update the ad/asset final URL."))

    if len(urls) > max_urls:
        d["puller"].coverage.append(CoverageRow(
            acct, "destination checks", "gap", len(ranked),
            f"{len(urls) - max_urls} lower-spend URLs beyond --max-urls cap not "
            f"checked this run"))


# ============================================================================
# SECTION 8 — MODULE 4: BUDGET & PACING  (pure math helpers are tested)
# ============================================================================

def month_pacing(mtd_spend, days_elapsed, days_in_month, monthly_budget):
    """Return (run_rate, projected_month_spend, pace_ratio vs budget)."""
    run_rate = mtd_spend / days_elapsed if days_elapsed else 0.0
    projected = run_rate * days_in_month
    ratio = projected / monthly_budget if monthly_budget else None
    return run_rate, projected, ratio


def run_module4(d, findings, today=None):
    acct = d["account"].name
    campaigns = d.get("campaigns") or []
    daily = d.get("campaign_daily") or []
    if not campaigns:
        return

    today = today or dt.date.today()
    days_in_month = (dt.date(today.year + (today.month == 12),
                             (today.month % 12) + 1, 1) - dt.timedelta(days=1)).day
    month_prefix = today.strftime("%Y-%m")
    days_elapsed = max(1, today.day - 1)   # through yesterday

    by_campaign = defaultdict(lambda: {"cost": 0.0, "conv_value": 0.0, "conv": 0.0,
                                       "mtd": 0.0, "lost_is": [], "channel": ""})
    for r in daily:
        c = by_campaign[r["id"]]
        c["cost"] += r["cost"]
        c["conv_value"] += r["conv_value"]
        c["conv"] += r["conv"]
        c["channel"] = r["channel"]
        if r["date"].startswith(month_prefix) and r["date"] != str(today):
            c["mtd"] += r["cost"]
        # PMAX 0.0-sentinel guard: lost-IS on PMAX is unreliable, treat as unknown
        if r["channel"] != "PERFORMANCE_MAX" and r["lost_is"] is not None:
            c["lost_is"].append(r["lost_is"])

    total_cost = sum(c["cost"] for c in by_campaign.values())
    total_value = sum(c["conv_value"] for c in by_campaign.values())
    acct_roas = total_value / total_cost if total_cost else 0.0

    for camp in campaigns:
        if camp["status"] != "ENABLED" or not camp["budget"]:
            continue
        perf = by_campaign.get(camp["id"])
        if not perf:
            continue
        monthly_budget = camp["budget"] * days_in_month
        run_rate, projected, ratio = month_pacing(
            perf["mtd"], days_elapsed, days_in_month, monthly_budget)
        name = camp["name"]
        roas = perf["conv_value"] / perf["cost"] if perf["cost"] else 0.0
        avg_lost = (sum(perf["lost_is"]) / len(perf["lost_is"])) if perf["lost_is"] else None

        if ratio is not None and ratio > THRESHOLDS["PACE_HIGH"]:
            findings.append(Finding(
                acct, "4. Budget & pacing", "Pacing over budget", "MEDIUM", name,
                f"MTD {perf['mtd']:,.2f}; projected {projected:,.2f} vs monthly "
                f"budget ≈{monthly_budget:,.2f} ({ratio:.0%}).",
                max(0.0, projected - monthly_budget),
                "Trim daily budget or accept the overspend deliberately."))
        elif ratio is not None and ratio < THRESHOLDS["PACE_LOW"] and perf["mtd"] > 0:
            findings.append(Finding(
                acct, "4. Budget & pacing", "Pacing far under budget", "LOW", name,
                f"MTD {perf['mtd']:,.2f}; projected {projected:,.2f} vs monthly "
                f"budget ≈{monthly_budget:,.2f} ({ratio:.0%}).",
                None, "Reallocate the idle budget or lower it."))

        if avg_lost is not None and avg_lost >= THRESHOLDS["LOST_IS_MIN"] \
                and roas >= acct_roas and perf["cost"] > 0:
            headroom_value = perf["cost"] * (avg_lost / (1 - avg_lost)) * roas \
                if avg_lost < 1 else 0.0
            findings.append(Finding(
                acct, "4. Budget & pacing", "Budget-limited winner", "HIGH", name,
                f"Losing {avg_lost:.0%} impression share to budget at ROAS "
                f"{roas:,.2f} (account avg {acct_roas:,.2f}). Est. forgone conv "
                f"value ≈{headroom_value:,.0f}/window.",
                None, "Raise this budget; fund it from an overfunded loser below."))
        if avg_lost is not None and avg_lost < 0.02 and perf["cost"] > 0 \
                and acct_roas > 0 and roas < acct_roas * THRESHOLDS["OVERFUNDED_ROAS_FRACTION"]:
            findings.append(Finding(
                acct, "4. Budget & pacing", "Overfunded loser", "MEDIUM", name,
                f"Unconstrained by budget, yet ROAS {roas:,.2f} is under "
                f"{THRESHOLDS['OVERFUNDED_ROAS_FRACTION']:.0%} of account avg "
                f"{acct_roas:,.2f}. Spend {perf['cost']:,.2f} in window.",
                None, "Cut budget/bids here; move funds to the limited winners."))

    for rec in (d.get("recommendations") or []):
        delta_cost = rec["potential_cost"] - rec["base_cost"]
        delta_conv = rec["potential_conv"] - rec["base_conv"]
        findings.append(Finding(
            acct, "4. Budget & pacing", f"Google recommendation: {rec['type']}",
            "LOW", rec["campaign"] or "(account)",
            f"Predicted impact: Δcost {delta_cost:+,.2f}, Δconv {delta_conv:+,.1f}. "
            f"Reported deterministically — never auto-applied.",
            None, "Review in the UI; apply only if it matches strategy."))


# ============================================================================
# SECTION 9 — MODULE 5 (PHASE-1 SLICE): CHANGE-EVENT ARCHIVE
# ============================================================================
# The API only reaches back 30 days; requires change_date_time predicate and
# LIMIT (max 10,000 rows). We page by day (splitting a day into 4 chunks when
# it hits the cap) and archive to history/change_events/<cid>/<date>.jsonl.
# Daily runs make the 30-day window an unlimited local archive — the point.

CHANGE_DIR = HISTORY_DIR / "change_events"

CHANGE_QUERY = """
    SELECT change_event.change_date_time, change_event.user_email,
           change_event.client_type, change_event.change_resource_type,
           change_event.change_resource_name, change_event.resource_change_operation,
           change_event.changed_fields, change_event.campaign, change_event.ad_group
    FROM change_event
    WHERE change_event.change_date_time BETWEEN '{start}' AND '{end}'
    ORDER BY change_event.change_date_time
    LIMIT 10000
"""


def _change_row(r):
    ce = r.change_event
    return {
        "time": ce.change_date_time,
        "user": ce.user_email,
        "client": ce.client_type.name,
        "resource_type": ce.change_resource_type.name,
        "resource": ce.change_resource_name,
        "operation": ce.resource_change_operation.name,
        "fields": list(ce.changed_fields.paths),
        "campaign": ce.campaign,
        "ad_group": ce.ad_group,
        "diff": _extract_diff(ce),
    }


def _extract_diff(ce):
    """Best-effort old→new values for each changed field path."""
    type_attr = ce.change_resource_type.name.lower()
    diff = {}
    try:
        old_msg = getattr(ce.old_resource, type_attr, None)
        new_msg = getattr(ce.new_resource, type_attr, None)
        for path in list(ce.changed_fields.paths)[:25]:
            diff[path] = {"old": _walk(old_msg, path), "new": _walk(new_msg, path)}
    except Exception:
        pass
    return diff


def _walk(msg, path):
    cur = msg
    try:
        for part in path.split("."):
            if cur is None:
                return None
            cur = getattr(cur, part)
        if hasattr(cur, "name"):        # enum
            cur = cur.name
        s = str(cur)
        return s[:300]
    except Exception:
        return None


def archived_days(customer_id):
    folder = CHANGE_DIR / customer_id
    if not folder.exists():
        return set()
    return {f.stem for f in folder.glob("*.jsonl")}


def archive_change_events(d, coverage):
    account = d["account"]
    client = d["puller"].client
    folder = CHANGE_DIR / account.customer_id
    folder.mkdir(parents=True, exist_ok=True)

    have = archived_days(account.customer_id)
    today = dt.date.today()
    # Re-pull the 2 most recent days each run (they may have been partial).
    days = []
    for back in range(1, 30):
        day = today - dt.timedelta(days=back)
        if str(day) not in have or back <= 2:
            days.append(day)

    pulled_rows = 0
    for day in sorted(days):
        rows, capped = _pull_change_day(client, account.customer_id, day)
        if rows is None:
            coverage.append(CoverageRow(
                account.name, f"change_event {day}", "gap", 0,
                "pull failed — day left unarchived, will retry next run"))
            continue
        path = folder / f"{day}.jsonl"
        with open(path, "w", encoding="utf-8") as f:
            for row in rows:
                f.write(json.dumps(row, ensure_ascii=False) + "\n")
        pulled_rows += len(rows)
        if capped:
            coverage.append(CoverageRow(
                account.name, f"change_event {day}", "gap", len(rows),
                "a 6-hour chunk hit the 10k cap — day may be incomplete"))
    coverage.append(CoverageRow(
        account.name, "change_event archive", "ok", pulled_rows,
        f"{len(days)} day(s) archived this run; "
        f"{len(archived_days(account.customer_id))} days total in history/"))


def _pull_change_day(client, customer_id, day):
    """Pull one day; if it hits the 10k LIMIT, split into 4 six-hour chunks."""
    def q(start, end):
        query = CHANGE_QUERY.format(start=start, end=end)
        return [_change_row(r) for r in stream(client, customer_id, query)]
    try:
        rows = q(f"{day} 00:00:00", f"{day} 23:59:59")
        if len(rows) < 10000:
            return rows, False
        rows, capped = [], False
        for h in (0, 6, 12, 18):
            chunk = q(f"{day} {h:02d}:00:00", f"{day} {h + 5:02d}:59:59")
            capped = capped or len(chunk) >= 10000
            rows.extend(chunk)
        return dedupe_changes(rows), capped
    except Exception:
        return None, False


def dedupe_changes(rows):
    seen, out = set(), []
    for r in rows:
        key = (r["time"], r["resource"], r["operation"], tuple(r["fields"]))
        if key not in seen:
            seen.add(key)
            out.append(r)
    return out


def load_change_log(customer_id, limit=5000):
    folder = CHANGE_DIR / customer_id
    rows = []
    if not folder.exists():
        return rows
    for f in sorted(folder.glob("*.jsonl"), reverse=True):
        with open(f, encoding="utf-8") as fh:
            for line in fh:
                try:
                    rows.append(json.loads(line))
                except json.JSONDecodeError:
                    continue
        if len(rows) >= limit:
            break
    rows.sort(key=lambda r: r["time"], reverse=True)
    return rows[:limit]


# ============================================================================
# SECTION 10 — HEALTH SCORE  (pure, tested)
# ============================================================================

def health_score(findings_for_account):
    """0–100; per-module penalty = Σ severity weights, capped per module."""
    per_module = defaultdict(float)
    for f in findings_for_account:
        if f.module in MODULE_SCORE_CAP:
            per_module[f.module] += SEVERITY_PENALTY.get(f.severity, 1.0)
    breakdown = {}
    total = 0.0
    for module, cap in MODULE_SCORE_CAP.items():
        pen = min(cap, per_module.get(module, 0.0))
        breakdown[module] = round(cap - pen, 1)
        total += pen
    return round(max(0.0, 100.0 - total), 1), breakdown


SCORES_CSV = HISTORY_DIR / "health_scores.csv"


def append_health_history(run_date, account_name, score, breakdown):
    SCORES_CSV.parent.mkdir(parents=True, exist_ok=True)
    new = not SCORES_CSV.exists()
    with open(SCORES_CSV, "a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if new:
            w.writerow(["date", "account", "score"] + list(MODULE_SCORE_CAP))
        w.writerow([run_date, account_name, score] +
                   [breakdown[m] for m in MODULE_SCORE_CAP])


def previous_score(account_name, before_date):
    if not SCORES_CSV.exists():
        return None
    last = None
    with open(SCORES_CSV, newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            if row["account"] == account_name and row["date"] < before_date:
                last = float(row["score"])
    return last


def snapshot_budgets(run_date, all_data):
    """Daily budget snapshot = the budget history no human maintains."""
    folder = HISTORY_DIR / "budgets"
    folder.mkdir(parents=True, exist_ok=True)
    snap = {}
    for d in all_data:
        camps = d.get("campaigns") or []
        snap[d["account"].name] = [
            {"campaign_id": c["id"], "campaign": c["name"], "status": c["status"],
             "daily_budget": c["budget"], "shared": c["budget_shared"]}
            for c in camps]
    with open(folder / f"{run_date}.json", "w", encoding="utf-8") as f:
        json.dump(snap, f, indent=1)


# ============================================================================
# SECTION 11 — EXCEL WORKBOOK (one workbook per run)
# ============================================================================

def write_workbook(path, run_date, all_data, findings, coverage, url_results,
                   scores, narrative, change_logs=None, days_archived=None):
    """change_logs / days_archived default to the on-disk archive in history/."""
    if change_logs is None:
        change_logs = {d["account"].name: load_change_log(d["account"].customer_id, 2000)
                       for d in all_data}
    if days_archived is None:
        days_archived = {d["account"].name: len(archived_days(d["account"].customer_id))
                         for d in all_data}
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    HEAD_FILL = PatternFill("solid", fgColor="1F3864")
    HEAD_FONT = Font(color="FFFFFF", bold=True)
    SEV_FILL = {"CRITICAL": PatternFill("solid", fgColor="C00000"),
                "HIGH": PatternFill("solid", fgColor="ED7D31"),
                "MEDIUM": PatternFill("solid", fgColor="FFD966"),
                "LOW": PatternFill("solid", fgColor="D9D9D9")}

    wb = Workbook()

    def sheet(title, headers, rows, widths=None, money_cols=(), sev_col=None):
        ws = wb.create_sheet(title)
        ws.append(headers)
        for i, _ in enumerate(headers, 1):
            c = ws.cell(row=1, column=i)
            c.fill, c.font = HEAD_FILL, HEAD_FONT
            c.alignment = Alignment(vertical="center")
        for r in rows:
            ws.append(list(r))
        for i, w in enumerate(widths or [18] * len(headers), 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        for col in money_cols:
            for row in range(2, ws.max_row + 1):
                ws.cell(row=row, column=col).number_format = "#,##0.00"
        if sev_col:
            for row in range(2, ws.max_row + 1):
                v = ws.cell(row=row, column=sev_col).value
                if v in SEV_FILL:
                    ws.cell(row=row, column=sev_col).fill = SEV_FILL[v]
        ws.freeze_panes = "A2"
        return ws

    # --- Executive Summary ---------------------------------------------------
    ws = wb.active
    ws.title = "Executive Summary"
    ws["A1"] = f"Account Guardian — {run_date}"
    ws["A1"].font = Font(size=16, bold=True)
    total_waste = sum(f.monthly_cost or 0 for f in findings)
    ws["A3"] = f"Sum of priced findings ({THRESHOLDS['WINDOW_DAYS']}d window):"
    ws["E3"] = total_waste
    ws["E3"].number_format = "#,##0.00"
    ws["E3"].font = Font(bold=True)
    ws["A4"] = ("One problem can trigger several checks (e.g. a dead ad group and "
                "its broken URL), so this sum can count the same spend twice.")
    ws["A4"].font = Font(italic=True, color="808080")

    ws["A5"] = "Account health"
    ws["A5"].font = Font(bold=True)
    ws.append([])
    ws.append(["Account", "Score /100", "Δ vs previous run"] + list(MODULE_SCORE_CAP))
    hdr_row = ws.max_row
    for i in range(1, 8):
        c = ws.cell(row=hdr_row, column=i)
        c.fill, c.font = HEAD_FILL, HEAD_FONT
    for name, (score, breakdown, prev) in scores.items():
        delta = "" if prev is None else f"{score - prev:+.1f}"
        ws.append([name, score, delta] + [breakdown[m] for m in MODULE_SCORE_CAP])

    ws.append([])
    ws.append(["Top findings (by priced waste, then severity)"])
    ws.cell(row=ws.max_row, column=1).font = Font(bold=True)
    ws.append(["Account", "Module", "Severity", "Finding", "Entity", "Cost"])
    hdr_row = ws.max_row
    for i in range(1, 7):
        c = ws.cell(row=hdr_row, column=i)
        c.fill, c.font = HEAD_FILL, HEAD_FONT
    top = sorted(findings, key=lambda f: (-(f.monthly_cost or 0),
                                          SEVERITY_ORDER.get(f.severity, 9)))[:20]
    for f in top:
        ws.append([f.account, f.module, f.severity, f.check, f.entity[:80],
                   f.monthly_cost or ""])
    for col, w in zip("ABCDEF", (14, 22, 10, 38, 50, 12)):
        ws.column_dimensions[col].width = w

    if narrative:
        ws.append([])
        ws.append(["Narrative"])
        ws.cell(row=ws.max_row, column=1).font = Font(bold=True)
        for line in narrative.splitlines():
            ws.append([line])

    # --- Findings Worklist -----------------------------------------------------
    ordered = sorted(findings, key=lambda f: (SEVERITY_ORDER.get(f.severity, 9),
                                              -(f.monthly_cost or 0)))
    sheet("Findings Worklist",
          ["Account", "Module", "Check", "Severity", "Entity", "Detail",
           f"Cost ({THRESHOLDS['WINDOW_DAYS']}d)", "Exact fix", "Done?"],
          [(f.account, f.module, f.check, f.severity, f.entity, f.detail,
            f.monthly_cost if f.monthly_cost is not None else "", f.fix, "")
           for f in ordered],
          widths=[12, 20, 32, 10, 45, 60, 12, 45, 8], money_cols=(7,), sev_col=4)

    # --- Change Log --------------------------------------------------------------
    change_rows = []
    for account_name, log in change_logs.items():
        for r in log:
            diff = "; ".join(f"{k}: {v['old']} → {v['new']}"
                             for k, v in list(r.get("diff", {}).items())[:5])
            change_rows.append((account_name, r["time"], r["user"], r["client"],
                                r["resource_type"], r["operation"],
                                ", ".join(r["fields"][:8]), diff, r["resource"]))
    change_rows.sort(key=lambda r: r[1], reverse=True)
    sheet("Change Log",
          ["Account", "When", "Who", "Client", "Resource type", "Op",
           "Changed fields", "Old → New", "Resource"],
          change_rows[:5000],
          widths=[12, 20, 28, 22, 20, 10, 40, 60, 50])

    # --- Change Impact (Phase 2 placeholder — archive is accruing) ---------------
    sheet("Change Impact",
          ["Account", "Days of change history archived", "Status"],
          [(name, n, "Phase 2 joins these changes to before/after performance. "
                     "The archive is accruing daily so the join is ready when built.")
           for name, n in days_archived.items()],
          widths=[14, 30, 90])

    # --- Destination Health -------------------------------------------------------
    verdict_order = {"broken": 0, "soft404": 1, "oos": 2, "suspect": 3, "ok": 4}
    url_rows = sorted(url_results, key=lambda r: (verdict_order.get(r["verdict"], 9),
                                                  -r["spend"]))
    sheet("Destination Health",
          ["Account", "Verdict", "Reason", f"Spend ({THRESHOLDS['WINDOW_DAYS']}d)",
           "Status", "Redirects", "URL", "Final URL"],
          [(r["account"], r["verdict"], r["reason"], r["spend"], r["status"],
            r["chain"], r["url"], r["final_url"]) for r in url_rows],
          widths=[12, 10, 30, 12, 8, 9, 60, 60], money_cols=(4,))

    # --- Budget & Pacing ------------------------------------------------------------
    budget_rows = []
    for d in all_data:
        for c in (d.get("campaigns") or []):
            if c["status"] == "REMOVED":
                continue
            budget_rows.append((d["account"].name, c["name"], c["status"],
                                c["channel"], c["budget"],
                                c["budget_recommended"] or "",
                                "shared" if c["budget_shared"] else ""))
    sheet("Budget & Pacing",
          ["Account", "Campaign", "Status", "Channel", "Daily budget (API)",
           "Google suggested", "Shared?"],
          budget_rows, widths=[12, 45, 10, 18, 16, 16, 9], money_cols=(5, 6))

    # --- Data Coverage ---------------------------------------------------------------
    sheet("Data Coverage",
          ["Account", "Source", "Status", "Rows", "Note"],
          [(c.account, c.source, c.status.upper(), c.rows, c.note) for c in coverage],
          widths=[12, 40, 8, 10, 90])

    # --- How To Read -------------------------------------------------------------------
    rules = [
        ("Read-only guarantee", "This tool never writes to any ad account. It recommends; a person applies."),
        ("No manual input", "Every figure comes from the Ads API or this tool's own archive in history/."),
        ("Window", f"All trailing metrics cover the last {THRESHOLDS['WINDOW_DAYS']} days."),
        ("Negative collision", "EXACT blocks identical text; PHRASE blocks in-order containment; BROAD blocks all-words. 'close' = singular/plural stem match, lower confidence."),
        ("Low QS", f"Flagged at quality score ≤ {THRESHOLDS['QS_MAX']} with spend ≥ {THRESHOLDS['MIN_SPEND_FOR_FINDING']}. NULL QS is a stated gap, never a zero."),
        ("Budget-limited winner", f"search_budget_lost_impression_share ≥ {THRESHOLDS['LOST_IS_MIN']:.0%} with ROAS ≥ account average. PMAX lost-IS is unreliable (0.0 sentinel) and excluded."),
        ("Pacing", f"Projected month spend vs daily-budget × days-in-month; flagged over {THRESHOLDS['PACE_HIGH']:.0%} or under {THRESHOLDS['PACE_LOW']:.0%}."),
        ("Soft-404", f"200-status pages that redirect to homepage, carry not-found language, or are under {THRESHOLDS['SOFT404_MIN_BODY']} bytes."),
        ("Landing spend", "landing_page_view sums are reconciled against campaign totals; over-attribution is scaled down, never up."),
        ("Health score", "100 minus per-module penalties (CRITICAL 8, HIGH 5, MEDIUM 2, LOW 1), capped per module: " + ", ".join(f"{m} {int(c)}" for m, c in MODULE_SCORE_CAP.items()) + "."),
        ("Change archive", "change_event pulled daily, paged by day (max 10k rows/query), stored forever in history/change_events/."),
        ("Gaps", "Anything unreachable appears on Data Coverage as a GAP with the fix — never as zeros."),
    ]
    sheet("How To Read", ["Rule", "Definition"], rules, widths=[26, 140])

    wb.save(path)


# ============================================================================
# SECTION 12 — NARRATIVE (optional LLM, hard-fenced; --no-llm = full report)
# ============================================================================

def deterministic_narrative(run_date, scores, findings):
    total = sum(f.monthly_cost or 0 for f in findings)
    crit = sum(1 for f in findings if f.severity == "CRITICAL")
    lines = [f"Run {run_date}: {len(findings)} findings across "
             f"{len(scores)} account(s); {crit} critical; "
             f"priced findings sum to {total:,.2f} (checks can overlap)."]
    for name, (score, _, prev) in scores.items():
        trend = "" if prev is None else f" ({score - prev:+.1f} vs last run)"
        lines.append(f"  {name}: health {score}/100{trend}.")
    return "\n".join(lines)


def gemini_narrative(run_date, scores, findings):
    """LLM writes prose ONLY; every number it emits must already exist in the
    deterministic figures or the narrative is discarded."""
    api_key = _env("GEMINI_API_KEY")
    if not api_key:
        return None
    try:
        import google.generativeai as genai
        genai.configure(api_key=api_key)
        total = sum(f.monthly_cost or 0 for f in findings)
        facts = {
            "run_date": run_date,
            "total_priced_waste": round(total, 2),
            "finding_count": len(findings),
            "critical_count": sum(1 for f in findings if f.severity == "CRITICAL"),
            "accounts": {n: s for n, (s, _, _) in scores.items()},
            "top_findings": [
                {"account": f.account, "check": f.check, "cost": round(f.monthly_cost or 0, 2)}
                for f in sorted(findings, key=lambda f: -(f.monthly_cost or 0))[:5]],
        }
        model = genai.GenerativeModel(_env("GEMINI_MODEL", "gemini-2.0-flash"))
        resp = model.generate_content(
            "Write a 4-6 sentence executive summary of this ads audit. Use ONLY "
            "numbers present in the JSON, verbatim. No advice beyond the data.\n"
            + json.dumps(facts))
        text = resp.text.strip()
        if _narrative_numbers_ok(text, facts):
            return text
    except Exception:
        pass
    return None


def _narrative_numbers_ok(text, facts):
    allowed = set()

    def collect(v):
        if isinstance(v, dict):
            for x in v.values():
                collect(x)
        elif isinstance(v, list):
            for x in v:
                collect(x)
        elif isinstance(v, (int, float)):
            for fmt in (f"{v:,.2f}", f"{v:,.1f}", f"{v:,.0f}", str(v),
                        f"{v:.2f}", f"{v:.1f}", f"{v:.0f}" if isinstance(v, float) else str(v)):
                allowed.add(fmt.replace(",", ""))
    collect(facts)
    for num in re.findall(r"(?<![\w.])\d[\d,]*(?:\.\d+)?(?!\w)", text):
        clean = num.replace(",", "").rstrip(".")
        if clean and clean not in allowed and not re.match(r"^(19|20)\d\d", clean):
            return False
    return True


# ============================================================================
# SECTION 13 — PROBE  (--probe: the four "probe first" API items)
# ============================================================================

def run_probe(account):
    print(f"Probing {account.name} ({account.customer_id})\n")
    client = build_client(account)
    cid = account.customer_id

    print("1) change_event rows/day (10k cap risk):")
    for back in range(1, 8):
        day = dt.date.today() - dt.timedelta(days=back)
        q = CHANGE_QUERY.format(start=f"{day} 00:00:00", end=f"{day} 23:59:59")
        try:
            n = sum(1 for _ in stream(client, cid, q))
            flag = "  <-- AT CAP, needs sub-day paging" if n >= 10000 else ""
            print(f"   {day}: {n}{flag}")
        except Exception as e:
            print(f"   {day}: FAILED {type(e).__name__}: {str(e)[:150]}")

    print("\n2) recommendation types:")
    try:
        types = Counter(r.recommendation.type_.name for r in stream(
            client, cid, "SELECT recommendation.type FROM recommendation"))
        print("   " + (", ".join(f"{t}×{n}" for t, n in types.most_common()) or "none"))
    except Exception as e:
        print(f"   FAILED {type(e).__name__}: {str(e)[:150]}")

    print("\n3) asset_group.ad_strength:")
    try:
        rows = list(stream(client, cid, """
            SELECT asset_group.id, asset_group.name, asset_group.ad_strength
            FROM asset_group LIMIT 20"""))
        for r in rows[:10]:
            print(f"   {r.asset_group.name}: {r.asset_group.ad_strength.name}")
        if not rows:
            print("   no asset groups (no PMAX on this account)")
    except Exception as e:
        print(f"   FAILED {type(e).__name__}: {str(e)[:150]}")

    print("\n4) asset_group_asset policy coverage by type:")
    try:
        cov = Counter()
        for r in stream(client, cid, """
            SELECT asset.type, asset_group_asset.policy_summary.approval_status
            FROM asset_group_asset LIMIT 5000"""):
            cov[(r.asset.type_.name,
                 r.asset_group_asset.policy_summary.approval_status.name)] += 1
        for (atype, status), n in sorted(cov.items()):
            print(f"   {atype:20s} {status:20s} ×{n}")
        if not cov:
            print("   no asset_group_asset rows")
    except Exception as e:
        print(f"   FAILED {type(e).__name__}: {str(e)[:150]}")
    return 0


# ============================================================================
# SECTION 14 — SELF-TEST  (--self-test: the silent-failure logic, no API)
# ============================================================================

def _t_normalize():
    assert normalize_kw('  "Running  Shoes"  ') == "running shoes"
    assert normalize_kw("+running +shoes") == "running shoes"
    assert normalize_kw("[running shoes]") == "running shoes"
    assert normalize_kw("trail-running shoe's") == "trail running shoe s"


def _t_exact_negative():
    assert negative_blocks("running shoes", "EXACT", "Running Shoes")
    assert not negative_blocks("running shoes", "EXACT", "running shoes sale")


def _t_phrase_negative():
    assert negative_blocks("running shoes", "PHRASE", "buy running shoes online")
    assert not negative_blocks("running shoes", "PHRASE", "shoes for running")


def _t_broad_negative():
    assert negative_blocks("shoes running", "BROAD", "running shoes sale")
    assert not negative_blocks("shoes running kids", "BROAD", "running shoes sale")


def _t_close_variant():
    assert not negative_blocks("sneaker", "BROAD", "sneakers for men", close=False)
    assert negative_blocks("sneaker", "BROAD", "sneakers for men", close=True)
    assert not negative_blocks("dres", "EXACT", "dress", close=True)


def _t_collision_confidence():
    negs = [{"text": "sneaker", "match": "BROAD"}]
    out = find_collisions([{"text": "sneakers", "match": "EXACT"}], negs)
    assert len(out) == 1 and out[0][2] == "close"
    out = find_collisions([{"text": "sneaker", "match": "EXACT"}], negs)
    assert out[0][2] == "literal"


def _t_empty_never_collides():
    assert not negative_blocks("", "BROAD", "running shoes")
    assert not negative_blocks("running", "BROAD", "")


def _t_pacing():
    run_rate, projected, ratio = month_pacing(1000.0, 10, 30, 3000.0)
    assert (run_rate, projected, ratio) == (100.0, 3000.0, 1.0)
    _, projected, ratio = month_pacing(0.0, 0, 30, 0.0)
    assert projected == 0.0 and ratio is None


def _page(status=200, req="https://x.com/product/a", final=None, body="x" * 2000,
          title="Product A"):
    return classify_page(status, req, final or req, body, title)


def _t_hard_errors():
    assert _page(status=404)[0] == "broken"
    assert _page(status=500)[0] == "broken"
    assert classify_page(None, "https://x.com/a", "https://x.com/a", None)[0] == "broken"


def _t_soft404():
    verdict, reason = _page(final="https://x.com/")
    assert verdict == "soft404" and "homepage" in reason
    assert _page(body="Sorry, this page cannot be found." + "x" * 2000)[0] == "soft404"
    assert _page(title="404 Not Found")[0] == "soft404"


def _t_out_of_stock():
    assert _page(body="This item is currently unavailable." + "x" * 2000)[0] == "oos"
    assert _page(body='"availability":"OutOfStock"' + "x" * 2000)[0] == "oos"


def _t_thin_and_ok():
    assert _page(body="tiny")[0] == "suspect"
    assert _page()[0] == "ok"


def _t_over_attribution():
    lp = [{"url": "https://x.com/a", "campaign_id": 1, "cost": 70.0},
          {"url": "https://x.com/b", "campaign_id": 1, "cost": 70.0}]
    spend, scale = reconcile_landing_spend(lp, {1: 100.0})
    assert abs(scale[1] - 100.0 / 140.0) < 1e-9
    assert abs(spend["https://x.com/a"] - 50.0) < 1e-9
    spend, scale = reconcile_landing_spend(lp[:1], {1: 100.0})
    assert scale[1] == 1.0 and spend["https://x.com/a"] == 70.0


def _t_url_key_merge():
    lp = [{"url": "https://X.com/a?utm=1", "campaign_id": 1, "cost": 10.0},
          {"url": "https://x.com/a/", "campaign_id": 1, "cost": 10.0}]
    spend, _ = reconcile_landing_spend(lp, {1: 100.0})
    assert spend == {"https://x.com/a": 20.0}


def _t_dedupe_changes():
    row = {"time": "2026-09-27 10:00:00", "resource": "customers/1/campaigns/2",
           "operation": "UPDATE", "fields": ["status"]}
    assert len(dedupe_changes([dict(row), dict(row), dict(row, operation="CREATE")])) == 2


def _t_health_score():
    def f(module, sev):
        return Finding("A", module, "c", sev, "e", "d", None, "f")
    assert health_score([])[0] == 100.0
    assert health_score([f("2. Policy & delivery", "CRITICAL")])[0] == 92.0
    score, breakdown = health_score([f("2. Policy & delivery", "CRITICAL")] * 10)
    assert score == 75.0 and breakdown["2. Policy & delivery"] == 0.0
    everything = [f(m, "CRITICAL") for m in MODULE_SCORE_CAP] * 20
    assert health_score(everything)[0] == 0.0


def _t_narrative_fence():
    facts = {"total": 1234.56, "accounts": {"Store1": 87.0}}
    assert _narrative_numbers_ok("Waste was 1,234.56 and Store1 scored 87.", facts)
    assert not _narrative_numbers_ok("Waste was 9,999.99 somehow.", facts)


def _t_env_parser():
    import tempfile
    with tempfile.TemporaryDirectory() as tmp:
        p = Path(tmp) / ".env"
        p.write_text('# c\nGUARDIAN_TEST_A="x y"\nGUARDIAN_TEST_B=1\n\nbad line\n',
                     encoding="utf-8")
        os.environ.pop("GUARDIAN_TEST_A", None)
        os.environ["GUARDIAN_TEST_B"] = "keep"
        load_env_file(p)
        assert os.environ.pop("GUARDIAN_TEST_A") == "x y"
        assert os.environ.pop("GUARDIAN_TEST_B") == "keep"


def run_self_test():
    tests = [(n, f) for n, f in globals().items() if n.startswith("_t_") and callable(f)]
    failed = 0
    for name, fn in tests:
        try:
            fn()
            print(f"  ok    {name[3:]}")
        except Exception as exc:
            failed += 1
            print(f"  FAIL  {name[3:]}: {type(exc).__name__} {exc}")
    print(f"\n{len(tests) - failed}/{len(tests)} passed")
    return 1 if failed else 0


# ============================================================================
# SECTION 15 — INIT  (--init: write the .env template)
# ============================================================================

ENV_TEMPLATE = """\
# Account Guardian settings. Never commit this file.

# --- Google Ads API credentials ---
GOOGLE_ADS_DEVELOPER_TOKEN=
GOOGLE_ADS_CLIENT_ID=
GOOGLE_ADS_CLIENT_SECRET=
GOOGLE_ADS_REFRESH_TOKEN=

# Default MCC (login-customer-id) most accounts sit under
GOOGLE_ADS_LOGIN_CUSTOMER_ID=

# --- Accounts: Name:customer_id[:login_customer_id_override],... ---
# Accounts under a different manager account carry the third field (that
# manager's ID). Digits only or dashed, both accepted.
# ACCOUNTS=Store1:111-111-1111,Store2:222-222-2222:999-999-9999
ACCOUNTS=

# --- Optional: Gemini writes the executive narrative (numbers are fenced) ---
GEMINI_API_KEY=
GEMINI_MODEL=gemini-2.0-flash

# --- Thresholds (defaults shown; uncomment to override) ---
# GUARDIAN_QS_MAX=3
# GUARDIAN_LOST_IS_MIN=0.10
# GUARDIAN_WINDOW_DAYS=30
# GUARDIAN_MAX_URLS=500
# GUARDIAN_URL_TIMEOUT=10
# GUARDIAN_URL_WORKERS=16
# GUARDIAN_SOFT404_MIN_BODY=512
# GUARDIAN_MAX_REDIRECTS=5
# GUARDIAN_PACE_HIGH=1.15
# GUARDIAN_PACE_LOW=0.70
# GUARDIAN_OVERFUNDED_ROAS_FRACTION=0.5
# GUARDIAN_MIN_SPEND_FOR_FINDING=5.0
"""


def run_init():
    if ENV_FILE.exists():
        print(f"{ENV_FILE} already exists; left untouched.")
        return 1
    ENV_FILE.write_text(ENV_TEMPLATE, encoding="utf-8")
    print(f"Wrote {ENV_FILE}. Fill in the credentials and ACCOUNTS, then run "
          f"--list-accounts.")
    return 0


# ============================================================================
# SECTION 16 — DEMO  (--demo: sample workbook from made-up data, no API)
# ============================================================================
# Two fictional stores run through the real audit modules, so the sample
# shows genuine engine output. Nothing is read from or written to history/.

DEMO_OUT = ROOT / "Account_Guardian_SAMPLE.xlsx"


def _demo_daily(campaign_id, name, channel, cost, value, conv, lost_is):
    rows = []
    for back in range(1, 31):
        day = dt.date.today() - dt.timedelta(days=back)
        wobble = 1 + 0.08 * ((back * 7) % 5 - 2) / 2
        rows.append({"id": campaign_id, "name": name, "channel": channel,
                     "date": str(day), "cost": round(cost * wobble, 2),
                     "clicks": int(cost * wobble / 1.2), "impressions": int(cost * 40),
                     "conv": conv * wobble, "conv_value": round(value * wobble, 2),
                     "lost_is": lost_is})
    return rows


def _demo_campaign(cid, name, channel, budget, status="ENABLED"):
    return {"id": cid, "name": name, "status": status, "channel": channel,
            "bidding": "MAXIMIZE_CONVERSION_VALUE", "budget_id": cid * 10,
            "budget": budget, "budget_delivery": "STANDARD", "budget_shared": False,
            "budget_recommended": 0.0}


def _demo_kw(cid, camp, agid, ag, crit, text, match, qs=7, negative=False,
             status="ENABLED"):
    return {"campaign_id": cid, "campaign": camp, "ad_group_id": agid,
            "ad_group": ag, "ad_group_status": "ENABLED", "criterion_id": crit,
            "text": text, "match": match, "status": status, "negative": negative,
            "qs": qs, "urls": []}


def _demo_ad(cid, camp, agid, ag, ad_id, url, approval="APPROVED",
             strength="GOOD", status="ENABLED"):
    return {"campaign_id": cid, "campaign": camp, "ad_group_id": agid,
            "ad_group": ag, "ad_group_status": "ENABLED", "ad_id": ad_id,
            "ad_type": "RESPONSIVE_SEARCH_AD",
            "status": status, "approval": approval, "review": "REVIEWED",
            "strength": strength, "urls": [url]}


def _demo_store_a(coverage):
    site = "https://store-a.example"
    s_run, s_brand, pmax, clear = (
        "Search - Running Shoes", "Search - Brand", "PMax - All Products",
        "Search - Clearance")
    men, women, kids, sale = ("Men's Running", "Women's Running", "Kids Sneakers",
                              "Shoe Sale")
    d = {"account": Account("Demo Store A", "1111111111", ""),
         "customer": [{"name": "Demo Store A", "currency": "USD", "tz": "America/New_York"}]}
    d["campaigns"] = [_demo_campaign(101, s_run, "SEARCH", 150),
                      _demo_campaign(102, s_brand, "SEARCH", 50),
                      _demo_campaign(103, pmax, "PERFORMANCE_MAX", 300),
                      _demo_campaign(104, clear, "SEARCH", 80)]
    d["campaign_daily"] = (
        _demo_daily(101, s_run, "SEARCH", 140, 900, 9, 0.25)          # limited winner
        + _demo_daily(102, s_brand, "SEARCH", 30, 400, 4, 0.0)
        + _demo_daily(103, pmax, "PERFORMANCE_MAX", 360, 1500, 15, 0.0)  # over-pacing
        + _demo_daily(104, clear, "SEARCH", 75, 60, 1, 0.0))            # overfunded loser
    d["keywords"] = [
        _demo_kw(101, s_run, 1001, men, 1, "running shoes for men", "PHRASE"),
        _demo_kw(101, s_run, 1001, men, 2, "mens running shoes", "EXACT"),
        _demo_kw(101, s_run, 1001, men, 3, "trail running shoes", "BROAD", qs=2),
        _demo_kw(101, s_run, 1002, women, 4, "womens running shoes", "EXACT"),
        _demo_kw(101, s_run, 1002, women, 5, "running shoes for men", "PHRASE"),
        _demo_kw(101, s_run, 1003, kids, 6, "kids sneakers", "BROAD", qs=None),
        _demo_kw(104, clear, 1004, sale, 7, "shoe sale", "BROAD"),
    ]
    d["campaign_negatives"] = [{"campaign_id": 101, "text": "mens", "match": "BROAD"}]
    d["shared_negatives"] = [
        {"set_id": 9, "set_name": "Brand Safety", "text": "free", "match": "BROAD"},
        {"set_id": 9, "set_name": "Brand Safety", "text": "trail", "match": "PHRASE"}]
    d["campaign_shared_sets"] = [{"campaign_id": 101, "set_id": 9}]
    d["keyword_metrics"] = [
        {"ad_group_id": 1001, "criterion_id": 1, "cost": 900.0, "clicks": 700, "conv": 30, "conv_value": 3600},
        {"ad_group_id": 1001, "criterion_id": 2, "cost": 612.40, "clicks": 480, "conv": 20, "conv_value": 2400},
        {"ad_group_id": 1001, "criterion_id": 3, "cost": 264.80, "clicks": 80, "conv": 2, "conv_value": 210},
        {"ad_group_id": 1002, "criterion_id": 4, "cost": 1180.0, "clicks": 950, "conv": 38, "conv_value": 4300},
        {"ad_group_id": 1002, "criterion_id": 5, "cost": 455.60, "clicks": 360, "conv": 5, "conv_value": 520},
        {"ad_group_id": 1003, "criterion_id": 6, "cost": 182.30, "clicks": 150, "conv": 1, "conv_value": 60},
    ]
    d["ad_group_metrics"] = [
        {"ad_group_id": 1001, "cost": 1777.20, "conv": 52, "conv_value": 6210},
        {"ad_group_id": 1002, "cost": 1635.60, "conv": 43, "conv_value": 4820},
        {"ad_group_id": 1003, "cost": 182.30, "conv": 1, "conv_value": 60},
        {"ad_group_id": 1004, "cost": 241.50, "conv": 1, "conv_value": 45}]
    d["ads"] = [
        _demo_ad(101, s_run, 1001, men, 1, f"{site}/running/men"),
        _demo_ad(101, s_run, 1001, men, 2, f"{site}/running/men", approval="APPROVED_LIMITED"),
        _demo_ad(101, s_run, 1002, women, 3, f"{site}/running/women", strength="POOR"),
        _demo_ad(101, s_run, 1003, kids, 4, f"{site}/kids/sneakers", approval="DISAPPROVED"),
        _demo_ad(101, s_run, 1003, kids, 5, f"{site}/kids/sneakers", approval="DISAPPROVED"),
        _demo_ad(104, clear, 1004, sale, 6, f"{site}/sale", status="PAUSED"),
        _demo_ad(102, s_brand, 1005, "Brand", 7, f"{site}/"),
    ]
    d["asset_groups"] = [
        {"campaign_id": 103, "campaign": pmax, "id": 31, "name": "All Products",
         "status": "ENABLED", "urls": [f"{site}/collections/all"], "strength": "AVERAGE"},
        {"campaign_id": 103, "campaign": pmax, "id": 32, "name": "Summer Collection",
         "status": "ENABLED", "urls": [f"{site}/collections/summer"], "strength": "GOOD"}]
    d["asset_group_assets"] = [
        {"campaign_id": 103, "asset_group_id": 31, "asset_group": "All Products",
         "field_type": "MARKETING_IMAGE", "approval": "DISAPPROVED",
         "asset_id": 1, "asset_type": "IMAGE"},
        {"campaign_id": 103, "asset_group_id": 31, "asset_group": "All Products",
         "field_type": "HEADLINE", "approval": "APPROVED", "asset_id": 2, "asset_type": "TEXT"}]
    d["sitelinks"] = [
        {"asset_id": 51, "text": "Clearance", "urls": [f"{site}/clearance"], "approval": "APPROVED"},
        {"asset_id": 52, "text": "Gift Cards", "urls": [f"{site}/gift-cards"], "approval": "DISAPPROVED"}]
    d["landing_pages"] = [
        {"url": f"{site}/running/men", "campaign_id": 101, "cost": 1780.0, "clicks": 0, "conv": 0, "conv_value": 0},
        {"url": f"{site}/running/women", "campaign_id": 101, "cost": 1640.0, "clicks": 0, "conv": 0, "conv_value": 0},
        {"url": f"{site}/kids/sneakers", "campaign_id": 101, "cost": 182.30, "clicks": 0, "conv": 0, "conv_value": 0},
        {"url": f"{site}/collections/all", "campaign_id": 103, "cost": 6900.0, "clicks": 0, "conv": 0, "conv_value": 0},
        {"url": f"{site}/collections/summer", "campaign_id": 103, "cost": 1420.0, "clicks": 0, "conv": 0, "conv_value": 0},
        {"url": f"{site}/clearance", "campaign_id": 104, "cost": 612.0, "clicks": 0, "conv": 0, "conv_value": 0},
        {"url": f"{site}/gift-cards", "campaign_id": 102, "cost": 96.0, "clicks": 0, "conv": 0, "conv_value": 0}]
    d["recommendations"] = [
        {"type": "CAMPAIGN_BUDGET", "campaign": s_run, "base_cost": 4200.0,
         "potential_cost": 5600.0, "base_conv": 270.0, "potential_conv": 352.0}]
    pages = {
        f"{site}/running/men": (200, f"{site}/running/men", "x" * 5000, "Men's Running Shoes"),
        f"{site}/running/women": (200, f"{site}/running/women", "x" * 5000, "Women's Running Shoes"),
        f"{site}/kids/sneakers": (404, f"{site}/kids/sneakers", "", "Not Found"),
        f"{site}/collections/all": (200, f"{site}/collections/all", "x" * 9000, "Shop All"),
        f"{site}/collections/summer": (200, f"{site}/", "x" * 9000, "Store A Home"),
        f"{site}/clearance": (200, f"{site}/clearance", "Everything here is sold out." + "x" * 4000, "Clearance"),
        f"{site}/gift-cards": (404, f"{site}/gift-cards", "", "Not Found"),
        f"{site}/": (200, f"{site}/", "x" * 9000, "Store A Home"),
    }
    return d, pages


def _demo_store_b(coverage):
    site = "https://store-b.example"
    bed, pmax = "Search - Bedding", "PMax - Home"
    d = {"account": Account("Demo Store B", "2222222222", ""),
         "customer": [{"name": "Demo Store B", "currency": "USD", "tz": "America/Chicago"}]}
    d["campaigns"] = [_demo_campaign(201, bed, "SEARCH", 100),
                      _demo_campaign(202, pmax, "PERFORMANCE_MAX", 200)]
    d["campaign_daily"] = (_demo_daily(201, bed, "SEARCH", 95, 520, 6, 0.03)
                           + _demo_daily(202, pmax, "PERFORMANCE_MAX", 170, 780, 8, 0.0))
    d["keywords"] = [
        _demo_kw(201, bed, 2001, "Duvet Covers", 11, "duvet covers", "EXACT"),
        _demo_kw(201, bed, 2001, "Duvet Covers", 12, "queen duvet cover", "PHRASE"),
        _demo_kw(201, bed, 2002, "Bedding Sale", 13, "bedding sale", "PHRASE")]
    d["campaign_negatives"], d["shared_negatives"], d["campaign_shared_sets"] = [], [], []
    d["keyword_metrics"] = [
        {"ad_group_id": 2001, "criterion_id": 11, "cost": 1500.0, "clicks": 1200, "conv": 110, "conv_value": 9200},
        {"ad_group_id": 2001, "criterion_id": 12, "cost": 700.0, "clicks": 560, "conv": 40, "conv_value": 3900},
        {"ad_group_id": 2002, "criterion_id": 13, "cost": 650.0, "clicks": 500, "conv": 30, "conv_value": 2500}]
    d["ad_group_metrics"] = [
        {"ad_group_id": 2001, "cost": 2200.0, "conv": 150, "conv_value": 13100},
        {"ad_group_id": 2002, "cost": 650.0, "conv": 30, "conv_value": 2500}]
    d["ads"] = [
        _demo_ad(201, bed, 2001, "Duvet Covers", 21, f"{site}/bedding/duvet-covers"),
        _demo_ad(201, bed, 2002, "Bedding Sale", 22, f"{site}/bedding/sale",
                 approval="APPROVED_LIMITED")]
    d["asset_groups"] = [
        {"campaign_id": 202, "campaign": pmax, "id": 41, "name": "Home Essentials",
         "status": "ENABLED", "urls": [f"{site}/home"], "strength": "EXCELLENT"}]
    d["asset_group_assets"], d["sitelinks"] = [], []
    d["landing_pages"] = [
        {"url": f"{site}/bedding/duvet-covers", "campaign_id": 201, "cost": 2200.0, "clicks": 0, "conv": 0, "conv_value": 0},
        {"url": f"{site}/bedding/sale", "campaign_id": 201, "cost": 650.0, "clicks": 0, "conv": 0, "conv_value": 0},
        {"url": f"{site}/home", "campaign_id": 202, "cost": 5100.0, "clicks": 0, "conv": 0, "conv_value": 0}]
    d["recommendations"] = []
    pages = {
        f"{site}/bedding/duvet-covers": (200, f"{site}/bedding/duvet-covers", "x" * 6000, "Duvet Covers"),
        f"{site}/bedding/sale": (200, f"{site}/bedding/sale", "<html>Loading</html>", "Sale"),
        f"{site}/home": (200, f"{site}/home", "x" * 8000, "Home Essentials"),
    }
    return d, pages


def _demo_change_log():
    base = dt.datetime.combine(dt.date.today(), dt.time(9, 0))

    def ev(days, hours, user, client, rtype, op, field, old, new, resource):
        when = base - dt.timedelta(days=days, hours=hours)
        return {"time": when.strftime("%Y-%m-%d %H:%M:%S"), "user": user,
                "client": client, "resource_type": rtype, "operation": op,
                "fields": [field], "diff": {field: {"old": old, "new": new}},
                "resource": resource}
    return {
        "Demo Store A": [
            ev(2, 1, "manager@example.com", "GOOGLE_ADS_WEB_CLIENT", "CAMPAIGN", "UPDATE",
               "maximize_conversion_value.target_roas", "3.8", "4.5", "PMax - All Products"),
            ev(5, 3, "manager@example.com", "GOOGLE_ADS_WEB_CLIENT", "CAMPAIGN_CRITERION",
               "CREATE", "keyword.text", "", "mens", "Search - Running Shoes"),
            ev(9, 0, "", "GOOGLE_ADS_RECOMMENDATIONS", "CAMPAIGN_BUDGET", "UPDATE",
               "amount_micros", "120000000", "150000000", "Search - Running Shoes"),
            ev(14, 6, "analyst@example.com", "GOOGLE_ADS_EDITOR", "AD_GROUP_AD", "UPDATE",
               "status", "ENABLED", "PAUSED", "Search - Clearance › Shoe Sale"),
            ev(21, 2, "manager@example.com", "GOOGLE_ADS_WEB_CLIENT", "SHARED_CRITERION",
               "CREATE", "keyword.text", "", "trail", "Brand Safety list"),
        ],
        "Demo Store B": [
            ev(3, 4, "owner@example.com", "GOOGLE_ADS_WEB_CLIENT", "CAMPAIGN_BUDGET",
               "UPDATE", "amount_micros", "80000000", "100000000", "Search - Bedding"),
        ],
    }


def run_demo():
    run_date = f"{dt.date.today()} (SAMPLE, made-up data)"
    findings, coverage, url_results = [], [], []
    all_data = []
    for build in (_demo_store_a, _demo_store_b):
        d, pages = build(coverage)
        d["puller"] = argparse.Namespace(coverage=coverage)

        def fake_check(url, pages=pages):
            status, final, body, title = pages.get(url, (404, url, "", "Not Found"))
            verdict, reason = classify_page(status, url, final, body, title)
            return {"url": url, "final_url": final, "status": status,
                    "chain": 0 if final == url else 1, "verdict": verdict,
                    "reason": reason}

        for source in ("campaign + campaign_budget", "ad_group_criterion", "ad_group_ad",
                       "asset_group (PMAX)", "landing_page_view (unexpanded)"):
            coverage.append(CoverageRow(d["account"].name, source, "ok",
                                        len(d.get("keywords") or []), "DEMO data"))
        run_module1(d, findings)
        run_module2(d, findings)
        run_module3(d, findings, False, THRESHOLDS["MAX_URLS"], url_results, fake_check)
        run_module4(d, findings)
        all_data.append(d)

    prev = {"Demo Store A": 26.0, "Demo Store B": 94.0}
    scores = {}
    for d in all_data:
        name = d["account"].name
        score, breakdown = health_score([f for f in findings if f.account == name])
        scores[name] = (score, breakdown, prev[name])

    change_logs = _demo_change_log()
    write_workbook(DEMO_OUT, run_date, all_data, findings, coverage, url_results,
                   scores, deterministic_narrative(run_date, scores, findings),
                   change_logs=change_logs,
                   days_archived={"Demo Store A": 30, "Demo Store B": 30})

    total = sum(f.monthly_cost or 0 for f in findings)
    print(f"{len(findings)} findings, priced findings sum {total:,.2f} (made-up data)")
    for name, (score, _, p) in scores.items():
        print(f"  {name}: {score}/100  ({score - p:+.1f})")
    print(f"\nSample workbook: {DEMO_OUT}")
    return 0


# ============================================================================
# SECTION 17 — MAIN
# ============================================================================

def run(argv=None):
    ap = argparse.ArgumentParser(description="Account Guardian — read-only ads audit")
    ap.add_argument("--init", action="store_true",
                    help="write a .env template next to this script")
    ap.add_argument("--self-test", action="store_true",
                    help="run the built-in unit tests (no API access needed)")
    ap.add_argument("--demo", action="store_true",
                    help="write a sample workbook from made-up data (no API needed)")
    ap.add_argument("--probe", metavar="ACCOUNT", nargs="?", const="",
                    help="probe the four 'probe first' API items on one account")
    ap.add_argument("--list-accounts", action="store_true")
    ap.add_argument("--accounts", help="comma-separated account names to include")
    ap.add_argument("--window", type=int, default=THRESHOLDS["WINDOW_DAYS"],
                    help="trailing metric window in days (default 30)")
    ap.add_argument("--skip-url-checks", action="store_true",
                    help="skip live HTTP destination checks (Module 3)")
    ap.add_argument("--max-urls", type=int, default=THRESHOLDS["MAX_URLS"])
    ap.add_argument("--no-llm", action="store_true",
                    help="deterministic narrative only (full report either way)")
    args = ap.parse_args(argv)
    THRESHOLDS["WINDOW_DAYS"] = args.window

    if args.init:
        return run_init()
    if args.self_test:
        return run_self_test()
    if args.demo:
        return run_demo()

    accounts = load_accounts()
    if not accounts:
        raise SystemExit("No accounts configured. Run `python guardian.py --init`, "
                         "then fill ACCOUNTS=Name:customer_id[:login_customer_id],...")
    if args.probe is not None:
        target = next((a for a in accounts
                       if not args.probe or a.name.lower() == args.probe.lower()), None)
        if not target:
            raise SystemExit(f"No account {args.probe!r} in .env ACCOUNTS")
        return run_probe(target)
    if args.list_accounts:
        for a in accounts:
            login = a.login_customer_id or "(default MCC)"
            print(f"{a.name:15s} {a.customer_id}  login-customer-id: {login}")
        return 0
    if args.accounts:
        wanted = {n.strip().lower() for n in args.accounts.split(",")}
        accounts = [a for a in accounts if a.name.lower() in wanted]
        if not accounts:
            raise SystemExit(f"No configured account matches {args.accounts!r}")

    run_date = str(dt.date.today())
    findings: list[Finding] = []
    coverage: list[CoverageRow] = []
    url_results: list[dict] = []
    all_data = []

    for account in accounts:
        print(f"[{account.name}] pulling…")
        try:
            d = pull_account_data(account, coverage, args.window)
        except SystemExit:
            raise
        except Exception as exc:
            coverage.append(CoverageRow(
                account.name, "account", "gap", 0,
                f"account unreachable: {type(exc).__name__}: {str(exc)[:200]} — "
                f"check login-customer-id override and API access, then re-run"))
            continue
        all_data.append(d)

        print(f"[{account.name}] auditing…")
        run_module1(d, findings)
        run_module2(d, findings)
        run_module3(d, findings, args.skip_url_checks, args.max_urls, url_results)
        run_module4(d, findings)
        print(f"[{account.name}] archiving change events…")
        archive_change_events(d, coverage)

    if not all_data:
        print("Every account failed — see console; nothing to report.", file=sys.stderr)
        return 1

    # Currency guard: cross-account totals only when currencies match.
    currencies = {(d.get("customer") or [{}])[0].get("currency") for d in all_data
                  if d.get("customer")}
    if len(currencies) > 1:
        coverage.append(CoverageRow(
            "(all)", "currency check", "gap", 0,
            f"Accounts use different currencies {sorted(currencies)} — cross-"
            f"account totals in this workbook mix currencies; read per-account."))

    scores = {}
    for d in all_data:
        name = d["account"].name
        mine = [f for f in findings if f.account == name]
        score, breakdown = health_score(mine)
        prev = previous_score(name, run_date)
        scores[name] = (score, breakdown, prev)
        append_health_history(run_date, name, score, breakdown)

    snapshot_budgets(run_date, all_data)

    narrative = None if args.no_llm else gemini_narrative(run_date, scores, findings)
    narrative = narrative or deterministic_narrative(run_date, scores, findings)

    REPORTS_DIR.mkdir(exist_ok=True)
    out = REPORTS_DIR / f"Account_Guardian_{run_date}.xlsx"
    write_workbook(out, run_date, all_data, findings, coverage, url_results,
                   scores, narrative)

    total_waste = sum(f.monthly_cost or 0 for f in findings)
    print(f"\n{len(findings)} findings, priced findings sum {total_waste:,.2f} "
          f"({args.window}d window)")
    for name, (score, _, prev) in scores.items():
        trend = "" if prev is None else f"  ({score - prev:+.1f})"
        print(f"  {name}: {score}/100{trend}")
    print(f"\nWorkbook: {out}")
    return 0


if __name__ == "__main__":
    sys.exit(run())
