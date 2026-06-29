import json
from datetime import date

from app.services.db import neon_db

SENSITIVE_KEY_PARTS = (
    "token",
    "authorization",
    "password",
    "secret",
    "api_key",
    "apikey",
    "cookie",
    "credential",
)


def _as_text(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _first_text(mapping: dict, keys: tuple[str, ...]) -> str:
    for key in keys:
        value = _as_text(mapping.get(key))
        if value:
            return value
    return ""


def _normalize_date(value: object) -> str:
    raw = _as_text(value)
    if not raw:
        return ""
    digits = "".join(ch for ch in raw if ch.isdigit())
    if len(digits) == 8:
        return date(int(digits[:4]), int(digits[4:6]), int(digits[6:8])).isoformat()
    return date.fromisoformat(raw[:10]).isoformat()


def _redact_payload(value: object) -> object:
    if isinstance(value, dict):
        redacted = {}
        for key, child in value.items():
            lowered = str(key).lower()
            if any(part in lowered for part in SENSITIVE_KEY_PARTS):
                redacted[key] = "[redacted]"
            else:
                redacted[key] = _redact_payload(child)
        return redacted
    if isinstance(value, list):
        return [_redact_payload(item) for item in value]
    return value


def _comparison_site_ids(payload: dict) -> list[str]:
    raw_publishers = (
        payload.get("comparisonPublishers")
        or payload.get("competitorPublishers")
        or payload.get("competitors")
        or []
    )
    if not isinstance(raw_publishers, list):
        return []

    site_ids: list[str] = []
    for publisher in raw_publishers:
        if not isinstance(publisher, dict):
            continue
        site_id = _first_text(
            publisher,
            (
                "primarySourceId",
                "sourceID",
                "sourceId",
                "siteID",
                "siteId",
                "publisherSiteId",
                "affiliateSiteId",
            ),
        )
        if site_id:
            site_ids.append(site_id)
        if len(site_ids) == 4:
            break
    return site_ids


def build_publisher_qbr_record(payload: object) -> dict | None:
    if not isinstance(payload, dict):
        return None

    primary_site_id = _first_text(
        payload,
        (
            "primarySourceId",
            "sourceID",
            "sourceId",
            "siteID",
            "siteId",
            "publisherSiteId",
            "affiliateSiteId",
        ),
    )
    if not primary_site_id:
        return None

    from_date = _normalize_date(payload.get("fromDate") or payload.get("startDate") or payload.get("from"))
    to_date = _normalize_date(payload.get("toDate") or payload.get("endDate") or payload.get("to"))
    currency_code = _as_text(payload.get("currencyCode") or payload.get("currency") or "EUR").upper()

    if not from_date or not to_date or not currency_code:
        return None

    return {
        "primary_publisher_site_id": primary_site_id,
        "comparison_publisher_site_ids": _comparison_site_ids(payload),
        "from_date": from_date,
        "to_date": to_date,
        "currency_code": currency_code,
        "request_payload": _redact_payload(payload),
    }


async def save_publisher_qbr_request(record: dict) -> str:
    rows = await neon_db.query(
        "INSERT INTO publisher_qbr "
        "(primary_publisher_site_id, comparison_publisher_site_ids, from_date, to_date, currency_code, request_payload) "
        "VALUES ($1, $2, $3::date, $4::date, $5, $6::jsonb) "
        "RETURNING id",
        [
            record["primary_publisher_site_id"],
            record["comparison_publisher_site_ids"],
            date.fromisoformat(record["from_date"]),
            date.fromisoformat(record["to_date"]),
            record["currency_code"],
            json.dumps(record["request_payload"]),
        ],
    )
    return str(rows[0]["id"])

