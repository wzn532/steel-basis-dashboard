#!/usr/bin/env python3
"""Backfill RB futures closes for Jiangsu rebar spot series."""

from __future__ import annotations

import json
import sys
import time
import urllib.error
import urllib.request
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import dashboard_app as app  # noqa: E402


SPOT_KEYS = tuple(
    spot_key
    for spot_key, config in app.SPOT_SERIES.items()
    if config["product"] == "RB"
)
MONTHS = (1, 5, 10)


def needed_by_date() -> dict[str, list[tuple[str, int]]]:
    needed: dict[str, list[tuple[str, int]]] = {}
    with app.connect() as conn:
        rows = conn.execute(
            """
            SELECT DISTINCT date
            FROM spot_series_prices
            WHERE spot_key IN ({})
            ORDER BY date
            """.format(",".join("?" for _ in SPOT_KEYS)),
            SPOT_KEYS,
        ).fetchall()
        dates = [row["date"] for row in rows]

        for date_text in dates:
            for month in MONTHS:
                contract = app.next_contract_code(date_text, "RB", month)
                existing = conn.execute(
                    """
                    SELECT 1
                    FROM futures_prices
                    WHERE date = ?
                        AND product = 'RB'
                        AND contract_code = ?
                    """,
                    (date_text, contract),
                ).fetchone()
                if not existing:
                    needed.setdefault(date_text, []).append((contract, month))
    return needed


def fetch_shfe_rows(date_text: str) -> list[dict]:
    request = urllib.request.Request(
        app.shfe_daily_url(date_text),
        headers={
            "User-Agent": "Mozilla/5.0 steel-basis-dashboard/1.0",
            "Accept": "application/json,text/plain,*/*",
        },
    )
    with urllib.request.urlopen(request, timeout=6) as response:
        raw = response.read().decode("utf-8", errors="replace")
    payload = json.loads(raw)
    return payload.get("o_curinstrument") or payload.get("data") or []


def fetch_date(date_text: str, contracts: list[tuple[str, int]]) -> tuple[str, list[tuple[str, int, float]], str | None]:
    try:
        rows = fetch_shfe_rows(date_text)
    except urllib.error.HTTPError as exc:
        return date_text, [], f"HTTP {exc.code}"
    except Exception as exc:
        return date_text, [], str(exc)

    found: list[tuple[str, int, float]] = []
    errors: list[str] = []
    for contract, month in contracts:
        try:
            found.append((contract, month, app.close_from_daily_rows(rows, contract)))
        except Exception as exc:
            errors.append(f"{contract}: {exc}")
    return date_text, found, "; ".join(errors) if errors and not found else None


def main() -> None:
    app.init_db()
    needed = needed_by_date()
    print(f"candidate_dates={len(needed)} candidate_cells={sum(len(v) for v in needed.values())}", flush=True)
    saved = 0
    by_contract: dict[str, int] = {}
    errors: list[str] = []
    started = time.time()

    with ThreadPoolExecutor(max_workers=8) as pool:
        futures = {
            pool.submit(fetch_date, date_text, contracts): date_text
            for date_text, contracts in needed.items()
        }
        for i, future in enumerate(as_completed(futures), 1):
            date_text, found, error = future.result()
            for contract, month, close in found:
                app.save_futures(date_text, "RB", month, contract, close, "shfe")
                saved += 1
                by_contract[contract] = by_contract.get(contract, 0) + 1
            if error and len(errors) < 20:
                errors.append(f"{date_text}: {error}")
            if i % 50 == 0:
                print(f"processed={i}/{len(needed)} saved={saved}", flush=True)

    points = {
        spot_key: {month: len(app.basis_payload(None, month, spot_key)["points"]) for month in MONTHS}
        for spot_key in SPOT_KEYS
    }
    totals = {}
    with app.connect() as conn:
        rows = conn.execute(
            """
            SELECT contract_code, COUNT(*) AS n, MIN(date) AS mn, MAX(date) AS mx
            FROM futures_prices
            WHERE product = 'RB'
            GROUP BY contract_code
            ORDER BY contract_code
            """
        ).fetchall()
        totals = {row["contract_code"]: {"count": row["n"], "min": row["mn"], "max": row["mx"]} for row in rows}

    print(
        {
            "saved": saved,
            "by_contract": by_contract,
            "totals": totals,
            "points": points,
            "sample_errors": errors,
            "seconds": round(time.time() - started, 1),
        },
        flush=True,
    )


if __name__ == "__main__":
    main()
