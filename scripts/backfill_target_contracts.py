#!/usr/bin/env python3
"""Backfill selected HC contracts from SHFE daily data."""

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


TARGETS = [
    ("HC2501", 1),
    ("HC2505", 5),
    ("HC2610", 10),
    ("HC2701", 1),
]


def missing_by_date() -> dict[str, list[tuple[str, int]]]:
    needed: dict[str, list[tuple[str, int]]] = {}
    with app.connect() as conn:
        for contract, month in TARGETS:
            start, end, _label = app.contract_window(contract, month)
            rows = conn.execute(
                """
                SELECT s.date
                FROM spot_prices s
                LEFT JOIN futures_prices f
                    ON f.date = s.date
                    AND f.product = 'HC'
                    AND f.contract_code = ?
                WHERE s.date BETWEEN ? AND ?
                    AND f.date IS NULL
                ORDER BY s.date
                """,
                (contract, start.isoformat(), end.isoformat()),
            ).fetchall()
            for row in rows:
                needed.setdefault(row["date"], []).append((contract, month))
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
    needed = missing_by_date()
    print(f"candidate_dates={len(needed)} candidate_cells={sum(len(v) for v in needed.values())}", flush=True)

    saved = 0
    errors: list[str] = []
    by_contract: dict[str, int] = {}
    started = time.time()

    with ThreadPoolExecutor(max_workers=8) as pool:
        futures = {
            pool.submit(fetch_date, date_text, contracts): (date_text, contracts)
            for date_text, contracts in needed.items()
        }
        for i, future in enumerate(as_completed(futures), 1):
            date_text, found, error = future.result()
            for contract, month, close in found:
                app.save_futures(date_text, "HC", month, contract, close, "shfe")
                saved += 1
                by_contract[contract] = by_contract.get(contract, 0) + 1
            if error and len(errors) < 20:
                errors.append(f"{date_text}: {error}")
            if i % 50 == 0:
                print(f"processed={i}/{len(needed)} saved={saved}", flush=True)

    totals = {}
    with app.connect() as conn:
        for contract, _month in TARGETS:
            row = conn.execute(
                """
                SELECT COUNT(*) AS n, MIN(date) AS mn, MAX(date) AS mx
                FROM futures_prices
                WHERE product = 'HC' AND contract_code = ?
                """,
                (contract,),
            ).fetchone()
            totals[contract] = {"count": row["n"], "min": row["mn"], "max": row["mx"]}

    points = {
        "HC01": len(app.basis_payload("HC", 1)["points"]),
        "HC05": len(app.basis_payload("HC", 5)["points"]),
        "HC10": len(app.basis_payload("HC", 10)["points"]),
    }
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
