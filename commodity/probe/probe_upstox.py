"""
MCX Upstox probe (round 4) - complete field dump for one live option key.

Dumps the FULL /market-quote/quotes response for MCX_FO|579520 (SILVER PE
26 MAY 27) so we can see every field: last_price, oi, timestamp, depth, etc.
"""

import json
import os
import sys
from typing import Optional

import requests
from dotenv import load_dotenv

BASE_URL = "https://api.upstox.com/v2"


def main() -> int:
    load_dotenv()
    token = os.getenv("UPSTOX_API_TOKEN", "").strip()
    print("MCX Upstox probe (round 4) - full field dump")
    print("=" * 50)
    if not token:
        print("ERROR: missing UPSTOX_API_TOKEN")
        return 1

    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}

    # The key that returned real structure in round 2.
    key = "MCX_FO|579520"

    r = requests.get(
        f"{BASE_URL}/market-quote/quotes",
        headers=headers,
        params={"instrument_key": key},
        timeout=30,
    )
    print(f"GET /market-quote/quotes?instrument_key={key} -> {r.status_code}")
    if r.status_code == 200:
        data = r.json().get("data", {}).get(key, {})
        print(json.dumps(data, indent=2, ensure_ascii=False))
    else:
        print(r.text[:1000])

    print("\n" + "=" * 50)
    print("Done.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
