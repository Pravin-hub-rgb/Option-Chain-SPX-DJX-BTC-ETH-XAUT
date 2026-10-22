import asyncio, sys
sys.path.insert(0, '.')
import main as M

async def run():
    print('=== SPOT PRICES (app reads spot_price) ===')
    for i in range(3):
        for asset in ['BTC', 'ETH', 'XAUT']:
            d = await asyncio.to_thread(M._get_json, f"{M.REST_BASE}/v2/tickers/{asset}USD", None)
            res = d.get('result') or {}
            print(f"  {asset}: spot_price={res.get('spot_price')} | mark_price={float(res.get('mark_price')):.4f} | close={res.get('close')}")
        await asyncio.sleep(0.4)

    print()
    print('=== OI DATA (fetch_tickers after fix) ===')
    for asset in ['BTC', 'ETH', 'XAUT']:
        products = await asyncio.to_thread(M.fetch_products, asset)
        expiries = [p.get('settlement_time', '')[:10] for p in products if p.get('settlement_time')]
        results = await asyncio.gather(*[asyncio.to_thread(M.fetch_tickers, asset, e) for e in expiries[:3]], return_exceptions=True)
        ok = sum(1 for r in results if not isinstance(r, BaseException))
        first = results[0]
        if isinstance(first, list) and len(first) > 0:
            first = first[0]
        print(f"  {asset}: {ok}/3 OK | first spot_price sample: {first.get('spot_price', 'N/A')} | OI sample: {first.get('oi', 'N/A')}")

asyncio.run(run())
