import asyncio
import time

async def tick(n):
    for i in range(n):
        print(f"tick {i}", flush=True)
        await asyncio.sleep(1)

asyncio.run(tick(5))
