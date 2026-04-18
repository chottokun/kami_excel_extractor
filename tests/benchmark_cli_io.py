import asyncio
import json
import time
import os
from pathlib import Path

async def heartbeat(interval=0.01):
    max_gap = 0
    count = 0
    try:
        while True:
            t0 = time.perf_counter()
            await asyncio.sleep(interval)
            t1 = time.perf_counter()
            gap = (t1 - t0) - interval
            if gap > max_gap:
                max_gap = gap
            count += 1
    except asyncio.CancelledError:
        return max_gap

def sync_write(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

async def async_write(path, data):
    def _write():
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
    await asyncio.to_thread(_write)

async def main():
    # Create large dummy data
    data = {"items": [{"id": i, "value": "x" * 100} for i in range(100000)]}
    path = Path("test_benchmark.json")

    print("--- Testing Sync Write ---")
    hb_task = asyncio.create_task(heartbeat())
    await asyncio.sleep(0.1) # Let heartbeat stabilize

    t0 = time.perf_counter()
    sync_write(path, data)
    t1 = time.perf_counter()

    await asyncio.sleep(0.1)
    hb_task.cancel()
    max_gap = await hb_task
    print(f"Sync Write took: {t1-t0:.4f}s")
    print(f"Heartbeat max gap: {max_gap:.4f}s")

    if os.path.exists(path): os.remove(path)

    print("\n--- Testing Async (to_thread) Write ---")
    hb_task = asyncio.create_task(heartbeat())
    await asyncio.sleep(0.1)

    t0 = time.perf_counter()
    await async_write(path, data)
    t1 = time.perf_counter()

    await asyncio.sleep(0.1)
    hb_task.cancel()
    max_gap = await hb_task
    print(f"Async Write took: {t1-t0:.4f}s")
    print(f"Heartbeat max gap: {max_gap:.4f}s")

    if os.path.exists(path): os.remove(path)

if __name__ == "__main__":
    asyncio.run(main())
