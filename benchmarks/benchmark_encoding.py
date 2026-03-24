import asyncio
import time
import base64
from pathlib import Path
import os
import shutil
from kami_excel_extractor.core import KamiExcelExtractor

def create_dummy_images(num_images, size_kb=100):
    os.makedirs("bench_images", exist_ok=True)
    image_paths = []
    content = b"0" * (size_kb * 1024)
    for i in range(num_images):
        p = Path(f"bench_images/img_{i}.png")
        p.write_bytes(content)
        image_paths.append(p)
    return image_paths

async def benchmark_concurrent_reads(extractor, image_paths):
    start_time = time.perf_counter()

    tasks = []
    for p in image_paths:
        tasks.append(extractor._encode_image_to_base64_url(p))

    results = await asyncio.gather(*tasks)

    end_time = time.perf_counter()
    return end_time - start_time

async def main():
    extractor = KamiExcelExtractor(api_key="fake")
    num_images = 100
    image_paths = create_dummy_images(num_images, size_kb=500) # 500KB each

    print(f"Benchmarking {num_images} concurrent image encodings (Async optimized)...")
    duration = await benchmark_concurrent_reads(extractor, image_paths)
    print(f"Total time: {duration:.4f}s")
    print(f"Average time: {duration/num_images:.4f}s")

    shutil.rmtree("bench_images")

if __name__ == "__main__":
    asyncio.run(main())
