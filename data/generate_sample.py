#!/usr/bin/env python3
"""Generate sample CSV datasets for QtSpreadsheet testing.
Usage: python3 generate_sample.py [rows]
Default: 100,000 rows. For 1M row stress test: python3 generate_sample.py 1000000
"""
import csv, random, datetime, sys, time

PRODUCTS = ["Widget A","Widget B","Gadget X","Gadget Y","Device Pro","Device Lite",
            "Module Z","Module W","Pack Plus","Pack Basic","Ultra Kit","Starter Kit"]
REGIONS  = ["North","South","East","West","Central","International"]
REPS     = ["Alice Smith","Bob Jones","Carol Lee","David Kim","Eva Martinez",
            "Frank Chen","Grace Patel","Henry Brown"]

def generate(path, rows):
    start = time.time()
    random.seed(42)
    base  = datetime.date(2020,1,1)
    with open(path,"w",newline="",encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["ID","Date","Product","Region","SalesRep",
                    "Units","UnitPrice","Revenue","Discount","NetRevenue","Category"])
        for i in range(1, rows+1):
            prod  = random.choice(PRODUCTS)
            units = random.randint(1, 500)
            price = round(random.uniform(9.99, 499.99), 2)
            rev   = round(units * price, 2)
            disc  = round(random.uniform(0, 0.25), 3)
            net   = round(rev * (1-disc), 2)
            cat   = ("Electronics" if any(k in prod for k in ("Device","Gadget")) else
                     "Hardware"    if "Widget" in prod else
                     "Bundle"      if "Kit"    in prod else "Accessories")
            w.writerow([i,
                        (base + datetime.timedelta(days=random.randint(0,1460))).isoformat(),
                        prod, random.choice(REGIONS), random.choice(REPS),
                        units, price, rev, disc, net, cat])
            if i % 100_000 == 0:
                print(f"  {i:>10,} rows  ({time.time()-start:.1f}s)")
    elapsed = time.time() - start
    import os
    size_mb = os.path.getsize(path) / 1_048_576
    print(f"\nGenerated {rows:,} rows  →  {path}")
    print(f"File size : {size_mb:.1f} MB")
    print(f"Time      : {elapsed:.2f}s")

if __name__ == "__main__":
    n = int(sys.argv[1]) if len(sys.argv) > 1 else 100_000
    out = f"sample_{n//1000}k.csv" if n >= 1000 else f"sample_{n}.csv"
    generate(out, n)
