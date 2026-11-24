"""python -m report [input folder] [--products products.csv] [--out output]
Builds output/sales_report.xlsx and output/summary.txt (a short, email-ready summary)."""
import argparse
import time
from pathlib import Path

from .load_clean import load_all
from .workbook import build


def main() -> None:
    p = argparse.ArgumentParser(description='Combine messy branch sales exports into one finished Excel report.')
    p.add_argument('input', type=Path, nargs='?', default=Path('samples/input'))
    p.add_argument('--products', type=Path, default=Path('samples/products.csv'))
    p.add_argument('--out', type=Path, default=Path('output'))
    a = p.parse_args()
    t0 = time.time()
    clean, exceptions, files = load_all(a.input, a.products)
    a.out.mkdir(parents=True, exist_ok=True)
    s = build(clean, exceptions, files, a.out / 'sales_report.xlsx')
    k = s['kpi']
    growth = f"{s['growth']:+.1%} on {s['prev']}" if s['growth'] is not None else 'first month'
    text = (f"Sales report to {s['last']}: net sales Rs. {k['Net sales']:,.0f} across {len(files)} branch exports "
            f"({growth} for the latest month). Best branch in {s['last']}: {s['best_branch']} "
            f"(Rs. {s['best_sales']:,.0f}). Top product: {s['top_product']}. "
            f"{s['exceptions']} lines need a look; see the Exceptions sheet.")
    (a.out / 'summary.txt').write_text(text + '\n', encoding='utf-8')
    print(text)
    print(f"{len(clean)} clean lines, {len(exceptions)} exceptions, {time.time() - t0:.1f} s. Saved {a.out / 'sales_report.xlsx'}")


if __name__ == '__main__':
    main()
