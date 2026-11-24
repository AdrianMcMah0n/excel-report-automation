"""Read every branch export, whatever its layout, into one clean table plus a list of exceptions.

- Finds the header row (title rows above it are skipped) and maps each branch's column names to
  date / product / qty / price / total.
- Parses dates (dd/mm/yyyy is read day first) and amounts ('1,234.50', 'Rs. 1,234', '₹1,234').
- Matches product names to the product list, including variants like 'steel bottle 1 l'.
- Removes re-exported duplicate lines, checks totals, and checks each file's own 'Total' row if it has one.
"""
import csv
import re
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook
from rapidfuzz import fuzz, process

SYNONYMS = {
    'invoice': ['invoice', 'invoice no', 'order no', 'txn id', 'order id', 'bill no'],
    'date': ['date', 'order date', 'txn date', 'invoice date', 'sale date'],
    'product': ['product', 'item', 'sku name', 'product name', 'description'],
    'qty': ['qty', 'quantity', 'units', 'qty sold'],
    'price': ['unit price', 'price', 'rate', 'selling price'],
    'total': ['total', 'amount', 'amount (inr)', 'sales', 'line total', 'net amount'],
}
LOOKUP = {name: field for field, names in SYNONYMS.items() for name in names}


def read_rows(path: Path) -> list[list]:
    if path.suffix.lower() == '.csv':
        with path.open(encoding='utf-8-sig', newline='') as f:
            return [row for row in csv.reader(f)]
    ws = load_workbook(path, data_only=True).active
    return [list(r) for r in ws.iter_rows(values_only=True)]


def find_header(rows: list[list]) -> tuple[int, dict]:
    """First row (within the top 10) where at least 3 cells are known column names."""
    for i, row in enumerate(rows[:10]):
        cols = {LOOKUP[str(c).strip().lower()]: j for j, c in enumerate(row) if c and str(c).strip().lower() in LOOKUP}
        if len(cols) >= 3:
            return i, cols
    raise ValueError('no header row found')


def to_date(v) -> date | None:
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v or '').strip()
    for fmt in ('%d/%m/%Y', '%Y-%m-%d', '%d-%m-%Y', '%d.%m.%Y'):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def to_number(v) -> float | None:
    if isinstance(v, (int, float)):
        return float(v)
    s = re.sub(r'(?i)rs\.?|inr|₹|,|\s', '', str(v or ''))
    try:
        return float(s)
    except ValueError:
        return None


class ProductMatcher:
    def __init__(self, products: list[dict]):
        self.by_key = {self.key(p['product']): p for p in products}

    @staticmethod
    def key(name: str) -> str:
        return re.sub(r'[\s\-]+', ' ', str(name)).strip().lower()

    def match(self, name: str) -> tuple[dict | None, str]:
        k = self.key(name)
        if k in self.by_key:
            return self.by_key[k], ''
        best = process.extractOne(k, list(self.by_key), scorer=fuzz.token_sort_ratio)
        if best and best[1] >= 85:
            hit = self.by_key[best[0]]
            return hit, f"'{name}' read as '{hit['product']}'"
        return None, ''


def load_all(input_dir: Path, products_csv: Path):
    with products_csv.open(encoding='utf-8') as f:
        matcher = ProductMatcher(list(csv.DictReader(f)))
    clean, exceptions, files = [], [], []
    for path in sorted(input_dir.glob('*.*')):
        if path.suffix.lower() not in ('.csv', '.xlsx'):
            continue
        branch = path.stem.split('_')[0].title()
        rows = read_rows(path)
        h, cols = find_header(rows)
        seen, kept, excluded, footer, read_total, excluded_amount = set(), 0, 0, None, 0.0, 0.0
        for n, row in enumerate(rows[h + 1:], start=h + 2):
            cells = [row[cols[f]] if cols.get(f) is not None and cols[f] < len(row) else None for f in SYNONYMS]
            if all(c in (None, '') for c in cells):
                continue
            if any(str(c or '').strip().lower() == 'total' for c in cells[:2]):
                footer = to_number(cells[5])
                continue
            inv = str(cells[0] or '').strip()
            d, name, qty, price, total = to_date(cells[1]), str(cells[2] or '').strip(), to_number(cells[3]), to_number(cells[4]), to_number(cells[5])
            read_total += total or 0
            where = {'file': path.name, 'row': n, 'branch': branch, 'values': ' | '.join(str(c) for c in cells)}

            def flag(issue, action):
                nonlocal excluded_amount
                exceptions.append({**where, 'issue': issue, 'action': action})
                if action == 'Excluded':
                    excluded_amount += total or 0

            if d is None:
                flag('missing or unreadable date', 'Excluded'); excluded += 1; continue
            prod, note = matcher.match(name)
            if prod is None:
                flag(f"product '{name}' is not in the product list", 'Excluded'); excluded += 1; continue
            # same invoice + product + qty + amount = the line was exported twice (without invoice numbers,
            # two genuine sales of the same item on the same day would look identical, so they are required)
            sig = (inv, prod['product'], qty, total)
            if inv and sig in seen:
                flag(f'duplicate of invoice {inv} earlier in the same export', 'Excluded'); excluded += 1; continue
            seen.add(sig)
            if note:
                flag(note, 'Kept (name corrected)')
            if qty is not None and price is not None and total is not None and abs(qty * price - total) > 0.5:
                flag(f'total {total:,.2f} does not match {qty:g} x {price:,.2f} = {qty * price:,.2f}', 'Kept, please check')
            if qty is not None and qty < 0:
                flag('return (negative quantity)', 'Kept as a return')
            clean.append({'invoice': inv, 'branch': branch, 'date': d, 'month': d.strftime('%Y-%m'), 'product': prod['product'],
                          'category': prod['category'], 'qty': qty, 'unit_price': price, 'total': total, 'source': path.name})
            kept += 1
        files.append({'file': path.name, 'branch': branch, 'header_row': h + 1, 'lines_kept': kept, 'lines_excluded': excluded,
                      'excluded_amount': round(excluded_amount, 2),
                      'file_total_row': '-' if footer is None else ('matches all lines read' if abs(footer - read_total) < 1 else
                                                                    f'MISMATCH: file says {footer:,.0f}, lines add up to {read_total:,.0f}')})
    return clean, exceptions, files
