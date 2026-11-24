"""Write the finished report: Summary (KPIs, branch x month, top products, categories, file checks),
Charts (native Excel charts), Clean data (filterable Excel table) and Exceptions."""
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

MONEY = '₹#,##0'
NAVY, LIGHT = '1F3864', 'D9E1F2'


def _table(ws, df: pd.DataFrame, top: int, left: int = 1, title: str | None = None, money=(), pct=()) -> int:
    if title:
        ws.cell(top, left, title).font = Font(bold=True, size=12, color=NAVY)
        top += 1
    for j, col in enumerate(df.columns):
        c = ws.cell(top, left + j, col)
        c.font, c.fill = Font(bold=True, color='FFFFFF'), PatternFill('solid', fgColor=NAVY)
    for i, row in enumerate(df.itertuples(index=False), 1):
        for j, v in enumerate(row):
            c = ws.cell(top + i, left + j, None if pd.isna(v) else v)
            name = df.columns[j]
            c.number_format = MONEY if name in money else ('0.0%' if name in pct else c.number_format)
    return top + len(df) + 2


def build(clean: list[dict], exceptions: list[dict], files: list[dict], out: Path) -> dict:
    df = pd.DataFrame(clean)
    months = sorted(df.month.unique())
    last, prev = months[-1], (months[-2] if len(months) > 1 else None)
    pivot = df.pivot_table(index='branch', columns='month', values='total', aggfunc='sum').reindex(columns=months)
    pivot.loc['Total'] = pivot.sum()
    pivot['Total'] = pivot[months].sum(axis=1)
    if prev:
        pivot[f'{last} vs {prev}'] = pivot[last] / pivot[prev] - 1
    pivot = pivot.reset_index()
    top = (df.groupby(['product', 'category']).agg(units=('qty', 'sum'), sales=('total', 'sum')).reset_index()
           .sort_values('sales', ascending=False))
    cats = df.groupby('category').agg(sales=('total', 'sum')).reset_index().sort_values('sales', ascending=False)
    cats['share'] = cats.sales / cats.sales.sum()
    returns = df[df.qty < 0].total.sum()
    kpi = {'Net sales': df.total.sum(), 'Sales lines': len(df), 'Units': int(df.qty.sum()),
           'Average line value': df.total.mean(), 'Returns': returns}

    wb = Workbook()
    ws = wb.active
    ws.title = 'Summary'
    ws['A1'] = f'Sales report: {months[0]} to {last}'
    ws['A1'].font = Font(bold=True, size=16, color=NAVY)
    ws['A2'] = f'{len(files)} branch exports combined; {len(exceptions)} items in the Exceptions sheet.'
    for j, (k, v) in enumerate(kpi.items()):
        ws.cell(4, 1 + j * 2, k).font = Font(bold=True, color='595959')
        c = ws.cell(5, 1 + j * 2, v)
        c.font = Font(bold=True, size=14)
        c.number_format = MONEY if k in ('Net sales', 'Average line value', 'Returns') else '#,##0'
    row = _table(ws, pivot, 8, title='Sales by branch and month', money=months + ['Total'], pct=[c for c in pivot.columns if ' vs ' in str(c)])
    if prev:
        col = get_column_letter(len(pivot.columns))
        rng = f'{col}10:{col}{9 + len(pivot)}'
        ws.conditional_formatting.add(rng, CellIsRule(operator='greaterThan', formula=['0'], font=Font(color='006100'), fill=PatternFill('solid', fgColor='C6EFCE')))
        ws.conditional_formatting.add(rng, CellIsRule(operator='lessThan', formula=['0'], font=Font(color='9C0006'), fill=PatternFill('solid', fgColor='FFC7CE')))
    row = _table(ws, top.head(8), row, title='Top products', money=['sales'])
    row = _table(ws, cats, row, title='Sales by category', money=['sales'], pct=['share'])
    _table(ws, pd.DataFrame(files), row, title='Files read', money=['excluded_amount'])
    for r in ws.iter_rows(min_row=10, max_row=9 + len(pivot)):
        if r[0].value == 'Total':
            for c in r:
                c.font = Font(bold=True)
    for col, w in zip('ABCDEFGHIJ', (24, 16, 14, 14, 16, 18, 24, 14, 14, 14)):
        ws.column_dimensions[col].width = w

    ch = wb.create_sheet('Charts')
    _table(ch, pivot[pivot.branch != 'Total'][['branch'] + months], 1, money=months)
    bar = BarChart(); bar.type = 'col'; bar.title = 'Sales by branch and month'; bar.y_axis.title = 'INR'
    bar.add_data(Reference(ch, min_col=2, max_col=1 + len(months), min_row=1, max_row=1 + len(pivot) - 1), titles_from_data=True)
    bar.set_categories(Reference(ch, min_col=1, min_row=2, max_row=len(pivot)))
    bar.width, bar.height = 18, 9
    bar.x_axis.delete = bar.y_axis.delete = False   # openpyxl hides axes by default
    bar.y_axis.numFmt = '#,##0'
    bar.y_axis.majorGridlines = None
    ch.add_chart(bar, 'H2')
    tot = df.groupby('month').total.sum().reset_index()
    start = len(pivot) + 4
    _table(ch, tot, start, money=['total'])
    line = LineChart(); line.title = 'Total sales per month'; line.legend = None; line.varyColors = False
    line.add_data(Reference(ch, min_col=2, min_row=start, max_row=start + len(tot)), titles_from_data=True)
    line.set_categories(Reference(ch, min_col=1, min_row=start + 1, max_row=start + len(tot)))
    line.width, line.height = 18, 8
    line.x_axis.delete = line.y_axis.delete = False
    line.y_axis.numFmt = '#,##0'
    for srs in line.series:
        srs.smooth = False
        srs.marker.symbol = 'circle'
    ch.add_chart(line, 'H22')

    cd = wb.create_sheet('Clean data')
    cols = ['invoice', 'branch', 'date', 'month', 'product', 'category', 'qty', 'unit_price', 'total', 'source']
    cd.append(cols)
    for r in clean:
        cd.append([r[c] for c in cols])
    t = Table(displayName='SalesLines', ref=f'A1:{get_column_letter(len(cols))}{len(clean) + 1}')
    t.tableStyleInfo = TableStyleInfo(name='TableStyleMedium2', showRowStripes=True)
    cd.add_table(t)
    for i, w in enumerate((11, 10, 12, 9, 24, 12, 7, 11, 12, 20), 1):
        cd.column_dimensions[get_column_letter(i)].width = w
    for r in cd.iter_rows(min_row=2):
        r[2].number_format = 'dd-mmm-yyyy'
        r[7].number_format = r[8].number_format = MONEY
    cd.freeze_panes = 'A2'

    ex = wb.create_sheet('Exceptions')
    exdf = pd.DataFrame(exceptions, columns=['file', 'row', 'branch', 'issue', 'action', 'values'])
    _table(ex, exdf, 1)
    for col, w in zip('ABCDEF', (20, 6, 9, 60, 22, 60)):
        ex.column_dimensions[col].width = w
    for r in ex.iter_rows(min_row=2):
        for c in r:
            c.alignment = Alignment(vertical='top', wrap_text=True)
    wb.save(out)

    growth = pivot.loc[pivot.branch == 'Total', f'{last} vs {prev}'].iloc[0] if prev else None
    best = pivot[pivot.branch != 'Total'].sort_values(last, ascending=False).iloc[0]
    return {'kpi': kpi, 'growth': growth, 'best_branch': best.branch, 'best_sales': best[last], 'last': last, 'prev': prev,
            'top_product': top.iloc[0]['product'], 'exceptions': len(exceptions)}
