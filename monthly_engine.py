"""Generic PhilRice Specific Tag forecasting engine (chronological 80/20 validation)."""
from __future__ import annotations
import math
import warnings
from pathlib import Path
import numpy as np
import pandas as pd
from statsmodels.tsa.arima.model import ARIMA
from statsmodels.tsa.holtwinters import Holt
from weekly_engine import evaluate_series as shared_evaluate, display_model
from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.drawing.image import Image as ExcelImage
import io
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.ticker import PercentFormatter, MultipleLocator
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

REQUIRED = ['Employee', 'Month', 'Specific Tag', '% of Total Workload']
CANONICAL_TAGS = ['core research function', 'core administration function', 'extension services', 'ad hoc', 'Research Enablement', 'RSS', 'SMART Farm', 'RCEF', 'RiceBIS', 'Malusog Rice', 'Special Project Support']
# Full intended ARIMA grid: p=0..2, d=0..1, q=0..2, with p+q<=2 (12 orders).
ORDERS = [(p,d,q) for d in (0,1) for p in range(3) for q in range(3) if p + q <= 2]


def read_inputs(path: Path):
    """Read long-form monthly workload records from any Excel sheet with required headers."""
    book = pd.read_excel(path, sheet_name=None)
    if "Data" in book:
        book = {"Data": book["Data"]}
    frames = []
    for sheet, frame in book.items():
        frame.columns = [str(x).strip() for x in frame.columns]
        if set(REQUIRED).issubset(frame.columns):
            subset = frame[REQUIRED].copy()
            subset = subset.dropna(how='all')
            subset['source_sheet'] = sheet
            frames.append(subset)
    if not frames:
        raise ValueError('No sheet contains Employee, Month, Specific Tag, % of Total Workload. Use the included template.')
    df = pd.concat(frames, ignore_index=True)
    # Do not discard rows with blank workload shares silently: require an explicit 0%.
    missing_meta = df[['Employee','Month','Specific Tag']].isna().any(axis=1)
    if missing_meta.any():
        raise ValueError('Employee, Month and Specific Tag cannot be blank. Remove incomplete rows.')
    if df['% of Total Workload'].isna().any():
        raise ValueError('Blank workload % detected. Enter 0% for tags with no workload.')
    df['Employee'] = df['Employee'].astype(str).str.strip()
    df['Specific Tag'] = df['Specific Tag'].astype(str).str.strip()
    tag_lookup = {t.casefold(): t for t in CANONICAL_TAGS}
    df['Specific Tag'] = df['Specific Tag'].map(lambda t: tag_lookup.get(t.casefold(), t))
    dates = pd.to_datetime(df['Month'], errors='coerce')
    if dates.isna().any():
        raise ValueError('Some Month entries are invalid. Use dates such as January 2025.')
    df['Month'] = dates.dt.to_period('M').dt.to_timestamp()
    raw = df['% of Total Workload']
    text = raw.astype(str).str.strip()
    perc_text = text.str.endswith('%')
    val = pd.to_numeric(text.str.rstrip('%'), errors='coerce')
    if val.isna().any() or not np.isfinite(val.to_numpy(float)).all() or (val < 0).any():
        raise ValueError('Workload percentage must be a valid nonnegative number.')
    # Excel % cells are fractions; textual values like 35% are percentage points.
    fraction = val.copy()
    fraction[perc_text] /= 100
    # Ambiguous numeric values between 0 and 1 are treated as fractions.
    fraction[~perc_text & (val > 1)] /= 100
    if (fraction > 1.0001).any():
        raise ValueError('Workload percentages must not exceed 100%.')
    df['share'] = fraction.astype(float)
    result = {}
    for employee, group in df.groupby('Employee', sort=True):
        if group.duplicated(['Month','Specific Tag']).any():
            raise ValueError(f'{employee}: duplicate Month and Specific Tag entries; combine them before uploading.')
        wide = group.pivot(index='Month', columns='Specific Tag', values='share').sort_index()
        # Every monthly dataset includes ALL eleven PhilRice Specific Tags, even if
        # the source omitted their rows. Those missing tag/month pairs mean 0%.
        # Explicit zero workload values are preserved, never excluded from modeling.
        columns = CANONICAL_TAGS + sorted(set(wide.columns) - set(CANONICAL_TAGS))
        wide = wide.reindex(columns=columns)
        if wide.isna().any().any():
            raise ValueError(f'{employee}: include all 11 Specific Tags for every month; enter explicit 0% values where applicable.')
        expected = pd.date_range(wide.index[0], wide.index[-1], freq='MS')
        if len(wide) < 6:
            raise ValueError(f'{employee}: at least 6 monthly observations are required; found {len(wide)}.')
        if len(expected) != len(wide):
            raise ValueError(f'{employee}: missing months in historical series. Include every month consecutively.')
        totals = wide.sum(axis=1)
        if ((totals - 1).abs() > .015).any():
            months = ', '.join(wide.index[(totals-1).abs() > .015].strftime('%b %Y')[:4])
            raise ValueError(f'{employee}: workload shares should add to 100% per month. Check {months}.')
        result[employee] = wide
    return result


def evaluate_series(series, horizon):
    fit = shared_evaluate(series, horizon)
    fit['train_months'] = fit['train_weeks']
    fit['validation_months'] = fit['validation_weeks']
    fit['model'] = display_model(fit['model'], fit['order'], fit['extra'])
    # Label already contains the ARIMA order and drift specification.
    fit['order'] = None
    return fit


def run_all(datasets, horizon, report):
    """Forecast each employee/tag, updating the supplied progress callback after every tag."""
    jobs = sum(len(m.columns) for m in datasets.values())
    done = 0; result = {}
    for employee, matrix in datasets.items():
        tags = {}
        for tag in matrix.columns:
            tags[tag] = evaluate_series(matrix[tag].to_numpy(float), horizon)
            done += 1
            report(done, jobs, f'{employee} — {tag} ({done}/{jobs})')
        dates = pd.date_range(matrix.index[-1] + pd.offsets.MonthBegin(1), periods=horizon, freq='MS')
        raw = np.column_stack([tags[t]['forecast'] for t in matrix.columns])
        total = raw.sum(axis=1)
        for i in range(horizon):
            if total[i] > 0: raw[i,:] /= total[i]
            else: raw[i,:] = matrix.iloc[-1].to_numpy(float)
        for j,tag in enumerate(matrix.columns):
            tags[tag]['forecast_normalized'] = raw[:,j]
        result[employee] = {'history':matrix, 'tags':tags, 'dates':dates}
    return result



def make_chart_image(dates, lines, title, every=1, width=11.8, height=4.9):
    """Render a legible Excel-compatible chart with explicit date and percent ticks."""
    fig, ax = plt.subplots(figsize=(width, height), dpi=130)
    x = np.arange(len(dates))
    for name, values in lines.items():
        ax.plot(x, np.asarray(values, dtype=float), linewidth=1.7, marker='o' if len(x)<13 else None,
                markersize=3, label=name)
    ax.set_xlim(-0.3, max(1,len(x)-0.7))
    # Use an adaptive y-axis so genuine month-to-month movement is visible.
    # The previous fixed 0-100% axis visually flattened small but meaningful changes.
    finite = np.concatenate([np.asarray(v, dtype=float).ravel() for v in lines.values()]) if lines else np.array([0.0])
    finite = finite[np.isfinite(finite)]
    if finite.size:
        lo = max(0.0, float(finite.min()))
        hi = min(1.0, float(finite.max()))
    else:
        lo, hi = 0.0, 1.0
    span = hi - lo
    pad = max(0.02, span * 0.15)
    y0 = max(0.0, lo - pad)
    y1 = min(1.0, hi + pad)
    if y1 - y0 < 0.10:
        center = (y0 + y1) / 2
        y0 = max(0.0, center - 0.05)
        y1 = min(1.0, center + 0.05)
    if y1 <= y0:
        y0, y1 = 0.0, 0.10
    ax.set_ylim(y0, y1)
    ax.set_title(title, fontsize=12, weight='bold', pad=12)
    ax.set_ylabel('% of Total Workload', fontsize=10, labelpad=9)
    ax.set_xlabel('Month and Year', fontsize=10, labelpad=9)
    tick_step = 0.05 if (y1-y0) <= 0.35 else 0.10
    ax.yaxis.set_major_locator(MultipleLocator(tick_step))
    ax.yaxis.set_major_formatter(PercentFormatter(xmax=1.0, decimals=0))
    tick_idx = np.arange(0, len(x), max(1, every))
    if len(x) and len(x)-1 not in tick_idx: tick_idx = np.append(tick_idx,len(x)-1)
    labels = [pd.Timestamp(d).strftime('%b %Y') for d in dates]
    ax.set_xticks(tick_idx, [labels[i] for i in tick_idx], rotation=35, ha='right', fontsize=9)
    ax.tick_params(axis='y', labelsize=9, labelleft=True)
    ax.tick_params(axis='x', labelbottom=True, bottom=True)
    ax.grid(axis='y', alpha=.22)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.legend(loc='upper center', bbox_to_anchor=(.5,-.31), ncol=3,
              frameon=False, fontsize=7.5)
    fig.subplots_adjust(left=.105, right=.97, top=.87, bottom=.29)
    stream = io.BytesIO()
    fig.savefig(stream, format='png', dpi=130, facecolor='white')
    plt.close(fig)
    stream.seek(0)
    img = ExcelImage(stream)
    img._philrice_buffer = stream
    return img

def write_workbook(results, path):
    """Create client-facing Excel forecast, model-validation tables, and labeled charts."""
    wb = Workbook(); overview = wb.active; overview.title = 'Overview'
    overview.append(['PhilRice Workload Forecast — Model Validation and Future Shares'])
    overview.append(['Employee','Specific Tag','History Months','Train Months','Validation Months','Model','ARIMA Raw RMSE','Holt Raw RMSE','Selected Raw RMSE'])
    for employee, item in results.items():
        matrix = item['history']; dates = item['dates']; tags = item['tags']
        for tag, fit in tags.items():
            overview.append([employee,tag,len(matrix),fit['train_months'],fit['validation_months'],
                             fit['model']+(f' {fit["order"]}' if fit['order'] else ''),
                             None if np.isinf(fit['arima_rmse']) else fit['arima_rmse'],
                             None if np.isinf(fit['holt_rmse']) else fit['holt_rmse'],fit['selected_rmse']])
        safe = ''.join(c if c not in '[]:*?/\\' else '_' for c in employee)[:23]
        sh = wb.create_sheet(safe + ' Forecast')
        sh.append([f'Forecast — {employee}']); sh.append(['Month'] + list(matrix.columns))
        for i,d in enumerate(dates):
            sh.append([d.strftime('%b %Y')] + [float(tags[t]['forecast_normalized'][i]) for t in matrix.columns])
            # Month/year is stored as a readable chart category label.
            for c in range(2,sh.max_column+1): sh.cell(sh.max_row,c).number_format='0.0%'
        forecast_lines = {tag: tags[tag]['forecast_normalized'] for tag in matrix.columns}
        interval = max(1, math.ceil(len(dates)/12))
        image = make_chart_image(dates,forecast_lines,
                                 f'{employee} — Monthly Specific Tag Forecast', every=interval)
        sh.add_image(image, f'A{sh.max_row+4}')
        val=wb.create_sheet(safe + ' Validation'); val.append([f'Historical actual vs rolling forecast — {employee}'])
        val.append(['Specific Tag','Month','Actual','Validation Forecast','Selected Model','Raw RMSE'])
        for tag, fit in tags.items():
            vd=matrix.index[-fit['validation_months']:]
            for i,d in enumerate(vd):
                val.append([tag,d.to_pydatetime(),float(fit['validation_actual'][i]),float(fit['validation_pred'][i]),fit['model'],fit['selected_rmse']])
                val.cell(val.max_row,2).number_format='mmmm yyyy'
                val.cell(val.max_row,3).number_format='0.0%';val.cell(val.max_row,4).number_format='0.0%'
        start=val.max_row+3
        for tag, fit in tags.items():
            val.cell(start,1, f'Validation Chart — {tag} (including zero-series)')
            val.cell(start+1,1,'Month'); val.cell(start+1,2,'Actual');val.cell(start+1,3,'Forecast')
            vd=matrix.index[-fit['validation_months']:]
            for i,d in enumerate(vd):
                val.cell(start+2+i,1,d.strftime('%b %Y'))
                val.cell(start+2+i,2,float(fit['validation_actual'][i])).number_format='0.0%'
                val.cell(start+2+i,3,float(fit['validation_pred'][i])).number_format='0.0%'
            image = make_chart_image(vd, {
                'Actual':fit['validation_actual'], 'Forecast':fit['validation_pred']},
                f'{employee} — {tag} — Validation', every=max(1, math.ceil(len(vd)/10)),
                width=9.5, height=4.5)
            val.add_image(image, f'H{start}')
            start += max(32,len(vd)+6)
    for sh in wb:
        sh.freeze_panes='B3'
        for col in sh.columns:
            letter=get_column_letter(col[0].column)
            sh.column_dimensions[letter].width=min(48,max(15,max(len(str(c.value or '')) for c in list(col)[:min(150,len(col))])+2))
        for cell in sh[2]:
            cell.fill=PatternFill('solid', fgColor='244D71');cell.font=Font(color='FFFFFF',bold=True)
    for row in overview.iter_rows(min_row=3):
        for cell in row[6:9]: cell.number_format='0.0000000000'
    for sheet in wb:
        if sheet.title.endswith(' Validation'):
            for row in sheet.iter_rows(min_row=3, max_col=6):
                if isinstance(row[5].value, (int,float)): row[5].number_format='0.0000000000'
    wb.save(path)
