"""PhilRice weekly Standard Work Time forecasting engine."""
from __future__ import annotations
import io
import math
import warnings
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.ticker import PercentFormatter
import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from statsmodels.tsa.arima.model import ARIMA
from statsmodels.tsa.holtwinters import Holt

REQUIRED = ['Employee', 'Week Start', '% Standard Work Time']
ORDERS = [(p,d,q) for d in (0,1) for p in range(3) for q in range(3) if p + q <= 2]
# Longer autoregressive memory can represent recurring weekly behavior without a seasonal model.
ORDERS += [(p,d,0) for d in (0,1) for p in (3,4,5,6)]
def display_model(kind, order=None, extra=None):
    return f'ARIMA {order}' + (' with drift' if extra == 't' else '') if kind == 'ARIMA' else "Holt's Trend"


def read_inputs(path: Path):
    """Read one-row-per-employee-per-week Standard Work Time records."""
    book = pd.read_excel(path, sheet_name=None)
    # Prefer the actual Data sheet; guide/example sheets are never observations.
    if "Data" in book:
        book = {"Data": book["Data"]}
    frames = []
    for sheet, frame in book.items():
        frame.columns = [str(x).strip() for x in frame.columns]
        if set(REQUIRED).issubset(frame.columns):
            subset = frame[REQUIRED].copy().dropna(how='all')
            subset['source_sheet'] = sheet
            frames.append(subset)
    if not frames:
        raise ValueError('No sheet contains Employee, Week Start, % Standard Work Time.')

    df = pd.concat(frames, ignore_index=True)
    if df[['Employee','Week Start']].isna().any(axis=1).any():
        raise ValueError('Employee and Week Start cannot be blank.')
    if df['% Standard Work Time'].isna().any():
        raise ValueError('Blank % Standard Work Time detected. Enter the observed weekly value.')

    df['Employee'] = df['Employee'].astype(str).str.strip()
    if (df['Employee'] == '').any():
        raise ValueError('Employee cannot be blank.')

    dates = pd.to_datetime(df['Week Start'], errors='coerce')
    if dates.isna().any():
        raise ValueError('Some Week Start entries are invalid. Use real Excel dates such as July 1, 2025.')
    df['Week Start'] = dates.dt.normalize()

    raw = df['% Standard Work Time']
    text = raw.astype(str).str.strip()
    perc_text = text.str.endswith('%')
    val = pd.to_numeric(text.str.rstrip('%'), errors='coerce')
    if val.isna().any() or not np.isfinite(val.to_numpy(float)).all():
        raise ValueError('% Standard Work Time must contain valid numbers or percentages.')
    fraction = val.astype(float).copy()
    fraction[perc_text] /= 100.0
    fraction[~perc_text & (val > 5)] /= 100.0
    if (fraction < 0).any():
        raise ValueError('% Standard Work Time cannot be negative.')
    df['swt'] = fraction

    result = {}
    warnings_out = []
    for employee, group in df.groupby('Employee', sort=True):
        group = group.sort_values('Week Start')
        if group.duplicated(['Week Start']).any():
            raise ValueError(f'{employee}: duplicate Week Start entries. Keep only one value per employee per week.')
        if len(group) < 12:
            raise ValueError(f'{employee}: at least 12 consecutive weekly observations are required; found {len(group)}.')
        dates = group['Week Start'].to_numpy(dtype='datetime64[D]')
        gaps = np.diff(dates).astype('timedelta64[D]').astype(int)
        if len(gaps) and not np.all(gaps == 7):
            bad_at = int(np.where(gaps != 7)[0][0])
            a = pd.Timestamp(dates[bad_at]).strftime('%d %b %Y')
            b = pd.Timestamp(dates[bad_at+1]).strftime('%d %b %Y')
            raise ValueError(f'{employee}: weekly dates must be exactly 7 days apart with no missing weeks. Check {a} to {b}.')
        series = pd.Series(group['swt'].to_numpy(float), index=pd.DatetimeIndex(group['Week Start']), name=employee)
        result[employee] = series
        extreme = group.loc[group['swt'] > 2.0, ['Week Start','swt']]
        if not extreme.empty:
            labels = ', '.join(f"{r['Week Start']:%d %b %Y} ({r['swt']:.0%})" for _, r in extreme.head(5).iterrows())
            warnings_out.append(f'{employee}: values above 200% detected — {labels}. Forecasting will continue, but review these entries.')
    return result, warnings_out


def _future(series, kind, horizon, order=None, extra=None):
    series = np.asarray(series, dtype=float)
    with warnings.catch_warnings():
        warnings.simplefilter('ignore')
        if kind == 'ARIMA':
            fit = ARIMA(series, order=order, trend=extra, enforce_stationarity=True, enforce_invertibility=True).fit()
            if not fit.mle_retvals.get('converged', True):
                raise ValueError('ARIMA optimization did not converge')
        elif kind == "Holt's Trend":
            fit = Holt(series, damped_trend=False, initialization_method='estimated').fit(optimized=True)
        else:
            raise ValueError(f'Unsupported forecasting method: {kind}')
        forecast = np.asarray(fit.forecast(horizon), dtype=float)
        if not np.all(np.isfinite(forecast)):
            raise ValueError('Model produced nonfinite forecasts.')
        return forecast


def _fit_predict(series, kind, order=None, extra=None):
    return float(_future(series, kind, 1, order, extra)[0])


def evaluate_series(series, horizon):
    """Compare only ARIMA and undamped Holt trend on identical rolling holdout weeks.

    RMSE is sqrt(mean((actual-prediction)**2)), on the stored fraction scale.
    No multiplication by 100, normalization, or percentage formatting is applied.
    """
    series = np.asarray(series, dtype=float)
    n = len(series)
    validation = max(1, math.ceil(n * 0.20))
    training = n - validation
    actual = series[training:]
    family = dict(arima_rmse=float('inf'), holt_rmse=float('inf'))
    candidates = [('ARIMA', order, trend) for order in ORDERS
                  if training >= max(4, 3 * (order[0] + order[2] + 1))
                  for trend in ([None, 't'] if order[1] == 1 else [None])]
    candidates.append(("Holt's Trend", None, None))
    ranked = []
    scores = []
    for kind, order, extra in candidates:
        try:
            preds = np.asarray([_fit_predict(series[:at], kind, order, extra) for at in range(training, n)])
            score = float(np.sqrt(np.mean((actual-preds)**2)))
            if not np.isfinite(score):
                raise ValueError('Nonfinite RMSE')
            key = 'arima_rmse' if kind == 'ARIMA' else 'holt_rmse'
            family[key] = min(family[key], score)
            ranked.append((score, kind, order, extra, preds))
            scores.append((display_model(kind, order, extra), score, 'Validated'))
        except Exception as exc:
            scores.append((display_model(kind, order, extra), None, str(exc)))
    # Full-history refits must succeed. Never silently substitute a different method.
    for score, kind, order, extra, preds in sorted(ranked, key=lambda x: x[0]):
        try:
            forecast = _future(series, kind, horizon, order, extra)
        except Exception:
            continue
        return dict(model=kind, order=order, extra=extra, **family, selected_rmse=score,
                    validation_actual=actual, validation_pred=preds,
                    forecast=np.maximum(0, forecast), train_weeks=training,
                    validation_weeks=validation, candidate_scores=scores)
    raise ValueError('No ARIMA or Holt trend candidate could be validated and refitted.')


def _weekly_dates_for_months(last_date, months):
    start = pd.Timestamp(last_date) + pd.Timedelta(days=7)
    end = pd.Timestamp(last_date) + pd.DateOffset(months=int(months))
    dates = pd.date_range(start=start, end=end, freq='7D')
    if len(dates) == 0:
        dates = pd.DatetimeIndex([start])
    return dates


def run_all(datasets, forecast_months, report):
    """Forecast each employee independently; output remains weekly for the requested calendar-month horizon."""
    employees = list(datasets)
    total = len(employees)
    result = {}
    done = 0
    def job(emp):
        series = datasets[emp]
        dates = _weekly_dates_for_months(series.index[-1], forecast_months)
        fit = evaluate_series(series.to_numpy(float), len(dates))
        return emp, {'history':series, 'fit':fit, 'dates':dates, 'forecast_months':int(forecast_months)}
    with ThreadPoolExecutor(max_workers=min(4, max(1,total))) as pool:
        futures = {pool.submit(job, emp): emp for emp in employees}
        for future in as_completed(futures):
            emp, item = future.result()
            result[emp] = item
            done += 1
            report(done, total, f'{emp} ({done}/{total})')
    return dict(sorted(result.items()))


def make_chart_image(dates, lines, title, width=11.2, height=4.6, legend=True):
    fig, ax = plt.subplots(figsize=(width,height), dpi=125)
    x = np.arange(len(dates))
    for name, values in lines.items():
        ax.plot(x, np.asarray(values,float), linewidth=1.9, marker='o' if len(x)<=20 else None, markersize=3, label=name)
    finite = np.concatenate([np.asarray(v,float).ravel() for v in lines.values()])
    finite = finite[np.isfinite(finite)]
    lo, hi = (float(finite.min()), float(finite.max())) if finite.size else (0.0,1.0)
    span = hi-lo
    pad = max(0.015, span*0.15)
    y0 = max(0.0, lo-pad); y1 = hi+pad
    # Keep the axis tight enough that small but real weekly movements remain visible.
    if y1-y0 < 0.06:
        center=(y0+y1)/2; y0=max(0,center-0.03); y1=center+0.03
    ax.set_ylim(y0,y1)
    ax.set_title(title, fontsize=12, weight='bold', pad=10)
    ax.set_ylabel('% Standard Work Time')
    ax.set_xlabel('Week Start')
    ax.yaxis.set_major_formatter(PercentFormatter(xmax=1.0, decimals=0))
    every=max(1, math.ceil(len(x)/12))
    ticks=np.arange(0,len(x),every)
    if len(x) and len(x)-1 not in ticks: ticks=np.append(ticks,len(x)-1)
    labels=[pd.Timestamp(d).strftime('%d %b %Y') for d in dates]
    ax.set_xticks(ticks,[labels[i] for i in ticks],rotation=35,ha='right',fontsize=8.5)
    ax.grid(axis='y',alpha=.22); ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
    if legend and len(lines) > 1:
        ax.legend(loc='upper center',bbox_to_anchor=(.5,-.27),ncol=3,frameon=False,fontsize=8)
    fig.subplots_adjust(left=.10,right=.97,top=.87,bottom=.28)
    stream=io.BytesIO(); fig.savefig(stream,format='png',dpi=125,facecolor='white'); plt.close(fig); stream.seek(0)
    img=ExcelImage(stream); img._philrice_buffer=stream; return img


def _raw_rmse(value):
    return None if value is None or np.isinf(value) else float(value)


def write_workbook(results, warnings_list, path):
    wb=Workbook(); ov=wb.active; ov.title='Overview'
    ov.append(['PhilRice Standard Work Time Forecast — Validation and Weekly Forecasts'])
    ov.append(['Employee','History Weeks','Train Weeks','Validation Weeks','Forecast Months','Forecast Weeks','Selected Model',
               'Best ARIMA Raw RMSE',"Holt's Trend Raw RMSE",'Selected Raw RMSE','Historical Average','Latest Actual'])
    for employee,item in results.items():
        s=item['history']; f=item['fit']
        ov.append([employee,len(s),f['train_weeks'],f['validation_weeks'],item['forecast_months'],len(item['dates']),display_model(f['model'],f.get('order'),f.get('extra')),
                   _raw_rmse(f['arima_rmse']),_raw_rmse(f['holt_rmse']),_raw_rmse(f['selected_rmse']),float(s.mean()),float(s.iloc[-1])])
    if warnings_list:
        warn=wb.create_sheet('Input Warnings'); warn.append(['Warning'])
        for w in warnings_list: warn.append([w])
    for employee,item in results.items():
        s=item['history']; f=item['fit']; dates=item['dates']
        safe=''.join(c if c not in '[]:*?/\\' else '_' for c in employee)[:23]
        sh=wb.create_sheet(safe+' Forecast')
        sh.append([f'Weekly % Standard Work Time Forecast — {employee}'])
        sh.append(['Week Start','Forecast % Standard Work Time'])
        for d,v in zip(dates,f['forecast']):
            sh.append([d.to_pydatetime(),float(v)]); sh.cell(sh.max_row,1).number_format='dd mmm yyyy'; sh.cell(sh.max_row,2).number_format='0.0%'
        # Forecast chart intentionally shows forecast only, with a tight y-axis for clearer trend/fluctuation visibility.
        sh.add_image(make_chart_image(dates,{'Forecast':f['forecast']},f'{employee} — Forecast Only',legend=False),f'A{sh.max_row+4}')

        val=wb.create_sheet(safe+' Validation')
        model_label=display_model(f['model'],f.get('order'),f.get('extra'))
        val.append([f'Rolling Validation — {employee}'])
        val.append(['Week Start','Actual','Validation Forecast','Selected Model','Selected Raw RMSE'])
        vd=s.index[-f['validation_weeks']:]
        for d,a,p in zip(vd,f['validation_actual'],f['validation_pred']):
            val.append([d.to_pydatetime(),float(a),float(p),model_label,_raw_rmse(f['selected_rmse'])])
            val.cell(val.max_row,1).number_format='dd mmm yyyy'; val.cell(val.max_row,2).number_format='0.0%'; val.cell(val.max_row,3).number_format='0.0%'; val.cell(val.max_row,5).number_format='0.0000000000'
        val.add_image(make_chart_image(vd,{'Actual':f['validation_actual'],'Validation Forecast':f['validation_pred']},
                                       f'{employee} — Validation ({model_label})'),f'G3')

    for sh in wb:
        sh.freeze_panes='B3'
        for col in sh.columns:
            letter=get_column_letter(col[0].column)
            sh.column_dimensions[letter].width=min(48,max(15,max(len(str(c.value or '')) for c in list(col)[:min(150,len(col))])+2))
        for c in sh[2]:
            c.fill=PatternFill('solid',fgColor='244D71'); c.font=Font(color='FFFFFF',bold=True)
    # Raw RMSE retains the model target's fractional scale.
    for row in ov.iter_rows(min_row=3):
        for col_idx in range(8,11):
            row[col_idx-1].number_format='0.0000000000'
        row[10].number_format='0.0%'
        row[11].number_format='0.0%'
    wb.save(path)
