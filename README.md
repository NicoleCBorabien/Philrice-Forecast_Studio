# PhilRice Forecast Studio

Two forecast types in one local app:
- Percentage Standard Work Time: weekly forecasts per employee.
- Workload by Specific Tag: monthly allocation across all 11 tags.

## Run on Windows
1. Install Python 3.12 or 3.13 from python.org, including the Python launcher.
2. Extract this whole folder. Double-click RUN_APP.bat.
3. First run installs dependencies and needs internet access. The browser then opens automatically. Forecast calculations run locally.
4. Choose a forecast type, download its template, fill the Data sheet, upload it, choose 1–24 months, and click RUN FORECAST.
5. Download the Excel results when processing completes. Keep the console open while using the app; close it to stop the server.

This is a complete source/build package, not a precompiled Windows EXE. To build an EXE locally, double-click BUILD_EXE.bat. It appears in dist/PhilRice_Forecast_Studio.exe. The included GitHub Actions workflow also builds on Windows; upload the folder contents to the repository root, including .github. Windows compilation and launch were not performed in this Linux environment.

## Data preparation
Use the exact headers from the template. Enter observations only in Data. Guides and examples are ignored. Original template contents and sheet names are unchanged; only download filenames differ.

Weekly: Employee, Week Start, % Standard Work Time. At least 12 consecutive weeks; dates exactly 7 days apart; no duplicates or blanks. Excel Percentage cells recommended: 90% is stored as 0.90. Values above 100% are permitted. Missing weeks are rejected, not converted to zero.

Monthly: Employee, Month, Specific Tag, % of Total Workload. At least 6 consecutive months. Include all 11 tags each month; explicit 0% for zero workload. No duplicate employee/month/tag combinations or missing tag values. Shares should total 100% (validation tolerance 1.5 percentage points, inherited from the original app). Use Excel dates for Month. Short monthly histories are exploratory.

Do not upload an untouched blank template. Completed examples and generated sample results are in examples/.

## Forecasting
Both modes use the same expanded nonseasonal ARIMA search and Holt’s undamped linear trend. ARIMA starts with 12 low-order specifications and adds AR orders 3–6, plus integrated models with drift. Higher orders require sufficient initial training observations. Models must converge and produce finite forecasts. Failed full-history refits are excluded; no silent fallback forecasting method is substituted.

The last ceil(20%) observations are held out for rolling one-step validation. Each prediction uses only preceding observations. Select lowest raw RMSE, then refit on all history. Scores are on the fractional scale (RMSE 0.05 corresponds to 5 percentage points), not multiplied by 100. These validation scores are also used for model selection, not an independent final test.

Weekly forecasts are generated every 7 days through the requested calendar-month horizon. Monthly forecasts are generated once per month. Negative future forecasts are floored at zero. Monthly tag forecasts are normalized across tags to 100%; raw validation RMSE is calculated before normalization. A flat forecast remains possible when supported by the selected model.

All outputs include overview, forecasts, validation, raw RMSE, and charts. No Methodology or Candidate RMSE sheet is included. Files are processed by a server bound to 127.0.0.1. Temporary uploads/results are kept in the operating-system temporary directory. No external service is used for forecasting.
