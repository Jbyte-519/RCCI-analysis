# RCCI Dahsboard2 (React)

This is a separate React/Vite dashboard. It does not modify `RUN_1/rcci_gui.py` or the original `RCCI_REACT_DASHBOARD`.

## Start

From the RCCI project directory, in two PowerShell terminals:

```powershell
.\.venv\Scripts\python.exe .\Dahsboard2\server.py
cd .\Dahsboard2
npm install
npm run dev
```

Open `http://127.0.0.1:5180`. The API uses port `8766` so it can run beside the existing dashboard.

Upload the instrument workbook, consecutive-cycle workbook, and performance report. Choose a report folder in the dashboard before running; Edge/Chrome writes the generated HTML, Excel, CSV, and trace files there. The API also keeps a local copy under `Dahsboard2/results`. The API selects the report row using the `LoadXX Nm` token in the selected filename and delegates calculations to the existing analyzer, including reports and plots. Blank measurements remain unavailable rather than being replaced with zero.
