# Active RCCI workspace

Open this folder in VS Code:

```text
C:\Users\Joel Babu\Desktop\Maj Pro\RCCI RUN DATA
```

The dashboards are inside its `RCCI` directory:

```text
RCCI RUN DATA
├─ .venv
└─ RCCI
   ├─ Dahsboard2   React dashboard
   ├─ Dahsboard3   CustomTkinter dashboard
   └─ 1, 2, ...    experimental run folders
```

React dashboard (two terminals):

```powershell
cd "C:\Users\Joel Babu\Desktop\Maj Pro\RCCI RUN DATA"
.\.venv\Scripts\python.exe .\RCCI\Dahsboard2\server.py
cd .\RCCI\Dahsboard2
npm run dev
```

Open `http://127.0.0.1:5180`.

CustomTkinter dashboard:

```powershell
cd "C:\Users\Joel Babu\Desktop\Maj Pro\RCCI RUN DATA"
.\.venv\Scripts\python.exe .\RCCI\Dahsboard3\rcci_customtkinter.py
```

The dashboards use a local copy of the analyzer (`rcci_engine.py`) and do not depend on a `RUN_1` folder. Upload the run files from folders such as `RCCI\1` through the dashboard file pickers.
