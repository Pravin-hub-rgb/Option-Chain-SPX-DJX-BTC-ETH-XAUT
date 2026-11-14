# DeltaX Option Chain

Three Python + xlwings tools that stream live options data into desktop Excel (Windows only) over COM. Each product is built as a standalone PyInstaller executable.

## Products

| Product | Underlying | Feed | Workbooks |
|---|---|---|---|
| **US Option Chain** | SPX, DJX, US stocks | Databento OPRA (live) with Tencent/Yahoo/CBOE fallbacks | `spx_option_chain.xlsx`, `djx_option_chain.xlsx`, `us_stock_option_chain.xlsx` |
| **Delta Option Chain** | BTC, ETH, XAUT | Delta Exchange WebSocket + REST | `btc_chain.xlsx`, `eth_chain.xlsx`, `xaut_chain.xlsx` |
| **MCX Commodity Option Chain** | Gold, Silver, Crude Oil | MCX feed | `gold_chain.xlsx`, `silver_chain.xlsx`, `crudeoil_chain.xlsx` |

## Repository layout

```
.
├── Deltax-option-chain/        # crypto option chain product
│   ├── source/                 # source code, tests, specs, build scripts
│   └── client package/         # built client package (Delta)
├── US option chain/            # US equity/index option chain product
│   ├── source/                 # source code, tests, specs, build scripts
│   └── client package/         # built client package (US)
├── Commodity/                  # MCX commodity option chain product
│   ├── source/                 # source code, tests, specs, build scripts
│   └── probe/                  # Upstox API probe scripts
├── DELIVERY/                   # delivery ZIPs (gitignored)
├── KeyGen (distributor only)/  # license key tooling (gitignored)
├── documents/                  # internal docs (mostly gitignored)
├── package_delivery.ps1        # combined ZIP delivery script (run only when shipping)
├── .gitignore
└── README.md
```

Each `source/` directory is self-contained: it holds the product's `main.py`, module files, tests, `config.json`, `requirements.txt`, a PyInstaller `.spec`, and a `build_exe.bat` / `build_final.bat`.

## Requirements

- Windows + Microsoft Excel
- Python 3.10
- pip dependencies (see each product's `requirements.txt`)

```powershell
cd "US option chain/source"     ; pip install -r requirements.txt
cd "Deltax-option-chain/source"; pip install -r requirements.txt
```

## Self-test

Run the built-in diagnostic before anything else:

```powershell
cd "US option chain/source"     ; python doctor.py
cd "Deltax-option-chain/source"; python doctor.py
```

This verifies Python packages, Excel connectivity, the activation key, the data feed, and workbook integrity, printing a PASS/FAIL for each check.

## Tests

From each product's `source/` directory:

```powershell
cd "US option chain/source"     ; python -m unittest discover -s tests -p "test_*.py" -t .
cd "Deltax-option-chain/source"; python -m unittest discover -s tests -p "test_*.py" -t .
```

Or with pytest:

```powershell
cd "US option chain/source"     ; pytest tests\
cd "Deltax-option-chain/source"; pytest tests\
```

## Build

```powershell
cd "Deltax-option-chain/source"; cmd /c build_final.bat   # builds + packages Delta
cd "US option chain/source"     ; cmd /c build_exe.bat      # builds US
```

After building, delete any delivery ZIPs and verify the `build_info.txt` stamp matches the source hash. See `documents/BUILD_AND_DELIVERY.md` for the full build/delivery checklist.

## Security

- `license.json`, `*.key`, `*.pem`, `.env`, `databento.key`, `KeyGen/` folders, `KEY_ACTIVATION.md`, `understanding/`, `OWNER-ONLY/`, and the `documents/` directory are all **gitignored** and must never be committed or shipped.
- The activation key is entered once per install and stored locally as `license.json`.
- The Databento API key (if used) is stored under `%LOCALAPPDATA%` and is never read by a frozen build.
- The bundled `.env.example` is a placeholder only — paste your key into `.env` (which is gitignored).

## Quick start (US product)

1. Run `python doctor.py` and confirm all checks PASS.
2. Double-click `us_option_chain.exe` (or `python main.py` from source).
3. Enter the activation key, then your Databento API key.
4. Excel opens with three workbooks — leave them open.

## Coexistence

US Option Chain and Delta Option Chain can run simultaneously in the same Excel instance; each manages its own workbooks and calculation-mode lifecycle without fighting the other.

## Links

- Client setup guide: see `HOW-TO-USE.txt` inside each product's `source/` directory.
- Delta chain architecture notes: see `README_CHAIN.md` in `Deltax-option-chain/source/`.
- Options trading primer: see `OPTIONS-BEGINNER-GUIDE.md`.
