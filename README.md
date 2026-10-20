# Option Chain — Source Code

Real-time option chain data for cryptocurrency and US equity markets, delivered
as an Excel-native desktop application.

| Component | Markets | Data Source |
|---|---|---|
| **Delta Option Chain** | BTC, ETH, XAUT (Gold) | Delta Exchange REST + WebSocket |
| **US Option Chain** | SPX, DJX, and individual equities | BigClawd / CBOE / Yahoo Finance |

---

## Features

- **Live Excel output** — data is written directly into `.xlsx` workbooks;
  no browser, no separate viewer
- **Three expiry blocks per sheet** — near, medium, and far dated contracts
  side by side
- **Colour-coded strike bands** — ATM highlighted, in-the-money calls and puts
  shaded by moneyness
- **User-adjustable strike range** and expiry selection via in-sheet dropdowns
- **Automatic connection recovery** — polling continues through network
  interruptions and resumes without restart
- **Remote-access detection** — the application terminates if a screen-sharing
  or remote-desktop utility is detected
- **Monthly licensing** — requires a product key on first run and on expiry

---

## Repository Structure

```
Delta Option Chain/
├── source/                    Application source and build inputs
│   ├── main.py                Core logic: market data, Excel writer, licensing
│   ├── config.py              Configuration loader
│   ├── config.json            Runtime settings
│   ├── requirements.txt       Python dependencies
│   └── *_chain.xlsx           Excel workbook templates
└── client package/            End-user documentation and configuration

US Option Chain/
├── source/                    Application source and build inputs
│   ├── main.py                Core logic: market data, Excel writer, licensing
│   ├── config.py              Configuration loader
│   ├── config.json            Runtime settings
│   ├── build_exe.bat          Windows build script (PyInstaller)
│   ├── OptionChain.spec       PyInstaller specification
│   ├── requirements.txt       Python dependencies
│   └── *_option_chain.xlsx    Excel workbook templates
└── client package/            End-user documentation and configuration

```

Owner-only licensing tools and documentation are kept outside this repository
and are never published or shipped to end users.

## Architecture

### Data pipeline

```
Market API  →  in-memory store  →  chain builder  →  Excel writer
```

1. **Ingest** — REST endpoints populate the initial state; a WebSocket stream
   (Delta) and periodic polling apply incremental updates.
2. **Store** — contracts are held in a per-asset dictionary keyed by symbol.
3. **Build** — the chain builder groups contracts by strike, selects a window
   centred on the at-the-money strike, merges the call and put sides, and
   computes the put-call ratio.
4. **Write** — the writer pushes a DataFrame into the workbook, applies number
   formats, and paints the colour bands.

### Sheet layout

Each expiry block occupies nine columns:

```
Call LTP | Call Bid | Call Ask | Call OI | Strike | Put OI | Put Bid | Put Ask | Put LTP
```

Row 1 carries the spot price, strike count, and put-call ratio for the block.
The expiry selector sits in row 3; the strike-range limit sits in `A2`.

### Colour convention

| Colour | Meaning |
|---|---|
| Yellow | At-the-money row |
| Green | In-the-money calls (strike below spot) |
| Red | In-the-money puts (strike above spot) |
| Dark blue | Column headers |
| Light yellow | Editable input cells |

---

## Build Environment

```
Python 3.10.11
Nuitka 4.2.2
PyInstaller (see requirements.txt)
```

Both applications require the **xlwings** package, which drives Excel through
COM automation. Microsoft Excel must be installed on the target machine.

---

## Building

### Delta (Nuitka)

```bash
cd "Delta Option Chain/source"
pip install -r requirements.txt
python -m nuitka --onefile --windows-console-mode=disable \
    --enable-plugin=tk-inter --assume-yes-for-downloads \
    --output-filename=OptionChain.exe main.py
```

### US (PyInstaller)

```bat
cd "US Option Chain\source"
pip install -r requirements.txt
build_exe.bat
```

`build_exe.bat` runs PyInstaller and copies the resulting executable together
with the workbook templates into the client package directory.

> **Note:** Python 3.13 and later may fail to build against the pinned
> dependencies. Use Python 3.10 or 3.11 in that case.

---

## Runtime Requirements

- Windows 10 or later
- Microsoft Excel (desktop edition — COM automation is required)
- Internet connectivity for market data

---

## Configuration

`config.json` exposes the polling and refresh intervals:

| Key | Purpose |
|---|---|
| `refresh_interval_seconds` | Excel write cadence |
| `oi_poll_seconds` | Open interest polling interval |
| `expiry_refresh_seconds` | Expiry list refresh interval |
| `spot_poll_seconds` | Spot price polling interval |
| `http_timeout_seconds` | Network request timeout |

---

## Licensing

End-user builds require a product key on first run and when the license cycle
ends. Activation state is stored in `license.json` beside the executable.
Owner-only key generation and activation documentation are kept outside this
repository.

---

## Secrets

Never commit API keys, `.env` files, activation files, or owner-only licensing
tools. The application reads the Databento key from the Windows user profile,
not from the repository or the distributed package.
