# Option Chain — Source Backup

Delta (BTC/ETH/XAUT) + US (SPX/DJX/Stocks) — dono ka complete source code.

## Ye kya hai

Ye folder **sirf source code** hai. Client ko jo bheja gaya wo
`DELIVERY\` folder me hai (ZIP files).

Ye backup isliye hai ki agar future me kuch repair karna ho —
ya Delta Exchange ka API change kar de — to **rebuild** kiya ja sake.
Bina iske, naya exe nahi ban sakta.

## Folder structure

```
Delta Option Chain\
   source\              main.py (license + API logic), config.py, templates
   client package\      docs + config (exe nahi — rebuild hoga)

US Option Chain\
   source\              main.py, config.py, build_exe.bat, OptionChain.spec
   client package\      docs + config

KeyGen (distributor only)\
   KeyGen.exe           (already built — backup me hai)
   keygen.py            CLI key generator
   keygen_gui.py        GUI key generator (KeyGen.exe ka source)

understanding\
   KEY_ACTIVATION.md    licensing system ka pura documentation
```

## Rebuild kaise karein

### Delta (Nuitka)
```bash
cd "Delta Option Chain/source"
pip install -r requirements.txt
python -m nuitka --onefile --windows-console-mode=disable ^
    --enable-plugin=tk-inter --assume-yes-for-downloads ^
    --output-filename=OptionChain.exe main.py
```
Output: `OptionChain.exe` (~39 MB)

### US (PyInstaller)
```bat
cd "US Option Chain\source"
pip install -r requirements.txt
build_exe.bat
```
`build_exe.bat` PyInstaller chalata hai aur `..\client package\` me
exe + xlsx templates copy kar deta hai.

### KeyGen (Nuitka)
```bash
cd "KeyGen (distributor only)"
python -m nuitka --onefile --windows-console-mode=disable ^
    --enable-plugin=tk-inter --assume-yes-for-downloads ^
    --output-filename=KeyGen.exe keygen_gui.py
```

## Build environment (jo use kiya gaya tha)

```
Python 3.10.11
Nuitka 4.2.2
PyInstaller (requirements.txt me)
```

Naya Python version (3.13+) pe build fail ho sakta hai —
tab `pip install "python<3.12"` karke 3.10/3.11 use karo.

## IMPORTANT — kya NAHI karna

- `KeyGen.exe` kabhi public repo me mat daalo
- Ye repo PRIVATE rakho (GitHub private / local git)
- Source me secret hai (XOR-encoded) — isliye private zaroori hai

## License system

Details: `understanding/KEY_ACTIVATION.md`

Cycle = har mahine ka last Friday.
Key = 100 variants per cycle, sab valid.
Accept rule: `today < date <= today + 45 days`.
