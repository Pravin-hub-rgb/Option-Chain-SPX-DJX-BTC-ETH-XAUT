# Key Activation — Owner Guide (DeltaX Option Chain)

> **Sirf tumhare liye** — yeh folder kabhi client zip me mat daalo.
> Client ko sirf ye jaata hai: `OptionChain.exe` + 3 xlsx + `config.json` + README/instruction/FEATURES.

---

## 1. Key ka formula (hashed / opaque / multi-variant)

Key me **date seedha dikhti nahi** — layers hain:

```
raw     = int(DDMMYYYY of cycle-end date)          # jaise 30/10/2026 -> 30102026
token   = raw XOR mask8                            # mask8 = SHA256(secret) se bana
variant = 0-99                                     # har click alag number
vmask   = SHA256(secret, variant)                  # per-variant scramble
key     = (token XOR vmask)*10000 + variant*100 + SHA256(secret, token, variant)[:2]
```

- **Cycle-end date** (validity ka last din):
  - **Delta (BTC/ETH/XAUT)** = mahine ki **last Friday**
  - **US (S&P 500)** = mahine ki **last Friday** (US market close ke baad expire)
- **100 variants per date**: har variant ka **POORA alag number** dikhta hai
  (XOR scramble — koi shared prefix nahi), **sab valid** (app sab accept karti hai)
  - Keygen screen par number hamesha **13 digit** ka dikhata hai (zero-padded).
    Ye sirf dikhane ka format hai — app `int()` se pad bhul jata hai, chalta hi hai.
- Date hidden (double XOR) + SHA256 check (variant ke saath) — bina secret ke
  **har variant ka** guess/forge impossible
- **Secret dono tools me SAME** (Delta + US) → **ek hi key dono me chalti hai**

### Secret kahan hai
- Live value **intentionally NOT written here** — is file me koi plaintext secret
  nahi hona chahiye. Jise ye value chahiye, sirf owner ke paas rakho.
- Source me `_SE` (XOR-encoded byte list) + `_SK` (XOR mask) se runtime pe decode
  hota hai — exe me `strings` chalao toh secret/offset **dikhta nahi**
- ⚠️ Ye `understanding/` folder CLIENT ko KABHI mat dena. Isme likha hua koi bhi
  secret = unlimited key generation.

### Examples (cycle-end = last Friday — dono tools ke liye SAME rule)
| Cycle end   | variant 0 | other variants |
|-------------|-----------|----------------|
| 25/09/2026  | **266571090051** | (Sept — client wala example) |
| 30/10/2026  | **190906610001** | v7 → `155219470795` (bilkul alag prefix) |
| 27/11/2026  | **244891410065** | `key_for(d, 1..99)` koi bhi |

---

## 2. Key kaise generate karein

```bash
python keygen.py               # current + next cycle keys (variant 0)
python keygen.py 30 10 2026    # explicit date (dd mm yyyy) — koi bhi cycle date
python keygen.py 30 10 2026 7  # same date, variant 0-99 (koi bhi chalega)
```

**Distributor ke liye `KeyGen.exe`** (ban chuka hai `keygen_gui.py` se) — usko
sirf program kholna hai, EK key screen pe + Copy button. Date ki sochna nahi:
program khud sabse lambi valid key (aaj ke liye) dikhata hai; cycle change hue toh
khud nayi cycle ki key. Round-robin variant → har click pe alag number.

**Rules (client se confirmed):**
- Validity = calendar-anchored: **cycle-end din tak** (chahe install 2 din pehle ho)
- Install **usi cycle-end ko ya uske baad** → key agli cycle ki
- Har cycle = nayi alag key; **har variant = alag number, same validity**

Owner `keygen.py` kabhi zip/package me mat daalna. **End-user package me
`KeyGen.exe` bhi NAHI** — sirf distributor ko.

---

## 3. Client ke system par kya hota hai (flow)

1. **First run**: `license.json` nahi hai → `license_is_valid() = False`
2. Popup khulta hai: *"License expired — Enter product key"*
3. Key daali → `key_to_date()`:
   - **SHA256 check digits verify** (variant ke saath — galat/guess key yahin reject)
   - XOR se `raw` decode → date nikli (variant koi bhi ho, date same)
4. **Accept rule**: `aaj < date ≤ aaj + 45 din`
5. Accept → `license.json` likh jaata hai (HMAC signature + `last_seen` ke saath)
6. Cancel/blank → error → app **band** (silent exit)
7. **Chalte-chalte har ~60s** license dobara check hota hai — cycle-end khatam hote hi
   dobara popup aata hai
8. **Clock rollback guard**: `last_seen` aaj se aage ka dikhe → reject

---

## 4. license.json kya hota hai

```json
{
  "valid_through": "2026-10-30",   // ISO date, tak chalega (cycle end)
  "issued":       "2026-09-26",    // kab banaya
  "key_month":    "2026-10",       // valid-through ka month (sig field)
  "last_seen":    "2026-09-26",    // aakhri baar kab chala (clock guard)
  "sig":          "hmac-sha256-hex"
}
```

- **Signature**: `HMAC-SHA256(key=SECRET, msg="valid_through|issued|key_month")`
  — `last_seen` sig me shamil NAHI (roz update hota hai)
- Client file edit nahi kar sakta — sig toot jayega
- Location: **exe ke folder me** (`BASE_DIR\license.json`)

---

## 5. Robustness (test — ALL PASS dono tools)

- Round-trip: 400+ random dates + known Fridays (2020-2100) — key → date sab pass
- **100 variants** of same date: sab decode → same date, sab distinct, sab accepted
- **Tampered key** (1/7/100/9999/10000 digit badla) → SHA256 check fail → reject
- **Date visibility**: key me `2026` jaisa kuch nahi dikhta (XOR mask)
- Garbage keys (`abc`, khaali, negative, 11-digit) → clean reject, crash nahi
- Window: past/today/+62 reject, future/+45 accept
- `30/02` jaise impossible dates → `date()` ValueError → reject
- last-Friday math: `last_friday(2028, 2) = 25/02/2028` (leap)
- Clock rollback (`last_seen` future) → reject; tampered `valid_through` → sig fail
- Legacy license (bina last_seen) bhi chalta hai

---

## 6. Apne machine par khud test karna

- Root me `license.json` nahi hai (unactivated ship) — exe chalaoge toh popup aayega
- Apna license wapas chahiye:
  ```bash
  copy understanding\license.json license.json    # root me → 30/10 tak chalega
  ```
- Ya key daal kar activate: `190906610001` (cycle end 30/10/2026, variant 0)
- Naya key: `python keygen.py`

---

## 7. Shipping checklist (kya kisko jaata hai)

**End-user package** (alag-alag PCs pe logon ko):
OptionChain.exe + xlsx + config.json + README/instruction/FEATURES
→ **`keygen.py`, `keygen_gui.py`, `KeyGen.exe`, `KEY_GENERATION.txt`,
`license.json`, `understanding/` — KUCH NAHI jaata**

**Distributor package** (ek main banda):
- Same end-user package (uska apne PC ke liye) +
- **`KeyGen.exe` + `KEY_GENERATION.txt`** (sirf uske liye — aage wo khud
  generate + distribute karega, tumhe har mahine nahi bulana)
- `KeyGen.exe` kabhi end-user zip me mat daalna

| Owner-only (ship MAT karo) |
|----------------------------|
| `keygen.py`, `keygen_gui.py`, `main.py`, `config.py` (source) |
| `understanding/` (yeh folder), active `license.json` |
| nuitka build dirs, logs, `kg.dat` (variant counter) |

- Package docs me koi formula/177/secret nahi hai — **checked CLEAN**
- `KEY_GENERATION.txt` me koi formula/secret nahi — sirf usage + backup warning

---

## 8. Key security notes (honest version)

**Jo protection hai:**
- Secret **plaintext me exe me nahi** (XOR-encoded) — `strings`/basic search se nikal nahi aata
- Key me **date chhupi hai** (XOR) + **SHA256 check digits (variant ke saath)** —
  bina secret ke guess/forge impossible (100 variants × 100 checks = brute-force space)
- license.json **HMAC** — file edit karke date extend nahi kar sakte
- **Clock rollback guard** (`last_seen`)
- Cycle rotation — leaked key bhi cycle end pe dead
- Exes **Nuitka** (C-compiled) — Python bytecode jaisa seedha padhna nahi

**Jo nahi hai:**
- **100% crack-proof nahi** — determined banda exe ko deep RE karke secret + mask nikal
  sakta hai (Nuitka compiled hai, toh effort zyada lagega — bytecode jaisa aasaan nahi)
- Offline validation ka yahi limit hai — 100% bachaav chahiye toh **server-side check**
  ya **machine-ID binding** lagta hai (abhi scope me nahi)

**Practice:**
- `keygen.py`, offset/mask/secret wala code, `understanding/` — kabhi ship mat karo
- Secret ki value kabhi client ko mat batana (yeh doc sirf tumhare paas)
