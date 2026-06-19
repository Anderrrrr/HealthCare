# HealthCare

Google Apps Script care log app.

## Files

- `Index.html`: web UI.
- `程式碼.gs`: Apps Script backend and Google Sheets setup.

## Insulin Settings

Admin login: `admin` / `1234`, then open `藥物`.

### 長效型胰島素

Use one line per shot:

```text
07:00=12
```

Default: 07:00, 12U.

### 短效型胰島素

Short-acting insulin follows blood sugar entries. Use one line per range:

```text
0-200=0
201-250=4
251-300=8
301-350=8
351-400=10
401-=送醫院
```

Default rule:

- `<200`: no insulin. The app stores this as `0-200=0`.
- `201-250`: 4U.
- `251-300`: 8U.
- `301-350`: 8U.
- `351-400`: 10U.
- `>400`: send to hospital.

## Check

Run:

```bash
node insulin-scale.test.js
```
