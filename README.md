# HealthCare

This is a small care-tracking project I built for my own family.

After a family member got sick, we needed a simple way for caregivers to record daily care items: blood sugar, blood pressure, medication, insulin, breathing training, rehab, urine/stool, and weight. I am open-sourcing it in case it helps another family facing a similar situation.

It is not a medical device. Please follow your doctor or nurse's instructions for all medication and insulin decisions.

## What It Uses

- Google Apps Script
- Google Sheets as the database
- One web UI file: `Index.html`
- One backend file: `程式碼.gs`

## Admin

Default admin login:

- Account: `admin`
- Password: `1234`

The admin backend has three main pages:

- `藥物`
- `設定`
- `人員`

## Insulin

Open admin backend, then go to `藥物`.

### 長效型胰島素

Add one row per shot:

- Time
- Dose

Default:

- 07:00, 12U

### 短效型胰島素

Add one row per blood sugar range:

- Blood sugar range
- Dose or action

Default:

- 0-200: no insulin
- 201-250: 4U
- 251-300: 8U
- 301-350: 8U
- 351-400: 10U
- 401 and above: send to hospital

After a caregiver records blood sugar, the same blood sugar row shows the suggested short-acting insulin dose. The caregiver must press confirm after giving the shot, and the app records that short-acting insulin dose for that time.

## Check

Run:

```bash
node insulin-scale.test.js
```
