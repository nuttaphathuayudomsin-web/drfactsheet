# DR Factsheet Generator

Streamlit app to generate DR factsheet PPTX files from a form — no Google Workspace required.

## Setup (Local)

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Run the app
streamlit run app.py
```

Then open http://localhost:8501 in your browser.

## Deploy to Streamlit Community Cloud (Free)

1. Push this folder to a **GitHub repo** (can be private)
2. Go to https://share.streamlit.io
3. Click **New app** → select your repo → set main file to `app.py`
4. Click **Deploy** — done, you get a shareable URL

Share the URL with your team (up to 5 people for free tier).

## Password

The default password is `ktbdr2025`. To change it, edit line 3 of `app.py`:

```python
APP_PASSWORD = "your_new_password_here"
```

For Streamlit Cloud, store it in **Secrets** instead of hardcoding:
1. Streamlit Cloud dashboard → your app → **Settings → Secrets**
2. Add: `APP_PASSWORD = "your_password"`
3. In `app.py` change to: `APP_PASSWORD = st.secrets["APP_PASSWORD"]`

## How to use

1. Upload your `template.pptx` (the one with `{{placeholders}}`)
2. Fill in the 8 form fields
3. Click **Generate Factsheet**
4. Review the preview on the right
5. Click **Download** to get the filled `.pptx`

## Form fields

| Field | Description |
|-------|-------------|
| Ticker | e.g. SUNNY80 |
| รหัสหลักทรัพย์ | e.g. 2383 HK |
| ชื่อบริษัท (EN) | e.g. Sunny Optical Technology (Group) Co., Ltd. |
| ตลาดจดทะเบียน (ชื่อเต็ม) | e.g. Hong Kong Stock Exchange |
| ชื่อย่อตลาด | e.g. HKEX, NDX, TSE |
| จำนวนหน่วย | default 10,000,000,000 |
| อัตราอ้างอิง | e.g. 100 → becomes "1 : 100" |
| วันซื้อขายวันแรก | e.g. 11 มี.ค. 69 |
| ลิงก์ Filing | SEC URL for QR code (optional) |

## Fixed values (hardcoded, never change)

- Exchange: ตลาดหลักทรัพย์แห่งประเทศไทย (SET)
- Depositary: ธนาคารกรุงไทย จำกัด (มหาชน)
- Offering type: Direct Listing
- Price info: เป็นไปตามกลไกราคา...
- KTB Contact: 02-208-3748, 02-208-4669

## Template placeholders

Make sure your `.pptx` template contains these exact placeholder strings:

```
{{ticker}}              {{full_name_thai}}      {{full_name_eng}}
{{exchange}}            {{underlying_stock}}    {{underlying_exchange}}
{{depositary}}          {{offering_type}}       {{total_units}}
{{ratio}}               {{first_trading_date}}  {{price_info}}
{{foreign_exchange}}    {{filing_url}}          {{ktb_contact}}
{{qr_code}}  ← text box sized to your desired QR dimensions
```
