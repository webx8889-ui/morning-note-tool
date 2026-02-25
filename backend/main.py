# backend/main.py
from fastapi import FastAPI, UploadFile, File
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from io import BytesIO

from openpyxl import load_workbook
from datetime import date, datetime


app = FastAPI(title="Morning Note HTML Generator")

origins = [
    "http://127.0.0.1:5500",
    "http://localhost:5500",
    "http://127.0.0.1:8000",
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ---------- Helper functions ----------

def cell(ws, addr):
    v = ws[addr].value
    return v.strip() if isinstance(v, str) else (v or "")

def fmt_number(value, decimals=2):
    if value is None or value == "":
        return ""
    try:
        num = float(str(value).replace(",", ""))
        fmt = f"{{:,.{decimals}f}}"
        return fmt.format(num)
    except ValueError:
        return str(value)

def fmt_percent(value, decimals=2, show_sign=True):
    if value is None or value == "":
        return ""
    s = str(value).strip()
    try:
        if s.endswith("%"):
            num = float(s.replace("%", "").replace(",", ""))
        else:
            num = float(s.replace(",", ""))
            if abs(num) < 1:
                num = num * 100
        sign = "+" if show_sign and num > 0 else ""
        fmt = f"{{:.{decimals}f}}"
        return f"{sign}{fmt.format(num)}%"
    except ValueError:
        return s

def perc_color(value_str):
    s = str(value_str).strip()
    if not s:
        return "#111827"
    if s.startswith("-"):
        return "#b91c1c"
    return "#15803d"

def fmt_date(v):
    if isinstance(v, datetime):
        return v.strftime("%d %b %Y")
    elif isinstance(v, (int, float)):
        return str(v)
    elif isinstance(v, str):
        for fmt in ("%d-%m-%Y", "%Y-%m-%d", "%d %b %Y"):
            try:
                d = datetime.strptime(v.strip(), fmt)
                return d.strftime("%d %b %Y")
            except ValueError:
                continue
        return v
    return ""

# ---------- HTML template (unchanged, trimmed here for brevity) ----------

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>{title}</title>
</head>
<body style="margin:0; padding:0; background-color:#f3f4f6; font-family: system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; color:#111827;">

<table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color:#f3f4f6;">
  <tr>
    <td align="center" valign="top">

      <table width="100%" cellpadding="0" cellspacing="0" border="0" style="max-width:1024px; padding:24px 16px 40px;">
        <tr>
          <td>

            <!-- HEADER -->
            <table width="100%" cellpadding="0" cellspacing="0" border="0">
              <tr>
                <td align="left" valign="middle" style="font-size:14px; font-weight:600; color:#4b5563;">
                  {date_line}
                </td>
                <td align="right" valign="middle">
                  <a href="{podcast_link}">
                    <img src="https://ekyc.bajajfinservsecurities.in/ekyc/assets/BfslLogo-B8OKXnFb.svg"
                         alt="Bajaj Broking" style="width:150px; height:auto;"/>
                  </a>
                </td>
              </tr>
            </table>

            <!-- MAIN TITLE -->
            <table width="100%" cellpadding="0" cellspacing="0" border="0">
              <tr>
                <td style="font-size:28px; font-weight:700; color:#0f172a; padding:16px 0 16px 0;">
                  {main_heading}
                </td>
              </tr>
            </table>

            <!-- MARKET SNAPSHOT CARD -->
            <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#ffffff; border-radius:12px; box-shadow:0 8px 20px rgba(15,23,42,0.06); margin-bottom:16px;">
              <tr>
                <td style="padding:16px 18px;">

                  <table width="100%" cellpadding="0" cellspacing="0" border="0">
                    <tr>
                      <td style="font-size:18px; font-weight:700; color:#0f172a; padding:0 0 12px 0;">
                        {market_snapshot_heading}
                      </td>
                    </tr>
                  </table>

                  <!-- ROW 1: Gift Nifty | Nifty 50 | Sensex -->
                  <table width="100%" cellpadding="0" cellspacing="0" border="0">
                    <tr>
                      <td valign="top" width="33.33%" style="padding:0 8px 12px 0;">
                        <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#f9fafb; border-radius:10px;">
                          <tr>
                            <td style="padding:10px 12px;">
                              <div style="font-size:11px; text-transform:uppercase; letter-spacing:0.05em; color:#6b7280; margin-bottom:4px;">{gift_label}</div>
                              <div style="font-size:16px; font-weight:600; color:#111827;">{gift_value}</div>
                              <div style="font-size:11px; margin-top:2px; color:{gift_color}; font-weight:500;">{gift_change}</div>
                            </td>
                          </tr>
                        </table>
                      </td>

                      <td valign="top" width="33.33%" style="padding:0 8px 12px 0;">
                        <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#f9fafb; border-radius:10px;">
                          <tr>
                            <td style="padding:10px 12px;">
                              <div style="font-size:11px; text-transform:uppercase; letter-spacing:0.05em; color:#6b7280; margin-bottom:4px;">{nifty_label}</div>
                              <div style="font-size:16px; font-weight:600; color:#111827;">{nifty_value}</div>
                              <div style="font-size:11px; margin-top:2px; color:{nifty_color}; font-weight:500;">{nifty_change}</div>
                            </td>
                          </tr>
                        </table>
                      </td>

                      <td valign="top" width="33.33%" style="padding:0 0 12px 0;">
                        <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#f9fafb; border-radius:10px;">
                          <tr>
                            <td style="padding:10px 12px;">
                              <div style="font-size:11px; text-transform:uppercase; letter-spacing:0.05em; color:#6b7280; margin-bottom:4px;">{sensex_label}</div>
                              <div style="font-size:16px; font-weight:600; color:#111827;">{sensex_value}</div>
                              <div style="font-size:11px; margin-top:2px; color:{sensex_color}; font-weight:500;">{sensex_change}</div>
                            </td>
                          </tr>
                        </table>
                      </td>
                    </tr>

                    <!-- ROW 2: Bank Nifty | India VIX | USDINR -->
                    <tr>
                      <td valign="top" width="33.33%" style="padding:0 8px 12px 0;">
                        <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#f9fafb; border-radius:10px;">
                          <tr>
                            <td style="padding:10px 12px;">
                              <div style="font-size:11px; text-transform:uppercase; letter-spacing:0.05em; color:#6b7280; margin-bottom:4px;">{bank_label}</div>
                              <div style="font-size:16px; font-weight:600; color:#111827;">{bank_value}</div>
                              <div style="font-size:11px; margin-top:2px; color:{bank_color}; font-weight:500;">{bank_change}</div>
                            </td>
                          </tr>
                        </table>
                      </td>

                      <td valign="top" width="33.33%" style="padding:0 8px 12px 0;">
                        <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#f9fafb; border-radius:10px;">
                          <tr>
                            <td style="padding:10px 12px;">
                              <div style="font-size:11px; text-transform:uppercase; letter-spacing:0.05em; color:#6b7280; margin-bottom:4px;">{vix_label}</div>
                              <div style="font-size:16px; font-weight:600; color:#111827;">{vix_value}</div>
                              <div style="font-size:11px; margin-top:2px; color:{vix_color}; font-weight:500;">{vix_change}</div>
                            </td>
                          </tr>
                        </table>
                      </td>

                      <td valign="top" width="33.33%" style="padding:0 0 12px 0;">
                        <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#f9fafb; border-radius:10px;">
                          <tr>
                            <td style="padding:10px 12px;">
                              <div style="font-size:11px; text-transform:uppercase; letter-spacing:0.05em; color:#6b7280; margin-bottom:4px;">{usdinr_label}</div>
                              <div style="font-size:16px; font-weight:600; color:#111827;">{usdinr_value}</div>
                              <div style="font-size:11px; margin-top:2px; color:{usdinr_color}; font-weight:500;">{usdinr_change}</div>
                            </td>
                          </tr>
                        </table>
                      </td>
                    </tr>
                  </table>

                  <!-- PODCAST SECTION -->
                  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:12px;">
                    <tr>
                      <td>
                        <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#0f172a; border-radius:12px; color:#e5e7eb;">
                          <tr>
                            <td style="padding:16px 18px;">
                              <table cellpadding="0" cellspacing="0" border="0" style="margin-bottom:8px;">
                                <tr>
                                  <td style="padding:4px 10px; border-radius:999px; background:#1f2937; color:#e5e7eb; font-size:11px; font-weight:500;">
                                    {podcast_tagline}
                                  </td>
                                </tr>
                              </table>
                              <p style="font-size:12px; margin:8px 0 0; line-height:1.5; color:#e5e7eb;">
                                {podcast_para1}
                              </p>
                              <table cellpadding="0" cellspacing="0" border="0" style="margin-top:8px;">
                                <tr>
                                  <td align="center" style="background:#005DAC; border-radius:999px;">
                                    <a href="{podcast_link}" style="display:block; padding:8px 16px; font-size:12px; font-weight:600; color:#facc15; text-decoration:none;">
                                      🎧 LISTEN NOW 🎧
                                    </a>
                                  </td>
                                </tr>
                              </table>
                              <div style="font-size:11px; margin-top:6px; color:#e5e7eb;">
                                {podcast_footer}
                              </div>
                            </td>
                          </tr>
                        </table>
                      </td>
                    </tr>
                  </table>

                </td>
              </tr>
            </table>

            <!-- INDIA MARKET RECAP -->
            <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:24px;">
              <tr>
                <td style="font-size:18px; font-weight:700; color:#0f172a; padding-bottom:12px;">
                  {recap_heading}
                </td>
              </tr>
            </table>

            <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#ffffff; border-radius:12px; box-shadow:0 8px 20px rgba(15,23,42,0.06); margin-bottom:16px;">
              <tr>
                <td style="padding:16px 18px;">
                  <p style="font-size:14px; line-height:1.6; margin:0 0 12px 0;">
                    {recap_para1}
                  </p>
                  <p style="font-size:14px; line-height:1.6; margin:0 0 12px 0;">
                    {recap_para2}
                  </p>
                  <p style="font-size:14px; line-height:1.6; margin:0 0 12px 0;">
                    {recap_para3}
                  </p>
                  <p style="font-size:11px; color:#6b7280; margin:0;">
                    {recap_source}
                  </p>
                </td>
              </tr>
            </table>

            <!-- TRADING PLAYBOOK -->
            <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:24px;">
              <tr>
                <td style="font-size:18px; font-weight:700; color:#0f172a; padding-bottom:12px;">
                  {playbook_heading}
                </td>
              </tr>
            </table>

            <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#ffffff; border-radius:12px; box-shadow:0 8px 20px rgba(15,23,42,0.06); margin-bottom:16px;">
              <tr>
                <td style="padding:16px 18px;">

                  <table width="100%" cellpadding="0" cellspacing="0" border="0">
                    <tr>
                      <td style="font-size:14px; font-weight:600; color:#374151; padding-bottom:4px;">{flows_heading}</td>
                    </tr>
                    # <tr>
                    #   <td style="font-size:11px; color:#6b7280; padding-bottom:10px;">{flows_subtext}</td>
                    # </tr>
                  </table>

                  <!-- FII/DII TABLE -->
                  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="font-size:12px; border-collapse:collapse;">
                    <thead>
                      <tr>
                        <th align="left" style="padding:8px 6px; background:#eff6ff; border-bottom:1px solid #e5e7eb; font-size:11px; text-transform:uppercase; letter-spacing:0.04em; color:#111827;">Participant</th>
                        <th align="left" style="padding:8px 6px; background:#eff6ff; border-bottom:1px solid #e5e7eb; font-size:11px; text-transform:uppercase; letter-spacing:0.04em; color:#111827;">Previous Day (&#8377; Cr)</th>
                        <th align="left" style="padding:8px 6px; background:#eff6ff; border-bottom:1px solid #e5e7eb; font-size:11px; text-transform:uppercase; letter-spacing:0.04em; color:#111827;">MTD (&#8377; Cr)</th>
                        <th align="left" style="padding:8px 6px; background:#eff6ff; border-bottom:1px solid #e5e7eb; font-size:11px; text-transform:uppercase; letter-spacing:0.04em; color:#111827;">YTD (&#8377; Cr)</th>
                      </tr>
                    </thead>
                    <tbody>
                      <tr>
                        <td style="padding:8px 6px; border-bottom:1px solid #e5e7eb;">FII</td>
                        <td style="padding:8px 6px; border-bottom:1px solid #e5e7eb; color:{fii_prev_color}; font-weight:500;">{fii_prev}</td>
                        <td style="padding:8px 6px; border-bottom:1px solid #e5e7eb; color:{fii_mtd_color}; font-weight:500;">{fii_mtd}</td>
                        <td style="padding:8px 6px; border-bottom:1px solid #e5e7eb; color:{fii_ytd_color}; font-weight:500;">{fii_ytd}</td>
                      </tr>
                      <tr>
                        <td style="padding:8px 6px; border-bottom:1px solid #e5e7eb;">DII</td>
                        <td style="padding:8px 6px; border-bottom:1px solid #e5e7eb; color:{dii_prev_color}; font-weight:500;">{dii_prev}</td>
                        <td style="padding:8px 6px; border-bottom:1px solid #e5e7eb; color:{dii_mtd_color}; font-weight:500;">{dii_mtd}</td>
                        <td style="padding:8px 6px; border-bottom:1px solid #e5e7eb; color:{dii_ytd_color}; font-weight:500;">{dii_ytd}</td>
                      </tr>
                    </tbody>
                  </table>
                  # <p style="font-size:11px; color:#6b7280; margin:8px 0 0 0;">{flows_source}</p>

                  <!-- RANGE LEVELS -->
                  <p style="font-size:11px; color:#6b7280; margin:16px 0 4px 0;"><strong>{range_heading}</strong></p>
                  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="font-size:12px; border-collapse:collapse;">
                    <thead>
                      <tr>
                        <th align="left" style="padding:8px 6px; background:#eff6ff; border-bottom:1px solid #e5e7eb; font-size:11px; text-transform:uppercase; letter-spacing:0.04em; color:#111827;">Index</th>
                        <th align="left" style="padding:8px 6px; background:#eff6ff; border-bottom:1px solid #e5e7eb; font-size:11px; text-transform:uppercase; letter-spacing:0.04em; color:#111827;">Support 1</th>
                        <th align="left" style="padding:8px 6px; background:#eff6ff; border-bottom:1px solid #e5e7eb; font-size:11px; text-transform:uppercase; letter-spacing:0.04em; color:#111827;">Support 2</th>
                        <th align="left" style="padding:8px 6px; background:#eff6ff; border-bottom:1px solid #e5e7eb; font-size:11px; text-transform:uppercase; letter-spacing:0.04em; color:#111827;">Resistance 1</th>
                        <th align="left" style="padding:8px 6px; background:#eff6ff; border-bottom:1px solid #e5e7eb; font-size:11px; text-transform:uppercase; letter-spacing:0.04em; color:#111827;">Resistance 2</th>
                      </tr>
                    </thead>
                    <tbody>
                      <tr>
                        <td style="padding:8px 6px; border-bottom:1px solid #e5e7eb;">{range_row1_index}</td>
                        <td style="padding:8px 6px; border-bottom:1px solid #e5e7eb;">{range_row1_s1}</td>
                        <td style="padding:8px 6px; border-bottom:1px solid #e5e7eb;">{range_row1_s2}</td>
                        <td style="padding:8px 6px; border-bottom:1px solid #e5e7eb;">{range_row1_r1}</td>
                        <td style="padding:8px 6px; border-bottom:1px solid #e5e7eb;">{range_row1_r2}</td>
                      </tr>
                      <tr>
                        <td style="padding:8px 6px; border-bottom:1px solid #e5e7eb;">{range_row2_index}</td>
                        <td style="padding:8px 6px; border-bottom:1px solid #e5e7eb;">{range_row2_s1}</td>
                        <td style="padding:8px 6px; border-bottom:1px solid #e5e7eb;">{range_row2_s2}</td>
                        <td style="padding:8px 6px; border-bottom:1px solid #e5e7eb;">{range_row2_r1}</td>
                        <td style="padding:8px 6px; border-bottom:1px solid #e5e7eb;">{range_row2_r2}</td>
                      </tr>
                    </tbody>
                  </table>
                  # <p style="font-size:11px; color:#6b7280; margin:8px 0 0 0;">{range_comment}</p>

                </td>
              </tr>
            </table>

            <!-- GLOBAL MARKET PULSE -->
            <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:24px;">
              <tr>
                <td style="font-size:18px; font-weight:700; color:#0f172a; padding-bottom:12px;">
                  {global_heading}
                </td>
              </tr>
            </table>

            <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#ffffff; border-radius:12px; box-shadow:0 8px 20px rgba(15,23,42,0.06); margin-bottom:16px;">
              <tr>
                <td style="padding:16px 18px;">
                  <p style="font-size:14px; line-height:1.6; margin:0 0 10px 0;">
                    {global_para1}
                  </p>
                  # <p style="font-size:14px; line-height:1.6; margin:0 0 10px 0;">
                  #   {global_para2}
                  # </p>
                  # <p style="font-size:11px; color:#6b7280; margin:0;">
                  #   {global_source}
                  # </p>
                </td>
              </tr>
            </table>

            <!-- STOCKS IN FOCUS -->
            <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:24px;">
              <tr>
                <td style="font-size:18px; font-weight:700; color:#0f172a; padding-bottom:12px;">
                  {stocks_heading}
                </td>
              </tr>
            </table>

            <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#ffffff; border-radius:12px; box-shadow:0 8px 20px rgba(15,23,42,0.06); margin-bottom:16px;">
              <tr>
                <td style="padding:16px 18px;">
                  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="font-size:12px; border-collapse:collapse;">
                    <thead>
                      <tr>
                        <th align="left" style="padding:8px 6px; background:#eff6ff; border-bottom:1px solid #e5e7eb; font-size:11px; text-transform:uppercase; letter-spacing:0.04em; color:#111827;">Symbol</th>
                        <th align="left" style="padding:8px 6px; background:#eff6ff; border-bottom:1px solid #e5e7eb; font-size:11px; text-transform:uppercase; letter-spacing:0.04em; color:#111827;">Price %</th>
                        <th align="left" style="padding:8px 6px; background:#eff6ff; border-bottom:1px solid #e5e7eb; font-size:11px; text-transform:uppercase; letter-spacing:0.04em; color:#111827;">OI %</th>
                        <th align="left" style="padding:8px 6px; background:#eff6ff; border-bottom:1px solid #e5e7eb; font-size:11px; text-transform:uppercase; letter-spacing:0.04em; color:#111827;">Interpretation</th>
                      </tr>
                    </thead>
                    # <tbody>
                    #   {stocks_rows_html}
                    # </tbody>
                  </table>
                  # <p style="font-size:11px; color:#6b7280; margin:8px 0 0 0;">
                  #   {stocks_source}
                  # </p>
                </td>
              </tr>
            </table>

            <!-- STOCKS IN NEWS / CORPORATE HIGHLIGHTS -->
            <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:24px;">
              <tr>
                <td style="font-size:18px; font-weight:700; color:#0f172a; padding-bottom:12px;">
                  {corp_heading}
                </td>
              </tr>
            </table>

            <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#ffffff; border-radius:12px; box-shadow:0 8px 20px rgba(15,23,42,0.06);">
              <tr>
                <td style="padding:16px 18px;">
                  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="font-size:12px; border-collapse:collapse;">
                    <thead>
                      <tr>
                        <th align="left" style="padding:8px 6px; background:#eff6ff; border-bottom:1px solid #e5e7eb; font-size:11px; text-transform:uppercase; letter-spacing:0.04em; color:#111827;">Company / Theme</th>
                        <th align="left" style="padding:8px 6px; background:#eff6ff; border-bottom:1px solid #e5e7eb; font-size:11px; text-transform:uppercase; letter-spacing:0.04em; color:#111827;">Update</th>
                      </tr>
                    </thead>
                    <tbody>
                      {corp_rows_html}
                    </tbody>
                  </table>
                  <p style="font-size:11px; color:#6b7280; margin:8px 0 0 0;">
                    {corp_source}
                  </p>
                </td>
              </tr>
            </table>

            <!-- DISCLAIMER -->
            <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:24px;">
              <tr>
                <td style="padding:0px; font-size:11px; color:#374151; line-height:1.6;">
                  <p style="margin:0 0 8px 0;"><strong>Note:</strong> {note_text}</p>
                  <p style="margin:0 0 8px 0;"><strong>Disclaimer:</strong> Please do not reply to this email, as responses to this address will not be received or monitored. For any queries, feel free to contact us at <a href="mailto:connect@bajajbroking.in" style="color:#1414DF; font-weight:450;">connect@bajajbroking.in</a>.</p>
                  <p style="margin:0 0 8px 0; line-height:16px;">Investments in the securities market are subject to market risk. Read all related documents carefully before investing. REG OFFICE: Bajaj Auto Limited Complex, Mumbai Pune Road, Akurdi, Pune 411035. Corp. Office: Bajaj Financial Securities Limited. 1st Floor, Mantri IT Park, Tower B, Unit No. 9 &amp; 10, Viman Nagar, Pune, Maharashtra 411014. SEBI Registration No.: INZ000218931 | BSE Cash/F&amp;O/CDS Member ID: 6706 | NSE Cash/F&amp;O/CDS Member ID: 90177 | SEBI DP Registration No.: IN-DP-418-2019 | CDSL DP No.: 12088600 | NSDL DP No.: IN304030 | AMFI Registration No.: ARN-163403 | SEBI Registration No. (Research Analyst/Entity): INH000010043 | Website: <a href="https://www.bajajbroking.in" style="color:#1414DF; font-weight:450;">https://www.bajajbroking.in</a></p>
                  <p style="margin:0 0 8px 0;">Compliance Officer: Mr. Boudhayan Ghosh (For Broking/DP/Research) | Email: <a href="mailto:compliance_sec@bajajbroking.in" style="color:#1414DF; font-weight:450;">compliance_sec@bajajbroking.in</a> | Contact No.: 020-4857 4486.</p>
                  <p style="margin:0 0 8px 0;">For any queries/information, you can call the toll-free number 1800-833-8888 or write to us on <a href="mailto:connect@bajajbroking.in" style="color:#1414DF; font-weight:450;">connect@bajajbroking.in</a>. For any Investor Grievances, write to compliance on <a href="mailto:compliance_sec@bajajbroking.in" style="color:#1414DF; font-weight:450;">compliance_sec@bajajbroking.in</a> or <a href="mailto:compliance_dp@bajajbroking.in" style="color:#1414DF; font-weight:450;">compliance_dp@bajajbroking.in</a></p>
                  <p style="margin:0;">Kindly refer to <a href="https://www.bajajbroking.in/disclaimer" style="color:#1414DF; font-weight:450;">https://www.bajajbroking.in/disclaimer</a> for detailed disclaimer and risk factors.</p>
                </td>
              </tr>
            </table>

          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>

</body>
</html>

"""

# ---------- Core generator using Excel ----------

def generate_html_from_excel(excel_bytes: bytes) -> str:
    wb = load_workbook(BytesIO(excel_bytes), data_only=True)
    ws = wb["Sheet1"]

    # ----- MARKET SNAPSHOT -----

    raw_gift_value   = cell("D9")
    raw_gift_change  = cell("D10")
    raw_nifty_value  = cell("G9")
    raw_nifty_change = cell("G10")
    raw_sensex_value = cell("J9")
    raw_sensex_change= cell("J10")

    raw_bank_value   = cell("D13")
    raw_bank_change  = cell("D14")
    raw_vix_value    = cell("G13")
    raw_vix_change   = cell("G14")
    raw_usdinr_value = cell("J13")
    raw_usdinr_change= cell("J14")


    gift_value   = fmt_number(raw_gift_value, 1)
    gift_change  = fmt_percent(raw_gift_change, 2)
    nifty_value  = fmt_number(raw_nifty_value, 2)
    nifty_change = fmt_percent(raw_nifty_change, 2)
    sensex_value = fmt_number(raw_sensex_value, 2)
    sensex_change= fmt_percent(raw_sensex_change, 2)

    bank_value   = fmt_number(raw_bank_value, 2)
    bank_change  = fmt_percent(raw_bank_change, 2)
    vix_value    = fmt_number(raw_vix_value, 2)
    vix_change   = fmt_percent(raw_vix_change, 2)
    usdinr_value = fmt_number(raw_usdinr_value, 4)
    usdinr_change= fmt_percent(raw_usdinr_change, 2)

    # ----- FII/DII + RANGE -----
    playbook_heading = cell(ws, "A19")
    flows_heading    = cell(ws, "A20")
    flows_subtext    = cell(ws, "A22")

    raw_fii_prev = cell(ws, "B24")
    raw_fii_mtd  = cell(ws, "C24")
    raw_fii_ytd  = cell(ws, "D24")

    raw_dii_prev = cell(ws, "B25")
    raw_dii_mtd  = cell(ws, "C25")
    raw_dii_ytd  = cell(ws, "D25")

    fii_prev = fmt_number(raw_fii_prev, 0)
    fii_mtd  = fmt_number(raw_fii_mtd, 2)
    fii_ytd  = fmt_number(raw_fii_ytd, 2)

    dii_prev = fmt_number(raw_dii_prev, 0)
    dii_mtd  = fmt_number(raw_dii_mtd, 2)
    dii_ytd  = fmt_number(raw_dii_ytd, 2)

    flows_source = cell(ws, "A26")

    # ----- F&O points -----
    nifty_points_html = ""
    for row in range(45, 48):
        txt = cell(ws, f"A{row}")
        if txt:
            nifty_points_html += f"<li>{txt}</li>"

    bank_nifty_points_html = ""
    for row in range(51, 54):
        txt = cell(ws, f"A{row}")
        if txt:
            bank_nifty_points_html += f"<li>{txt}</li>"        

    # ----- STOCKS IN FOCUS -----
    stocks_rows_html = ""
    for r in range(73, 89):
        bucket = cell(ws, f"A{r}")
        stock  = cell(ws, f"B{r}")
        raw_price = cell(ws, f"C{r}")
        raw_oi    = cell(ws, f"D{r}")
        interp    = cell(ws, f"E{r}")

        if not (bucket or stock or raw_price or raw_oi or interp):
            break

        price = fmt_percent(raw_price, 2, show_sign=False)
        oi    = fmt_percent(raw_oi, 2, show_sign=False)

        stocks_rows_html += f"""
      <tr>
        <td style="padding:8px 6px; border-bottom:1px solid #e5e7eb;">{bucket}</td>
        <td style="padding:8px 6px; border-bottom:1px solid #e5e7eb;">{stock}</td>
        <td style="padding:8px 6px; border-bottom:1px solid #e5e7eb;">{price}</td>
        <td style="padding:8px 6px; border-bottom:1px solid #e5e7eb;">{oi}</td>
        <td style="padding:8px 6px; border-bottom:1px solid #e5e7eb;">{interp}</td>
      </tr>"""

    # ----- KEY EVENTS -----
    events_rows_html = ""
    for r in range(94, 111):
        raw_date  = ws[f"A{r}"].value
        ev_date   = fmt_date(raw_date)
        ev_country = cell(ws, f"B{r}")
        ev_event   = cell(ws, f"C{r}")

        if not (ev_date or ev_country or ev_event):
            break

        events_rows_html += f"""
      <tr>
        <td style="padding:8px 6px; border-bottom:1px solid #e5e7eb;">{ev_date}</td>
        <td style="padding:8px 6px; border-bottom:1px solid #e5e7eb;">{ev_country}</td>
        <td style="padding:8px 6px; border-bottom:1px solid #e5e7eb;">{ev_event}</td>
      </tr>"""

    # ----- CORPORATE HIGHLIGHTS -----
    corp_rows_html = ""
    for r in range(105, 115):
        company = cell(ws, f"A{r}")
        update  = cell(ws, f"B{r}")
        if not (company or update):
            break
        corp_rows_html += f"""
      <tr>
        <td style="padding:8px 6px; border-bottom:1px solid #e5e7eb;"><strong>{company}</strong></td>
        <td style="padding:8px 6px; border-bottom:1px solid #e5e7eb;">{update}</td>
      </tr>"""

    # ----- CONTEXT (same as earlier script) -----
    context = {
        # top level
    # "title":            cell("A3"),
    "date_line":        cell("D2"),
    "main_heading":     cell("D3"),

    # market snapshot – labels
    "market_snapshot_heading": cell("D6"),
    "gift_label":       cell("G8"),
    "nifty_label":      cell("B5"),
    "sensex_label":     cell("J8"),
    "bank_label":       cell("D12"),
    "vix_label":        cell("G12"),
    "usdinr_label":     cell("J12"),

    # market snapshot 
    "gift_value":   gift_value,
    "gift_change":  gift_change,
    "gift_color":   perc_color(gift_change),

    "nifty_value":  nifty_value,
    "nifty_change": nifty_change,
    "nifty_color":  perc_color(nifty_change),

    "sensex_value": sensex_value,
    "sensex_change":sensex_change,
    "sensex_color": perc_color(sensex_change),

    "bank_value":   bank_value,
    "bank_change":  bank_change,
    "bank_color":   perc_color(bank_change),

    "vix_value":    vix_value,
    "vix_change":   vix_change,
    "vix_color":    perc_color(vix_change),

    "usdinr_value": usdinr_value,
    "usdinr_change":usdinr_change,
    "usdinr_color": perc_color(usdinr_change),

    # podcast section
    "podcast_tagline":       cell("D18"),
    "podcast_para1":         cell("D19"),
    # "podcast_para2":         cell("D8"),
    "podcast_link":          cell("D22"),
    # "podcast_button_label":  cell("B17"),
    

    # recap section
    "recap_heading":   cell("D24"),
    "recap_para1":     cell("D25"),
    # "recap_para2":     cell("A15"),
    # "recap_para3":     cell("A16"),
    "recap_source":    cell("A17"),

    # trading playbook / flows
    "playbook_heading": playbook_heading,
    "flows_heading":    flows_heading,
    "flows_subtext":    flows_subtext,

   "fii_prev":        fii_prev,
    "fii_prev_color":  perc_color(fii_prev),
    "fii_mtd":         fii_mtd,
    "fii_mtd_color":   perc_color(fii_mtd),
    "fii_ytd":         fii_ytd,
    "fii_ytd_color":   perc_color(fii_ytd),

    "dii_prev":        dii_prev,
    "dii_prev_color":  perc_color(dii_prev),
    "dii_mtd":         dii_mtd,
    "dii_mtd_color":   perc_color(dii_mtd),
    "dii_ytd":         dii_ytd,
    "dii_ytd_color":   perc_color(dii_ytd),
    "flows_source":    flows_source,

    # range table
    "range_heading":   cell("D36"),
    "range_row1_index": cell("D38"),
    "range_row1_s1":    fmt_number(cell("E38"), 0),
    "range_row1_s2":    fmt_number(cell("G38"), 0),
    "range_row1_r1":    fmt_number(cell("I38"), 0),
    "range_row1_r2":    fmt_number(cell("K38"), 0),

    "range_row2_index": cell("A31"),
    "range_row2_s1":    fmt_number(cell("E39"), 0),
    "range_row2_s2":    fmt_number(cell("G39"), 0),
    "range_row2_r1":    fmt_number(cell("I39"), 0),
    "range_row2_r2":    fmt_number(cell("K39"), 0),

    # "range_comment":   cell("A32"),

    # outlook
    "outlook_heading": cell("A35"),
    "outlook_para1":   cell("A37"),
    "outlook_para2":   cell("A39"),
    "outlook_para3":   cell("A41"),

    # F&O
    "fo_heading":          cell("A43"),
    "nifty_option_heading": cell("A44"),
    

    "nifty_option_points" : nifty_points_html,
    "bank_option_heading":  cell("A50"),
    "bank_option_points":   nifty_points_html,
    "fo_source":            cell("A56"),
    "fo_ban_stocks":        cell("A57"),

    # global market
    "global_heading": cell("D41"),
    "global_para1":   cell("D42"),
    # "global_source":  cell("A63"),

    # sector
    "sector_heading": cell("A65"),
    "sector_leaders": cell("A67"),
    "sector_laggards":cell("E67"),
    "sector_source":  cell("A69"),

    # stocks in focus 
    "stocks_heading":   cell("D47"),
    "stocks_rows_html": stocks_rows_html,
    # "stocks_source":    cell("A90"),

    # key events
    "events_heading":   cell("A92"),
    "events_rows_html": events_rows_html,
    "events_source":    cell("A101"),

    # corporate highlights
    "corp_heading":   cell("D54"),
     "corp_rows_html": corp_rows_html,
    # "corp_source":    cell("A116"),

    # disclaimer block
    "note_text":           cell("A118"),
    "disclaimer_main":     cell("B76"),
    "disclaimer_reg":      cell("B77"),
    "compliance_contact":  cell("B78"),
    "support_contact":     cell("B79"),
    "disclaimer_link_text":cell("B80"),
    }

    wb.close()
    return HTML_TEMPLATE.format(**context)

# ---------- FastAPI endpoint ----------

@app.post("/generate-html")
async def generate_html(file: UploadFile = File(...)):
    """
    Upload Excel, get Morning Note HTML.
    """
    excel_bytes = await file.read()
    html_str = generate_html_from_excel(excel_bytes)

    buf = BytesIO(html_str.encode("utf-8"))
    filename = f"MorningNote_{date.today():%d-%b-%Y}.html"

    return StreamingResponse(
        buf,
        media_type="text/html",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )
