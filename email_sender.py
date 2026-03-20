# ============================================================
#  email_sender.py  —  用 Outlook 寄送風控整合日報
#  需要本機已安裝 Outlook 並登入
# ============================================================

import sys
from pathlib import Path
from datetime import date


def send_report(report_date: date, html_path: Path, data: dict):
    """
    用 Outlook COM 介面寄出報告
    - 信件本文：今日重點說明
    - 附件：當天 HTML 報告
    """
    try:
        import win32com.client
    except ImportError:
        print("  [ERROR] 找不到 pywin32，請先執行：py -m pip install pywin32")
        return False

    import config

    # ── 組合信件本文（今日重點說明）──────────────────────────
    alert_items = data.get("alert_items", [])
    date_str = report_date.strftime("%Y/%m/%d")

    def _ai_color(t):
        return {"red": "#c62828", "yellow": "#b45309", "blue": "#1565c0"}.get(t, "#4a6080")

    def _ai_icon(t):
        return {"red": "🔴", "yellow": "🟡", "blue": "🔵"}.get(t, "🔵")

    items_html = ""
    for ai in alert_items:
        color = _ai_color(ai["type"])
        icon  = _ai_icon(ai["type"])
        items_html += f"""
        <tr>
          <td style="padding:5px 12px;font-size:13px;color:{color};">
            {icon} {ai['text']}
          </td>
        </tr>"""

    if not items_html:
        items_html = """
        <tr>
          <td style="padding:5px 12px;font-size:13px;color:#1a9e6a;">
            ✅ 今日各項指標正常
          </td>
        </tr>"""

    body_html = f"""
<html>
<body style="font-family:system-ui,sans-serif;font-size:13px;color:#1a2535;margin:0;padding:0;">
<div style="max-width:600px;margin:20px auto;padding:0 16px;">

  <!-- Header -->
  <div style="border-bottom:2px solid #1565c0;padding-bottom:10px;margin-bottom:16px;">
    <div style="font-size:17px;font-weight:700;">風險管理整合日報</div>
    <div style="font-size:11px;color:#8a9bb5;margin-top:3px;">INTEGRATED RISK MANAGEMENT DAILY REPORT</div>
  </div>

  <!-- 日期 -->
  <div style="font-size:12px;color:#4a6080;margin-bottom:12px;">
    資料日期：<strong>{date_str}</strong>　｜　風險管理部
  </div>

  <!-- 今日重點說明 -->
  <div style="background:#fef2f2;border:1px solid #fccaca;border-left:4px solid #c62828;
              border-radius:6px;padding:12px 16px;margin-bottom:20px;">
    <div style="font-size:11px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;
                color:#c62828;font-family:monospace;margin-bottom:8px;">
      ⚡ 今日重點說明
    </div>
    <table style="width:100%;border-collapse:collapse;">
      {items_html}
    </table>
  </div>

  <!-- 附件說明 -->
  <div style="background:#f8f9fb;border:1px solid #e3e8ef;border-radius:6px;
              padding:12px 16px;margin-bottom:20px;">
    <div style="font-size:12px;color:#4a6080;">
      📎 完整報告請見附件：<strong>{html_path.name}</strong><br>
      <span style="font-size:11px;color:#8a9bb5;">
        下載後以瀏覽器開啟即可檢視
      </span>
    </div>
  </div>

  <!-- Footer -->
  <div style="font-size:10px;color:#8a9bb5;border-top:1px solid #e3e8ef;padding-top:10px;">
    本信件由風控日報系統自動產生　｜　如有疑問請聯繫風險管理部
  </div>

</div>
</body>
</html>"""

    # ── 收件人 ────────────────────────────────────────────────
    to_list = getattr(config, "EMAIL_TO", [])
    cc_list = getattr(config, "EMAIL_CC", [])
    subject_tpl = getattr(config, "EMAIL_SUBJECT", "【風控整合日報】{date}")
    subject = subject_tpl.format(date=date_str)

    if not to_list:
        print("  [WARN] config.EMAIL_TO 為空，請先設定收件人")
        return False

    # ── 用 Outlook COM 建立郵件 ───────────────────────────────
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)   # 0 = olMailItem

        mail.Subject = subject
        mail.HTMLBody = body_html
        mail.To = "; ".join(to_list)
        if cc_list:
            mail.CC = "; ".join(cc_list)

        # 附加 HTML 報告
        if html_path.exists():
            mail.Attachments.Add(str(html_path.resolve()))
        else:
            print(f"  [WARN] 找不到 HTML 檔案：{html_path}")

        mail.Send()
        print(f"  OK 信件已寄出 → {', '.join(to_list)}")
        return True

    except Exception as e:
        print(f"  [ERROR] 寄信失敗：{e}")
        return False


def preview_report(report_date: date, html_path: Path, data: dict):
    """
    開啟 Outlook 草稿（不直接寄出，讓使用者先確認）
    """
    try:
        import win32com.client
    except ImportError:
        print("  [ERROR] 找不到 pywin32，請先執行：py -m pip install pywin32")
        return False

    import config

    alert_items = data.get("alert_items", [])
    date_str = report_date.strftime("%Y/%m/%d")
    subject_tpl = getattr(config, "EMAIL_SUBJECT", "【風控整合日報】{date}")
    subject = subject_tpl.format(date=date_str)
    to_list = getattr(config, "EMAIL_TO", [])
    cc_list = getattr(config, "EMAIL_CC", [])

    def _ai_color(t):
        return {"red": "#c62828", "yellow": "#b45309", "blue": "#1565c0"}.get(t, "#4a6080")
    def _ai_icon(t):
        return {"red": "🔴", "yellow": "🟡", "blue": "🔵"}.get(t, "🔵")

    items_html = ""
    for ai in alert_items:
        items_html += f'<p style="color:{_ai_color(ai["type"])};margin:4px 0;">{_ai_icon(ai["type"])} {ai["text"]}</p>'
    if not items_html:
        items_html = '<p style="color:#1a9e6a;">✅ 今日各項指標正常</p>'

    body_html = f"""<html><body style="font-family:system-ui;font-size:13px;">
<h3 style="color:#1565c0;border-bottom:1px solid #ccc;padding-bottom:8px;">
  風險管理整合日報　{date_str}
</h3>
<div style="background:#fef2f2;border-left:4px solid #c62828;padding:10px 14px;margin:12px 0;border-radius:4px;">
  <strong style="color:#c62828;font-size:11px;">⚡ 今日重點說明</strong><br>
  {items_html}
</div>
<p style="color:#4a6080;font-size:12px;">📎 完整報告請見附件：{html_path.name}</p>
<hr style="border:none;border-top:1px solid #e3e8ef;margin-top:20px;">
<p style="font-size:10px;color:#8a9bb5;">本信件由風控日報系統自動產生</p>
</body></html>"""

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.Subject = subject
        mail.HTMLBody = body_html
        mail.To = "; ".join(to_list)
        if cc_list:
            mail.CC = "; ".join(cc_list)
        if html_path.exists():
            mail.Attachments.Add(str(html_path.resolve()))
        mail.Display()   # 開啟草稿視窗，不直接寄出
        print("  OK Outlook 草稿已開啟，請確認後手動按送出")
        return True
    except Exception as e:
        print(f"  [ERROR] 開啟草稿失敗：{e}")
        return False
