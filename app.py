"""
FUR CPR Automation System - Backend v3
Dragon Studios for FUR Store (furstores.myshopify.com)
"""

import os, re, time, json, threading, uuid
from datetime import datetime
from flask import Flask, request, jsonify, render_template, send_file
from flask_cors import CORS
import pdfplumber
import requests
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from io import BytesIO

app = Flask(__name__)
CORS(app)

# In-memory stores
jobs    = {}   # job_id -> job data
uploads = {}   # upload_id -> parsed cpr_data

SHOPIFY_STORE = "furstores.myshopify.com"
API_VERSION   = "2024-10"


def get_access_token(credential):
    """
    Get Shopify access token.
    Supports:
    - shpat_xxx (legacy admin token) — use directly
    - client_id|client_secret (new Dev Dashboard) — exchange for token
    """
    credential = credential.strip()
    
    # All known token formats — use directly
    if (credential.startswith("shpat_") or 
        credential.startswith("atkn_") or
        len(credential) > 30):  # Any long token string, try directly
        return credential, None
    
    # New Dev Dashboard: Client ID|Secret format
    if "|" in credential:
        parts = credential.split("|", 1)
        client_id = parts[0].strip()
        client_secret = parts[1].strip()
        try:
            resp = requests.post(
                f"https://{SHOPIFY_STORE}/admin/oauth/access_token",
                json={
                    "client_id": client_id,
                    "client_secret": client_secret,
                    "grant_type": "client_credentials"
                },
                timeout=15
            )
            data = resp.json()
            if "access_token" in data:
                return data["access_token"], None
            return None, f"Token exchange failed: {data.get('error_description', str(data))}"
        except Exception as e:
            return None, f"Token exchange error: {e}"
    
    # Try as-is (might be a valid token we don't recognize)
    return credential, None


def shopify_gql(token, query, variables=None):
    """Safe Shopify GraphQL call — always returns a dict."""
    url = f"https://{SHOPIFY_STORE}/admin/api/{API_VERSION}/graphql.json"
    try:
        resp = requests.post(
            url,
            headers={
                "Content-Type": "application/json",
                "X-Shopify-Access-Token": token.strip()
            },
            json={"query": query, "variables": variables or {}},
            timeout=30
        )
        result = resp.json()
        if isinstance(result, dict):
            return result
        return {"errors": [{"message": f"Non-dict response: {str(result)[:100]}"}]}
    except Exception as e:
        return {"errors": [{"message": str(e)}]}


def parse_cpr_pdf(file_bytes):
    """Parse PostEx CPR PDF → delivered / returned lists."""
    tracking_re = re.compile(r'\b(\d{14})\b')
    cod_re      = re.compile(r'(\d{1,3}(?:,\d{3})*\.\d{2})')
    cpr_no_re   = re.compile(r'CPR[-\s]?\w+')
    date_re     = re.compile(r'\d{2}/\d{2}/\d{4}')

    delivered, returned = [], []
    cpr_number, cpr_date = "", ""
    summary = {}

    with pdfplumber.open(BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""

            if not cpr_number:
                m = cpr_no_re.search(text)
                if m:
                    cpr_number = m.group(0).strip()
            if not cpr_date:
                m = date_re.search(text)
                if m:
                    cpr_date = m.group(0)

            if "Delivered" in text and not summary:
                dm = re.search(r'Delivered\s+(\d+)\s+([\d,]+\.\d+)', text)
                rm = re.search(r'Returned\s+(\d+)\s+([\d,]+\.\d+)', text)
                if dm:
                    summary['delivered_count'] = int(dm.group(1))
                    summary['cod_collected']   = float(dm.group(2).replace(',', ''))
                if rm:
                    summary['returned_count'] = int(rm.group(1))

            lines = text.split('\n')
            for j, line in enumerate(lines):
                tm = tracking_re.search(line)
                if not tm:
                    continue
                tracking = tm.group(1)
                nxt = lines[j+1].strip() if j+1 < len(lines) else ""
                combined = line + " " + nxt

                if "Delivered" in combined:
                    status = "delivered"
                elif "Return" in combined:
                    status = "returned"
                else:
                    continue

                cods = cod_re.findall(combined)
                cod  = float(cods[0].replace(',', '')) if cods else 0.0

                entry = {"tracking": tracking, "cod": cod, "status": status}
                if status == "delivered":
                    delivered.append(entry)
                else:
                    returned.append(entry)

    return {
        "cpr_number": cpr_number,
        "cpr_date":   cpr_date,
        "delivered":  delivered,
        "returned":   returned,
        "summary":    summary,
    }


def fetch_shopify_orders(token, job_id):
    """Fetch all orders and index by tracking number."""
    all_orders = {}
    cursor = None
    page   = 0

    q = """
    query GetOrders($after: String) {
      orders(first: 50, after: $after, sortKey: CREATED_AT,
             query: "created_at:>=2026-05-01 created_at:<=2026-07-31") {
        pageInfo { hasNextPage endCursor }
        edges { node {
          id name displayFinancialStatus cancelledAt
          fulfillments(first: 1) { trackingInfo { number } }
        }}
      }
    }"""

    while True:
        page += 1
        data = shopify_gql(token, q, {"after": cursor} if cursor else {})

        if data.get("errors"):
            err0 = data["errors"][0]
            if isinstance(err0, dict):
                msg = err0.get("message", str(err0))
            else:
                msg = str(err0)
            if "401" in msg or "nvalid" in msg or "ccess" in msg:
                return None, "Token galat hai — Shopify Admin se naya shpat_ token banao aur dobara try karo"
            return None, f"Shopify error: {msg}"

        gdata = data.get("data") or {}
        orders_obj = gdata.get("orders") or {}
        edges = orders_obj.get("edges") or []

        for edge in edges:
            node = edge.get("node") or {}
            fuls = node.get("fulfillments") or []
            if fuls:
                trk_list = fuls[0].get("trackingInfo") or []
                if trk_list:
                    num = trk_list[0].get("number", "")
                    if num:
                        all_orders[num] = node

        jobs[job_id]["progress"] = min(40, page * 6)
        jobs[job_id]["log"].append({
            "type": "info",
            "msg": f"Page {page}: {len(all_orders)} orders indexed"
        })

        page_info = orders_obj.get("pageInfo") or {}
        if not page_info.get("hasNextPage"):
            break
        cursor = page_info.get("endCursor")
        time.sleep(0.25)

    return all_orders, None


def run_automation(job_id, token, cpr_data):
    """Main automation thread."""
    job = jobs[job_id]
    job["status"] = "running"

    def log(msg, t="info", order=None):
        job["log"].append({"type": t, "msg": msg, "order": order})

    try:
        # Resolve token (supports shpat_ and client_id|secret)
        resolved_token, token_err = get_access_token(token)
        if token_err:
            job["status"] = "error"
            job["error"] = token_err
            log(f"Token error: {token_err}", "error")
            return
        token = resolved_token
        
        log("═══ PHASE 1: Fetching Shopify orders ═══")
        all_orders, err = fetch_shopify_orders(token, job_id)
        if err:
            job["status"] = "error"
            job["error"]  = err
            log(err, "error")
            return

        log(f"Total orders indexed: {len(all_orders)}", "success")

        # Phase 2: Returns
        log("═══ PHASE 2: Verifying returns ═══")
        returns_voided, returns_not_canceled, returns_not_found = [], [], []

        for ret in (cpr_data.get("returned") or []):
            o = all_orders.get(ret["tracking"])
            if not o:
                returns_not_found.append(ret)
                log(f"Not in Shopify: {ret['tracking']}", "warn")
            elif o.get("displayFinancialStatus") == "VOIDED" or o.get("cancelledAt"):
                returns_voided.append({**ret, "orderName": o["name"], "orderId": o["id"]})
                log(f"Canceled ✓  {ret['tracking']} → #{o['name']}", "success", o["name"])
            else:
                returns_not_canceled.append({**ret, "orderName": o["name"],
                                             "orderId": o["id"], "status": o.get("displayFinancialStatus","")})
                log(f"NOT CANCELED: {ret['tracking']} → #{o['name']}", "error", o["name"])

        job["progress"] = 45

        # Phase 3: Mark paid
        log("═══ PHASE 3: Marking delivered orders as PAID ═══")
        paid, skipped, errors, not_found = [], [], [], []

        mutation = """
        mutation MarkPaid($id: ID!) {
          orderMarkAsPaid(input: { id: $id }) {
            order { name displayFinancialStatus }
            userErrors { field message }
          }
        }"""

        delivered = cpr_data.get("delivered") or []
        total = len(delivered)

        for i, d in enumerate(delivered):
            o = all_orders.get(d["tracking"])
            job["progress"] = 45 + int((i / max(total,1)) * 50)

            if not o:
                not_found.append(d)
                log(f"Not found: {d['tracking']}", "warn")
                continue

            name   = o.get("name", "?")
            status = o.get("displayFinancialStatus", "")

            if d["cod"] == 0:
                skipped.append({**d, "orderName": name, "reason": "COD=0"})
                log(f"COD=0 skip: #{name}", "skip", name)
                continue
            if status == "PAID":
                skipped.append({**d, "orderName": name, "reason": "Already PAID"})
                log(f"Already paid: #{name}", "skip", name)
                continue
            if status == "VOIDED" or o.get("cancelledAt"):
                skipped.append({**d, "orderName": name, "reason": "Canceled"})
                log(f"Canceled skip: #{name}", "skip", name)
                continue

            res  = shopify_gql(token, mutation, {"id": o["id"]})
            errs = (res.get("data") or {}).get("orderMarkAsPaid", {}).get("userErrors") or []

            if res.get("errors") or errs:
                msg = (res.get("errors") or [{}])[0].get("message") or (errs[0].get("message") if errs else "Unknown")
                errors.append({**d, "orderName": name, "error": msg})
                log(f"FAILED #{name}: {msg}", "error", name)
            else:
                paid.append({**d, "orderName": name, "orderId": o["id"]})
                log(f"PAID ✓  {d['tracking']} → #{name} · Rs {d['cod']:,.0f}", "success", name)

            time.sleep(0.2)

        job["progress"] = 100
        job["status"]   = "done"
        job["results"]  = {
            "cpr_number": cpr_data.get("cpr_number", ""),
            "cpr_date":   cpr_data.get("cpr_date", ""),
            "paid": paid, "skipped": skipped, "errors": errors, "not_found": not_found,
            "returns_voided": returns_voided,
            "returns_not_canceled": returns_not_canceled,
            "returns_not_found": returns_not_found,
        }

        paid_total = sum(o["cod"] for o in paid)
        log(f"DONE — Paid: {len(paid)} · Rs {paid_total:,.0f} | Skipped: {len(skipped)} | Errors: {len(errors)}", "success")
        log(f"Returns: {len(returns_voided)}/{len(cpr_data.get('returned',[]))} voided | Action needed: {len(returns_not_canceled)}", "info")

    except Exception as e:
        import traceback
        job["status"] = "error"
        job["error"]  = str(e)
        log(f"Unexpected error: {e}\n{traceback.format_exc()}", "error")


def make_excel_report(results):
    wb = openpyxl.Workbook()
    hf = PatternFill("solid", fgColor="1C3557")
    hFont = Font(color="FFFFFF", bold=True)
    gf = PatternFill("solid", fgColor="C6EFCE")
    rf = PatternFill("solid", fgColor="FFC7CE")
    af = PatternFill("solid", fgColor="FFEB9C")

    def sheet(ws, title, headers, rows, fill=None):
        ws.title = title
        ws.append(headers)
        for c in ws[1]:
            c.fill = hf; c.font = hFont
        for row in rows:
            ws.append(row)
            if fill:
                for c in ws[ws.max_row]:
                    c.fill = fill
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = 20

    ws0 = wb.active
    ws0.title = "Summary"
    ws0["A1"] = "FUR CPR Automation Report"
    ws0["A1"].font = Font(size=14, bold=True, color="1C3557")
    ws0["A2"] = f"CPR: {results.get('cpr_number','')} | Date: {results.get('cpr_date','')} | Generated: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    paid_total = sum(o["cod"] for o in results["paid"])
    for row in [
        ["METRIC","VALUE"], ["Marked PAID", len(results["paid"])],
        ["Total COD", f"Rs {paid_total:,.2f}"],
        ["Skipped", len(results["skipped"])], ["Errors", len(results["errors"])],
        ["Not Found", len(results["not_found"])],
        ["Returns Verified Voided", len(results["returns_voided"])],
        ["Returns Action Needed", len(results["returns_not_canceled"])],
    ]:
        ws0.append(row)
    ws0.column_dimensions["A"].width = 28
    ws0.column_dimensions["B"].width = 20

    sheet(wb.create_sheet(), "Paid Orders",
          ["Tracking", "Order", "COD"],
          [[o["tracking"], f"#{o['orderName']}", f"Rs {o['cod']:,.2f}"] for o in results["paid"]], gf)

    sheet(wb.create_sheet(), "Skipped",
          ["Tracking", "Order", "COD", "Reason"],
          [[o["tracking"], f"#{o.get('orderName','')}", f"Rs {o['cod']:,.2f}", o.get("reason","")] for o in results["skipped"]], af)

    if results["errors"]:
        sheet(wb.create_sheet(), "Errors",
              ["Tracking", "Order", "COD", "Error"],
              [[o["tracking"], f"#{o.get('orderName','')}", f"Rs {o['cod']:,.2f}", o.get("error","")] for o in results["errors"]], rf)

    sheet(wb.create_sheet(), "Returns Voided",
          ["Tracking", "Order", "Status"],
          [[o["tracking"], f"#{o.get('orderName','')}", "VOIDED"] for o in results["returns_voided"]], gf)

    if results["returns_not_canceled"]:
        sheet(wb.create_sheet(), "Returns-Action Needed",
              ["Tracking", "Order", "Status"],
              [[o["tracking"], f"#{o.get('orderName','')}", o.get("status","")] for o in results["returns_not_canceled"]], rf)

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ── Routes ──────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return jsonify({"error": "No file"}), 400
    f = request.files["file"]
    if not f.filename.lower().endswith(".pdf"):
        return jsonify({"error": "PDF only"}), 400
    try:
        data = parse_cpr_pdf(f.read())
        # Store server-side, return only an ID
        uid = str(uuid.uuid4())
        uploads[uid] = data
        return jsonify({
            "ok": True,
            "upload_id":  uid,
            "cpr_number": data["cpr_number"],
            "cpr_date":   data["cpr_date"],
            "delivered":  len(data["delivered"]),
            "returned":   len(data["returned"]),
            "summary":    data["summary"],
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/run", methods=["POST"])
def run_job():
    try:
        body = request.get_json(force=True, silent=True) or {}
    except Exception:
        body = {}

    token     = str(body.get("token") or "").strip()
    upload_id = str(body.get("upload_id") or "").strip()

    if not token:
        return jsonify({"error": "Shopify API token required"}), 400
    if not upload_id or upload_id not in uploads:
        return jsonify({"error": "CPR data not found — please re-upload the PDF"}), 400

    cpr_data = uploads[upload_id]

    job_id = f"job_{int(time.time()*1000)}"
    jobs[job_id] = {
        "id": job_id, "status": "queued",
        "progress": 0, "log": [], "results": None, "error": None,
        "started": datetime.now().isoformat(),
    }

    t = threading.Thread(target=run_automation, args=(job_id, token, cpr_data), daemon=True)
    t.start()
    return jsonify({"job_id": job_id})


@app.route("/api/status/<job_id>")
def job_status(job_id):
    job = jobs.get(job_id)
    if not job:
        return jsonify({"error": "Job not found"}), 404

    resp = {
        "status":   job["status"],
        "progress": job["progress"],
        "log":      job["log"],
        "error":    job.get("error"),
    }

    if job["status"] == "done" and job.get("results"):
        r = job["results"]
        resp["summary"] = {
            "paid":           len(r["paid"]),
            "paid_total":     sum(o["cod"] for o in r["paid"]),
            "skipped":        len(r["skipped"]),
            "errors":         len(r["errors"]),
            "not_found":      len(r["not_found"]),
            "returns_voided": len(r["returns_voided"]),
            "returns_action": len(r["returns_not_canceled"]),
        }
        if r["returns_not_canceled"]:
            resp["returns_not_canceled"] = r["returns_not_canceled"]

    return jsonify(resp)


@app.route("/api/report/<job_id>")
def download_report(job_id):
    job = jobs.get(job_id)
    if not job or job["status"] != "done":
        return jsonify({"error": "Job not complete"}), 404
    buf = make_excel_report(job["results"])
    cpr = job["results"].get("cpr_number", "CPR")
    return send_file(buf,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=f"FUR-{cpr}-{datetime.now().strftime('%Y-%m-%d')}.xlsx")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print(f"\n🚀 FUR CPR App v3 on http://localhost:{port}\n")
    app.run(debug=False, host="0.0.0.0", port=port)
