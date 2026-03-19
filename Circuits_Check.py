from flask import Flask, render_template, request,send_file
import requests
import pandas as pd
import re
import io

USERNAME = "mustaha"
PASSWORD = "mustaha@#$M2"

LOGIN_URL = "https://cm.earthlink.iq/api/auth"
Circuits_VLANs_URL = "https://cm.earthlink.iq/api/circuits/vlans/"
Circuits_URL = "https://cm.earthlink.iq/api/circuits/"

app = Flask(__name__)

# ---------- Border Port Capacity ----------
BORDER_PORT_CAPACITY = {
    "ARAR": 810000,
    "Badra": 700000,
    "IBK": 2050000,
    "Muntheria": 400000,
    "RAB3A": 2700000,
    "Safwan": 400000
}

# ---------- Login ----------
def get_token(username, password):
    response = requests.post(
        LOGIN_URL,
        json={"username": username, "password": password}
    )
    response.raise_for_status()
    token = response.json().get("value")
    if not token:
        raise Exception("Login failed")
    return token

token = get_token(USERNAME, PASSWORD)
headers = {"Authorization": f"Bearer {token}"}

# ---------- Fetch VLAN Data ----------
def fetch_vlans():
    all_data = []
    page = 1
    while True:
        params = {"page": page, "size": 100}
        response = requests.get(Circuits_VLANs_URL, headers=headers, params=params)
        response.raise_for_status()
        data_json = response.json()
        page_data = data_json["value"]["data"]
        if not page_data:
            break
        all_data.extend(page_data)
        total_pages = data_json["value"]["numberOfPages"]
        if page >= total_pages:
            break
        page += 1
    active_data = [item for item in all_data if not item.get("disabledAt")]
    rows = []
    for item in active_data:
        circuit = item.get("circuit")
        border_name = (
            circuit.get("border", {}).get("name")
            if isinstance(circuit, dict)
            else None
        )
        rows.append({
            "Service ID": str(item.get("serviceId", "")).strip(),
            "VLAN": str(item.get("vlan", "")).strip(),
            "Capacity": item.get("capacity") or 0,
            "Border Name": border_name
        })
    return pd.DataFrame(rows)

# ---------- Fetch Circuits ----------
def fetch_circuits():
    all_data = []
    page = 1
    while True:
        params = {"page": page, "size": 100}
        response = requests.get(Circuits_URL, headers=headers, params=params)
        response.raise_for_status()
        data_json = response.json()
        page_data = data_json["value"]["data"]
        if not page_data:
            break
        all_data.extend(page_data)
        total_pages = data_json["value"]["numberOfPages"]
        if page >= total_pages:
            break
        page += 1
    active_data = [item for item in all_data if not item.get("disabledAt")]
    rows = []
    for item in active_data:
        border = None
        if isinstance(item.get("border"), dict):
            border = item.get("border", {}).get("name")
        rows.append({
            "TotalCapacity": item.get("totalCapacity"),
            "Border": border,
            "SCIS": str(item.get("scis", "")).strip()
        })
    return pd.DataFrame(rows)

# ---------- Convert Gb to Mbps ----------
def convert_gb_to_mbps(value):
    try:
        number = float(value)
    except:
        number = 0
    if number < 0.01:
        return 0
    mbps = int(number * 1000)
    if mbps == 96000:
        mbps = 100000
    return mbps

# ---------- Build Port Capacity Map ----------
def build_port_capacity_map(df):
    mapping = {}
    for _, row in df.iterrows():
        scis_name = str(row.get("SCIS", "")).strip()
        m = re.fullmatch(r"(SCIS-\d{2})", scis_name)
        if m:
            port_value = convert_gb_to_mbps(row.get("TotalCapacity", 0))
            mapping[m.group(1)] = port_value
    return mapping

# ---------- Build D3 Tree ----------
def build_d3_tree(df):
    df = df.copy()
    patterns = {
        "FTTH": r"^V\d{2,4}\sFTTH$",
        "Wireless": r"^V\d{2,4}\sHulum(?:\sVIP)?$",
        "Hala FTTH": r"^V\d{2,4}\sVRF-HalaFTTH$",
        "Sur3at Albarq": r"^V\d{2,4}\sSur3at Albarq$",
        "Tech Resource": r"^V\d{2,4}\sVRF-TechRes$",
        "Skopesky": r"^V\d{2,4}\sVRF-SS$"
    }
    tree = {"name": "Total Capacity", "value": 0, "children": []}
    borders = []
    for border in df["Border Name"].dropna().unique():
        border_df = df[df["Border Name"] == border]
        border_node = {"name": border, "value": 0, "children": []}
        for name, pattern in patterns.items():
            value = border_df[
                border_df["VLAN"].str.contains(pattern, regex=True, na=False)
            ]["Capacity"].sum() * 1000
            if value > 0:
                border_node["children"].append({
                    "name": name,
                    "value": int(value)
                })
                border_node["value"] += int(value)
        if border_node["children"]:
            borders.append(border_node)
    borders.sort(key=lambda x: x["value"], reverse=True)
    total_sum = sum(b["value"] for b in borders)
    tree["children"] = borders
    tree["value"] = int(total_sum)
    return tree

# ---------- Border Used Capacity ----------
def build_border_used_capacity(tree):
    used_map = {}
    for border in tree.get("children", []):
        name = border.get("name")
        value = border.get("value", 0)
        used_map[name] = value
    return used_map

# ---------- Load Data ----------
df_vlans = fetch_vlans()
df_circuits = fetch_circuits()
port_capacity_map = build_port_capacity_map(df_circuits)
capacity_tree = build_d3_tree(df_vlans)
border_used_capacity = build_border_used_capacity(capacity_tree)

# ---------- SCIS Options ----------

def get_scis_options(df):
    scis_with_border = []
    seen_scis = set()  # لتجنب التكرار

    for sid in df.get("Service ID", []):
        sid = str(sid).strip()
        if sid.startswith("SCIS-"):
            num_part = sid.split("-")[1].zfill(2)
            if num_part in seen_scis:
                continue  # نتجاوز إذا سبق إضافته
            circuit_df = df[df["Service ID"].str.contains(f"SCIS-{num_part}")]
            total_capacity = circuit_df["Capacity"].sum()
            if total_capacity > 0:
                border_name = circuit_df["Border Name"].iloc[0] if not circuit_df.empty else ""
                scis_with_border.append((num_part, border_name))
                seen_scis.add(num_part)

    # ترتيب حسب البورد ثم الرقم
    border_order = list(BORDER_PORT_CAPACITY.keys())
    scis_with_border.sort(key=lambda x: (border_order.index(x[1]) if x[1] in border_order else 999, int(x[0])))
    return [x[0] for x in scis_with_border]
# ---------- Prepare SCIS Table with Rowspan ----------
def prepare_table_with_rowspan(df, selected_scis):
    if selected_scis == "all":
        tables = []
        for scis in get_scis_options(df):
            table = prepare_table_with_rowspan(df, scis)
            if table:
                tables.append(table)
        return tables
    else:
        filtered_df = df[df['Service ID'].str.contains(f"-{selected_scis}")]
        filtered_df = filtered_df[filtered_df['Capacity'] > 0]  # فقط Capacities > 0
        if filtered_df.empty:
            return None

        # تحويل Gb إلى Mbps
        filtered_df["Capacity"] = filtered_df["Capacity"].apply(lambda x: int(x * 1000))
        filtered_df = filtered_df.drop_duplicates(subset=['VLAN', 'Capacity'])
        total_capacity = filtered_df["Capacity"].sum()
        filtered_df = filtered_df.sort_values(by="Capacity", ascending=False)

        border_name = filtered_df["Border Name"].iloc[0]  # <-- هنا نضيف Border
        vlan_rows = [{"VLAN": row["VLAN"], "Capacity (Mbps)": f"{row['Capacity']:,}", "Border": border_name} 
                     for _, row in filtered_df.iterrows()]  # نضيف Border لكل صف

        circuit_name = f"SCIS-{selected_scis}"
        port_capacity = port_capacity_map.get(circuit_name, 0)

        return {
            "Circuit Name": circuit_name,
            "Port Capacity": f"{port_capacity:,}",
            "Border Name": border_name,  # احتفظ بالبوردر للـ merge لو حبيت
            "Total (Mbps)": f"{total_capacity:,}",
            "rows": vlan_rows,
            "rowspan": len(vlan_rows)
        }

# ---------- Prepare Border Table (existing) ----------
def prepare_border_table(selected_border):
    if selected_border == "all":
        all_tables = []
        for border in BORDER_PORT_CAPACITY.keys():
            table = prepare_border_table(border)
            if table and table['rows']:
                all_tables.append(table)
        return all_tables

    else:
        border_df = df_circuits[df_circuits["Border"] == selected_border]
        all_rows = []

        for _, row in border_df.iterrows():
            circuit = row["SCIS"]
            port = convert_gb_to_mbps(row["TotalCapacity"])
            if port == 0:
                continue

            vlans_df = df_vlans[df_vlans['Service ID'].str.contains(circuit)]
            vlans_df = vlans_df[vlans_df['Capacity'] > 0]
            vlans_df = vlans_df.drop_duplicates(subset=['VLAN', 'Capacity'])
            vlans_df = vlans_df.sort_values(by='Capacity', ascending=False)

            if vlans_df.empty:
                circuit_rows = [{
                    "Circuit": circuit,
                    "Port": f"{port:,}",
                    "VLAN": "No VLAN",
                    "Capacity (Mbps)": "No Capacity"
                }]
                total_capacity_circuit = "No Capacity"
            else:
                circuit_rows = []
                for _, vlan_row in vlans_df.iterrows():
                    capacity_value = int(vlan_row["Capacity"] * 1000)
                    circuit_rows.append({
                        "Circuit": circuit,
                        "Port": f"{port:,}",
                        "VLAN": vlan_row["VLAN"],
                        "Capacity (Mbps)": f"{capacity_value:,}"
                    })
                total_capacity_circuit = f"{int(vlans_df['Capacity'].sum() * 1000):,}"

            for idx, r in enumerate(circuit_rows):
                r['total_capacity_circuit'] = total_capacity_circuit
                if idx == 0:
                    r['rowspan_circuit'] = len(circuit_rows)
                    r['show_circuit'] = True
                    r['show_port'] = True
                else:
                    r['rowspan_circuit'] = 0
                    r['show_circuit'] = False
                    r['show_port'] = False
                all_rows.append(r)

        all_rows = sorted(
            all_rows,
            key=lambda x: -int(x["Port"].replace(",", ""))
        )

        border_total = BORDER_PORT_CAPACITY.get(selected_border, 0)
        used_capacity = border_used_capacity.get(selected_border, 0)
        percent = int((used_capacity / border_total) * 100) if border_total > 0 else 0

        return {
            "Border": selected_border,
            "Border Port": f"{border_total:,}",
            "Used Capacity": f"{used_capacity:,} ({percent}%)",
            "Used_Percent": percent,
            "rows": all_rows,
            "rowspan": len(all_rows)
        }

# ---------- Flask ----------
@app.route("/", methods=["GET", "POST"])
def index():
    selected_scis = None
    table = None
    selected_border = None
    border_table = None
    scis_options = get_scis_options(df_vlans)
    border_options = list(BORDER_PORT_CAPACITY.keys()) 
    if request.method == "POST":
        if request.form.get("scis_select"):
            selected_scis = request.form.get("scis_select")
            if selected_scis == "all":
                table = prepare_table_with_rowspan(df_vlans, "all")
            else:
                table = prepare_table_with_rowspan(df_vlans, selected_scis)
        if request.form.get("border_select"):
            selected_border = request.form.get("border_select")
            if selected_border == "all":
                border_table = []
                for b in BORDER_PORT_CAPACITY.keys():
                    t = prepare_border_table(b)
                    if t:
                        border_table.append(t)
            else:
                border_table = prepare_border_table(selected_border)
    return render_template(
        "index.html",
        scis_options=scis_options,
        selected_scis=selected_scis,
        table=table,
        capacity_tree=capacity_tree,
        border_options=border_options,
        selected_border=selected_border,
        border_table=border_table
    )




@app.route("/export", methods=["POST"])
def export_excel():
    import io
    from flask import request, send_file
    import pandas as pd

    def parse_number(value):
        """
        يحول القيمة إلى int مع إزالة النسبة المئوية، أو يرجع 0
        """
        try:
            value_str = str(value).split("(")[0]  # إزالة النسبة
            return int(value_str.replace(",", "").strip())
        except:
            return 0

    selected_scis = request.form.get("scis_select")
    selected_border = request.form.get("border_select")

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet("Data")
        writer.sheets["Data"] = worksheet

        # ===== تنسيقات =====
        header_format = workbook.add_format({
            'bold': True, 'align': 'center', 'valign': 'vcenter',
            'border': 1, 'bg_color': '#2563eb', 'color': 'white'
        })

        cell_format = workbook.add_format({
            'align': 'center', 'valign': 'vcenter', 'border': 1
        })

        number_format = workbook.add_format({
            'align': 'center', 'valign': 'vcenter',
            'border': 1, 'num_format': '#,##0'
        })

        green = workbook.add_format({'align':'center','valign':'vcenter','border':1,'bg_color':'#16a34a','color':'white'})
        yellow = workbook.add_format({'align':'center','valign':'vcenter','border':1,'bg_color':'#f59e0b','color':'white'})
        red = workbook.add_format({'align':'center','valign':'vcenter','border':1,'bg_color':'#dc2626','color':'white'})

        row = 0

        # ===== تحديد الوضع =====
        tables = []
        is_border = False

        if selected_border:
            is_border = True
            if selected_border == "all":
                tables = [prepare_border_table(b) for b in BORDER_PORT_CAPACITY if prepare_border_table(b)]
            else:
                tables = [prepare_border_table(selected_border)]

        elif selected_scis:
            if selected_scis == "all":
                scis_list = get_scis_options(df_vlans)
                tables = [prepare_table_with_rowspan(df_vlans, s) for s in scis_list]
                tables = [t for t in tables if t]
                tables.sort(key=lambda x: x["Border Name"])
            else:
                t = prepare_table_with_rowspan(df_vlans, selected_scis)
                if t:
                    tables = [t]

        else:
            return "No data selected"

        # ===== رؤوس الأعمدة =====
        if is_border:
            headers = [
                "Border",
                "Port Capacity (Mbps)",
                "Used Capacity (Mbps)",
                "Circuit",
                "Port Capacity (Mbps)",
                "VLAN",
                "Capacity (Mbps)",
                "Total (Mbps)"
            ]
        else:
            headers = [
                "Circuit",
                "Port Capacity (Mbps)",
                "VLAN",
                "Capacity (Mbps)",
                "Total (Mbps)",
                "Border"
            ]

        for col, h in enumerate(headers):
            worksheet.write(row, col, h, header_format)
        row += 1

        # ===== كتابة البيانات =====
        for t in tables:
            if not t:
                continue

            # ===== BORDER =====
            if is_border:
                rows = t["rows"]
                span = len(rows)

                border = t["Border"]
                port = parse_number(t["Border Port"])
                used = parse_number(t["Used Capacity"])
                percent = t["Used_Percent"]

                fmt = green if percent <= 60 else yellow if percent <= 80 else red

                # كتابة Border وPort وUsed Capacity
                worksheet.merge_range(row, 0, row+span-1, 0, border, cell_format)
                worksheet.merge_range(row, 1, row+span-1, 1, port, number_format)
                worksheet.merge_range(row, 2, row+span-1, 2, f"{used:,} ({percent}%)", fmt)

                i = 0
                while i < len(rows):
                    r = rows[i]
                    cspan = r.get("rowspan_circuit", 1)

                    if cspan > 1:
                        worksheet.merge_range(row, 3, row+cspan-1, 3, r["Circuit"], cell_format)
                        worksheet.merge_range(row, 4, row+cspan-1, 4, parse_number(r["Port"]), number_format)
                        worksheet.merge_range(row, 7, row+cspan-1, 7, parse_number(r["total_capacity_circuit"]), number_format)
                    else:
                        # صف واحد → نترك الأعمدة فارغة عدا البيانات
                        worksheet.write(row, 3, r["Circuit"], cell_format)
                        worksheet.write(row, 4, parse_number(r["Port"]), number_format)
                        worksheet.write(row, 7, parse_number(r["total_capacity_circuit"]), number_format)

                    for j in range(cspan):
                        rr = rows[i+j]
                        worksheet.write(row, 5, rr["VLAN"], cell_format)
                        worksheet.write(row, 6, parse_number(rr["Capacity (Mbps)"]), number_format)
                        row += 1

                    i += cspan

            # ===== SCIS =====
            else:
                rows = t["rows"]
                span = len(rows)

                circuit = t["Circuit Name"]
                port = parse_number(t["Port Capacity"])
                total = parse_number(t["Total (Mbps)"])
                border = t["Border Name"]

                if span > 1:
                    worksheet.merge_range(row, 0, row+span-1, 0, circuit, cell_format)
                    worksheet.merge_range(row, 1, row+span-1, 1, port, number_format)
                    worksheet.merge_range(row, 4, row+span-1, 4, total, number_format)
                    worksheet.merge_range(row, 5, row+span-1, 5, border, cell_format)
                else:
                    # صف واحد → نترك الأعمدة فارغة عدا البيانات
                    worksheet.write(row, 0, circuit, cell_format)
                    worksheet.write(row, 1, port, number_format)
                    worksheet.write(row, 4, total, number_format)
                    worksheet.write(row, 5, border, cell_format)

                for r in rows:
                    worksheet.write(row, 2, r["VLAN"], cell_format)
                    worksheet.write(row, 3, parse_number(r["Capacity (Mbps)"]), number_format)
                    row += 1

        worksheet.set_column(0, 10, 22)

    output.seek(0)

    return send_file(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="Circuit_Data.xlsx"
    )
    
    
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
    
