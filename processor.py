"""
processor.py — Core Sales Data Engine & Accounting Formulas
Implements the exact logic of Access query `qryDailySales`:
- Filters valid data lines: WHERE F1 Between 1 And 999
- Deduplicates against tbl_old
- Intelligent showroom resolution (mapping to Short name, CityNo, Branch)
- Exact mathematical formulas:
    Total (المبيعات) = NetTotal + Returns (where Returns is negative)
    Commissions rounded to integer (DecimalPlaces = 0)
    Sorting: ORDER BY CityNo ASC, Showroom (Short) ASC, ClosingDate ASC
- Provides detailed rows, grouped showroom totals, and grand totals.
"""

import os
import re
from datetime import datetime, timedelta
from typing import List, Dict, Any, Optional, Set, Tuple
import pandas as pd
from sqlalchemy.orm import Session, joinedload

from database import SessionLocal, Showroom, City, Branch, OldClosing


def parse_date_value(val: Any) -> Tuple[str, Optional[int]]:
    """
    Parses date and returns (formatted_date_str 'DD/MM/YYYY', closing_hour_int).
    Closing hour (0-23) is used for color-coded audit conditional formatting.
    """
    if pd.isna(val) or val is None or val == "":
        return "", None

    # Excel serial number handling (e.g., 46321 -> 26/10/2026)
    try:
        num = float(val)
        # Check if it looks like an Excel serial date (e.g. 40000 to 60000)
        if 30000 <= num <= 70000:
            dt = datetime(1899, 12, 30) + timedelta(days=num)
            return dt.strftime("%d/%m/%Y"), dt.hour
    except (ValueError, TypeError):
        pass

    s = str(val).strip()
    hour = None

    # Handle datetime string formats
    formats = [
        ("%Y-%m-%d %H:%M:%S", True),
        ("%d/%m/%Y %H:%M:%S", True),
        ("%Y/%m/%d %H:%M:%S", True),
        ("%Y-%m-%dT%H:%M:%S", True),
        ("%Y-%m-%d", False),
        ("%d/%m/%Y", False),
        ("%Y/%m/%d", False),
    ]

    for fmt, has_time in formats:
        try:
            dt = datetime.strptime(s, fmt)
            if has_time:
                hour = dt.hour
            return dt.strftime("%d/%m/%Y"), hour
        except ValueError:
            continue

    return s, None


def normalize_text(text_str: str) -> str:
    """Normalizes string for robust fuzzy/containment matching."""
    if not text_str:
        return ""
    s = str(text_str).strip()
    s = re.sub(r"[\s\(\)\.\-\_\/]+", " ", s)
    return s.strip()


class ShowroomResolver:
    """Fast in-memory resolver matching raw Odoo descriptions to canonical showrooms."""
    def __init__(self, db: Session):
        self.showrooms = db.query(Showroom).options(
            joinedload(Showroom.city).joinedload(City.branch)
        ).filter(Showroom.is_active == True).all()
        self.mapping = {}
        for s in self.showrooms:
            self.mapping[s.DESC.strip()] = s
            self.mapping[normalize_text(s.DESC)] = s
            self.mapping[s.Short.strip()] = s
            self.mapping[normalize_text(s.Short)] = s
            if s.No:
                self.mapping[f"NO_{s.No.strip()}"] = s

    def resolve(self, raw_name: str) -> Optional[Showroom]:
        if not raw_name:
            return None
        name_clean = str(raw_name).strip()

        # 1. Exact match
        if name_clean in self.mapping:
            return self.mapping[name_clean]

        # 2. Normalized match
        norm = normalize_text(name_clean)
        if norm in self.mapping:
            return self.mapping[norm]

        # 3. Special handling for Al-Bawadi vs Al-Bawadi MS
        has_ms = bool(re.search(r"\bMS\b", name_clean, re.IGNORECASE))
        if "البوادي" in name_clean:
            for s in self.showrooms:
                if has_ms and s.No == "122":  # البوادي MS
                    return s
                elif not has_ms and s.No == "112":  # البوادي
                    return s

        # 4. Check for 3-digit code
        codes = re.findall(r"\b(\d{3})\b", raw_name)
        for c in codes:
            if f"NO_{c}" in self.mapping:
                return self.mapping[f"NO_{c}"]

        # 5. Longest short-name substring match
        sorted_showrooms = sorted(self.showrooms, key=lambda s: len(s.Short or ""), reverse=True)
        for s in sorted_showrooms:
            if s.Short and s.Short in name_clean:
                return s
            arabic_part = "".join(re.findall(r"[\u0600-\u06FF\s]+", name_clean)).strip()
            if arabic_part and (arabic_part in s.DESC or s.Short in arabic_part):
                return s

        return None


def to_float(val: Any) -> float:
    """Safe float conversion."""
    if val is None or pd.isna(val) or val == "":
        return 0.0
    try:
        if isinstance(val, (int, float)):
            return float(val)
        s = str(val).replace(",", "").strip()
        return float(s)
    except Exception:
        return 0.0


class OldClosingsFileUploadedError(Exception):
    """Raised when the uploaded file is recognized as an old closings archive (tblOLD / qryExportOldDetailed)."""
    def __init__(self, message: str, stats: Dict[str, Any]):
        super().__init__(message)
        self.stats = stats


class SalesDataProcessor:
    """
    Parses raw daily sales data (from Excel or Odoo XML-RPC), applies Access qryDailySales
    accounting rules, filters previously approved sessions, and produces data ready for reports.
    """
    def __init__(self, db: Optional[Session] = None):
        self.db = db or SessionLocal()
        self.resolver = ShowroomResolver(self.db)

    def get_old_closing_numbers(self) -> Set[str]:
        """Fetches all closing numbers currently in tbl_old via fast scalar index scan."""
        from sqlalchemy import select
        closings = self.db.execute(select(OldClosing.closing_no)).scalars().all()
        return set(c.strip() for c in closings if c)

    def import_old_closings_file(
        self,
        file_path_or_bytes: Any,
        batch_id: str = "MANUAL_IMPORT"
    ) -> Dict[str, Any]:
        """
        Parses an Excel or CSV file containing old closings (e.g. qryExportOldDetailed.xlsx),
        identifies closing numbers, and bulk-inserts new ones into tbl_old.
        """
        if hasattr(file_path_or_bytes, "seek"):
            file_path_or_bytes.seek(0)

        # Handle CSV vs Excel
        try:
            if isinstance(file_path_or_bytes, str) and file_path_or_bytes.endswith(".csv"):
                df = pd.read_csv(file_path_or_bytes)
            else:
                df = pd.read_excel(file_path_or_bytes)
        except Exception:
            if hasattr(file_path_or_bytes, "seek"):
                file_path_or_bytes.seek(0)
            df = pd.read_csv(file_path_or_bytes)

        # Detect closing number column
        c_no_col = None
        for col in df.columns:
            clean_c = str(col).strip().lower().replace("_", "").replace(" ", "")
            if clean_c in ["closingno", "closingnum", "closingnumber", "رقمإقفال", "رقمالإقفال"]:
                c_no_col = col
                break
        
        if c_no_col is None:
            # Look for column containing 'POS/'
            for col in df.columns:
                sample_vals = df[col].dropna().astype(str).head(10).tolist()
                if any(v.startswith("POS/") for v in sample_vals):
                    c_no_col = col
                    break
        
        if c_no_col is None:
            raise ValueError("لم يتم العثور على عمود أرقام الإقفالات (ClosingNO أو رقم الإقفال) في هذا الملف.")

        # Optional columns
        showroom_col = next((c for c in df.columns if any(k in str(c).lower() for k in ["showroom", "معرض", "فرع"])), None)
        date_col = next((c for c in df.columns if any(k in str(c).lower() for k in ["date", "تاريخ"])), None)
        total_col = next((c for c in df.columns if any(k in str(c).lower() for k in ["nettotal", "total", "إجمالي", "مبيعات"])), None)

        existing = set(r[0] for r in self.db.query(OldClosing.closing_no).all())
        total_in_file = len(df)
        new_objects = []

        for _, row in df.iterrows():
            raw_c_no = row.get(c_no_col)
            if pd.isna(raw_c_no):
                continue
            c_no = str(raw_c_no).strip()
            if not c_no or c_no in existing:
                continue

            existing.add(c_no)
            sh_val = str(row.get(showroom_col, "")).strip() if showroom_col and pd.notna(row.get(showroom_col)) else ""
            date_val = str(row.get(date_col, "")).strip() if date_col and pd.notna(row.get(date_col)) else ""
            net_val = to_float(row.get(total_col, 0.0)) if total_col and pd.notna(row.get(total_col)) else 0.0

            new_objects.append(OldClosing(
                closing_no=c_no,
                showroom_short=sh_val,
                closing_date=date_val,
                net_total=net_val,
                batch_id=batch_id
            ))

        if new_objects:
            self.db.bulk_save_objects(new_objects)
            self.db.commit()

        final_total = self.db.query(OldClosing).count()
        return {
            "total_in_file": total_in_file,
            "new_added": len(new_objects),
            "already_existed": total_in_file - len(new_objects),
            "final_total": final_total
        }

    def parse_excel_file(
        self,
        file_path_or_bytes: Any,
        exclude_old: bool = True
    ) -> List[Dict[str, Any]]:
        """
        Reads Excel file (`Pos Closing Daily Sales.xlsx` or `.csv`), applies `qryDailySales` filter:
        WHERE F1 Between '1' And '999' AND F3 Not In (SELECT closing_no FROM tbl_old).
        """
        old_closings = self.get_old_closing_numbers() if exclude_old else set()

        if hasattr(file_path_or_bytes, "seek"):
            file_path_or_bytes.seek(0)

        # Handle both CSV and Excel files
        if isinstance(file_path_or_bytes, str) and file_path_or_bytes.endswith(".csv"):
            lines = []
            with open(file_path_or_bytes, "r", encoding="utf-8-sig") as f:
                for line in f:
                    lines.append(line.strip().split(","))
        else:
            df_raw = pd.read_excel(file_path_or_bytes, header=None)
            lines = df_raw.values.tolist()

        # Check if the uploaded file is actually an Old Closings export (e.g., qryExportOldDetailed)
        first_row_str = " ".join(str(x).lower() for x in lines[0] if pd.notna(x)) if lines else ""
        if "closingno" in first_row_str or "qryexportold" in str(file_path_or_bytes).lower():
            stats = self.import_old_closings_file(file_path_or_bytes, batch_id="FILE_UPLOAD_OLD")
            raise OldClosingsFileUploadedError(
                f"تم التعرف على هذا الملف كأرشيف إقفالات معتمدة (tblOLD). تم استيراد وتحديث {stats['new_added']} إقفال جديد في قاعدة البيانات (الإجمالي: {stats['final_total']}).",
                stats=stats
            )

        # Locate header row containing 'م' and 'الفرع'
        header_idx = -1
        for i, r in enumerate(lines):
            r_str = [str(x) for x in r if pd.notna(x)]
            if len(r_str) > 2 and (r_str[0].strip() == "م" or "م" in r_str[0]) and any("الفرع" in s for s in r_str[:3]):
                header_idx = i
                break

        if header_idx == -1:
            # Fallback: header might be on row 7
            header_idx = 7 if len(lines) > 7 else 0

        raw_rows = lines[header_idx + 1:]
        records = []

        for r in raw_rows:
            if len(r) < 10:
                continue
            f1 = str(r[0]).strip() if pd.notna(r[0]) else ""
            # Filter F1 Between "1" And "999" (ignoring header or summary totals)
            if not (f1.isdigit() and 1 <= int(f1) <= 999):
                continue

            raw_desc = str(r[1]).strip() if pd.notna(r[1]) else ""
            closing_no = str(r[2]).strip() if pd.notna(r[2]) else ""

            # Exclude already approved sessions (matching Access WHERE F3 Not In tblOLD)
            if exclude_old and closing_no in old_closings:
                continue

            # Resolve showroom details
            sh = self.resolver.resolve(raw_desc)
            short_name = sh.Short if sh else raw_desc
            city_no = sh.CityNo if sh and sh.CityNo else "999"
            city_name = sh.city.CityName if sh and sh.city else "غير محدد"
            branch_name = sh.city.branch.BranchName if sh and sh.city and sh.city.branch else "عام"

            # Parse date & time
            date_raw = r[3] if len(r) > 3 else ""
            date_str, closing_hour = parse_date_value(date_raw)

            # Amounts
            raw_open = to_float(r[5]) if len(r) > 5 else 0.0
            raw_returns = to_float(r[6]) if len(r) > 6 else 0.0  # Usually negative, e.g. -232.00
            raw_tax = to_float(r[7]) if len(r) > 7 else 0.0
            raw_net_sales = to_float(r[8]) if len(r) > 8 else 0.0
            raw_net_total = to_float(r[9]) if len(r) > 9 else 0.0

            # CRITICAL ACCESS FORMULA: Total (المبيعات) = NetTotal (الإجمالي) + Returns (المردودات)
            # Example: 3580.90 + (-232.00) = 3348.90
            calculated_total = raw_net_total + raw_returns

            raw_credit = to_float(r[10]) if len(r) > 10 else 0.0
            raw_trans = to_float(r[11]) if len(r) > 11 else 0.0
            raw_cash = to_float(r[12]) if len(r) > 12 else 0.0
            raw_comm = round(to_float(r[13])) if len(r) > 13 else 0  # Rounded to integer
            raw_exp = to_float(r[14]) if len(r) > 14 else 0.0
            raw_diff = to_float(r[15]) if len(r) > 15 else 0.0

            records.append({
                "selected": True,
                "is_old": False,
                "seq": int(f1),
                "raw_desc": raw_desc,
                "showroom": short_name,
                "city_no": city_no,
                "city_name": city_name,
                "branch_name": branch_name,
                "closing_no": closing_no,
                "date": date_str,
                "closing_hour": closing_hour,
                "total": calculated_total,
                "returns": raw_returns,
                "tax": raw_tax,
                "net_sales": raw_net_sales,
                "net_total": raw_net_total,
                "open": raw_open,
                "credit": raw_credit,
                "trans": raw_trans,
                "cash": raw_cash,
                "commissions": raw_comm,
                "expenses": raw_exp,
                "diff": raw_diff,
                "acc_adjustment": 0.0,
            })

        # Order BY CityNo ASC, Showroom ASC, date ASC
        records.sort(key=lambda x: (x["city_no"], x["showroom"], x["date"], x["closing_no"]))
        return records

    def process_selected_records(
        self,
        records: List[Dict[str, Any]]
    ) -> Dict[str, Any]:
        """
        Takes user-selected (checkbox checked) records and prepares:
        1. detailed_rows: For rep00DailySalesReport and rep01DailySalesReportColor
        2. totals_rows: Compact rows aggregated by (CityNo, Showroom) for rep00DailySalesReportTotals
        3. grand_totals: Computed sums across all columns + unique showroom count
        """
        active_records = [r for r in records if r.get("selected", True) and not r.get("is_old", False)]

        # Sort: ORDER BY CityNo ASC, Showroom ASC, date ASC
        active_records.sort(key=lambda x: (x["city_no"], x["showroom"], x["date"], x["closing_no"]))

        # 1. Detailed rows
        detailed_rows = []
        for idx, r in enumerate(active_records, start=1):
            row_copy = dict(r)
            row_copy["index"] = idx
            detailed_rows.append(row_copy)

        # 2. Compact Totals: Aggregate by (CityNo, Showroom)
        showroom_groups: Dict[Tuple[str, str], Dict[str, Any]] = {}
        for r in active_records:
            key = (r["city_no"], r["showroom"])
            if key not in showroom_groups:
                showroom_groups[key] = {
                    "city_no": r["city_no"],
                    "showroom": r["showroom"],
                    "city_name": r["city_name"],
                    "net_sales": 0.0,
                    "net_total": 0.0,
                    "credit": 0.0,
                    "trans": 0.0,
                    "commissions": 0,
                    "acc_adjustment": 0.0,
                    "sessions_count": 0,
                }
            g = showroom_groups[key]
            g["net_sales"] += r["net_sales"]
            g["net_total"] += r["net_total"]
            g["credit"] += r["credit"]
            g["trans"] += r["trans"]
            g["commissions"] += r["commissions"]
            g["acc_adjustment"] += r.get("acc_adjustment", 0.0)
            g["sessions_count"] += 1

        # Sort totals by CityNo ASC, Showroom ASC
        sorted_groups = sorted(showroom_groups.values(), key=lambda x: (x["city_no"], x["showroom"]))
        totals_rows = []
        for idx, g in enumerate(sorted_groups, start=1):
            g_copy = dict(g)
            g_copy["index"] = idx
            totals_rows.append(g_copy)

        # 3. Grand Totals
        unique_showrooms_count = len(showroom_groups)
        grand_totals = {
            "unique_showrooms_count": unique_showrooms_count,
            "sessions_count": len(active_records),
            "total": sum(r["total"] for r in active_records),
            "returns": sum(r["returns"] for r in active_records),
            "tax": sum(r["tax"] for r in active_records),
            "net_sales": sum(r["net_sales"] for r in active_records),
            "net_total": sum(r["net_total"] for r in active_records),
            "open": sum(r["open"] for r in active_records),
            "credit": sum(r["credit"] for r in active_records),
            "trans": sum(r["trans"] for r in active_records),
            "cash": sum(r["cash"] for r in active_records),
            "commissions": sum(r["commissions"] for r in active_records),
            "expenses": sum(r["expenses"] for r in active_records),
            "diff": sum(r["diff"] for r in active_records),
            "acc_adjustment": sum(r.get("acc_adjustment", 0.0) for r in active_records),
        }

        return {
            "detailed_rows": detailed_rows,
            "totals_rows": totals_rows,
            "grand_totals": grand_totals,
        }


def generate_periodic_excel(period_data: Dict[str, Any]) -> bytes:
    """
    Generates an executive Excel (.xlsx) file for any periodic showroom sales dataset (Monthly, Quarterly, Annual, etc.).
    - Columns: م, المعرض, الإجمالي, الضريبة, الصافي
    - Sorted by Net sales descending
    - Formatted currencies, borders, RTL view, and grand totals row.
    """
    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from io import BytesIO

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = str(period_data.get("period_label", "تقرير المبيعات"))[:30]
    ws.views.sheetView[0].rightToLeft = True

    # 1. Main Title
    ws.merge_cells("A1:E1")
    title_cell = ws["A1"]
    title_cell.value = period_data["title"]
    title_cell.font = Font(name="Calibri", size=15, bold=True, color="1F4E79")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 34

    # 2. Table Headers at row 3
    headers = ["م", "المعرض", "الإجمالي", "الضريبة", "الصافي"]
    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="000000")
    thin_border = Border(
        left=Side(style="thin", color="808080"),
        right=Side(style="thin", color="808080"),
        top=Side(style="thin", color="808080"),
        bottom=Side(style="thin", color="808080")
    )

    ws.row_dimensions[3].height = 25
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col_idx, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    # 3. Data rows
    for r_idx, r in enumerate(period_data["rows"], start=4):
        ws.row_dimensions[r_idx].height = 20

        c_seq = ws.cell(row=r_idx, column=1, value=r["seq"])
        c_seq.alignment = Alignment(horizontal="center", vertical="center")
        c_seq.border = thin_border

        c_sh = ws.cell(row=r_idx, column=2, value=r["showroom"])
        c_sh.alignment = Alignment(horizontal="right", vertical="center")
        c_sh.border = thin_border

        c_tot = ws.cell(row=r_idx, column=3, value=r["total"])
        c_tot.number_format = "#,##0.00"
        c_tot.alignment = Alignment(horizontal="center", vertical="center")
        c_tot.border = thin_border

        c_tax = ws.cell(row=r_idx, column=4, value=r["tax"])
        c_tax.number_format = "#,##0.00"
        c_tax.alignment = Alignment(horizontal="center", vertical="center")
        c_tax.border = thin_border

        c_net = ws.cell(row=r_idx, column=5, value=r["net"])
        c_net.number_format = "#,##0.00"
        c_net.font = Font(name="Calibri", size=10, bold=True)
        c_net.alignment = Alignment(horizontal="center", vertical="center")
        c_net.border = thin_border

    # 4. Total row
    last_row = 4 + len(period_data["rows"])
    ws.row_dimensions[last_row].height = 25
    total_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    double_border = Border(
        left=Side(style="thin", color="808080"),
        right=Side(style="thin", color="808080"),
        top=Side(style="thin", color="808080"),
        bottom=Side(style="double", color="000000")
    )

    c_tot_label = ws.cell(row=last_row, column=1, value="المجموع")
    c_tot_label.font = Font(name="Calibri", size=11, bold=True)
    c_tot_label.alignment = Alignment(horizontal="center", vertical="center")
    c_tot_label.fill = total_fill
    c_tot_label.border = double_border

    c_sh_empty = ws.cell(row=last_row, column=2, value="")
    c_sh_empty.fill = total_fill
    c_sh_empty.border = double_border

    gt = period_data["grand_totals"]
    for c_idx, val in [(3, gt["total"]), (4, gt["tax"]), (5, gt["net"])]:
        cell = ws.cell(row=last_row, column=c_idx, value=val)
        cell.number_format = "#,##0.00"
        cell.font = Font(name="Calibri", size=11, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = total_fill
        cell.border = double_border

    # 5. Set column widths
    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 16
    ws.column_dimensions["E"].width = 18

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def generate_monthly_excel(monthly_data: Dict[str, Any]) -> bytes:
    """Wrapper for backwards compatibility."""
    return generate_periodic_excel(monthly_data)


def generate_comparative_excel(comp_data: Dict[str, Any]) -> bytes:
    """
    Generates an executive Excel (.xlsx) file for Comparative Sales & Growth analysis.
    - Columns: م, المعرض, مبيعات الفترة السابقة, مبيعات الفترة الحالية, الفارق, نسبة النمو
    - Highlights positive growth in green and decline in red
    - Formatted currencies, percentages, borders, RTL view, and grand totals row.
    """
    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from io import BytesIO

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "تقرير المقارنة"
    ws.views.sheetView[0].rightToLeft = True

    # 1. Main Title
    ws.merge_cells("A1:F1")
    title_cell = ws["A1"]
    title_cell.value = comp_data["title"]
    title_cell.font = Font(name="Calibri", size=15, bold=True, color="1F4E79")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 34

    # 2. Table Headers at row 3
    prev_lbl = comp_data.get("previous_label", "الفترة السابقة")
    curr_lbl = comp_data.get("current_label", "الفترة الحالية")
    headers = [
        "م",
        "المعرض",
        f"مبيعات {prev_lbl} (الصافي)",
        f"مبيعات {curr_lbl} (الصافي)",
        "الفارق (ر.س)",
        "نسبة النمو (%)"
    ]
    header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    thin_border = Border(
        left=Side(style="thin", color="808080"),
        right=Side(style="thin", color="808080"),
        top=Side(style="thin", color="808080"),
        bottom=Side(style="thin", color="808080")
    )

    ws.row_dimensions[3].height = 28
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col_idx, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    # Fills for growth/decline
    green_font = Font(name="Calibri", size=10, bold=True, color="006100")
    red_font = Font(name="Calibri", size=10, bold=True, color="9C0006")
    green_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    red_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")

    # 3. Data rows
    for r_idx, r in enumerate(comp_data["rows"], start=4):
        ws.row_dimensions[r_idx].height = 20

        c_seq = ws.cell(row=r_idx, column=1, value=r["seq"])
        c_seq.alignment = Alignment(horizontal="center", vertical="center")
        c_seq.border = thin_border

        c_sh = ws.cell(row=r_idx, column=2, value=r["showroom"])
        c_sh.alignment = Alignment(horizontal="right", vertical="center")
        c_sh.border = thin_border

        c_prev = ws.cell(row=r_idx, column=3, value=r["previous_net"])
        c_prev.number_format = "#,##0.00"
        c_prev.alignment = Alignment(horizontal="center", vertical="center")
        c_prev.border = thin_border

        c_curr = ws.cell(row=r_idx, column=4, value=r["current_net"])
        c_curr.number_format = "#,##0.00"
        c_curr.font = Font(name="Calibri", size=10, bold=True)
        c_curr.alignment = Alignment(horizontal="center", vertical="center")
        c_curr.border = thin_border

        # Diff
        c_diff = ws.cell(row=r_idx, column=5, value=r["diff"])
        c_diff.number_format = "+#,##0.00;-#,##0.00;0.00"
        c_diff.alignment = Alignment(horizontal="center", vertical="center")
        c_diff.border = thin_border

        # Growth %
        c_pct = ws.cell(row=r_idx, column=6, value=r["growth_pct"] / 100.0)
        c_pct.number_format = "+0.0%;-0.0%;0.0%"
        c_pct.alignment = Alignment(horizontal="center", vertical="center")
        c_pct.border = thin_border

        if r["diff"] > 0:
            c_diff.font = green_font
            c_pct.font = green_font
            c_pct.fill = green_fill
        elif r["diff"] < 0:
            c_diff.font = red_font
            c_pct.font = red_font
            c_pct.fill = red_fill

    # 4. Total row
    last_row = 4 + len(comp_data["rows"])
    ws.row_dimensions[last_row].height = 25
    total_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    double_border = Border(
        left=Side(style="thin", color="808080"),
        right=Side(style="thin", color="808080"),
        top=Side(style="thin", color="808080"),
        bottom=Side(style="double", color="000000")
    )

    c_tot_label = ws.cell(row=last_row, column=1, value="المجموع")
    c_tot_label.font = Font(name="Calibri", size=11, bold=True)
    c_tot_label.alignment = Alignment(horizontal="center", vertical="center")
    c_tot_label.fill = total_fill
    c_tot_label.border = double_border

    c_sh_empty = ws.cell(row=last_row, column=2, value="")
    c_sh_empty.fill = total_fill
    c_sh_empty.border = double_border

    gt = comp_data["grand_totals"]
    c_gt_prev = ws.cell(row=last_row, column=3, value=gt["previous_net"])
    c_gt_prev.number_format = "#,##0.00"
    c_gt_prev.font = Font(name="Calibri", size=11, bold=True)
    c_gt_prev.alignment = Alignment(horizontal="center", vertical="center")
    c_gt_prev.fill = total_fill
    c_gt_prev.border = double_border

    c_gt_curr = ws.cell(row=last_row, column=4, value=gt["current_net"])
    c_gt_curr.number_format = "#,##0.00"
    c_gt_curr.font = Font(name="Calibri", size=11, bold=True)
    c_gt_curr.alignment = Alignment(horizontal="center", vertical="center")
    c_gt_curr.fill = total_fill
    c_gt_curr.border = double_border

    c_gt_diff = ws.cell(row=last_row, column=5, value=gt["diff"])
    c_gt_diff.number_format = "+#,##0.00;-#,##0.00;0.00"
    c_gt_diff.font = Font(name="Calibri", size=11, bold=True, color="006100" if gt["diff"] >= 0 else "9C0006")
    c_gt_diff.alignment = Alignment(horizontal="center", vertical="center")
    c_gt_diff.fill = total_fill
    c_gt_diff.border = double_border

    c_gt_pct = ws.cell(row=last_row, column=6, value=gt["growth_pct"] / 100.0)
    c_gt_pct.number_format = "+0.0%;-0.0%;0.0%"
    c_gt_pct.font = Font(name="Calibri", size=11, bold=True, color="006100" if gt["diff"] >= 0 else "9C0006")
    c_gt_pct.alignment = Alignment(horizontal="center", vertical="center")
    c_gt_pct.fill = green_fill if gt["diff"] >= 0 else red_fill
    c_gt_pct.border = double_border

    # 5. Column widths
    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 22
    ws.column_dimensions["E"].width = 18
    ws.column_dimensions["F"].width = 18

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()
