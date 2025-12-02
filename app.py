# -*- coding: utf-8 -*-
"""
Flask web port of your Tkinter Orders Manager.
Single-file app: run with `python app.py` then open http://127.0.0.1:5000

Key features kept:
- Passcode gate (1977)
- XLSX datastore with identical columns and logic
- PDF import (page-by-page) with the same extract rules
- Invoice PDF match -> auto mark Delivered
- Search, barcode (mark Returned), edit/delete, dedupe, move to Shipping with group product name
- Pending list with date filters
- Detailed stats (summary / by price / daily trend)

Folders auto-created under a per-user data dir (like the desktop version).
"""


from __future__ import annotations
import os
import re
import io
import sys
import traceback
from pathlib import Path
from datetime import datetime, date
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address
from werkzeug.utils import secure_filename
from flask import (
    Flask, render_template_string, request, redirect, url_for,
    session, flash, send_from_directory, abort
)

import requests  # تأكد pip install requests


import pandas as pd
import pdfplumber




try:
    import openpyxl  # noqa: F401
    from openpyxl.utils import get_column_letter
except Exception:  # pragma: no cover
    get_column_letter = None

# ----------------------------- CONFIG ---------------------------------
PASSCODE = "1977"
SECRET_KEY = os.environ.get("SECRET_KEY", "dev-secret-change-me")
# Telegram config (ضع القيم في متغيرات البيئة أو مباشرة للتجربة)
TELEGRAM_BOT_TOKEN = "8311293130:AAF5ALNUB9DZkJQ6KWoEYSiBedZxZneu6S8"
TELEGRAM_CHAT_ID = "-5043262753"  # ID الكروب     # مثال: '-1001234567890' أو ID الحساب


# ------------------------- SAFE PATH HELPERS ---------------------------
def is_frozen():
    return getattr(sys, "frozen", False)


def app_dir():
    if is_frozen():
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent


def user_data_dir():
    if os.name == "nt":
        base = os.environ.get("APPDATA") or str(Path.home() / "AppData" / "Roaming")
        p = Path(base) / "OrdersManagerWeb"
    else:
        p = Path.home() / ".local" / "share" / "OrdersManagerWeb"
    p.mkdir(parents=True, exist_ok=True)
    (p / "uploads").mkdir(exist_ok=True)
    return p


def resource_path(*parts):
    return str((app_dir() / Path(*parts)).resolve())

# ------------------------------ STORAGE -------------------------------
STATUS_READY = "قيد التجهيز"
STATUS_SHIPPING = "قيد التوصيل"
STATUS_DELIVERED = "تم التوصيل"
STATUS_RETURNED = "راجع"

BASE_COLUMNS = [
    "Product Name",
    "Page Name",
    "Transaction ID",
    "Time and Date",
    "Contact Numbers",
    "Address",
    "Order Price",
    "Status",
    "Return Reason",
    "Notes",
    "Client Orders Count",
]

EXCEL_FILE = str((user_data_dir() / "orders_data.xlsx").resolve())
ERROR_LOG = str((user_data_dir() / "error.log").resolve())
UPLOAD_DIR = str((user_data_dir() / "uploads").resolve())

# ------------------------------ UTILS ---------------------------------

def now_str():
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')

def send_telegram(msg: str):
    """
    إرسال رسالة بسيطة إلى تلغرام.
    يعتمد على TELEGRAM_BOT_TOKEN و TELEGRAM_CHAT_ID من متغيرات البيئة.
    """
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        return  # لو مو متهيئة، نطنش بصمت حتى ما يوقع البرنامج

    if requests is None:
        return  # لو مكتبة requests مو منصبة

    try:
        url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
        requests.post(url, data={
            "chat_id": TELEGRAM_CHAT_ID,
            "text": msg
        }, timeout=5)
    except Exception as e:
        # نسجل الخطأ في اللوج بدون ما نوقف التطبيق
        try:
            _fatal_box("Telegram send failed", e)
        except Exception:
            pass

def normalize_digits(s: str) -> str:
    if s is None:
        return ""
    trans = {
        ord('٠'): '0', ord('١'): '1', ord('٢'): '2', ord('٣'): '3', ord('٤'): '4',
        ord('٥'): '5', ord('٦'): '6', ord('٧'): '7', ord('٨'): '8', ord('٩'): '9',
        ord('۰'): '0', ord('۱'): '1', ord('۲'): '2', ord('۳'): '3', ord('۴'): '4',
        ord('۵'): '5', ord('۶'): '6', ord('۷'): '7', ord('۸'): '8', ord('۹'): '9',
        ord('\u066C'): ',',  # ARABIC THOUSANDS SEPARATOR -> ,
        ord('\u200f'): None, ord('\u200e'): None,  # RLM/LRM
    }
    return str(s).translate(trans)


def to_int(num_str: str):
    if not num_str:
        return None
    s = normalize_digits(num_str).replace(",", "").replace(" ", "")
    if not re.search(r'\d', s):
        return None
    try:
        return int(re.search(r'(\d+)', s).group(1))
    except Exception:
        return None


class DataStore:
    def __init__(self, path):
        self.path = path
        self.df = self._load_or_create()
        self._ensure_index()

    def _load_or_create(self):
        path = Path(self.path)
        if not path.exists():
            df = pd.DataFrame(columns=BASE_COLUMNS)
            try:
                with pd.ExcelWriter(self.path, engine="openpyxl") as writer:
                    df.to_excel(writer, index=False, sheet_name="Sheet1")
                    if get_column_letter is not None:
                        ws = writer.sheets["Sheet1"]
                        tid_idx = BASE_COLUMNS.index("Transaction ID") + 1
                        for cell in ws[get_column_letter(tid_idx)]:
                            cell.number_format = "@"
            except Exception:
                df.to_excel(self.path, index=False)
            return df

        try:
            df = pd.read_excel(self.path, dtype=str)
        except Exception:
            df = pd.read_excel(self.path)
            if "Transaction ID" in df.columns:
                df["Transaction ID"] = df["Transaction ID"].astype(str)

        if "Product Name" not in df.columns:
            if "Title" in df.columns:
                df.rename(columns={"Title": "Product Name"}, inplace=True)
            else:
                df["Product Name"] = pd.NA
        if "Order Price" not in df.columns:
            df["Order Price"] = pd.NA
        for old_col in ["Delivery Type", "Delivery Cost", "Pieces Count", "Page Number"]:
            if old_col in df.columns:
                df.drop(columns=[old_col], inplace=True)
        for c in BASE_COLUMNS:
            if c not in df.columns:
                df[c] = pd.NA
        df["Transaction ID"] = df["Transaction ID"].astype(str).str.strip()
        df["Order Price"] = pd.to_numeric(df["Order Price"], errors="coerce")
        df["Status"] = df["Status"].fillna(STATUS_READY)
        # ensure new columns ordering
        df = df[BASE_COLUMNS]
        return df

    def _ensure_index(self):
        if "Transaction ID" not in self.df.columns:
            self.df["Transaction ID"] = ""
        try:
            self.df.set_index("Transaction ID", drop=False, inplace=True)
        except Exception:
            pass

    def save(self):
        to_save = self.df.reset_index(drop=True).copy()
        to_save["Transaction ID"] = to_save["Transaction ID"].astype(str)
        try:
            with pd.ExcelWriter(self.path, engine="openpyxl") as writer:
                to_save.to_excel(writer, index=False, sheet_name="Sheet1")
                if get_column_letter is not None:
                    ws = writer.sheets["Sheet1"]
                    tid_idx = BASE_COLUMNS.index("Transaction ID") + 1
                    for cell in ws[get_column_letter(tid_idx)]:
                        cell.number_format = "@"
        except Exception:
            to_save.to_excel(self.path, index=False)

    def exists(self, txn):
        return str(txn).strip() in self.df.index

    def get_row(self, txn):
        txn = str(txn).strip()
        if self.exists(txn):
            return self.df.loc[txn]
        return None

    def upsert_row(self, row_dict: dict):
        txn = str(row_dict.get("Transaction ID", "")).strip()
        if not txn or not re.fullmatch(r'\d{6,}', txn):
            return False, "Transaction ID غير صالح (أرقام فقط وبحد أدنى 6 خانات)."
        row_dict = row_dict.copy()
        if not row_dict.get("Status"):
            row_dict["Status"] = STATUS_READY
        for c in BASE_COLUMNS:
            if c not in row_dict:
                row_dict[c] = pd.NA
        if self.exists(txn):
            for k, v in row_dict.items():
                self.df.at[txn, k] = v
            return True, "تم التحديث"
        else:
            new_df = pd.DataFrame([row_dict], columns=BASE_COLUMNS)
            new_df["Transaction ID"] = new_df["Transaction ID"].astype(str).str.strip()
            new_df.set_index("Transaction ID", drop=False, inplace=True)
            self.df = pd.concat([self.df, new_df], axis=0, ignore_index=False)
            return True, "تمت الإضافة"

    def update_status(self, txn, new_status, return_reason=None):
        txn = str(txn).strip()
        if not self.exists(txn):
            return False, "الشحنة غير موجودة"
        old_status = self.df.at[txn, "Status"] if "Status" in self.df.columns else None
        self.df.at[txn, "Status"] = new_status
        if return_reason is not None:
            self.df.at[txn, "Return Reason"] = return_reason
        ret = {"msg": "تم تحديث الحالة", "old": old_status, "new": new_status, "row": self.df.loc[txn] }
        # inventory hook
        try:
            adjust_inventory_on_transition(ret['row'], old_status, new_status)
        except Exception:
            pass
        return True, ret

    def drop_by_txn(self, txn):
        txn = str(txn).strip()
        if not self.exists(txn):
            return 0
        self.df = self.df.drop(index=txn)
        return 1

    def drop_duplicates_keep_last(self):
        before = len(self.df)
        self.df = (
            self.df.reset_index(drop=True)
                   .drop_duplicates(subset=["Transaction ID"], keep="last")
        )
        self._ensure_index()
        after = len(self.df)
        return before - after

    def stats_global(self, df=None):
        d = self.df if df is None else df
        total_orders = len(d)
        total_amount = pd.to_numeric(d["Order Price"], errors="coerce").sum()
        delivered = (d["Status"] == STATUS_DELIVERED).sum()
        returned = (d["Status"] == STATUS_RETURNED).sum()
        shipping = (d["Status"] == STATUS_SHIPPING).sum()
        ready = (d["Status"] == STATUS_READY).sum()
        pct = lambda x: (x / total_orders * 100) if total_orders else 0.0
        return {
            "العدد الكلي للطلبات": total_orders,
            "المجموع المالي (Order Price)": float(total_amount or 0),
            f"عدد {STATUS_DELIVERED}": delivered,
            f"عدد {STATUS_RETURNED}": returned,
            f"عدد {STATUS_SHIPPING}": shipping,
            f"عدد {STATUS_READY}": ready,
            f"نسبة {STATUS_DELIVERED} %": round(pct(delivered), 2),
            f"نسبة {STATUS_RETURNED} %": round(pct(returned), 2),
            f"نسبة {STATUS_SHIPPING} %": round(pct(shipping), 2),
            f"نسبة {STATUS_READY} %": round(pct(ready), 2),
        }

    def stats_by_product_price(self, df=None):
        d = self.df if df is None else df
        d = d.copy()
        d["Order Price"] = pd.to_numeric(d["Order Price"], errors="coerce")
        cols = [
            "السعر", "عدد الطلبات",
            STATUS_DELIVERED, STATUS_RETURNED, STATUS_SHIPPING, STATUS_READY,
            "المبلغ المُسلَّم", "نسبة الراجع %"
        ]
        if d.empty or d["Order Price"].isna().all():
            return pd.DataFrame(columns=cols)
        rows = []
        for price, g in d.groupby("Order Price", dropna=False):
            total = len(g)
            delivered = (g["Status"] == STATUS_DELIVERED).sum()
            returned = (g["Status"] == STATUS_RETURNED).sum()
            shipping = (g["Status"] == STATUS_SHIPPING).sum()
            ready = (g["Status"] == STATUS_READY).sum()
            delivered_amount = pd.to_numeric(
                g.loc[g["Status"] == STATUS_DELIVERED, "Order Price"], errors="coerce"
            ).sum()
            return_rate = (returned / total * 100) if total else 0.0
            rows.append({
                "السعر": price,
                "عدد الطلبات": total,
                STATUS_DELIVERED: delivered,
                STATUS_RETURNED: returned,
                STATUS_SHIPPING: shipping,
                STATUS_READY: ready,
                "المبلغ المُسلَّم": float(delivered_amount or 0),
                "نسبة الراجع %": round(return_rate, 2),
            })
        out_df = pd.DataFrame(rows, columns=cols)
        if not out_df.empty:
            out_df = out_df.sort_values(
                by=["المبلغ المُسلَّم", "عدد الطلبات"],
                ascending=[False, False],
                na_position="last"
            )
        return out_df

    def daily_trend(self, df=None):
        d = self.df if df is None else df
        d = d.copy()
        d["Time and Date"] = pd.to_datetime(d["Time and Date"], errors="coerce")
        d = d.dropna(subset=["Time and Date"])
        d["Date"] = d["Time and Date"].dt.date
        daily = d.groupby("Date").size().reset_index(name="Order Count").sort_values("Date")
        daily["Trend"] = daily["Order Count"].diff().apply(
            lambda x: "ارتفاع" if x and x > 0 else ("انخفاض" if x and x < 0 else "ثابت")
        )
        return daily
    
_data_root = Path(EXCEL_FILE).parent
# ------------------------------ APP -----------------------------------
app = Flask(__name__)
app.secret_key = SECRET_KEY
store = DataStore(EXCEL_FILE)

# --------------------------- TEMPLATES --------------------------------
limiter = Limiter(
    key_func=get_remote_address,
    app=app,
    default_limits=["200 per hour"]
)

INVENTORY_HTML = r"""
{% extends 'base.html' %}
{% block content %}
<div class="row g-3">
  <div class="col-xl-8">
    <div class="card p-3">
      <div class="d-flex justify-content-between align-items-center mb-2">
        <h5 class="mb-0">المخزن</h5>
        <span class="badge bg-secondary">عدد السجلات: {{ rows|length }}</span>
      </div>
      <div class="table-responsive">
        <table class="table table-striped align-middle">
          <thead>
            <tr>
              <th>الكود</th>
              <th>الاسم</th>
              <th>النوع</th>
              <th>الكمية</th>
              <th>أمتار القماش</th>
              <th>متر/قطعة</th>
              <th>تكلفة خياطة</th>
              <th>تكاليف أخرى</th>
              <th>سعر البيع</th>
              <th>إجراءات</th>
            </tr>
          </thead>
          <tbody>
            {% for r in rows %}
            <tr>
              <td>{{ r['Product Code'] }}</td>
              <td>{{ r['Product Name'] }}</td>
              <td>{{ r['Type'] }}</td>
              <td>{{ r['Quantity'] }}</td>
              <td>{{ r['Fabric Meters'] }}</td>
              <td>{{ r['Meters per Unit'] }}</td>
              <td>{{ r['Sewing Cost'] }}</td>
              <td>{{ r['Other Costs'] }}</td>
              <td>{{ r['Sale Price'] }}</td>
              <td class="text-nowrap">
                <button class="btn btn-sm btn-success" data-bs-toggle="modal" data-bs-target="#addQtyModal" data-name="{{ r['Product Name'] }}">+ إضافة كمية</button>
                <button class="btn btn-sm btn-outline-danger ms-1" data-bs-toggle="modal" data-bs-target="#takeQtyModal" data-name="{{ r['Product Name'] }}">- سحب كمية</button>
              </td>
            </tr>
            {% endfor %}
          </tbody>
        </table>
      </div>
    </div>
  </div>
  <div class="col-xl-4">
    <div class="card p-3">
      <h6 class="mb-3">إضافة صنف جديد</h6>
      <form method="post" action="{{ url_for('inventory_add') }}" class="row g-2">
        <div class="col-12">
          <label class="form-label">اسم المنتج</label>
          <input required name="name" class="form-control" placeholder="مثال: عباءة موديل 123" autofocus autocomplete="off">
        </div>
        <div class="col-12">
          <label class="form-label">نوع البضاعة</label>
          <select name="type" class="form-select">
            <option value="">—</option>
            <option>ملابس أطفال</option>
            <option>نساء</option>
            <option>عباءة</option>
            <option>سوت</option>
          </select>
        </div>
        <div class="col-6"><label class="form-label">الكمية</label><input name="qty" type="number" class="form-control" value="0" inputmode="numeric" pattern="[0-9]*"></div>
        <div class="col-6"><label class="form-label">أمتار القماش</label><input name="fabric" type="number" step="0.01" class="form-control" value="0" inputmode="decimal"></div>
        <div class="col-6"><label class="form-label">متر/قطعة</label><input name="mpu" type="number" step="0.01" class="form-control" value="0" inputmode="decimal"></div>
        <div class="col-6"><label class="form-label">تكلفة الخياطة</label><input name="sew" type="number" step="0.01" class="form-control" value="0" inputmode="decimal"></div>
        <div class="col-6"><label class="form-label">تكاليف أخرى</label><input name="other" type="number" step="0.01" class="form-control" value="0" inputmode="decimal"></div>
        <div class="col-6"><label class="form-label">سعر البيع</label><input name="price" type="number" step="0.01" class="form-control" value="0" inputmode="decimal"></div>
        <div class="col-12"><button class="btn btn-dark w-100">إضافة</button></div>
      </form>
    </div>
  </div>
</div>

<!-- Modal: إضافة كمية -->
<div class="modal fade" id="addQtyModal" tabindex="-1">
  <div class="modal-dialog">
    <form method="post" action="{{ url_for('inventory_adjust_bulk') }}" class="modal-content">
      <div class="modal-header"><h6 class="modal-title">إضافة كمية للمخزن</h6><button type="button" class="btn-close" data-bs-dismiss="modal"></button></div>
      <div class="modal-body">
        <input type="hidden" name="name" id="addQtyName">
        <div class="mb-2">
          <label class="form-label">الكمية التي ستُضاف</label>
          <input required name="qty" type="number" class="form-control" value="1" min="1" inputmode="numeric" pattern="[0-9]*" autofocus>
        </div>
      </div>
      <div class="modal-footer">
        <button class="btn btn-success">إضافة</button>
      </div>
    </form>
  </div>
</div>

<!-- Modal: سحب كمية -->
<div class="modal fade" id="takeQtyModal" tabindex="-1">
  <div class="modal-dialog">
    <form method="post" action="{{ url_for('inventory_adjust_bulk') }}" class="modal-content">
      <div class="modal-header"><h6 class="modal-title">سحب كمية من المخزن</h6><button type="button" class="btn-close" data-bs-dismiss="modal"></button></div>
      <div class="modal-body">
        <input type="hidden" name="name" id="takeQtyName">
        <div class="mb-2">
          <label class="form-label">الكمية التي ستُسحب</label>
          <input required name="qty" type="number" class="form-control" value="-1" step="1" inputmode="numeric" pattern="-?[0-9]*" autofocus>
          <div class="form-text">استخدم قيمة سالبة للسحب (مثال: -5)</div>
        </div>
      </div>
      <div class="modal-footer">
        <button class="btn btn-danger">سحب</button>
      </div>
    </form>
  </div>
</div>

<!-- Feedback Modal (after redirect) -->
<div class="modal fade" id="feedbackModal" tabindex="-1">
  <div class="modal-dialog">
    <div class="modal-content">
      <div class="modal-header"><h6 class="modal-title">تحديث المخزن</h6><button type="button" class="btn-close" data-bs-dismiss="modal"></button></div>
      <div class="modal-body">
        {% if added %}
          تم إضافة <b>{{ added }}</b> قطعة إلى المنتج <b>{{ name }}</b>.
        {% elif taken %}
          تم سحب <b>{{ taken }}</b> قطعة من المنتج <b>{{ name }}</b>.
        {% endif %}
      </div>
      <div class="modal-footer"><button class="btn btn-secondary" data-bs-dismiss="modal">إغلاق</button></div>
    </div>
  </div>
</div>

<script>
  const addQtyModal = document.getElementById('addQtyModal');
  addQtyModal?.addEventListener('show.bs.modal', event => {
    const btn = event.relatedTarget; const name = btn.getAttribute('data-name');
    document.getElementById('addQtyName').value = name;
  });
  const takeQtyModal = document.getElementById('takeQtyModal');
  takeQtyModal?.addEventListener('show.bs.modal', event => {
    const btn = event.relatedTarget; const name = btn.getAttribute('data-name');
    document.getElementById('takeQtyName').value = name;
  });
  // auto show feedback if present
  {% if added or taken %}
  const fb = new bootstrap.Modal(document.getElementById('feedbackModal'));
  fb.show();
  {% endif %}
</script>
{% endblock %}
"""

# (Bootstrap from CDN; RTL-friendly)


# ----------------------------- ISSUES TEMPLATE --------------------------
ISSUES_HTML = r"""
{% extends 'base.html' %}
{% block content %}
<div class="card p-3">
  <div class="d-flex justify-content-between align-items-center mb-2">
    <h6 class="mb-0">المشاكل</h6>
    <span class="badge bg-secondary">عدد السجلات: {{ rows|length }}</span>
  </div>
  <form method="post" action="{{ url_for('issues_add') }}" enctype="multipart/form-data" class="row g-2 mb-3">
    <div class="col-md-4"><input required name="title" class="form-control" placeholder="عنوان المشكلة" autocomplete="off"></div>
    <div class="col-md-5"><input name="desc" class="form-control" placeholder="وصف مختصر"></div>
    <div class="col-md-2"><input type="file" name="image" accept="image/*" class="form-control"></div>
    <div class="col-md-1"><button class="btn btn-dark w-100">رفع</button></div>
  </form>
  <div class="table-responsive">
    <table class="table table-striped align-middle">
      <thead><tr><th>#</th><th>العنوان</th><th>الوصف</th><th>الصورة</th><th>الحالة</th><th>الحلّ</th><th>أُنشئت</th><th>إجراءات</th></tr></thead>
      <tbody>
        {% for r in rows %}
        <tr>
          <td>{{ r['ID'] }}</td>
          <td>{{ r['Title'] }}</td>
          <td>{{ r['Description'] }}</td>
          <td>{% if r['ImagePath'] %}<img src="/static-proxy?f={{ r['ImagePath'] }}" style="height:56px">{% endif %}</td>
          <td>{{ r['Status'] }}</td>
          <td>{{ r['Solver'] }}</td>
          <td>{{ r['CreatedAt'] }}</td>
          <td class="text-nowrap">
            {% if r['Status']!='Solved' %}
            <form method="post" action="{{ url_for('issues_solve') }}" class="d-inline">
              <input type="hidden" name="id" value="{{ r['ID'] }}">
              <input name="solver" class="form-control form-control-sm d-inline-block" style="width:140px" placeholder="اسم الحلّال" required>
              <button class="btn btn-sm btn-success ms-1">تم الحل</button>
            </form>
            {% endif %}
            <a class="btn btn-sm btn-outline-danger ms-1" href="{{ url_for('issues_delete', iid=r['ID']) }}" onclick="return confirm('حذف المشكلة؟');">حذف</a>
          </td>
        </tr>
        {% endfor %}
      </tbody>
    </table>
  </div>
</div>
{% endblock %}
"""



SEAMSTRESS_HTML = r"""
{% extends 'base.html' %}
{% block content %}
<div class="row g-3">
  <div class="col-xl-7">
    <div class="card p-3">
      <div class="d-flex justify-content-between align-items-center mb-2">
        <h6 class="mb-0">الخياطات</h6>
        <span class="badge bg-secondary">عدد السجلات: {{ seamstresses|length }}</span>
      </div>
      <div class="table-responsive">
        <table class="table table-striped align-middle">
          <thead><tr><th>#</th><th>الاسم</th><th>الهاتف</th><th>ملاحظات</th><th>فعّالة</th><th>إجراءات</th></tr></thead>
          <tbody>
            {% for r in seamstresses %}
            <tr>
              <td>{{ r['ID'] }}</td>
              <td>{{ r['Name'] }}</td>
              <td>{{ r['Phone'] }}</td>
              <td>{{ r['Notes'] }}</td>
              <td>{{ 'نعم' if r['Active'] else 'لا' }}</td>
              <td class="text-nowrap">
                <button class="btn btn-sm btn-outline-primary" data-bs-toggle="modal" data-bs-target="#editSeam" data-id="{{r['ID']}}" data-name="{{r['Name']}}" data-phone="{{r['Phone']}}" data-notes="{{r['Notes']}}" data-active="{{r['Active']}}">تعديل</button>
                <a class="btn btn-sm btn-outline-danger" href="{{ url_for('seam_delete', sid=r['ID']) }}" onclick="return confirm('حذف {{r['Name']}}؟');">حذف</a>
              </td>
            </tr>
            {% endfor %}
          </tbody>
        </table>
      </div>
    </div>
      <form method="get" class="row g-2 mb-2">
        <div class="col-md-3">
          <label class="form-label">من تاريخ</label>
          <input type="date" name="from" class="form-control" value="{{ dfrom or '' }}">
        </div>
        <div class="col-md-3">
          <label class="form-label">إلى تاريخ</label>
          <input type="date" name="to" class="form-control" value="{{ dto or '' }}">
        </div>
        <div class="col-md-3">
          <label class="form-label">الخياطة</label>
          <select name="sid" class="form-select">
            <option value="">الكل</option>
            {% for r in seamstresses %}
              <option value="{{ r['ID'] }}" {{ 'selected' if sel_sid and sel_sid|int == r['ID'] else '' }}>
                {{ r['Name'] }}
              </option>
            {% endfor %}
          </select>
        </div>
        <div class="col-md-3">
          <label class="form-label">الحالة</label>
          <select name="paid" class="form-select">
            <option value="">الكل</option>
            <option value="paid" {{ 'selected' if sel_paid=='paid' else '' }}>مدفوع</option>
            <option value="unpaid" {{ 'selected' if sel_paid=='unpaid' else '' }}>غير مدفوع</option>
          </select>
        </div>
        <div class="col-12 text-end">
          <button class="btn btn-secondary btn-sm mt-1">تطبيق</button>
          <a href="{{ url_for('seam_home') }}" class="btn btn-outline-secondary btn-sm mt-1">إلغاء</a>
        </div>
      </form>

    <div class="card p-3 mt-3">
      <div class="d-flex justify-content-between align-items-center mb-2">
        <h6 class="mb-0">سجل الإنجاز اليومي</h6>
        <span class="badge bg-secondary">عدد السجلات: {{ logs|length }}</span>
      </div>
      <div class="table-responsive">
        <table class="table table-striped align-middle">
          <thead><tr><th>#</th><th>التاريخ</th><th>الخياطة</th><th>الموديل</th><th>القطع</th><th>سعر/قطعة</th><th>الإجمالي</th><th>مدفوع</th><th>إجراءات</th></tr></thead>
          <tbody>
            {% for r in logs %}
            <tr>
              <td>{{ r['LogID'] }}</td>
              <td>{{ r['Date'] }}</td>
              <td>{{ seam_name_map.get(r['SeamstressID'], r['SeamstressID']) }}</td>
              <td>{{ r['Model'] }}</td>
              <td>{{ r['Pieces'] }}</td>
              <td>{{ r['UnitCost'] }}</td>
              <td>{{ r['Total'] }}</td>
              <td>{{ 'نعم' if r['Paid'] else 'لا' }}</td>
              <td>
                {% if not r['Paid'] %}
                <a class="btn btn-sm btn-success" href="{{ url_for('sew_mark_paid', log_id=r['LogID']) }}">تصفية</a>
                {% else %}
                <a class="btn btn-sm btn-outline-secondary" href="{{ url_for('sew_mark_unpaid', log_id=r['LogID']) }}">إلغاء التصفية</a>
                {% endif %}
              </td>
            </tr>
            {% endfor %}
          </tbody>
        </table>
      </div>
    </div>
  </div>

  <div class="col-xl-5">
    <div class="card p-3">
      <h6 class="mb-3">إضافة خياطة</h6>
      <form method="post" action="{{ url_for('seam_add') }}" class="row g-2">
        <div class="col-6"><label class="form-label">الاسم</label><input required name="name" class="form-control" autocomplete="off"></div>
        <div class="col-6"><label class="form-label">الهاتف</label><input name="phone" class="form-control" inputmode="numeric" pattern="[0-9]*"></div>
        <div class="col-12"><label class="form-label">ملاحظات</label><input name="notes" class="form-control"></div>
        <div class="col-12"><button class="btn btn-dark w-100">إضافة</button></div>
      </form>
    </div>

    <div class="card p-3 mt-3">
      <h6 class="mb-3">تسجيل إنجاز اليوم</h6>
      <form method="post" action="{{ url_for('sew_add_log') }}" class="row g-2">
        <div class="col-6">
          <label class="form-label">الخياطة</label>
          <select name="sid" class="form-select" required>
            <option value="">—</option>
            {% for r in seamstresses %}
              <option value="{{ r['ID'] }}">{{ r['Name'] }}</option>
            {% endfor %}
          </select>
        </div>
        <div class="col-6"><label class="form-label">اسم الموديل</label><input required name="model" class="form-control" autocomplete="off"></div>
        <div class="col-6"><label class="form-label">عدد القطع</label><input required type="number" name="pieces" class="form-control" min="1" value="1" inputmode="numeric" pattern="[0-9]*"></div>
        <div class="col-6"><label class="form-label">سعر الخياطة/قطعة</label><input required type="number" step="0.01" name="unit" class="form-control" value="0" inputmode="decimal"></div>
        <div class="col-12"><button class="btn btn-success w-100">تسجيل & زيادة المخزون</button></div>
      </form>
    </div>
  </div>
</div>

<!-- Modal تعديل خياطة -->
<div class="modal fade" id="editSeam" tabindex="-1">
  <div class="modal-dialog">
    <form method="post" action="{{ url_for('seam_edit') }}" class="modal-content">
      <div class="modal-header"><h6 class="modal-title">تعديل خياطة</h6><button type="button" class="btn-close" data-bs-dismiss="modal"></button></div>
      <div class="modal-body">
        <input type="hidden" name="id" id="seamID">
        <div class="mb-2"><label class="form-label">الاسم</label><input name="name" id="seamName" class="form-control"></div>
        <div class="mb-2"><label class="form-label">الهاتف</label><input name="phone" id="seamPhone" class="form-control" inputmode="numeric" pattern="[0-9]*"></div>
        <div class="mb-2"><label class="form-label">ملاحظات</label><input name="notes" id="seamNotes" class="form-control"></div>
        <div class="form-check"><input class="form-check-input" type="checkbox" name="active" id="seamActive"><label class="form-check-label" for="seamActive">فعّالة</label></div>
      </div>
      <div class="modal-footer"><button class="btn btn-primary">حفظ</button></div>
    </form>
  </div>
</div>

<script>
  const editSeam = document.getElementById('editSeam');
  editSeam?.addEventListener('show.bs.modal', e => {
    const b = e.relatedTarget;
    document.getElementById('seamID').value = b.getAttribute('data-id');
    document.getElementById('seamName').value = b.getAttribute('data-name');
    document.getElementById('seamPhone').value = b.getAttribute('data-phone');
    document.getElementById('seamNotes').value = b.getAttribute('data-notes');
    document.getElementById('seamActive').checked = (b.getAttribute('data-active') === 'True');
  });
</script>
{% endblock %}
"""

# ---------------------------- CUTTING TEMPLATE --------------------------
CUTTING_HTML = r"""
{% extends 'base.html' %}
{% block content %}
<style>
  /* تلوين صفوف جدول الفصال حسب الحالة */
  .cutting-table tbody tr.row-working  { background-color: #fff3cd !important; }  /* قيد العمل - أصفر فاتح */
  .cutting-table tbody tr.row-done     { background-color: #d4edda !important; }  /* مكتمل - أخضر فاتح */
  .cutting-table tbody tr.row-rejected { background-color: #f8d7da !important; }  /* مرفوض - أحمر فاتح */
  .cutting-table tbody tr.row-pending  { background-color: #e2e3e5 !important; }  /* قيد الانتظار - رمادي فاتح */
</style>

<div class="row g-3">
  <div class="col-xl-5">
    <div class="card p-3">
      <h6 class="mb-3">إنشاء فصل جديد</h6>
      <form method="post" action="{{ url_for('cutting_add') }}" enctype="multipart/form-data" class="row g-2">
        <div class="col-12">
          <label class="form-label">اسم الموديل</label>
          <input required name="model" class="form-control" autocomplete="off">
        </div>
        <div class="col-6">
          <label class="form-label">موعد الفصال</label>
          <input required type="date" name="due" class="form-control">
        </div>
        <div class="col-6">
          <label class="form-label">عدد القطع المطلوبة</label>
          <input required type="number" name="qty" class="form-control" min="1" value="1" inputmode="numeric" pattern="[0-9]*">
        </div>
        <div class="col-12">
          <label class="form-label">ملاحظات</label>
          <input name="notes" class="form-control">
        </div>
        <div class="col-12">
          <label class="form-label">صورة الموديل</label>
          <input type="file" name="image" accept="image/*" class="form-control">
        </div>
        <div class="col-12">
          <button class="btn btn-dark w-100">إنشاء</button>
        </div>
      </form>
    </div>
  </div>

  <div class="col-xl-7">
    <div class="card p-3">
      <div class="d-flex justify-content-between align-items-center mb-2">
        <h6 class="mb-0">طلبات الفصال</h6>
        <span class="badge bg-secondary">عدد السجلات: {{ rows|length }}</span>
      </div>

      <div class="table-responsive">
        <table class="table table-striped align-middle cutting-table">
          <thead>
            <tr>
              <th>#</th>
              <th>الموديل</th>
              <th>الصورة</th>
              <th>الموعد</th>
              <th>المطلوب</th>
              <th>الحالة</th>
              <th>ملاحظات</th>
              <th>سبب الرفض</th>
              <th>إجراءات</th>
            </tr>
          </thead>
          <tbody>
            {% for r in rows %}
            {% set st = r['Status'] %}
            <tr
              class="
                {% if st == 'قيد العمل' %}
                  row-working
                {% elif st == 'مكتمل' %}
                  row-done
                {% elif st == 'مرفوض' %}
                  row-rejected
                {% elif st == 'قيد الانتظار' %}
                  row-pending
                {% endif %}
              "
            >
              <td>{{ r['ID'] }}</td>
              <td>{{ r['Model'] }}</td>
              <td>
                {% if r['ImagePath'] %}
                  <img src="/static-proxy?f={{ r['ImagePath'] }}" style="height:56px">
                {% endif %}
              </td>
              <td>{{ r['DueDate'] }}</td>
              <td>{{ r['RequiredQty'] }}</td>
              <td>
                {% if st == 'قيد العمل' %}
                  <span class="badge bg-warning text-dark">{{ st }}</span>
                {% elif st == 'مكتمل' %}
                  <span class="badge bg-success">{{ st }}</span>
                {% elif st == 'مرفوض' %}
                  <span class="badge bg-danger">{{ st }}</span>
                {% elif st == 'قيد الانتظار' %}
                  <span class="badge bg-secondary">{{ st }}</span>
                {% else %}
                  <span class="badge bg-light text-dark">{{ st }}</span>
                {% endif %}
              </td>
              <td>{{ r['Notes'] }}</td>
              <td>{{ r['RejectionReason'] }}</td>
              <td class="text-nowrap">
                <a class="btn btn-sm btn-outline-secondary"
                   href="{{ url_for('cutting_status', cid=r['ID'], s='قيد الانتظار') }}">
                   انتظار
                </a>
                <a class="btn btn-sm btn-primary"
                   href="{{ url_for('cutting_status', cid=r['ID'], s='قيد العمل') }}">
                   عمل
                </a>
                <a class="btn btn-sm btn-success"
                   href="{{ url_for('cutting_status', cid=r['ID'], s='مكتمل') }}">
                   مكتمل
                </a>
                <button class="btn btn-sm btn-outline-danger"
                        data-bs-toggle="modal"
                        data-bs-target="#rejectModal"
                        data-id="{{ r['ID'] }}">
                  رفض
                </button>
                <a class="btn btn-sm btn-outline-danger"
                   href="{{ url_for('cutting_delete', cid=r['ID']) }}"
                   onclick="return confirm('حذف الفصال؟');">
                  حذف
                </a>
              </td>
            </tr>
            {% endfor %}
          </tbody>
        </table>
      </div>
    </div>
  </div>
</div>

<div class="modal fade" id="rejectModal" tabindex="-1">
  <div class="modal-dialog">
    <form method="post" action="{{ url_for('cutting_reject') }}" class="modal-content">
      <div class="modal-header">
        <h6 class="modal-title">رفض طلب فصال</h6>
        <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
      </div>
      <div class="modal-body">
        <input type="hidden" name="id" id="rejID">
        <label class="form-label">سبب الرفض</label>
        <textarea required name="reason" class="form-control"></textarea>
      </div>
      <div class="modal-footer">
        <button class="btn btn-danger">رفض</button>
      </div>
    </form>
  </div>
</div>

<script>
  const rej = document.getElementById('rejectModal');
  rej?.addEventListener('show.bs.modal', e => {
    document.getElementById('rejID').value = e.relatedTarget.getAttribute('data-id');
  });
</script>

{% endblock %}
"""



BASE_HTML = r"""
<!doctype html>
<html lang="ar" dir="rtl">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{{ title or 'نظام إدارة الطلبات (ويب)' }}</title>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
  <style>
    body{background:#f8f9fb}
    .table thead th{white-space:nowrap}
    .card{box-shadow:0 2px 10px rgba(0,0,0,.06)}
    .form-control, .btn{border-radius:0.75rem}
  </style>
</head>
<body>
<nav class="navbar navbar-expand-lg bg-white border-bottom">
  <div class="container-fluid">
    <a class="navbar-brand fw-bold" href="{{ url_for('home') }}">🗂️ نظام الطلبات</a>
    <div class="d-flex">
      {% if session.get('auth') %}
      <a class="btn btn-sm btn-outline-secondary me-2" href="{{ url_for('download_excel') }}">تنزيل ملف Excel</a>
      <a class="btn btn-sm btn-danger" href="{{ url_for('logout') }}">تسجيل خروج</a>
      {% endif %}
    </div>
  </div>
</nav>

<div class="container py-4">
  {% with messages = get_flashed_messages(with_categories=true) %}
    {% if messages %}
      {% for cat, msg in messages %}
        <div class="alert alert-{{ 'success' if cat=='ok' else ('danger' if cat=='err' else 'info') }}">{{ msg }}</div>
      {% endfor %}
    {% endif %}
  {% endwith %}
  {% block content %}{% endblock %}
</div>
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js"></script>
</body>
</html>
"""

LOGIN_HTML = r"""
{% extends 'base.html' %}
{% block content %}
<div class="row justify-content-center">
  <div class="col-md-5">
    <div class="card p-4">
      <h5 class="mb-3">أدخل رمز الدخول</h5>
      <form method="post">
        <div class="mb-3">
          <input required name="code" type="password" class="form-control form-control-lg" placeholder="••••">
        </div>
        <button class="btn btn-primary w-100">دخول</button>
      </form>
    </div>
  </div>
</div>
{% endblock %}
"""

HOME_HTML = r"""
{% extends 'base.html' %}
{% block content %}
<div class="row g-3">
  <div class="col-xl-8">
    <div class="card p-3">
      <form class="row g-2 align-items-end" method="get" action="{{ url_for('home') }}">
        <div class="col-md-3">
          <label class="form-label">بحث</label>
          <input name="q" value="{{ q or '' }}" class="form-control" placeholder="كلمة مفتاحية" autofocus>
        </div>
        <div class="col-md-3">
          <label class="form-label">اسم المنتج</label>
          <select name="product" class="form-select">
            <option value="">الكل</option>
            {% for p in all_products %}
              <option value="{{p}}" {{ 'selected' if sel_product==p else '' }}>{{p}}</option>
            {% endfor %}
          </select>
        </div>
        <div class="col-md-3">
          <label class="form-label">اسم البيج</label>
          <select name="page" class="form-select">
            <option value="">الكل</option>
            {% for p in all_pages %}
              <option value="{{p}}" {{ 'selected' if sel_page==p else '' }}>{{p}}</option>
            {% endfor %}
          </select>
        </div>
        <div class="col-md-3">
          <label class="form-label">من تاريخ</label>
          <input name="from" type="date" class="form-control" value="{{ dfrom or '' }}">
        </div>
        <div class="col-md-3">
          <label class="form-label">إلى تاريخ</label>
          <input name="to" type="date" class="form-control" value="{{ dto or '' }}">
        </div>
        <div class="col-md-3 text-end align-self-end">
          <button class="btn btn-secondary mt-2">تطبيق</button>
          <a href="{{ url_for('home') }}" class="btn btn-outline-secondary mt-2">إلغاء</a>
        </div>
      </form>
    </div>

    <div class="card p-3 mt-3">
      <div class="d-flex justify-content-between align-items-center mb-2">
        <h6 class="mb-0">الطلبات</h6>
        <span class="badge bg-secondary">عدد السجلات: {{ rows|length }}</span>
      </div>
      <div class="table-responsive">
        <table class="table table-striped align-middle">
          <thead><tr>
            {% for c in columns %}<th>{{ c }}</th>{% endfor %}
            <th>إجراءات</th>
          </tr></thead>
          <tbody>
            {% for r in rows %}
            <tr>
              {% for c in columns %}<td>{{ r.get(c,'') }}</td>{% endfor %}
              <td class="text-nowrap">
                <a class="btn btn-sm btn-outline-primary" href="{{ url_for('edit', txn=r['Transaction ID']) }}">تعديل</a>
                <a class="btn btn-sm btn-outline-danger" href="{{ url_for('delete', txn=r['Transaction ID']) }}" onclick="return confirm('تأكيد حذف {{ r['Transaction ID'] }}؟')">حذف</a>
              </td>
            </tr>
            {% endfor %}
          </tbody>
        </table>
      </div>
    </div>
  </div>

  <div class="col-xl-4">
    <div class="card p-3">
      <h6 class="mb-3">قارئ باركود (تحديث إلى راجع)</h6>
      <form method="post" action="{{ url_for('mark_returned') }}" class="row g-2">
        <div class="col-8"><input required name="txn" class="form-control" placeholder="Transaction ID"></div>
        <div class="col-4"><button class="btn btn-warning w-100">تحديث</button></div>
      </form>
    </div>

    <div class="card p-3 mt-3">
      <h6 class="mb-3">استيراد من PDF</h6>
      <form method="post" action="{{ url_for('upload_pdf') }}" enctype="multipart/form-data" class="row g-2">
        <div class="col-12">
          <input required class="form-control" type="file" name="pdf" accept="application/pdf">
        </div>
        <div class="col-12"><button class="btn btn-primary w-100">إضافة ملف PDF</button></div>
      </form>
      <hr>
      <h6 class="mb-3">فاتورة مطابقة (تسليم تلقائي)</h6>
      <form method="post" action="{{ url_for('upload_invoice') }}" enctype="multipart/form-data" class="row g-2">
        <div class="col-12">
          <input required class="form-control" type="file" name="pdf" accept="application/pdf">
        </div>
        <div class="col-12"><button class="btn btn-success w-100">رفع فاتورة</button></div>
      </form>
    </div>

    <div class="card p-3 mt-3">
      <div class="d-grid gap-2">
        <a class="btn btn-outline-secondary" href="{{ url_for('dedupe') }}">حذف مكرر</a>
        <a class="btn btn-outline-secondary" href="{{ url_for('move_to_shipping') }}">تحديث إلى قيد التوصيل</a>
        <a class="btn btn-outline-secondary" href="{{ url_for('returns_bulk') }}">إدارة راجع</a>
        <a class="btn btn-outline-secondary" href="{{ url_for('delivered_bulk') }}">إدارة تم التوصيل</a>
        <a class="btn btn-outline-secondary" href="{{ url_for('pending') }}">الطلبات قيد التوصيل</a>
        <a class="btn btn-outline-primary" href="{{ url_for('stats') }}">الإحصائيات (مفصّل)</a>
            <a class="btn btn-outline-dark" href="{{ url_for('seam_home') }}">الخياطات</a>
    <a class="btn btn-outline-dark" href="{{ url_for('issues_home') }}">المشاكل</a>
    <a class="btn btn-outline-dark" href="{{ url_for('cutting_home') }}">طلبات الفصال</a>
        <a class="btn btn-outline-dark" href="{{ url_for('inventory_home') }}">المخزن</a>
      </div>
    </div>
  </div>
</div>
{% endblock %}
"""

EDIT_HTML = r"""
{% extends 'base.html' %}
{% block content %}
<div class="card p-3">
  <h5 class="mb-3">تعديل الطلب {{ txn }}</h5>
  <form method="post" class="row g-3">
    {% for c in columns %}
    <div class="col-md-6">
      <label class="form-label">{{ c }}</label>
      <input class="form-control" name="{{ c }}" value="{{ row.get(c,'') }}">
    </div>
    {% endfor %}
    <div class="col-12 text-end">
      <button class="btn btn-primary">حفظ</button>
      <a class="btn btn-outline-secondary" href="{{ url_for('home') }}">إلغاء</a>
    </div>
  </form>
</div>
{% endblock %}
"""

BULK_HTML = r"""
{% extends 'base.html' %}
{% block content %}
<div class="card p-3">
  <h5 class="mb-3">{{ title }}</h5>
  {% if product_name is not none %}
  <form method="post" class="row g-2 mb-3">
    <div class="col-md-5">
      <label class="form-label">اسم المنتج (للمجموعة)</label>
      <input name="product_name" class="form-control" value="{{ product_name or '' }}" placeholder="مثال: عباءة موديل 123">
    </div>
    <div class="col-md-4">
      <label class="form-label">اسم البيج</label>
      <select name="page_name" class="form-select">
        <option value="">بدون</option>
        {% for p in PAGES or [] %}
          <option value="{{p}}" {{ 'selected' if page_name==p else '' }}>{{p}}</option>
        {% endfor %}
      </select>
    </div>
    <div class="col-md-3 align-self-end"><button name="apply_name" value="1" class="btn btn-outline-primary w-100">تطبيق</button></div>
  </form>
  {% endif %}

  <form method="post" class="row g-2">
    <div class="col-md-6">
      <label class="form-label">رقم الشحنة</label>
      <input required name="txn" class="form-control" placeholder="Transaction ID">
    </div>
    <div class="col-md-3 align-self-end">
      <button class="btn btn-secondary w-100">إضافة إلى القائمة</button>
    </div>
    {% if action_label %}
    <div class="col-md-3 align-self-end">
      <button name="apply_all" value="1" class="btn btn-primary w-100">{{ action_label }}</button>
    </div>
    {% endif %}
  </form>

  <div class="table-responsive mt-3">
    <table class="table table-sm table-striped"><thead><tr>
      {% for h in headers %}<th>{{ h }}</th>{% endfor %}
    </tr></thead><tbody>
      {% for r in items %}
      <tr>
        {% for h in headers %}<td>{{ r.get(h,'') }}</td>{% endfor %}
      </tr>
      {% endfor %}
    </tbody></table>
  </div>
</div>
{% endblock %}
"""

PENDING_HTML = r"""
{% extends 'base.html' %}
{% block content %}
<div class="card p-3">
  <h5 class="mb-3">الطلبات قيد التوصيل</h5>
  <form method="get" class="row g-2">
    <div class="col-md-3"><label class="form-label">من</label><input name="from" type="date" class="form-control" value="{{ dfrom or '' }}"></div>
    <div class="col-md-3"><label class="form-label">إلى</label><input name="to" type="date" class="form-control" value="{{ dto or '' }}"></div>
    <div class="col-md-3 align-self-end"><button class="btn btn-secondary w-100">تصفية</button></div>
  </form>
  <div class="table-responsive mt-3">
    <table class="table table-striped">
      <thead><tr><th>Transaction ID</th><th>Time and Date</th><th>Order Price</th><th>Status</th></tr></thead>
      <tbody>
        {% for r in rows %}
        <tr><td>{{ r['Transaction ID'] }}</td><td>{{ r['Time and Date'] }}</td><td>{{ r['Order Price'] }}</td><td>{{ r['Status'] }}</td></tr>
        {% endfor %}
      </tbody>
    </table>
  </div>
</div>
{% endblock %}
"""

STATS_HTML = r"""
{% extends 'base.html' %}
{% block content %}
<form method="get" class="card p-3 mb-3">
  <div class="row g-2">
    <div class="col-md-3"><label class="form-label">من</label><input name="from" type="date" class="form-control" value="{{ dfrom or '' }}"></div>
    <div class="col-md-3"><label class="form-label">إلى</label><input name="to" type="date" class="form-control" value="{{ dto or '' }}"></div>
    <div class="col-md-3"><label class="form-label">اسم البيج</label>
      <select name="page" class="form-select">
        <option value="">الكل</option>
        {% for p in pages %}<option value="{{p}}" {{ 'selected' if sel_page==p else '' }}>{{p}}</option>{% endfor %}
      </select>
    </div>
    <div class="col-md-3 align-self-end"><button class="btn btn-secondary w-100">تطبيق</button></div>
  </div>
</form>

<div class="row g-3">
  <div class="col-xl-6">
    <div class="card p-3">
      <h6>ملخص عام</h6>
      <div class="row row-cols-1 row-cols-md-2 g-2 mt-2">
        {% for k, v in summary.items() %}
        <div class="col"><div class="border rounded p-3"> <div class="small text-muted">{{ k }}</div><div class="fw-bold fs-5">{{ v }}</div></div></div>
        {% endfor %}
        <div class="col"><div class="border rounded p-3"> <div class="small text-muted">الإيراد المُسلّم</div><div class="fw-bold fs-5">{{ revenue }}</div></div></div>
      </div>
    </div>
  </div>
  <div class="col-xl-6">
    <div class="card p-3">
      <h6>حسب السعر (Order Price)</h6>
      <div class="table-responsive mt-2">
        <table class="table table-sm table-striped">
          <thead><tr>{% for h in price_cols %}<th>{{ h }}</th>{% endfor %}</tr></thead>
          <tbody>
            {% for r in by_price %}
            <tr>{% for h in price_cols %}<td>{{ r.get(h,'') }}</td>{% endfor %}</tr>
            {% endfor %}
          </tbody>
        </table>
      </div>
    </div>
  </div>
</div>

<div class="card p-3 mt-3">
  <h6>اتجاه يومي</h6>
  <div class="table-responsive">
    <table class="table table-striped">
      <thead><tr><th>التاريخ</th><th>عدد الطلبات</th><th>الاتجاه</th></tr></thead>
      <tbody>
        {% for r in daily %}
        <tr><td>{{ r['Date'] }}</td><td>{{ r['Order Count'] }}</td><td>{{ r['Trend'] }}</td></tr>
      {% endfor %}
      </tbody>
    </table>
  </div>
</div>
{% endblock %}
"""

# Register templates in-memory (use DictLoader so `{% extends 'base.html' %}` works)
from jinja2 import DictLoader
app.jinja_loader = DictLoader({
    'base.html': BASE_HTML,
    'login.html': LOGIN_HTML,
    'home.html': HOME_HTML,
    'edit.html': EDIT_HTML,
    'bulk.html': BULK_HTML,
    'pending.html': PENDING_HTML,
    'stats.html': STATS_HTML,
    'inventory.html': INVENTORY_HTML,
    'seamstress.html': SEAMSTRESS_HTML,
    'issues.html': ISSUES_HTML,
    'cutting.html': CUTTING_HTML,
})

# --------------------------- AUTH DECORATOR ----------------------------
from functools import wraps

def login_required(fn):
    @wraps(fn)
    def _wrap(*args, **kwargs):
        if not session.get('auth'):
            return redirect(url_for('login'))
        return fn(*args, **kwargs)
    return _wrap

# ---------------------------- EXTRACTORS -------------------------------

def extract_from_text(text: str):
    text = normalize_digits(text)
    lines = [ln.strip() for ln in (text or "").splitlines() if ln.strip()]
    full = "\n".join(lines)

    txn = None
    m = re.search(r'رقم\s*(?:الشحنة|الوصل|الطلب)\s*[:：]?\s*(\d{6,})', full)
    if m:
        txn = m.group(1)
    else:
        m2 = re.search(r'(?<!\d)(?!07)\d{8,14}(?!\d)', full)
        if m2:
            txn = m2.group(0)

    phones = re.findall(r'(07\d{9})', full)
    seen = set(); uniq = []
    for p in phones:
        if p not in seen:
            seen.add(p); uniq.append(p)
    phone_str = ", ".join(uniq) if uniq else None

    def parse_price_from_lines(ls):
        label = r'(المبلغ(?:\s*الكلي)?(?:\s*ل?لفاتورة)?|السعر|قيمة\s*الطلب|Price|Total|IQD|دينار|د\.ع)'
        num   = r'(\d{1,3}(?:,\d{3})+|\d{4,9})'
        for ln in ls:
            cand = ln
            m1 = re.search(fr'{label}[^\d]{{0,40}}{num}', cand)
            if m1:
                v = int(m1.group(2 if m1.lastindex and m1.lastindex >= 2 else 1).replace(",", ""))
                if str(v).endswith("000"):
                    return v
            m2 = re.search(fr'{num}\s*{label}', cand)
            if m2:
                v = int(m2.group(1).replace(",", ""))
                if str(v).endswith("000"):
                    return v
        all_nums = [int(n.replace(",", "")) for n in re.findall(r'(\d{1,3}(?:,\d{3})+|\d{4,9})', full)]
        candidates = [n for n in all_nums if str(n).endswith("000")]
        return max(candidates) if candidates else None

    order_price = parse_price_from_lines(lines)

    def parse_address(ls):
        for i, ln in enumerate(ls):
            m = re.search(r'(?:العنوان|عنوان\s*الزبون|Address)\s*[:：]?\s*(.+)$', ln)
            if m and m.group(1).strip():
                return m.group(1).strip(" ,:؛-")
            if any(lbl in ln for lbl in ("العنوان", "عنوان الزبون", "Address")):
                if i+1 < len(ls) and ls[i+1].strip():
                    return ls[i+1].strip(" ,:؛-")
                if i > 0 and ls[i-1].strip():
                    return ls[i-1].strip(" ,:؛-")
        return None

    address = parse_address(lines)
    return txn, phone_str, order_price, address
# ---------------------------- SEAM / SEW STORE --------------------------
# ------------------------------ ISSUES STORE ----------------------------
class IssuesStore:
    COLS = ['ID', 'Title', 'Description', 'ImagePath', 'Status', 'Solver', 'CreatedAt']

    def __init__(self, root_dir: Path):
        self.path = root_dir / 'issues.xlsx'
        self.df = self._load()

    def _load(self):
        if not self.path.exists():
            df = pd.DataFrame(columns=self.COLS)
            df.to_excel(self.path, index=False)
            return df
        df = pd.read_excel(self.path)
        for c in self.COLS:
            if c not in df.columns:
                df[c] = pd.NA
        return df[self.COLS]

    def _save(self):
        self.df.to_excel(self.path, index=False)

    def _next_id(self):
        if self.df.empty:
            return 1
        vals = pd.to_numeric(self.df['ID'], errors='coerce').dropna()
        return int(vals.max() + 1) if len(vals) else 1

    def add_issue(self, title, desc='', img_path=''):
        new_id = self._next_id()
        row = {
            'ID': new_id,
            'Title': title,
            'Description': desc,
            'ImagePath': img_path,
            'Status': 'Open',
            'Solver': '',
            'CreatedAt': now_str(),
        }
        self.df = pd.concat([self.df, pd.DataFrame([row])], ignore_index=True)
        self._save()

    def solve(self, iid, solver):
        idx = self.df[self.df['ID'] == iid].index
        if not len(idx):
            return
        i = idx[0]
        self.df.at[i, 'Status'] = 'Solved'
        self.df.at[i, 'Solver'] = solver
        self._save()

    def delete(self, iid):
        self.df = self.df[self.df['ID'] != iid]
        self._save()


issues = IssuesStore(_data_root)

class SeamStore:
    MAST_COLS = ['ID', 'Name', 'Phone', 'Notes', 'Active']
    LOG_COLS = ['LogID', 'Date', 'SeamstressID', 'Model', 'Pieces', 'UnitCost', 'Total', 'Paid']

    def __init__(self, root_dir: Path):
        self.mast_path = root_dir / 'seamstresses.xlsx'
        self.log_path = root_dir / 'sewing_logs.xlsx'
        self.mast = self._load_mast()
        self.log = self._load_log()

    def _load_mast(self):
        if not self.mast_path.exists():
            df = pd.DataFrame(columns=self.MAST_COLS)
            df.to_excel(self.mast_path, index=False)
            return df
        df = pd.read_excel(self.mast_path)
        for c in self.MAST_COLS:
            if c not in df.columns:
                df[c] = pd.NA
        return df[self.MAST_COLS]

    def _load_log(self):
        if not self.log_path.exists():
            df = pd.DataFrame(columns=self.LOG_COLS)
            df.to_excel(self.log_path, index=False)
            return df
        df = pd.read_excel(self.log_path)
        for c in self.LOG_COLS:
            if c not in df.columns:
                df[c] = pd.NA
        return df[self.LOG_COLS]

    def _save_mast(self):
        self.mast.to_excel(self.mast_path, index=False)

    def _save_log(self):
        self.log.to_excel(self.log_path, index=False)

    def _next_id(self, col_name, df):
        if df.empty or col_name not in df.columns:
            return 1
        vals = pd.to_numeric(df[col_name], errors='coerce')
        vals = vals.dropna()
        return int(vals.max() + 1) if len(vals) else 1

    def add_seamstress(self, name, phone='', notes=''):
        new_id = self._next_id('ID', self.mast)
        row = {
            'ID': new_id,
            'Name': name,
            'Phone': phone,
            'Notes': notes,
            'Active': True,
        }
        self.mast = pd.concat([self.mast, pd.DataFrame([row])], ignore_index=True)
        self._save_mast()

    def update_seamstress(self, sid, **kwargs):
        idx = self.mast[self.mast['ID'] == sid].index
        if not len(idx):
            return
        i = idx[0]
        for k, v in kwargs.items():
            if k in self.mast.columns:
                self.mast.at[i, k] = v
        self._save_mast()

    def delete_seamstress(self, sid):
        self.mast = self.mast[self.mast['ID'] != sid]
        # حذف السجلات المرتبطة من سجل الإنجاز
        self.log = self.log[self.log['SeamstressID'] != sid]
        self._save_mast()
        self._save_log()

    def add_log(self, sid, model, pieces, unit_cost):
        log_id = self._next_id('LogID', self.log)
        total = float(pieces) * float(unit_cost)
        row = {
            'LogID': log_id,
            'Date': date.today().isoformat(),
            'SeamstressID': sid,
            'Model': model,
            'Pieces': pieces,
            'UnitCost': unit_cost,
            'Total': total,
            'Paid': False,
        }
        self.log = pd.concat([self.log, pd.DataFrame([row])], ignore_index=True)
        self._save_log()
        # زيادة المخزون تلقائيًا بالموديل وعدد القطع
        try:
            inventory.adjust_quantity(model, pieces)
        except Exception:
            pass

    def set_paid(self, log_id, paid: bool):
        idx = self.log[self.log['LogID'] == log_id].index
        if not len(idx):
            return
        self.log.at[idx[0], 'Paid'] = bool(paid)
        self._save_log()


# إنشاء كائن seams

seams = SeamStore(_data_root)


# ------------------------------- INVENTORY ------------------------------
class InventoryStore:
    COLS = [
        'Product Code','Product Name','Type','Quantity','Fabric Meters','Meters per Unit',
        'Sewing Cost','Other Costs','Sale Price'
    ]
    def __init__(self, path):
        self.path = str(Path(path).with_name('inventory.xlsx'))
        self.df = self._load()
    def _load(self):
        p = Path(self.path)
        if not p.exists():
            df = pd.DataFrame(columns=self.COLS)
            df.to_excel(self.path, index=False)
            return df
        df = pd.read_excel(self.path)
        for c in self.COLS:
            if c not in df.columns:
                df[c] = pd.NA
        return df[self.COLS]
    def save(self):
        self.df.to_excel(self.path, index=False)
    def next_code(self):
        prefix = 'INV'
        nums = [int(str(x).replace(prefix,'') or 0) for x in self.df['Product Code'].dropna().astype(str) if str(x).startswith(prefix)]
        n = (max(nums) if nums else 0) + 1
        return f'{prefix}{n:04d}'
    def add_item(self, row):
        row = {**{c: pd.NA for c in self.COLS}, **row}
        self.df = pd.concat([self.df, pd.DataFrame([row])], ignore_index=True)
        self.save()
    def adjust_quantity(self, name, delta):
        idx = self.df[self.df['Product Name'].astype(str)==str(name)].index
        if not len(idx):
            return
        i = idx[0]
        q = pd.to_numeric(self.df.at[i,'Quantity'], errors='coerce')
        q = int(q) if pd.notna(q) else 0
        self.df.at[i,'Quantity'] = q + int(delta)
        # meters per unit
        mpu = pd.to_numeric(self.df.at[i,'Meters per Unit'], errors='coerce')
        mpu = float(mpu) if pd.notna(mpu) else 0
        fm = pd.to_numeric(self.df.at[i,'Fabric Meters'], errors='coerce')
        fm = float(fm) if pd.notna(fm) else 0
        self.df.at[i,'Fabric Meters'] = max(0.0, fm - (mpu*delta)) if delta>0 else max(0.0, fm)
        self.save()

inventory = InventoryStore(EXCEL_FILE)

# hook: adjust inventory when status transitions
def adjust_inventory_on_transition(row, old_status, new_status):
    try:
        name = row.get('Product Name')
        if not name:
            return
        # READY -> SHIPPING: decrement 1
        if old_status == STATUS_READY and new_status == STATUS_SHIPPING:
            inventory.adjust_quantity(name, -1)
        # SHIPPING -> RETURNED: add back 1
        if old_status == STATUS_SHIPPING and new_status == STATUS_RETURNED:
            inventory.adjust_quantity(name, +1)
    except Exception:
        pass

# --------------------------- INVENTORY TEMPLATES ------------------------

# --------------------------- EXTRA UPLOAD HELPERS ----------------------
UPLOAD_DIR = user_data_dir() / 'uploads'
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

ALLOWED_IMG_EXT = {'.png', '.jpg', '.jpeg', '.webp'}

def _is_allowed_image(filename: str) -> bool:
    ext = Path(filename).suffix.lower()
    return ext in ALLOWED_IMG_EXT

def _save_image(file_storage):
    if not file_storage or not file_storage.filename:
        return ''
    if not _is_allowed_image(file_storage.filename):
        return ''
    fname = secure_filename(file_storage.filename)
    dst = UPLOAD_DIR / (datetime.now().strftime('%Y%m%d%H%M%S_') + fname)
    file_storage.save(dst)
    return str(dst)

# --------------------------- SEAMSTRESS TEMPLATE ------------------------



# ------------------------------- ROUTES --------------------------------
@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        code = (request.form.get('code') or '').strip()
        if code == PASSCODE:
            session['auth'] = True
            return redirect(url_for('home'))
        flash('رمز غير صحيح', 'err')
    return render_template_string(LOGIN_HTML)


@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))


@app.route('/')
@login_required
def home():
    q = (request.args.get('q') or '').strip()
    prod = (request.args.get('product') or '').strip()
    page = (request.args.get('page') or '').strip()
    dfrom = request.args.get('from')
    dto = request.args.get('to')

    d = store.df.copy()
    # text search
    if q:
        mask = pd.Series(False, index=d.index)
        for c in BASE_COLUMNS:
            if c in d.columns:
                mask = mask | d[c].astype(str).str.contains(re.escape(q), na=False)
        d = d[mask]
    # product/page filter
    if prod:
        d = d[d['Product Name'].astype(str) == prod]
    if page:
        d = d[d['Page Name'].astype(str) == page]
    # date range
    if 'Time and Date' in d.columns:
        d['Time and Date'] = pd.to_datetime(d['Time and Date'], errors='coerce')
        if dfrom:
            start = datetime.strptime(dfrom, '%Y-%m-%d')
            d = d[d['Time and Date'] >= start]
        if dto:
            end = datetime.strptime(dto, '%Y-%m-%d')
            d = d[d['Time and Date'] <= end]
        d = d.sort_values('Time and Date', ascending=False, na_position='last')
        d['Time and Date'] = d['Time and Date'].dt.strftime('%Y-%m-%d %H:%M:%S')

    rows = d.fillna("").to_dict(orient='records')
    # populate filter dropdowns
    all_products = sorted(list({str(x) for x in store.df['Product Name'].dropna().unique()}))
    all_pages = sorted(list({str(x) for x in store.df['Page Name'].dropna().unique()}))
    return render_template_string(HOME_HTML, columns=BASE_COLUMNS, rows=rows, q=q,
                                  all_products=all_products, all_pages=all_pages,
                                  sel_product=prod, sel_page=page, dfrom=dfrom, dto=dto)


@app.route('/mark_returned', methods=['POST'])
@login_required
def mark_returned():
    txn = (request.form.get('txn') or '').strip()
    ok, msg = store.update_status(txn, STATUS_RETURNED)
    if ok:
        store.save(); flash('تم تحديث الحالة إلى راجع', 'ok')
    else:
        flash(msg, 'err')
    return redirect(url_for('home'))


@app.route('/upload_pdf', methods=['POST'])
@login_required
def upload_pdf():
    file = request.files.get('pdf')
    if not file:
        flash('يرجى اختيار ملف PDF', 'err'); return redirect(url_for('home'))
    path = Path(UPLOAD_DIR) / f"import_{int(datetime.now().timestamp())}.pdf"
    file.save(path)

    client_count = {}
    added, updated = 0, 0
    page_errors = []
    try:
        with pdfplumber.open(str(path)) as pdf:
            for page_num, page in enumerate(pdf.pages, start=1):
                try:
                    text = page.extract_text() or ""
                    txn, phone_str, order_price, address = extract_from_text(text)
                    if not txn:
                        continue
                    main_phone = None
                    if phone_str:
                        main_phone = phone_str.split(',')[0].strip()
                        client_count[main_phone] = client_count.get(main_phone, 0) + 1
                    page_data = {
                        "Product Name": pd.NA,
                        "Transaction ID": str(txn),
                        "Time and Date": now_str(),
                        "Contact Numbers": phone_str,
                        "Address": address,
                        "Order Price": order_price,
                        "Status": STATUS_READY,
                        "Return Reason": "لا يوجد",
                        "Notes": None,
                        "Client Orders Count": client_count.get(main_phone, 1) if main_phone else pd.NA,
                    }
                    ok, msg = store.upsert_row(page_data)
                    if ok and msg == "تمت الإضافة":
                        added += 1
                    elif ok and msg == "تم التحديث":
                        updated += 1
                except Exception as pe:
                    page_errors.append((page_num, f"{type(pe).__name__}: {pe}"))
        store.save()
        info = f"تمت معالجة PDF. المضاف: {added} | المحدّث: {updated}"
        if page_errors:
            info += f" | تعذّر قراءة {len(page_errors)} صفحة"
        flash(info, 'ok')
    except Exception as e:
        _fatal_box('فشل استيراد PDF', e)
        flash('فشل استيراد PDF', 'err')
    return redirect(url_for('home'))


@app.route('/upload_invoice', methods=['POST'])
@login_required
def upload_invoice():
    file = request.files.get('pdf')
    if not file:
        flash('يرجى اختيار ملف PDF', 'err'); return redirect(url_for('home'))
    path = Path(UPLOAD_DIR) / f"invoice_{int(datetime.now().timestamp())}.pdf"
    file.save(path)

    updated_rows, skipped_rows = [], []
    try:
        with pdfplumber.open(str(path)) as pdf:
            for page in pdf.pages:
                text = normalize_digits(page.extract_text() or "")
                for ln in text.split("\n"):
                    ln = ln.strip()
                    m = re.search(r'(\d{6,})\s+((?:\d{1,3}(?:,\d{3})+|\d{4,9}))', ln)
                    if not m:
                        continue
                    txn = m.group(1).strip()
                    price_val = to_int(m.group(2))
                    if price_val is None or not str(price_val).endswith("000"):
                        continue
                    if store.exists(txn):
                        exist = store.get_row(txn)
                        exist_price = pd.to_numeric(exist.get("Order Price"), errors="coerce")
                        if pd.notna(exist_price) and int(exist_price) == int(price_val):
                            store.update_status(txn, STATUS_DELIVERED)
                            updated_rows.append((txn, price_val, "OK"))
                        else:
                            skipped_rows.append((txn, price_val, f"سعر مختلف (المسجل: {exist_price})"))
                    else:
                        skipped_rows.append((txn, price_val, "الشحنة غير موجودة"))
        store.save()
        flash(f"تم التحديث: {len(updated_rows)} | لم يتم: {len(skipped_rows)}", 'ok')
    except Exception as e:
        _fatal_box('فشل رفع الفاتورة', e)
        flash('فشل رفع الفاتورة', 'err')
    return redirect(url_for('home'))


@app.route('/dedupe')
@login_required
def dedupe():
    removed = store.drop_duplicates_keep_last()
    store.save()
    flash(f"تم حذف {removed} صف مكرر.", 'ok')
    return redirect(url_for('home'))


@app.route('/delete/<txn>')
@login_required
def delete(txn):
    deleted = store.drop_by_txn(txn)
    if deleted:
        store.save(); flash('تم الحذف', 'ok')
    else:
        flash('الشحنة غير موجودة', 'err')
    return redirect(url_for('home'))


@app.route('/edit/<txn>', methods=['GET', 'POST'])
@login_required
def edit(txn):
    if not store.exists(txn):
        abort(404)
    if request.method == 'POST':
        new_vals = {c: request.form.get(c) for c in BASE_COLUMNS}
        if 'Order Price' in new_vals:
            new_vals['Order Price'] = pd.to_numeric(new_vals['Order Price'], errors='coerce')
        ok, msg = store.upsert_row(new_vals)
        if ok:
            store.save(); flash('تم التعديل', 'ok'); return redirect(url_for('home'))
        flash(msg, 'err')
    row = store.get_row(txn).fillna("").to_dict()
    return render_template_string(EDIT_HTML, txn=txn, columns=BASE_COLUMNS, row=row)


@app.route('/move-to-shipping', methods=['GET', 'POST'])
@login_required
def move_to_shipping():
    session.setdefault('shipping_items', [])
    session['shipping_items'] = list(dict.fromkeys(session['shipping_items']))
    headers = ['Transaction ID', 'Page', 'Product', 'Status']
    title = 'تحديث الحالة إلى قيد التوصيل'
    product_name = session.get('product_name', '')
    page_name = session.get('page_name', '')

    PAGES = ['فاتنة','لمسة حرير','براعم','أنيقا','خيوط']

    if request.method == 'POST':
        if request.form.get('apply_name'):
            name = (request.form.get('product_name') or '').strip()
            pg = (request.form.get('page_name') or '').strip()
            session['product_name'] = name
            session['page_name'] = pg
            count = 0
            if session['shipping_items']:
                for txn in session['shipping_items']:
                    if store.exists(txn):
                        if name:
                            store.df.at[txn, 'Product Name'] = name
                        if pg:
                            store.df.at[txn, 'Page Name'] = pg
                        count += 1
                store.save()
                flash(f'تم تطبيق الاسم/البيج على {count} شحنة', 'ok')
            return redirect(url_for('move_to_shipping'))
        if request.form.get('apply_all'):
            flash('تم تحديث الحالات الحالية إلى قيد التوصيل', 'ok')
            return redirect(url_for('move_to_shipping'))
        txn = (request.form.get('txn') or '').strip()
        ok, info = store.update_status(txn, STATUS_SHIPPING)
        if ok:
            # set product/page immediately if chosen
            if store.exists(txn):
                if product_name:
                    store.df.at[txn, 'Product Name'] = product_name
                if page_name:
                    store.df.at[txn, 'Page Name'] = page_name
            if txn not in session['shipping_items']:
                session['shipping_items'].append(txn)
            store.save()
        else:
            flash(info, 'err')
        return redirect(url_for('move_to_shipping'))

    def row(txn):
        p = store.get_row(txn) if store.exists(txn) else None
        if p is not None:
            try:
                page_val = p.get('Page Name', '')
                prod_val = p.get('Product Name', '')
            except Exception:
                # fallback if p is a plain dict
                page_val = p['Page Name'] if isinstance(p, dict) and 'Page Name' in p else ''
                prod_val = p['Product Name'] if isinstance(p, dict) and 'Product Name' in p else ''
        else:
            page_val, prod_val = '', ''
        return {"Transaction ID": txn, "Page": page_val, "Product": prod_val, "Status": STATUS_SHIPPING}
    items = [row(t) for t in session['shipping_items']]
    return render_template_string(BULK_HTML, title=title, headers=headers, items=items,
                                  action_label=f"تطبيق الكل -> {STATUS_SHIPPING}",
                                  product_name=product_name,
                                  PAGES=PAGES, page_name=page_name)


@app.route('/returns-bulk', methods=['GET', 'POST'])
@login_required
def returns_bulk():
    session.setdefault('returns_items', [])
    session['returns_items'] = list(dict.fromkeys(session['returns_items']))
    headers = ['Transaction ID', 'Status', 'Reason']
    title = 'إدارة الطلبات الراجعة'
    if request.method == 'POST':
        if request.form.get('apply_all'):
            for txn in session['returns_items']:
                if store.exists(txn):
                    store.update_status(txn, STATUS_RETURNED)
            store.save(); flash('تم تحديث الحالات', 'ok')
            return redirect(url_for('returns_bulk'))
        txn = (request.form.get('txn') or '').strip()
        if txn and txn not in session['returns_items']:
            session['returns_items'].append(txn)
        return redirect(url_for('returns_bulk'))
    items = [{"Transaction ID": t, "Status": STATUS_RETURNED, "Reason": ""} for t in session['returns_items']]
    return render_template_string(BULK_HTML, title=title, headers=headers, items=items,
                                  action_label=f"تطبيق الكل -> {STATUS_RETURNED}", product_name=None)


@app.route('/delivered-bulk', methods=['GET', 'POST'])
@login_required
def delivered_bulk():
    session.setdefault('delivered_items', [])
    session['delivered_items'] = list(dict.fromkeys(session['delivered_items']))
    headers = ['Transaction ID', 'Order Price', 'Status']
    title = 'إدارة الطلبات التي تم تسليمها'
    if request.method == 'POST':
        if request.form.get('apply_all'):
            for txn in session['delivered_items']:
                if store.exists(txn):
                    store.update_status(txn, STATUS_DELIVERED)
            store.save(); flash('تم تحديث الحالات', 'ok')
            return redirect(url_for('delivered_bulk'))
        txn = (request.form.get('txn') or '').strip()
        if txn and txn not in session['delivered_items']:
            session['delivered_items'].append(txn)
        return redirect(url_for('delivered_bulk'))
    def row(txn):
        pr = ''
        if store.exists(txn):
            pr = store.get_row(txn).get('Order Price', '')
        return {"Transaction ID": txn, "Order Price": pr, "Status": STATUS_DELIVERED}
    items = [row(t) for t in session['delivered_items']]
    return render_template_string(BULK_HTML, title=title, headers=headers, items=items,
                                  action_label=f"تطبيق الكل -> {STATUS_DELIVERED}", product_name=None)


@app.route('/pending')
@login_required
def pending():
    dfrom = request.args.get('from')
    dto = request.args.get('to')
    d = store.df.copy()
    d = d[d['Status'] == STATUS_SHIPPING]
    d['Time and Date'] = pd.to_datetime(d['Time and Date'], errors='coerce')
    if dfrom:
        start = datetime.strptime(dfrom, '%Y-%m-%d')
        d = d[d['Time and Date'] >= start]
    if dto:
        end = datetime.strptime(dto, '%Y-%m-%d')
        d = d[d['Time and Date'] <= end]
    d = d.sort_values('Time and Date', ascending=False)
    out = []
    for _, r in d.iterrows():
        ts = r['Time and Date']
        ts = ts.strftime('%Y-%m-%d %H:%M:%S') if pd.notna(ts) else ''
        out.append({'Transaction ID': r['Transaction ID'], 'Time and Date': ts,
                    'Order Price': r['Order Price'], 'Status': r['Status']})
    return render_template_string(PENDING_HTML, rows=out, dfrom=dfrom, dto=dto)


@app.route('/stats', methods=['GET', 'POST'])
@login_required
def stats():
    # Gate with secondary passcode 998144
    if not session.get('stats_auth'):
        if request.method == 'POST' and (request.form.get('code') or '').strip() == '998144':
            session['stats_auth'] = True
        else:
            return render_template_string("""
            {% extends 'base.html' %}
            {% block content %}
            <div class='row justify-content-center'>
              <div class='col-md-5'><div class='card p-4'>
                <h6 class='mb-3'>رمز دخول الإحصائيات</h6>
                <form method='post'>
                  <input name='code' type='password' class='form-control mb-3' placeholder='••••••'>
                  <button class='btn btn-primary w-100'>دخول</button>
                </form>
              </div></div></div>
            {% endblock %}
            """)

    dfrom = request.args.get('from')
    dto = request.args.get('to')
    sel_page = (request.args.get('page') or '').strip()

    d = store.df.copy()
    d['Time and Date'] = pd.to_datetime(d['Time and Date'], errors='coerce')
    if dfrom:
        start = datetime.strptime(dfrom, '%Y-%m-%d')
        d = d[d['Time and Date'] >= start]
    if dto:
        end = datetime.strptime(dto, '%Y-%m-%d')
        d = d[d['Time and Date'] <= end]
    if sel_page:
        d = d[d['Page Name'].astype(str) == sel_page]

    summary = store.stats_global(d)

    by_price_df = store.stats_by_product_price(d)
    by_price = by_price_df.fillna("").to_dict(orient='records') if not by_price_df.empty else []

    daily_df = store.daily_trend(d)
    daily = []
    if not daily_df.empty:
        for _, r in daily_df.iterrows():
            daily.append({'Date': r['Date'].strftime('%Y-%m-%d') if hasattr(r['Date'],'strftime') else str(r['Date']),
                          'Order Count': int(r['Order Count']), 'Trend': r['Trend']})

    # simple revenue/profit by page (profit based on inventory costs if available)
    rev = pd.to_numeric(d.loc[d['Status']==STATUS_DELIVERED,'Order Price'], errors='coerce').sum()

    return render_template_string(STATS_HTML, summary=summary, by_price=by_price,
                                  price_cols=["السعر","عدد الطلبات",STATUS_DELIVERED,STATUS_RETURNED,STATUS_SHIPPING,STATUS_READY,"المبلغ المُسلَّم","نسبة الراجع %"],
                                  daily=daily, dfrom=dfrom, dto=dto, sel_page=sel_page,
                                  pages=sorted(list({str(x) for x in store.df['Page Name'].dropna().unique()})),
                                  revenue=rev)


@app.route('/download/excel')
@login_required
def download_excel():
    # make sure latest is saved, then send
    store.save()
    d = Path(EXCEL_FILE).parent
    return send_from_directory(str(d), Path(EXCEL_FILE).name, as_attachment=True)


@app.route('/inventory')
@login_required
def inventory_home():
    rows = inventory.df.fillna("").to_dict(orient='records')
    added = request.args.get('added')
    taken = request.args.get('taken')
    name = request.args.get('name')
    return render_template_string(INVENTORY_HTML, rows=rows, added=added, taken=taken, name=name)

@app.route('/inventory/add', methods=['POST'])
@login_required
def inventory_add():
    name = (request.form.get('name') or '').strip()
    if not name:
        flash('يرجى إدخال اسم المنتج', 'err'); return redirect(url_for('inventory_home'))
    row = {
        'Product Code': inventory.next_code(),
        'Product Name': name,
        'Type': (request.form.get('type') or '').strip(),
        'Quantity': int(request.form.get('qty') or 0),
        'Fabric Meters': float(request.form.get('fabric') or 0),
        'Meters per Unit': float(request.form.get('mpu') or 0),
        'Sewing Cost': float(request.form.get('sew') or 0),
        'Other Costs': float(request.form.get('other') or 0),
        'Sale Price': float(request.form.get('price') or 0),
    }
    inventory.add_item(row)
    flash('تمت إضافة الصنف', 'ok')
    return redirect(url_for('inventory_home'))

@app.route('/inventory/adjust', methods=['POST'])
@login_required
def inventory_adjust():
    name = (request.form.get('name') or '').strip()
    try:
        delta = int(request.form.get('delta') or 0)
    except Exception:
        delta = 0
    if not name or not delta:
        flash('بيانات غير مكتملة', 'err'); return redirect(url_for('inventory_home'))
    inventory.adjust_quantity(name, delta)
    flash('تم تعديل الكمية', 'ok')
    # redirect with modal params
    if delta>0:
        return redirect(url_for('inventory_home', added=str(delta), name=name))
    else:
        return redirect(url_for('inventory_home', taken=str(abs(delta)), name=name))

@app.route('/inventory/adjust-bulk', methods=['POST'])
@login_required
def inventory_adjust_bulk():
    name = (request.form.get('name') or '').strip()
    try:
        qty = int(request.form.get('qty') or 0)
    except Exception:
        qty = 0
    if not name or qty == 0:
        flash('يرجى إدخال اسم المنتج والكمية', 'err'); return redirect(url_for('inventory_home'))
    inventory.adjust_quantity(name, qty)
    flash('تم تعديل الكمية', 'ok')
    if qty>0:
        return redirect(url_for('inventory_home', added=str(qty), name=name))
    else:
        return redirect(url_for('inventory_home', taken=str(abs(qty)), name=name))
    inventory.adjust_quantity(name, delta)
    flash('تم تعديل الكمية', 'ok')
    return redirect(url_for('inventory_home'))

@app.route('/seamstresses')
@login_required
def seam_home():
    # الخياطات
    seamstresses_df = seams.mast.fillna('')
    seamstresses = seamstresses_df.to_dict(orient='records')
    seam_name_map = {r['ID']: r['Name'] for _, r in seamstresses_df.iterrows()}

    # قيم الفلتر من الـ query string
    dfrom = request.args.get('from') or ''
    dto = request.args.get('to') or ''
    sel_sid = request.args.get('sid') or ''
    sel_paid = request.args.get('paid') or ''

    logs = []
    if hasattr(seams, 'log') and isinstance(seams.log, pd.DataFrame) and not seams.log.empty:
        logs_df = seams.log.copy().fillna('')

        # تحويل التاريخ لنوع datetime حتى نفلتر صح
        logs_df['Date'] = pd.to_datetime(logs_df['Date'], errors='coerce')

        if dfrom:
            start = datetime.strptime(dfrom, '%Y-%m-%d')
            logs_df = logs_df[logs_df['Date'] >= start]
        if dto:
            end = datetime.strptime(dto, '%Y-%m-%d')
            logs_df = logs_df[logs_df['Date'] <= end]

        if sel_sid:
            try:
                sid_int = int(sel_sid)
                logs_df = logs_df[logs_df['SeamstressID'] == sid_int]
            except Exception:
                pass

        if sel_paid in ('paid', 'unpaid'):
            if sel_paid == 'paid':
                logs_df = logs_df[logs_df['Paid'] == True]
            else:
                logs_df = logs_df[logs_df['Paid'] == False]

        logs_df = logs_df.sort_values(by='Date', ascending=False)
        # تنسيق التاريخ للعرض
        logs_df['Date'] = logs_df['Date'].dt.strftime('%Y-%m-%d')
        logs = logs_df.to_dict(orient='records')

    return render_template_string(
        SEAMSTRESS_HTML,
        seamstresses=seamstresses,
        logs=logs,
        seam_name_map=seam_name_map,
        dfrom=dfrom,
        dto=dto,
        sel_sid=sel_sid,
        sel_paid=sel_paid,
    )

@app.route('/seam/add', methods=['POST'])
@login_required
def seam_add():
    name = (request.form.get('name') or '').strip()
    if not name:
        flash('يرجى إدخال الاسم', 'err'); return redirect(url_for('seam_home'))
    seams.add_seamstress(name, (request.form.get('phone') or '').strip(), (request.form.get('notes') or '').strip())
    flash('تمت الإضافة', 'ok'); return redirect(url_for('seam_home'))

@app.route('/seam/edit', methods=['POST'])
@login_required
def seam_edit():
    try:
        sid = int(request.form.get('id') or 0)
    except Exception:
        sid = 0
    if not sid:
        flash('معرّف غير صالح', 'err'); return redirect(url_for('seam_home'))
    seams.update_seamstress(sid, Name=request.form.get('name', ''), Phone=request.form.get('phone', ''), Notes=request.form.get('notes', ''), Active=bool(request.form.get('active')))
    flash('تم الحفظ', 'ok'); return redirect(url_for('seam_home'))

@app.route('/seam/delete/<int:sid>')
@login_required
def seam_delete(sid):
    seams.delete_seamstress(sid)
    flash('تم الحذف', 'ok'); return redirect(url_for('seam_home'))

@app.route('/sew/add', methods=['POST'])
@login_required
def sew_add_log():
    try:
        sid = int(request.form.get('sid') or 0)
        pieces = int(request.form.get('pieces') or 0)
        unit = float(request.form.get('unit') or 0)
    except Exception:
        flash('بيانات غير صالحة', 'err'); return redirect(url_for('seam_home'))
    model = (request.form.get('model') or '').strip()
    if not sid or not model or pieces<=0:
        flash('الرجاء إدخال الخياطة، الموديل، وعدد صحيح', 'err'); return redirect(url_for('seam_home'))
    seams.add_log(sid, model, pieces, unit)
    flash('تم تسجيل الإنجاز وزيادة المخزون', 'ok')
    return redirect(url_for('seam_home'))

@app.route('/sew/paid/<int:log_id>')
@login_required
def sew_mark_paid(log_id):
    seams.set_paid(log_id, True); flash('تمت التصفية', 'ok'); return redirect(url_for('seam_home'))

@app.route('/sew/unpaid/<int:log_id>')
@login_required
def sew_mark_unpaid(log_id):
    seams.set_paid(log_id, False); flash('تم الإلغاء', 'ok'); return redirect(url_for('seam_home'))

# ------------------------------ ISSUES ROUTES ---------------------------
@app.route('/issues')
@login_required
def issues_home():
    rows = issues.df.fillna('').sort_values(by='CreatedAt', ascending=False).to_dict(orient='records') if not issues.df.empty else []
    return render_template_string(ISSUES_HTML, rows=rows)

@app.route('/issues/add', methods=['POST'])
@login_required
@limiter.limit('20/minute')
def issues_add():
    title = (request.form.get('title') or '').strip()
    if not title:
        flash('العنوان مطلوب', 'err'); return redirect(url_for('issues_home'))
    img = request.files.get('image')
    img_path = _save_image(img)
    desc = (request.form.get('desc') or '').strip()

    issues.add_issue(title, desc, img_path)

    # 🔔 إشعار تلغرام عند إضافة مشكلة جديدة
    try:
        msg = (
            "⚠️ تم تسجيل مشكلة جديدة\n"
            f"العنوان: {title}\n"
            f"الوصف: {desc or 'لا يوجد'}\n"
            f"الوقت: {now_str()}"
        )
        send_telegram(msg)
    except Exception:
        pass

    flash('تمت إضافة المشكلة', 'ok')
    return redirect(url_for('issues_home'))


@app.route('/issues/solve', methods=['POST'])
@login_required
def issues_solve():
    try:
        iid = int(request.form.get('id') or 0)
    except Exception:
        iid = 0
    solver = (request.form.get('solver') or '').strip()
    if not iid or not solver:
        flash('بيانات غير مكتملة', 'err'); return redirect(url_for('issues_home'))
    issues.solve(iid, solver)
    flash('تم الحل', 'ok'); return redirect(url_for('issues_home'))

@app.route('/issues/delete/<int:iid>')
@login_required
def issues_delete(iid):
    issues.delete(iid); flash('تم الحذف', 'ok'); return redirect(url_for('issues_home'))

@app.route('/static-proxy')
@login_required
def static_proxy():
    # لعرض الصور المخزّنة خارج static
    from flask import send_file, request as _rq
    f = _rq.args.get('f')
    if not f or not Path(f).exists():
        return ('', 404)
    return send_file(f)

# ------------------------------ CUTTINGS STORE --------------------------
class CuttingsStore:
    COLS = ['ID', 'Model', 'ImagePath', 'DueDate', 'RequiredQty',
            'Status', 'Notes', 'RejectionReason', 'CreatedAt']

    def __init__(self, root_dir: Path):
        self.path = root_dir / 'cuttings.xlsx'
        self.df = self._load()

    def _load(self):
        if not self.path.exists():
            df = pd.DataFrame(columns=self.COLS)
            df.to_excel(self.path, index=False)
            return df
        df = pd.read_excel(self.path)
        for c in self.COLS:
            if c not in df.columns:
                df[c] = pd.NA
        return df[self.COLS]

    def _save(self):
        self.df.to_excel(self.path, index=False)

    def _next_id(self):
        if self.df.empty:
            return 1
        vals = pd.to_numeric(self.df['ID'], errors='coerce').dropna()
        return int(vals.max() + 1) if len(vals) else 1

    def add(self, model, due, qty, notes='', img_path=''):
        new_id = self._next_id()
        row = {
            'ID': new_id,
            'Model': model,
            'ImagePath': img_path,
            'DueDate': due,
            'RequiredQty': qty,
            'Status': 'قيد الانتظار',
            'Notes': notes,
            'RejectionReason': '',
            'CreatedAt': now_str(),
        }
        self.df = pd.concat([self.df, pd.DataFrame([row])], ignore_index=True)
        self._save()

    def update_status(self, cid, status, reason=None):
        idx = self.df[self.df['ID'] == cid].index
        if not len(idx):
            return
        i = idx[0]
        self.df.at[i, 'Status'] = status
        if reason is not None:
            self.df.at[i, 'RejectionReason'] = reason
        self._save()

    def delete(self, cid):
        self.df = self.df[self.df['ID'] != cid]
        self._save()


cuttings = CuttingsStore(_data_root)


# ------------------------------ CUTTING ROUTES --------------------------
@app.route('/cutting')
@login_required
def cutting_home():
    rows = cuttings.df.fillna('').sort_values(by='CreatedAt', ascending=False).to_dict(orient='records') if not cuttings.df.empty else []
    return render_template_string(CUTTING_HTML, rows=rows)

@app.route('/cutting/add', methods=['POST'])
@login_required
def cutting_add():
    model = (request.form.get('model') or '').strip()
    due = (request.form.get('due') or '').strip()
    try:
        qty = int(request.form.get('qty') or 0)
    except Exception:
        qty = 0
    if not model or not due or qty<=0:
        flash('بيانات غير مكتملة', 'err'); return redirect(url_for('cutting_home'))
    img = request.files.get('image')
    img_path = _save_image(img)
    notes = (request.form.get('notes') or '').strip()

    # إضافة الفصال في الإكسل
    cuttings.add(model, due, qty, notes, img_path)

    # 🔔 إشعار تلغرام
    try:
        msg = (
            "🧵 تم إنشاء طلب فصال جديد\n"
            f"الموديل: {model}\n"
            f"الكمية المطلوبة: {qty}\n"
            f"موعد الفصال: {due}\n"
            f"ملاحظات: {notes or 'لا يوجد'}\n"
            f"الوقت: {now_str()}"
        )
        send_telegram(msg)
    except Exception:
        pass

    flash('تم إنشاء طلب الفصال', 'ok')
    return redirect(url_for('cutting_home'))

@app.route('/cutting/status/<int:cid>')
@login_required
def cutting_status(cid):
    s = (request.args.get('s') or '').strip()
    if s not in ['قيد الانتظار','قيد العمل','مكتمل','مرفوض']:
        flash('حالة غير صالحة', 'err'); return redirect(url_for('cutting_home'))
    cuttings.update_status(cid, s)
    flash('تم التحديث', 'ok'); return redirect(url_for('cutting_home'))

@app.route('/cutting/reject', methods=['POST'])
@login_required
def cutting_reject():
    try:
        cid = int(request.form.get('id') or 0)
    except Exception:
        cid = 0
    reason = (request.form.get('reason') or '').strip()
    if not cid or not reason:
        flash('بيانات غير مكتملة', 'err'); return redirect(url_for('cutting_home'))
    cuttings.update_status(cid, 'مرفوض', reason)
    flash('تم الرفض', 'ok'); return redirect(url_for('cutting_home'))

@app.route('/cutting/delete/<int:cid>')
@login_required
def cutting_delete(cid):
    cuttings.delete(cid); flash('تم الحذف', 'ok'); return redirect(url_for('cutting_home'))

# --------------------------- ERROR HANDLING ----------------------------

def _fatal_box(title, exc):
    try:
        with open(ERROR_LOG, 'a', encoding='utf-8') as f:
            f.write(f"[{now_str()}] {title}: {type(exc).__name__}: {exc}\n")
            f.write(traceback.format_exc() + "\n")
    except Exception:
        pass


# ------------------------------ RUN -----------------------------------
if __name__ == '__main__':
    app.run(debug=True)
