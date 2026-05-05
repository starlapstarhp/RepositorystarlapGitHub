import os
import platform
import sqlite3
import pandas as pd
import flet as ft
from flet import *
from reportlab.lib import colors
from reportlab.platypus import TableStyle, SimpleDocTemplate, Table, Paragraph
from reportlab.lib.units import mm
from reportlab.lib.pagesizes import landscape, letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.styles import getSampleStyleSheet
import arabic_reshaper
from bidi.algorithm import get_display

# --- 1. إعداد المسارات المتوافقة مع أندرويد ---
data_dir = os.environ.get("FLET_APP_STORAGE_DATA", os.getcwd())
db_path = os.path.join(data_dir, "company.db")

# --- 2. كلاس إدارة قاعدة البيانات ---
class DatabaseHelper:
    def __init__(self):
        self.conn = sqlite3.connect(db_path, check_same_thread=False)
        self.cursor = self.conn.cursor()
        self.create_table()

    def create_table(self):
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS employees (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                job_title TEXT NOT NULL,
                salary REAL NOT NULL,
                is_active INTEGER DEFAULT 1 
            )
        """)
        self.conn.commit()

    def add_employee(self, name, job, salary):
        self.cursor.execute("INSERT INTO employees (name, job_title, salary, is_active) VALUES (?, ?, ?, 1)", (name, job, salary))
        self.conn.commit()

    def get_all_employees(self):
        self.cursor.execute("SELECT * FROM employees ORDER BY id DESC")
        return self.cursor.fetchall()

    def delete_employee(self, emp_id):
        self.cursor.execute("DELETE FROM employees WHERE id = ?", (emp_id,))
        self.conn.commit()

    def update_status__employee(self, emp_id, Checkbox_value):
        self.cursor.execute("UPDATE employees SET is_active = ? WHERE id = ?", (int(Checkbox_value), emp_id))
        self.conn.commit()

# --- 3. أدوات مساعدة للخط العربي ---
try:
    # تأكد من وضع ملف arial.ttf بجانب ملف الكود
    pdfmetrics.registerFont(TTFont('ArabicFont', 'arial.ttf'))
    pdfmetrics.registerFont(TTFont('ArabicFont-Bold', 'arial.ttf')) 
except:
    print("تنبيه: لم يتم العثور على ملف الخط")

def fix_arabic(text):
    reshaped_text = arabic_reshaper.reshape(str(text))
    return get_display(reshaped_text)

# --- 4. واجهة المستخدم الرئيسية ---
def main(page: Page):
    page.title = "نظام إدارة الموظفين"
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
    page.scroll = ft.ScrollMode.AUTO
    page.rtl = True # تفعيل دعم العربية في الواجهة

    db = DatabaseHelper()

    # حقول الإدخال
    name_field = ft.TextField(label="اسم الموظف", expand=True)
    job_field = ft.TextField(label="المسمى الوظيفي", expand=True)
    salary_field = ft.TextField(label="الراتب", keyboard_type=ft.KeyboardType.NUMBER, expand=True)

    my_dt = ft.DataTable(
        columns=[
            ft.DataColumn(Text("الرقم")),
            ft.DataColumn(Text("الاسم")),
            ft.DataColumn(Text("المسمى")),
            ft.DataColumn(Text("الراتب")),
            ft.DataColumn(Text("حذف")),
            ft.DataColumn(Text("الحالة")),
        ],
        rows=[]
    )

    def load_data():
        my_t_conn = sqlite3.connect(db_path)
        my_t_conn.row_factory = sqlite3.Row 
        cursor = my_t_conn.cursor()
        cursor.execute("SELECT * FROM employees")
        rows_data = cursor.fetchall()
        
        my_dt.rows.clear()
        for x in rows_data:
            my_dt.rows.append(DataRow(
                cells=[
                    DataCell(Text(x['id'])),
                    DataCell(Text(x['name'])),
                    DataCell(Text(x['job_title'])),
                    DataCell(Text(x['salary'])),
                    DataCell(ft.IconButton(icon=icons.Icons.DELETE, icon_color="red", on_click=lambda e, id=x['id']: delete_clicked(id))),
                    DataCell(Checkbox(value=bool(x['is_active']), on_change=lambda e, id=x['id']: update_status(id, e.control.value)))
                ]
            ))
        my_t_conn.close()
        page.update()

    def add_clicked(e):
        if name_field.value and job_field.value and salary_field.value:
            try:
                db.add_employee(name_field.value, job_field.value, float(salary_field.value))
                name_field.value = ""; job_field.value = ""; salary_field.value = ""
                load_data()
            except ValueError:
                page.snack_bar = SnackBar(Text("خطأ في القيمة الرقمية")); page.snack_bar.open = True
        page.update()

    def delete_clicked(emp_id):
        db.delete_employee(emp_id)
        load_data()

    def update_status(emp_id, value):
        db.update_status__employee(emp_id, value)
        load_data()

    def printpdfnow(e):
        output_path = os.path.join(data_dir, "export_arabic.pdf")
        doc = SimpleDocTemplate(output_path, pagesize=landscape(letter))
        data = [[fix_arabic("الرقم"), fix_arabic("الاسم"), fix_arabic("المسمى"), fix_arabic("الراتب")]]
        
        conn = sqlite3.connect(db_path)
        conn.row_factory = sqlite3.Row
        rows = conn.execute("SELECT * FROM employees WHERE is_active = 1").fetchall()
        for item in rows:
            data.append([fix_arabic(item['id']), fix_arabic(item['name']), fix_arabic(item['job_title']), fix_arabic(item['salary'])])
        
        table = Table(data, colWidths=[30*mm, 50*mm, 50*mm, 30*mm])
        style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#FCF8BE")),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, -1), 'ArabicFont'),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ])
        table.setStyle(style)
        doc.build([Paragraph(fix_arabic("تقرير الموظفين"), getSampleStyleSheet()['Normal']), table])

        # فتح الملف
        if platform.system() == "Windows": os.startfile(output_path)
        else: page.launch_url(f"file://{output_path}")

    def export_to_excel_3(e):
        output_path = os.path.join(data_dir, "export_employees.xlsx")
        conn = sqlite3.connect(db_path)
        df = pd.read_sql_query("SELECT id, name, job_title, salary FROM employees WHERE is_active = 1", conn)
        df.columns = ['الرقم', 'الاسم', 'المسمى الوظيفي', 'الراتب']
        writer = pd.ExcelWriter(output_path, engine='xlsxwriter')
        df.to_excel(writer, index=False, sheet_name='الموظفين')
        writer.close()
        conn.close()
        if platform.system() == "Windows": os.startfile(output_path)
        else: page.launch_url(f"file://{output_path}")

    # بناء الواجهة
    page.add(
        ft.Column([
            ft.Row([name_field, job_field, salary_field]),
            ft.ElevatedButton("حفظ الموظف", icon=icons.Icons.SAVE, on_click=add_clicked),
            ft.Divider(),
            ft.Row([
                ft.Button("طباعة PDF", icon="print", on_click=printpdfnow),
                ft.Button("تصدير Excel", icon="table_chart", on_click=export_to_excel_3)
            ], alignment=MainAxisAlignment.CENTER),
            ft.Divider(),
            my_dt
        ])
    )
    load_data()

ft.app(main)

