import os
import subprocess
import tempfile
import threading
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox
from PyPDF2 import PdfReader, PdfWriter
from plyer import notification

# ==== إعداد مسار Adobe Reader ====
sumatra_path = "C:/Tools/SumatraPDF/SumatraPDF.exe"

# ==== الوظيفة الأساسية ====

def select_file():
    filepath = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if filepath:
        file_entry.delete(0, END)
        file_entry.insert(0, filepath)

def start_print_thread():
    thread = threading.Thread(target=print_pdf)
    thread.start()

def print_pdf():
    file_path = file_entry.get().strip()
    pages_input = pages_entry.get().strip()
    copies_input = copies_entry.get().strip()

    if not os.path.isfile(file_path):
        messagebox.showerror("خطأ", "الملف غير موجود!")
        return

    if not pages_input or not copies_input:
        messagebox.showerror("خطأ", "رجاءً أدخل الصفحات وعدد النسخ.")
        return

    try:
        copies = int(copies_input)
    except ValueError:
        messagebox.showerror("خطأ", "عدد النسخ يجب أن يكون رقمًا صحيحًا.")
        return

    # معالجة الصفحات المطلوبة
    pages_to_print = []
    try:
        if "-" in pages_input:
            start, end = map(int, pages_input.split("-"))
            pages_to_print = list(range(start - 1, end))
        else:
            pages_to_print = [int(p) - 1 for p in pages_input.split(",")]
    except Exception:
        messagebox.showerror("خطأ", "تنسيق الصفحات غير صحيح! (مثلا 1-3 أو 1,2,5)")
        return

    # إنشاء ملف PDF مؤقت
    try:
        reader = PdfReader(file_path)
        writer = PdfWriter()

        for page_num in pages_to_print:
            if 0 <= page_num < len(reader.pages):
                writer.add_page(reader.pages[page_num])
            else:
                messagebox.showwarning("تحذير", f"الصفحة {page_num + 1} غير موجودة!")

        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
            temp_pdf_path = temp_pdf.name
            writer.write(temp_pdf)

        # إرسال الملف للطباعة
        for i in range(copies):
            subprocess.run([
               sumatra_path,
               "-print-to-default",
               temp_pdf_path
            ])
        
        # عرض إشعار بعد الطباعة
        notification.notify(
            title="تمت الطباعة ✅",
            message=f"تم طباعة {copies} نسخة بنجاح!",
            app_name = "Slatuna",
            timeout=5  # بالثواني
        )

    except Exception as e:
        messagebox.showerror("خطأ", f"حدث خطأ أثناء المعالجة: {e}")
    
    finally:
        try:
            os.remove(temp_pdf_path)
        except:
            pass

# ==== بناء الواجهة ====

app = ttk.Window(themename="solar")
app.title("طباعة PDF محدد")
app.geometry("450x350")

frame = ttk.Frame(app, padding=20)
frame.pack(fill=BOTH, expand=YES)

ttk.Label(frame, text="اختر ملف PDF:", font=("Arial", 12)).pack(anchor=W, pady=(0, 5))

file_frame = ttk.Frame(frame)
file_frame.pack(fill=X, pady=5)

file_entry = ttk.Entry(file_frame, width=40)
file_entry.pack(side=LEFT,expand=YES, fill=X, padx=(0, 5))

browse_button = ttk.Button(file_frame, text="استعراض", command=select_file, bootstyle=PRIMARY)
browse_button.pack()

ttk.Label(frame, text="الصفحات المطلوبة (مثلا 1-3 أو 1,2,5):", font=("Arial", 12)).pack(anchor=W, pady=(15, 5))
pages_entry = ttk.Entry(frame)
pages_entry.pack(fill=X, pady=5)

ttk.Label(frame, text="عدد النسخ:", font=("Arial", 12)).pack(anchor=W, pady=(15, 5))
copies_entry = ttk.Entry(frame)
copies_entry.pack(fill=X, pady=5)

print_button = ttk.Button(frame, text="طباعة", command=start_print_thread, bootstyle=SUCCESS)
print_button.pack(pady=20, fill=X)

app.mainloop()