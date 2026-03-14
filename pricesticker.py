import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
import os

class LemonLabelGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("Lemon Pharmacy Label Maker")
        self.root.geometry("400x250")
        self.root.configure(bg="#f0f0f0")

        self.label = tk.Label(root, text="Select Excel File to Generate Labels", 
                              font=("Arial", 12), bg="#f0f0f0", pady=20)
        self.label.pack()

        self.btn_select = tk.Button(root, text="Browse Excel File", 
                                    command=self.process_file, 
                                    bg="#2ecc71", fg="white", 
                                    font=("Arial", 10, "bold"), 
                                    padx=20, pady=10)
        self.btn_select.pack(pady=10)

        self.status = tk.Label(root, text="Status: Ready", bg="#f0f0f0", fg="gray")
        self.status.pack(side="bottom", pady=10)

    def process_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not file_path:
            return

        try:
            # قراءة الإكسيل (تأكد من ترتيب الأعمدة كما ذكرت)
            df = pd.read_excel(file_path)
            
            output_name = "Lemon_Labels_Output.pdf"
            # إعداد الكانفاس بابعاد A4 (أو يمكنك تخصيص الحجم إذا كانت ملصقات صغيرة)
            c = canvas.Canvas(output_name, pagesize=A4)
            width, height = A4

            for index, row in df.iterrows():
                # استلام البيانات من الأعمدة بالترتيب
                intl_code = str(row.iloc[0])   # العمود 1: الكود الدولي
                ascon_code = str(row.iloc[1])  # العمود 2: كود ASCON
                item_name = str(row.iloc[2])   # العمود 3: الاسم الإنجليزي
                tax_val = float(row.iloc[3])   # العمود 4: نسبة الضريبة (مثلاً 15)
                price_raw = float(row.iloc[4]) # العمود 5: السعر قبل الضريبة

                # حساب السعر النهائي
                if tax_val > 0:
                    final_price = price_raw + (price_raw * (tax_val / 100))
                else:
                    final_price = price_raw
                
                price_text = "{:.2f}".format(final_price)

                # --- رسم التصميم (محاكاة للملف المرفق) ---
                
                # 1. Header
                c.setFont("Helvetica-Bold", 18)
                c.drawCentredString(width/2, height - 40*mm, "Lemon Pharmacy Group")

                # 2. Price (الكبير في المنتصف)
                c.setFont("Helvetica-Bold", 60)
                c.drawCentredString(width/2, height - 70*mm, price_text)

                # 3. Tax info & Currency
                c.setFont("Helvetica-Bold", 14)
                tax_label = f"S.R. {int(tax_val)}% VAT"
                c.drawCentredString(width/2, height - 85*mm, tax_label)

                # 4. Codes (International on Left, ASCON on Right)
                c.setFont("Helvetica", 12)
                c.drawString(40*mm, height - 100*mm, intl_code)
                c.drawRightString(width - 40*mm, height - 100*mm, ascon_code)

                # 5. Item Name (English Only)
                c.setFont("Helvetica-Bold", 16)
                # عمل Wrap بسيط للنصوص الطويلة
                c.drawCentredString(width/2, height - 120*mm, item_name[:40]) 
                if len(item_name) > 40:
                    c.drawCentredString(width/2, height - 128*mm, item_name[40:80])

                # 6. Footer (Website)
                c.setFont("Helvetica", 12)
                c.drawCentredString(width/2, height - 150*mm, "www.lemon.sa")

                # إنهاء الصفحة الحالية والبدء في صفحة جديدة للصنف التالي
                c.showPage()

            c.save()
            messagebox.showinfo("Success", f"PDF Generated successfully:\n{output_name}")
            self.status.config(text="Status: Completed", fg="green")
            
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            self.status.config(text="Status: Error", fg="red")

if __name__ == "__main__":
    root = tk.Tk()
    app = LemonLabelGenerator(root)
    root.mainloop()
