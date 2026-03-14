import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib.utils import simpleSplit
import os

class LemonLabelGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("Lemon Sticker Maker 60x40")
        self.root.geometry("400x250")

        self.label = tk.Label(root, text="Select Excel File to Generate Stickers", pady=20)
        self.label.pack()

        self.btn_select = tk.Button(root, text="Generate Stickers (60x40mm)", 
                                    command=self.process_file, 
                                    bg="#2ecc71", fg="white", padx=20, pady=10)
        self.btn_select.pack(pady=10)

    def process_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not file_path:
            return

        try:
            df = pd.read_excel(file_path).dropna(how='all')
            output_name = "Lemon_Stickers_Final.pdf"
            
            # مقاس الاستيكر 60 مم عرض × 40 مم طول
            sw = 60 * mm
            sh = 40 * mm
            
            c = canvas.Canvas(output_name, pagesize=(sw, sh))

            for index, row in df.iterrows():
                # استلام البيانات
                intl_code = str(row.iloc[0])
                ascon_code = str(row.iloc[1])
                item_name = str(row.iloc[2])
                tax_percent = float(row.iloc[3])
                price_base = float(row.iloc[4])

                # حساب السعر النهائي: السعر + (السعر * نسبة الضريبة)
                final_price = price_base + (price_base * (tax_percent / 100))
                price_num_str = "{:.2f}".format(final_price)

                # --- رسم التصميم ---
                
                # 1. الاسم بالإنجليزي (فوق يسار مع خاصية التفاف النص)
                c.setFont("Helvetica-Bold", 8)
                lines = simpleSplit(item_name, "Helvetica-Bold", 8, sw - 10*mm)
                y_text = sh - 6*mm
                for line in lines[:2]: # السماح بسطرين فقط للاسم
                    c.drawString(4*mm, y_text, line)
                    y_text -= 3.5*mm

                # 2. السعر في المنتصف (الرقم كبير)
                c.setFont("Helvetica-Bold", 28)
                num_width = c.stringWidth(price_num_str, "Helvetica-Bold", 28)
                c.drawString((sw - num_width)/2 - 4*mm, sh/2 + 1*mm, price_num_str)
                
                # كلمة S.R بجانب السعر (خط صغير مثل الـ VAT)
                c.setFont("Helvetica", 7)
                c.drawString((sw - num_width)/2 + num_width - 3*mm, sh/2 + 1*mm, "S.R")

                # 3. الضريبة (في الوسط على اليمين - خط صغير)
                c.setFont("Helvetica", 7)
                tax_str = f"{int(tax_percent)}% VAT"
                c.drawRightString(sw - 4*mm, sh/2 + 1*mm, tax_str)

                # 4. الأكواد (INTERNATIONAL CODE / ASCON CODE) - خط أكبر و Bold
                c.setFont("Helvetica-Bold", 10)
                codes_text = f"{intl_code}   /   {ascon_code}"
                c.drawCentredString(sw/2, sh/2 - 9*mm, codes_text)

                # 5. التذييل (اسم المجموعة والموقع)
                c.setFont("Helvetica-Bold", 8)
                footer_text = "Lemon pharmacy Group       www.lemon.sa"
                c.drawCentredString(sw/2, 6*mm, footer_text)

                c.showPage()

            c.save()
            messagebox.showinfo("Success", "PDF Stickers Created Successfully!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = LemonLabelGenerator(root)
    root.mainloop()
