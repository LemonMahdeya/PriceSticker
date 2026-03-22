import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib.utils import simpleSplit


class LemonLabelGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("Lemon Sticker Maker 60x40")
        self.root.geometry("400x250")

        tk.Label(root, text="Select Excel File to Generate Stickers", pady=20).pack()

        tk.Button(
            root,
            text="Generate Stickers (60x40mm)",
            command=self.process_file,
            bg="#2ecc71",
            fg="white",
            padx=20,
            pady=10
        ).pack(pady=10)

    def safe_str(self, value):
        """تحويل آمن إلى string بدون nan"""
        if pd.isna(value):
            return ""
        return str(value).strip()

    def safe_float(self, value):
        """تحويل آمن إلى رقم"""
        try:
            return float(value)
        except:
            return 0.0

    def process_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if not file_path:
            return

        try:
            df = pd.read_excel(file_path).dropna(how='all')

            if df.shape[1] < 5:
                raise Exception("Excel file must have at least 5 columns")

            output_name = "Lemon_Stickers_Final.pdf"

            sw, sh = 60 * mm, 40 * mm
            c = canvas.Canvas(output_name, pagesize=(sw, sh))

            for _, row in df.iterrows():

                # ✅ Intl Code
                intl_code_raw = row.iloc[0]
                try:
                    intl_code = str(int(intl_code_raw)) if pd.notna(intl_code_raw) else ""
                except:
                    intl_code = ""

                # ✅ ASCON Code
                ascon_code = self.safe_str(row.iloc[1])
                if ascon_code and not ascon_code.startswith("0"):
                    ascon_code = "0" + ascon_code

                # ✅ باقي البيانات
                item_name = self.safe_str(row.iloc[2])
                tax_percent = self.safe_float(row.iloc[3])
                price_input = self.safe_float(row.iloc[4])

                price_display = "{:.2f}".format(price_input)

                # ------------------ الرسم ------------------

                # الاسم
                c.setFont("Helvetica-Bold", 8)
                name_lines = simpleSplit(item_name, "Helvetica-Bold", 8, sw - 12 * mm)
                y_pos = sh - 5 * mm
                for line in name_lines[:2]:
                    c.drawString(4 * mm, y_pos, line)
                    y_pos -= 3.2 * mm

                # السعر
                c.setFont("Helvetica-Bold", 32)
                p_width = c.stringWidth(price_display, "Helvetica-Bold", 32)
                start_x = (sw - p_width) / 2
                c.drawString(start_x - 2 * mm, sh / 2 + 1 * mm, price_display)

                # S.R
                c.setFont("Helvetica", 8)
                sr_x = start_x + p_width + 1 * mm
                sr_y = sh / 2 + 1 * mm
                c.drawString(sr_x, sr_y, "S.R")

                # VAT
                c.setFont("Helvetica", 6)
                c.drawString(sr_x, sr_y + 4 * mm, f"{int(tax_percent)}% VAT")

                # الأكواد
                c.setFont("Helvetica-Bold", 10)
                c.drawCentredString(
                    sw / 2,
                    sh / 2 - 9 * mm,
                    f"{intl_code}   /   {ascon_code}"
                )

                # الفوتر
                c.setFont("Helvetica-Bold", 8)
                c.drawCentredString(
                    sw / 2,
                    6 * mm,
                    "Lemon pharmacy Group       www.lemon.sa"
                )

                c.showPage()

            c.save()
            messagebox.showinfo("Success", "Done! Check Lemon_Stickers_Final.pdf")

        except Exception as e:
            messagebox.showerror("Error", f"Error: {str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = LemonLabelGenerator(root)
    root.mainloop()
