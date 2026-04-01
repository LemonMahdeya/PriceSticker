import os
import sys
import win32print
import win32ui
from PIL import Image, ImageWin, ImageDraw, ImageFont

def print_fridge_label():
    try:
        # 1. إعدادات المسار (للتوافق مع ملف EXE)
        if getattr(sys, 'frozen', False):
            base_path = os.path.dirname(sys.executable)
        else:
            base_path = os.path.dirname(__file__)
        
        image_path = os.path.join(base_path, "snow.png")

        if not os.path.exists(image_path):
            print("خطأ: لم يتم العثور على صورة snow.png في المجلد")
            return

        # 2. إعدادات الورقة (80mm تقريباً تعادل 575 بكسل بـ 203 DPI)
        width = 575 
        height = 600 # طول مرن حسب الحاجة
        img = Image.new('RGB', (width, height), color=(255, 255, 255))
        draw = ImageDraw.Draw(img)

        # 3. كتابة النص FRIDGE ITEM
        # ملاحظة: يمكنك تحميل خط عريض إذا أردت، هنا نستخدم الخط الافتراضي
        try:
            font = ImageFont.truetype("arial.ttf", 60)
        except:
            font = ImageFont.load_default()

        text = "FRIDGE ITEM"
        # توسيط النص
        bbox = draw.textbbox((0, 0), text, font=font)
        text_width = bbox[2] - bbox[0]
        draw.text(((width - text_width) // 2, 40), text, fill=(0, 0, 0), font=font)

        # 4. إضافة صورة الثلج وتكبيرها
        snow_img = Image.open(image_path).convert("RGBA")
        # جعل عرض الصورة 400 بكسل مع الحفاظ على التناسب
        ratio = 400 / float(snow_img.size[0])
        new_height = int(float(snow_img.size[1]) * float(ratio))
        snow_img = snow_img.resize((400, new_height), Image.Resampling.LANCZOS)
        
        # دمج صورة الثلج في منتصف الورقة
        img.paste(snow_img, ((width - 400) // 2, 150), snow_img)

        # 5. إرسال الأمر للطابعة الافتراضية
        printer_name = win32print.GetDefaultPrinter()
        hDC = win32ui.CreateDC()
        hDC.CreatePrinterDC(printer_name)
        
        hDC.StartDoc("Fridge Label")
        hDC.StartPage()

        dib = ImageWin.Dib(img)
        dib.draw(hDC.GetHandleOutput(), (0, 0, width, height))

        hDC.EndPage()
        hDC.EndDoc()
        hDC.DeleteDC()
        
        print("تم إرسال الطلب للطابعة بنجاح.")

    except Exception as e:
        print(f"حدث خطأ: {e}")

if __name__ == "__main__":
    print_fridge_label()
    # البرنامج سيغلق تلقائياً بعد انتهاء الوظيفة
