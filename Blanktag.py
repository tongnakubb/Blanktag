import customtkinter
import tkinter
from tkinter import messagebox
import threading
import os
import qrcode
import os.path
import re
import json
import sys
import io 
from datetime import date
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.utils import ImageReader

# --- Helper Functions ---
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

NUMBER_SAVE_FILE = "last_number.txt"
CONFIG_FILE = "config.json"

# --- PDF Generation Function ---
def create_label_pdf(output_filename, start_num, end_num, prefix, digit_count, status_callback, completion_callback):
    try:
        status_callback("กำลังตรวจสอบฟอนต์...", None)
        font_path = resource_path("THSarabunNew.ttf") 
        if not os.path.exists(font_path): 
            status_callback("⚠️ ไม่พบฟอนต์ THSarabunNew.ttf", "#EBA403")
            completion_callback(False, 0)
            return
        
        pdfmetrics.registerFont(TTFont('ThaiFont', font_path))
        thai_font = 'ThaiFont'
        
        logo_path = resource_path("logo.png") 
        if not os.path.exists(logo_path): 
            status_callback("⚠️ ไม่พบไฟล์โลโก้ (logo.png)", "#EBA403")
            completion_callback(False, 0)
            return
        
        PAGE_WIDTH, PAGE_HEIGHT = A4
        c = canvas.Canvas(output_filename, pagesize=A4)
        margin_x = 10 * mm
        margin_y = 10 * mm
        gap_y = 5 * mm
        total_gap_space = 3 * gap_y
        label_width = PAGE_WIDTH - (2 * margin_x)
        label_height = (PAGE_HEIGHT - (2 * margin_y) - total_gap_space) / 4
        
        positions = [
            (margin_x, PAGE_HEIGHT-margin_y-label_height), 
            (margin_x, PAGE_HEIGHT-margin_y-(2*label_height)-gap_y), 
            (margin_x, PAGE_HEIGHT-margin_y-(3*label_height)-(2*gap_y)), 
            (margin_x, margin_y)
        ]
        
        current_date_str = date.today().strftime("%d-%b-%Y")
        logo_width = 30 * mm
        logo_height = 20 * mm
        
        label_count = 0
        total_labels = (end_num - start_num) + 1
        
        for i, num in enumerate(range(start_num, end_num + 1)):
            status_callback(f"กำลังสร้างป้ายที่ {i + 1}/{total_labels}...", None)
            

            number_part = f"{num:0{digit_count}d}"
            qr_data = f"{prefix}{number_part}"
            
            qr_img_obj = qrcode.make(qr_data)
            img_buffer = io.BytesIO()
            qr_img_obj.save(img_buffer, format='PNG')
            img_buffer.seek(0)
            qr_image_reader = ImageReader(img_buffer)
            
            pos_index = label_count % 4
            x_start, y_start = positions[pos_index]
            
            c.rect(x_start, y_start, label_width, label_height)
            
            qr_size = 20 * mm
            qr_x = x_start+label_width-qr_size-(5*mm)
            qr_y = y_start+(label_height-qr_size)/2
            c.drawImage(qr_image_reader, qr_x, qr_y, width=qr_size, height=qr_size, mask='auto')
            
            c.setFont("Helvetica-Bold", 14)
            c.drawCentredString(qr_x+qr_size/2, qr_y-3*mm, qr_data)
            
            text_x = x_start+10*mm
            line_start_x = text_x+15*mm
            line_end_x = qr_x-(10*mm)
            
            c.setFont(thai_font, 14)
            c.drawString(text_x, y_start+label_height-(20*mm), "Part:")
            c.line(line_start_x, y_start+label_height-(22*mm), line_end_x, y_start+label_height-(22*mm))
            
            c.drawString(text_x, y_start+label_height-(35*mm), "Qty:")
            c.line(line_start_x, y_start+label_height-(37*mm), line_end_x, y_start+label_height-(37*mm))
            
            c.drawString(text_x, y_start+label_height-(50*mm), "Name:")
            c.line(line_start_x, y_start+label_height-(52*mm), line_end_x, y_start+label_height-(52*mm))
            
            # Draw Logo
            logo_pos_x = x_start+label_width-logo_width-(1*mm)
            logo_pos_y = y_start+label_height-logo_height-(1*mm)
            c.drawImage(resource_path(logo_path), logo_pos_x, logo_pos_y, width=logo_width, height=logo_height, preserveAspectRatio=True, anchor='ne')
            
            # Draw Date
            c.setFont("Helvetica", 8)
            date_x_pos = x_start+label_width-(3*mm)
            date_y_pos = y_start+(3*mm)
            c.drawRightString(date_x_pos, date_y_pos, current_date_str)
            
            label_count += 1
            if label_count % 4 == 0 and num < end_num: 
                c.showPage()
                
        c.save()
        status_callback(f"✅ สร้างไฟล์สำเร็จ! กำลังเปิด...", "green")
        try:
            os.startfile(output_filename)
        except AttributeError:
            pass
        status_callback(f"✅ เปิดไฟล์ PDF แล้ว! กดพิมพ์ได้เลย", "green")
        completion_callback(True, end_num)
        
    except Exception as e:
        status_callback(f"❌ เกิดข้อผิดพลาด: {e}", "red")
        completion_callback(False, 0)

# --- Settings Window Class ---
class SettingsWindow(customtkinter.CTkToplevel):
    def __init__(self, master, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.iconbitmap(resource_path("logo1.ico")) 
        self.title("ตั้งค่าโปรแกรม")
        self.geometry("400x470")
        self.main_app = master
        self.grid_columnconfigure(0, weight=1)

        format_frame = customtkinter.CTkFrame(self, corner_radius=10)
        format_frame.grid(row=0, column=0, padx=20, pady=10, sticky="ew")
        format_frame.grid_columnconfigure(0, weight=1)
        
        format_label = customtkinter.CTkLabel(format_frame, text="Settings", font=customtkinter.CTkFont(size=14, weight="bold"))
        format_label.grid(row=0, column=0, padx=15, pady=(10,5))
        
        prefix_label = customtkinter.CTkLabel(format_frame, text="ตัวอักษรนำหน้า (Prefix):")
        prefix_label.grid(row=1, column=0, padx=20, pady=0, sticky="w")
        self.prefix_entry = customtkinter.CTkEntry(format_frame, placeholder_text="เช่น B")
        self.prefix_entry.grid(row=2, column=0, padx=20, pady=5, sticky="ew")
        
        digits_label = customtkinter.CTkLabel(format_frame, text="จำนวนหลักของตัวเลขทั้งหมด:")
        digits_label.grid(row=3, column=0, padx=20, pady=0, sticky="w")
        self.digits_entry = customtkinter.CTkEntry(format_frame, placeholder_text="เช่น 7")
        self.digits_entry.grid(row=4, column=0, padx=20, pady=5, sticky="ew")
        
        self.save_button = customtkinter.CTkButton(format_frame, text="บันทึกการตั้งค่ารูปแบบ", command=self.save_config)
        self.save_button.grid(row=5, column=0, padx=20, pady=(10,15))
        
        counter_frame = customtkinter.CTkFrame(self, corner_radius=10)
        counter_frame.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        counter_frame.grid_columnconfigure(0, weight=1)
        
        counter_label = customtkinter.CTkLabel(counter_frame, text="Counter", font=customtkinter.CTkFont(size=14, weight="bold"))
        counter_label.grid(row=0, column=0, padx=15, pady=(10,5))
        
        self.manual_set_label = customtkinter.CTkLabel(counter_frame, text="ตั้งค่าเลขล่าสุดด้วยตนเอง:")
        self.manual_set_label.grid(row=1, column=0, padx=20, pady=(10,0), sticky="w")
        self.manual_set_entry = customtkinter.CTkEntry(counter_frame, placeholder_text="เช่น 0000000")
        self.manual_set_entry.grid(row=2, column=0, padx=20, pady=5, sticky="ew")
        
        self.manual_set_button = customtkinter.CTkButton(counter_frame, text="บันทึกเลขใหม่", command=self.set_manual_counter)
        self.manual_set_button.grid(row=3, column=0, padx=20, pady=(5,10))
        
        self.reset_button = customtkinter.CTkButton(counter_frame, text="รีเซ็ต (เริ่มนับ 1 ใหม่)", command=self.reset_counter, fg_color="#D35400", hover_color="#A94442")
        self.reset_button.grid(row=4, column=0, padx=20, pady=(5, 15))
        
        self.prefix_entry.insert(0, self.main_app.config.get("prefix", ""))
        self.digits_entry.insert(0, str(self.main_app.config.get("digits", 7)))
        
        self.transient(self.master)
        self.grab_set()

    def set_manual_counter(self):
        try:
            new_num = int(self.manual_set_entry.get())
            if new_num < 0: return
            answer = messagebox.askyesno("ยืนยันการตั้งค่า", f"คุณแน่ใจหรือไม่ว่าต้องการเปลี่ยนเลขล่าสุดให้เป็น {new_num}?\nครั้งต่อไปจะเริ่มนับที่ {new_num + 1}")
            if answer: 
                self.main_app.set_main_counter(new_num)
                self.destroy()
        except (ValueError, TypeError): 
            messagebox.showerror("ผิดพลาด", "กรุณากรอกเป็นตัวเลขเท่านั้น")

    def reset_counter(self):
        answer = messagebox.askyesno("ยืนยันการรีเซ็ต", "คุณแน่ใจหรือไม่ว่าต้องการรีเซ็ตตัวนับเลข?\nค่าที่บันทึกไว้จะหายไปและจะเริ่มนับจาก 1 ใหม่")
        if answer: 
            self.main_app.reset_main_counter()
            self.destroy()

    def save_config(self):
        prefix = self.prefix_entry.get()
        try: 
            digits = int(self.digits_entry.get())
        except ValueError: 
            return
        
        self.main_app.config["prefix"] = prefix
        self.main_app.config["digits"] = digits
        self.main_app.save_config()
        self.main_app.update_display()
        messagebox.showinfo("สำเร็จ", "บันทึกการตั้งค่ารูปแบบเรียบร้อยแล้ว")

# --- Reprint Window Class (FIXED VERSION) ---
class ReprintWindow(customtkinter.CTkToplevel):
    def __init__(self, master, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.iconbitmap(resource_path("logo1.ico")) 
        self.title("Reprint")
        self.geometry("450x300")
        self.main_app = master
        self.grid_columnconfigure(0, weight=1)

        main_frame = customtkinter.CTkFrame(self, corner_radius=10)
        main_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        main_frame.grid_columnconfigure(0, weight=1)

        # แปลง digits ให้เป็น int เพื่อความชัวร์สำหรับการเติมเลข 0
        try:
            self.digits = int(self.main_app.config.get('digits', 7))
        except:
            self.digits = 7

        title_text = f"Reprint: {self.main_app.config.get('prefix', '')} + ตัวเลข {self.digits} หลัก"
        title_label = customtkinter.CTkLabel(main_frame, text=title_text, font=customtkinter.CTkFont(size=14, weight="bold"))
        title_label.grid(row=0, column=0, padx=20, pady=(15, 10))

        start_label = customtkinter.CTkLabel(main_frame, text="From BlankTag ที่ต้องการพิมพ์ซ้ำ:")
        start_label.grid(row=1, column=0, padx=20, pady=(10, 0), sticky="w")
        self.start_entry = customtkinter.CTkEntry(main_frame, placeholder_text=f"เช่น {1:0{self.digits}d}")
        self.start_entry.grid(row=2, column=0, padx=20, pady=5, sticky="ew")

        end_label = customtkinter.CTkLabel(main_frame, text="To BlankTag (ถ้าต้องการพิมพ์แค่ใบเดียวให้เว้นว่าง):")
        end_label.grid(row=3, column=0, padx=20, pady=(10, 0), sticky="w")
        self.end_entry = customtkinter.CTkEntry(main_frame, placeholder_text=f"เช่น {2:0{self.digits}d}")
        self.end_entry.grid(row=4, column=0, padx=20, pady=5, sticky="ew")

        self.generate_button = customtkinter.CTkButton(self, text="สร้างและเปิดเพื่อพิมพ์ซ้ำ", command=self.start_reprint_generation, corner_radius=8)
        self.generate_button.grid(row=1, column=0, padx=20, pady=10)
        
        self.status_label = customtkinter.CTkLabel(self, text="กรอกหมายเลขที่ต้องการพิมพ์ซ้ำ")
        self.status_label.grid(row=2, column=0, padx=20, pady=(0, 20))
        
        self.transient(self.master)
        self.grab_set()

    # --- ส่วนที่แก้ไขเพื่อป้องกัน Error Color is None ---
    def _internal_update_status(self, message, color):
        # ถ้า color มีค่า ให้ใส่สี / ถ้าไม่มี (None) ให้ใส่แค่ข้อความ
        if color:
            self.status_label.configure(text=message, text_color=color)
        else:
            self.status_label.configure(text=message) 

    def update_status_safe(self, message, color):
        # เรียกใช้ฟังก์ชันภายในผ่าน .after เพื่อความปลอดภัยของ Thread
        self.after(0, self._internal_update_status, message, color)
    # -----------------------------------------------

    def generation_completed_safe(self, success, last_num_generated):
        self.after(0, lambda: self._on_generation_finished(success))

    def _on_generation_finished(self, success):
        self.generate_button.configure(state="normal")
        if not success:
            self.status_label.configure(text="การสร้างไฟล์ล้มเหลว", text_color="red")

    def start_reprint_generation(self):
        prefix = self.main_app.config.get("prefix", "")
        try:
            self.digits = int(self.main_app.config.get("digits", 7))
        except ValueError:
            self.digits = 7

        try:
            start_str = self.start_entry.get()
            end_str = self.end_entry.get()
            
            if not start_str: 
                self.update_status_safe("❌ กรุณากรอกหมายเลขเริ่มต้น", "#EBA403")
                return
            
            start_num = int(start_str) 
            
            if not end_str: 
                end_num = start_num
            else: 
                end_num = int(end_str)
        except ValueError:
            self.update_status_safe("❌ กรุณากรอกเป็นตัวเลขเท่านั้น", "#EBA403")
            return

        if start_num > end_num:
            self.update_status_safe("❌ เลขเริ่มต้นต้องไม่มากกว่าเลขสิ้นสุด", "#EBA403")
            return

        self.generate_button.configure(state="disabled")
        
        safe_prefix = re.sub(r'[\\/*?:"<>|]', "_", prefix) 

        filename = f"Reprint_{safe_prefix}{start_num:0{self.digits}d}-{safe_prefix}{end_num:0{self.digits}d}.pdf"
        
        thread = threading.Thread(target=create_label_pdf, args=(filename, start_num, end_num, prefix, self.digits, self.update_status_safe, self.generation_completed_safe))
        thread.start()

# --- Main Application Class ---
class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        self.title("โปรแกรมสร้าง BlankTag")
        self.iconbitmap(resource_path("logo1.ico")) 
        self.geometry("563x380")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)
        
        self.load_config()
        
        top_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color="transparent")
        top_frame.grid(row=0, column=0, columnspan=2, padx=10, pady=10, sticky="ew")
        top_frame.grid_columnconfigure(0, weight=1)
        
        self.last_num_label = customtkinter.CTkLabel(top_frame, text="กำลังโหลด...", font=customtkinter.CTkFont(size=15, weight="bold"), wraplength=450, justify="left")
        self.last_num_label.grid(row=0, column=0, sticky="w", padx=10)
        
        self.theme_switch = customtkinter.CTkSwitch(top_frame, text="Dark Mode", command=lambda: customtkinter.set_appearance_mode("dark" if self.theme_switch.get() else "light"))
        self.theme_switch.grid(row=0, column=1, sticky="e")
        self.theme_switch.select()
        
        main_frame = customtkinter.CTkFrame(self, corner_radius=10)
        main_frame.grid(row=1, column=0, columnspan=2, padx=20, pady=10, sticky="ew")
        main_frame.grid_columnconfigure(0, weight=1)
        
        quantity_label = customtkinter.CTkLabel(main_frame, text="จำนวน BlankTag ที่ต้องการสร้าง (รันต่อจากเลขล่าสุด)")
        quantity_label.grid(row=2, column=0, padx=20, pady=(10, 0), sticky="w")
        
        self.quantity_entry = customtkinter.CTkEntry(main_frame, placeholder_text="เช่น 50", corner_radius=8)
        self.quantity_entry.grid(row=3, column=0, padx=20, pady=(5, 20), sticky="ew")
        
        bottom_frame = customtkinter.CTkFrame(self, corner_radius=10)
        bottom_frame.grid(row=2, column=0, columnspan=2, padx=20, pady=10, sticky="nsew")
        bottom_frame.grid_columnconfigure(0, weight=1)
        bottom_frame.grid_rowconfigure(1, weight=1)
        
        self.generate_button = customtkinter.CTkButton(bottom_frame, text="พิมพ์ BlankTag", command=self.start_generation, corner_radius=8, font=customtkinter.CTkFont(size=14, weight="bold"))
        self.generate_button.grid(row=0, column=0, padx=20, pady=15)
        
        self.status_label = customtkinter.CTkLabel(bottom_frame, text="กรอกข้อมูลแล้วกดปุ่ม", wraplength=400)
        self.status_label.grid(row=1, column=0, padx=20, pady=5, sticky="n")
        
        action_frame = customtkinter.CTkFrame(bottom_frame, fg_color="transparent")
        action_frame.grid(row=2, column=0, padx=20, pady=(10, 15))
        
        self.reprint_button = customtkinter.CTkButton(action_frame, text="Reprint", command=self.open_reprint_window, fg_color="transparent", border_width=1, width=120, text_color=("#1A1A1A", "#DCE4EE"))
        self.reprint_button.pack(side="left", padx=5)
        
        self.settings_button = customtkinter.CTkButton(action_frame, text="Settings", command=self.open_settings_window, fg_color="transparent", border_width=1, width=120, text_color=("#1A1A1A", "#DCE4EE"))
        self.settings_button.pack(side="left", padx=5)
        
        self.toplevel_window = None
        self.load_last_number()

    def reset_main_counter(self): 
        self.set_main_counter(0, "รีเซ็ตตัวนับเลขเรียบร้อยแล้ว")

    def set_main_counter(self, number_to_set, message="ตั้งค่าเลขล่าสุดเรียบร้อยแล้ว"): 
        self.save_last_number(number_to_set)
        self.update_status(message, "green")

    def open_settings_window(self):
        if self.toplevel_window is None or not self.toplevel_window.winfo_exists():
            self.toplevel_window = SettingsWindow(self)
            self.toplevel_window.iconbitmap(resource_path("logo1.ico")) 
        self.toplevel_window.focus()

    def open_reprint_window(self):
        if self.toplevel_window is None or not self.toplevel_window.winfo_exists():
            self.toplevel_window = ReprintWindow(self)
            self.toplevel_window.iconbitmap(resource_path("logo1.ico")) 
        self.toplevel_window.focus()

    def load_config(self):
        try:
            with open(CONFIG_FILE, 'r') as f: 
                self.config = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError): 
            self.config = {"prefix": "B", "digits": 7}

    def save_config(self):
        with open(CONFIG_FILE, 'w') as f: 
            json.dump(self.config, f, indent=4)

    def update_display(self):
        prefix = self.config.get("prefix", "")
        digits = self.config.get("digits", 7)
        last_num_str = f"{self.last_number:0{digits}d}"
        next_num_str = f"{self.last_number + 1:0{digits}d}"
        
        if self.last_number == 0: 
            self.last_num_label.configure(text=f"ยังไม่มีการบันทึก (จะเริ่มที่ {prefix}{next_num_str})")
        else: 
            self.last_num_label.configure(text=f"BlankTag ล่าสุด: {prefix}{last_num_str} (BlankTag ต่อไปเริ่มที่ {prefix}{next_num_str})")

    def load_last_number(self):
        try:
            if os.path.exists(NUMBER_SAVE_FILE):
                with open(NUMBER_SAVE_FILE, 'r') as f: 
                    last_num = int(f.read().strip())
            else: 
                last_num = 0
            self.last_number = last_num
        except (ValueError, FileNotFoundError): 
            self.last_number = 0
        self.update_display()

    def save_last_number(self, num_to_save):
        with open(NUMBER_SAVE_FILE, 'w') as f: 
            f.write(str(num_to_save))
        self.load_last_number()

    def generation_completed(self, success, last_num_generated):
        if success: 
            self.save_last_number(last_num_generated)
        self.generate_button.configure(state="normal")

    def _internal_update_status(self, message, color):
        color_map = {"green": "#2ECC71", "red": "#E74C3C", "#EBA403": "#EBA403"}
        text_color = color_map.get(color, customtkinter.ThemeManager.theme["CTkLabel"]["text_color"])
        self.status_label.configure(text=message, text_color=text_color)

    def update_status(self, message, color):
        self.after(0, lambda: self._internal_update_status(message, color))

    def start_generation(self):
        prefix = self.config.get("prefix", "")
        digits = self.config.get("digits", 7)
        try: 
            quantity = int(self.quantity_entry.get())
        except ValueError: 
            self.update_status("❌ กรุณากรอกจำนวนเป็นตัวเลข", "#EBA403")
            return
            
        if quantity <= 0: 
            self.update_status("❌ จำนวนต้องมากกว่า 0", "#EBA403")
            return
            
        start_num = self.last_number + 1
        end_num = start_num + quantity - 1
        
        self.generate_button.configure(state="disabled")
        thread = threading.Thread(target=create_label_pdf, args=("labels.pdf", start_num, end_num, prefix, digits, self.update_status, self.generation_completed))
        thread.start()

if __name__ == "__main__":
    app = App()
    app.mainloop()