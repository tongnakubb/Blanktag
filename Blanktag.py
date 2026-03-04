import customtkinter
import tkinter
from tkinter import messagebox
import threading
import os
import qrcode
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


def get_app_dir():
    """ ล็อก Path ปัจจุบันให้อยู่ที่เดียวกับไฟล์ .exe หรือ .py เสมอ ป้องกันไฟล์เซฟผิดที่ """
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


CONFIG_FILE = os.path.join(get_app_dir(), "config.json")


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
            

            logo_pos_x = x_start+label_width-logo_width-(1*mm)
            logo_pos_y = y_start+label_height-logo_height-(1*mm)
            c.drawImage(resource_path(logo_path), logo_pos_x, logo_pos_y, width=logo_width, height=logo_height, preserveAspectRatio=True, anchor='ne')
            

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

class SettingsWindow(customtkinter.CTkToplevel):
    def __init__(self, generator_frame, mode_name, *args, **kwargs):
        super().__init__(generator_frame.main_app, *args, **kwargs)
        try:
            self.iconbitmap(resource_path("logo1.ico")) 
        except:
            pass
        self.title(f"ตั้งค่าโปรแกรม ({mode_name})")
        self.geometry("400x470")
        self.generator_frame = generator_frame
        self.main_app = generator_frame.main_app
        self.mode_name = mode_name
        self.grid_columnconfigure(0, weight=1)

        format_frame = customtkinter.CTkFrame(self, corner_radius=10)
        format_frame.grid(row=0, column=0, padx=20, pady=10, sticky="ew")
        format_frame.grid_columnconfigure(0, weight=1)
        
        format_label = customtkinter.CTkLabel(format_frame, text=f"Settings - {mode_name}", font=customtkinter.CTkFont(size=14, weight="bold"))
        format_label.grid(row=0, column=0, padx=15, pady=(10,5))
        
        prefix_label = customtkinter.CTkLabel(format_frame, text="ตัวอักษรนำหน้า (Prefix):")
        prefix_label.grid(row=1, column=0, padx=20, pady=0, sticky="w")
        self.prefix_entry = customtkinter.CTkEntry(format_frame, placeholder_text="เช่น B หรือ NC")
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
        
        mode_config = self.main_app.config.get(self.mode_name, {})
        self.prefix_entry.insert(0, mode_config.get("prefix", ""))
        self.digits_entry.insert(0, str(mode_config.get("digits", 7)))
        
        self.transient(self.master)
        self.grab_set()

    def set_manual_counter(self):
        try:
            new_num = int(self.manual_set_entry.get())
            if new_num < 0: return
            answer = messagebox.askyesno("ยืนยันการตั้งค่า", f"คุณแน่ใจหรือไม่ว่าต้องการเปลี่ยนเลขล่าสุดให้เป็น {new_num}?\nครั้งต่อไปจะเริ่มนับที่ {new_num + 1}")
            if answer: 
                self.generator_frame.set_main_counter(new_num)
                self.destroy()
        except (ValueError, TypeError): 
            messagebox.showerror("ผิดพลาด", "กรุณากรอกเป็นตัวเลขเท่านั้น")

    def reset_counter(self):
        answer = messagebox.askyesno("ยืนยันการรีเซ็ต", "คุณแน่ใจหรือไม่ว่าต้องการรีเซ็ตตัวนับเลข?\nค่าที่บันทึกไว้จะหายไปและจะเริ่มนับจาก 1 ใหม่")
        if answer: 
            self.generator_frame.reset_main_counter()
            self.destroy()

    def save_config(self):
        prefix = self.prefix_entry.get()
        try: 
            digits = int(self.digits_entry.get())
        except ValueError: 
            return
            
        if self.mode_name not in self.main_app.config:
            self.main_app.config[self.mode_name] = {}
            
        self.main_app.config[self.mode_name]["prefix"] = prefix
        self.main_app.config[self.mode_name]["digits"] = digits
        self.main_app.save_config()
        self.generator_frame.update_display()
        messagebox.showinfo("สำเร็จ", f"บันทึกการตั้งค่าสำหรับ {self.mode_name} เรียบร้อยแล้ว")

class ReprintWindow(customtkinter.CTkToplevel):
    def __init__(self, generator_frame, mode_name, *args, **kwargs):
        super().__init__(generator_frame.main_app, *args, **kwargs)
        try:
            self.iconbitmap(resource_path("logo1.ico")) 
        except:
            pass
        self.title(f"Reprint ({mode_name})")
        self.geometry("450x300")
        self.generator_frame = generator_frame
        self.main_app = generator_frame.main_app
        self.mode_name = mode_name
        self.grid_columnconfigure(0, weight=1)

        main_frame = customtkinter.CTkFrame(self, corner_radius=10)
        main_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        main_frame.grid_columnconfigure(0, weight=1)

        mode_config = self.main_app.config.get(self.mode_name, {})
        try:
            self.digits = int(mode_config.get('digits', 7))
        except:
            self.digits = 7

        title_text = f"Reprint: {mode_config.get('prefix', '')} + ตัวเลข {self.digits} หลัก"
        title_label = customtkinter.CTkLabel(main_frame, text=title_text, font=customtkinter.CTkFont(size=14, weight="bold"))
        title_label.grid(row=0, column=0, padx=20, pady=(15, 10))

        start_label = customtkinter.CTkLabel(main_frame, text="From Tag ที่ต้องการพิมพ์ซ้ำ:")
        start_label.grid(row=1, column=0, padx=20, pady=(10, 0), sticky="w")
        self.start_entry = customtkinter.CTkEntry(main_frame, placeholder_text=f"เช่น {1:0{self.digits}d}")
        self.start_entry.grid(row=2, column=0, padx=20, pady=5, sticky="ew")

        end_label = customtkinter.CTkLabel(main_frame, text="To Tag (ถ้าต้องการพิมพ์แค่ใบเดียวให้เว้นว่าง):")
        end_label.grid(row=3, column=0, padx=20, pady=(10, 0), sticky="w")
        self.end_entry = customtkinter.CTkEntry(main_frame, placeholder_text=f"เช่น {2:0{self.digits}d}")
        self.end_entry.grid(row=4, column=0, padx=20, pady=5, sticky="ew")

        btn_color = "#D35400" if mode_name == "NC Tag" else ["#3a7ebf", "#1f538d"]
        hover_color = "#E67E22" if mode_name == "NC Tag" else ["#325882", "#14375e"]

        self.generate_button = customtkinter.CTkButton(self, text="สร้างและเปิดเพื่อพิมพ์ซ้ำ", command=self.start_reprint_generation, corner_radius=8, fg_color=btn_color, hover_color=hover_color)
        self.generate_button.grid(row=1, column=0, padx=20, pady=10)
        
        self.status_label = customtkinter.CTkLabel(self, text="กรอกหมายเลขที่ต้องการพิมพ์ซ้ำ")
        self.status_label.grid(row=2, column=0, padx=20, pady=(0, 20))
        
        self.transient(self.master)
        self.grab_set()

    def _internal_update_status(self, message, color):
        if color:
            self.status_label.configure(text=message, text_color=color)
        else:
            self.status_label.configure(text=message) 

    def update_status_safe(self, message, color):
        self.after(0, self._internal_update_status, message, color)

    def generation_completed_safe(self, success, last_num_generated):
        self.after(0, lambda: self._on_generation_finished(success))

    def _on_generation_finished(self, success):
        self.generate_button.configure(state="normal")
        if not success:
            self.status_label.configure(text="การสร้างไฟล์ล้มเหลว", text_color="red")

    def start_reprint_generation(self):
        mode_config = self.main_app.config.get(self.mode_name, {})
        prefix = mode_config.get("prefix", "")
        
        try:
            self.digits = int(mode_config.get("digits", 7))
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

        filename = os.path.join(get_app_dir(), f"Reprint_{self.mode_name.replace(' ', '')}_{safe_prefix}{start_num:0{self.digits}d}-{safe_prefix}{end_num:0{self.digits}d}.pdf")
        
        thread = threading.Thread(target=create_label_pdf, args=(filename, start_num, end_num, prefix, self.digits, self.update_status_safe, self.generation_completed_safe))
        thread.start()

class GeneratorFrame(customtkinter.CTkFrame):
    def __init__(self, master, main_app, mode_name, *args, **kwargs):
        super().__init__(master, corner_radius=10, *args, **kwargs)
        self.main_app = main_app
        self.mode_name = mode_name
        self.toplevel_window = None
        self.grid_columnconfigure(0, weight=1)
        
        title_color = "#3498DB" if mode_name == "BlankTag" else "#E67E22"
        
        title_label = customtkinter.CTkLabel(self, text=f"{mode_name}", font=customtkinter.CTkFont(size=20, weight="bold"), text_color=title_color)
        title_label.grid(row=0, column=0, pady=(15, 5))
        
        self.last_num_label = customtkinter.CTkLabel(self, text="กำลังโหลด...", font=customtkinter.CTkFont(size=14, weight="bold"), wraplength=350)
        self.last_num_label.grid(row=1, column=0, pady=10, padx=10)
        
        qty_label = customtkinter.CTkLabel(self, text="จำนวนที่ต้องการสร้าง:")
        qty_label.grid(row=2, column=0, pady=(10,0), sticky="w", padx=30)
        
        self.quantity_entry = customtkinter.CTkEntry(self, placeholder_text="เช่น 50", corner_radius=8)
        self.quantity_entry.grid(row=3, column=0, pady=(5, 15), padx=30, sticky="ew")
        
        btn_color = "#D35400" if mode_name == "NC Tag" else ["#3a7ebf", "#1f538d"]
        hover_color = "#E67E22" if mode_name == "NC Tag" else ["#325882", "#14375e"]

        self.generate_button = customtkinter.CTkButton(self, text=f"พิมพ์ {mode_name}", command=self.start_generation, font=customtkinter.CTkFont(size=14, weight="bold"), corner_radius=8, fg_color=btn_color, hover_color=hover_color)
        self.generate_button.grid(row=4, column=0, pady=10, padx=30, sticky="ew")
        
        self.status_label = customtkinter.CTkLabel(self, text="พร้อมใช้งาน", wraplength=300)
        self.status_label.grid(row=5, column=0, pady=5)
        
        action_frame = customtkinter.CTkFrame(self, fg_color="transparent")
        action_frame.grid(row=6, column=0, pady=15)
        
        self.reprint_button = customtkinter.CTkButton(action_frame, text="Reprint", command=self.open_reprint_window, fg_color="transparent", border_width=1, width=100, text_color=("#1A1A1A", "#DCE4EE"))
        self.reprint_button.pack(side="left", padx=5)
        
        self.settings_button = customtkinter.CTkButton(action_frame, text="Settings", command=self.open_settings_window, fg_color="transparent", border_width=1, width=100, text_color=("#1A1A1A", "#DCE4EE"))
        self.settings_button.pack(side="left", padx=5)
        
        self.load_last_number()

    def get_number_save_file(self):
        mode_clean = self.mode_name.replace(" ", "")
        return os.path.join(get_app_dir(), f"last_number_{mode_clean}.txt")

    def reset_main_counter(self): 
        self.set_main_counter(0, f"รีเซ็ตตัวนับเลข {self.mode_name} เรียบร้อยแล้ว")

    def set_main_counter(self, number_to_set, message="ตั้งค่าเลขล่าสุดเรียบร้อยแล้ว"): 
        self.save_last_number(number_to_set)
        self.update_status(message, "green")

    def open_settings_window(self):
        if self.toplevel_window is None or not self.toplevel_window.winfo_exists():
            self.toplevel_window = SettingsWindow(self, self.mode_name)
        self.toplevel_window.focus()

    def open_reprint_window(self):
        if self.toplevel_window is None or not self.toplevel_window.winfo_exists():
            self.toplevel_window = ReprintWindow(self, self.mode_name)
        self.toplevel_window.focus()

    def update_display(self):
        mode_config = self.main_app.config.get(self.mode_name, {})
        prefix = mode_config.get("prefix", "")
        digits = mode_config.get("digits", 7)
        
        last_num_str = f"{self.last_number:0{digits}d}"
        next_num_str = f"{self.last_number + 1:0{digits}d}"
        
        if self.last_number == 0: 
            self.last_num_label.configure(text=f"เริ่มที่: {prefix}{next_num_str}")
        else: 
            self.last_num_label.configure(text=f"ล่าสุด: {prefix}{last_num_str}  (คิวต่อไป: {prefix}{next_num_str})")

    def load_last_number(self):
        save_file = self.get_number_save_file()
        try:
            if os.path.exists(save_file):
                with open(save_file, 'r') as f: 
                    last_num = int(f.read().strip())
            else: 
                last_num = 0
            self.last_number = last_num
        except (ValueError, FileNotFoundError): 
            self.last_number = 0
        self.update_display()

    def save_last_number(self, num_to_save):
        save_file = self.get_number_save_file()
        with open(save_file, 'w') as f: 
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
        mode_config = self.main_app.config.get(self.mode_name, {})
        prefix = mode_config.get("prefix", "")
        digits = mode_config.get("digits", 7)
        
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
        
        output_pdf = os.path.join(get_app_dir(), f"labels_{self.mode_name.replace(' ', '')}.pdf")
        thread = threading.Thread(target=create_label_pdf, args=(output_pdf, start_num, end_num, prefix, digits, self.update_status, self.generation_completed))
        thread.start()

class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        self.title("ระบบออก Tag")
        try:
            self.iconbitmap(resource_path("logo1.ico")) 
        except:
            pass
        self.geometry("850x450") 
        self.grid_columnconfigure(0, weight=1, uniform="group1")
        self.grid_columnconfigure(1, weight=1, uniform="group1")
        self.grid_rowconfigure(1, weight=1)
        
        self.load_config()
        
        header_frame = customtkinter.CTkFrame(self, fg_color="transparent")
        header_frame.grid(row=0, column=0, columnspan=2, sticky="ew", padx=10, pady=5)
        header_frame.grid_columnconfigure(0, weight=1)
        
        title_label = customtkinter.CTkLabel(header_frame, text="ระบบสร้าง BlankTag และ NC Tag", font=customtkinter.CTkFont(size=20, weight="bold"))
        title_label.grid(row=0, column=0, sticky="w", padx=10)
        
        self.theme_switch = customtkinter.CTkSwitch(header_frame, text="Dark Mode", command=lambda: customtkinter.set_appearance_mode("dark" if self.theme_switch.get() else "light"))
        self.theme_switch.grid(row=0, column=1, sticky="e", padx=10)
        self.theme_switch.select()
        
        self.frame_blank = GeneratorFrame(self, self, "BlankTag")
        self.frame_blank.grid(row=1, column=0, padx=(20, 10), pady=10, sticky="nsew")
        
        self.frame_nc = GeneratorFrame(self, self, "NC Tag")
        self.frame_nc.grid(row=1, column=1, padx=(10, 20), pady=10, sticky="nsew")
        
    def load_config(self):
        default_config = {
            "BlankTag": {"prefix": "B", "digits": 7},
            "NC Tag": {"prefix": "NC", "digits": 7}
        }
        try:
            with open(CONFIG_FILE, 'r') as f: 
                data = json.load(f)
                if "prefix" in data:
                    self.config = default_config
                    self.config["BlankTag"]["prefix"] = data.get("prefix", "B")
                    self.config["BlankTag"]["digits"] = data.get("digits", 7)
                    self.save_config()
                else:
                    self.config = data
                    for mode in default_config:
                        if mode not in self.config:
                            self.config[mode] = default_config[mode]
        except (FileNotFoundError, json.JSONDecodeError): 
            self.config = default_config
            self.save_config()

    def save_config(self):
        with open(CONFIG_FILE, 'w') as f: 
            json.dump(self.config, f, indent=4)

if __name__ == "__main__":
    app = App()
    app.mainloop()