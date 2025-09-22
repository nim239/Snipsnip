import csv
import os
import sys
from pathlib import Path
import xml.etree.ElementTree as ET
from xml.dom import minidom
import subprocess
import json
import customtkinter as ctk
from tkinter import filedialog, messagebox
from collections import Counter
import shutil
import threading
import re
import urllib.request
import io

# --- Backend Logic ---
def get_video_metadata(video_path):
    command = ['ffprobe', '-v', 'error', '-show_streams', '-show_format', '-of', 'json', str(video_path)]
    try:
        result = subprocess.run(command, capture_output=True, text=True, check=True, encoding='utf-8')
        return json.loads(result.stdout)
    except (subprocess.CalledProcessError, FileNotFoundError, KeyError, json.JSONDecodeError):
        return None

def time_to_frames(t, fps):
    if not t or not isinstance(t, str): return 0
    t = t.strip()
    if t.count(':') == 3:
        parts = t.split(':'); h, m, s, f = map(int, parts)
        return (h * 3600 + m * 60 + s) * fps + f
    parts = [float(x) for x in t.split(':')]
    if len(parts) == 2: m, s = parts; return round((m * 60 + s) * fps)
    elif len(parts) == 3: h, m, s = parts; return round((h * 3600 + m * 60 + s) * fps)
    raise ValueError(f"Dinh dang thoi gian khong hop le: {t}")

def parse_inout(s, fps):
    s = s.strip()
    if '-' not in s:
        last_colon_index = s.rfind(':')
        if last_colon_index != -1:
            ins = s[:last_colon_index]; outs = s[last_colon_index+1:]
            s = f"{ins} - {outs}"
        else: raise ValueError(f"Dinh dang InOut khong hop le: {s}")
    ins, outs = s.split('-')
    return time_to_frames(ins.strip(), float(fps)), time_to_frames(outs.strip(), float(fps))

def find_video_file(video_folder, base_filename):
    if not base_filename: return None
    folder = Path(video_folder)
    p_filename = Path(base_filename)
    candidates = [
        folder / f"{p_filename.stem}_Proxy.mp4",
        folder / f"{p_filename.name}_Proxy.mp4",
        folder / p_filename.name
    ]
    for path in candidates:
        if path.exists():
            return str(path)
    return None

# --- Custom Dialogs ---
class ErrorEditorDialog(ctk.CTkToplevel):
    def __init__(self, parent, error_rows):
        super().__init__(parent)
        self.parent = parent
        self.error_rows_with_indices = error_rows
        self.current_error_index = 0
        self.title("Sửa lỗi dữ liệu CSV"); self.geometry("600x400"); self.resizable(False, False); self.transient(parent); self.lift(); self.focus_force()
        self.info_label = ctk.CTkLabel(self, text="", justify="left", font=("Consolas", 12)); self.info_label.pack(pady=10, padx=20, fill="x")
        self.filename_entry = ctk.CTkEntry(self, width=560); self.filename_entry.pack(pady=5, padx=20, fill="x")
        self.timecode_entry = ctk.CTkEntry(self, width=560); self.timecode_entry.pack(pady=5, padx=20, fill="x")
        self.status_label = ctk.CTkLabel(self, text="", text_color="gray"); self.status_label.pack(pady=5, padx=20)
        self.progress_label = ctk.CTkLabel(self, text=""); self.progress_label.pack(pady=(5, 0), padx=20)
        self.progress_bar = ctk.CTkProgressBar(self, width=560); self.progress_bar.set(0); self.progress_bar.pack_forget()
        self.button_frame = ctk.CTkFrame(self, fg_color="transparent"); self.button_frame.pack(pady=10)
        self.find_copy_button = ctk.CTkButton(self.button_frame, text="Tìm & Chép File...", command=self.find_and_copy_file)
        self.create_gap_button = ctk.CTkButton(self.button_frame, text="Tạo Gap & Tiếp", command=self.create_gap)
        self.save_button = ctk.CTkButton(self.button_frame, text="Lưu & Kiểm tra lại", command=self.save_and_recheck)
        self.skip_button = ctk.CTkButton(self.button_frame, text="Bỏ qua ->", command=self.next_error)
        self.finish_button = ctk.CTkButton(self.button_frame, text="Gap Tất Cả & Đóng", command=self.close_and_process_remaining)
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.load_current_error()
    def load_current_error(self, event=None):
        if self.current_error_index >= len(self.error_rows_with_indices):
            messagebox.showinfo("Hoàn tất", "Đã duyệt qua tất cả các lỗi."); self.destroy(); return
        original_index, row = self.error_rows_with_indices[self.current_error_index]
        status = row.get('status')
        total_errors = len(self.error_rows_with_indices)
        progress_text = f"Lỗi {self.current_error_index + 1} / {total_errors}"
        info = (f"Tìm thấy {total_errors} lỗi. ({progress_text})\n" f"Dòng CSV gốc số: {original_index + 1}\n" f"Lỗi hiện tại: {status.upper() if status else 'N/A'}")
        self.info_label.configure(text=info)
        for widget in self.button_frame.winfo_children(): widget.pack_forget()
        if status == 'File not found':
            self.timecode_entry.configure(state="disabled"); self.filename_entry.configure(state="normal")
            self.find_copy_button.pack(side="left", padx=5); self.create_gap_button.pack(side="left", padx=5)
        else:
            self.timecode_entry.configure(state="normal"); self.filename_entry.configure(state="normal")
            self.save_button.pack(side="left", padx=10); self.skip_button.pack(side="left", padx=10)
        self.finish_button.pack(side="left", padx=10)
        self.filename_entry.delete(0, "end"); self.filename_entry.insert(0, row.get('filename', ''))
        self.timecode_entry.delete(0, "end"); self.timecode_entry.insert(0, row.get('Time in - time out', ''))
        self.status_label.configure(text="Chọn một hành động hoặc sửa thông tin rồi Lưu"); self.progress_label.configure(text=""); self.progress_bar.pack_forget()
    def find_and_copy_file(self, event=None):
        original_index, row = self.error_rows_with_indices[self.current_error_index]
        missing_filename = row.get('filename')
        src_path = filedialog.askopenfilename()
        if not src_path: return
        if Path(src_path).name != missing_filename: messagebox.showerror("Sai tên file", f"Bạn đã chọn file '{Path(src_path).name}'.\nVui lòng chọn đúng file có tên là '{missing_filename}'."); return
        dest_folder = self.parent.full_video_path
        if not dest_folder or not Path(dest_folder).is_dir(): messagebox.showerror("Lỗi", "Thư mục video nguồn không hợp lệ!"); return
        self.progress_bar.pack(pady=(0, 10), padx=20)
        for widget in self.button_frame.winfo_children(): widget.configure(state="disabled")
        copy_thread = threading.Thread(target=self.threaded_copy_with_progress, args=(src_path, dest_folder, original_index)); copy_thread.start()
    def threaded_copy_with_progress(self, src, dest_folder, original_index):
        dest = Path(dest_folder) / Path(src).name
        total_size = Path(src).stat().st_size
        copied_size = 0
        try:
            with open(src, 'rb') as fsrc, open(dest, 'wb') as fdest:
                while True:
                    chunk = fsrc.read(4096 * 1024)
                    if not chunk: break
                    fdest.write(chunk); copied_size += len(chunk)
                    self.after(0, self.update_progress, copied_size / total_size, copied_size, total_size)
            self.after(0, self.finish_copy, original_index, True)
        except Exception as e: self.after(0, self.finish_copy, e, False)
    def update_progress(self, progress, copied_size, total_size):
        self.progress_bar.set(progress); self.progress_label.configure(text=f"Đang chép: {copied_size / (1024*1024):.1f} / {total_size / (1024*1024):.1f} MB")
    def finish_copy(self, result_or_index, success):
        self.progress_bar.pack_forget()
        for widget in self.button_frame.winfo_children(): widget.configure(state="normal")
        if success:
            original_index = result_or_index
            row = self.parent.processed_data[original_index]
            updated_row, _ = self.parent._validate_row(row)
            self.parent.processed_data[original_index] = updated_row
            self.parent.log_message(f"Đã chép và xác thực lại file cho dòng {original_index + 1}", 'success'); self.next_error()
        else: messagebox.showerror("Lỗi sao chép", f"Không thể sao chép file: {result_or_index}")
    def create_gap(self, event=None):
        original_index, row = self.error_rows_with_indices[self.current_error_index]
        try:
            timeline_fps_val = self.parent.fps_map[self.parent.fps_var.get()]
            fps = self.parent.most_common_fps if timeline_fps_val == 'auto' else timeline_fps_val
            in_f, out_f = parse_inout(row.get('Time in - time out', ''), fps)
            if (out_f - in_f) <= 0: raise ValueError("Duration is not positive")
            row['type'] = 'gap'; row['status'] = 'gap'
            self.parent.processed_data[original_index] = row
            self.parent.log_message(f"Đã tạo khoảng trống cho dòng {original_index + 1}", 'success'); self.next_error()
        except Exception as e: self.status_label.configure(text=f"Không thể tạo khoảng trống. Lỗi timecode? ({e})", text_color="red")
    def save_and_recheck(self, event=None):
        original_index, row = self.error_rows_with_indices[self.current_error_index]
        row['filename'] = self.filename_entry.get().strip(); row['Time in - time out'] = self.timecode_entry.get().strip()
        updated_row, _ = self.parent._validate_row(row)
        self.parent.processed_data[original_index] = updated_row
        new_status = updated_row.get('status')
        if new_status == 'ok':
            self.status_label.configure(text=f"Thành công! Trạng thái mới: OK", text_color="#00AA00"); self.after(1200, self.next_error)
        else:
            self.status_label.configure(text=f"Vẫn còn lỗi! Trạng thái mới: {new_status.upper()}", text_color="#F44336")
            self.info_label.configure(text=f"Đang sửa dòng CSV số: {original_index + 1} / {len(self.parent.processed_data)}\nLỗi hiện tại: {new_status.upper()}")
    def next_error(self, event=None): self.current_error_index += 1; self.load_current_error()
    def close_and_process_remaining(self, event=None):
        remaining_count = len(self.error_rows_with_indices) - self.current_error_index
        if remaining_count > 0:
            self.parent.log_message(f"Chuyển {remaining_count} lỗi còn lại thành khoảng trống...")
            for i in range(self.current_error_index, len(self.error_rows_with_indices)):
                original_index, row = self.error_rows_with_indices[i]
                row['type'] = 'gap'; row['status'] = 'gap'
                self.parent.processed_data[original_index] = row
            self.parent.log_message("Đã xử lý xong các lỗi còn lại.", 'success')
        self.destroy()

# --- Main Application Class (Wizard UI) ---
class AutoCutApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        config_dir = Path.home() / ".autocut_gui_config"
        config_dir.mkdir(exist_ok=True)
        self.CONFIG_FILE = config_dir / "config.json"
        self.full_csv_path = ""; self.full_video_path = ""; self.full_xml_path = ""
        self.processed_data = []; self.most_common_fps = 25.0
        self.resolution_map = {"1080p (Full HD)": ("1920", "1080"), "2K / QHD": ("2560", "1440"), "4K UHD": ("3840", "2160"), "Tùy chỉnh...": "custom"}
        self.fps_map = {"24 fps": 24.0, "25 fps": 25.0, "29.97 fps (DF)": 29.97, "30 fps": 30.0, "59.94 fps (DF)": 59.94, "60 fps": 60.0, "Tự động theo media": "auto"}
        self.CANONICAL_HEADERS = ['filename', 'Time in - time out', 'type', 'codec', 'framerate', 'color_profile', 'duration_frames', 'status']
        self._setup_ui(); self._create_widgets()

    def _setup_ui(self):
        self.title("SnipSnip (v25.09.19) by NamNhọ@visualStation")
        self.geometry("1000x750")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        ctk.set_appearance_mode("dark")
        self.log_colors = ['#00AA00', '#00AAAA', '#AA00AA', '#FFAA00', '#5555FF', '#55FF55', '#55FFFF', '#FF5555', '#FF55FF', '#FFFFFF']
        self.log_color_index = 0

    def _create_widgets(self):
        self.status_label = ctk.CTkLabel(self, text="", font=ctk.CTkFont(size=16), anchor="center"); self.status_label.grid(row=0, column=0, padx=20, pady=(10,10))
        container = ctk.CTkFrame(self, fg_color="transparent"); container.grid(row=1, column=0, sticky='nsew', padx=20, pady=10); container.grid_columnconfigure(0, weight=1); container.grid_rowconfigure(0, weight=1)
        self.screen1_script = self._create_screen1_script(container)
        self.screen2_video = self._create_screen2_video(container)
        self.screen3_main = self._create_screen3_main(container)
        for frame in [self.screen1_script, self.screen2_video, self.screen3_main]: frame.grid(row=0, column=0, sticky='nsew')
        self.load_config()
        self._show_screen(1)

    def _show_screen(self, screen_number):
        if screen_number == 1: self.status_label.configure(text="Bước 1: Cho xin kịch bản đi bạn ơi\n ( giờ chỉ đang hỗ trợ file CSV \n xoá bớt mấy note linh tinh trong ggsheet \n tách sẵn filename + timecode thành cột  rồi bấm tải về csv \n hoặc link Google Sheet - hơi hên xui)", justify="center"); self.screen1_script.tkraise()
        elif screen_number == 2: self.status_label.configure(text="Bước 2: Chọn thư mục chứa source video đã tải về", justify="center"); self.screen2_video.tkraise()
        elif screen_number == 3: self.status_label.configure(text="Bước 3: Tạo Sequence với thông số ở dưới?", justify="center"); self.screen3_main.tkraise()

    def _create_screen1_script(self, parent):
        frame = ctk.CTkFrame(parent, fg_color="transparent"); frame.grid_columnconfigure(0, weight=1); frame.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(frame, text="Xin chào! Hãy bắt đầu nào.", font=ctk.CTkFont(size=28, weight="bold")).pack(pady=(20, 40))
        button_frame = ctk.CTkFrame(frame, fg_color="transparent"); button_frame.pack(pady=20)
        ctk.CTkButton(button_frame, text="📂 Chọn file Kịch bản trong PC ?", width=300, height=60, command=self.browse_csv, font=ctk.CTkFont(size=16)).pack(pady=15, padx=20)
        ctk.CTkButton(button_frame, text="🔗 Kịch bản trong link Google Sheet", width=300, height=60, command=self.import_from_google_sheet, font=ctk.CTkFont(size=16)).pack(pady=15, padx=20)
        return frame

    def _create_screen2_video(self, parent):
        frame = ctk.CTkFrame(parent, fg_color="transparent"); frame.grid_columnconfigure(0, weight=1); frame.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(frame, text="Tuyệt vời! Giờ hãy cho tôi biết video của bạn ở đâu.", font=ctk.CTkFont(size=24, weight="bold")).pack(pady=(20, 40))
        ctk.CTkButton(frame, text="📂 Chọn thư mục Video", width=300, height=60, command=self.browse_video_folder, font=ctk.CTkFont(size=16)).pack(pady=20)
        return frame

    def _create_screen3_main(self, parent):
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.grid_columnconfigure(1, weight=1)
        frame.grid_rowconfigure(0, weight=1)
        left_frame = ctk.CTkFrame(frame, corner_radius=0, fg_color="transparent"); left_frame.grid(row=0, column=0, sticky='nsew', padx=(0, 6)); left_frame.grid_columnconfigure(0, weight=1); left_frame.grid_rowconfigure(5, weight=1)
        right_frame = ctk.CTkFrame(frame, corner_radius=12); right_frame.grid(row=0, column=1, sticky='nsew', padx=(6, 0)); right_frame.grid_columnconfigure(0, weight=1); right_frame.grid_rowconfigure(1, weight=1)
        main_frame = ctk.CTkFrame(left_frame, corner_radius=12); main_frame.grid(row=0, column=0, sticky='ew', pady=(0,12)); main_frame.grid_columnconfigure(1, weight=1)
        self.csv_entry = ctk.CTkEntry(main_frame, placeholder_text="..."); self.csv_entry.configure(state="disabled")
        self.video_entry = ctk.CTkEntry(main_frame, placeholder_text="..."); self.video_entry.configure(state="disabled")
        self.xml_entry = ctk.CTkEntry(main_frame, placeholder_text="Chưa chọn nơi lưu...")
        ctk.CTkLabel(main_frame, text="Kịch bản:").grid(row=0, column=0, padx=(10,5), pady=5, sticky="w"); self.csv_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew");
        ctk.CTkLabel(main_frame, text="Source Video:").grid(row=1, column=0, padx=(10,5), pady=5, sticky="w"); self.video_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew");
        ctk.CTkLabel(main_frame, text="Đặt tên XML:").grid(row=2, column=0, padx=(10,5), pady=5, sticky="w"); self.xml_entry.grid(row=2, column=1, padx=5, pady=5, sticky="ew");
        self.browse_xml_btn = ctk.CTkButton(main_frame, text="Lưu tại...", width=80, command=self.browse_xml_output); self.browse_xml_btn.grid(row=2, column=2, padx=10, pady=5)
        options_frame = ctk.CTkFrame(left_frame, corner_radius=12); options_frame.grid(row=1, column=0, sticky='ew', pady=12); options_frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(options_frame, text="Resolution:").grid(row=0, column=0, padx=(10,5), pady=10, sticky="e")
        self.resolution_var = ctk.StringVar(value="1080p (Full HD)")
        self.resolution_menu = ctk.CTkOptionMenu(options_frame, variable=self.resolution_var, values=list(self.resolution_map.keys()), command=self._on_resolution_change); self.resolution_menu.grid(row=0, column=1, padx=5, pady=10, sticky="w")
        self.custom_res_frame = ctk.CTkFrame(options_frame, fg_color="transparent"); self.custom_res_frame.grid(row=0, column=2, padx=10, pady=5, sticky="w")
        vcmd = (self.register(lambda P: P.isdigit() or P == ""), '%P')
        self.custom_width_entry = ctk.CTkEntry(self.custom_res_frame, width=60, placeholder_text="Width", validate="key", validatecommand=vcmd); self.custom_width_entry.grid(row=0, column=0, padx=0)
        ctk.CTkLabel(self.custom_res_frame, text="x").grid(row=0, column=1, padx=5)
        self.custom_height_entry = ctk.CTkEntry(self.custom_res_frame, width=60, placeholder_text="Height", validate="key", validatecommand=vcmd); self.custom_height_entry.grid(row=0, column=2, padx=0)
        self.custom_res_frame.grid_remove()
        ctk.CTkLabel(options_frame, text="Timeline FPS:").grid(row=1, column=0, padx=(10,5), pady=10, sticky="e"); 
        self.fps_var = ctk.StringVar(value="Tự động theo media"); 
        self.fps_menu = ctk.CTkOptionMenu(options_frame, variable=self.fps_var, values=list(self.fps_map.keys())); self.fps_menu.grid(row=1, column=1, padx=5, pady=10, sticky="w")
        action_frame = ctk.CTkFrame(left_frame, corner_radius=12); action_frame.grid(row=2, column=0, sticky='ew'); action_frame.grid_columnconfigure((0, 1), weight=1)
        self.scan_button = ctk.CTkButton(action_frame, text="🔍 Scan & Kiểm tra", height=40, command=self.scan_data, corner_radius=8); self.scan_button.grid(row=0, column=0, pady=10, padx=(10,5), sticky="ew")
        self.generate_button = ctk.CTkButton(action_frame, text="🚀 Tạo XML", height=40, command=self.generate_xml, corner_radius=8, state="disabled"); self.generate_button.grid(row=0, column=1, pady=10, padx=(5,10), sticky="ew")
        self.scan_progress_label = ctk.CTkLabel(left_frame, text="", anchor="w"); self.scan_progress_label.grid(row=3, column=0, sticky="ew", padx=10, pady=(5,0))
        self.scan_progress_bar = ctk.CTkProgressBar(left_frame, corner_radius=8); self.scan_progress_bar.grid(row=4, column=0, sticky="ew", padx=10, pady=(0,5))
        self.scan_progress_label.grid_remove(); self.scan_progress_bar.grid_remove()
        self.console_text = ctk.CTkTextbox(left_frame, height=200, corner_radius=8); self.console_text.grid(row=5, column=0, sticky="nsew", pady=12); self.console_text.configure(state="disabled")
        ctk.CTkLabel(right_frame, text="CSV Data Preview").grid(row=0, column=0, pady=5)
        self.csv_preview_text = ctk.CTkTextbox(right_frame, corner_radius=8, font=("Consolas", 11)); self.csv_preview_text.grid(row=1, column=0, sticky="nsew", padx=8, pady=(0,8)); self.csv_preview_text.configure(state="disabled")
        return frame

    def _on_resolution_change(self, choice):
        if choice == "Tùy chỉnh...": self.custom_res_frame.grid()
        else: self.custom_res_frame.grid_remove()

    def log_message(self, message, status_type=None):
        self.console_text.configure(state="normal"); color = self.log_colors[self.log_color_index]; tag_name = f'color_{self.log_color_index}'; self.console_text.tag_config(tag_name, foreground=color); self.console_text.insert("end", message + '\n', tag_name); self.console_text.see("end"); self.console_text.configure(state="disabled"); self.log_color_index = (self.log_color_index + 1) % len(self.log_colors); self.update_idletasks()

    def save_config(self):
        config_data = {'csv_path': str(self.full_csv_path), 'video_path': str(self.full_video_path), 'xml_path': str(self.full_xml_path)}
        with open(self.CONFIG_FILE, 'w') as f: json.dump(config_data, f, indent=4)

    def load_config(self):
        try:
            if self.CONFIG_FILE.exists():
                with open(self.CONFIG_FILE, 'r') as f: config_data = json.load(f)
                self.full_csv_path = config_data.get('csv_path', '')
                self.full_video_path = config_data.get('video_path', '')
                self.full_xml_path = config_data.get('xml_path', '')
                if self.full_csv_path: self.csv_entry.insert(0, Path(self.full_csv_path).name)
                if self.full_video_path: self.video_entry.insert(0, Path(self.full_video_path).name)
                if self.full_xml_path: self.xml_entry.insert(0, Path(self.full_xml_path).name)
        except (json.JSONDecodeError, KeyError) as e:
            self.log_message(f"Lỗi đọc file config: {e}", 'error')

    def browse_csv(self):
        path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")]);
        if path:
            self.full_csv_path = path
            self.csv_entry.delete(0, "end"); self.csv_entry.insert(0, Path(path).name)
            self.log_message(f"Đã chọn kịch bản: {Path(path).name}")
            self.save_config(); self._show_screen(2)

    def browse_video_folder(self):
        path = filedialog.askdirectory();
        if path:
            self.full_video_path = path
            self.video_entry.delete(0, "end"); self.video_entry.insert(0, Path(path).name)
            self.log_message(f"Đã chọn thư mục video: {Path(path).name}")
            self.save_config(); self._show_screen(3)

    def browse_xml_output(self):
        initial_file = Path(self.full_csv_path).stem + "_Final.xml" if self.full_csv_path else "output.xml"
        path = filedialog.asksaveasfilename(defaultextension=".xml", filetypes=[("XML files", "*.xml")], initialfile=initial_file)
        if path: self.full_xml_path = path; self.xml_entry.delete(0, "end"); self.xml_entry.insert(0, Path(path).name); self.log_message(f"File XML sẽ được lưu tại: {Path(path).name}"); self.save_config()

    def import_from_google_sheet(self):
        dialog = ctk.CTkInputDialog(text="Dán link Google Sheet vào đây:", title="Nhập từ Google Sheet"); url = dialog.get_input()
        if not url: return
        try:
            match = re.search(r'/spreadsheets/d/([a-zA-Z0-9-_]+)(/.*)?gid=([0-9]+)', url)
            if not match: raise ValueError("Link Google Sheet không hợp lệ hoặc không chứa GID.")
            sheet_id, gid = match.group(1), match.group(3)
            download_url = f'https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}'
            self.log_message("Đang tải dữ liệu từ Google Sheet..."); self.update_idletasks()
            with urllib.request.urlopen(download_url) as response:
                if response.getcode() != 200: raise Exception(f"Lỗi máy chủ: {response.getcode()}")
                raw_data = response.read()
            self.log_message("Đang dọn dẹp dữ liệu..."); self.update_idletasks()
            cleaned_csv_content, skipped_rows_info = self.clean_google_sheet_data(raw_data)
            if not cleaned_csv_content: raise Exception("Không tìm thấy dữ liệu hợp lệ sau khi dọn dẹp.")
            temp_dir = self.CONFIG_FILE.parent; temp_csv_path = temp_dir / "g_sheet_import.csv"
            temp_csv_path.write_text(cleaned_csv_content, encoding='utf-8')
            if skipped_rows_info:
                reason_counts = Counter(info['reason'] for info in skipped_rows_info)
                summary_messages = [f"- {count} dòng: {reason}" for reason, count in reason_counts.items()]
                summary_text = "\n".join(summary_messages)
                self.log_message(f"Đã bỏ qua {len(skipped_rows_info)} dòng không hợp lệ. Chi tiết:", 'error'); self.log_message(summary_text, 'error')
                messagebox.showinfo("Dọn dẹp dữ liệu", f"Đã bỏ qua {len(skipped_rows_info)} dòng không hợp lệ.\n\nChi tiết:\n{summary_text}")
            self.full_csv_path = str(temp_csv_path)
            self.csv_entry.delete(0, "end"); self.csv_entry.insert(0, temp_csv_path.name)
            self.log_message(f"Đã nhập thành công từ Google Sheet: {temp_csv_path.name}", 'success')
            self.save_config(); self._show_screen(2)
        except Exception as e: self.log_message(f"Lỗi khi nhập từ Google Sheet: {e}", 'error'); messagebox.showerror("Lỗi", f"Không thể nhập dữ liệu từ Google Sheet:\n{e}")

    def clean_google_sheet_data(self, raw_data):
        cleaned_rows = []; skipped_rows_info = []
        decoded_content = raw_data.decode('utf-8', errors='ignore')
        csv_file = io.StringIO(decoded_content); reader = csv.reader(csv_file)
        timecode_segment_pattern = r'(?:\d{1,2}:)?\d{1,2}:\d{2}\s*-\s*(?:\d{1,2}:)?\d{1,2}:\d{2}'
        for i, row in enumerate(reader):
            row_num = i + 1
            if any('(BỎ)' in str(cell).upper() for cell in row): skipped_rows_info.append({'row_index': row_num, 'raw_row': ','.join(row), 'reason': 'Chứa từ khóa "(BỎ)"'}); continue
            if len(row) < 5:
                if len(row) > 0 and 'FRAME' in str(row[0]).upper() and len(row) > 2 and 'TIMECODE' in str(row[2]).upper() and len(row) > 4 and 'SOURCE' in str(row[4]).upper(): continue
                skipped_rows_info.append({'row_index': row_num, 'raw_row': ','.join(row), 'reason': 'Dòng quá ngắn để chứa dữ liệu hợp lệ'}); continue
            found_filename = None; found_timecode = None
            for cell_content in row:
                cell_content_stripped = str(cell_content).strip()
                if not found_filename and any(cell_content_stripped.upper().endswith(ext) for ext in ['.MP4', '.MOV', '.MXF', '.MTS', '.AVI', '.WMV', '.FLV', '.WEBM']): found_filename = cell_content_stripped
                temp_timecode_raw = cell_content_stripped.replace('\n', ' ')
                match_colon_timecode = re.match(r'(\d{1,2}:\d{2}):(\d{1,2}:\d{2})', temp_timecode_raw)
                if match_colon_timecode: temp_timecode_raw = f"{match_colon_timecode.group(1)}-{match_colon_timecode.group(2)}"
                if not found_timecode and re.search(timecode_segment_pattern, temp_timecode_raw): found_timecode = temp_timecode_raw
                if found_filename and found_timecode: break
            if found_filename and found_timecode:
                timecode_raw = found_timecode; filename_raw = found_filename
                is_video = any(filename_raw.upper().endswith(ext) for ext in ['.MP4', '.MOV', '.MXF', '.MTS', '.AVI', '.WMV', '.FLV', '.WEBM'])
                found_timecode_segments_in_raw = re.findall(timecode_segment_pattern, timecode_raw)
                if is_video and found_timecode_segments_in_raw:
                    timecode_cleaned_full = re.sub(r'\(.*?\)', '', timecode_raw).strip()
                    segments = re.findall(timecode_segment_pattern, timecode_cleaned_full)
                    if segments: [cleaned_rows.append([filename_raw, segment.strip()]) for segment in segments]
                    else: skipped_rows_info.append({'row_index': row_num, 'raw_row': ','.join(row), 'reason': 'Timecode không hợp lệ sau khi làm sạch'})
                else: skipped_rows_info.append({'row_index': row_num, 'raw_row': ','.join(row), 'reason': 'Không phải dòng dữ liệu hợp lệ (thiếu tên file/timecode)'}); continue
        if not cleaned_rows: return "", skipped_rows_info
        output = io.StringIO(); writer = csv.writer(output); writer.writerow(['filename', 'Time in - time out']); writer.writerows(cleaned_rows)
        return output.getvalue(), skipped_rows_info

    def _read_csv_rows(self):
        rows = []
        with open(self.full_csv_path, 'r', newline='', encoding='utf-8-sig') as f:
            try: has_header = csv.Sniffer().has_header(f.read(2048)) 
            except csv.Error: has_header = False
            f.seek(0); use_fallback = False
            if has_header:
                reader = csv.DictReader(f)
                def normalize(s): return s.lower().strip().replace(" ", "").replace("-", "")
                field_map = {normalize(field): field for field in reader.fieldnames}
                filename_header = next((field_map[k] for k in field_map if any(keyword in k for keyword in ['file', 'tên', 'tệp'])), None)
                timecode_header = next((field_map[k] for k in field_map if any(keyword in k for keyword in ['time', 'inout', 'thờigian'])), None)
                if filename_header and timecode_header:
                    self.log_message("Đã nhận diện cột filename và timecode từ header.")
                    for row in reader:
                        if not row.get(filename_header) or row.get(filename_header) == filename_header: continue
                        rows.append({'filename': row[filename_header].replace('\n',' ').strip(), 'Time in - time out': row[timecode_header].replace('\n',' ').strip()})
                else: self.log_message("Không nhận diện được cột từ header, chuyển sang chế độ dự phòng.", 'error'); use_fallback = True
            if not has_header or use_fallback:
                self.log_message("Đang đọc file theo thứ tự cột mặc định."); f.seek(0); reader = csv.reader(f)
                for r in reader:
                    if len(r) >= 2 and r[0] and r[0] != 'filename': rows.append({'filename': r[0].strip().replace('\n',' '), 'Time in - time out': r[1].strip().replace('\n',' ')})
        return rows

    def _setup_preview_tags(self):
        status_colors = {"ok": "#2ECC71", "file not found": "#E74C3C", "cannot open media": "#E74C3C", "invalid fps (0)": "#E74C3C", "FFProbe Error": "#E74C3C", "time format error": "#E67E22", "out time < in time": "#E67E22", "gap": "#F1C40F", "skipped": "#F1C40F"}
        for status, color in status_colors.items(): self.csv_preview_text.tag_config(status.replace(" ", "_"), foreground=color)
        self.fps_colors = {"24": "#3498DB", "25": "#2ECC71", "30": "#9B59B6", "60": "#E67E22"}
        for fps, color in self.fps_colors.items(): self.csv_preview_text.tag_config(f"fps_{fps}", foreground=color)
        self.csv_preview_text.tag_config("default", foreground="#FFFFFF")

    def _format_preview_row(self, row, i):
        parts = []; status = row.get('status', 'ok'); fname = row.get('filename', '')
        if len(fname) > 26: fname = fname[:25] + "..."
        if status == 'gap':
            line_content = (f"{i:<4} | {fname:<28} | {row.get('Time in - time out', ''):<22} | " f"{row.get('codec', ''):<7} | {row.get('framerate', ''):<8} | " f"{row.get('color_profile', ''):<10} | {row.get('duration_frames', ''):<9} | {status.upper()}")
            parts.append((line_content, "gap")); return parts
        color_str = row.get('color_profile', '')
        if len(color_str) > 9: color_str = color_str[:8] + ".."
        part1 = f"{i:<4} | {fname:<28} | {row.get('Time in - time out', ''):<22} | {row.get('codec', ''):<7} | "; parts.append((part1, "default"))
        framerate_str = row.get('framerate', ''); part_fps = f"{framerate_str:<8}"
        fps_int_str = framerate_str.split('.')[0] if '.' in framerate_str else framerate_str
        fps_tag = f"fps_{fps_int_str}" if fps_int_str in self.fps_colors else "default"; parts.append((part_fps, fps_tag))
        part2 = f" | {color_str:<10} | {row.get('duration_frames', ''):<9} | "; parts.append((part2, "default"))
        status_tag = status.replace(" ", "_"); parts.append((f"{status.upper()}", status_tag)); return parts

    def _update_csv_preview(self, rows_to_preview):
        self.csv_preview_text.configure(state="normal"); self.csv_preview_text.delete("1.0", "end"); self._setup_preview_tags()
        header = f"{ 'STT':<4} | {'FILENAME':<28} | {'TIMECODE':<22} | {'CODEC':<7} | {'FPS':<8} | {'COLOR':<10} | {'DURATION':<9} | {'STATUS'}\n";
        sep = "-" * (len(header) + 5) + "\n"; self.csv_preview_text.insert("end", header, "default"); self.csv_preview_text.insert("end", sep, "default")
        for i, row in enumerate(rows_to_preview):
            line_parts = self._format_preview_row(row, i + 1)
            for text, tag in line_parts: self.csv_preview_text.insert("end", text, tag)
            self.csv_preview_text.insert("end", "\n", "default")
        self.csv_preview_text.configure(state="disabled")

    def _validate_row(self, row):
        filename = row.get("filename", "").strip(); row['status'] = 'ok'; framerate = None; row['type'] = ''; row['codec'] = ''; row['color_profile'] = 'N/A'
        if not filename: row['status'] = 'skipped'; return row, framerate
        video_path = find_video_file(self.full_video_path, filename)
        if video_path is None: row['status'] = 'File not found'; return row, framerate
        row['full_path'] = video_path; row['type'] = 'clip'
        try:
            metadata = get_video_metadata(video_path)
            if metadata and metadata.get('streams'):
                video_stream = next((s for s in metadata['streams'] if s.get('codec_type') == 'video'), None)
                row['audio_tracks'] = sum(1 for s in metadata.get('streams', []) if s.get('codec_type') == 'audio')
                start_timecode_str = '00:00:00:00'
                if video_stream and 'tags' in video_stream: start_timecode_str = next((v for k, v in video_stream['tags'].items() if 'timecode' in k.lower()), start_timecode_str)
                if 'format' in metadata and 'tags' in metadata['format']: start_timecode_str = next((v for k, v in metadata['format']['tags'].items() if 'timecode' in k.lower()), start_timecode_str)
                row['start_timecode'] = start_timecode_str
                if video_stream:
                    row['codec'] = video_stream.get('codec_name', ''); row['width'] = video_stream.get('width', '1920'); row['height'] = video_stream.get('height', '1080')
                    color_transfer = video_stream.get('color_transfer'); color_space = video_stream.get('color_space'); color_primaries = video_stream.get('color_primaries')
                    if color_transfer and color_transfer not in ['unknown', 'bt709', 'smpte170m', 'bt470bg', 'bt601']: row['color_profile'] = color_transfer
                    elif (color_transfer in ['smpte170m', 'bt601']) or (color_space in ['smpte170m', 'bt601']): row['color_profile'] = 'Rec.601'
                    elif (color_transfer == 'bt709') or (color_space == 'bt709') or (color_primaries == 'bt709'): row['color_profile'] = 'Rec.709'
                    else: row['color_profile'] = 'N/A'
                else: row['codec'] = Path(filename).suffix
                if video_stream:
                    r_frame_rate = video_stream.get('r_frame_rate', '0/1')
                    try: num, den = map(float, r_frame_rate.split('/')); fps = num / den if den != 0 else 0
                    except (ValueError, ZeroDivisionError): fps = 0
                    if not fps or fps == 0: row['status'] = 'Invalid FPS (0)'; return row, framerate
                    row["framerate"] = f"{fps:.3f}"; framerate = round(fps, 3)
                    nb_frames = video_stream.get('nb_frames')
                    if nb_frames and int(nb_frames) > 0: row["duration_frames"] = str(int(nb_frames))
                    elif 'duration' in video_stream: duration = float(video_stream.get('duration', '0')); row["duration_frames"] = str(int(duration * fps))
                    else: row["duration_frames"] = '0'
                else: row['type'] = 'title'; row['status'] = 'Cannot open media'; return row, framerate
            else: row['codec'] = Path(filename).suffix; row['type'] = 'title'; row['status'] = 'Cannot open media'; return row, framerate
        except Exception as e: row['type'] = 'title'; row['status'] = f'FFProbe Error: {e}'; return row, framerate
        if row['status'] == 'ok' and row.get('type') == 'clip':
            try:
                in_f, out_f = parse_inout(row.get('Time in - time out', ''), float(row.get('framerate', 25.0)))
                if in_f >= out_f: row['status'] = 'Out time < In time'
            except (ValueError, TypeError):
                row['status'] = 'Time format error'
        return row, framerate

    def scan_data(self):
        if not all([self.full_csv_path, self.full_video_path]): messagebox.showerror("Ối!", "Bạn ơi, chọn file CSV và thư mục video trước đã nhé!"); self.log_message("Thiếu file CSV hoặc thư mục video!", 'error'); return
        self.scan_button.configure(state='disabled'); self.generate_button.configure(state='disabled')
        self.scan_progress_label.grid(); self.scan_progress_bar.grid(); self.scan_progress_bar.set(0); self.scan_progress_label.configure(text="Chuẩn bị quét..."); self.update_idletasks()
        scan_thread = threading.Thread(target=self.threaded_scan_data); scan_thread.start()

    def threaded_scan_data(self):
        try:
            rows = self._read_csv_rows()
            if not rows: self.after(0, self.finish_scan, None); return
            self.after(0, self.log_message, "Bắt đầu quét và xác thực dữ liệu...")
            self.processed_data = []; scanned_framerates = []
            num_rows = len(rows)
            for i, row in enumerate(rows):
                progress = (i + 1) / num_rows; filename = Path(row.get('filename', 'N/A')).name
                self.after(0, self.update_scan_progress, progress, f"Đang quét {i+1}/{num_rows}: {filename}")
                validated_row, framerate = self._validate_row(row)
                self.processed_data.append(validated_row)
                if framerate: scanned_framerates.append(framerate)
            if scanned_framerates: self.most_common_fps = Counter(scanned_framerates).most_common(1)[0][0]
            self.after(0, self.finish_scan, self.processed_data)
        except Exception as e: self.after(0, self.finish_scan, e)

    def update_scan_progress(self, progress, message): self.scan_progress_bar.set(progress); self.scan_progress_label.configure(text=message)

    def finish_scan(self, result):
        self.scan_progress_label.grid_remove(); self.scan_progress_bar.grid_remove(); self.scan_button.configure(state='normal')
        if isinstance(result, Exception): messagebox.showerror("Lỗi khi quét!", f"Đã có lỗi xảy ra:\n{result}"); self.log_message(f"Lỗi khi quét data: {result}", 'error'); return
        if result is None: self.log_message("CSV rỗng hoặc không đọc được dữ liệu.", 'error'); return
        self.log_message(f"FPS phổ biến nhất trong media là: {self.most_common_fps}")
        self._update_csv_preview(self.processed_data)
        error_rows = [(i, row) for i, row in enumerate(self.processed_data) if row.get('status') not in ['ok', 'skipped', 'gap'] and row.get('filename')]
        if error_rows: self.log_message(f"Tìm thấy {len(error_rows)} lỗi. Mở trình chỉnh sửa..."); editor = ErrorEditorDialog(self, error_rows); self.wait_window(editor); self._update_csv_preview(self.processed_data)
        try:
            with open(self.full_csv_path, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.DictWriter(f, fieldnames=self.CANONICAL_HEADERS, extrasaction='ignore')
                writer.writeheader(); writer.writerows(self.processed_data)
            self.log_message("Quét và cập nhật CSV thành công!", 'success')

            # --- USER WARNING ---
            self.console_text.configure(state="normal")
            separator = "\n" + "="*70 + "\n"
            self.console_text.insert("end", separator, "default")
            
            warning_tag = "warning_highlight"
            self.console_text.tag_config(warning_tag, background="yellow", foreground="black", justify='center')
            
            warning_message = "Phiên Bản Google sheet đọc data éo chính xác đâu bro! đtao đang mò dở đoạn đấy nên tốt nhất là tạm thời tự check số lượng clip rồi fill thêm vào CSV sau bước này dùm tao hehe !"
            self.console_text.insert("end", f"\n{warning_message}\n\n", warning_tag)
            
            self.console_text.insert("end", "="*70 + "\n", "default")
            self.console_text.see("end")
            self.console_text.configure(state="disabled")
            # --- END USER WARNING ---

            self.generate_button.configure(state="normal")
        except Exception as e: messagebox.showerror("Lỗi khi lưu CSV!", f"Không thể ghi lại file CSV:\n{e}"); self.log_message(f"Lỗi ghi file CSV: {e}", 'error')

    def _write_xml_to_file(self, xmeml_element, file_path):
        xml_string = ET.tostring(xmeml_element, 'utf-8')
        pretty_xml = minidom.parseString(xml_string).toprettyxml(indent="  ")
        with open(file_path, "w", encoding="utf-8") as f:
            f.write(pretty_xml)

    def _create_xml_scaffold(self, sequence_name, timeline_fps, width, height):
        xmeml = ET.Element("xmeml", version="4"); sequence = ET.SubElement(xmeml, "sequence", id="sequence-1"); ET.SubElement(sequence, "name").text = sequence_name
        rate_seq = ET.SubElement(sequence, "rate"); ET.SubElement(rate_seq, "timebase").text = str(round(timeline_fps)); ET.SubElement(rate_seq, "ntsc").text = "FALSE"
        media = ET.SubElement(sequence, "media"); video = ET.SubElement(media, "video"); vid_format = ET.SubElement(video, "format"); vid_sample_chars = ET.SubElement(vid_format, "samplecharacteristics")
        rate_format = ET.SubElement(vid_sample_chars, "rate"); ET.SubElement(rate_format, "timebase").text = str(round(timeline_fps)); ET.SubElement(vid_sample_chars, "width").text = width; ET.SubElement(vid_sample_chars, "height").text = height; ET.SubElement(vid_sample_chars, "pixelaspectratio").text = "square"
        video_track = ET.SubElement(video, "track"); audio = ET.SubElement(media, "audio"); timeline_audio_tracks = [ET.SubElement(audio, "track") for _ in range(8)]
        return xmeml, sequence, video_track, timeline_audio_tracks

    def _create_file_node(self, row, file_id):
        source_fps = float(row.get('framerate')); video_path = Path(row.get('full_path')); file_duration = int(row.get('duration_frames', 0)); start_timecode_str = row.get('start_timecode', '00:00:00:00'); start_offset_frames = time_to_frames(start_timecode_str, source_fps); num_audio_tracks_in_file = row.get('audio_tracks', 0); video_width = row.get('width', '1920'); video_height = row.get('height', '1080')
        file_el = ET.Element("file", id=file_id); ET.SubElement(file_el, "name").text = video_path.name; path_uri = video_path.as_uri(); formatted_uri = path_uri.replace('file:///', 'file://localhost/').replace(':', '%3a', 1); ET.SubElement(file_el, "pathurl").text = formatted_uri; ET.SubElement(file_el, "duration").text = str(file_duration)
        timecode_el = ET.SubElement(file_el, "timecode"); tc_rate = ET.SubElement(timecode_el, "rate"); ET.SubElement(tc_rate, "timebase").text = str(round(source_fps)); is_ntsc = '29.97' in str(source_fps) or '59.94' in str(source_fps); ET.SubElement(tc_rate, "ntsc").text = "TRUE" if is_ntsc else "FALSE"; ET.SubElement(timecode_el, "string").text = start_timecode_str; ET.SubElement(timecode_el, "frame").text = str(int(start_offset_frames)); ET.SubElement(timecode_el, "displayformat").text = "DF" if is_ntsc else "NDF"
        media_el = ET.SubElement(file_el, "media"); vid_media_el = ET.SubElement(media_el, "video"); vid_sample_chars = ET.SubElement(vid_media_el, "samplecharacteristics"); rate_el = ET.SubElement(vid_sample_chars, "rate"); ET.SubElement(rate_el, "timebase").text = str(round(source_fps)); ET.SubElement(vid_sample_chars, "width").text = str(video_width); ET.SubElement(vid_sample_chars, "height").text = str(video_height)
        aud_media_el = ET.SubElement(media_el, "audio"); ET.SubElement(aud_media_el, "channelcount").text = str(num_audio_tracks_in_file) if num_audio_tracks_in_file > 0 else "2"
        return file_el

    def _link_clip_group(self, clip_group, clip_count):
        for source_item, _, _ in clip_group:
            for target_item, target_type, target_track_idx in clip_group:
                link = ET.SubElement(source_item, "link"); ET.SubElement(link, "linkclipref").text = target_item.get("id"); ET.SubElement(link, "mediatype").text = target_type; ET.SubElement(link, "trackindex").text = str(target_track_idx); ET.SubElement(link, "clipindex").text = str(clip_count)
                if source_item.get("id") != target_item.get("id"): ET.SubElement(link, "groupindex").text = "1"

    def finish_generate_xml(self, success, result):
        if success:
            message = result; self.log_message(message, 'success'); messagebox.showinfo("Xong!", f"File XML của bạn đã sẵn sàng tại:\n{self.full_xml_path}")
            try:
                output_folder = Path(self.full_xml_path).parent
                if output_folder: 
                    os.startfile(output_folder)
            except Exception as e: 
                self.log_message(f"Không thể tự động mở thư mục: {e}", 'error')
        else: error = result; messagebox.showerror("Ôi không!", f"Đã có lỗi nghiêm trọng xảy ra khi tạo XML:\n{error}"); self.log_message(f"Lỗi rồi bạn ơi: {error}", 'error')
        self.scan_button.configure(state='normal'); self.generate_button.configure(state='normal'); self.log_message("Sẵn sàng cho lần chạy tiếp theo.")

    def threaded_generate_xml(self):
        try:
            selected_res = self.resolution_var.get()
            if selected_res == "Tùy chỉnh...":
                w_str = self.custom_width_entry.get(); h_str = self.custom_height_entry.get()
                if not w_str.isdigit() or not h_str.isdigit() or int(w_str) == 0 or int(h_str) == 0: raise ValueError("Width và Height tùy chỉnh phải là các số hợp lệ và lớn hơn 0.")
                width, height = w_str, h_str
            else: width, height = self.resolution_map[selected_res]
            selected_fps_str = self.fps_var.get(); timeline_fps_val = self.fps_map[selected_fps_str]
            TIMELINE_FPS = self.most_common_fps if timeline_fps_val == 'auto' else timeline_fps_val
            self.log_message(f"Tạo XML với Resolution: {width}x{height}, FPS: {TIMELINE_FPS}")
            sequence_name = Path(self.full_csv_path).stem + "_FinalSequence"
            xmeml, sequence, video_track, timeline_audio_tracks = self._create_xml_scaffold(sequence_name, TIMELINE_FPS, width, height)
            total_duration_on_timeline = 0; clip_count = 0
            for row in self.processed_data:
                if row.get('type') == 'gap' and row.get('status') == 'gap':
                    try:
                        in_f, out_f = parse_inout(row.get('Time in - time out', ''), TIMELINE_FPS)
                        gap_duration_frames = out_f - in_f
                        if gap_duration_frames > 0: 
                            total_duration_on_timeline += gap_duration_frames
                    except: 
                        continue
                    continue
                if row.get('status') != 'ok' or row.get('type') != 'clip': continue
                try:
                    source_fps = float(row.get('framerate')); in_frame_raw, out_frame_raw = parse_inout(row.get('Time in - time out', ''), source_fps); file_duration = int(row.get('duration_frames', 0)); start_timecode_str = row.get('start_timecode', '00:00:00:00'); start_offset_frames = time_to_frames(start_timecode_str, source_fps)
                    in_frame = max(0, in_frame_raw - start_offset_frames); out_frame = min(file_duration, out_frame_raw - start_offset_frames)
                    if in_frame >= out_frame: continue
                    clip_duration_source_frames = out_frame - in_frame; clip_duration_timeline_frames = round(clip_duration_source_frames * (TIMELINE_FPS / source_fps))
                    if clip_duration_timeline_frames <= 0: continue
                    clip_count += 1; file_id = f"file_{clip_count}"; clip_name = row.get('filename', '').strip()
                    file_el = self._create_file_node(row, file_id)
                    vid_clipitem = ET.SubElement(video_track, "clipitem", id=f"vid_clip_{clip_count}"); ET.SubElement(vid_clipitem, "name").text = clip_name; ET.SubElement(vid_clipitem, "start").text = str(total_duration_on_timeline); ET.SubElement(vid_clipitem, "end").text = str(total_duration_on_timeline + clip_duration_timeline_frames); ET.SubElement(vid_clipitem, "in").text = str(in_frame); ET.SubElement(vid_clipitem, "out").text = str(out_frame); vid_clipitem.append(file_el)
                    clip_group = [(vid_clipitem, 'video', 1)]
                    num_audio_tracks_in_file = row.get('audio_tracks', 0)
                    for j in range(min(num_audio_tracks_in_file, 8)):
                        aud_clipitem = ET.SubElement(timeline_audio_tracks[j], "clipitem", id=f"aud_clip_{clip_count}_{j+1}"); ET.SubElement(aud_clipitem, "name").text = clip_name; ET.SubElement(aud_clipitem, "start").text = str(total_duration_on_timeline); ET.SubElement(aud_clipitem, "end").text = str(total_duration_on_timeline + clip_duration_timeline_frames); ET.SubElement(aud_clipitem, "in").text = str(in_frame); ET.SubElement(aud_clipitem, "out").text = str(out_frame); ET.SubElement(aud_clipitem, "file", id=file_id)
                        sourcetrack = ET.SubElement(aud_clipitem, "sourcetrack"); ET.SubElement(sourcetrack, "mediatype").text = "audio"; ET.SubElement(sourcetrack, "trackindex").text = str(j + 1); clip_group.append((aud_clipitem, 'audio', j + 1))
                    self._link_clip_group(clip_group, clip_count)
                    total_duration_on_timeline += clip_duration_timeline_frames
                except (ValueError, TypeError, KeyError) as e: self.log_message(f"Lỗi xử lý dòng cho clip '{row.get('filename')}': {e}. Bỏ qua.", 'error'); continue
            ET.SubElement(sequence, "duration").text = str(total_duration_on_timeline)
            self._write_xml_to_file(xmeml, self.full_xml_path)
            success_message = f"Voilà! Đã tạo xong file XML: {Path(self.full_xml_path).name} với {clip_count} clip."
            self.after(0, self.finish_generate_xml, True, success_message)
        except Exception as e: self.after(0, self.finish_generate_xml, False, e)

    def generate_xml(self):
        if not self.full_xml_path: messagebox.showerror("Ối!", "Bạn ơi, chưa chọn nơi lưu file XML!"); return
        if not self.processed_data: messagebox.showerror("Ối!", "Chưa có dữ liệu để tạo XML. Hãy Scan trước nhé!"); return
        self.scan_button.configure(state='disabled'); self.generate_button.configure(state='disabled'); self.update_idletasks()
        xml_thread = threading.Thread(target=self.threaded_generate_xml); xml_thread.start()

if __name__ == "__main__":
    app = AutoCutApp()
    app.mainloop()
