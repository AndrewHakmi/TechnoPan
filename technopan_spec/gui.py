
import os
import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
import sys
from pathlib import Path
import logging

# Set appearance mode and default color theme
ctk.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
ctk.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

class TextRedirector(object):
    def __init__(self, widget, tag="stdout"):
        self.widget = widget
        self.tag = tag

    def write(self, str):
        self.widget.configure(state="normal")
        self.widget.insert("end", str, (self.tag,))
        self.widget.see("end")
        self.widget.configure(state="disabled")
        
    def flush(self):
        pass

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Window setup
        self.title("TechnoPan Specification Generator")
        self.geometry("900x600")
        
        # Grid layout (2 columns)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Sidebar
        self.sidebar_frame = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, rowspan=4, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(4, weight=1)

        self.logo_label = ctk.CTkLabel(self.sidebar_frame, text="TechnoPan\nGenerator", font=ctk.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))

        self.appearance_mode_label = ctk.CTkLabel(self.sidebar_frame, text="Appearance Mode:", anchor="w")
        self.appearance_mode_label.grid(row=5, column=0, padx=20, pady=(10, 0))
        self.appearance_mode_optionemenu = ctk.CTkOptionMenu(self.sidebar_frame, values=["Light", "Dark", "System"],
                                                                       command=self.change_appearance_mode_event)
        self.appearance_mode_optionemenu.grid(row=6, column=0, padx=20, pady=(10, 10))
        
        # Main Content Area
        self.main_frame = ctk.CTkFrame(self, corner_radius=10)
        self.main_frame.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)
        self.main_frame.grid_columnconfigure(0, weight=1)

        # 1. Input Section
        self.label_input = ctk.CTkLabel(self.main_frame, text="Input Settings", font=ctk.CTkFont(size=16, weight="bold"))
        self.label_input.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="w")

        # DWG File Selection
        self.file_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.file_frame.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        self.file_frame.grid_columnconfigure(1, weight=1)
        
        self.btn_select_file = ctk.CTkButton(self.file_frame, text="Select DWG File", command=self.select_file)
        self.btn_select_file.grid(row=0, column=0, padx=(0, 10))
        
        self.entry_file_path = ctk.CTkEntry(self.file_frame, placeholder_text="No file selected")
        self.entry_file_path.grid(row=0, column=1, sticky="ew")
        
        # Config Selection
        self.config_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.config_frame.grid(row=2, column=0, padx=20, pady=10, sticky="ew")
        self.config_frame.grid_columnconfigure(1, weight=1)
        
        self.label_config = ctk.CTkLabel(self.config_frame, text="Configuration:")
        self.label_config.grid(row=0, column=0, padx=(0, 10))
        
        # Load configs dynamically
        self.config_files = self.get_config_files()
        self.option_config = ctk.CTkOptionMenu(self.config_frame, values=self.config_files)
        self.option_config.grid(row=0, column=1, sticky="ew")
        if "default.yml" in self.config_files:
            self.option_config.set("default.yml")

        # 2. Output Section
        self.label_output = ctk.CTkLabel(self.main_frame, text="Output Settings", font=ctk.CTkFont(size=16, weight="bold"))
        self.label_output.grid(row=3, column=0, padx=20, pady=(20, 10), sticky="w")
        
        self.output_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.output_frame.grid(row=4, column=0, padx=20, pady=10, sticky="ew")
        self.output_frame.grid_columnconfigure(1, weight=1)
        
        self.btn_select_out = ctk.CTkButton(self.output_frame, text="Output Folder", command=self.select_output_folder)
        self.btn_select_out.grid(row=0, column=0, padx=(0, 10))
        
        self.entry_out_path = ctk.CTkEntry(self.output_frame, placeholder_text="Same as input folder")
        self.entry_out_path.grid(row=0, column=1, sticky="ew")

        # Filename
        self.filename_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.filename_frame.grid(row=5, column=0, padx=20, pady=10, sticky="ew")
        
        self.label_filename = ctk.CTkLabel(self.filename_frame, text="Output Filename:")
        self.label_filename.grid(row=0, column=0, padx=(0, 10))
        
        self.entry_filename = ctk.CTkEntry(self.filename_frame, placeholder_text="spec.xlsx")
        self.entry_filename.insert(0, "spec.xlsx")
        self.entry_filename.grid(row=0, column=1, sticky="ew", ipadx=50)

        # 3. Action Section
        self.btn_process = ctk.CTkButton(self.main_frame, text="GENERATE SPECIFICATION", 
                                         font=ctk.CTkFont(size=15, weight="bold"),
                                         height=50, fg_color="green", hover_color="darkgreen",
                                         command=self.start_processing_thread)
        self.btn_process.grid(row=6, column=0, padx=20, pady=30, sticky="ew")

        # 4. Log Console
        self.console = ctk.CTkTextbox(self.main_frame, height=150)
        self.console.grid(row=7, column=0, padx=20, pady=(0, 20), sticky="nsew")
        self.main_frame.grid_rowconfigure(7, weight=1)
        
        # Redirect stdout/stderr
        # sys.stdout = TextRedirector(self.console, "stdout")
        # sys.stderr = TextRedirector(self.console, "stderr")
        
        # Set defaults
        self.appearance_mode_optionemenu.set("System")

    def get_config_files(self):
        config_dir = Path("configs")
        if not config_dir.exists():
            return ["default.yml"]
        files = [f.name for f in config_dir.glob("*.yml")]
        return sorted(files) if files else ["default.yml"]

    def change_appearance_mode_event(self, new_appearance_mode: str):
        ctk.set_appearance_mode(new_appearance_mode)

    def select_file(self):
        filename = filedialog.askopenfilename(filetypes=[("DWG Files", "*.dwg"), ("All Files", "*.*")])
        if filename:
            self.entry_file_path.delete(0, "end")
            self.entry_file_path.insert(0, filename)
            
            # Auto-suggest output name
            base_name = Path(filename).stem
            out_name = f"{base_name}_spec.xlsx"
            self.entry_filename.delete(0, "end")
            self.entry_filename.insert(0, out_name)

    def select_output_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.entry_out_path.delete(0, "end")
            self.entry_out_path.insert(0, folder)

    def log(self, message):
        def _log():
            self.console.configure(state="normal")
            self.console.insert("end", message + "\n")
            self.console.see("end")
            self.console.configure(state="disabled")
        self.after(0, _log)

    def start_processing_thread(self):
        thread = threading.Thread(target=self.process)
        thread.start()

    def process(self):
        dwg_path_str = self.entry_file_path.get()
        if not dwg_path_str:
            self.log("Error: Please select a DWG file.")
            return

        dwg_path = Path(dwg_path_str)
        if not dwg_path.exists():
            self.log(f"Error: File not found: {dwg_path}")
            return

        config_name = self.option_config.get()
        config_path = Path("configs") / config_name
        
        out_folder = self.entry_out_path.get()
        if not out_folder:
            out_folder = dwg_path.parent
        else:
            out_folder = Path(out_folder)
            
        out_filename = self.entry_filename.get()
        if not out_filename:
            out_filename = "spec.xlsx"
            
        out_path = out_folder / out_filename

        self.btn_process.configure(state="disabled", text="Processing...")
        self.log("-" * 40)
        self.log(f"Starting processing for: {dwg_path.name}")
        self.log(f"Config: {config_name}")
        self.log(f"Output: {out_path}")
        
        try:
            # Import core logic here to avoid UI freeze during import if heavy
            from technopan_spec.config import load_config
            from technopan_spec.dxf import extract_panels_from_dxf
            from technopan_spec.spec import build_panel_rows, write_spec_xlsx
            
            self.log("Loading configuration...")
            cfg = load_config(config_path)
            
            self.log("Extracting panels (this may take a while)...")
            panels = extract_panels_from_dxf(dwg_path, cfg)
            self.log(f"Extracted {len(panels)} panels/items.")
            
            if not panels:
                self.log("Warning: No panels found! Check your configuration and layers.")
            
            self.log("Building specification rows...")
            rows = build_panel_rows(panels)
            
            self.log("Writing Excel file...")
            write_spec_xlsx(out_path, rows, title="Спецификация")
            
            self.log(f"Success! File saved to:\n{out_path}")
            messagebox.showinfo("Success", f"Specification generated successfully!\n\n{out_path}")
            
        except Exception as e:
            self.log(f"ERROR: {str(e)}")
            import traceback
            self.log(traceback.format_exc())
            messagebox.showerror("Error", f"An error occurred:\n{str(e)}")
            
        finally:
            self.btn_process.configure(state="normal", text="GENERATE SPECIFICATION")

if __name__ == "__main__":
    app = App()
    app.mainloop()
