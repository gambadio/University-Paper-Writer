import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import requests
import os
import PyPDF2
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt, Inches, Cm
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.section import WD_SECTION_START
import re
import json
from datetime import datetime

class ScriptsWindow(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("Manage Scripts")
        self.geometry("600x400")
        self.grab_set()  # Make the window modal
        self.create_widgets()

    def create_widgets(self):
        self.scripts_listbox = tk.Listbox(self, width=80, height=15)
        self.scripts_listbox.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

        buttons_frame = ttk.Frame(self)
        buttons_frame.pack(pady=10, fill=tk.X)

        upload_btn = ttk.Button(buttons_frame, text="Upload Script", command=self.upload_script)
        upload_btn.pack(side=tk.LEFT, padx=5)

        move_up_btn = ttk.Button(buttons_frame, text="Move Up", command=self.move_up)
        move_up_btn.pack(side=tk.LEFT, padx=5)

        move_down_btn = ttk.Button(buttons_frame, text="Move Down", command=self.move_down)
        move_down_btn.pack(side=tk.LEFT, padx=5)

        delete_btn = ttk.Button(buttons_frame, text="Delete Selected", command=self.delete_selected)
        delete_btn.pack(side=tk.LEFT, padx=5)

        save_btn = ttk.Button(buttons_frame, text="Save Changes", command=self.save_changes)
        save_btn.pack(side=tk.LEFT, padx=5)

        close_btn = ttk.Button(buttons_frame, text="Close", command=self.destroy)
        close_btn.pack(side=tk.RIGHT, padx=5)

        self.update_listbox()

    def upload_script(self):
        file_types = [("PDF files", "*.pdf"), ("Text files", "*.txt"), ("All files", "*.*")]
        file_paths = filedialog.askopenfilenames(title="Select Script(s) or Paper(s)", filetypes=file_types)
        for file_path in file_paths:
            self.parent.upload_script(file_path)
        self.update_listbox()

    def move_up(self):
        selection = self.scripts_listbox.curselection()
        if selection:
            index = selection[0]
            if index > 0:
                self.parent.scripts[index], self.parent.scripts[index - 1] = self.parent.scripts[index - 1], self.parent.scripts[index]
                self.update_listbox()
                self.scripts_listbox.select_set(index - 1)

    def move_down(self):
        selection = self.scripts_listbox.curselection()
        if selection:
            index = selection[0]
            if index < len(self.parent.scripts) - 1:
                self.parent.scripts[index], self.parent.scripts[index + 1] = self.parent.scripts[index + 1], self.parent.scripts[index]
                self.update_listbox()
                self.scripts_listbox.select_set(index + 1)

    def delete_selected(self):
        selection = self.scripts_listbox.curselection()
        if selection:
            index = selection[0]
            del self.parent.scripts[index]
            self.update_listbox()

    def save_changes(self):
        self.parent.update_system_prompt()
        self.parent.save_all_settings()
        messagebox.showinfo("Success", "Changes saved and system prompt updated.")

    def update_listbox(self):
        self.scripts_listbox.delete(0, tk.END)
        for script in self.parent.scripts:
            self.scripts_listbox.insert(tk.END, script[0])

    def destroy(self):
        self.grab_release()  # Release the modal state before destroying
        super().destroy()

class InstructionsWindow(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("Manage Instructions")
        self.geometry("600x400")
        self.grab_set()  # Make the window modal
        self.create_widgets()

    def create_widgets(self):
        self.instructions_listbox = tk.Listbox(self, width=80, height=15)
        self.instructions_listbox.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

        buttons_frame = ttk.Frame(self)
        buttons_frame.pack(pady=10, fill=tk.X)

        upload_btn = ttk.Button(buttons_frame, text="Upload Instruction", command=self.upload_instruction)
        upload_btn.pack(side=tk.LEFT, padx=5)

        move_up_btn = ttk.Button(buttons_frame, text="Move Up", command=self.move_up)
        move_up_btn.pack(side=tk.LEFT, padx=5)

        move_down_btn = ttk.Button(buttons_frame, text="Move Down", command=self.move_down)
        move_down_btn.pack(side=tk.LEFT, padx=5)

        delete_btn = ttk.Button(buttons_frame, text="Delete Selected", command=self.delete_selected)
        delete_btn.pack(side=tk.LEFT, padx=5)

        save_btn = ttk.Button(buttons_frame, text="Save Changes", command=self.save_changes)
        save_btn.pack(side=tk.LEFT, padx=5)

        close_btn = ttk.Button(buttons_frame, text="Close", command=self.destroy)
        close_btn.pack(side=tk.RIGHT, padx=5)

        self.update_listbox()

    def upload_instruction(self):
        file_types = [("PDF files", "*.pdf"), ("Text files", "*.txt"), ("All files", "*.*")]
        file_paths = filedialog.askopenfilenames(title="Select Instruction File(s)", filetypes=file_types)
        for file_path in file_paths:
            self.parent.upload_instruction(file_path)
        self.update_listbox()

    def move_up(self):
        selection = self.instructions_listbox.curselection()
        if selection:
            index = selection[0]
            if index > 0:
                self.parent.instructions[index], self.parent.instructions[index - 1] = self.parent.instructions[index - 1], self.parent.instructions[index]
                self.update_listbox()
                self.instructions_listbox.select_set(index - 1)

    def move_down(self):
        selection = self.instructions_listbox.curselection()
        if selection:
            index = selection[0]
            if index < len(self.parent.instructions) - 1:
                self.parent.instructions[index], self.parent.instructions[index + 1] = self.parent.instructions[index + 1], self.parent.instructions[index]
                self.update_listbox()
                self.instructions_listbox.select_set(index + 1)

    def delete_selected(self):
        selection = self.instructions_listbox.curselection()
        if selection:
            index = selection[0]
            del self.parent.instructions[index]
            self.update_listbox()

    def save_changes(self):
        self.parent.update_system_prompt()
        self.parent.save_all_settings()
        messagebox.showinfo("Success", "Changes saved and system prompt updated.")

    def update_listbox(self):
        self.instructions_listbox.delete(0, tk.END)
        for instruction in self.parent.instructions:
            self.instructions_listbox.insert(tk.END, instruction[0])

    def destroy(self):
        self.grab_release()  # Release the modal state before destroying
        super().destroy()

class FormattingWindow(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("Formatting Options")
        self.geometry("400x500")
        self.create_widgets()

    def create_widgets(self):
        main_frame = ttk.Frame(self, padding="20")
        main_frame.pack(expand=True)

        ttk.Label(main_frame, text="Font Name:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.font_name_var = tk.StringVar(value=self.parent.font_name)
        font_options = ["Times New Roman", "Arial", "Calibri", "Cambria", "Verdana"]
        self.font_name_combobox = ttk.Combobox(main_frame, textvariable=self.font_name_var, values=font_options, state="readonly", width=30)
        self.font_name_combobox.grid(row=0, column=1, sticky=tk.W+tk.E, pady=5, padx=5)

        ttk.Label(main_frame, text="Line Spacing:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.line_spacing_var = tk.StringVar(value=self.parent.line_spacing)
        line_spacing_options = ["Single", "1.5 lines", "Double"]
        self.line_spacing_combobox = ttk.Combobox(main_frame, textvariable=self.line_spacing_var, values=line_spacing_options, state="readonly", width=30)
        self.line_spacing_combobox.grid(row=1, column=1, sticky=tk.W+tk.E, pady=5, padx=5)

        ttk.Label(main_frame, text="Font Sizes (pt):").grid(row=2, column=0, sticky=tk.W, pady=5)
        sizes_frame = ttk.Frame(main_frame)
        sizes_frame.grid(row=2, column=1, sticky=tk.W+tk.E, pady=5, padx=5)

        ttk.Label(sizes_frame, text="Normal Text:").grid(row=0, column=0, sticky=tk.W)
        self.font_size_normal_var = tk.IntVar(value=self.parent.font_size_normal)
        self.font_size_normal_spinbox = ttk.Spinbox(sizes_frame, from_=8, to=72, textvariable=self.font_size_normal_var, width=5)
        self.font_size_normal_spinbox.grid(row=0, column=1, sticky=tk.W+tk.E)

        ttk.Label(sizes_frame, text="Heading 1:").grid(row=1, column=0, sticky=tk.W)
        self.font_size_heading1_var = tk.IntVar(value=self.parent.font_size_heading1)
        self.font_size_heading1_spinbox = ttk.Spinbox(sizes_frame, from_=8, to=72, textvariable=self.font_size_heading1_var, width=5)
        self.font_size_heading1_spinbox.grid(row=1, column=1, sticky=tk.W+tk.E)

        ttk.Label(sizes_frame, text="Heading 2:").grid(row=2, column=0, sticky=tk.W)
        self.font_size_heading2_var = tk.IntVar(value=self.parent.font_size_heading2)
        self.font_size_heading2_spinbox = ttk.Spinbox(sizes_frame, from_=8, to=72, textvariable=self.font_size_heading2_var, width=5)
        self.font_size_heading2_spinbox.grid(row=2, column=1, sticky=tk.W+tk.E)

        ttk.Label(sizes_frame, text="Heading 3:").grid(row=3, column=0, sticky=tk.W)
        self.font_size_heading3_var = tk.IntVar(value=self.parent.font_size_heading3)
        self.font_size_heading3_spinbox = ttk.Spinbox(sizes_frame, from_=8, to=72, textvariable=self.font_size_heading3_var, width=5)
        self.font_size_heading3_spinbox.grid(row=3, column=1, sticky=tk.W+tk.E)

        ttk.Label(main_frame, text="Margins (cm):").grid(row=3, column=0, sticky=tk.W, pady=5)
        margins_frame = ttk.Frame(main_frame)
        margins_frame.grid(row=3, column=1, sticky=tk.W+tk.E, pady=5, padx=5)

        ttk.Label(margins_frame, text="Top:").grid(row=0, column=0, sticky=tk.W)
        self.margin_top_var = tk.DoubleVar(value=self.parent.margin_top)
        self.margin_top_spinbox = ttk.Spinbox(margins_frame, from_=0, to=10, increment=0.1, textvariable=self.margin_top_var, width=5)
        self.margin_top_spinbox.grid(row=0, column=1, sticky=tk.W+tk.E)

        ttk.Label(margins_frame, text="Bottom:").grid(row=1, column=0, sticky=tk.W)
        self.margin_bottom_var = tk.DoubleVar(value=self.parent.margin_bottom)
        self.margin_bottom_spinbox = ttk.Spinbox(margins_frame, from_=0, to=10, increment=0.1, textvariable=self.margin_bottom_var, width=5)
        self.margin_bottom_spinbox.grid(row=1, column=1, sticky=tk.W+tk.E)

        ttk.Label(margins_frame, text="Left:").grid(row=2, column=0, sticky=tk.W)
        self.margin_left_var = tk.DoubleVar(value=self.parent.margin_left)
        self.margin_left_spinbox = ttk.Spinbox(margins_frame, from_=0, to=10, increment=0.1, textvariable=self.margin_left_var, width=5)
        self.margin_left_spinbox.grid(row=2, column=1, sticky=tk.W+tk.E)

        ttk.Label(margins_frame, text="Right:").grid(row=3, column=0, sticky=tk.W)
        self.margin_right_var = tk.DoubleVar(value=self.parent.margin_right)
        self.margin_right_spinbox = ttk.Spinbox(margins_frame, from_=0, to=10, increment=0.1, textvariable=self.margin_right_var, width=5)
        self.margin_right_spinbox.grid(row=3, column=1, sticky=tk.W+tk.E)

        save_btn = ttk.Button(main_frame, text="Save", command=self.save_formatting)
        save_btn.grid(row=4, column=0, pady=10)

        default_btn = ttk.Button(main_frame, text="Back to Default", command=self.set_default_formatting)
        default_btn.grid(row=4, column=1, pady=10)

    def save_formatting(self):
        self.parent.font_name = self.font_name_var.get()
        self.parent.line_spacing = self.line_spacing_var.get()
        self.parent.font_size_normal = self.font_size_normal_var.get()
        self.parent.font_size_heading1 = self.font_size_heading1_var.get()
        self.parent.font_size_heading2 = self.font_size_heading2_var.get()
        self.parent.font_size_heading3 = self.font_size_heading3_var.get()
        self.parent.margin_top = self.margin_top_var.get()
        self.parent.margin_bottom = self.margin_bottom_var.get()
        self.parent.margin_left = self.margin_left_var.get()
        self.parent.margin_right = self.margin_right_var.get()
        self.parent.save_all_settings()
        self.destroy()

    def set_default_formatting(self):
        self.font_name_var.set("Times New Roman")
        self.line_spacing_var.set("1.5 lines")
        self.font_size_normal_var.set(12)
        self.font_size_heading1_var.set(16)
        self.font_size_heading2_var.set(14)
        self.font_size_heading3_var.set(12)
        self.margin_top_var.set(2.0)
        self.margin_bottom_var.set(2.0)
        self.margin_left_var.set(2.0)
        self.margin_right_var.set(2.0)

class SettingsWindow(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("Settings")
        self.geometry("400x300")
        self.grab_set()  # Make the window modal
        self.create_widgets()

    def create_widgets(self):
        main_frame = ttk.Frame(self, padding="20")
        main_frame.pack(expand=True, fill=tk.BOTH)

        ttk.Label(main_frame, text="API Key:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.api_key_entry = ttk.Entry(main_frame, width=40, show="*")
        self.api_key_entry.grid(row=0, column=1, pady=5)
        self.api_key_entry.insert(0, self.parent.api_key)

        ttk.Label(main_frame, text="First Name:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.first_name_entry = ttk.Entry(main_frame, width=40)
        self.first_name_entry.grid(row=1, column=1, pady=5)
        self.first_name_entry.insert(0, self.parent.first_name)

        ttk.Label(main_frame, text="Last Name:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.last_name_entry = ttk.Entry(main_frame, width=40)
        self.last_name_entry.grid(row=2, column=1, pady=5)
        self.last_name_entry.insert(0, self.parent.last_name)

        ttk.Label(main_frame, text="Date (YYYY-MM-DD):").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.date_entry = ttk.Entry(main_frame, width=40)
        self.date_entry.grid(row=3, column=1, pady=5)
        self.date_entry.insert(0, self.parent.date)

        save_btn = ttk.Button(main_frame, text="Save", command=self.save_settings)
        save_btn.grid(row=4, column=0, pady=20)

        close_btn = ttk.Button(main_frame, text="Close", command=self.destroy)
        close_btn.grid(row=4, column=1, pady=20)

    def save_settings(self):
        self.parent.api_key = self.api_key_entry.get().strip()
        self.parent.first_name = self.first_name_entry.get().strip()
        self.parent.last_name = self.last_name_entry.get().strip()
        self.parent.date = self.date_entry.get().strip()
        self.parent.save_all_settings()
        messagebox.showinfo("Settings", "Settings saved successfully.")

    def destroy(self):
        self.grab_release()
        super().destroy()

class ClaudeApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("University Paper Generator")
        self.geometry("800x600")

        # Initialize variables
        self.api_key = ""
        self.default_system_prompt = (
            "You are to write a university paper based on the provided scientific papers and study scripts. "
            "Determine an appropriate title for the paper. The paper should be formatted as a real university paper suitable for submission, "
            "including chapters, sections, headings, and citations in Harvard style. "
            "Include a bibliography at the end. Do not include a table of contents. "
            "Take content from the scripts provided below and cite them appropriately. "
            "Write in first person singular, as if by {first_name} {last_name}. "
            "The paper should be dated {date}. "
            "Remember to sometimes make minor spelling mistakes as a student around B2/C1 English level might. "
            "Your specific instructions are:\n\n"
            "{instructions}\n\n"
            "The scripts are provided in the following format:\n\n"
            "{scripts}\n\n"
            "Please output the paper in Markdown format with clear markers for headings and sections. "
            "Use '#' for main headings, '##' for subheadings, and '###' for sub-subheadings. "
            "Use **bold** and *italic* text where appropriate. Include bullet points and numbered lists if necessary. "
            "Ensure that citations are properly formatted in Harvard style and included within the text. "
            "At the beginning of the paper, include a title page containing the paper's title, your name, and date. "
            "Enclose the title page content between '####TITLE PAGE####' and '####END TITLE PAGE####'.\n\n"
            "# Introduction\n"
            "...\n\n"
            "# Conclusion\n"
            "...\n\n"
            "# Bibliography\n"
            "...\n\n"
            "IMPORTANT: IF THE INSTRUCTIONS OR DETAILS SUCH AS TITLE PAGE, "
            "BIBLIOGRAPHY, CITATION STYLE, ETC., ARE PROVIDED REGARDING STRUCTURING THE PAPER, "
            "PLEASE FOLLOW THOSE INSTEAD OF THE ONES LISTED ABOVE. MAKE USE OF FULL MAX TOKEN OUTPUT OF 8192"
        )
        self.system_prompt = self.default_system_prompt
        self.scripts = []
        self.instructions = []
        self.first_name = ""
        self.last_name = ""
        self.date = datetime.now().strftime("%Y-%m-%d")

        # Formatting options
        self.font_name = "Times New Roman"
        self.font_size_normal = 12
        self.font_size_heading1 = 16
        self.font_size_heading2 = 14
        self.font_size_heading3 = 12
        self.line_spacing = "1.5 lines"
        self.margin_top = 2.0
        self.margin_bottom = 2.0
        self.margin_left = 2.0
        self.margin_right = 2.0

        # Load saved settings
        self.load_all_settings()

        # Build UI
        self.create_widgets()

    def create_widgets(self):
        style = ttk.Style()
        style.theme_use('clam')

        # Define styles for custom button colors
        style.configure("Green.TButton", foreground="black", background="lightgreen")
        style.map("Green.TButton",
                  background=[('active', 'green')])

        style.configure("Blue.TButton", foreground="black", background="lightblue")
        style.map("Blue.TButton",
                  background=[('active', 'blue')])

        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Top buttons frame
        top_buttons_frame = ttk.Frame(main_frame)
        top_buttons_frame.grid(row=0, column=0, columnspan=3, sticky=tk.W+tk.E, pady=5)

        # Settings button (gear symbol)
        settings_btn = ttk.Button(top_buttons_frame, text="⚙", command=self.open_settings_window)
        settings_btn.pack(side=tk.LEFT, padx=5)

        # Formatting options button
        formatting_btn = ttk.Button(top_buttons_frame, text="Formatting Options", command=self.open_formatting_window)
        formatting_btn.pack(side=tk.LEFT, padx=5)

        # System Prompt
        ttk.Label(main_frame, text="System Prompt:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.system_prompt_text = tk.Text(main_frame, wrap=tk.WORD, height=10)
        self.system_prompt_text.grid(row=1, column=1, columnspan=2, pady=5, padx=5, sticky=tk.W+tk.E+tk.N+tk.S)
        self.system_prompt_text.insert(tk.END, self.system_prompt)

        # Buttons
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.grid(row=2, column=0, columnspan=3, pady=10)

        manage_scripts_btn = ttk.Button(buttons_frame, text="Manage Scripts", command=self.open_scripts_window)
        manage_scripts_btn.pack(side=tk.LEFT, padx=5)

        manage_instructions_btn = ttk.Button(buttons_frame, text="Manage Instructions", command=self.open_instructions_window)
        manage_instructions_btn.pack(side=tk.LEFT, padx=5)

        generate_paper_btn = ttk.Button(buttons_frame, text="Generate Paper", command=self.send_request, style="Green.TButton")
        generate_paper_btn.pack(side=tk.LEFT, padx=5)

        save_output_btn = ttk.Button(buttons_frame, text="Save Output", command=self.save_output, style="Blue.TButton")
        save_output_btn.pack(side=tk.LEFT, padx=5)

        save_settings_btn = ttk.Button(buttons_frame, text="Save All Settings", command=self.save_all_settings)
        save_settings_btn.pack(side=tk.LEFT, padx=5)

        reset_system_prompt_btn = ttk.Button(buttons_frame, text="Reset System Prompt", command=self.reset_system_prompt)
        reset_system_prompt_btn.pack(side=tk.LEFT, padx=5)

        # Output Text Area
        self.output_text = tk.Text(main_frame, wrap=tk.WORD, height=20)
        self.output_text.grid(row=3, column=0, columnspan=3, pady=10, sticky=tk.W+tk.E+tk.N+tk.S)

        # Configure grid weights
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(1, weight=1)
        main_frame.rowconfigure(3, weight=2)

    def open_settings_window(self):
        SettingsWindow(self)

    def open_formatting_window(self):
        FormattingWindow(self)

    def open_scripts_window(self):
        ScriptsWindow(self)

    def open_instructions_window(self):
        InstructionsWindow(self)

    def upload_script(self, file_path):
        file_name = os.path.basename(file_path)
        file_extension = os.path.splitext(file_path)[1].lower()

        if file_extension == '.pdf':
            # Extract text from PDF
            try:
                with open(file_path, 'rb') as f:
                    reader = PyPDF2.PdfReader(f)
                    text = ""
                    for page in reader.pages:
                        page_text = page.extract_text()
                        if page_text:
                            text += page_text
                self.scripts.append((file_name, text))
            except Exception as e:
                messagebox.showerror("Error", f"Error reading PDF file {file_name}: {e}")
        else:
            # Read text file
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    text = f.read()
                self.scripts.append((file_name, text))
            except UnicodeDecodeError:
                try:
                    with open(file_path, 'r', encoding='latin-1') as f:
                        text = f.read()
                    self.scripts.append((file_name, text))
                except Exception as e:
                    messagebox.showerror("Error", f"Error reading text file {file_name}: {e}")

    def upload_instruction(self, file_path):
        file_name = os.path.basename(file_path)
        file_extension = os.path.splitext(file_path)[1].lower()

        if file_extension == '.pdf':
            # Extract text from PDF
            try:
                with open(file_path, 'rb') as f:
                    reader = PyPDF2.PdfReader(f)
                    text = ""
                    for page in reader.pages:
                        page_text = page.extract_text()
                        if page_text:
                            text += page_text
                self.instructions.append((file_name, text))
            except Exception as e:
                messagebox.showerror("Error", f"Error reading PDF file {file_name}: {e}")
        else:
            # Read text file
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    text = f.read()
                self.instructions.append((file_name, text))
            except UnicodeDecodeError:
                try:
                    with open(file_path, 'r', encoding='latin-1') as f:
                        text = f.read()
                    self.instructions.append((file_name, text))
                except Exception as e:
                    messagebox.showerror("Error", f"Error reading text file {file_name}: {e}")

    def update_system_prompt(self):
        formatted_instructions = "\n\n".join([f"Instruction {i+1} ({name}):\n{content}\n!!!this is the next document!!!" for i, (name, content) in enumerate(self.instructions)])
        formatted_scripts = "\n\n".join([f"Script {i+1} ({name}):\n{content}\n!!!this is the next document!!!" for i, (name, content) in enumerate(self.scripts)])
        self.system_prompt = self.system_prompt_text.get(1.0, tk.END).format(
            scripts=formatted_scripts,
            instructions=formatted_instructions,
            first_name=self.first_name,
            last_name=self.last_name,
            date=self.date
        )

    def send_request(self):
        if not self.instructions:
            messagebox.showerror("Error", "Please upload instruction files first.")
            return
        if not self.scripts:
            messagebox.showerror("Error", "Please upload script files first.")
            return
        if not self.first_name or not self.last_name:
            messagebox.showerror("Error", "Please enter your first and last name in the settings.")
            return
        if not self.api_key:
            messagebox.showerror("Error", "Please enter your API key in the settings.")
            return

        self.update_system_prompt()

        # Prepare the messages
        messages = [
            {"role": "user", "content": self.system_prompt}
        ]

        # API call parameters
        api_url = "https://api.anthropic.com/v1/messages"
        headers = {
            "x-api-key": self.api_key,
            "anthropic-version": "2023-06-01",
            "content-type": "application/json",
            "anthropic-beta": "max-tokens-3-5-sonnet-2024-07-15"
        }
        data = {
            "model": "claude-3-5-sonnet-20240620",
            "max_tokens": 8192,
            "messages": messages
        }

        try:
            # Show a loading message
            self.output_text.delete(1.0, tk.END)
            self.output_text.insert(tk.END, "Generating response, please wait...")
            self.update_idletasks()

            response = requests.post(api_url, headers=headers, json=data)

            if response.status_code == 200:
                result = response.json()
                content = result['content'][0]['text']
                response_text = content.strip()
                self.output_text.delete(1.0, tk.END)
                self.output_text.insert(tk.END, response_text)
                messagebox.showinfo("Success", "Paper generated.")
            else:
                error_message = response.text
                self.output_text.delete(1.0, tk.END)
                messagebox.showerror("Error", f"API Error {response.status_code}: {error_message}")
        except Exception as e:
            self.output_text.delete(1.0, tk.END)
            messagebox.showerror("Error", f"Error making API request: {e}")

    def save_output(self):
        output = self.output_text.get(1.0, tk.END).strip()
        if not output:
            messagebox.showerror("Error", "No output to save.")
            return

        save_path = filedialog.asksaveasfilename(title="Save Output as Word File", defaultextension=".docx", filetypes=[("Word Document", "*.docx")])
        if save_path:
            try:
                document = Document()

                # Document settings
                section = document.sections[0]
                section.page_height = Inches(11)
                section.page_width = Inches(8.5)
                section.left_margin = Cm(self.margin_left)
                section.right_margin = Cm(self.margin_right)
                section.top_margin = Cm(self.margin_top)
                section.bottom_margin = Cm(self.margin_bottom)

                # Define styles
                styles = document.styles

                # Title style
                style_title = styles.add_style('TitleStyle', WD_STYLE_TYPE.PARAGRAPH)
                style_title.font.size = Pt(self.font_size_heading1)
                style_title.font.bold = True
                style_title.font.name = self.font_name

                # Heading 1 style
                style_heading1 = styles.add_style('Heading1Custom', WD_STYLE_TYPE.PARAGRAPH)
                style_heading1.base_style = styles['Heading 1']
                style_heading1.font.size = Pt(self.font_size_heading1)
                style_heading1.font.bold = True
                style_heading1.font.name = self.font_name

                # Heading 2 style
                style_heading2 = styles.add_style('Heading2Custom', WD_STYLE_TYPE.PARAGRAPH)
                style_heading2.base_style = styles['Heading 2']
                style_heading2.font.size = Pt(self.font_size_heading2)
                style_heading2.font.bold = True
                style_heading2.font.name = self.font_name

                # Heading 3 style
                style_heading3 = styles.add_style('Heading3Custom', WD_STYLE_TYPE.PARAGRAPH)
                style_heading3.base_style = styles['Heading 3']
                style_heading3.font.size = Pt(self.font_size_heading3)
                style_heading3.font.bold = True
                style_heading3.font.name = self.font_name

                # Normal text style
                style_normal = styles['Normal']
                style_normal.font.size = Pt(self.font_size_normal)
                style_normal.font.name = self.font_name

                # Set line spacing for paragraph styles
                line_spacing = 1.0 if self.line_spacing == "Single" else 1.5 if self.line_spacing == "1.5 lines" else 2.0
                for style in [style_normal, style_title, style_heading1, style_heading2, style_heading3]:
                    style.paragraph_format.line_spacing = Pt(line_spacing * 12)  # Approximation

                # Process the content
                paragraphs = output.strip().split('\n')
                i = 0
                while i < len(paragraphs):
                    para = paragraphs[i].strip()
                    if not para:
                        i += 1
                        continue

                    # Handle Title Page
                    if para == '####TITLE PAGE####':
                        title_page_content = []
                        i += 1
                        while i < len(paragraphs) and paragraphs[i].strip() != '####END TITLE PAGE####':
                            title_page_content.append(paragraphs[i].strip())
                            i += 1
                        # Add Title Page
                        for line in title_page_content:
                            p = document.add_paragraph(line, style='TitleStyle')
                            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        i += 1
                        continue

                    # Handle Headings
                    if para.startswith('# '):
                        document.add_heading(para[2:].strip(), level=1)
                    elif para.startswith('## '):
                        document.add_heading(para[3:].strip(), level=2)
                    elif para.startswith('### '):
                        document.add_heading(para[4:].strip(), level=3)
                    elif re.match(r'^\d+\.', para):
                        # Numbered list
                        p = document.add_paragraph(style='List Number')
                        self._add_runs(p, para)
                    elif para.startswith('- '):
                        # Bullet list
                        p = document.add_paragraph(style='List Bullet')
                        self._add_runs(p, para[2:].strip())
                    elif para.startswith('|') and para.endswith('|'):
                        # Possible Markdown Table
                        table_lines = [para]
                        i += 1
                        while i < len(paragraphs) and paragraphs[i].strip().startswith('|') and paragraphs[i].strip().endswith('|'):
                            table_lines.append(paragraphs[i].strip())
                            i += 1
                        i -= 1  # Adjust for the outer loop
                        table = self._parse_markdown_table(table_lines)
                        if table:
                            document.add_table(rows=0, cols=len(table[0]))
                            word_table = document.add_table(rows=len(table), cols=len(table[0]))
                            for row_idx, row in enumerate(table):
                                for col_idx, cell in enumerate(row):
                                    word_table.cell(row_idx, col_idx).text = cell
                            for row in word_table.rows:
                                for cell in row.cells:
                                    for paragraph in cell.paragraphs:
                                        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    else:
                        # Regular paragraph
                        p = document.add_paragraph(style='Normal')
                        self._add_runs(p, para)
                    i += 1

                # Add page numbers
                self._add_page_numbers(document.sections[0])

                # Save the document
                document.save(save_path)
                messagebox.showinfo("Success", f"Output saved to {save_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Error saving Word file: {e}")

    def _parse_markdown_table(self, table_lines):
        try:
            # Split each line by '|' and remove empty strings
            table = []
            for line in table_lines:
                cells = [cell.strip() for cell in line.strip('|').split('|')]
                table.append(cells)
            return table
        except Exception:
            return None

    def _add_page_numbers(self, section):
        footer = section.footer
        paragraph = footer.paragraphs[0]
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = paragraph.add_run()
        fldSimple = OxmlElement('w:fldSimple')
        fldSimple.set(qn('w:instr'), 'PAGE')
        run._r.append(fldSimple)

    def _add_runs(self, paragraph, text):
        # This method adds runs to the paragraph, handling bold and italic text
        parts = re.split(r'(\*\*.*?\*\*|\*.*?\*)', text)
        for part in parts:
            if part.startswith('**') and part.endswith('**'):
                run = paragraph.add_run(part[2:-2])
                run.bold = True
            elif part.startswith('*') and part.endswith('*'):
                run = paragraph.add_run(part[1:-1])
                run.italic = True
            else:
                paragraph.add_run(part)

    def load_all_settings(self):
        try:
            with open('claude_app_settings.json', 'r') as f:
                settings = json.load(f)
                self.api_key = settings.get('api_key', '')
                self.first_name = settings.get('first_name', '')
                self.last_name = settings.get('last_name', '')
                self.font_name = settings.get('font_name', 'Times New Roman')
                self.font_size_normal = settings.get('font_size_normal', 12)
                self.font_size_heading1 = settings.get('font_size_heading1', 16)
                self.font_size_heading2 = settings.get('font_size_heading2', 14)
                self.font_size_heading3 = settings.get('font_size_heading3', 12)
                self.line_spacing = settings.get('line_spacing', '1.5 lines')
                self.margin_top = settings.get('margin_top', 2.0)
                self.margin_bottom = settings.get('margin_bottom', 2.0)
                self.margin_left = settings.get('margin_left', 2.0)
                self.margin_right = settings.get('margin_right', 2.0)
                self.system_prompt = settings.get('system_prompt', self.default_system_prompt)
        except FileNotFoundError:
            pass  # It's okay if the file doesn't exist yet

    def save_all_settings(self):
        self.system_prompt = self.system_prompt_text.get(1.0, tk.END).strip()
        settings = {
            'api_key': self.api_key,
            'first_name': self.first_name,
            'last_name': self.last_name,
            'font_name': self.font_name,
            'font_size_normal': self.font_size_normal,
            'font_size_heading1': self.font_size_heading1,
            'font_size_heading2': self.font_size_heading2,
            'font_size_heading3': self.font_size_heading3,
            'line_spacing': self.line_spacing,
            'margin_top': self.margin_top,
            'margin_bottom': self.margin_bottom,
            'margin_left': self.margin_left,
            'margin_right': self.margin_right,
            'system_prompt': self.system_prompt
        }
        try:
            with open('claude_app_settings.json', 'w') as f:
                json.dump(settings, f)
            messagebox.showinfo("Settings", "All settings saved successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Error saving settings: {e}")

    def reset_system_prompt(self):
        self.system_prompt_text.delete(1.0, tk.END)
        self.system_prompt_text.insert(tk.END, self.default_system_prompt)
        messagebox.showinfo("System Prompt", "System prompt reset to default.")

if __name__ == "__main__":
    app = ClaudeApp()
    app.mainloop()