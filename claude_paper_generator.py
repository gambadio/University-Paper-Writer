import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox, ttk
import requests
import os
import PyPDF2
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt, Inches
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import re
import json
import markdown
from bs4 import BeautifulSoup
from datetime import datetime

class ScriptsWindow(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("Manage Scripts")
        self.geometry("500x400")
        
        self.create_widgets()

    def create_widgets(self):
        self.scripts_listbox = tk.Listbox(self, width=70, height=15)
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
                # Swap in parent list
                self.parent.scripts[index], self.parent.scripts[index - 1] = self.parent.scripts[index - 1], self.parent.scripts[index]
                # Update listbox
                self.update_listbox()
                # Reselect the moved item
                self.scripts_listbox.select_set(index - 1)

    def move_down(self):
        selection = self.scripts_listbox.curselection()
        if selection:
            index = selection[0]
            if index < len(self.parent.scripts) - 1:
                # Swap in parent list
                self.parent.scripts[index], self.parent.scripts[index + 1] = self.parent.scripts[index + 1], self.parent.scripts[index]
                # Update listbox
                self.update_listbox()
                # Reselect the moved item
                self.scripts_listbox.select_set(index + 1)

    def delete_selected(self):
        selection = self.scripts_listbox.curselection()
        if selection:
            index = selection[0]
            del self.parent.scripts[index]
            self.update_listbox()

    def update_listbox(self):
        self.scripts_listbox.delete(0, tk.END)
        for script in self.parent.scripts:
            self.scripts_listbox.insert(tk.END, script[0])

class InstructionsWindow(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("Manage Instructions")
        self.geometry("500x400")
        
        self.create_widgets()

    def create_widgets(self):
        self.instructions_listbox = tk.Listbox(self, width=70, height=15)
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
                # Swap in parent list
                self.parent.instructions[index], self.parent.instructions[index - 1] = self.parent.instructions[index - 1], self.parent.instructions[index]
                # Update listbox
                self.update_listbox()
                # Reselect the moved item
                self.instructions_listbox.select_set(index - 1)

    def move_down(self):
        selection = self.instructions_listbox.curselection()
        if selection:
            index = selection[0]
            if index < len(self.parent.instructions) - 1:
                # Swap in parent list
                self.parent.instructions[index], self.parent.instructions[index + 1] = self.parent.instructions[index + 1], self.parent.instructions[index]
                # Update listbox
                self.update_listbox()
                # Reselect the moved item
                self.instructions_listbox.select_set(index + 1)

    def delete_selected(self):
        selection = self.instructions_listbox.curselection()
        if selection:
            index = selection[0]
            del self.parent.instructions[index]
            self.update_listbox()

    def update_listbox(self):
        self.instructions_listbox.delete(0, tk.END)
        for instruction in self.parent.instructions:
            self.instructions_listbox.insert(tk.END, instruction[0])

class ClaudeApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("University Paper Generator")
        self.geometry("800x700")
        
        # Initialize variables
        self.api_key = ""
        self.system_prompt = (
            "You are to write a university paper based on the provided scientific papers and study scripts. "
            "Determine an appropriate title for the paper. The paper should be perfectly formatted as a real university paper suitable for submission, "
            "including chapters, sections, headings, and citations in Harvard style. "
            "Include a table of contents at the beginning and a bibliography at the end. "
            "Take content from the scripts provided below and cite them appropriately. "
            "Write in first person singular, as if by {first_name} {last_name}. "
            "The paper should be dated {date}. "
            "Remember to sometimes make minor spelling mistakes as a student around B2/C1 English level might. "
            "Your specific instructions are:\n\n"
            "{instructions}\n\n"
            "The scripts are provided in the following format:\n\n"
            "{scripts}\n\n"
            "Please output the paper in detailed Markdown format, using appropriate Markdown syntax for proper styling. Use headings, subheadings, bold and italic text where appropriate. Include bullet points, numbered lists, tables, and images if necessary. "
            "Ensure that citations are properly formatted in Harvard style and included within the text. "
            "Make sure the Markdown output is detailed and properly formatted, so that when converted to a Word document, the paper looks professional.\n\n"
            "Use the following structure:\n\n"
            "# Title of the Paper\n\n"
            "## Abstract\n\n"
            "Abstract text...\n\n"
            "## Table of Contents\n\n"
            "1. Introduction\n"
            "2. Chapter 1: Chapter Title\n"
            "   - Subsection 1.1\n"
            "   - Subsection 1.2\n"
            "3. Chapter 2: Chapter Title\n"
            "4. Conclusion\n"
            "5. Bibliography\n\n"
            "## Introduction\n\n"
            "Introduction text...\n\n"
            "## Chapter 1: Chapter Title\n\n"
            "Chapter text...\n\n"
            "### Subsection 1.1\n\n"
            "Subsection text...\n\n"
            "### Subsection 1.2\n\n"
            "Subsection text...\n\n"
            "## Chapter 2: Chapter Title\n\n"
            "...\n\n"
            "## Conclusion\n\n"
            "Conclusion text...\n\n"
            "## Bibliography\n\n"
            "- Reference 1 in Harvard style\n"
            "- Reference 2 in Harvard style\n\n"
            "Ensure that all headings are properly marked with '#' symbols, and use bold, italics, bullet points, and other Markdown formatting where appropriate."
        )
        self.scripts = []
        self.instructions = []
        self.first_name = ""
        self.last_name = ""
        self.date = datetime.now().strftime("%Y-%m-%d")
        
        # Load saved settings
        self.load_settings()
        
        # Build UI
        self.create_widgets()
    
    def create_widgets(self):
        style = ttk.Style()
        style.theme_use('clam')

        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # API Key
        ttk.Label(main_frame, text="API Key:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.api_key_entry = ttk.Entry(main_frame, width=50, show='*')
        self.api_key_entry.grid(row=0, column=1, pady=5, padx=5, sticky=tk.W+tk.E)
        self.api_key_entry.insert(0, self.api_key)
        ttk.Button(main_frame, text="Save API Key", command=self.save_api_key).grid(row=0, column=2, pady=5, padx=5)

        # First Name
        ttk.Label(main_frame, text="First Name:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.first_name_entry = ttk.Entry(main_frame, width=50)
        self.first_name_entry.grid(row=1, column=1, pady=5, padx=5, sticky=tk.W+tk.E)
        self.first_name_entry.insert(0, self.first_name)

        # Last Name
        ttk.Label(main_frame, text="Last Name:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.last_name_entry = ttk.Entry(main_frame, width=50)
        self.last_name_entry.grid(row=2, column=1, pady=5, padx=5, sticky=tk.W+tk.E)
        self.last_name_entry.insert(0, self.last_name)

        # Date
        ttk.Label(main_frame, text="Date:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.date_entry = ttk.Entry(main_frame, width=50)
        self.date_entry.grid(row=3, column=1, pady=5, padx=5, sticky=tk.W+tk.E)
        self.date_entry.insert(0, self.date)

        # Instructions Button
        ttk.Button(main_frame, text="Manage Instructions", command=self.open_instructions_window).grid(row=4, column=0, columnspan=3, pady=10)

        # System Prompt
        ttk.Label(main_frame, text="System Prompt:").grid(row=5, column=0, sticky=tk.W, pady=5)
        self.system_prompt_text = tk.Text(main_frame, wrap=tk.WORD, height=10)
        self.system_prompt_text.grid(row=5, column=1, columnspan=2, pady=5, padx=5, sticky=tk.W+tk.E+tk.N+tk.S)
        self.system_prompt_text.insert(tk.END, self.system_prompt)

        # Buttons
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.grid(row=6, column=0, columnspan=3, pady=10)

        ttk.Button(buttons_frame, text="Manage Scripts", command=self.open_scripts_window).pack(side=tk.LEFT, padx=5)
        ttk.Button(buttons_frame, text="Generate Paper", command=self.send_request).pack(side=tk.LEFT, padx=5)
        ttk.Button(buttons_frame, text="Save Output", command=self.save_output).pack(side=tk.LEFT, padx=5)
        ttk.Button(buttons_frame, text="Save Settings", command=self.save_settings).pack(side=tk.LEFT, padx=5)

        # Output Text Area
        self.output_text = tk.Text(main_frame, wrap=tk.WORD, height=10)
        self.output_text.grid(row=7, column=0, columnspan=3, pady=10, sticky=tk.W+tk.E+tk.N+tk.S)

        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(5, weight=1)
        main_frame.rowconfigure(7, weight=1)

    def load_settings(self):
        try:
            with open('claude_app_settings.json', 'r') as f:
                settings = json.load(f)
                self.api_key = settings.get('api_key', '')
                self.first_name = settings.get('first_name', '')
                self.last_name = settings.get('last_name', '')
                self.system_prompt = settings.get('system_prompt', self.system_prompt)
        except FileNotFoundError:
            pass  # It's okay if the file doesn't exist yet

    def save_settings(self):
        settings = {
            'api_key': self.api_key_entry.get(),
            'first_name': self.first_name_entry.get(),
            'last_name': self.last_name_entry.get(),
            'system_prompt': self.system_prompt_text.get(1.0, tk.END)
        }
        with open('claude_app_settings.json', 'w') as f:
            json.dump(settings, f)
        messagebox.showinfo("Settings", "Settings saved successfully.")

    def save_api_key(self):
        self.api_key = self.api_key_entry.get()
        messagebox.showinfo("API Key", "API Key saved successfully.")
    
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

    def send_request(self):
        if not self.instructions:
            messagebox.showerror("Error", "Please upload instruction files first.")
            return
        if not self.scripts:
            messagebox.showerror("Error", "Please upload script files first.")
            return
        if not self.first_name_entry.get() or not self.last_name_entry.get():
            messagebox.showerror("Error", "Please enter your first and last name.")
            return
        if not self.api_key_entry.get():
            messagebox.showerror("Error", "Please enter your API key.")
            return
        
        self.first_name = self.first_name_entry.get()
        self.last_name = self.last_name_entry.get()
        self.api_key = self.api_key_entry.get()
        self.date = self.date_entry.get()
        self.system_prompt = self.system_prompt_text.get(1.0, tk.END)
        
        # Prepare the messages
        formatted_instructions = "\n\n".join([f"Instruction {i+1} ({name}):\n{content}" for i, (name, content) in enumerate(self.instructions)])
        formatted_scripts = "\n\n".join([f"Script {i+1} ({name}):\n{content}" for i, (name, content) in enumerate(self.scripts)])
        system_message = self.system_prompt.format(
            scripts=formatted_scripts,
            instructions=formatted_instructions,
            first_name=self.first_name,
            last_name=self.last_name,
            date=self.date
        )
        messages = [
            {"role": "user", "content": system_message}
        ]
        
        # API call parameters
        api_url = "https://api.anthropic.com/v1/complete"
        headers = {
            "x-api-key": self.api_key,
            "anthropic-version": "2023-06-01",
            "content-type": "application/json",
        }
        data = {
            "model": "claude-2",
            "max_tokens_to_sample": 8192,
            "prompt": "\n\n".join([f"{msg['role']}: {msg['content']}" for msg in messages]) + "\n\nAssistant:",
            "stop_sequences": ["\n\nHuman:"],
        }
        
        try:
            # Show a loading message
            self.output_text.delete(1.0, tk.END)
            self.output_text.insert(tk.END, "Generating response, please wait...")
            self.update_idletasks()
            
            response = requests.post(api_url, headers=headers, json=data)
            if response.status_code == 200:
                result = response.json()
                content = result['completion']
                # Extract text content
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
        save_path = filedialog.asksaveasfilename(title="Save Output as Word File", defaultextension=".docx",
                                                 filetypes=[("Word Document", "*.docx")])
        if save_path:
            try:
                document = Document()
                
                # Document settings
                section = document.sections[0]
                section.page_height = Inches(11)
                section.page_width = Inches(8.5)
                section.left_margin = Inches(1)
                section.right_margin = Inches(1)
                section.top_margin = Inches(1)
                section.bottom_margin = Inches(1)
                
                # Define styles
                styles = document.styles
                
                # Title style
                style_title = styles.add_style('TitleStyle', WD_STYLE_TYPE.PARAGRAPH)
                style_title.font.size = Pt(20)
                style_title.font.bold = True
                
                # Heading 1 style
                style_heading1 = styles.add_style('Heading1', WD_STYLE_TYPE.PARAGRAPH)
                style_heading1.base_style = styles['Heading 1']
                style_heading1.font.size = Pt(16)
                style_heading1.font.bold = True
                
                # Heading 2 style
                style_heading2 = styles.add_style('Heading2', WD_STYLE_TYPE.PARAGRAPH)
                style_heading2.base_style = styles['Heading 2']
                style_heading2.font.size = Pt(14)
                style_heading2.font.bold = True
                
                # Heading 3 style
                style_heading3 = styles.add_style('Heading3', WD_STYLE_TYPE.PARAGRAPH)
                style_heading3.base_style = styles['Heading 3']
                style_heading3.font.size = Pt(12)
                style_heading3.font.bold = True
                
                # Normal text style
                style_normal = styles['Normal']
                style_normal.font.size = Pt(12)
                
                # Citation style
                style_citation = styles.add_style('Citation', WD_STYLE_TYPE.PARAGRAPH)
                style_citation.font.size = Pt(12)
                style_citation.font.italic = True
                
                # Convert Markdown to HTML
                html = markdown.markdown(output)
                
                # Parse HTML
                soup = BeautifulSoup(html, 'html.parser')
                
                # Process the HTML elements
                for element in soup.descendants:
                    if element.name == 'h1':
                        # Title
                        p = document.add_paragraph(element.get_text(strip=True), style='TitleStyle')
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    elif element.name == 'h2':
                        # Heading level 1
                        heading = element.get_text(strip=True)
                        if heading.lower() == 'bibliography':
                            document.add_page_break()
                        document.add_paragraph(heading, style='Heading1')
                    elif element.name == 'h3':
                        # Heading level 2
                        document.add_paragraph(element.get_text(strip=True), style='Heading2')
                    elif element.name == 'h4':
                        # Heading level 3
                        document.add_paragraph(element.get_text(strip=True), style='Heading3')
                    elif element.name == 'p':
                        # Regular paragraph
                        paragraph = document.add_paragraph(style='Normal')
                        self._add_runs(paragraph, element)
                    elif element.name == 'em' or element.name == 'i':
                        pass  # Handled in _add_runs
                    elif element.name == 'strong' or element.name == 'b':
                        pass  # Handled in _add_runs
                    elif element.name == 'ul':
                        # Unordered list
                        for li in element.find_all('li', recursive=False):
                            paragraph = document.add_paragraph(style='List Bullet')
                            self._add_runs(paragraph, li)
                    elif element.name == 'ol':
                        # Ordered list
                        for li in element.find_all('li', recursive=False):
                            paragraph = document.add_paragraph(style='List Number')
                            self._add_runs(paragraph, li)
                    elif element.name == 'blockquote':
                        # Blockquote
                        paragraph = document.add_paragraph(style='Intense Quote')
                        self._add_runs(paragraph, element)
                
                # Add page numbers
                self._add_page_numbers(document.sections[0])
                
                document.save(save_path)
                messagebox.showinfo("Success", f"Output saved to {save_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Error saving Word file: {e}")

    def _add_page_numbers(self, section):
        footer = section.footer
        paragraph = footer.paragraphs[0]
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = paragraph.add_run()
        fldSimple = OxmlElement('w:fldSimple')
        fldSimple.set(qn('w:instr'), 'PAGE')
        run._r.append(fldSimple)

    def _add_runs(self, paragraph, element):
        for node in element.descendants:
            if isinstance(node, str):
                paragraph.add_run(node)
            elif node.name == 'strong' or node.name == 'b':
                run = paragraph.add_run(node.get_text())
                run.bold = True
            elif node.name == 'em' or node.name == 'i':
                run = paragraph.add_run(node.get_text())
                run.italic = True
            elif node.name == 'a':
                run = paragraph.add_run(node.get_text())
                run.font.underline = True
            elif node.name == 'img':
                img_src = node.get('src')
                if img_src:
                    try:
                        response = requests.get(img_src)
                        if response.status_code == 200:
                            from io import BytesIO
                            image_stream = BytesIO(response.content)
                            paragraph.add_run().add_picture(image_stream)
                    except Exception as e:
                        pass  # Image couldn't be loaded
            # Add more styles as needed

if __name__ == "__main__":
    app = ClaudeApp()
    app.mainloop()
