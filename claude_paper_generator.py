import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import requests
import os
import PyPDF2
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.section import WD_SECTION
from docx.shared import Pt, Inches
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
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
            "Your task is to write a professional university-level paper based on the provided scientific papers and study scripts. "
            "The paper should be thoroughly researched, well-structured, and ready for submission. Please adhere to the following guidelines:\n\n"
            "- **Title and Author**: Determine an appropriate and compelling title for the paper. The paper should be written by {first_name} {last_name}. Include the author's name and the date {date} in the header of each page.\n\n"
            "- **Formatting**: The paper should follow standard academic formatting:\n"
            "  - Use Times New Roman font, size 12 for the main text.\n"
            "  - Set line spacing to 1.5.\n"
            "  - Include page numbers at the bottom center of each page.\n"
            "  - Margins should be 1 inch on all sides.\n"
            "  - Include a header with the author's name and date.\n\n"
            "- **Structure**:\n"
            "  - **Title Page**: Include the title of the paper, the author's name, the course name, instructor's name, and the date, centered on the page.\n"
            "  - **Abstract**: Provide a concise summary of the paper's content.\n"
            "  - **Table of Contents**: List all the sections and subsections with page numbers.\n"
            "  - **Introduction**: Introduce the topic and outline the purpose of the paper.\n"
            "  - **Main Body**: Organize the content into chapters and sections with appropriate headings.\n"
            "  - **Conclusion**: Summarize the findings and discuss their implications.\n"
            "  - **Bibliography**: Include all references in Harvard citation style.\n\n"
            "- **Content**:\n"
            "  - Incorporate content from the provided scripts and cite them appropriately within the text.\n"
            "  - Ensure all information is accurate and derived from credible sources.\n"
            "  - Write in the first person singular, as if by {first_name} {last_name}.\n"
            "  - Occasionally include minor spelling mistakes to emulate writing by a student at B2/C1 English level.\n\n"
            "- **Style**:\n"
            "  - Use formal academic language throughout.\n"
            "  - Employ proper grammar and syntax.\n"
            "  - Use headings, subheadings, bullet points, numbered lists, tables, and figures where appropriate.\n"
            "  - Ensure that all citations are properly formatted in Harvard style within the text and in the bibliography.\n\n"
            "- **Markdown Output**:\n"
            "  - Output the paper in detailed Markdown format.\n"
            "  - Use appropriate Markdown syntax for headings, bold, italics, lists, tables, and images.\n"
            "  - Ensure that the Markdown is structured such that, when converted to a Word document, the paper appears professional and adheres to the formatting guidelines.\n\n"
            "Please incorporate the specific instructions provided below:\n\n"
            "{instructions}\n\n"
            "Use the scripts provided below as sources:\n\n"
            "{scripts}\n\n"
            "Ensure that the final output is a comprehensive, well-structured, and professionally formatted university paper that meets all the above guidelines."
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
            pass

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

        # Build the messages
        messages = [{"role": "user", "content": system_message}]

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
                content = result['content']
                # Extract text content
                response_text = ""
                for block in content:
                    if block['type'] == 'text':
                        response_text += block['text']
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
                
                # Set default font and size
                style = document.styles['Normal']
                font = style.font
                font.name = 'Times New Roman'
                font.size = Pt(12)
                style.paragraph_format.line_spacing = 1.5

                # Document settings
                section = document.sections[0]
                section.page_height = Inches(11)
                section.page_width = Inches(8.5)
                section.left_margin = Inches(1)
                section.right_margin = Inches(1)
                section.top_margin = Inches(1)
                section.bottom_margin = Inches(1)
                
                # Add headers and footers
                self._add_header_and_footer(document.sections[0])
                
                # Define styles
                styles = document.styles
                
                style_title = styles.add_style('TitleStyle', WD_STYLE_TYPE.PARAGRAPH)
                style_title.font.name = 'Times New Roman'
                style_title.font.size = Pt(20)
                style_title.font.bold = True
                style_title.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
                style_title.paragraph_format.line_spacing = 1.5
                
                style_heading1 = styles.add_style('Heading1', WD_STYLE_TYPE.PARAGRAPH)
                style_heading1.base_style = styles['Heading 1']
                style_heading1.font.name = 'Times New Roman'
                style_heading1.font.size = Pt(16)
                style_heading1.font.bold = True
                style_heading1.paragraph_format.line_spacing = 1.5
                
                style_heading2 = styles.add_style('Heading2', WD_STYLE_TYPE.PARAGRAPH)
                style_heading2.base_style = styles['Heading 2']
                style_heading2.font.name = 'Times New Roman'
                style_heading2.font.size = Pt(14)
                style_heading2.font.bold = True
                style_heading2.paragraph_format.line_spacing = 1.5
                
                style_heading3 = styles.add_style('Heading3', WD_STYLE_TYPE.PARAGRAPH)
                style_heading3.base_style = styles['Heading 3']
                style_heading3.font.name = 'Times New Roman'
                style_heading3.font.size = Pt(12)
                style_heading3.font.bold = True
                style_heading3.paragraph_format.line_spacing = 1.5
                
                style_normal = styles['Normal']
                style_normal.font.name = 'Times New Roman'
                style_normal.font.size = Pt(12)
                style_normal.paragraph_format.line_spacing = 1.5
                
                style_citation = styles.add_style('Citation', WD_STYLE_TYPE.PARAGRAPH)
                style_citation.font.name = 'Times New Roman'
                style_citation.font.size = Pt(12)
                style_citation.font.italic = True
                style_citation.paragraph_format.line_spacing = 1.5
                
                # Convert Markdown to HTML
                html = markdown.markdown(output)
                
                # Parse HTML
                soup = BeautifulSoup(html, 'html.parser')
                
                # Process the HTML elements
                for element in soup.descendants:
                    if element.name == 'h1':
                        text = element.get_text(strip=True)
                        if 'title page' in text.lower():
                            continue
                        p = document.add_paragraph(text, style='TitleStyle')
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    elif element.name == 'h2':
                        heading = element.get_text(strip=True)
                        if heading.lower() == 'bibliography':
                            document.add_page_break()
                        document.add_paragraph(heading, style='Heading1')
                    elif element.name == 'h3':
                        document.add_paragraph(element.get_text(strip=True), style='Heading2')
                    elif element.name == 'h4':
                        document.add_paragraph(element.get_text(strip=True), style='Heading3')
                    elif element.name == 'p':
                        paragraph = document.add_paragraph(style='Normal')
                        self._add_runs(paragraph, element)
                    elif element.name == 'em' or element.name == 'i':
                        pass
                    elif element.name == 'strong' or element.name == 'b':
                        pass
                    elif element.name == 'ul':
                        for li in element.find_all('li', recursive=False):
                            paragraph = document.add_paragraph(style='List Bullet')
                            self._add_runs(paragraph, li)
                    elif element.name == 'ol':
                        for li in element.find_all('li', recursive=False):
                            paragraph = document.add_paragraph(style='List Number')
                            self._add_runs(paragraph, li)
                    elif element.name == 'blockquote':
                        paragraph = document.add_paragraph(style='Intense Quote')
                        self._add_runs(paragraph, element)
                
                # Save the document
                document.save(save_path)
                messagebox.showinfo("Success", f"Output saved to {save_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Error saving Word file: {e}")

    def _add_header_and_footer(self, section):
        # Header
        header = section.header
        header_paragraph = header.paragraphs[0]
        header_paragraph.text = f"{self.first_name} {self.last_name} - {self.date}"
        header_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        header_paragraph.style.font.name = 'Times New Roman'
        header_paragraph.style.font.size = Pt(12)

        # Footer
        footer = section.footer
        footer_paragraph = footer.paragraphs[0]
        footer_paragraph.style.font.name = 'Times New Roman'
        footer_paragraph.style.font.size = Pt(12)
        footer_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = footer_paragraph.add_run()
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
                        pass

if __name__ == "__main__":
    app = ClaudeApp()
    app.mainloop()
