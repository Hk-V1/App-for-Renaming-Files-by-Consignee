import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import zipfile
import os
import shutil
import tempfile
import re
from pathlib import Path
from collections import defaultdict
import pdfplumber


class PDFRenamerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Renaming Tool")
        self.root.geometry("900x600")
        self.temp_dir = None
        self.pdf_files = []
        self.renamed_files = {}
        self.setup_ui()
        
    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=2)
        main_frame.columnconfigure(1, weight=3)
        main_frame.rowconfigure(0, weight=1)     
        left_frame = ttk.LabelFrame(main_frame, text="PDF Files", padding="10")
        left_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 5))       
        list_scroll = ttk.Scrollbar(left_frame, orient=tk.VERTICAL)
        self.file_listbox = tk.Listbox(left_frame, selectmode=tk.MULTIPLE, 
                                       yscrollcommand=list_scroll.set,
                                       font=("Consolas", 9))
        list_scroll.config(command=self.file_listbox.yview)
        
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        list_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        right_frame = ttk.Frame(main_frame, padding="10")
        right_frame.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S))
        right_frame.rowconfigure(1, weight=1)
        right_frame.columnconfigure(0, weight=1)
        
        btn_frame = ttk.Frame(right_frame)
        btn_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.upload_btn = ttk.Button(btn_frame, text="Upload ZIP", 
                                     command=self.upload_zip, width=20)
        self.upload_btn.pack(pady=5, fill=tk.X)
        
        self.rename_btn = ttk.Button(btn_frame, text="Rename Selected", 
                                     command=self.rename_selected, 
                                     state=tk.DISABLED, width=20)
        self.rename_btn.pack(pady=5, fill=tk.X)
        
        self.download_btn = ttk.Button(btn_frame, text="Download ZIP", 
                                       command=self.download_zip, 
                                       state=tk.DISABLED, width=20)
        self.download_btn.pack(pady=5, fill=tk.X)
        
        status_frame = ttk.LabelFrame(right_frame, text="Status Log", padding="10")
        status_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        status_frame.rowconfigure(0, weight=1)
        status_frame.columnconfigure(0, weight=1)
        
        log_scroll = ttk.Scrollbar(status_frame, orient=tk.VERTICAL)
        self.log_text = tk.Text(status_frame, wrap=tk.WORD, 
                                yscrollcommand=log_scroll.set,
                                font=("Consolas", 9), height=20)
        log_scroll.config(command=self.log_text.yview)
        
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        log_scroll.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        self.log("Ready. Please upload a ZIP file containing PDFs.")
        
    def log(self, message):
        """Add message to log window"""
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
        
    def upload_zip(self):
        """Handle ZIP file upload and extraction"""
        zip_path = filedialog.askopenfilename(
            title="Select ZIP file",
            filetypes=[("ZIP files", "*.zip"), ("All files", "*.*")]
        )
        
        if not zip_path:
            return
            
        if self.temp_dir and os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir)
            
        self.temp_dir = tempfile.mkdtemp()
        
        try:
            self.log(f"Extracting ZIP file: {Path(zip_path).name}")
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(self.temp_dir)
                
            self.pdf_files = []
            for root, dirs, files in os.walk(self.temp_dir):
                for file in files:
                    if file.lower().endswith('.pdf'):
                        self.pdf_files.append(os.path.join(root, file))
                        
            if not self.pdf_files:
                messagebox.showwarning("No PDFs", "No PDF files found in the ZIP archive.")
                self.log("ERROR: No PDF files found in ZIP.")
                return
                
            self.file_listbox.delete(0, tk.END)
            for pdf in self.pdf_files:
                self.file_listbox.insert(tk.END, Path(pdf).name)
                
            self.log(f"Found {len(self.pdf_files)} PDF file(s).")
            self.rename_btn.config(state=tk.NORMAL)
            self.renamed_files = {}
            self.download_btn.config(state=tk.DISABLED)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to extract ZIP: {str(e)}")
            self.log(f"ERROR: {str(e)}")
            
    def extract_consignee_name(self, pdf_path):
        """Extract and clean consignee name from PDF"""
        try:
            with pdfplumber.open(pdf_path) as pdf:
                text = ""
                for page in pdf.pages:
                    text += page.extract_text() or ""
                    
            lines = text.split('\n')
            consignee_text = None
            for i, line in enumerate(lines):
                if re.search(r"Consignee\s*\(Ship to\)", line, re.IGNORECASE):
                    if i + 1 < len(lines):
                        consignee_text = lines[i + 1].strip()
                        break
            
            if not consignee_text:
                return None
            
            clean_pattern = r"(Buyer'?s?\s*Order\s*No\.?|Dated)"
            consignee_text = re.split(clean_pattern, consignee_text, flags=re.IGNORECASE)[0].strip()           
            consignee_text = re.sub(r'[^a-zA-Z0-9\s_-]', '', consignee_text)           
            consignee_text = consignee_text.replace(' ', '_')           
            consignee_text = re.sub(r'_+', '_', consignee_text)           
            consignee_text = consignee_text.strip('_')
            
            return consignee_text if consignee_text else None
            
        except Exception as e:
            self.log(f"ERROR reading {Path(pdf_path).name}: {str(e)}")
            return None
            
    def rename_selected(self):
        """Rename selected PDF files based on consignee name"""
        selected_indices = self.file_listbox.curselection()
        
        if not selected_indices:
            messagebox.showwarning("No Selection", "Please select at least one PDF file to rename.")
            return
            
        self.log("\n" + "="*50)
        self.log("Starting renaming process...")
        self.log("="*50)
        
        name_counts = defaultdict(int)
        self.renamed_files = {}
        
        for idx in selected_indices:
            pdf_path = self.pdf_files[idx]
            original_name = Path(pdf_path).name
            
            self.log(f"\nProcessing: {original_name}")
            
            consignee_name = self.extract_consignee_name(pdf_path)
            
            if not consignee_name:
                self.log(f"  ⚠ WARNING: Could not find consignee name. Skipping.")
                continue
                
            name_counts[consignee_name] += 1
            count = name_counts[consignee_name]
            
            if count > 1:
                new_name = f"{consignee_name}_{count}.pdf"
            else:
                new_name = f"{consignee_name}.pdf"
                
            new_path = os.path.join(os.path.dirname(pdf_path), new_name)
            
            try:
                os.rename(pdf_path, new_path)
                self.renamed_files[pdf_path] = new_path
                self.log(f"  ✓ Renamed to: {new_name}")
                
                self.pdf_files[idx] = new_path
                
            except Exception as e:
                self.log(f"  ✗ ERROR renaming file: {str(e)}")
                
        if self.renamed_files:
            self.log(f"\n{'='*50}")
            self.log(f"Successfully renamed {len(self.renamed_files)} file(s).")
            self.log(f"{'='*50}")
            self.download_btn.config(state=tk.NORMAL)
            
            self.file_listbox.delete(0, tk.END)
            for pdf in self.pdf_files:
                self.file_listbox.insert(tk.END, Path(pdf).name)
        else:
            self.log("\nNo files were renamed.")
            
    def download_zip(self):
        """Create and download ZIP file with renamed PDFs"""
        save_path = filedialog.asksaveasfilename(
            title="Save renamed files as",
            defaultextension=".zip",
            filetypes=[("ZIP files", "*.zip"), ("All files", "*.*")],
            initialfile="renamed_files.zip"
        )
        
        if not save_path:
            return
            
        try:
            self.log(f"\nCreating ZIP archive...")
            
            with zipfile.ZipFile(save_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for pdf_path in self.pdf_files:
                    if os.path.exists(pdf_path):
                        arcname = Path(pdf_path).name
                        zipf.write(pdf_path, arcname)
                        
            self.log(f"✓ ZIP file saved: {Path(save_path).name}")
            self.log(f"  Location: {save_path}")
            messagebox.showinfo("Success", f"ZIP file created successfully!\n{save_path}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create ZIP: {str(e)}")
            self.log(f"ERROR: {str(e)}")
            
    def on_closing(self):
        """Clean up temp directory on exit"""
        if self.temp_dir and os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir)
        self.root.destroy()


def main():
    root = tk.Tk()
    app = PDFRenamerApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()


if __name__ == "__main__":
    main()