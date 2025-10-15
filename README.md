# 📦 Consignee-Based PDF Renaming Tool

### 🧠 Summary
A Python-based Windows desktop app that uploads a ZIP of PDFs, extracts each file’s **“Consignee (Ship to)”** name, renames the files accordingly (removing unwanted text like *“Buyer's Order No.”* or *“Dated”*), and exports them as a new ZIP.

---

### ⚙️ Tech Stack
- **Language:** Python 3  
- **GUI:** Tkinter  
- **Libraries:** pdfplumber, zipfile, tempfile, shutil, re, os, uuid, pathlib

---

### 🌟 Features
- Upload ZIP file containing PDFs  
- Extract all PDF files and list them in a left panel  
- Auto-detect “Consignee (Ship to)” or “Ship to” fields  
- Clean extracted names (remove “Buyer's Order No.”, “Dated”)  
- Rename files and handle duplicates with serial numbers  
- Export renamed PDFs as a new ZIP  

---

### 🔄 Workflow
1. Upload ZIP → Extract PDFs  
2. Select files → Click *Rename*  
3. App reads consignee name and renames files  
4. Download the final renamed ZIP  

---

### 💡 Use Case
Helps logistics and documentation teams quickly rename shipment or invoice PDFs based on consignee names for organized file management.
