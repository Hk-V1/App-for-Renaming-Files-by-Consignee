import streamlit as st
import pandas as pd
import re
import zipfile
import io
import tempfile
import os
from pathlib import Path
import shutil

# PDF processing libraries
try:
    import pdfplumber
    PDF_LIB = 'pdfplumber'
except ImportError:
    try:
        from PyPDF2 import PdfReader
        PDF_LIB = 'pypdf2'
    except ImportError:
        PDF_LIB = None

# Excel processing
try:
    import openpyxl
    EXCEL_SUPPORT = True
except ImportError:
    EXCEL_SUPPORT = False


def extract_consignee_from_text(text):
    """
    Extract consignee name from text using regex patterns.
    Only extracts the line immediately below "Consignee (Ship to)".
    Returns the extracted name or None if not found.
    """
    if not text:
        return None
    
    # Primary pattern: Only match "Consignee (Ship to)" specifically
    # This will capture the content on the next line after the label
    pattern = r"Consignee\s*\(Ship\s*to\)\s*[:\-]?\s*\n?\s*([^\n\r]+)"
    
    match = re.search(pattern, text, re.IGNORECASE)
    if match:
        consignee = match.group(1).strip()
        # Clean up the consignee name
        consignee = re.sub(r'\s+', ' ', consignee)  # Remove extra spaces
        if len(consignee) > 0 and len(consignee) < 100:  # Reasonable length
            return consignee
    
    return None


def extract_from_pdf(file_path):
    """Extract text from PDF and find consignee."""
    try:
        if PDF_LIB == 'pdfplumber':
            with pdfplumber.open(file_path) as pdf:
                text = ""
                for page in pdf.pages[:3]:  # Check first 3 pages
                    text += page.extract_text() or ""
                return extract_consignee_from_text(text)
        
        elif PDF_LIB == 'pypdf2':
            with open(file_path, 'rb') as f:
                pdf_reader = PdfReader(f)
                text = ""
                for page in pdf_reader.pages[:3]:  # Check first 3 pages
                    text += page.extract_text() or ""
            return extract_consignee_from_text(text)
        
        else:
            return None
    except Exception as e:
        return None


def extract_from_excel(file_path):
    """Extract consignee from Excel/CSV file."""
    try:
        # Try reading as Excel first
        try:
            df = pd.read_excel(file_path)
        except:
            # Try as CSV
            df = pd.read_csv(file_path)
        
        # Search in all cells for consignee information
        text = df.to_string()
        consignee = extract_consignee_from_text(text)
        
        if not consignee:
            # Try searching column names and values
            for col in df.columns:
                if 'consignee' in str(col).lower() or 'ship to' in str(col).lower():
                    if not df[col].empty:
                        value = str(df[col].iloc[0])
                        if value and value != 'nan':
                            return value.strip()
        
        return consignee
    except Exception as e:
        return None


def extract_from_text(file_path):
    """Extract consignee from text file."""
    try:
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
            text = f.read()
        return extract_consignee_from_text(text)
    except Exception as e:
        return None


def sanitize_filename(name):
    """Convert consignee name to valid filename."""
    # Remove invalid filename characters
    name = re.sub(r'[<>:"/\\|?*]', '', name)
    # Replace spaces with underscores
    name = name.replace(' ', '_')
    # Remove leading/trailing dots and spaces
    name = name.strip('. ')
    # Limit length
    if len(name) > 50:
        name = name[:50]
    return name


def process_file_from_zip(file_path, filename):
    """
    Process a single file and extract consignee.
    Returns: (original_name, new_name, consignee)
    """
    file_ext = Path(filename).suffix.lower()
    
    consignee = None
    
    # Determine file type and extract
    if file_ext == '.pdf':
        consignee = extract_from_pdf(file_path)
    elif file_ext in ['.xlsx', '.xls', '.csv']:
        consignee = extract_from_excel(file_path)
    elif file_ext == '.txt':
        consignee = extract_from_text(file_path)
    
    # Generate new filename
    if consignee:
        new_name = sanitize_filename(consignee) + file_ext
    else:
        new_name = filename
    
    return filename, new_name, consignee


def extract_and_process_zip(uploaded_zip):
    """
    Extract ZIP, process files, rename them, and return results + new ZIP.
    Returns: (results_list, zip_buffer)
    """
    results = []
    
    # Create temporary directories
    with tempfile.TemporaryDirectory() as temp_extract_dir:
        with tempfile.TemporaryDirectory() as temp_output_dir:
            
            # Extract uploaded ZIP
            with zipfile.ZipFile(uploaded_zip, 'r') as zip_ref:
                zip_ref.extractall(temp_extract_dir)
            
            # Get all files from extracted directory
            all_files = []
            for root, dirs, files in os.walk(temp_extract_dir):
                for file in files:
                    # Skip hidden files and __MACOSX
                    if not file.startswith('.') and '__MACOSX' not in root:
                        all_files.append(os.path.join(root, file))
            
            if not all_files:
                return [], None
            
            # Process each file
            for file_path in all_files:
                filename = os.path.basename(file_path)
                
                # Process and get new name
                original_name, new_name, consignee = process_file_from_zip(file_path, filename)
                
                # Copy file with new name to output directory
                output_path = os.path.join(temp_output_dir, new_name)
                
                # Handle duplicate filenames
                counter = 1
                base_name = Path(new_name).stem
                extension = Path(new_name).suffix
                while os.path.exists(output_path):
                    new_name = f"{base_name}_{counter}{extension}"
                    output_path = os.path.join(temp_output_dir, new_name)
                    counter += 1
                
                shutil.copy2(file_path, output_path)
                
                # Store results
                results.append({
                    'Original File Name': original_name,
                    'Extracted Consignee': consignee if consignee else "Not found",
                    'New File Name': new_name
                })
            
            # Create ZIP from output directory
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_out:
                for root, dirs, files in os.walk(temp_output_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        arcname = os.path.basename(file_path)
                        zip_out.write(file_path, arcname)
            
            zip_buffer.seek(0)
            return results, zip_buffer


# Streamlit App
def main():
    st.set_page_config(
        page_title="Consignee File Renamer (ZIP)", 
        page_icon="📦", 
        layout="wide"
    )
    
    st.title("📦 Consignee-Based File Renamer (ZIP Upload)")
    st.markdown("""
    Upload a **ZIP file** containing multiple documents (PDFs, Excel, CSV, or text files).  
    The app will:
    1. Extract all files from the ZIP
    2. Detect and extract the **Consignee (Ship to)** information
    3. Rename files based on the consignee name
    4. Provide a new ZIP file with renamed files for download
    """)
    
    # Check for required libraries
    warnings = []
    if PDF_LIB is None:
        warnings.append("⚠️ PDF processing unavailable. Install pdfplumber: `pip install pdfplumber`")
    
    if not EXCEL_SUPPORT:
        warnings.append("⚠️ Excel support limited. Install openpyxl: `pip install openpyxl`")
    
    if warnings:
        for warning in warnings:
            st.warning(warning)
    
    st.markdown("---")
    
    # File upload
    st.header("1️⃣ Upload ZIP File")
    uploaded_zip = st.file_uploader(
        "Choose a ZIP file containing your documents",
        type=['zip'],
        help="Upload a ZIP archive with PDF, Excel, CSV, or text files"
    )
    
    if uploaded_zip:
        st.success(f"✅ ZIP file uploaded: **{uploaded_zip.name}**")
        
        # Process button
        col1, col2, col3 = st.columns([1, 1, 2])
        with col1:
            process_button = st.button("🚀 Process Files", type="primary", use_container_width=True)
        
        if process_button:
            st.header("2️⃣ Processing Files")
            
            with st.spinner("Extracting and processing files..."):
                # Process the ZIP file
                results, zip_buffer = extract_and_process_zip(uploaded_zip)
            
            if not results:
                st.error("❌ No valid files found in the ZIP archive.")
                return
            
            st.success("✅ Processing complete!")
            
            # Display results table
            st.header("3️⃣ Processing Results")
            results_df = pd.DataFrame(results)
            st.dataframe(results_df, use_container_width=True, height=400)
            
            # Statistics
            st.subheader("📊 Statistics")
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Files", len(results))
            with col2:
                success_count = sum(1 for r in results if r['Extracted Consignee'] != "Not found")
                st.metric("Consignee Found", success_count)
            with col3:
                failed_count = len(results) - success_count
                st.metric("Not Found", failed_count)
            
            # Download section
            st.header("4️⃣ Download Renamed Files")
            
            st.download_button(
                label="⬇️ Download ZIP with Renamed Files",
                data=zip_buffer,
                file_name="renamed_files.zip",
                mime="application/zip",
                type="primary",
                use_container_width=True
            )
            
            st.success("✅ Your renamed files are ready to download!")
            
            # Additional info
            with st.expander("ℹ️ View Processing Details"):
                st.write("**Extraction Patterns Used:**")
                st.code("""
- "Consignee (Ship to):"
- "Consignee:"
- "Ship to:"
- Case-insensitive matching
                """)
                st.write("**File Name Sanitization:**")
                st.write("- Special characters removed")
                st.write("- Spaces replaced with underscores")
                st.write("- Length limited to 50 characters")
    
    else:
        st.info("👆 Please upload a ZIP file to begin")
    
    # Footer
    st.markdown("---")
    st.markdown("""
    <div style='text-align: center; color: #666;'>
        <p><strong>Supported File Types:</strong> PDF, Excel (.xlsx, .xls), CSV, Text (.txt)</p>
        <p><strong>Note:</strong> Files without detectable consignee information will keep their original names</p>
    </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
