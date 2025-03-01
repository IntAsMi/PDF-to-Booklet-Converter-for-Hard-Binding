import streamlit as st
import os
import tempfile
import subprocess
import io
from PIL import Image
import base64
import time
import sys

def install_missing_packages():
    required_packages = ['PyPDF2', 'fitz', 'pikepdf', 'reportlab', 'PyMuPDF', 'Pillow']
    installed = []
    
    for package in required_packages:
        try:
            if package == 'fitz':
                __import__('fitz')
            elif package == 'PyMuPDF':
                __import__('fitz')  # PyMuPDF provides the fitz module
            elif package == 'Pillow':
                __import__('PIL')  # Pillow provides the PIL module
            else:
                __import__(package.lower().replace('-', '_'))
            st.sidebar.success(f"✓ {package} is already installed.")
        except ImportError:
            st.sidebar.info(f"Installing {package}...")
            try:
                if package == 'PyMuPDF':
                    subprocess.check_call([sys.executable, "-m", "pip", "install", package])
                else:
                    subprocess.check_call([sys.executable, "-m", "pip", "install", package])
                st.sidebar.success(f"✓ {package} installed successfully.")
                installed.append(package)
            except Exception as e:
                st.sidebar.error(f"Failed to install {package}: {str(e)}")
                
    if installed:
        st.sidebar.warning("Installed new packages. Please restart the app if you encounter any issues.")
    
    return True

# Try importing after installation
try:
    import PyPDF2
    import fitz  # PyMuPDF
    from pikepdf import Pdf
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import letter, A4, legal
except ImportError as e:
    st.error(f"Error importing required libraries: {str(e)}")
    st.warning("Please restart the app after package installation.")
    st.stop()

def calculate_booklet_page_count(page_count):
    """Calculate the number of pages needed for a booklet (multiple of 4)"""
    return (page_count + 3) // 4 * 4

def create_booklet_pdf(input_pdf_path, output_pdf_path, page_size='A4', 
                       binding_margin=10, bleed=0, crop_marks=False, 
                       signatures=None, include_blank=True, compression_level=2):
    """
    Convert a regular PDF into a booklet PDF suitable for printing and binding.
    
    Args:
        input_pdf_path: Path to the input PDF file
        output_pdf_path: Path where the booklet PDF will be saved
        page_size: The size of the output pages ('A4', 'Letter', 'Legal', etc.)
        binding_margin: Extra margin in mm for the binding edge
        bleed: Bleed area in mm
        crop_marks: Whether to include crop marks
        signatures: Number of pages per signature (must be multiple of 4), or None for a single signature
        include_blank: Whether to add blank pages to make the total a multiple of 4
        compression_level: PDF compression level (0-4, where 0 is no compression, 4 is maximum)
    """
    try:
        # Open the input PDF directly with PyMuPDF for efficiency
        input_doc = fitz.open(input_pdf_path)
        input_page_count = len(input_doc)
        
        # Determine total pages needed (must be a multiple of 4 for booklet printing)
        total_pages = calculate_booklet_page_count(input_page_count)
        
        # Define the page size
        if page_size.upper() == 'A4':
            page_width, page_height = 210, 297  # mm
        elif page_size.upper() == 'LETTER':
            page_width, page_height = 215.9, 279.4  # mm
        elif page_size.upper() == 'LEGAL':
            page_width, page_height = 215.9, 355.6  # mm
        else:
            page_width, page_height = 210, 297  # Default to A4
        
        # Create a PDF with PyMuPDF for better control and visual quality
        doc = fitz.open()
        
        # Calculate spread dimensions (two pages side by side)
        spread_width = (page_width * 2) + (binding_margin * 2) + (bleed * 2)
        spread_height = page_height + (bleed * 2)
        
        # Determine signature count
        if signatures is None:
            # Single signature (all pages together)
            signature_ranges = [(0, total_pages)]
        else:
            # Multiple signatures
            signature_size = (signatures + 3) // 4 * 4  # Ensure multiple of 4
            signature_count = (total_pages + signature_size - 1) // signature_size
            signature_ranges = []
            for i in range(signature_count):
                start = i * signature_size
                end = min(start + signature_size, total_pages)
                # Ensure each signature has multiple of 4 pages
                if (end - start) % 4 != 0:
                    end = start + ((end - start + 3) // 4 * 4)
                signature_ranges.append((start, end))
        
        # Process each signature
        for sig_start, sig_end in signature_ranges:
            sig_pages = sig_end - sig_start
            
            # Process pages in booklet order for this signature
            for i in range(sig_pages // 2):
                # Create a new spread (page in the output PDF)
                spread_page = doc.new_page(width=spread_width, height=spread_height)
                
                # Calculate booklet indices for this pair
                left_idx = sig_start + sig_pages - 1 - i
                right_idx = sig_start + i
                
                # Insert left page (verso) - handle pages beyond actual content
                if left_idx < input_page_count:
                    left_rect = fitz.Rect(
                        bleed, 
                        bleed, 
                        bleed + page_width,
                        bleed + page_height
                    )
                    # Using direct page reference instead of creating intermediate PDFs
                    spread_page.show_pdf_page(left_rect, input_doc, left_idx)
                
                # Insert right page (recto) - handle pages beyond actual content
                if right_idx < input_page_count:
                    right_rect = fitz.Rect(
                        bleed + page_width + binding_margin * 2, 
                        bleed, 
                        bleed + page_width * 2 + binding_margin * 2,
                        bleed + page_height
                    )
                    # Using direct page reference instead of creating intermediate PDFs
                    spread_page.show_pdf_page(right_rect, input_doc, right_idx)
                
                # Add crop marks if requested
                if crop_marks:
                    # Create crop mark lines
                    mark_length = 10  # mm
                    line_width = 0.2  # mm
                    
                    # Convert to PDF points (1/72 inch) from mm
                    ml = mark_length * 72 / 25.4
                    lw = line_width * 72 / 25.4
                    
                    # Draw crop marks
                    spread_page.draw_line((bleed-ml, bleed), (bleed, bleed), color=(0, 0, 0), width=lw)
                    spread_page.draw_line((bleed, bleed-ml), (bleed, bleed), color=(0, 0, 0), width=lw)
                    
                    spread_page.draw_line((bleed+page_width, bleed-ml), (bleed+page_width, bleed), color=(0, 0, 0), width=lw)
                    spread_page.draw_line((bleed+page_width+ml, bleed), (bleed+page_width, bleed), color=(0, 0, 0), width=lw)
                    
                    spread_page.draw_line((bleed-ml, bleed+page_height), (bleed, bleed+page_height), color=(0, 0, 0), width=lw)
                    spread_page.draw_line((bleed, bleed+page_height+ml), (bleed, bleed+page_height), color=(0, 0, 0), width=lw)
                    
                    spread_page.draw_line((bleed+page_width, bleed+page_height+ml), (bleed+page_width, bleed+page_height), color=(0, 0, 0), width=lw)
                    spread_page.draw_line((bleed+page_width+ml, bleed+page_height), (bleed+page_width, bleed+page_height), color=(0, 0, 0), width=lw)
                    
                    # Right page crop marks
                    right_x = bleed + page_width + binding_margin * 2
                    
                    spread_page.draw_line((right_x-ml, bleed), (right_x, bleed), color=(0, 0, 0), width=lw)
                    spread_page.draw_line((right_x, bleed-ml), (right_x, bleed), color=(0, 0, 0), width=lw)
                    
                    spread_page.draw_line((right_x+page_width, bleed-ml), (right_x+page_width, bleed), color=(0, 0, 0), width=lw)
                    spread_page.draw_line((right_x+page_width+ml, bleed), (right_x+page_width, bleed), color=(0, 0, 0), width=lw)
                    
                    spread_page.draw_line((right_x-ml, bleed+page_height), (right_x, bleed+page_height), color=(0, 0, 0), width=lw)
                    spread_page.draw_line((right_x, bleed+page_height+ml), (right_x, bleed+page_height), color=(0, 0, 0), width=lw)
                    
                    spread_page.draw_line((right_x+page_width, bleed+page_height+ml), (right_x+page_width, bleed+page_height), color=(0, 0, 0), width=lw)
                    spread_page.draw_line((right_x+page_width+ml, bleed+page_height), (right_x+page_width, bleed+page_height), color=(0, 0, 0), width=lw)
        
        # Save the output PDF with compression
        doc.save(output_pdf_path, 
                 garbage=4,  # Maximum garbage collection
                 clean=True,  # Clean redundant elements
                 deflate=True,  # Use deflate compression 
                 deflate_images=True,  # Compress images
                 deflate_fonts=True)  # Compress fonts
        
        # If pikepdf is available, use it to optimize the PDF further
        try:
            with Pdf.open(output_pdf_path) as pdf:
                # Set compression options
                pdf.save(output_pdf_path, 
                         object_stream_mode=2,  # Use object streams
                         compress_streams=True,  # Compress all streams
                         stream_decode_level=compression_level)  # Level of compression
        except Exception as e:
            st.warning(f"Additional optimization with pikepdf failed (but PDF was saved): {str(e)}")
            
        return True
        
    except Exception as e:
        st.error(f"Error creating booklet PDF: {str(e)}")
        raise
        return False

def get_preview_images(pdf_path, max_pages=4):
    """Generate preview images from the PDF"""
    preview_images = []
    try:
        pdf_document = fitz.open(pdf_path)
        page_count = min(len(pdf_document), max_pages)
        
        for page_num in range(page_count):
            page = pdf_document.load_page(page_num)
            pix = page.get_pixmap(matrix=fitz.Matrix(0.5, 0.5))
            img_data = pix.tobytes("png")
            preview_images.append(img_data)
            
        return preview_images
    except Exception as e:
        st.error(f"Error generating preview: {str(e)}")
        return []

def main():
    st.set_page_config(
        page_title="PDF Booklet Creator",
        page_icon="📚",
        layout="wide"
    )
    
    # Add custom CSS
    st.markdown("""
    <style>
    .main {
        padding: 2rem;
    }
    .stButton button {
        width: 100%;
        background-color: #4CAF50;
        color: white;
        font-weight: bold;
    }
    .preview-header {
        text-align: center;
        font-weight: bold;
        margin-bottom: 10px;
    }
    </style>
    """, unsafe_allow_html=True)
    
    st.title("📚 PDF Booklet Creator")
    st.subheader("Convert any PDF into a booklet format for hard binding")
    
    # Install missing packages if needed
    with st.sidebar:
        st.header("Package Installation")
        install_status = install_missing_packages()
    
    # File uploader
    uploaded_file = st.file_uploader("Upload your PDF file", type=["pdf"])
    
    if uploaded_file is not None:
        # Save the uploaded file to a temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            input_pdf_path = tmp_file.name
        
        # Get page count
        try:
            reader = PyPDF2.PdfReader(input_pdf_path)
            original_page_count = len(reader.pages)
            booklet_page_count = calculate_booklet_page_count(original_page_count)
            
            st.info(f"Original PDF has {original_page_count} pages. Booklet will have {booklet_page_count} pages (including any blank pages).")
            
            # Blank pages info
            if booklet_page_count > original_page_count:
                st.info(f"{booklet_page_count - original_page_count} blank pages will be added to make the total a multiple of 4 (required for booklet format).")
        except Exception as e:
            st.error(f"Error reading PDF: {str(e)}")
            st.stop()
            
        # Sidebar settings
        st.sidebar.header("Booklet Settings")
        
        page_size = st.sidebar.selectbox(
            "Page Size", 
            ["A4", "Letter", "Legal"],
            help="Size of the output paper"
        )
        
        binding_margin = st.sidebar.slider(
            "Binding Margin (mm)", 
            min_value=0, 
            max_value=30, 
            value=10,
            help="Extra margin for the binding edge"
        )
        
        bleed = st.sidebar.slider(
            "Bleed Area (mm)", 
            min_value=0, 
            max_value=10, 
            value=3,
            help="Extra area beyond the trim size that will be cut off"
        )
        
        crop_marks = st.sidebar.checkbox(
            "Add Crop Marks", 
            value=False,
            help="Add crop marks for professional printing"
        )
        
        # Calculate how many signatures (multiples of 4 pages) would be needed
        max_signature_size = min(80, booklet_page_count)  # Limit max signature size
        signature_options = [4, 8, 16, 24, 32, 40, 48, 56, 64, 72, 80]
        valid_signature_options = [s for s in signature_options if s <= max_signature_size]
        
        use_signatures = st.sidebar.checkbox(
            "Use Signatures", 
            value=False,
            help="Split booklet into separate signatures (groups of pages)"
        )
        
        signatures = None
        if use_signatures:
            signatures = st.sidebar.selectbox(
                "Pages per Signature", 
                valid_signature_options,
                index=min(2, len(valid_signature_options)-1),  # Default to 16 pages if available
                help="Number of pages in each signature (groups for binding)"
            )
        
        # Create a temporary file for the output PDF
        output_pdf_path = input_pdf_path.replace(".pdf", "_booklet.pdf")
        
        # Add compression option
        compression_level = st.sidebar.slider(
            "Compression Level", 
            min_value=0, 
            max_value=4, 
            value=2,
            help="Level of PDF compression (0=none, 4=maximum)"
        )
        
        optimize_images = st.sidebar.checkbox(
            "Optimize Images", 
            value=True,
            help="Reduce image quality to decrease file size"
        )
        
        if optimize_images:
            image_quality = st.sidebar.slider(
                "Image Quality", 
                min_value=25, 
                max_value=100, 
                value=80,
                help="Lower quality = smaller file (may affect image clarity)"
            )
        else:
            image_quality = 100
            
        # Process button
        if st.button("Create Booklet PDF", key="create_button"):
            with st.spinner("Creating booklet PDF..."):
                try:
                    # Get original file size
                    original_size = os.path.getsize(input_pdf_path) / (1024 * 1024)  # in MB
                    
                    success = create_booklet_pdf(
                        input_pdf_path=input_pdf_path,
                        output_pdf_path=output_pdf_path,
                        page_size=page_size,
                        binding_margin=binding_margin,
                        bleed=bleed,
                        crop_marks=crop_marks,
                        signatures=signatures,
                        compression_level=compression_level
                    )
                    
                    if success:
                        # Get file sizes
                        output_size = os.path.getsize(output_pdf_path) / (1024 * 1024)  # in MB
                        
                        # Advanced compression if file grew too much
                        if output_size > original_size * 2 and image_quality < 100:
                            st.info("Applying additional compression to reduce file size...")
                            
                            # Try additional optimization with pymupdf for images
                            try:
                                # Open the booklet PDF
                                doc = fitz.open(output_pdf_path)
                                
                                # Iterate through pages and compress images
                                for page_num in range(len(doc)):
                                    page = doc[page_num]
                                    image_list = page.get_images(full=True)
                                    
                                    for img_index, img in enumerate(image_list):
                                        xref = img[0]
                                        
                                        # Try to get image
                                        try:
                                            base_image = doc.extract_image(xref)
                                            image_bytes = base_image["image"]
                                            
                                            # Convert to PIL image
                                            import io
                                            from PIL import Image
                                            pil_image = Image.open(io.BytesIO(image_bytes))
                                            
                                            # Compress and replace
                                            temp_buffer = io.BytesIO()
                                            pil_image.save(temp_buffer, format="JPEG", 
                                                          quality=image_quality, optimize=True)
                                            temp_buffer.seek(0)
                                            
                                            # Replace image in PDF
                                            page.delete_image(img[1])  # delete old image
                                            page.insert_image(
                                                img[1],  # original rectangle
                                                stream=temp_buffer.getvalue(),
                                            )
                                        except Exception as img_err:
                                            # Skip this image if error
                                            continue
                                
                                # Save optimized PDF
                                doc.save(output_pdf_path, 
                                         garbage=4, 
                                         deflate=True, 
                                         deflate_images=True,
                                         deflate_fonts=True)
                                doc.close()
                                
                                # Update output size
                                output_size = os.path.getsize(output_pdf_path) / (1024 * 1024)
                            except Exception as e:
                                st.warning(f"Additional image compression failed: {str(e)}")
                        
                        # Try one more advanced compression with pikepdf for extremely large files
                        if output_size > original_size * 3:
                            try:
                                with Pdf.open(output_pdf_path) as pdf:
                                    # Remove unnecessary metadata and structure
                                    if "/Metadata" in pdf.Root:
                                        del pdf.Root.Metadata
                                    
                                    # Linearize and save with extreme compression
                                    pdf.save(output_pdf_path, 
                                            linearize=True,  # Optimize for web viewing
                                            object_stream_mode=2,  # Use object streams 
                                            compress_streams=True,  # Compress all streams
                                            stream_decode_level=3)  # High compression
                                    
                                # Update output size
                                output_size = os.path.getsize(output_pdf_path) / (1024 * 1024)
                            except Exception as e:
                                st.warning(f"Final optimization failed: {str(e)}")
                                
                        # Display size info
                        size_ratio = output_size / original_size
                        size_status = "✅ Good" if size_ratio < 3 else "⚠️ Large"
                        
                        st.success(f"Booklet PDF created successfully! {size_status}")
                        st.info(f"Original size: {original_size:.2f} MB → Booklet size: {output_size:.2f} MB ({size_ratio:.1f}x)")
                        
                        # Show optimization tips if file is still large
                        if size_ratio > 4:
                            st.warning("""
                            **Tips to reduce file size further:**
                            - Lower the image quality slider
                            - Enable image optimization
                            - Use an external tool like Adobe Acrobat or online PDF compressor
                            """)
                        
                        # Preview section
                        st.header("Preview")
                        preview_images = get_preview_images(output_pdf_path)
                        
                        if preview_images:
                            # Display preview images in a grid
                            cols = st.columns(min(len(preview_images), 2))
                            for i, img_data in enumerate(preview_images):
                                col_idx = i % 2
                                with cols[col_idx]:
                                    st.image(img_data, caption=f"Page {i+1}", use_container_width =True)
                        
                        # Download button
                        with open(output_pdf_path, "rb") as file:
                            output_file_data = file.read()
                            
                        st.download_button(
                            label="📥 Download Booklet PDF",
                            data=output_file_data,
                            file_name=f"booklet_{os.path.basename(uploaded_file.name)}",
                            mime="application/pdf",
                            key="download_button"
                        )
                    else:
                        st.error("Failed to create booklet PDF.")
                except Exception as e:
                    st.error(f"Error creating booklet PDF: {str(e)}")
                    
        # Cleanup temporary files
        try:
            if 'input_pdf_path' in locals():
                # Do not delete here, will be cleaned up when Python exits
                pass
        except:
            pass
    else:
        # Show sample layout
        st.markdown("""
        ### How it works:
        
        1. **Upload your PDF** - Any PDF can be converted to a booklet format
        2. **Customize settings** - Adjust margins, page size, and other options
        3. **Create and download** - Generate your booklet PDF ready for printing
        
        ### Booklet Format:
        
        A booklet arranges pages so that when printed double-sided and folded in half, 
        the pages appear in the correct reading order.
        
        For example, in a 4-page booklet:
        - Sheet 1 side 1: Page 4 | Page 1
        - Sheet 1 side 2: Page 2 | Page 3
        
        This app handles all page arrangements automatically!
        """)
        
        # Show example diagram
        cols = st.columns([1, 2, 1])
        with cols[1]:
            st.image("https://www.printninja.com/wp-content/uploads/2021/06/Book-Binding-Diagram.png", caption="Booklet Binding Example", use_container_width =True)

if __name__ == "__main__":
    main()