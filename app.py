"""
Enhanced DOCX Document Comparison Tool
A Flask-based application for comparing Microsoft Word documents with advanced features:
- Text content comparison with difflib
- Strikethrough detection and tracking
- Table structure and content analysis
- Image extraction and comparison
- Excel export with comprehensive reporting
- Encrypted document support
"""

import os
import io
import tempfile
import difflib
import re
from hashlib import md5
from collections import defaultdict
from typing import List, Dict, Any, Tuple, Optional
from datetime import datetime

# Third-party imports

from docx.shared import RGBColor
import msoffcrypto
import docx
import openpyxl
from flask import Flask, render_template, request, send_file, flash, jsonify, redirect, url_for
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
from PIL import Image as PILImage
import xml.etree.ElementTree as ET
from functools import lru_cache

# ================================
# APPLICATION CONFIGURATION
# ================================

app = Flask(__name__)
app.secret_key = "supersecretkey"  # Change this in production

# File upload configuration
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# ================================
# Performance Optimizations
# ================================

@lru_cache(maxsize=100)
def detect_strikethrough_cached(run_element):
    """Cached version of strikethrough detection"""
    return detect_strikethrough(run_element)

# ================================
# UTILITY FUNCTIONS
# ================================

def qn(tag: str) -> str:
    """
    Utility function to handle XML namespaces in DOCX documents.
    
    Args:
        tag (str): XML tag name
        
    Returns:
        str: Fully qualified namespace tag
    """
    if tag.startswith('{'):
        return tag
    return '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}' + tag


def is_encrypted(file_stream) -> bool:
    """
    Check if a document is encrypted without fully loading it.
    
    Args:
        file_stream: File stream object
        
    Returns:
        bool: True if document is encrypted, False otherwise
    """
    try:
        file_stream.seek(0)
        office_file = msoffcrypto.OfficeFile(file_stream)
        return office_file.is_encrypted()
    except Exception:
        return False
    finally:
        file_stream.seek(0)


def load_docx(file_stream, password: Optional[str] = None) -> docx.Document:
    """
    Load DOCX file, handling both encrypted and unencrypted documents.
    
    Args:
        file_stream: File stream object
        password (str, optional): Password for encrypted documents
        
    Returns:
        docx.Document: Loaded document object
        
    Raises:
        Exception: If document loading fails
    """
    file_stream.seek(0)
    
    if password and is_encrypted(file_stream):
        # Handle encrypted document
        decrypted = io.BytesIO()
        office_file = msoffcrypto.OfficeFile(file_stream)
        try:
            office_file.load_key(password=password)
            office_file.decrypt(decrypted)
            decrypted.seek(0)
            return docx.Document(decrypted)
        except Exception as e:
            raise Exception(f"Failed to decrypt document: {str(e)}")
    else:
        # Handle unencrypted document
        try:
            return docx.Document(file_stream)
        except Exception as e:
            # If document is encrypted but no password provided
            if "encrypted" in str(e).lower():
                raise Exception("Document is encrypted but no password was provided")
            raise Exception(f"Failed to load document: {str(e)}")


# ================================
# CONTENT EXTRACTION FUNCTIONS
# ================================

def extract_docx_content_enhanced(doc: docx.Document) -> List[Dict[str, Any]]:
    """
    Extract content from DOCX with enhanced strikethrough text handling.
    
    Args:
        doc (docx.Document): Document object to extract content from
        
    Returns:
        List[Dict]: List of content elements with metadata
    """
    content = []
    
    # Extract paragraphs with enhanced formatting detection
    for para_idx, para in enumerate(doc.paragraphs, start=1):
        if para.text.strip():
            runs_data = []
            strikethrough_found = False
            strikethrough_texts = []
            
            # Analyze each text run in the paragraph
            for run_idx, run in enumerate(para.runs):
                run_text = run.text.strip()
                if run_text:
                    # Detect formatting properties
                    strike = detect_strikethrough(run)
                    
                    run_data = {
                        "text": run_text,
                        "strikethrough": strike,
                        "bold": run.font.bold,
                        "italic": run.font.italic,
                        "underline": run.font.underline,
                        "font_name": run.font.name,
                        "font_size": run.font.size,
                        "font_color": run.font.color.rgb if run.font.color else None,
                    }
                    
                    if strike:
                        strikethrough_found = True
                        strikethrough_texts.append(run_text)
                    
                    runs_data.append(run_data)
            
            # Combine strikethrough texts
            strikethrough_text_combined = " | ".join(strikethrough_texts) if strikethrough_texts else ""
            
            content.append({
                "type": "paragraph",
                "index": para_idx,
                "text": para.text,
                "runs": runs_data,
                "style": para.style.name if para.style else "Normal",
                "has_strikethrough": strikethrough_found,
                "strikethrough_texts": strikethrough_texts,
                "strikethrough_combined": strikethrough_text_combined
            })
    
    # Extract tables (excluding first two tables which are often templates)
    for table_idx, table in enumerate(doc.tables, start=1):
        if table_idx <= 2:  # Skip first two tables
            continue
            
        table_data = {
            "type": "table",
            "index": table_idx,
            "rows": [],
            "has_strikethrough": False,
            "strikethrough_texts": []
        }
        
        # Process each row and cell in the table
        for row_idx, row in enumerate(table.rows, start=1):
            row_data = {"cells": []}
            for cell_idx, cell in enumerate(row.cells, start=1):
                cell_text = "\n".join(p.text for p in cell.paragraphs if p.text.strip())
                
                cell_has_strikethrough = False
                cell_strikethrough_texts = []
                
                # Check for strikethrough in cell content
                for para in cell.paragraphs:
                    for run in para.runs:
                        run_text = run.text.strip()
                        if run_text:
                            strike = detect_strikethrough(run)
                            if strike:
                                cell_has_strikethrough = True
                                cell_strikethrough_texts.append(run_text)
                                table_data["has_strikethrough"] = True
                                table_data["strikethrough_texts"].append(run_text)
                
                # Combine strikethrough texts
                cell_strikethrough_combined = " | ".join(cell_strikethrough_texts) if cell_strikethrough_texts else ""
                
                row_data["cells"].append({
                    "text": cell_text,
                    "row": row_idx,
                    "col": cell_idx,
                    "contains_tbd": "TBD" in cell_text.upper(),
                    "contains_code_names": bool(re.search(r'\[[A-Z]+\]', cell_text)),
                    "urls": re.findall(r'https?://[^\s<>"]+|www\.[^\s<>"]+', cell_text),
                    "has_strikethrough": cell_has_strikethrough,
                    "strikethrough_texts": cell_strikethrough_texts,
                    "strikethrough_combined": cell_strikethrough_combined
                })
            
            table_data["rows"].append(row_data)
        
        content.append(table_data)
    
    return content


def detect_strikethrough(run) -> bool:
    """
    Detect if text has strikethrough formatting using multiple methods.
    
    Args:
        run: DOCX text run object
        
    Returns:
        bool: True if strikethrough is detected
    """
    try:
        # Method 1: Check font property directly
        if run.font.strike:
            return True
            
        # Method 2: Check XML directly for more reliable detection
        if hasattr(run, '_element'):
            # Check for w:strike elements
            strike_elements = run._element.xpath('.//w:strike')
            if strike_elements:
                for strike_elm in strike_elements:
                    val = strike_elm.get(qn('w:val'))
                    if val is None or val != "false":
                        return True
            
            # Check for w:dstrike elements (double strikethrough)
            dstrike_elements = run._element.xpath('.//w:dstrike')
            if dstrike_elements:
                for dstrike_elem in dstrike_elements:
                    val = dstrike_elem.get(qn('w:val'))
                    if val is None or val in ['true', '1', 'on']:
                        return True
        
        # Method 3: Check run properties via XML
        if hasattr(run, '_r'):
            r_pr = run._r.rPr
            if r_pr is not None:
                strike = r_pr.find(qn('w:strike'))
                if strike is not None:
                    val = strike.get(qn('w:val'))
                    if val is None or val not in ['false', '0', 'off']:
                        return True
                    
        return False
    except Exception as e:
        print(f"Error in detect_strikethrough: {e}")
        return False

def detect_panels(doc: docx.Document) -> List[Dict[str, Any]]:
    """
    Detect panels in document based on section breaks, headings, or other markers.
    
    Args:
        doc (docx.Document): Document object
        
    Returns:
        List[Dict]: List of panel information
    """
    panels = []
    
    # Look for section breaks
    for sect in doc.sections:
        panels.append({
            "type": "section",
            "start": "beginning",
            "content": "Section content"
        })
    
    # Look for headings that might indicate panels
    for para in doc.paragraphs:
        if para.style.name.startswith('Heading'):
            panels.append({
                "type": "heading",
                "level": int(para.style.name.replace('Heading', '')) if para.style.name.replace('Heading', '').isdigit() else 1,
                "text": para.text,
                "position": len(panels) + 1
            })
    
    return panels

# ================================
# IMAGE HANDLING FUNCTIONS
# ================================

def get_docx_images(doc: docx.Document) -> List[Dict[str, Any]]:
    """
    Extract all images from DOCX document with comprehensive detection.
    
    Args:
        doc (docx.Document): Document object
        
    Returns:
        List[Dict]: List of image information dictionaries
    """
    images = []
    
    try:
        print("DEBUG: Starting image extraction from DOCX...")
        
        # Method 1: Check document relationships
        if hasattr(doc, 'part') and hasattr(doc.part, 'rels'):
            for rel_id, rel in doc.part.rels.items():
                if "image" in rel.reltype:
                    try:
                        img_bytes = rel.target_part.blob
                        img_hash = md5(img_bytes).hexdigest()
                        
                        # Get image dimensions using PIL
                        pil_img = PILImage.open(io.BytesIO(img_bytes))
                        width, height = pil_img.size
                        
                        images.append({
                            "hash": img_hash,
                            "bytes": img_bytes,
                            "width": width,
                            "height": height,
                            "format": pil_img.format,
                            "rel_id": rel_id,
                            "source": "document_relationships"
                        })
                    except Exception as e:
                        print(f"DEBUG: Error processing relationship image {rel_id}: {e}")
                        continue
        
        # Method 2: Check inline shapes
        try:
            for element in doc.element.body:
                for inline in element.xpath('.//wp:inline', namespaces={
                    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
                }):
                    blip = inline.xpath('.//a:blip', namespaces={
                        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
                    })
                    if blip:
                        embed_id = blip[0].get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                        if embed_id and embed_id in doc.part.rels:
                            try:
                                img_part = doc.part.rels[embed_id].target_part
                                img_bytes = img_part.blob
                                img_hash = md5(img_bytes).hexdigest()
                                
                                pil_img = PILImage.open(io.BytesIO(img_bytes))
                                width, height = pil_img.size
                                
                                images.append({
                                    "hash": img_hash,
                                    "bytes": img_bytes,
                                    "width": width,
                                    "height": height,
                                    "format": pil_img.format,
                                    "embed_id": embed_id,
                                    "source": "inline_shapes"
                                })
                            except Exception as e:
                                print(f"DEBUG: Error processing inline image {embed_id}: {e}")
                                continue
        except Exception as e:
            print(f"DEBUG: Error processing inline shapes: {e}")
        
        print(f"DEBUG: Total images extracted: {len(images)}")
        
    except Exception as e:
        print(f"DEBUG: Error in get_docx_images: {e}")
    
    return images


def get_table_images(doc: docx.Document) -> List[Dict[str, Any]]:
    """
    Extract images from inside table cells with proper XML parsing.
    
    Args:
        doc (docx.Document): Document object
        
    Returns:
        List[Dict]: List of table image information
    """
    table_images = []
    
    try:
        # Namespace map for XML parsing
        nsmap = {
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        }
        
        for table_idx, table in enumerate(doc.tables, start=1):
            if table_idx <= 2:  # Skip first two tables
                continue
            
            for row_idx, row in enumerate(table.rows, start=1):
                for cell_idx, cell in enumerate(row.cells, start=1):
                    # Convert cell to XML element for parsing
                    cell_xml = cell._element
                    
                    # Find all graphic elements in the cell
                    for graphic in cell_xml.xpath('.//w:drawing//a:graphic', namespaces=nsmap):
                        # Find blip elements which reference images
                        for blip in graphic.xpath('.//a:blip', namespaces=nsmap):
                            embed_id = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                            if embed_id and embed_id in doc.part.rels:
                                try:
                                    img_part = doc.part.rels[embed_id].target_part
                                    img_bytes = img_part.blob
                                    img_hash = md5(img_bytes).hexdigest()
                                    
                                    # Get image dimensions
                                    pil_img = PILImage.open(io.BytesIO(img_bytes))
                                    width, height = pil_img.size
                                    
                                    table_images.append({
                                        "table": table_idx,
                                        "row": row_idx,
                                        "col": cell_idx,
                                        "hash": img_hash,
                                        "bytes": img_bytes,
                                        "width": width,
                                        "height": height,
                                        "format": pil_img.format,
                                        "embed_id": embed_id
                                    })
                                except Exception as e:
                                    print(f"Error processing table image: {e}")
                                    continue
    except Exception as e:
        print(f"Error extracting table images: {e}")
    
    return table_images


def insert_image_into_excel(ws, img_bytes: bytes, cell_address: str, 
                           max_width: int = 120, max_height: int = 120) -> bool:
    """
    Insert an image into Excel worksheet with size adjustment.
    
    Args:
        ws: Excel worksheet object
        img_bytes (bytes): Image data
        cell_address (str): Target cell address (e.g., 'A1')
        max_width (int): Maximum image width
        max_height (int): Maximum image height
        
    Returns:
        bool: True if successful, False otherwise
    """
    if not img_bytes:
        return False
        
    try:
        # Create temporary file for the image
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_file:
            # Open and process the image
            pil_img = PILImage.open(io.BytesIO(img_bytes))
            
            # Convert to RGB if necessary
            if pil_img.mode in ('RGBA', 'LA', 'P'):
                background = PILImage.new('RGB', pil_img.size, (255, 255, 255))
                if pil_img.mode == 'P':
                    pil_img = pil_img.convert('RGBA')
                background.paste(pil_img, mask=pil_img.split()[-1] if pil_img.mode == 'RGBA' else None)
                pil_img = background
            
            # Resize image if too large
            original_width, original_height = pil_img.size
            width_ratio = max_width / original_width
            height_ratio = max_height / original_height
            ratio = min(width_ratio, height_ratio, 1.0)  # Don't enlarge small images
            
            if ratio < 1.0:
                new_width = int(original_width * ratio)
                new_height = int(original_height * ratio)
                pil_img = pil_img.resize((new_width, new_height), PILImage.Resampling.LANCZOS)
            
            # Save as PNG for better compatibility
            pil_img.save(tmp_file.name, format="PNG", optimize=True)
            
            # Create OpenPyXL image and add to worksheet
            img = XLImage(tmp_file.name)
            img.anchor = cell_address
            
            # Adjust cell size to fit image
            col_letter = cell_address[0]  # Get column letter
            row_num = int(cell_address[1:])  # Get row number
            
            # Set column width based on image width
            col_width = max(8.43, min(50, img.width * 0.14))
            ws.column_dimensions[col_letter].width = col_width
            
            # Set row height based on image height
            row_height = max(15, min(400, img.height * 0.75))
            ws.row_dimensions[row_num].height = row_height
            
            ws.add_image(img)
            
            # Clean up temporary file
            os.unlink(tmp_file.name)
            return True
            
    except Exception as e:
        print(f"Error inserting image into Excel: {e}")
        # Clean up temporary file if it exists
        if 'tmp_file' in locals() and os.path.exists(tmp_file.name):
            os.unlink(tmp_file.name)
        return False


# ================================
# COMPARISON FUNCTIONS
# ================================

def compare_images_enhanced(imgs1: List[Dict], imgs2: List[Dict]) -> List[Dict[str, Any]]:
    """
    Enhanced image comparison that properly captures deleted images.
    
    Args:
        imgs1 (List): Images from first document
        imgs2 (List): Images from second document
        
    Returns:
        List[Dict]: List of image changes
    """
    changes = []
    
    # Group images by hash for comparison
    hashes1 = {img["hash"]: img for img in imgs1}
    hashes2 = {img["hash"]: img for img in imgs2}
    
    print(f"DEBUG: Document 1 has {len(imgs1)} images, Document 2 has {len(imgs2)} images")
    
    # Find deleted images (in doc1 but not in doc2)
    for img_hash, img1 in hashes1.items():
        if img_hash not in hashes2:
            changes.append({
                "type": "image",
                "status": "deleted",
                "hash": img_hash,
                "old_img": img1["bytes"],
                "new_img": None,
                "width": img1["width"],
                "height": img1["height"],
                "alignment": "document_level",
                "change_detail": f"Image deleted - Original size: {img1['width']}x{img1['height']}"
            })
    
    # Find added images (in doc2 but not in doc1)
    for img_hash, img2 in hashes2.items():
        if img_hash not in hashes1:
            changes.append({
                "type": "image",
                "status": "added",
                "hash": img_hash,
                "old_img": None,
                "new_img": img2["bytes"],
                "width": img2["width"],
                "height": img2["height"],
                "alignment": "document_level",
                "change_detail": f"Image added - Size: {img2['width']}x{img2['height']}"
            })
    
    # Find modified images (same hash but different properties)
    common_hashes = set(hashes1.keys()) & set(hashes2.keys())
    for img_hash in common_hashes:
        img1 = hashes1[img_hash]
        img2 = hashes2[img_hash]
        
        # Check for size changes
        if img1["width"] != img2["width"] or img1["height"] != img2["height"]:
            changes.append({
                "type": "image",
                "status": "modified",
                "hash": img_hash,
                "old_img": img1["bytes"],
                "new_img": img2["bytes"],
                "width_original": img1["width"],
                "height_original": img1["height"],
                "width_modified": img2["width"],
                "height_modified": img2["height"],
                "alignment": "document_level",
                "change_detail": f"Size changed from {img1['width']}x{img1['height']} to {img2['width']}x{img2['height']}"
            })
    
    print(f"DEBUG: Total image changes detected: {len(changes)}")
    return changes


def compare_table_images_enhanced(imgs1: List[Dict], imgs2: List[Dict]) -> List[Dict[str, Any]]:
    """
    Enhanced table image comparison that properly captures deleted images.
    
    Args:
        imgs1 (List): Table images from first document
        imgs2 (List): Table images from second document
        
    Returns:
        List[Dict]: List of table image changes
    """
    changes = []
    
    print(f"DEBUG: Table images - Doc1: {len(imgs1)}, Doc2: {len(imgs2)}")
    
    # Group by table, row, column AND hash
    all_images1 = {(img["table"], img["row"], img["col"], img["hash"]): img for img in imgs1}
    all_images2 = {(img["table"], img["row"], img["col"], img["hash"]): img for img in imgs2}
    
    # Find deleted table images
    for img_key, img1 in all_images1.items():
        if img_key not in all_images2:
            table, row, col, img_hash = img_key
            changes.append({
                "type": "table_image",
                "status": "deleted",
                "table": table,
                "row": row,
                "col": col,
                "hash": img_hash,
                "old_img": img1["bytes"],
                "new_img": None,
                "width": img1["width"],
                "height": img1["height"],
                "alignment": f"Table {table}, Cell ({row},{col})",
                "change_detail": f"Table image deleted from Table {table}, Cell ({row},{col})"
            })
    
    # Find added table images
    for img_key, img2 in all_images2.items():
        if img_key not in all_images1:
            table, row, col, img_hash = img_key
            changes.append({
                "type": "table_image",
                "status": "added",
                "table": table,
                "row": row,
                "col": col,
                "hash": img_hash,
                "old_img": None,
                "new_img": img2["bytes"],
                "width": img2["width"],
                "height": img2["height"],
                "alignment": f"Table {table}, Cell ({row},{col})",
                "change_detail": f"Table image added to Table {table}, Cell ({row},{col})"
            })
    
    # Check for moved images (same hash, different position)
    all_hashes1 = {img["hash"] for img in imgs1}
    all_hashes2 = {img["hash"] for img in imgs2}
    common_hashes = all_hashes1 & all_hashes2
    
    for img_hash in common_hashes:
        # Find all occurrences of this image in both documents
        imgs1_with_hash = [img for img in imgs1 if img["hash"] == img_hash]
        imgs2_with_hash = [img for img in imgs2 if img["hash"] == img_hash]
        
        # Check if the image moved to a different location
        for img1 in imgs1_with_hash:
            for img2 in imgs2_with_hash:
                if (img1["table"] != img2["table"] or 
                    img1["row"] != img2["row"] or 
                    img1["col"] != img2["col"]):
                    
                    changes.append({
                        "type": "table_image",
                        "status": "moved",
                        "hash": img_hash,
                        "old_img": img1["bytes"],
                        "new_img": img2["bytes"],
                        "old_table": img1["table"],
                        "old_row": img1["row"],
                        "old_col": img1["col"],
                        "new_table": img2["table"],
                        "new_row": img2["row"],
                        "new_col": img2["col"],
                        "alignment": f"Moved from Table {img1['table']}, Cell ({img1['row']},{img1['col']}) to Table {img2['table']}, Cell ({img2['row']},{img2['col']})",
                        "change_detail": f"Image moved from Table {img1['table']}, Cell ({img1['row']},{img1['col']}) to Table {img2['table']}, Cell ({img2['row']},{img2['col']})"
                    })
    
    print(f"DEBUG: Total table image changes detected: {len(changes)}")
    return changes


def compare_text_content_enhanced(content1: List[Dict], content2: List[Dict]) -> List[Dict[str, Any]]:
    """
    Enhanced text comparison with detailed change tracking and strikethrough detection.
    
    Args:
        content1 (List): Content from first document
        content2 (List): Content from second document
        
    Returns:
        List[Dict]: List of text changes
    """
    changes = []
    strikethrough_changes = []  # Separate list for strikethrough changes
    processed_strikethrough_paragraphs = set()
    
    # Compare paragraphs
    para1 = [item for item in content1 if item["type"] == "paragraph"]
    para2 = [item for item in content2 if item["type"] == "paragraph"]
    
    print(f"DEBUG: Comparing {len(para1)} vs {len(para2)} paragraphs")
    
    for i in range(max(len(para1), len(para2))):
        if i < len(para1) and i < len(para2):
            p1 = para1[i]
            p2 = para2[i]
            
            # Get strikethrough information
            strike_original = p1["has_strikethrough"]
            strike_modified = p2["has_strikethrough"]
            strike_text_original = p1["strikethrough_combined"]
            strike_text_modified = p2["strikethrough_combined"]
            
            # Enhanced strikethrough detection
            strikethrough_runs_original = [run for run in p1["runs"] if run.get("strikethrough", False)]
            strikethrough_runs_modified = [run for run in p2["runs"] if run.get("strikethrough", False)]
            
            strikethrough_original = len(strikethrough_runs_original) > 0
            strikethrough_modified = len(strikethrough_runs_modified) > 0
            
            # Capture strikethrough text content
            strikethrough_text_original = " | ".join([run["text"] for run in strikethrough_runs_original])
            strikethrough_text_modified = " | ".join([run["text"] for run in strikethrough_runs_modified])
            
            #check if we've already processed strikethrough for this paragraph
            paragraph_key = f"para_{i+1}"
            
            # SEPARATE STRIKETHROUGH CHANGES - only add to strikethrough_changes list
            if (strike_original or strike_modified) and paragraph_key not in processed_strikethrough_paragraphs:
                if strike_original and strike_text_original.strip():
                    strikethrough_changes.append({
                        "type": "paragraph_strikethrough",
                        "index": i + 1,
                        "status": "strikethrough_original",
                        "original": p1["text"],
                        "modified": p2["text"],
                        "strikethrough_original": True,
                        "strikethrough_text_original": strike_text_original,
                        "strikethrough_modified": strike_modified,
                        "strikethrough_text_modified": strike_text_modified,
                        "change_detail": f"Strikethrough text found in original: {strike_text_original}",
                    })
                    processed_strikethrough_paragraphs.add(paragraph_key)
                
                if strike_modified and strike_text_modified.strip():
                    # Only add if it's new strikethrough (not already reported above)
                    if not strike_original or strike_text_original != strike_text_modified:
                        strikethrough_changes.append({
                            "type": "paragraph_strikethrough",
                            "index": i + 1,
                            "status": "strikethrough_modified",
                            "original": p1["text"],
                            "modified": p2["text"],
                            "strikethrough_original": strike_original,
                            "strikethrough_text_original": strike_text_original,
                            "strikethrough_modified": True,
                            "strikethrough_text_modified": strike_text_modified,
                            "change_detail": f"Strikethrough text found in modified: {strike_text_modified}",
                        })
                        
            # Check for text changes
            if p1["text"] != p2["text"]:
                # Use difflib to show exact differences
                diff = list(difflib.ndiff(p1["text"].split(), p2["text"].split()))
                added = [word[2:] for word in diff if word.startswith('+ ')]
                removed = [word[2:] for word in diff if word.startswith('- ')]
                
                # Check for various content patterns
                double_spaces_original = len(re.findall(r'  ', p1["text"]))
                double_spaces_modified = len(re.findall(r'  ', p2["text"]))
                
                tbd_original = "TBD" in p1["text"].upper()
                tbd_modified = "TBD" in p2["text"].upper()
                
                code_names_original = bool(re.search(r'\[[A-Z]+\]', p1["text"]))
                code_names_modified = bool(re.search(r'\[[A-Z]+\]', p2["text"]))
                
                urls_original = re.findall(r'https?://[^\s<>"]+|www\.[^\s<>"]+', p1["text"])
                urls_modified = re.findall(r'https?://[^\s<>"]+|www\.[^\s<>"]+', p2["text"])
                
                 # Only add to regular changes if there are actual text changes beyond strikethrough
                has_non_strikethrough_changes = (
                    added or removed or 
                    double_spaces_original != double_spaces_modified or
                    tbd_original != tbd_modified or
                    code_names_original != code_names_modified or
                    urls_original != urls_modified
                )
                
                if has_non_strikethrough_changes:
                    changes.append({
                        "type": "paragraph",
                        "index": i + 1,
                        "status": "modified",
                        "original": p1["text"],
                        "modified": p2["text"],
                        "added": " ".join(added),
                        "removed": " ".join(removed),
                        "double_spaces_original": double_spaces_original,
                        "double_spaces_modified": double_spaces_modified,
                        "tbd_original": tbd_original,
                        "tbd_modified": tbd_modified,
                        "code_names_original": code_names_original,
                        "code_names_modified": code_names_modified,
                        "urls_original": urls_original,
                        "urls_modified": urls_modified,
                        # Don't include strikethrough info in regular changes
                    })
            
                
            # Check for formatting changes even if text is the same
            elif p1["text"] == p2["text"] and paragraph_key not in processed_strikethrough_paragraphs:
                formatting_changes = []
                strikethrough_format_changes = []
                
                for run_idx in range(max(len(p1["runs"]), len(p2["runs"]))):
                    if run_idx < len(p1["runs"]) and run_idx < len(p2["runs"]):
                        run1 = p1["runs"][run_idx]
                        run2 = p2["runs"][run_idx]
                        
                        if run1["text"] == run2["text"]:
                            # Check for formatting differences
                            format_diffs = []
                            non_strikethrough_diffs = []
                            
                            for run_idx in range(max(len(p1["runs"]), len(p2["runs"]))):
                                if run_idx < len(p1["runs"]) and run_idx < len(p2["runs"]):
                                    run1 = p1["runs"][run_idx]
                                    run2 = p2["runs"][run_idx]
                                    
                                    if run1["text"] == run2["text"]:
                                        # Check for formatting differences
                                        format_diffs = []
                                        non_strikethrough_diffs = []
                                        
                                        if run1.get("bold") != run2.get("bold"):
                                            format_diffs.append("bold")
                                            non_strikethrough_diffs.append("bold")
                                        if run1.get("italic") != run2.get("italic"):
                                            format_diffs.append("italic")
                                            non_strikethrough_diffs.append("italic")
                                        if run1.get("underline") != run2.get("underline"):
                                            format_diffs.append("underline")
                                            non_strikethrough_diffs.append("underline")
                                        
                                        # Handle strikethrough formatting separately
                                        if run1.get("strikethrough") != run2.get("strikethrough"):
                                            format_diffs.append("strikethrough")
                                            if run1.get("strikethrough") and not run2.get("strikethrough"):
                                                strikethrough_format_changes.append(f"strikethrough removed: '{run1['text']}'")
                                            elif not run1.get("strikethrough") and run2.get("strikethrough"):
                                                strikethrough_format_changes.append(f"Strikethrough added: '{run1['text']}'")
                                        
                                        if run1.get("font_name") != run2.get("font_name"):
                                            format_diffs.append("font")
                                            non_strikethrough_diffs.append("font")
                                        if run1.get("font_size") != run2.get("font_size"):
                                            format_diffs.append("font size")
                                            non_strikethrough_diffs.append("font size")
                                        if run1.get("font_color") != run2.get("font_color"):
                                            format_diffs.append("font color")
                                            non_strikethrough_diffs.append("font color")
                                            
                                        if format_diffs:
                                            # Only add to regular changes if there are non-strikethrough formatting changes
                                            if non_strikethrough_diffs:
                                                formatting_changes.append({
                                                    "text": run1["text"],
                                                    "changes": non_strikethrough_diffs
                                                })
                            
                            # Add strikethrough formatting changes to strikethrough_changes list
                            if strikethrough_format_changes and paragraph_key not in processed_strikethrough_paragraphs:
                                unique_strikethrough_changes = list(set(strikethrough_format_changes))
                                strikethrough_changes.append({
                                    "type": "paragraph_formatting_strikethrough",
                                    "index": i + 1,
                                    "status": "strikethrough_formatting_changed",
                                    "original": p1["text"],
                                    "modified": p2["text"],
                                    "formatting_changes": unique_strikethrough_changes,
                                    "change_detail": "Strikethrough formatting changes:\n" + "\n".join(unique_strikethrough_changes),
                                })
                                processed_strikethrough_paragraphs.add(paragraph_key)
                            
                            # Only add to regular changes if there are non-strikethrough formatting changes
                            if formatting_changes:
                                change_detail = "Formatting changes:\n"
                                for fmt_change in formatting_changes:
                                    change_detail += f"'{fmt_change['text']}': {', '.join(fmt_change['changes'])}\n"
                                
                                changes.append({
                                    "type": "paragraph_formatting",
                                    "index": i + 1,
                                    "status": "formatting_changed",
                                    "original": p1["text"],
                                    "modified": p2["text"],
                                    "formatting_changes": formatting_changes,
                                    "change_detail": change_detail,
                                })
                            
                            
        elif i < len(para1):  # Deleted paragraph
            p1 = para1[i]
            # Check for strikethrough in deleted paragraph
            strikethrough_runs = [run for run in p1["runs"] if run.get("strikethrough", False)]
            strikethrough_text = " | ".join([run["text"] for run in strikethrough_runs])
            
            # Add strikethrough info to strikethrough_changes if present
            if strikethrough_text.strip():
                strikethrough_changes.append({
                    "type": "paragraph_strikethrough",
                    "index": i + 1,
                    "status": "deleted_strikethrough",
                    "original": p1["text"],
                    "modified": "",
                    "strikethrough_original": True,
                    "strikethrough_text_original": strikethrough_text,
                    "change_detail": f"Deleted paragraph contained strikethrough: {strikethrough_text}",
                })
            
            # Always add to regular changes for deleted paragraphs
            changes.append({
                "type": "paragraph",
                "index": i + 1,
                "status": "deleted",
                "original": p1["text"],
                "modified": "",
                "added": "",
                "removed": p1["text"],
            })
                
        elif i < len(para2):  # Added paragraph
            p2 = para2[i]
            # Check for strikethrough in added paragraph
            strikethrough_runs = [run for run in p2["runs"] if run.get("strikethrough", False)]
            strikethrough_text = " | ".join([run["text"] for run in strikethrough_runs])
            
            # Add strikethrough info to strikethrough_changes if present
            if strikethrough_text.strip():
                strikethrough_changes.append({
                    "type": "paragraph_strikethrough",
                    "index": i + 1,
                    "status": "added_strikethrough",
                    "original": "",
                    "modified": p2["text"],
                    "strikethrough_modified": True,
                    "strikethrough_text_modified": strikethrough_text,
                    "change_detail": f"Added paragraph contains strikethrough: {strikethrough_text}",
                })
            
            # Always add to regular changes for added paragraphs
            changes.append({
                "type": "paragraph",
                "index": i + 1,
                "status": "added",
                "original": "",
                "modified": p2["text"],
                "added": p2["text"],
                "removed": "",
            })
    # Remove duplicate strikethrough changes by paragraph
    unique_strikethrough_changes = []
    seen_paragraphs = set()
    
    for change in strikethrough_changes:
        para_key = f"para_{change.get('index')}"
        if para_key not in seen_paragraphs:
            unique_strikethrough_changes.append(change)
            seen_paragraphs.add(para_key)
        else:
            print(f"DEBUG: Removed duplicate strikethrough change for paragraph {change.get('index')}")
    
    # Combine regular changes with strikethrough changes (they'll be separated in Excel export)
    all_changes = changes + unique_strikethrough_changes
    print(f"DEBUG: Final changes - Regular: {len(changes)}, Strikethrough: {len(unique_strikethrough_changes)}, Total: {len(all_changes)}")
    return all_changes
    
    
def compare_tables(content1: List[Dict], content2: List[Dict]) -> List[Dict[str, Any]]:
    """
    Compare tables between two documents without duplication.
    Now separates strikethrough changes from regular table changes.
    """
    changes = []
    strikethrough_changes = []  # Separate list for table strikethrough changes
    
    tables1 = [item for item in content1 if item["type"] == "table"]
    tables2 = [item for item in content2 if item["type"] == "table"]
    
    for i in range(max(len(tables1), len(tables2))):
        if i < len(tables1) and i < len(tables2):
            table1 = tables1[i]
            table2 = tables2[i]
            
            # Compare table-level strikethrough - add to strikethrough_changes only
            if table1["has_strikethrough"] or table2["has_strikethrough"]:
                strike1_text = " | ".join(table1.get("strikethrough_texts", []))
                strike2_text = " | ".join(table2.get("strikethrough_texts", []))
                
                if strike1_text != strike2_text:
                    strikethrough_changes.append({
                        "type": "table_strikethrough",
                        "table_index": i + 1,
                        "status": "modified",
                        "original_strikethrough": strike1_text,
                        "modified_strikethrough": strike2_text,
                        "original_has_strike": table1["has_strikethrough"],
                        "modified_has_strike": table2["has_strikethrough"],
                        "change_detail": f"Table strikethrough content changed"
                    })

            # Compare row count - regular change
            if len(table1["rows"]) != len(table2["rows"]):
                changes.append({
                    "type": "table",
                    "index": i + 1,
                    "status": "structure_changed",
                    "detail": f"Row count changed from {len(table1['rows'])} to {len(table2['rows'])}",
                    "original": f"Table with {len(table1['rows'])} rows",
                    "modified": f"Table with {len(table2['rows'])} rows"
                })
                
            # Compare cell content - SINGLE PASS to avoid duplication
            for r_idx in range(max(len(table1["rows"]), len(table2["rows"]))):
                if r_idx < len(table1["rows"]) and r_idx < len(table2["rows"]):
                    row1 = table1["rows"][r_idx]
                    row2 = table2["rows"][r_idx]
                    
                    # Compare column count - regular change
                    if len(row1["cells"]) != len(row2["cells"]):
                        changes.append({
                            "type": "table",
                            "index": i + 1,
                            "status": "structure_changed",
                            "detail": f"Column count changed in row {r_idx+1} from {len(row1['cells'])} to {len(row2['cells'])}",
                            "original": f"Row with {len(row1['cells'])} columns",
                            "modified": f"Row with {len(row2['cells'])} columns"
                        })
                    
                    # Compare cell content - process each cell only once
                    for c_idx in range(max(len(row1["cells"]), len(row2["cells"]))):
                        if c_idx < len(row1["cells"]) and c_idx < len(row2["cells"]):
                            cell1 = row1["cells"][c_idx]
                            cell2 = row2["cells"][c_idx]
                            
                            # Check for strikethrough changes separately
                            strike1 = " | ".join(cell1.get("strikethrough_texts", []))
                            strike2 = " | ".join(cell2.get("strikethrough_texts", []))
                            
                            has_strikethrough_changes = (
                                cell1["has_strikethrough"] != cell2["has_strikethrough"] or
                                (cell1["has_strikethrough"] and cell2["has_strikethrough"] and strike1 != strike2)
                            )
                            
                            has_text_changes = cell1["text"] != cell2["text"]
                            
                            # Handle strikethrough changes separately
                            if has_strikethrough_changes:
                                strikethrough_changes.append({
                                    "type": "table_cell_strikethrough",
                                    "table_index": i + 1,
                                    "row": r_idx + 1,
                                    "col": c_idx + 1,
                                    "status": "strikethrough_modified",
                                    "original": cell1["text"],
                                    "modified": cell2["text"],
                                    "strikethrough_original": strike1,
                                    "strikethrough_modified": strike2,
                                    "has_strikethrough_original": cell1["has_strikethrough"],
                                    "has_strikethrough_modified": cell2["has_strikethrough"],
                                    "change_detail": "Strikethrough content changed in table cell"
                                })
                            
                            # Handle regular text changes
                            if has_text_changes and not has_strikethrough_changes:
                                changes.append({
                                    "type": "table_cell",
                                    "table_index": i + 1,
                                    "row": r_idx + 1,
                                    "col": c_idx + 1,
                                    "status": "modified",
                                    "original": cell1["text"],
                                    "modified": cell2["text"],
                                    "change_detail": "Text content modified in table cell"
                                })
                            elif has_text_changes and has_strikethrough_changes:
                                # If both text and strikethrough changed, include text change but not strikethrough info
                                changes.append({
                                    "type": "table_cell",
                                    "table_index": i + 1,
                                    "row": r_idx + 1,
                                    "col": c_idx + 1,
                                    "status": "modified",
                                    "original": cell1["text"],
                                    "modified": cell2["text"],
                                    "change_detail": "Text content modified in table cell"
                                })
                                
                        elif c_idx < len(row1["cells"]):
                            # Cell deleted - check for strikethrough
                            cell1 = row1["cells"][c_idx]
                            strike1 = " | ".join(cell1.get("strikethrough_texts", []))
                            
                            if cell1["has_strikethrough"] and strike1.strip():
                                strikethrough_changes.append({
                                    "type": "table_cell_strikethrough",
                                    "table_index": i + 1,
                                    "row": r_idx + 1,
                                    "col": c_idx + 1,
                                    "status": "deleted_strikethrough",
                                    "original": cell1["text"],
                                    "modified": "",
                                    "strikethrough_original": strike1,
                                    "has_strikethrough_original": True,
                                    "change_detail": "Deleted cell contained strikethrough text"
                                })
                            
                            changes.append({
                                "type": "table_cell",
                                "table_index": i + 1,
                                "row": r_idx + 1,
                                "col": c_idx + 1,
                                "status": "deleted",
                                "original": cell1["text"],
                                "modified": "",
                                "change_detail": "Cell deleted from table"
                            })
                        elif c_idx < len(row2["cells"]):
                            # Cell added - check for strikethrough
                            cell2 = row2["cells"][c_idx]
                            strike2 = " | ".join(cell2.get("strikethrough_texts", []))
                            
                            if cell2["has_strikethrough"] and strike2.strip():
                                strikethrough_changes.append({
                                    "type": "table_cell_strikethrough",
                                    "table_index": i + 1,
                                    "row": r_idx + 1,
                                    "col": c_idx + 1,
                                    "status": "added_strikethrough",
                                    "original": "",
                                    "modified": cell2["text"],
                                    "strikethrough_modified": strike2,
                                    "has_strikethrough_modified": True,
                                    "change_detail": "Added cell contains strikethrough text"
                                })
                            
                            changes.append({
                                "type": "table_cell",
                                "table_index": i + 1,
                                "row": r_idx + 1,
                                "col": c_idx + 1,
                                "status": "added",
                                "original": "",
                                "modified": cell2["text"],
                                "change_detail": "Cell added to table"
                            })
                            
                elif r_idx < len(table1["rows"]):
                    # Entire row deleted - regular change
                    row1 = table1["rows"][r_idx]
                    changes.append({
                        "type": "table_row",
                        "table_index": i + 1,
                        "row": r_idx + 1,
                        "status": "deleted",
                        "original": f"Row {r_idx+1} with {len(row1['cells'])} columns",
                        "modified": "",
                        "change_detail": f"Row {r_idx+1} deleted from table"
                    })
                elif r_idx < len(table2["rows"]):
                    # Entire row added - regular change
                    row2 = table2["rows"][r_idx]
                    changes.append({
                        "type": "table_row",
                        "table_index": i + 1,
                        "row": r_idx + 1,
                        "status": "added",
                        "original": "",
                        "modified": f"Row {r_idx+1} with {len(row2['cells'])} columns",
                        "change_detail": f"Row {r_idx+1} added to table"
                    })
                    
        elif i < len(tables1):
            # Entire table deleted - regular change
            changes.append({
                "type": "table",
                "index": i + 1,
                "status": "deleted",
                "original": f"Table with {len(tables1[i]['rows'])} rows",
                "modified": "",
                "change_detail": f"Table {i+1} deleted from document"
            })
        elif i < len(tables2):
            # Entire table added - regular change
            changes.append({
                "type": "table",
                "index": i + 1,
                "status": "added",
                "original": "",
                "modified": f"Table with {len(tables2[i]['rows'])} rows",
                "change_detail": f"Table {i+1} added to document"
            })
    
    # Combine regular changes with strikethrough changes
    all_changes = changes + strikethrough_changes
    return all_changes


def compare_panels(panels1: List[Dict], panels2: List[Dict]) -> List[Dict[str, Any]]:
    """
    Compare panels between documents.
    
    Args:
        panels1 (List): Panels from first document
        panels2 (List): Panels from second document
        
    Returns:
        List[Dict]: List of panel changes
    """
    changes = []
    
    # Simple comparison based on panel count and content
    if len(panels1) != len(panels2):
        changes.append({
            "type": "panel_count",
            "status": "changed",
            "original_count": len(panels1),
            "modified_count": len(panels2)
        })
    
    # Compare panel content
    for i in range(max(len(panels1), len(panels2))):
        if i < len(panels1) and i < len(panels2):
            p1 = panels1[i]
            p2 = panels2[i]
            
            if p1.get("text") != p2.get("text"):
                changes.append({
                    "type": "panel_content",
                    "index": i + 1,
                    "status": "modified",
                    "original": p1.get("text", ""),
                    "modified": p2.get("text", "")
                })
                
            if p1.get("type") != p2.get("type"):
                changes.append({
                    "type": "panel_type",
                    "index": i + 1,
                    "status": "modified",
                    "original": p1.get("type", ""),
                    "modified": p2.get("type", "")
                })
                
        elif i < len(panels1):
            changes.append({
                "type": "panel",
                "index": i + 1,
                "status": "deleted",
                "original": panels1[i].get("text", ""),
                "modified": ""
            })
        elif i < len(panels2):
            changes.append({
                "type": "panel",
                "index": i + 1,
                "status": "added",
                "original": "",
                "modified": panels2[i].get("text", "")
            })
    
    return changes


def compare_docx_enhanced(doc1: docx.Document, doc2: docx.Document) -> Tuple[List[Dict], Dict[str, Any]]:
    """
    Enhanced DOCX comparison with all requested features.
    
    Args:
        doc1 (docx.Document): First document to compare
        doc2 (docx.Document): Second document to compare
        
    Returns:
        Tuple[List[Dict], Dict]: Changes list and statistics dictionary
    """
    # Extract content with enhanced analysis
    content1 = extract_docx_content_enhanced(doc1)
    content2 = extract_docx_content_enhanced(doc2)
    
    # Extract images
    imgs1 = get_docx_images(doc1)
    imgs2 = get_docx_images(doc2)
    table_imgs1 = get_table_images(doc1)
    table_imgs2 = get_table_images(doc2)
    
    # Detect panels
    panels1 = detect_panels(doc1)
    panels2 = detect_panels(doc2)
    
    # Compare different aspects
    text_changes = compare_text_content_enhanced(content1, content2)
    table_changes = compare_tables(content1, content2)
    image_changes = compare_images_enhanced(imgs1, imgs2)
    table_image_changes = compare_table_images_enhanced(table_imgs1, table_imgs2)
    panel_changes = compare_panels(panels1, panels2)
    
    # Combine all changes
    all_changes = text_changes + table_changes + image_changes + table_image_changes + panel_changes
    all_changes = remove_duplicate_changes(all_changes)  # Remove any remaining duplicates
    # Count elements with enhanced statistics
    para_count = len([item for item in content1 if item["type"] == "paragraph"])
    
    # Count only tables that were processed (excluding first two)
    processed_tables1 = [item for item in content1 if item["type"] == "table"]
    processed_tables2 = [item for item in content2 if item["type"] == "table"]
    table_count = len(processed_tables1)
    
    image_count = len(imgs1)
    table_image_count = len(table_imgs1)
    panel_count = len(panels1)
    
    # Count changes by type
    change_counts = {
        "paragraph": len([c for c in all_changes if c["type"] in ["paragraph", "paragraph_formatting"]]),
        "table": len([c for c in all_changes if c["type"].startswith("table")]),
        "image": len([c for c in all_changes if c["type"] in ["image", "table_image"]]),
        "panel": len([c for c in all_changes if c["type"].startswith("panel")]),
        "total": len(all_changes)
    }
    
    stats = {
        "paragraphs": para_count,
        "tables": table_count,
        "images": image_count,
        "tables_exempted": 2, 
        "table_images": table_image_count,
        "panels": panel_count,
        "total_elements": para_count + table_count + panel_count,
        "changes": change_counts,
    }
    
    return all_changes, stats


# ================================
# EXCEL EXPORT FUNCTIONS
# ================================

def export_to_excel_enhanced(changes: List[Dict], stats: Dict[str, Any]) -> io.BytesIO:
    """
    Enhanced Excel export with strict separation of strikethrough changes.
    Strikethrough changes ONLY go to "Strikethrough Changes" sheet.
    """
    wb = openpyxl.Workbook()
    
    # Remove default sheet if it exists
    if 'Sheet' in wb.sheetnames:
        del wb['Sheet']
        
    # Create Summary sheet
    summary_ws = wb.create_sheet("Summary")
    summary_ws.sheet_view.showGridLines = False
    
    summary_ws.merge_cells('A1:D1')
    title_cell = summary_ws['A1']
    title_cell.value = "ENHANCED DOCUMENT COMPARISON REPORT"
    title_cell.font = Font(bold=True, size=16)
    title_cell.alignment = Alignment(horizontal='center')

    # Add summary information
    summary_data = [
        ["Document Statistics", ""],
        ["Total Paragraphs", stats["paragraphs"]],
        ["Total Tables Processed", stats["tables"]],
        ["Tables Exempted (First 2)", stats.get("tables_exempted", 2)],
        ["Total Images", stats["images"]],
        ["Total Table Images", stats["table_images"]],
        ["Total Panels", stats["panels"]],
        ["", ""],
        ["Change Statistics", ""],
        ["Paragraph Changes", stats["changes"]["paragraph"]],
        ["Table Changes", stats["changes"]["table"]],
        ["Image Changes", stats["changes"]["image"]],
        ["Panel Changes", stats["changes"]["panel"]],
        ["Total Changes Detected", stats["changes"]["total"]],
    ]
    
    for i, (label, value) in enumerate(summary_data, start=3):
        summary_ws[f'A{i}'] = label
        summary_ws[f'B{i}'] = value
        summary_ws[f'A{i}'].font = Font(bold=True)
        
    # Set column widths for summary sheet
    summary_ws.column_dimensions['A'].width = 25
    summary_ws.column_dimensions['B'].width = 15
    
    
    # SEPARATE STRIKETHROUGH CHANGES FROM REGULAR CHANGES WITH DEDUPLICATION
    strikethrough_changes = []
    regular_changes = []
    seen_strikethrough_paragraphs = set()
    
    for change in changes:
        change_type = change.get('type', '')
        
        # Identify strikethrough changes
        if any(keyword in change_type for keyword in ['strikethrough', 'strike']):
            para_index = change.get('index')
            para_key = f"para_{para_index}" if para_index else None
            
            # Deduplicate by paragraph index
            if para_key and para_key not in seen_strikethrough_paragraphs:
                strikethrough_changes.append(change)
                seen_strikethrough_paragraphs.add(para_key)
            elif not para_key:  # For changes without paragraph index (table strikethrough)
                # Create unique key for table strikethrough changes
                table_key = f"table_{change.get('table_index')}_{change.get('row')}_{change.get('col')}"
                if table_key not in seen_strikethrough_paragraphs:
                    strikethrough_changes.append(change)
                    seen_strikethrough_paragraphs.add(table_key)
        else:
            regular_changes.append(change)
    
    print(f"DEBUG: After deduplication - Strikethrough: {len(strikethrough_changes)}, Regular: {len(regular_changes)}")
    
    # Create Strikethrough Changes sheet (ONLY for strikethrough changes)
    if strikethrough_changes:
        strike_ws = wb.create_sheet("Strikethrough Changes")
        strike_headers = [
            "Change #", "Content Type", "Location", "Status",
            "Strikethrough Text", "Context", "Change Details"
        ]
        strike_ws.append(strike_headers)
        
        # Style strikethrough sheet headers
        for col in range(1, len(strike_headers) + 1):
            cell = strike_ws.cell(row=1, column=col)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="FF6B6B", end_color="FF6B6B", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            
            
        # Process strikethrough changes
        strike_idx = 2
        seen_strike_details = set()
        
        for change in strikethrough_changes:
            # Create unique identifier for this change to avoid duplicates
            change_id = f"{change.get('type')}_{change.get('index')}_{change.get('table_index')}_{change.get('row')}_{change.get('col')}_{change.get('status')}"
            
            if change_id in seen_strike_details:
                print(f"DEBUG: Skipping duplicate strikethrough change: {change_id}")
                continue
                
            seen_strike_details.add(change_id)
            
            # Set location based on change type
            if change.get("type") in ["paragraph_strikethrough", "paragraph_formatting_strikethrough"]:
                location = f"Paragraph {change.get('index', '')}"
                content_type = "Paragraph"
            elif change.get("type") in ["table_strikethrough", "table_cell_strikethrough"]:
                location = f"Table {change.get('table_index', '')}"
                if change.get("row"):
                    location += f", Cell ({change.get('row')},{change.get('col')})"
                content_type = "Table"
            else:
                location = "Unknown"
                content_type = "Unknown"
            
            # Extract strikethrough text
            strike_text = ""
            if change.get("strikethrough_text_original"):
                strike_text += f"Original: {change.get('strikethrough_text_original')}"
            if change.get("strikethrough_text_modified"):
                if strike_text:
                    strike_text += "\n"
                strike_text += f"Modified: {change.get('strikethrough_text_modified')}"
            if change.get("strikethrough_original") and not strike_text:
                strike_text = change.get("strikethrough_original", "")
            if change.get("strikethrough_modified") and not strike_text:
                strike_text = change.get("strikethrough_modified", "")
            
            # If no specific strikethrough text, use boolean indicators
            if not strike_text.strip():
                if change.get("strikethrough_original"):
                    strike_text = "Strikethrough in original document"
                elif change.get("strikethrough_modified"):
                    strike_text = "Strikethrough in modified document"
                    
            # Context information (truncated for readability)
            context = ""
            original_text = change.get("original", "")
            modified_text = change.get("modified", "")
            
            if original_text:
                context += f"Original: {original_text[:100]}{'...' if len(original_text) > 100 else ''}"
            if modified_text:
                if context:
                    context += "\n"
                context += f"Modified: {modified_text[:100]}{'...' if len(modified_text) > 100 else ''}"
            
            # Change details (clean up formatting changes)
            change_detail = change.get("change_detail", "")
            if "Strikethrough formatting changes:" in change_detail:
                # Remove duplicate entries from formatting changes
                lines = change_detail.split('\n')
                unique_lines = []
                seen_lines = set()
                
                for line in lines:
                    if line and line not in seen_lines:
                        unique_lines.append(line)
                        seen_lines.add(line)
                
                change_detail = '\n'.join(unique_lines)
            
            strike_ws.append([
                strike_idx-1,
                content_type,
                location,
                change.get("status", "").replace("_", " ").title(),
                strike_text,
                context,
                change_detail
            ])
            
            # Apply strikethrough formatting to the text cells
            text_cell = strike_ws.cell(row=strike_idx, column=5)  # Strikethrough Text column
            text_cell.font = Font(strike=True)
            
            # Color code based on status
            status = change.get("status", "").lower()
            if "original" in status or "deleted" in status:
                fill_color = PatternFill(start_color="FFCCCB", end_color="FFCCCB", fill_type="solid")
            elif "modified" in status or "added" in status:
                fill_color = PatternFill(start_color="FFFFE0", end_color="FFFFE0", fill_type="solid")
            else:
                fill_color = PatternFill(start_color="E6E6FA", end_color="E6E6FA", fill_type="solid")
            
            for col in range(1, len(strike_headers) + 1):
                strike_ws.cell(row=strike_idx, column=col).fill = fill_color
            
            strike_idx += 1
        
        # Set column widths for strikethrough sheet
        strike_widths = [8, 15, 20, 20, 40, 50, 40]
        for i, width in enumerate(strike_widths, start=1):
            strike_ws.column_dimensions[get_column_letter(i)].width = width

    # Create Detailed Changes sheet (ONLY for regular changes - NO STRIKETHROUGH)
    changes_ws = wb.create_sheet("Detailed Changes")
    headers = [
        "Change #", "Type", "Location", "Status",
        "Original Content", "Modified Content",
        "Change Details"
    ]
    changes_ws.append(headers)
    
    # Style main changes sheet headers
    for col in range(1, len(headers) + 1):
        cell = changes_ws.cell(row=1, column=col)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
    
    # Define color fills for different change types (NO STRIKETHROUGH COLORS)
    red_fill = PatternFill(start_color="FFCCCB", end_color="FFCCCB", fill_type="solid")  # Deleted
    yellow_fill = PatternFill(start_color="FFFFE0", end_color="FFFFE0", fill_type="solid")  # Added
    green_fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")  # Modified
    
    # Process regular changes (NO STRIKETHROUGH CHANGES)
    change_idx = 2
    for change in regular_changes:
        # Set location based on change type
        if change["type"] == "paragraph":
            location = f"Paragraph {change.get('index', '')}"
        elif change["type"] == "table":
            location = f"Table {change.get('index', '')}"
        elif change["type"] == "table_cell":
            location = f"Table {change.get('table_index', '')}, Cell ({change.get('row','')},{change.get('col','')})"
        elif change["type"] == "table_row":
            location = f"Table {change.get('table_index', '')}, Row {change.get('row','')}"
        elif change["type"] == "image":
            location = "Document"
        elif change["type"] == "table_image":
            location = f"Table {change.get('table','')}, Cell ({change.get('row','')},{change.get('col','')})"
        elif change["type"] == "panel":
            location = f"Panel {change.get('index', '')}"
        else:
            location = "Unknown"
        
        # Handle images
        if change["type"] in ["image", "table_image"]:
            if change["status"] == "deleted" and change.get("old_img"):
                try:
                    img = XLImage(io.BytesIO(change["old_img"]))
                    img.width, img.height = 40, 40  # thumbnail size
                    img.anchor = f'E{change_idx}'  # Original Content column
                    changes_ws.add_image(img)
                except Exception as e:
                    print(f"Error adding deleted image to Excel: {e}")
            elif change["status"] == "added" and change.get("new_img"):
                try:
                    img = XLImage(io.BytesIO(change["new_img"]))
                    img.width, img.height = 40, 40
                    img.anchor = f'F{change_idx}'  # Modified Content column
                    changes_ws.add_image(img)
                except Exception as e:
                    print(f"Error adding added image to Excel: {e}")

        # Determine fill color based on change type (NO STRIKETHROUGH LOGIC)
        fill = None
        if change["status"] == "deleted":
            fill = red_fill
        elif change["status"] == "added":
            fill = yellow_fill
        elif change["status"] in ["modified", "formatting_changed"]:
            fill = green_fill

        # Build change details (NO STRIKETHROUGH INFORMATION)
        change_details = ""
        if change["type"] == "paragraph":
            if change.get("double_spaces_original", 0) > 0 or change.get("double_spaces_modified", 0) > 0:
                change_details += f"Double spaces: {change.get('double_spaces_original', 0)} → {change.get('double_spaces_modified', 0)}\n"
            if change.get("tbd_original") or change.get("tbd_modified"):
                change_details += f"TBD content: {change.get('tbd_original')} → {change.get('tbd_modified')}\n"
            if change.get("code_names_original") or change.get("code_names_modified"):
                change_details += f"Code names: {change.get('code_names_original')} → {change.get('code_names_modified')}\n"
            if change.get("urls_original") or change.get("urls_modified"):
                change_details += f"URLs: {change.get('urls_original')} → {change.get('urls_modified')}\n"
                
        elif change["type"] == "paragraph_formatting":
            for fmt_change in change.get("formatting_changes", []):
                change_details += f"'{fmt_change['text']}': {', '.join(fmt_change['changes'])}\n"
                
        elif change["type"] in ["image", "table_image"]:
            details = []
            details.append(change.get('change_detail', ''))
            if change["status"] == "deleted":
                details.append(f"Image deleted. Size: {change.get('width','?')}x{change.get('height','?')}")
            elif change["status"] == "added":
                details.append(f"Image added. Size: {change.get('width','?')}x{change.get('height','?')}")
            change_details = "\n".join([d for d in details if d])
        
        # Handle image text content
        original_text = change.get("original", "")
        modified_text = change.get("modified", "")
        if change["type"] in ["image", "table_image"]:
            if change["status"] == "deleted":
                original_text = f"Image: {change.get('width','?')}x{change.get('height','?')}, Format: {change.get('format','?')}"
                modified_text = ""
            elif change["status"] == "added":  
                original_text = ""
                modified_text = f"Image: {change.get('width','?')}x{change.get('height','?')}, Format: {change.get('format','?')}"
        
        # Write row to changes sheet (NO STRIKETHROUGH INFORMATION)
        changes_ws.cell(row=change_idx, column=1, value=change_idx-1)  # Change #
        changes_ws.cell(row=change_idx, column=2, value=change["type"].replace("_", " ").title())
        changes_ws.cell(row=change_idx, column=3, value=location)
        changes_ws.cell(row=change_idx, column=4, value=change["status"].replace("_", " ").title())
        changes_ws.cell(row=change_idx, column=5, value=original_text)
        changes_ws.cell(row=change_idx, column=6, value=modified_text)
        changes_ws.cell(row=change_idx, column=7, value=change_details)
            
        # Apply fill to all cells in the row (NO STRIKETHROUGH FORMATTING)
        for col in range(1, 9):
            cell = changes_ws.cell(row=change_idx, column=col)
            if fill:
                cell.fill = fill
            # Add borders
            cell.border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
        
        change_idx += 1

    # Set column widths for changes sheet
    column_widths = [8, 15, 20, 12, 40, 40, 40, 30]
    for i, width in enumerate(column_widths, start=1):
        changes_ws.column_dimensions[get_column_letter(i)].width = width

    # Set active sheet to Detailed Changes
    wb.active = changes_ws
        
    # Save to bytes buffer
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


def remove_duplicate_changes(changes: List[Dict]) -> List[Dict]:
    """
    Remove duplicate changes based on unique identifiers.
    
    Args:
        changes (List): List of changes potentially with duplicates
        
    Returns:
        List[Dict]: List of unique changes
    """
    seen_changes = set()
    unique_changes = []
    
    for change in changes:
        # Create a unique identifier for each change
        if change["type"] == "table_cell":
            change_id = f"table_{change.get('table_index')}_cell_{change.get('row')}_{change.get('col')}"
        elif change["type"] == "paragraph":
            change_id = f"para_{change.get('index')}"
        elif change["type"] == "image":
            change_id = f"image_{change.get('hash', '')}"
        elif change["type"] == "table_image":
            change_id = f"table_image_{change.get('table')}_{change.get('row')}_{change.get('col')}_{change.get('hash', '')}"
        else:
            change_id = f"{change['type']}_{hash(str(change))}"
        
        if change_id not in seen_changes:
            seen_changes.add(change_id)
            unique_changes.append(change)
    
    print(f"DEBUG: Removed {len(changes) - len(unique_changes)} duplicate changes")
    return unique_changes

def validate_strikethrough_detection(doc: docx.Document):
    """
    Validate strikethrough detection for debugging purposes.
    
    Args:
        doc (docx.Document): Document to validate
    """
    print("=== STRIKETHROUGH VALIDATION ===")
    
    for para_idx, para in enumerate(doc.paragraphs[:20], start=1):  # Check first 20 paragraphs
        if para.text.strip():
            print(f"Paragraph {para_idx}: '{para.text[:50]}...'")
            
            for run_idx, run in enumerate(para.runs):
                if run.text.strip():
                    strike = detect_strikethrough(run)
                    if strike:
                        print(f"  Run {run_idx}: STRIKETHROUGH FOUND - '{run.text}'")
                    else:
                        print(f"  Run {run_idx}: No strikethrough - '{run.text}'")
    
    print("=== END VALIDATION ===")


# ================================
# FLASK ROUTES
# ================================

@app.route("/", methods=["GET", "POST"])
def index():
    """
    Main route for document comparison interface.
    Handles both GET requests (page display) and POST requests (document processing).
    """
    if request.method == "POST":
        is_download_request = request.form.get('download') == 'true'
        
        file1 = request.files.get("doc1")
        file2 = request.files.get("doc2")
        pass1 = request.form.get("password1", "").strip() or None
        pass2 = request.form.get("password2", "").strip() or None

        if not file1 or not file2:
            error_msg = "Please upload both documents."
            if is_download_request or request.headers.get('X-Requested-With') == 'XMLHttpRequest':
                return jsonify({"error": error_msg}), 400
            flash(f"⚠️ {error_msg}")
            return render_template("index.html")
            
        try:
            # Reset stream positions
            file1.stream.seek(0)
            file2.stream.seek(0)
            
            # Check if documents need passwords
            needs_pass1 = is_encrypted(file1.stream)
            needs_pass2 = is_encrypted(file2.stream)
            
            # Reset streams again after encryption check
            file1.stream.seek(0)
            file2.stream.seek(0)
            
            if needs_pass1 and not pass1:
                error_msg = "Document 1 is encrypted but no password was provided."
                if is_download_request or request.headers.get('X-Requested-With') == 'XMLHttpRequest':
                    return jsonify({"error": error_msg}), 400
                flash(f"⚠️ {error_msg}")
                return render_template("index.html")
                
            if needs_pass2 and not pass2:
                error_msg = "Document 2 is encrypted but no password was provided."
                if is_download_request or request.headers.get('X-Requested-With') == 'XMLHttpRequest':
                    return jsonify({"error": error_msg}), 400
                flash(f"⚠️ {error_msg}")
                return render_template("index.html")

            # Load documents
            doc1 = load_docx(file1.stream, pass1)
            doc2 = load_docx(file2.stream, pass2)

            # Compare documents with enhanced comparison
            changes, stats = compare_docx_enhanced(doc1, doc2)

            # For download requests, return the Excel file
            if is_download_request:
                excel_file = export_to_excel_enhanced(changes, stats)
                
                # Generate filename with timestamp
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                filename = f"document_comparison_{timestamp}.xlsx"
                return send_file(
                    excel_file,
                    as_attachment=True,
                    download_name=filename,
                    mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            # For AJAX requests, return JSON
            if request.headers.get('X-Requested-With') == 'XMLHttpRequest':
                # Convert changes to serializable format
                serializable_changes = []
                for change in changes:
                    serializable_change = change.copy()
                    # Remove image bytes as they can't be serialized to JSON
                    serializable_change.pop('old_img', None)
                    serializable_change.pop('new_img', None)
                    serializable_changes.append(serializable_change)
                
                return jsonify({
                    "changes": serializable_changes,
                    "stats": stats
                })
            else:
                # For regular form submission, show results on page
                return render_template("index.html", changes=changes, stats=stats)

        except Exception as e:
            error_msg = f"Error processing documents: {str(e)}"
            if is_download_request or request.headers.get('X-Requested-With') == 'XMLHttpRequest':
                return jsonify({"error": error_msg}), 400
            else:
                flash(f"❌ {error_msg}")
                return render_template("index.html")

    # Check if download was completed and refresh the page
    if request.cookies.get('download_complete') == 'true':
        response = render_template("index.html")
        response.set_cookie('download_complete', '', expires=0)
        return response

    return render_template("index.html")


# ================================
# APPLICATION ENTRY POINT
# ================================

if __name__ == "__main__":
    """
    Main entry point for the Flask application.
    """
    print("Starting Enhanced DOCX Comparison Tool...")
    print("Access the application at: http://localhost:5000")
    app.run(debug=True)