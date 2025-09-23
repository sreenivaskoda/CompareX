import os
import io
import tempfile
import difflib
import re
from hashlib import md5
from collections import defaultdict
from typing import List, Dict, Any, Tuple
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
from openpyxl.drawing.image import Image as XLImage

app = Flask(__name__)
app.secret_key = "supersecretkey"
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# ----------------- Utility Functions -----------------
def qn(tag):
    """Utility function to handle XML namespaces."""
    if tag.startswith('{'):
        return tag
    return '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}' + tag

def is_encrypted(file_stream):
    """Check if a document is encrypted without fully loading it."""
    try:
        file_stream.seek(0)
        office_file = msoffcrypto.OfficeFile(file_stream)
        return office_file.is_encrypted()
    except Exception:
        return False
    finally:
        file_stream.seek(0)

def load_docx(file_stream, password=None):
    """Load DOCX file, handling both encrypted and unencrypted documents."""
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

def extract_docx_content(doc):
    """Extract all content from DOCX in order (paragraphs, tables)."""
    content = []
    
    # Extract paragraphs
    for para in doc.paragraphs:
        if para.text.strip():
            content.append({
                "type": "paragraph",
                "text": para.text,
                "runs": [run.text for run in para.runs if run.text.strip()]
            })
    
    # Extract tables
    for table_idx, table in enumerate(doc.tables, start=1):
        table_data = {
            "type": "table",
            "index": table_idx,
            "rows": []
        }
        
        for row_idx, row in enumerate(table.rows, start=1):
            row_data = {"cells": []}
            for cell_idx, cell in enumerate(row.cells, start=1):
                cell_text = "\n".join(p.text for p in cell.paragraphs if p.text.strip())
                row_data["cells"].append({
                    "text": cell_text,
                    "row": row_idx,
                    "col": cell_idx
                })
            table_data["rows"].append(row_data)
        
        content.append(table_data)
    
    return content

def extract_docx_content_enhanced(doc):
    """Extract all content from DOCX with enhanced text analysis including color detection"""
    content = []
    
    # Extract paragraphs with detailed run information including colors
    for para_idx, para in enumerate(doc.paragraphs, start=1):
        if para.text.strip():
            runs_data = []
            colored_text_found = False
            
            for run_idx, run in enumerate(para.runs):
                if run.text.strip():
                    # Detect text color
                    text_color = detect_text_color(run)
                    
                    run_data = {
                        "text": run.text,
                        "strikethrough": detect_strikethrough(run),
                        "bold": run.font.bold,
                        "italic": run.font.italic,
                        "underline": run.font.underline,
                        "font_name": run.font.name,
                        "font_size": run.font.size,
                        "font_color": run.font.color.rgb if run.font.color else None,
                        "detected_color": text_color  # This is the key addition
                    }
                    
                    if text_color:
                        colored_text_found = True
                    
                    runs_data.append(run_data)
            
            content.append({
                "type": "paragraph",
                "index": para_idx,
                "text": para.text,
                "runs": runs_data,
                "style": para.style.name if para.style else "Normal",
                "has_colored_text": colored_text_found
            })
    
    # Extract tables with content, structure, and color detection
    for table_idx, table in enumerate(doc.tables, start=1):
        table_data = {
            "type": "table",
            "index": table_idx,
            "rows": []
        }
        
        for row_idx, row in enumerate(table.rows, start=1):
            row_data = {"cells": []}
            for cell_idx, cell in enumerate(row.cells, start=1):
                cell_text = "\n".join(p.text for p in cell.paragraphs if p.text.strip())
                
                # Detect cell background color
                cell_bg_color = detect_cell_color(cell)
                
                # Check for colored text in cell paragraphs
                cell_has_colored_text = False
                for para in cell.paragraphs:
                    for run in para.runs:
                        if run.text.strip() and detect_text_color(run):
                            cell_has_colored_text = True
                            break
                    if cell_has_colored_text:
                        break
                
                # Check for TBD content, code names, and URLs
                contains_tbd = "TBD" in cell_text.upper()
                contains_code_names = bool(re.search(r'\[[A-Z]+\]', cell_text))
                urls = re.findall(r'https?://[^\s<>"]+|www\.[^\s<>"]+', cell_text)
                
                row_data["cells"].append({
                    "text": cell_text,
                    "row": row_idx,
                    "col": cell_idx,
                    "contains_tbd": contains_tbd,
                    "contains_code_names": contains_code_names,
                    "urls": urls,
                    "cell_bg_color": cell_bg_color,
                    "has_colored_text": cell_has_colored_text
                })
            table_data["rows"].append(row_data)
        
        content.append(table_data)
    
    return content

def detect_strikethrough(run):
    """Detect if text has strikethrough formatting."""
    if run.font.strike:
        return True
    # Check XML directly for more reliable strikethrough detection
    if hasattr(run, '_element'):
        strike_elements = run._element.xpath('.//w:strike')
        if strike_elements and strike_elements[0].get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') != 'false':
            return True
    return False
def detect_text_color(run):
    """Detect if text has specific colors (Red, Yellow, Green)"""
    if run.font.color and run.font.color.rgb:
        color_rgb = run.font.color.rgb
        
        # Define the specific RGB values for Red, Yellow, Green
        COLOR_MAPPING = {
            'RED': RGBColor(255, 0, 0),        # Pure Red
            'YELLOW': RGBColor(255, 255, 0),   # Pure Yellow
            'GREEN': RGBColor(0, 255, 0),      # Pure Green
        }
        
        # Check for exact color matches
        for color_name, rgb_value in COLOR_MAPPING.items():
            if (color_rgb == rgb_value or 
                (hasattr(color_rgb, 'rgb') and color_rgb.rgb == rgb_value)):
                return color_name
        
        # Check for similar colors with tolerance
        if hasattr(color_rgb, 'rgb'):
            r, g, b = color_rgb.rgb
        else:
            # Handle different color representations
            try:
                if isinstance(color_rgb, int):
                    r = (color_rgb >> 16) & 0xFF
                    g = (color_rgb >> 8) & 0xFF
                    b = color_rgb & 0xFF
                else:
                    return None
            except:
                return None
        
        # Color matching with tolerance
        COLOR_RANGES = {
            'RED': ((200, 255), (0, 100), (0, 100)),      # High R, Low G, Low B
            'YELLOW': ((200, 255), (200, 255), (0, 100)), # High R, High G, Low B
            'GREEN': ((0, 100), (200, 255), (0, 100)),    # Low R, High G, Low B
        }
        
        for color_name, (r_range, g_range, b_range) in COLOR_RANGES.items():
            if (r_range[0] <= r <= r_range[1] and 
                g_range[0] <= g <= g_range[1] and 
                b_range[0] <= b <= b_range[1]):
                return color_name
    
    return None

def detect_cell_color(cell):
    """Detect cell background color in tables"""
    try:
        # Check cell shading
        if hasattr(cell, '_tc') and hasattr(cell._tc, 'get_or_add_tcPr'):
            tcPr = cell._tc.get_or_add_tcPr()
            shading = tcPr.find(qn('w:shd'))
            if shading is not None:
                fill_color = shading.get(qn('w:fill'))
                if fill_color:
                    # Convert hex to RGB
                    if fill_color.startswith('#'):
                        hex_color = fill_color.lstrip('#')
                        r = int(hex_color[0:2], 16)
                        g = int(hex_color[2:4], 16)
                        b = int(hex_color[4:6], 16)
                        
                        # Match against our color ranges
                        COLOR_RANGES = {
                            'RED': ((200, 255), (0, 100), (0, 100)),
                            'YELLOW': ((200, 255), (200, 255), (0, 100)),
                            'GREEN': ((0, 100), (200, 255), (0, 100)),
                        }
                        
                        for color_name, (r_range, g_range, b_range) in COLOR_RANGES.items():
                            if (r_range[0] <= r <= r_range[1] and 
                                g_range[0] <= g <= g_range[1] and 
                                b_range[0] <= b <= b_range[1]):
                                return color_name
    except Exception as e:
        print(f"Error detecting cell color: {e}")
    
    return None

def get_docx_images(doc):
    """Extract all images from DOCX document with comprehensive detection."""
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
                        
                        # Get image dimensions
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
                        print(f"DEBUG: Found image via relationships - Hash: {img_hash}, Size: {width}x{height}")
                    except Exception as e:
                        print(f"DEBUG: Error processing relationship image {rel_id}: {e}")
                        continue
        
        # Method 2: Check inline shapes (alternative image storage)
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
                                print(f"DEBUG: Found inline image - Hash: {img_hash}, Size: {width}x{height}")
                            except Exception as e:
                                print(f"DEBUG: Error processing inline image {embed_id}: {e}")
                                continue
        except Exception as e:
            print(f"DEBUG: Error processing inline shapes: {e}")
        
        print(f"DEBUG: Total images extracted: {len(images)}")
        
    except Exception as e:
        print(f"DEBUG: Error in get_docx_images: {e}")
    
    return images

def get_table_images(doc):
    """Extract images from inside table cells with proper XML parsing."""
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

def insert_image_into_excel(ws, img_bytes, cell_address, max_width=120, max_height=120):
    """Properly insert an image into Excel worksheet with size adjustment."""
    if not img_bytes:
        return False
        
    try:
        # Create temporary file for the image
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_file:
            # Open and process the image
            pil_img = PILImage.open(io.BytesIO(img_bytes))
            
            # Convert to RGB if necessary (for PNG transparency issues)
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
            img.anchor = cell_address  # e.g., 'A1'
            
            # Adjust cell size to fit image
            col_letter = cell_address[0]  # Get column letter from cell address
            row_num = int(cell_address[1:])  # Get row number
            
            # Set column width based on image width (approximate)
            col_width = max(8.43, min(50, img.width * 0.14))  # Convert pixels to Excel column width
            ws.column_dimensions[col_letter].width = col_width
            
            # Set row height based on image height (approximate)
            row_height = max(15, min(400, img.height * 0.75))  # Convert pixels to Excel row height
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
    
def compare_images_enhanced(imgs1, imgs2):
    """Enhanced image comparison that properly captures deleted images."""
    changes = []
    
    # Group images by hash for comparison
    hashes1 = {img["hash"]: img for img in imgs1}
    hashes2 = {img["hash"]: img for img in imgs2}
    
    print(f"DEBUG: Document 1 has {len(imgs1)} images, Document 2 has {len(imgs2)} images")
    print(f"DEBUG: Unique hashes in doc1: {len(hashes1)}, doc2: {len(hashes2)}")
    
    # Find deleted images (in doc1 but not in doc2)
    for img_hash, img1 in hashes1.items():
        if img_hash not in hashes2:
            print(f"DEBUG: Deleted image detected - Hash: {img_hash}")
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
            print(f"DEBUG: Added image detected - Hash: {img_hash}")
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
            print(f"DEBUG: Modified image detected - Hash: {img_hash}")
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

def compare_table_images_enhanced(imgs1, imgs2):
    """Enhanced table image comparison that properly captures deleted images."""
    changes = []
    
    print(f"DEBUG: Table images - Doc1: {len(imgs1)}, Doc2: {len(imgs2)}")
    
    # Group by table, row, column AND hash to track individual images
    grouped1 = defaultdict(list)
    grouped2 = defaultdict(list)
    
    for img in imgs1:
        key = (img["table"], img["row"], img["col"], img["hash"])
        grouped1[key].append(img)
        
    for img in imgs2:
        key = (img["table"], img["row"], img["col"], img["hash"])
        grouped2[key].append(img)
    
    # Get all unique image identifiers
    all_images1 = {(img["table"], img["row"], img["col"], img["hash"]): img for img in imgs1}
    all_images2 = {(img["table"], img["row"], img["col"], img["hash"]): img for img in imgs2}
    
    # Find deleted table images
    for img_key, img1 in all_images1.items():
        if img_key not in all_images2:
            table, row, col, img_hash = img_key
            print(f"DEBUG: Deleted table image detected - Table {table}, Cell ({row},{col}), Hash: {img_hash}")
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
            print(f"DEBUG: Added table image detected - Table {table}, Cell ({row},{col}), Hash: {img_hash}")
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
                    
                    print(f"DEBUG: Moved table image detected - Hash: {img_hash}")
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

def compare_text_content_enhanced(content1, content2):
    """Enhanced text comparison with detailed change tracking."""
    changes = []
    
    # Compare paragraphs
    para1 = [item for item in content1 if item["type"] == "paragraph"]
    para2 = [item for item in content2 if item["type"] == "paragraph"]
    
    for i in range(max(len(para1), len(para2))):
        if i < len(para1) and i < len(para2):
            p1 = para1[i]
            p2 = para2[i]
            
            # Check for colored text in both versions
            p1_colored_runs = [run for run in p1["runs"] if run.get("detected_color")]
            p2_colored_runs = [run for run in p2["runs"] if run.get("detected_color")]
            
            # Process colored runs from original document (p1)
            for run in p1_colored_runs:
                color = run["detected_color"]
                formatting_changes = map_color_to_formatting_changes(color, "original")
                if formatting_changes:
                    changes.append({
                        "type": "colored_content",
                        "content_type": "paragraph",
                        "index": i + 1,
                        "status": formatting_changes,
                        "color": color,
                        "text": run["text"],
                        "original_context": p1["text"],
                        "modified_context": p2["text"] if i < len(para2) else "",
                        "change_detail": f"Color-coded text detected: {color} - {run['text']}",
                        "is_color_based": True

                    })

            # Process colored runs from modified document (p2)
            for run in p2_colored_runs:
                color = run["detected_color"]
                formatting_changes = map_color_to_formatting_changes(color, "modified")
                if formatting_changes and not any(c.get("text") == run["text"] and c.get("color") == color 

                    for c in changes if c.get("is_color_based")):
                        changes.append({
                        "type": "colored_content",
                        "content_type": "paragraph",
                        "index": i + 1,
                        "status": formatting_changes,
                        "color": color,
                        "text": run["text"],
                        "original_context": p1["text"] if i < len(para1) else "",
                        "modified_context": p2["text"],
                        "change_detail": f"Color-coded text detected: {color} - {run['text']}",
                        "is_color_based": True
                    })
            # Check for text changes
            if p1["text"] != p2["text"]:
                # Use difflib to show exact differences
                diff = list(difflib.ndiff(p1["text"].split(), p2["text"].split()))
                added = [word[2:] for word in diff if word.startswith('+ ')]
                removed = [word[2:] for word in diff if word.startswith('- ')]
                
                # Check for double spaces
                double_spaces_original = len(re.findall(r'  ', p1["text"]))
                double_spaces_modified = len(re.findall(r'  ', p2["text"]))
                
                # Check for TBD content
                tbd_original = "TBD" in p1["text"].upper()
                tbd_modified = "TBD" in p2["text"].upper()
                
                # Check for code names
                code_names_original = bool(re.search(r'\[[A-Z]+\]', p1["text"]))
                code_names_modified = bool(re.search(r'\[[A-Z]+\]', p2["text"]))
                
                # Check for URL changes
                urls_original = re.findall(r'https?://[^\s<>"]+|www\.[^\s<>"]+', p1["text"])
                urls_modified = re.findall(r'https?://[^\s<>"]+|www\.[^\s<>"]+', p2["text"])
                
                # Check for strikethrough text
                strikethrough_original = any(run.get("strikethrough", False) for run in p1["runs"])
                strikethrough_modified = any(run.get("strikethrough", False) for run in p2["runs"])
                
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
                    "strikethrough_original": strikethrough_original,
                    "strikethrough_modified": strikethrough_modified,
                    "is_color_based": False
                })
            # Check for formatting changes even if text is the same
            elif p1["text"] == p2["text"]:
                # Compare runs for formatting changes
                formatting_changes = []
                for run_idx in range(max(len(p1["runs"]), len(p2["runs"]))):
                    if run_idx < len(p1["runs"]) and run_idx < len(p2["runs"]):
                        run1 = p1["runs"][run_idx]
                        run2 = p2["runs"][run_idx]
                        
                        if run1["text"] == run2["text"]:
                            # Check for formatting differences
                            format_diffs = []
                            if run1.get("bold") != run2.get("bold"):
                                format_diffs.append("bold")
                            if run1.get("italic") != run2.get("italic"):
                                format_diffs.append("italic")
                            if run1.get("underline") != run2.get("underline"):
                                format_diffs.append("underline")
                            if run1.get("strikethrough") != run2.get("strikethrough"):
                                format_diffs.append("strikethrough")
                            if run1.get("font_name") != run2.get("font_name"):
                                format_diffs.append("font")
                            if run1.get("font_size") != run2.get("font_size"):
                                format_diffs.append("font size")
                            if run1.get("font_color") != run2.get("font_color"):
                                format_diffs.append("font color")
                                
                            if format_diffs:
                                formatting_changes.append({
                                    "text": run1["text"],
                                    "changes": format_diffs
                                })
                
                if formatting_changes:
                    changes.append({
                        "type": "paragraph_formatting",
                        "index": i + 1,
                        "status": "formatting_changed",
                        "original": p1["text"],
                        "modified": p2["text"],
                        "formatting_changes": formatting_changes,
                        "is_color_based": False
                    })    
        elif i < len(para1):  # Deleted paragraph
            p1 = para1[i]
            # Check for colored text in deleted paragraph
            colored_runs = [run for run in p1["runs"] if run.get("detected_color")]
            for run in colored_runs:
                color = run["detected_color"]
                change_type = map_color_to_change_type(color, "original")
                if change_type:
                    changes.append({
                        "type": "colored_content",
                        "content_type": "paragraph",
                        "index": i + 1,
                        "status": change_type,
                        "color": color,
                        "text": run["text"],
                        "original_context": p1["text"],
                        "modified_context": "",
                        "change_detail": f"Color-coded text in deleted content: {color} - {run['text']}",
                        "is_color_based": True
                    })
            
            changes.append({
                "type": "paragraph",
                "index": i + 1,
                "status": "deleted",
                "original": p1["text"],
                "modified": "",
                "added": "",
                "removed": p1["text"],
                "is_color_based": False
            })
            
        elif i < len(para2):  # Added paragraph
            p2 = para2[i]
            # Check for colored text in added paragraph
            colored_runs = [run for run in p2["runs"] if run.get("detected_color")]
            for run in colored_runs:
                color = run["detected_color"]
                change_type = map_color_to_change_type(color, "modified")
                if change_type:
                    changes.append({
                        "type": "colored_content",
                        "content_type": "paragraph",
                        "index": i + 1,
                        "status": change_type,
                        "color": color,
                        "text": run["text"],
                        "original_context": "",
                        "modified_context": p2["text"],
                        "change_detail": f"Color-coded text in added content: {color} - {run['text']}",
                        "is_color_based": True
                    })
            
            changes.append({
                "type": "paragraph",
                "index": i + 1,
                "status": "added",
                "original": "",
                "modified": p2["text"],
                "added": p2["text"],
                "removed": "",
                "is_color_based": False
            })
    
    return changes
def map_color_to_formatting_changes(color, document_type):
    """Map detected colors to change types based on your requirements"""
    COLOR_MAPPING = {
        'RED': 'removed' if document_type == 'original' else 'review_required',
        'YELLOW': 'new_addition' if document_type == 'modified' else 'existing_addition',
        'GREEN': 'modified' if document_type == 'modified' else 'original_version'
    }
    return COLOR_MAPPING.get(color)

def compare_tables(content1, content2):
    """Compare tables between two documents."""
    changes = []
    tables1 = [item for item in content1 if item["type"] == "table"]
    tables2 = [item for item in content2 if item["type"] == "table"]
    
    for i in range(max(len(tables1), len(tables2))):
        if i < len(tables1) and i < len(tables2):
            table1 = tables1[i]
            table2 = tables2[i]
            
            # Compare row count
            if len(table1["rows"]) != len(table2["rows"]):
                changes.append({
                    "type": "table",
                    "index": i + 1,
                    "status": "structure_changed",
                    "detail": f"Row count changed from {len(table1['rows'])} to {len(table2['rows'])}",
                    "original": f"Table with {len(table1['rows'])} rows",
                    "modified": f"Table with {len(table2['rows'])} rows"
                })
            
            # Compare cell content
            for r_idx in range(max(len(table1["rows"]), len(table2["rows"]))):
                if r_idx < len(table1["rows"]) and r_idx < len(table2["rows"]):
                    row1 = table1["rows"][r_idx]
                    row2 = table2["rows"][r_idx]
                    
                    # Compare column count
                    if len(row1["cells"]) != len(row2["cells"]):
                        changes.append({
                            "type": "table",
                            "index": i + 1,
                            "status": "structure_changed",
                            "detail": f"Column count changed in row {r_idx+1} from {len(row1['cells'])} to {len(row2['cells'])}",
                            "original": f"Row with {len(row1['cells'])} columns",
                            "modified": f"Row with {len(row2['cells'])} columns"
                        })
                    
                    # Compare cell content
                    for c_idx in range(max(len(row1["cells"]), len(row2["cells"]))):
                        if c_idx < len(row1["cells"]) and c_idx < len(row2["cells"]):
                            cell1 = row1["cells"][c_idx]
                            cell2 = row2["cells"][c_idx]
                            
                            if cell1["text"] != cell2["text"]:
                                changes.append({
                                    "type": "table_cell",
                                    "table_index": i + 1,
                                    "row": r_idx + 1,
                                    "col": c_idx + 1,
                                    "status": "modified",
                                    "original": cell1["text"],
                                    "modified": cell2["text"]
                                })
                        elif c_idx < len(row1["cells"]):
                            changes.append({
                                "type": "table_cell",
                                "table_index": i + 1,
                                "row": r_idx + 1,
                                "col": c_idx + 1,
                                "status": "deleted",
                                "original": row1["cells"][c_idx]["text"],
                                "modified": ""
                            })
                        elif c_idx < len(row2["cells"]):
                            changes.append({
                                "type": "table_cell",
                                "table_index": i + 1,
                                "row": r_idx + 1,
                                "col": c_idx + 1,
                                "status": "added",
                                "original": "",
                                "modified": row2["cells"][c_idx]["text"]
                            })
                elif r_idx < len(table1["rows"]):
                    changes.append({
                        "type": "table_row",
                        "table_index": i + 1,
                        "row": r_idx + 1,
                        "status": "deleted",
                        "original": f"Row {r_idx+1} with {len(table1['rows'][r_idx]['cells'])} columns",
                        "modified": ""
                    })
                elif r_idx < len(table2["rows"]):
                    changes.append({
                        "type": "table_row",
                        "table_index": i + 1,
                        "row": r_idx + 1,
                        "status": "added",
                        "original": "",
                        "modified": f"Row {r_idx+1} with {len(table2['rows'][r_idx]['cells'])} columns"
                    })
        elif i < len(tables1):
            changes.append({
                "type": "table",
                "index": i + 1,
                "status": "deleted",
                "original": f"Table with {len(tables1[i]['rows'])} rows",
                "modified": ""
            })
        elif i < len(tables2):
            changes.append({
                "type": "table",
                "index": i + 1,
                "status": "added",
                "original": "",
                "modified": f"Table with {len(tables2[i]['rows'])} rows"
            })
    
    return changes

def detect_panels(doc):
    """Detect panels in document based on section breaks, headings, or other markers."""
    panels = []
    
    # Look for section breaks
    for sect in doc.sections:
        panels.append({
            "type": "section",
            "start": "beginning",  # This would need more sophisticated detection
            "content": "Section content"  # Placeholder
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

def compare_panels(panels1, panels2):
    """Compare panels between documents."""
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

def compare_docx_enhanced(doc1, doc2):
    """Enhanced DOCX comparison with all requested features."""
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
    
    # Separate color-based changes for better organization
    color_based_changes = [change for change in text_changes if change.get("is_color_based")]
    other_changes = [change for change in text_changes if not change.get("is_color_based")]
    
    # Combine all changes
    all_changes = text_changes + table_changes + image_changes + table_image_changes + panel_changes + color_based_changes + other_changes
    
    # Count elements with enhanced statistics
    para_count = len([item for item in content1 if item["type"] == "paragraph"])
    table_count = len([item for item in content1 if item["type"] == "table"])
    image_count = len(imgs1)
    table_image_count = len(table_imgs1)
    panel_count = len(panels1)
    # Update stats to include color-based changes
    color_count = len(color_based_changes)
    
    # Count changes by type
    change_counts = {
        "paragraph": len([c for c in all_changes if c["type"] in ["paragraph", "paragraph_formatting"]]),
        "table": len([c for c in all_changes if c["type"].startswith("table")]),
        "image": len([c for c in all_changes if c["type"] in ["image", "table_image"]]),
        "panel": len([c for c in all_changes if c["type"].startswith("panel")]),
        "color_based": color_count,
        "total": len(all_changes)
    }
    
    stats = {
        "paragraphs": para_count,
        "color_count" : color_count,
        "tables": table_count,
        "images": image_count,
        "table_images": table_image_count,
        "panels": panel_count,
        "total_elements": para_count + table_count + panel_count,
        "changes": change_counts,
    }
    
    return all_changes, stats


def export_to_excel_enhanced(changes, stats):
    """Enhanced Excel export with color-coding and better organization."""
    wb = openpyxl.Workbook()
    
    # Remove default sheet if it exists
    if 'Sheet' in wb.sheetnames:
        del wb['Sheet']
    
    # Separate color-based changes first
    color_based_changes = [change for change in changes if change.get("is_color_based")]
    color_change_count = len(color_based_changes)
    
    # Create Color-Based Changes sheet
    color_ws = wb.create_sheet("Color-Based Changes")
    # Headers for color-based changes
    color_headers = [
        "Change #", "Content Type", "Location", "Detected Color", 
        "Change Type", "Text Content", "Context", "Change Details"
    ]
    color_ws.append(color_headers)
    
    #  Style color sheet headers
    for col in range(1, len(color_headers) + 1):
        cell = color_ws.cell(row=1, column=col)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
    
    # Process color-based changes
    for idx, change in enumerate(color_based_changes, start=2):
        # Determine location
        if change["content_type"] == "paragraph":
            location = f"Paragraph {change.get('index', '')}"
        else:
            location = "Unknown"
            
        # Map color to change type and Excel fill
        color_fill = None
        if change["color"] == "RED":
            color_fill = PatternFill(start_color="FFCCCB", end_color="FFCCCB", fill_type="solid")
            change_type = "Removed Text"
        elif change["color"] == "YELLOW":
            color_fill = PatternFill(start_color="FFFFE0", end_color="FFFFE0", fill_type="solid")
            change_type = "New Addition"
        elif change["color"] == "GREEN":
            color_fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
            change_type = "Modified Copy"
        else:
            change_type = "Color-Coded"
            
        # Add row to color-based changes sheet
        row_data = [
            idx-1,
            change["content_type"].title(),
            location,
            change["color"],
            change_type,
            change["text"],
            f"Original: {change.get('original_context', '')}\nModified: {change.get('modified_context', '')}",
            change["change_detail"]
        ]
        color_ws.append(row_data)
        
        # Apply color coding to the entire row
        for col in range(1, len(row_data) + 1):
            if color_fill:
                color_ws.cell(row=idx, column=col).fill = color_fill
    
    # Set column widths for color sheet
    color_widths = [8, 15, 20, 12, 15, 40, 50, 40]
    for i, width in enumerate(color_widths, start=1):
        color_ws.column_dimensions[get_column_letter(i)].width = width
    
    summary_ws = wb.create_sheet("Summary")
    summary_ws.sheet_view.showGridLines = False
    
    summary_ws.merge_cells('A1:D1')
    title_cell = summary_ws['A1']
    title_cell.value = "ENHANCED DOCUMENT COMPARISON REPORT"
    title_cell.font = Font(bold=True, size=16)
    title_cell.alignment = Alignment(horizontal='center')

    summary_data = [
        ["Color-Based Changes Detected", color_change_count],
        ["Document Statistics", ""],
        ["Total Paragraphs", stats["paragraphs"]],
        ["Total Tables", stats["tables"]],
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

    changes_ws = wb.create_sheet("Changes")
    headers = [
        "Change #", "Type", "Location", "Status",
        "Original Content", "Modified Content",
        "Change Details", "Color Code"
    ]
    color_ws.append(headers)
    
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
    
    # Define color fills
    red_fill = PatternFill(start_color="FFCCCB", end_color="FFCCCB", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFFFE0", end_color="FFFFE0", fill_type="solid")
    green_fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
    
    # Process non-color-based changes
    non_color_changes = [change for change in changes if not change.get("is_color_based")]
    for idx, change in enumerate(non_color_changes, start=2):
        
        # Set location
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
                    img.anchor = f'E{idx}'  # Original Content column
                    changes_ws.add_image(img)
                except Exception as e:
                    print(f"Error adding deleted image to Excel: {e}")
            elif change["status"] == "added" and change.get("new_img"):
                try:
                    img = XLImage(io.BytesIO(change["new_img"]))
                    img.width, img.height = 40, 40
                    img.anchor = f'F{idx}'  # Modified Content column
                    changes_ws.add_image(img)
                except Exception as e:
                    print(f"Error adding added image to Excel: {e}")

        # Color and fill
        color_code = ""
        fill = None
        if change["status"] == "deleted":
            color_code = "Red - Removed"
            fill = red_fill
        elif change["status"] == "added":
            color_code = "Yellow - New Addition"
            fill = yellow_fill
        elif change["status"] in ["modified", "formatting_changed"]:
            color_code = "Green - Modified"
            fill = green_fill

        # Change details
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
            if change.get("strikethrough_original") or change.get("strikethrough_modified"):
                change_details += f"Strikethrough: {change.get('strikethrough_original')} → {change.get('strikethrough_modified')}\n"
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
        
        # Write row
        changes_ws.cell(row=idx, column=1, value=idx-1)  # Change #
        changes_ws.cell(row=idx, column=2, value=change["type"].replace("_", " ").title())
        changes_ws.cell(row=idx, column=3, value=location)
        changes_ws.cell(row=idx, column=4, value=change["status"].replace("_", " ").title())
        changes_ws.cell(row=idx, column=5, value=original_text)
        changes_ws.cell(row=idx, column=6, value=modified_text)
        changes_ws.cell(row=idx, column=7, value=change_details)
        changes_ws.cell(row=idx, column=8, value=color_code)
        
        # Apply fill to all cells in the row
        for col in range(1, 9):
            cell = changes_ws.cell(row=idx, column=col)
            if fill:
                cell.fill = fill
            # Add borders
            cell.border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )

    # Set column widths for changes sheet
    column_widths = [8, 15, 20, 12, 40, 40, 40, 15]
    for i, width in enumerate(column_widths, start=1):
        changes_ws.column_dimensions[get_column_letter(i)].width = width

    # Create Legend sheet
    legend_ws = wb.create_sheet("Color Legend")
    legend_data = [
        ["Color", "Meaning", "Description"],
        ["Red", "Removed Text", "Text marked with red color indicating removal"],
        ["Yellow", "New Addition", "Text marked with yellow color indicating new content"],
        ["Green", "Modified Copy", "Text marked with green color indicating modifications"]
    ]
    
    for row_idx, row_data in enumerate(legend_data, start=1):
        for col_idx, value in enumerate(row_data, start=1):
            cell = legend_ws.cell(row=row_idx, column=col_idx, value=value)
            if row_idx > 1:
                if value == "Red":
                    cell.fill = PatternFill(start_color="FFCCCB", end_color="FFCCCB", fill_type="solid")
                elif value == "Yellow":
                    cell.fill = PatternFill(start_color="FFFFE0", end_color="FFFFE0", fill_type="solid")
                elif value == "Green":
                    cell.fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
    
    # Set column widths for legend sheet
    legend_ws.column_dimensions['A'].width = 15
    legend_ws.column_dimensions['B'].width = 15
    legend_ws.column_dimensions['C'].width = 50

    # Set active sheet to Color-Based Changes
    wb.active = color_ws
    
    # Save to bytes buffer
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# ----------------- Flask Routes -----------------

@app.route("/", methods=["GET", "POST"])
def index():
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

            # For download requests, return the file
            if is_download_request:
                excel_file = export_to_excel_enhanced(changes, stats)
                return send_file(
                    excel_file,
                    as_attachment=True,
                    download_name="enhanced_document_comparison_report.xlsx",
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

if __name__ == "__main__":
    app.run(debug=True)