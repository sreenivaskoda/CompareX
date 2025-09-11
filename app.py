import os
import io
import tempfile
import difflib
from hashlib import md5
from collections import defaultdict

import msoffcrypto
import docx
import openpyxl
from flask import Flask, render_template, request, send_file, flash, jsonify, redirect, url_for
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
from PIL import Image as PILImage

app = Flask(__name__)
app.secret_key = "supersecretkey"
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# ----------------- Utility functions -----------------
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
    """Extract all content from DOCX in order (paragraphs, tables, images)."""
    content = []
    
    # Extract paragraphs
    for para in doc.paragraphs:
        if para.text.strip():
            content.append({
                "type": "paragraph",
                "text": para.text,
                "runs": [run.text for run in para.runs if run.text.strip()]
            })
    
    # Extract tables with content and structure
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

def get_docx_images(doc):
    """Extract all images with hash + raw bytes."""
    images = []
    for rel in doc.part.rels.values():
        if "image" in rel.reltype:
            img_bytes = rel.target_part.blob
            img_hash = md5(img_bytes).hexdigest()
            images.append({"hash": img_hash, "bytes": img_bytes})
    return images

def get_table_images(doc):
    """Extract images from inside table cells with location info."""
    table_images = []
    NSMAP = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}

    for t_idx, table in enumerate(doc.tables, start=1):
        for r_idx, row in enumerate(table.rows, start=1):
            for c_idx, cell in enumerate(row.cells, start=1):
                for drawing in cell._element.xpath(".//a:blip"):
                    embed = drawing.attrib.get(
                        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
                    )
                    if embed and embed in doc.part.rels:
                        img_part = doc.part.rels[embed].target_part
                        img_bytes = img_part.blob
                        img_hash = md5(img_bytes).hexdigest()
                        table_images.append({
                            "table": t_idx,
                            "row": r_idx,
                            "col": c_idx,
                            "hash": img_hash,
                            "bytes": img_bytes
                        })
    return table_images

def compare_text_content(content1, content2):
    """Compare text content between two documents."""
    changes = []
    
    # Compare paragraphs
    para1 = [item for item in content1 if item["type"] == "paragraph"]
    para2 = [item for item in content2 if item["type"] == "paragraph"]
    
    for i in range(max(len(para1), len(para2))):
        if i < len(para1) and i < len(para2):
            if para1[i]["text"] != para2[i]["text"]:
                # Use difflib to show exact differences
                diff = list(difflib.ndiff(para1[i]["text"].split(), para2[i]["text"].split()))
                added = [word[2:] for word in diff if word.startswith('+ ')]
                removed = [word[2:] for word in diff if word.startswith('- ')]
                
                changes.append({
                    "type": "paragraph",
                    "index": i + 1,
                    "status": "modified",
                    "original": para1[i]["text"],
                    "modified": para2[i]["text"],
                    "added": " ".join(added),
                    "removed": " ".join(removed)
                })
        elif i < len(para1):
            changes.append({
                "type": "paragraph",
                "index": i + 1,
                "status": "deleted",
                "original": para1[i]["text"],
                "modified": "",
                "added": "",
                "removed": para1[i]["text"]
            })
        elif i < len(para2):
            changes.append({
                "type": "paragraph",
                "index": i + 1,
                "status": "added",
                "original": "",
                "modified": para2[i]["text"],
                "added": para2[i]["text"],
                "removed": ""
            })
    
    return changes

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

def compare_images(imgs1, imgs2):
    """Compare document-level images by hash."""
    changes = []
    hashes1 = {img["hash"]: img for img in imgs1}
    hashes2 = {img["hash"]: img for img in imgs2}

    # Deleted images
    for h1, img1 in hashes1.items():
        if h1 not in hashes2:
            changes.append({
                "type": "image",
                "status": "deleted",
                "hash": h1,
                "old_img": img1["bytes"],
                "new_img": None
            })

    # Inserted images
    for h2, img2 in hashes2.items():
        if h2 not in hashes1:
            changes.append({
                "type": "image",
                "status": "added",
                "hash": h2,
                "old_img": None,
                "new_img": img2["bytes"]
            })

    return changes

def compare_table_images(imgs1, imgs2):
    """Compare images in table cells by hash and location."""
    changes = []
    # Group by table, row, col
    grouped1 = defaultdict(list)
    grouped2 = defaultdict(list)
    
    for img in imgs1:
        key = (img["table"], img["row"], img["col"])
        grouped1[key].append(img)
        
    for img in imgs2:
        key = (img["table"], img["row"], img["col"])
        grouped2[key].append(img)
    
    # Check all keys in both groups
    all_keys = set(grouped1.keys()) | set(grouped2.keys())
    
    for key in all_keys:
        imgs_in_cell1 = grouped1.get(key, [])
        imgs_in_cell2 = grouped2.get(key, [])
        
        hashes1 = {img["hash"] for img in imgs_in_cell1}
        hashes2 = {img["hash"] for img in imgs_in_cell2}
        
        # Images deleted from this cell
        for img in imgs_in_cell1:
            if img["hash"] not in hashes2:
                changes.append({
                    "type": "table_image",
                    "status": "deleted",
                    "table": key[0],
                    "row": key[1],
                    "col": key[2],
                    "hash": img["hash"],
                    "old_img": img["bytes"],
                    "new_img": None
                })
        
        # Images added to this cell
        for img in imgs_in_cell2:
            if img["hash"] not in hashes1:
                changes.append({
                    "type": "table_image",
                    "status": "added",
                    "table": key[0],
                    "row": key[1],
                    "col": key[2],
                    "hash": img["hash"],
                    "old_img": None,
                    "new_img": img["bytes"]
                })
    
    return changes

def compare_docx(doc1, doc2):
    """Compare two DOCX documents comprehensively."""
    # Extract content
    content1 = extract_docx_content(doc1)
    content2 = extract_docx_content(doc2)
    
    # Extract images
    imgs1 = get_docx_images(doc1)
    imgs2 = get_docx_images(doc2)
    table_imgs1 = get_table_images(doc1)
    table_imgs2 = get_table_images(doc2)
    
    # Compare different aspects
    text_changes = compare_text_content(content1, content2)
    table_changes = compare_tables(content1, content2)
    image_changes = compare_images(imgs1, imgs2)
    table_image_changes = compare_table_images(table_imgs1, table_imgs2)
    
    # Combine all changes
    all_changes = text_changes + table_changes + image_changes + table_image_changes
    
    # Count elements
    para_count = len([item for item in content1 if item["type"] == "paragraph"])
    table_count = len([item for item in content1 if item["type"] == "table"])
    image_count = len(imgs1)
    table_image_count = len(table_imgs1)
    
    stats = {
        "paragraphs": para_count,
        "tables": table_count,
        "images": image_count,
        "table_images": table_image_count,
        "total_elements": para_count + table_count
    }
    
    return all_changes, stats

def insert_image(ws, img_bytes, cell, max_size=(120, 120)):
    """Insert an image into Excel cell from raw bytes."""
    if not img_bytes:
        return
        
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
            pil_img = PILImage.open(io.BytesIO(img_bytes))
            pil_img.thumbnail(max_size)
            pil_img.save(tmp.name, format="PNG")
            img = XLImage(tmp.name)
            img.anchor = cell
            ws.add_image(img)
            os.unlink(tmp.name)
    except Exception as e:
        print(f"Error inserting image: {e}")

def export_to_excel(changes, stats):
    """Export comparison results to a well-formatted Excel file."""
    wb = openpyxl.Workbook()
    
    # Remove default sheet and create organized ones
    if 'Sheet' in wb.sheetnames:
        del wb['Sheet']
    
    # Summary sheet
    summary_ws = wb.create_sheet("Summary")
    summary_ws.sheet_view.showGridLines = False
    
    # Add title
    summary_ws.merge_cells('A1:D1')
    title_cell = summary_ws['A1']
    title_cell.value = "DOCUMENT COMPARISON REPORT"
    title_cell.font = Font(bold=True, size=16)
    title_cell.alignment = Alignment(horizontal='center')
    
    # Add summary stats
    summary_data = [
        ["Document Statistics", ""],
        ["Total Paragraphs", stats["paragraphs"]],
        ["Total Tables", stats["tables"]],
        ["Total Images", stats["images"]],
        ["Total Table Images", stats["table_images"]],
        ["", ""],
        ["Comparison Results", ""],
        ["Total Changes Detected", len(changes)],
    ]
    
    for i, (label, value) in enumerate(summary_data, start=3):
        summary_ws[f'A{i}'] = label
        summary_ws[f'B{i}'] = value
        summary_ws[f'A{i}'].font = Font(bold=True)
    
    # Changes sheet
    changes_ws = wb.create_sheet("Changes")
    
    # Define headers
    headers = [
        "Change #", "Type", "Location", "Status", 
        "Original Content", "Modified Content", 
        "Original Image", "Modified Image"
    ]
    
    changes_ws.append(headers)
    
    # Style headers
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
    
    # Add changes data
    for idx, change in enumerate(changes, start=2):
        # Determine location text based on change type
        if change["type"] == "paragraph":
            location = f"Paragraph {change['index']}"
        elif change["type"] == "table":
            location = f"Table {change['index']}"
        elif change["type"] == "table_cell":
            location = f"Table {change['table_index']}, Cell ({change['row']},{change['col']})"
        elif change["type"] == "table_row":
            location = f"Table {change['table_index']}, Row {change['row']}"
        elif change["type"] == "image":
            location = "Document"
        elif change["type"] == "table_image":
            location = f"Table {change['table']}, Cell ({change['row']},{change['col']})"
        else:
            location = "Unknown"
        
        # Add row data
        changes_ws.cell(row=idx, column=1, value=idx-1)  # Change #
        changes_ws.cell(row=idx, column=2, value=change["type"].replace("_", " ").title())
        changes_ws.cell(row=idx, column=3, value=location)
        changes_ws.cell(row=idx, column=4, value=change["status"].replace("_", " ").title())
        
        # Add text content
        original_text = change.get("original", "")
        modified_text = change.get("modified", "")
        
        changes_ws.cell(row=idx, column=5, value=original_text)
        changes_ws.cell(row=idx, column=6, value=modified_text)
        
        # Add images if available
        if "old_img" in change and change["old_img"]:
            insert_image(changes_ws, change["old_img"], f"G{idx}")
        if "new_img" in change and change["new_img"]:
            insert_image(changes_ws, change["new_img"], f"H{idx}")
    
    # Set column widths
    column_widths = [8, 12, 20, 12, 40, 40, 20, 20]
    for i, width in enumerate(column_widths, start=1):
        changes_ws.column_dimensions[get_column_letter(i)].width = width
    
    # Set the changes sheet as active
    wb.active = changes_ws
    
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

            # Compare documents
            changes, stats = compare_docx(doc1, doc2)

            # For download requests, return the file
            if is_download_request:
                excel_file = export_to_excel(changes, stats)
                return send_file(
                    excel_file,
                    as_attachment=True,
                    download_name="document_comparison_report.xlsx",
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