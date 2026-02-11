"""
Debug script to test actual OCR extraction flow with FIXED cropped-region extraction.
"""
import sys
sys.path.insert(0, '.')

from modules import SessionLocal, OCRTemplate, OCRPage, LabeledBox
import fitz
from PIL import Image
import numpy as np

# Constants
OCR_DPI = 300

def main():
    # Load template from DB
    session = SessionLocal()
    template = session.query(OCRTemplate).order_by(OCRTemplate.id.desc()).first()
    
    if not template:
        print("No templates found!")
        return
    
    print(f"\n=== Using Template: {template.name} ===")
    
    # Get first page
    ocr_page = session.query(OCRPage).filter(OCRPage.template_id == template.id).first()
    print(f"Template Page: {ocr_page.pdf_filename}")
    
    # Get label boxes
    label_boxes = session.query(LabeledBox).filter(
        LabeledBox.page_id == ocr_page.id,
        LabeledBox.box_type == 'label'
    ).all()
    
    for lb in label_boxes:
        anchors = [b for b in lb.children if b.box_type == 'anchor']
        values = [b for b in lb.children if b.box_type == 'value']
        
        if not anchors or not values:
            continue
        
        anchor = anchors[0]
        value = values[0]
        
        # Get search text
        anchor_search_text = anchor.anchor_text or ''
        if not anchor_search_text or anchor_search_text.startswith("Anchor:"):
            anchor_search_text = anchor.name.replace("Anchor: ", "").replace("...", "").strip()
        
        print(f"\nLabel: {lb.name}")
        print(f"  Anchor search: '{anchor_search_text}'")
        print(f"  Anchor pos: ({anchor.x:.0f}, {anchor.y:.0f})")
        print(f"  Value pos: ({value.x:.0f}, {value.y:.0f})")
        
        # Calculate relative offsets
        value_dx = value.x - anchor.x
        value_dy = value.y - anchor.y
        print(f"  Relative offset: dx={value_dx:.0f}, dy={value_dy:.0f}")
        print(f"  Value size: {value.width:.0f} x {value.height:.0f}")
    
    # Test extraction with a PDF
    print("\n\n=== Testing Extraction with Cropped OCR ===")
    
    import os
    test_pdfs = [f for f in os.listdir('.') if f.endswith('.pdf')]
    if not test_pdfs:
        print("No PDF files found!")
        return
    
    test_pdf = test_pdfs[0]
    print(f"Testing with: {test_pdf}")
    
    try:
        import easyocr
        reader = easyocr.Reader(['en'], gpu=False, download_enabled=False)
    except Exception as e:
        print(f"EasyOCR error: {e}")
        return
    
    doc = fitz.open(test_pdf)
    page = doc.load_page(0)
    
    # Convert to image
    pix = page.get_pixmap(matrix=fitz.Matrix(OCR_DPI/72, OCR_DPI/72))
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    img_array = np.array(img)
    
    print(f"Page image: {img.width} x {img.height}")
    
    # Run OCR on full page to find anchor
    print("Running full-page OCR...")
    ocr_results = reader.readtext(img_array)
    
    # Convert to word dicts
    words = []
    for bbox, text, conf in ocr_results:
        x_coords = [p[0] for p in bbox]
        y_coords = [p[1] for p in bbox]
        words.append({
            'text': text.strip(),
            'left': min(x_coords),
            'top': min(y_coords),
            'width': max(x_coords) - min(x_coords),
            'height': max(y_coords) - min(y_coords)
        })
    
    print(f"Found {len(words)} OCR results")
    
    # For each label
    for lb in label_boxes:
        anchors = [b for b in lb.children if b.box_type == 'anchor']
        values = [b for b in lb.children if b.box_type == 'value']
        
        if not anchors or not values:
            continue
        
        anchor = anchors[0]
        value = values[0]
        
        anchor_search = anchor.anchor_text or ''
        if not anchor_search or anchor_search.startswith("Anchor:"):
            anchor_search = anchor.name.replace("Anchor: ", "").replace("...", "").strip()
        
        print(f"\n=== Extracting: {lb.name} ===")
        print(f"Searching for anchor: '{anchor_search}'")
        
        # Search for anchor
        found_anchor = None
        search_lower = anchor_search.lower().strip()
        
        for word in words:
            word_lower = word['text'].lower()
            if search_lower in word_lower or word_lower in search_lower:
                found_anchor = word
                print(f"FOUND anchor: '{word['text']}' at ({word['left']:.0f}, {word['top']:.0f})")
                break
        
        if not found_anchor:
            print("Anchor NOT FOUND!")
            continue
        
        # Calculate value rect
        value_dx = value.x - anchor.x
        value_dy = value.y - anchor.y
        
        value_x = found_anchor['left'] + value_dx
        value_y = found_anchor['top'] + value_dy
        value_w = value.width
        value_h = value.height
        
        print(f"Value rect: ({value_x:.0f}, {value_y:.0f}, {value_w:.0f}x{value_h:.0f})")
        
        # CROP and run OCR on value region (THE FIX!)
        x1 = max(0, int(value_x))
        y1 = max(0, int(value_y))
        x2 = min(img.width, int(value_x + value_w))
        y2 = min(img.height, int(value_y + value_h))
        
        print(f"Cropping to: ({x1}, {y1}) - ({x2}, {y2})")
        
        if x2 <= x1 or y2 <= y1:
            print("Invalid crop region!")
            continue
        
        cropped = img.crop((x1, y1, x2, y2))
        cropped_array = np.array(cropped)
        
        print(f"Cropped image size: {cropped.width} x {cropped.height}")
        
        # Run OCR on cropped region
        print("Running OCR on cropped value region...")
        value_ocr = reader.readtext(cropped_array)
        
        value_texts = [text.strip() for bbox, text, conf in value_ocr if text.strip()]
        extracted_value = " ".join(value_texts)
        
        print(f"\n*** EXTRACTED VALUE: '{extracted_value}' ***")
    
    doc.close()
    session.close()
    print("\n=== DONE ===")

if __name__ == "__main__":
    main()
