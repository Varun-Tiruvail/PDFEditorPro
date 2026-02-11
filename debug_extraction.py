"""
DEBUG EXTRACTION SCRIPT v4 - FINAL APPROACH
"""
import fitz
import sqlite3

def main():
    print("=" * 60)
    print("DEBUG EXTRACTION SCRIPT v4")
    print("=" * 60)
    
    doc = fitz.open("test_rotated_statement.pdf")
    page = doc[0]
    
    # Read template
    conn = sqlite3.connect("data/automation_hub.db")
    cursor = conn.cursor()
    cursor.execute("""
        SELECT lb.name, lb.box_type, lb.x, lb.y, lb.width, lb.height
        FROM labeled_boxes lb WHERE lb.box_type IN ('anchor', 'value')
    """)
    boxes = cursor.fetchall()
    
    anchor = value = None
    for box in boxes:
        if box[1] == 'anchor':
            anchor = {'text': box[0], 'x': box[2], 'y': box[3], 'w': box[4], 'h': box[5]}
        elif box[1] == 'value':
            value = {'x': box[2], 'y': box[3], 'w': box[4], 'h': box[5]}
    
    print(f"\n[TEMPLATE - Visual Space]")
    print(f"  Anchor: '{anchor['text']}' at vis({anchor['x']:.1f}, {anchor['y']:.1f}) size {anchor['w']:.1f}x{anchor['h']:.1f}")
    print(f"  Value: at vis({value['x']:.1f}, {value['y']:.1f}) size {value['w']:.1f}x{value['h']:.1f}")
    
    # Visual offset
    vis_dx = value['x'] - anchor['x']  # 84 (value is to the RIGHT)
    vis_dy = value['y'] - anchor['y']  # 2 (value is slightly below)
    
    print(f"  Visual offset: dx={vis_dx:.1f} (right), dy={vis_dy:.1f} (down)")
    
    # Search for anchor
    found = page.search_for(anchor['text'])
    if not found:
        print("ERROR: Anchor not found!")
        return
    
    anchor_raw = found[0]
    print(f"\n[ANCHOR - Raw PDF Space]")
    print(f"  Found at: {anchor_raw}")
    
    print(f"\n[COORDINATE TRANSFORM]")
    print(f"  Visual dx={vis_dx:.1f} (right) -> Raw dy=-{vis_dx:.1f} (lower Y)")
    print(f"  Visual dy={vis_dy:.1f} (down) -> Raw dx=-{vis_dy:.1f} (left)")
    
    # For 90 degree rotation:
    # - Value to the RIGHT in visual = Value at LOWER Y in raw
    # - anchor_raw.y0 is where anchor text STARTS in raw Y
    # - But search_for returns where the specific text is found
    # - The value should be at lower Y values
    
    value_raw_y0 = anchor_raw.y0 - vis_dx
    value_raw_y1 = value_raw_y0 + value['w']  # Width in visual = height in raw Y
    
    print(f"\n[VALUE RECT CALCULATION]")
    print(f"  anchor_raw.y0 = {anchor_raw.y0:.1f}")
    print(f"  vis_dx = {vis_dx:.1f}")
    print(f"  value_raw_y0 = {anchor_raw.y0:.1f} - {vis_dx:.1f} = {value_raw_y0:.1f}")
    print(f"  value_raw_y1 = {value_raw_y0:.1f} + {value['w']:.1f} = {value_raw_y1:.1f}")
    
    value_rect = fitz.Rect(
        anchor_raw.x0 - 5,  # Same X range as anchor (with some buffer)
        value_raw_y0,
        anchor_raw.x1 + 5,
        value_raw_y1
    )
    
    print(f"\n[VALUE RECT] {value_rect}")
    extracted = page.get_text("text", clip=value_rect.normalize()).strip()
    print(f"[EXTRACTED] '{extracted}'")
    
    if "1234" in extracted:
        print("\n>> SUCCESS! Found the value!")
    else:
        print("\n>> Still not right, trying variations...")
        
        # Try a few variations
        for offset in [-20, -10, 0, 10, 20]:
            for size_adj in [0, 20, 40]:
                test_rect = fitz.Rect(
                    anchor_raw.x0 - 5,
                    anchor_raw.y0 - vis_dx + offset,
                    anchor_raw.x1 + 15,
                    anchor_raw.y0 - vis_dx + offset + value['w'] + size_adj
                )
                test_text = page.get_text("text", clip=test_rect.normalize()).strip()
                if "1234" in test_text:
                    print(f"  >> Found with offset={offset}, size_adj={size_adj}: '{test_text}'")
                    print(f"    Rect: {test_rect}")
    
    conn.close()
    doc.close()
    print("\n" + "=" * 60)

if __name__ == "__main__":
    main()
