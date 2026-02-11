"""
Generate a test PDF with 90° anti-clockwise rotation
The text appears straight when viewed, but the page has internal rotation of 90°
Aspect ratio 4:5 (width:height when viewed correctly)
"""
import fitz  # PyMuPDF

def generate_rotated_test_pdf():
    # Create a new PDF document
    doc = fitz.open()
    
    # For 4:5 aspect ratio when viewed (width:height = 4:5)
    # With 90° rotation, the internal page is actually 5:4 (height:width)
    # Standard letter size adapted: 400x500 points when viewed
    # Internal dimensions: 500x400 (because of 90° rotation)
    internal_width = 500   # This becomes visual height
    internal_height = 400  # This becomes visual width
    
    # Create page with internal dimensions
    page = doc.new_page(width=internal_width, height=internal_height)
    
    # Set page rotation to 90 degrees (anti-clockwise)
    page.set_rotation(90)
    
    # After rotation:
    # - Visual width = internal height = 400
    # - Visual height = internal width = 500
    # - Aspect ratio = 400:500 = 4:5 ✓
    
    # We need to insert text in the INTERNAL coordinate system
    # Since the page is rotated 90° anti-clockwise:
    # - Internal X becomes Visual Y (bottom to top)
    # - Internal Y becomes Visual X (left to right, inverted)
    
    # To write text that appears at visual position (vx, vy):
    # Internal position: (vy, internal_height - vx)
    # But we also need to rotate the text matrix by -90 to counter the rotation
    
    # Create text writer for rotated content
    # We'll insert text with a rotation matrix to make it appear straight
    
    # Define the content as a bank statement
    content = [
        ("ACME BANK STATEMENT", 200, 30, 16, True),
        ("Account Number: 1234-5678-9012", 30, 60, 10, False),
        ("Statement Period: January 2026", 30, 80, 10, False),
        ("Account Holder: John Doe", 30, 100, 10, False),
        ("", 0, 0, 0, False),
        ("TRANSACTION DETAILS", 30, 140, 12, True),
        ("Date          Description                    Amount      Balance", 30, 165, 9, False),
        ("-" * 70, 30, 178, 8, False),
        ("01/01/2026    Opening Balance                            $5,000.00", 30, 195, 9, False),
        ("01/05/2026    Grocery Store Purchase         -$125.50    $4,874.50", 30, 212, 9, False),
        ("01/08/2026    Salary Credit                +$3,500.00    $8,374.50", 30, 229, 9, False),
        ("01/10/2026    Electric Bill                  -$89.00     $8,285.50", 30, 246, 9, False),
        ("01/12/2026    Online Shopping               -$250.00     $8,035.50", 30, 263, 9, False),
        ("01/15/2026    ATM Withdrawal                -$200.00     $7,835.50", 30, 280, 9, False),
        ("01/18/2026    Restaurant                     -$45.00     $7,790.50", 30, 297, 9, False),
        ("01/20/2026    Transfer from Savings        +$500.00     $8,290.50", 30, 314, 9, False),
        ("01/22/2026    Insurance Premium             -$150.00     $8,140.50", 30, 331, 9, False),
        ("01/25/2026    Fuel Station                   -$60.00     $8,080.50", 30, 348, 9, False),
        ("01/28/2026    Subscription Service           -$15.99     $8,064.51", 30, 365, 9, False),
        ("01/31/2026    Closing Balance                            $8,064.51", 30, 382, 9, False),
        ("-" * 70, 30, 395, 8, False),
        ("", 0, 0, 0, False),
        ("SUMMARY", 30, 420, 12, True),
        ("Total Credits:  $4,000.00", 50, 445, 10, False),
        ("Total Debits:   -$935.49", 50, 465, 10, False),
        ("Net Change:     +$3,064.51", 50, 485, 10, False),
    ]
    
    # Insert text with proper rotation
    # For 90° CCW rotation, we need to map visual coords to internal coords
    # Visual (vx, vy) -> Internal (vy, internal_height - vx)
    # And rotate text by 90° to appear straight
    
    for item in content:
        if item[0] == "":
            continue
        text, visual_x, visual_y, fontsize, bold = item
        
        # Convert visual coordinates to internal coordinates
        # For 90° CCW rotation: internal_x = visual_y, internal_y = internal_height - visual_x
        internal_x = visual_y
        internal_y = internal_height - visual_x
        
        # Create insertion point
        point = fitz.Point(internal_x, internal_y)
        
        # Create rotation matrix (90° CW to counter the page's 90° CCW rotation)
        # This makes text appear horizontal when viewed
        rotate_matrix = fitz.Matrix(90)
        
        fontname = "helv" if not bold else "hebo"
        
        # Insert text with rotation
        page.insert_text(
            point,
            text,
            fontname=fontname,
            fontsize=fontsize,
            rotate=90  # Rotate text 90° to appear straight
        )
    
    # Save the PDF
    output_path = "test_rotated_statement.pdf"
    doc.save(output_path)
    doc.close()
    
    print(f"Created: {output_path}")
    print(f"Page rotation: 90° (anti-clockwise)")
    print(f"Visual aspect ratio: 4:5")
    print(f"Internal dimensions: {internal_width}x{internal_height}")
    print(f"Visual dimensions: {internal_height}x{internal_width}")
    
    # Verify the PDF
    verify_doc = fitz.open(output_path)
    verify_page = verify_doc[0]
    print(f"\nVerification:")
    print(f"  Page rotation: {verify_page.rotation}°")
    print(f"  Page rect: {verify_page.rect}")
    print(f"  MediaBox: {verify_page.mediabox}")
    verify_doc.close()
    
    return output_path

if __name__ == "__main__":
    generate_rotated_test_pdf()
