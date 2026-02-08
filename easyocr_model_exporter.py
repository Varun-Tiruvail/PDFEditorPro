"""
EasyOCR Model Exporter - Exports models to text files for offline transfer.

This script exports the EasyOCR model files to base64-encoded text files
that can be copied/pasted to an offline workspace and reconstructed.
Files are split into chunks < 95MB for GitHub compatibility.

Usage:
    python easyocr_model_exporter.py export   # Export models to text files
    python easyocr_model_exporter.py import   # Import models from text files
"""
import os
import sys
import base64
import json
import zipfile
import io
from pathlib import Path

# EasyOCR model directory
EASYOCR_DIR = Path.home() / ".EasyOCR" / "model"
EXPORT_DIR = Path(__file__).parent / "easyocr_models"
MAX_CHUNK_SIZE = 40 * 1024 * 1024  # 40MB per file (under GitHub's 100MB limit)


def get_model_files():
    """Get list of EasyOCR model files"""
    if not EASYOCR_DIR.exists():
        print(f"[ERROR] EasyOCR model directory not found: {EASYOCR_DIR}")
        print("[INFO] Run EasyOCR once with internet to download models first.")
        return []
    
    files = []
    for f in EASYOCR_DIR.iterdir():
        if f.is_file():
            files.append(f)
    return files


def export_models():
    """Export EasyOCR models to multiple text files (base64 encoded, <95MB each)"""
    files = get_model_files()
    if not files:
        print("[ERROR] No model files found to export.")
        return False
    
    print(f"[INFO] Found {len(files)} model files:")
    total_size = 0
    for f in files:
        size_mb = f.stat().st_size / (1024 * 1024)
        total_size += f.stat().st_size
        print(f"  - {f.name} ({size_mb:.1f} MB)")
    
    print(f"[INFO] Total size: {total_size / (1024*1024):.1f} MB")
    
    # Create in-memory zip file
    print("[INFO] Creating compressed archive...")
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
        for f in files:
            print(f"  Compressing {f.name}...")
            zf.write(f, f.name)
    
    # Encode to base64
    print("[INFO] Encoding to base64...")
    zip_data = zip_buffer.getvalue()
    b64_data = base64.b64encode(zip_data).decode('ascii')
    
    # Create export directory
    EXPORT_DIR.mkdir(exist_ok=True)
    
    # Split into chunks
    num_chunks = (len(b64_data) + MAX_CHUNK_SIZE - 1) // MAX_CHUNK_SIZE
    print(f"[INFO] Splitting into {num_chunks} file(s) (max 40MB each)...")
    
    for i in range(num_chunks):
        start = i * MAX_CHUNK_SIZE
        end = min((i + 1) * MAX_CHUNK_SIZE, len(b64_data))
        chunk_data = b64_data[start:end]
        
        chunk_file = EXPORT_DIR / f"models_part{i+1}.txt"
        with open(chunk_file, 'w') as f:
            f.write(chunk_data)
        
        chunk_size = chunk_file.stat().st_size / (1024*1024)
        print(f"  Created {chunk_file.name} ({chunk_size:.1f} MB)")
    
    # Write metadata file
    metadata = {
        "format": "easyocr_offline_v2",
        "files": [f.name for f in files],
        "total_size_mb": total_size / (1024 * 1024),
        "compressed_size_mb": len(zip_data) / (1024 * 1024),
        "num_chunks": num_chunks,
        "chunk_pattern": "models_part{n}.txt"
    }
    
    meta_file = EXPORT_DIR / "metadata.json"
    with open(meta_file, 'w') as f:
        json.dump(metadata, f, indent=2)
    
    print(f"\n[SUCCESS] Models exported to: {EXPORT_DIR}")
    print(f"[INFO] Created {num_chunks + 1} files:")
    print(f"  - metadata.json")
    for i in range(num_chunks):
        print(f"  - models_part{i+1}.txt")
    
    print(f"\n[IMPORTANT] To use on offline workspace:")
    print(f"  1. Copy the entire '{EXPORT_DIR.name}' folder to target workspace")
    print(f"  2. Run: python easyocr_model_exporter.py import")
    return True


def import_models():
    """Import EasyOCR models from the exported text files"""
    meta_file = EXPORT_DIR / "metadata.json"
    
    if not meta_file.exists():
        print(f"[ERROR] Metadata file not found: {meta_file}")
        print(f"[INFO] Copy '{EXPORT_DIR.name}' folder to this directory first.")
        return False
    
    print(f"[INFO] Reading metadata...")
    with open(meta_file, 'r') as f:
        metadata = json.load(f)
    
    if metadata.get("format") not in ["easyocr_offline_v1", "easyocr_offline_v2"]:
        print("[ERROR] Invalid export format.")
        return False
    
    print(f"[INFO] Contains {len(metadata['files'])} model files:")
    for name in metadata['files']:
        print(f"  - {name}")
    
    # Read all chunks
    num_chunks = metadata.get("num_chunks", 1)
    print(f"[INFO] Reading {num_chunks} chunk file(s)...")
    
    b64_data = ""
    for i in range(num_chunks):
        chunk_file = EXPORT_DIR / f"models_part{i+1}.txt"
        if not chunk_file.exists():
            print(f"[ERROR] Missing chunk file: {chunk_file}")
            return False
        
        print(f"  Reading {chunk_file.name}...")
        with open(chunk_file, 'r') as f:
            b64_data += f.read()
    
    # Create EasyOCR directory
    EASYOCR_DIR.mkdir(parents=True, exist_ok=True)
    
    # Decode and extract
    print("[INFO] Decoding base64 data...")
    zip_data = base64.b64decode(b64_data)
    
    print("[INFO] Extracting model files...")
    with zipfile.ZipFile(io.BytesIO(zip_data), 'r') as zf:
        for name in zf.namelist():
            print(f"  Extracting {name}...")
            zf.extract(name, EASYOCR_DIR)
    
    print(f"\n[SUCCESS] Models imported to: {EASYOCR_DIR}")
    print("[INFO] EasyOCR is now ready to use offline!")
    return True


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        print("\nCommands:")
        print("  export - Export models from this machine (requires internet-downloaded models)")
        print("  import - Import models from text files (for offline use)")
        print(f"\nExport directory: {EXPORT_DIR}")
        return
    
    cmd = sys.argv[1].lower()
    
    if cmd == "export":
        export_models()
    elif cmd == "import":
        import_models()
    else:
        print(f"[ERROR] Unknown command: {cmd}")
        print("Use 'export' or 'import'")


if __name__ == "__main__":
    main()
