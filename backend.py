import os
import re
import zipfile
import shutil
from pathlib import Path

def run_batch(input_docs, input_images, output_docs, log_func=print):
    skipped = []
    processed = 0

    for docx_file in Path(input_docs).glob("*.docx"):
        log_func(f"Xử lý: {docx_file.name}")

        sample_id = extractID(docx_file.name)
        if not sample_id:
            reason = "Không tìm thấy mã ID"
            skipped.append((docx_file.name, reason))
            log_func(f"\t{reason}")
            continue

        imageset = collectImagesWithID(input_images, sample_id)
        if not imageset:
            reason = f"Gặp trục trặc với ảnh cho mã {sample_id}"
            skipped.append((docx_file.name, reason))
            log_func(f"\t{reason}")
            continue

        out_path = Path(output_docs) / docx_file.name
        success = replaceImages(docx_file, imageset, out_path, log_func)

        if success:
            processed += 1
            log_func("\tThành công, đã chèn ảnh A và K")
        else:
            reason = "Không thể thay thế ảnh (placeholder không hợp lệ)"
            skipped.append((docx_file.name, reason))
            log_func(f"\t{reason}")
    return skipped, processed

def extractID(filename):
    match = re.search(r"(M\d+)", filename)
    return match.group(1) if match else None

def collectImagesWithID(image_dir, sample_id):
    images = [f for f in os.listdir(image_dir) if f.startswith(sample_id)]
    result = {"A": None, "K": None}

    for img in images:
        parts = img.split(".")
        if len(parts) >= 3:
            suffix = parts[-2].upper()
            if suffix in result:
                result[suffix] = os.path.join(image_dir, img)
    
    if len(images) != 2 or not result["A"] or not result["K"]:
        return None
    return result

def replaceImages(docx_path, images, output_path, log_func=print):
    with zipfile.ZipFile(docx_path, "r") as zin:
        with zipfile.ZipFile(output_path, "w") as zout:
            image_files = [n for n in zin.namelist() if n.startswith("word/media/image")]
            image_files.sort()

            if len(image_files) != 2:
                log_func(f"[!] {docx_path.name}: số ảnh placeholder trong template khác 2")
                return
            
            log_func(f"{docx_path.name}: placeholders = {image_files}")
            log_func(f"   -> {image_files[0]} <= {images['A']}")
            log_func(f"   -> {image_files[1]} <= {images['K']}")
            
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == image_files[0]:
                    with open(images["A"], "rb") as f:
                        data = f.read()
                elif item.filename == image_files[1]:
                    with open(images["K"], "rb") as f:
                        data = f.read()
                zout.writestr(item, data, compress_type=item.compress_type)

        try:
            with zipfile.ZipFile(output_path, "r") as testzip:
                bad = testzip.testzip()
                if bad is not None:
                    raise zipfile.BadZipFile(f"Corrupted entry: {bad}")
            return True
        except Exception as e:
            log_func(f"[!] Lỗi: file {output_path} bị hỏng ({e})")
            if output_path.exists():
                output_path.unlink()
            return False