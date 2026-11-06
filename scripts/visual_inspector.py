import os
import sys
import tempfile
import win32com.client
import fitz  # PyMuPDF

def inspect_and_render_docx(docx_path, output_dir=None):
    docx_path = os.path.abspath(docx_path)
    if not os.path.exists(docx_path):
        raise FileNotFoundError(f"DOCX not found: {docx_path}")

    # Внутренний системный scratch агента, проект остается абсолютно чистым
    if output_dir is None:
        output_dir = os.path.join(tempfile.gettempdir(), "antigravity_visual_inspect")
    os.makedirs(output_dir, exist_ok=True)

    base_name = os.path.splitext(os.path.basename(docx_path))[0]
    pdf_path = os.path.join(output_dir, f"{base_name}.pdf")

    # 1. Export via MS Word COM
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(docx_path)
        word_pages = doc.ComputeStatistics(2)
        doc.SaveAs(pdf_path, FileFormat=17)
        doc.Close(False)
    finally:
        word.Quit()

    # 2. Render each page to 150 DPI PNG for internal agent inspection
    pdf_doc = fitz.open(pdf_path)
    page_images = []
    page_stats = []

    for i, page in enumerate(pdf_doc):
        pix = page.get_pixmap(dpi=150)
        img_path = os.path.join(output_dir, f"page_{i+1}.png")
        pix.save(img_path)
        page_images.append(img_path)

        # Calculate text bounds and fill height
        text = page.get_text()
        rects = page.get_text("blocks")
        lines = [l.strip() for l in text.splitlines() if l.strip()]

        min_y = 9999
        max_y = 0
        for b in rects:
            if b[4].strip(): # has text
                min_y = min(min_y, b[1])
                max_y = max(max_y, b[3])

        page_h = page.rect.height
        fill_pct = ((max_y - min_y) / page_h) * 100 if max_y > min_y else 0

        page_stats.append({
            "page": i + 1,
            "chars": len(text),
            "lines": len(lines),
            "fill_pct": round(fill_pct, 1),
            "image": img_path,
            "top_line": lines[0] if lines else "",
            "bottom_line": lines[-1] if lines else ""
        })

    pdf_doc.close()
    if os.path.exists(pdf_path):
        try:
            os.remove(pdf_path)
        except Exception:
            pass

    return {
        "docx": docx_path,
        "total_pages": len(page_stats),
        "word_pages": word_pages,
        "pages": page_stats
    }

if __name__ == "__main__":
    target = sys.argv[1] if len(sys.argv) > 1 else r"d:\Антигравити\legal_estate\deals\ПДКП_Протасовы_Фирсова.docx"
    res = inspect_and_render_docx(target)
    print(f"=== VISUAL INSPECTION REPORT: {res['total_pages']} PAGES ===")
    for p in res['pages']:
        print(f"Page {p['page']}: chars={p['chars']}, lines={p['lines']}, fill={p['fill_pct']}% | Top: '{p['top_line'][:40]}' | Bottom: '{p['bottom_line'][:40]}'")
        print(f"  IMAGE: {p['image']}")
