import fitz

doc = fitz.open(r'C:\Users\SarvjeetSinghTomar\AppData\Local\Temp\pdf_extraction_workspace\ea5840d2b3359bbf\question.pdf')
page = doc[2]
# Render a small strip at the right of column 2 to see what's there visually
clip = fitz.Rect(560, 100, 590, 700)
pix = page.get_pixmap(clip=clip, dpi=150)
pix.save(r'C:\Users\SarvjeetSinghTomar\AppData\Local\Temp\right_edge_debug.png')
print("Saved debug image")

# Also list all image blocks on the page
blocks = page.get_text("dict")["blocks"]
img_blocks = [b for b in blocks if b["type"] == 1]
print(f"Image blocks on page: {len(img_blocks)}")
for b in img_blocks:
    print(f"  image bbox: {b['bbox']}")
