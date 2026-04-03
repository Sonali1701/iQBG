import fitz

doc = fitz.open(r'C:\Users\SarvjeetSinghTomar\AppData\Local\Temp\pdf_extraction_workspace\ea5840d2b3359bbf\question.pdf')
page = doc[2]
# Get a pixmap of the right edge of the page, top half (to avoid checking the whole page)
clip = fitz.Rect(500, 100, 595, 400)
pix = page.get_pixmap(clip=clip, colorspace=fitz.csGRAY)

# Pixmap is a grayscale byte array.
# For each x column, count how many dark pixels there are.
# Dark pixel: value < 100

dark_counts = [0] * clip.width

# A line is vertical. So if we find a column x where dark_counts is very high (say, close to clip.height),
# that x is the vertical line.
clip_h = int(clip.height)
clip_w = int(clip.width)

bytes_per_row = pix.stride
samples = pix.samples

for y in range(clip_h):
    for x in range(clip_w):
        val = samples[y * bytes_per_row + x]
        if val < 100:  # dark
            dark_counts[x] += 1

# Print x coordinates that have a high concentration of dark pixels
print("Dark columns (y=100 to 400):")
found=False
for x in range(clip_w):
    # A true line should have dark pixels on almost all y's
    if dark_counts[x] > clip_h * 0.5: 
        real_x = 500 + x
        print(f"Vertical line found at exact X coordinate: {real_x} (density: {dark_counts[x]}/{clip_h})")
        found=True

if not found:
    print("No vertical line found. Checking threshold 20%:")
    for x in range(clip_w):
        if dark_counts[x] > clip_h * 0.2: 
            real_x = 500 + x
            print(f"Possible line / text column at X coordinate: {real_x} (density: {dark_counts[x]}/{clip_h})")
    
