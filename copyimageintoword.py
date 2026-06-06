from pathlib import Path
import win32com.client

# =====================================================
# SETTINGS
# =====================================================

DOC_PATH = r"E:\VERTIV\flirtest\template2.docx"
OUTPUT_PATH = r"E:\VERTIV\flirtest\report.docx"
IMAGE_FOLDER = r"E:\VERTIV\flirtest"

PHOTO_ROW_HEIGHT = 50  # points
PHOTO_ROW_WIDTH = 70

# =====================================================
# GET IMAGES
# =====================================================

images = sorted(Path(IMAGE_FOLDER).glob("*.jpg"))

print(f"Found {len(images)} images")

# =====================================================
# OPEN WORD
# =====================================================

word = win32com.client.Dispatch("Word.Application")
word.Visible = True

doc = word.Documents.Open(DOC_PATH)

# =====================================================
# OPTIONAL: SET TABLE ROW HEIGHTS
# =====================================================

for table in doc.Tables:
    for row in table.Rows:
        try:
            row.HeightRule = 1      # wdRowHeightExactly
            row.Height = PHOTO_ROW_HEIGHT
        except:
            pass

# =====================================================
# INSERT LINKED IMAGES
# =====================================================

for table in doc.Tables:

    for row in table.Rows:

        for cell in row.Cells:

            try:
                cell_text = cell.Range.Text

                for idx, image in enumerate(images, start=1):

                    placeholder = f"{{{{PHOTO{idx}}}}}"

                    if placeholder in cell_text:

                        rng = cell.Range

                        # Remove Word end-of-cell markers
                        rng.End = rng.End - 2

                        text = rng.Text

                        # Remove all PHOTO placeholders
                        text = text.replace(placeholder, "")

                        # Clear placeholder
                        rng.Text = text
                        print(repr(cell.Range.Text))

                        word_path = str(image).replace("\\", "\\\\")
                        print(repr(cell.Range.Text))

                        field_code = (
                            f'INCLUDEPICTURE "{word_path}"\\d'
                        )

                        field = doc.Fields.Add(
                            Range=rng,
                            Type=-1,
                            Text=field_code,
                            PreserveFormatting=True
                        )

                        field.Update()

                        print(f"Linked {placeholder} -> {image}")
                        print(repr(cell.Range.Text))
                        break

            except Exception as e:
                print(f"Cell error: {e}")

# =====================================================
# UPDATE ALL FIELDS
# =====================================================

doc.Fields.Update()

# =====================================================
# RESIZE IMAGES TO FIT CELL
# =====================================================

for shape in doc.InlineShapes:

    try:
        # Preserve aspect ratio
        shape.LockAspectRatio = True

        # Set desired height
        shape.Height = PHOTO_ROW_HEIGHT
        shape.width = PHOTO_ROW_WIDTH

        print(
            f"Image resized -> "
            f"W={shape.Width:.1f} pt, "
            f"H={shape.Height:.1f} pt"
        )

    except Exception as e:
        print(f"Resize error: {e}")

# =====================================================
# SAVE
# =====================================================

doc.SaveAs(OUTPUT_PATH)

print(f"\nSaved: {OUTPUT_PATH}")
print(f"Images inserted: {doc.InlineShapes.Count}")


# Leave Word open for inspection
# Uncomment below if you want Word closed automatically

# doc.Close()
# word.Quit()