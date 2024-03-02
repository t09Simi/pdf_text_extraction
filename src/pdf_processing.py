import fitz
from PIL import Image
import sparrow_extraction
import pdfplumber
import centurion_extraction
import first_integrated
import enermech_extraction


def pdf_to_text(pdf_path):
    doc = pdfplumber.open(pdf_path)
    text = ""
    page = doc.pages[0]
    text += page.extract_text()
    return text


def search_keyword(text, keywords):
    for keyword in keywords:
        if keyword.lower() in text.lower():
            return keyword
    return None  # Return None if no keyword is found


def is_empty(text):
    return not bool(text.strip())




def pdf_to_img(pdf_path, images_path):
    # Open the PDF document
    doc = fitz.open(pdf_path)

    # Loop through each page
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)

        # Render the page to a pixmap
        pix = page.get_pixmap()

        # Convert pixmap to a PIL Image
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

        # Save the image with page number
        img.save(f"{images_path}/page_{page_num + 1}.jpg", "JPEG")

    print(f"Converted {len(doc)} pages to images in {images_path}")


def main():
    pdf_path = "../resources/EnerMech.pdf"
    images_path = "../resources/images"
    text_content = pdf_to_text(pdf_path)
    keywords = ["Sparrows", "Centurion", "First Integrated"]

    if is_empty(text_content):
        print(f"No text found,Converting the pdf to images")
        pdf_to_img(pdf_path, images_path)
        enermech_extraction.extract_enermech_pdf(pdf_path)

    else:
        found_keyword = search_keyword(text_content, keywords)
        print(f"keyword found : {found_keyword}")
        if found_keyword and found_keyword == "Sparrows":
            sparrow_extraction.extract_sparrow_pdf(pdf_path)
        elif found_keyword and found_keyword == "Centurion":
            centurion_extraction.extraction_centurion_pdf(pdf_path)
        elif found_keyword and found_keyword == "First Integrated":
            first_integrated.extract_first_integrated_pdf(pdf_path)
        else:
            print(f"No matching keywords found")


if __name__ == "__main__":
    main()