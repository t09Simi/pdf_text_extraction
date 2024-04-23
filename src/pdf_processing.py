import sparrow_extraction
import pdfplumber
import centurion_extraction
import first_integrated


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


def main():
    try:
        pdf_path = "../resources/CenturionLoft.pdf"
        images_path = "../resources/images"
        text_content = pdf_to_text(pdf_path)
        keywords = ["Sparrows", "Centurion", "First Integrated"]

        if is_empty(text_content):
            print(f"No text found")

        else:
            found_keyword = search_keyword(text_content, keywords)
            print(f"keyword found: {found_keyword}")
            if found_keyword and found_keyword == "Sparrows":
                sparrow_extraction.extract_sparrow_pdf(pdf_path)
            elif found_keyword and found_keyword == "Centurion":
                centurion_extraction.extraction_centurion_pdf(pdf_path)
            elif found_keyword and found_keyword == "First Integrated":
                first_integrated.extract_first_integrated_pdf(pdf_path)
            else:
                print(f"No matching keywords found")
    except Exception as e:
        print("Error in processing the PDF")


if __name__ == "__main__":
    main()