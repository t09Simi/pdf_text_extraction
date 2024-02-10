import sparrow_extraction
import pdfplumber


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
    pdf_path = "../resources/sparrows.pdf"
    text_content = pdf_to_text(pdf_path)
    keywords = ["Sparrows", "keyword2"]

    if is_empty(text_content):
        print(f"No text found")
    else:
        found_keyword = search_keyword(text_content, keywords)
        if found_keyword:
            print(f"keyword found : {found_keyword}")
            sparrow_extraction.extract_sparrow_pdf(pdf_path)
        else:
            print(f"No matching keywords found")


if __name__ == "__main__":
    main()