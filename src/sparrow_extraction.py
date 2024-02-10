import pdfplumber

def extract_sparrow_pdf(pdf_path):
    print("<------------extracting sparrow pdf------------>")
    # pdf_doc = pdfplumber.open(pdf_path)
    # page = pdf_doc.pages[0]
    # page_tables = page.extract_tables()[0]
    # table_data = page_tables[0][0].split('\n')
    # print(table_data)
    # print("Certificate No ::" , table_data[0].split(':')[-1].strip())



if __name__ == "__main__":
    extract_sparrow_pdf("../resources/sparrows.pdf")
