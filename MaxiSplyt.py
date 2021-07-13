import os
from PyPDF2 import PdfFileReader, PdfFileWriter

files = []

def title_page_find(text):
    if "Location" in text and "Asset" in text and "Job Plan" in text and "Reported By" in text:
        return True

def pdf_splitter(path):
    fname = os.path.splitext(os.path.basename(path))[0]
    printed = 1

    pdf = PdfFileReader(path)
    for page in range(pdf.getNumPages()):

        test = title_page_find(pdf.getPage(page).extractText())
     
        if test:
            pdf_writer = PdfFileWriter()
            pdf_writer.addPage(pdf.getPage(page))

            output_filename = '{}_Section_{}.pdf'.format(
                fname, printed)
            
            with open(output_filename, 'wb') as out:
                pdf_writer.write(out)
                files.append(output_filename)
                printed += 1
        elif pdf.getPage(page).extractText().isspace():
                continue
        else:
            pdf_writer.addPage(pdf.getPage(page))
            with open(output_filename, 'ab') as out:
                pdf_writer.write(out)

   # for file in files:
    #    os.startfile(files[f], "print")


if __name__ == '__main__':
    path = r"C:\Support\SilentPrintServlet.pdf"
    pdf_splitter(path)
