import os
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter

def separate_pdf_per_words(pdf_path, list_words, list_words_exclude, output_folder):
    with open(pdf_path, 'rb') as file:
        pdf = PdfReader(file)
        total_pages = len(pdf.pages)
        results = []  # Lista to storethe results (title file and amount of pages)

        for word in list_words:
            output = PdfWriter()  # Creates pdfwriter objet
            word_found = False  # indicates if word is present in a pdf page

            for page_num in range(total_pages):
                page = pdf.pages[page_num]
                content = page.extract_text().lower()  # Extrct pdf content

                if word.lower() in content:
                    # check if words need to be excluded
                    exclude = False
                    for word_exclude in list_words_exclude:
                        if word_exclude.lower() in content:
                            exclude = True
                            break

                    if not exclude:
                        output.add_page(page)
                        word_found = True

            if word_found:
                # generates output file
                file_name = f"pages_{word}.pdf"
                file_dir = os.path.join(output_folder, file_name)

                #save pdf that contains specific words
                with open(file_dir, 'wb') as output_file:
                    output.write(output_file)

                print(f"File  {file_name} was created with pages that contain the word {word}.")
                results.append((file_name, len(output.pages)))
            else:
                print(f"{word} was not found.")

        # Generar DataFrame con los resultados
        df_results = pd.DataFrame(results, columns=["Name", "Amount Pages"])

        # Guardar DataFrame en un archivo Excel
        excel = os.path.join(output_folder, "results.xlsx")
        df_results.to_excel(excel, index=False)
        print(f"File {excel} was created.")

# input file
input_pdf = 'input.pdf'

# Words that must be included to separate pdfs
words = ['PDF1', 'PDF2','PDF3','PDF4']

# list of words to be excluded
words_exclude = ['Do not print']

# Carpeta de salida para los archivos generados
output_folder = './output'

separate_pdf_per_words(input_pdf, words, words_exclude, output_folder)