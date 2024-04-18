import os
import win32com.client
from time import sleep


def pdfConverter(sleep_time=1):
    ppttoPDF = 32  # PowerPoint format type for PDF
    input_folder = input('Enter Folder Directory\n>>> ')

    # Create a subfolder 'merged' if it does not exist
    merged_folder = os.path.join(input_folder, 'merged')
    if not os.path.exists(merged_folder):
        os.makedirs(merged_folder)

    # Scan for files that need to be converted
    while any(f.endswith((".pptx", ".ppt")) for _, _, files in os.walk(input_folder) for f in files):
        for root, dirs, files in os.walk(input_folder):
            for file in files:
                sleep(sleep_time)
                if file.endswith(".pptx") or file.endswith(".ppt"):
                    extension_length = 5 if file.endswith(".pptx") else 4
                    try:
                        print(f'Trying to open {file}')
                        in_file = os.path.join(root, file)
                        powerpoint = win32com.client.Dispatch("Powerpoint.Application")
                        deck = powerpoint.Presentations.Open(in_file)
                        sleep(sleep_time)

                        # Save the PDF to the 'merged' subfolder
                        pdf_filename = file[:-extension_length] + ".pdf"
                        pdf_path = os.path.join(merged_folder, pdf_filename)
                        deck.SaveAs(pdf_path, ppttoPDF)  # formatType = 32 for ppt to pdf

                        deck.Close()
                        powerpoint.Quit()
                        print(f'Converted {file} successfully')

                        # Remove the original file
                        os.remove(in_file)
                    except Exception as e:
                        print(f'Failed to open {file}, will try again later: {e}')

    print('No more .ppt or .pptx files found')

if __name__ == '__main__':
    pdfConverter()
