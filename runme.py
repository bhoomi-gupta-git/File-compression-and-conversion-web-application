import os

def install_libraries():
    libraries = [
        'pdf2docx',
        'openpyxl',
        'reportlab',
        'tabula-py',
        'pandas',
        'pdfminer.six',
        'fpdf',
        'moviepy',
        'tkinter'
    ]

    for lib in libraries:
        print(f"Installing {lib}...")
        os.system(f"pip install {lib}")

    print("Installation complete.")

if __name__ == "__main__":
    install_libraries()
