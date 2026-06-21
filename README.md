# Pdf and Image Text Extraction using EasyOCR and PyMuPDF

## Project Overview
This project extracts text and tabular data from both text-based and scanned PDF documents using PyMuPDF and EasyOCR. The extracted content is saved in CSV, TXT, and DOCX formats.

## Installation Steps
1. Install Python 3.8 or later.
2. Install dependencies:

pip install easyocr PyMuPDF python-docx Pillow torch torchvision numpy opencv-python

## Execution Steps
1. Open a terminal in the project directory.
2. Run:
python ttpy.py
3. Select a pdf file.
4. Wait for processing to complete.
5. View generated CSV, TXT, and DOCX files.

## OCR Technology Used
EasyOCR is a deep-learning-based OCR engine used to extract text from scanned Pdfs and images.

Workflow:
- Convert Pdf pages to images
- Detect text region
- Recognize characters and words
- Reconstruct rows and tables
- Save extracted data

## Sample Outputs

TXT Output:
Name | Roll No | Marks
John | 101 | 85
Mary | 102 | 90

CSV Output:
Name,Roll No,Marks
John,101,85
Mary,102,90

Generated Files:
- Sample_OCR.csv
- Sample_OCR.txt
- Sample_OCR.docx

## Conclusion
The project automates Pdf text extraction using PyMuPDF and EleasyOCR and exports result into multiple formats.
