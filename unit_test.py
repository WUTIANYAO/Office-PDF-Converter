# -*- coding: utf-8 -*-
"""
Created on Thu Oct 24 22:02:59 2024

@author: Ken
"""

import sys
sys.path.append(r"C:\Users\Ken\Desktop\Test\Code")
from excel_to_pdf_converter import ExcelToPDFConverter
from pdf_to_Image_converter import PDFToImageConverter
from word_to_pdf_converter import WordToPDFConverter
from ppt_to_pdf_converter import PPTToPDFConverter


# Example usage
PDFPath = 'C:\\Users\\Ken\\Desktop\\Test\\output.pdf'
ImagePath = 'C:\\Users\\Ken\\Desktop\\Test\\'
converter = PDFToImageConverter(PDFPath, ImagePath)
converter.convert_pdf_to_images()


input_file_path = "C:\\Users\\Ken\\Desktop\\Test\\test.docx"
output_base_folder = "C:\\Users\\Ken\\Desktop\\Test\\output.pdf"
font_path = 'C:\\Users\\Ken\\Desktop\\Test\\ipaexg.ttf'  
converter = WordToPDFConverter(input_file_path,output_base_folder,font_path)
converter.convert_word_to_pdf()


excel_file_path = 'C:\\Users\\Ken\\Desktop\\Test\\test1.xlsx'  
pdf_file_path = 'C:\\Users\\Ken\\Desktop\\Test\\Excel.pdf'  
font_path = 'C:\\Users\\Ken\\Desktop\\Test\\ipaexg.ttf' 
converter = ExcelToPDFConverter(font_path)
converter.convert_excel_to_pdf(excel_file_path, pdf_file_path)


input_file_path = "C:\\Users\\Ken\\Desktop\\test\\test.pptx"
output_filename = "C:\\Users\\Ken\\Desktop\\Test\\output.pdf"
font_path = 'C:\\Users\\Ken\\Desktop\\Test\\ipaexg.ttf'
converter = PPTToPDFConverter(input_file_path, output_filename, font_path)
converter.convert_ppt_to_pdf()
