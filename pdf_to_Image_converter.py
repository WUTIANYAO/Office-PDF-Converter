# -*- coding: utf-8 -*-
"""
Created on Thu Oct 24 23:41:55 2024

@author: Ken
"""

import fitz  # PyMuPDF
from PIL import Image
import os

class PDFToImageConverter:
    def __init__(self, input_path, output_base_folder):
        """
        Initialize the PDFToImageConverter class.
        
        :param input_path: str, path to the input PDF file
        :param output_base_folder: str, path to the base folder where output images will be saved
        """
        self.input_path = input_path
        self.output_folder = self._create_output_folder(output_base_folder)

    def _create_output_folder(self, base_folder):
        """
        Create a directory for the output images based on the PDF file name.
        
        :param base_folder: str, the base directory where the output folder will be created
        :return: str, the path to the created output folder
        """
        # Extract base name of the PDF file without extension
        pdf_name = os.path.basename(self.input_path)
        folder_name = os.path.splitext(pdf_name)[0]
        output_folder = os.path.join(base_folder, folder_name)

        # Ensure the folder exists
        os.makedirs(output_folder, exist_ok=True)
        return output_folder

    def convert_pdf_to_images(self, zoom_factor=2.0):
        """
        Convert each page of a PDF file into high-resolution images and save them.
        
        :param zoom_factor: float, scaling factor to adjust the resolution of the output images
                            - Default is 2.0, meaning images are enlarged by 2 times (i.e., 4 times more pixels)
                            - Higher values will produce higher resolution images but require more memory and processing time
        """
        if not self.input_path.lower().endswith('.pdf'):
            raise ValueError("Input file is not a PDF.")

        # Open the PDF document
        pdf_document = fitz.open(self.input_path)
        try:
            for page_number in range(pdf_document.page_count):
                page = pdf_document.load_page(page_number)
                zoom_matrix = fitz.Matrix(zoom_factor, zoom_factor)
                pix = page.get_pixmap(matrix=zoom_matrix)
                image_path = os.path.join(self.output_folder, f"page_{page_number + 1}.png")
                image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                image.save(image_path, "PNG")
                print(f"Page {page_number + 1} saved as {image_path}")
        finally:
            pdf_document.close()
        print(f"All pages have been saved to the folder: {self.output_folder}")