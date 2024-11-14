# -*- coding: utf-8 -*-
"""
Created on Thu Oct 24 23:55:36 2024

@author: Ken
"""

import mammoth
import os
import textwrap
from bs4 import BeautifulSoup
from xhtml2pdf import pisa
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from xhtml2pdf.default import DEFAULT_FONT
import reportlab.lib.styles

class WordToPDFConverter:
    def __init__(self, input_file_path, pdf_file_path, font_path):
        """
        Initialize the WordToPDFConverter class.
        
        :param input_file_path: str, path to the input Word file
        :param pdf_file_path: str, full path to where the output PDF file will be saved
        :param font_path: str, path to the font file for PDF conversion
        """
        self.input_file_path = input_file_path
        self.pdf_file_path = pdf_file_path
        self.font_path = font_path

    def convert_docx_to_html(self, max_line_length=43, list_line_length=37):
        # Convert DOCX to HTML in-memory string
        with open(self.input_file_path, "rb") as docx_file:
            result = mammoth.convert_to_html(docx_file)
            html_content = result.value
            
            # Parse and process HTML
            soup = BeautifulSoup(html_content, "html.parser")
            self._process_html(soup, max_line_length, list_line_length)
            
            style_content = """
            <head>
              <meta charset="UTF-8"> 
              <style type="text/css">
                body {
                font-family: 'JapaneseFont'; 
                font-size: 12pt; 
                max-width: 794px; 
                margin: 0 auto; 
                padding: 10px; 
                white-space: normal;
                word-wrap: break-word;
                }
                .table-container {
                    max-width: 100%; 
                    padding: 0 50px; 
                }
                table {
                    width: 80%; 
                    border-collapse: collapse;
                    margin: 20px auto; 
                }
                th, td {
                    border: 1px solid black;
                    padding: 8px;
                    text-align: left;
                }
                p {
                    margin: 0; 
                    padding: 5px 0; 
                }
                li {
                    margin-bottom: 5px; 
                    white-space: pre-wrap; 
                }
                img {
                    display: block;
                    margin: 10px 0;
                    max-width: 100%; 
                    height: auto;
                }
              </style>
            </head>
            """
            
            complete_html = f"{style_content}\n<body>\n{soup}\n</body>\n"
            complete_html = complete_html.replace("<img", "<p><img").replace("/>", "/></p>")

        return complete_html

    def _process_html(self, soup, max_line_length, list_line_length):
        flatten_nested_lists(soup)
        remove_empty_tags(soup)
        clean_html(soup)
        break_text_into_lines(soup, max_line_length, list_line_length)

    def convert_html_to_pdf(self, html_content):
        # Ensure custom font is registered
        pdfmetrics.registerFont(TTFont('JapaneseFont', self.font_path))
        reportlab.lib.styles.ParagraphStyle.defaults['wordWrap'] = 'CJK'
        DEFAULT_FONT['helvetica'] = 'JapaneseFont'  

        # Convert in-memory HTML to PDF
        with open(self.pdf_file_path, 'wb') as pdf_file:
            pisa_status = pisa.CreatePDF(
                html_content,                  
                dest=pdf_file,                  
                encoding='utf-8',
                page_size='A4',               
            )

        if pisa_status.err:
            print("Error occurred during HTML to PDF conversion.")
        else:
            print(f"PDF file has been saved to: {self.pdf_file_path}")

        return pisa_status.err

    def convert_word_to_pdf(self):
        html_content = self.convert_docx_to_html()
        self.convert_html_to_pdf(html_content)

# Function Definitions
def flatten_nested_lists(soup):
    # Flatten nested lists
    for list_tag in soup.find_all(['ul', 'ol']):
        while True:
            nested = list_tag.find(['ul', 'ol'])
            if not nested:
                break
            for li in nested.find_all('li', recursive=False):
                list_tag.append(li)
            nested.decompose()

def remove_empty_tags(soup):
    # 去除空的标签
    for tag in soup.find_all():
        if tag.name in ['p', 'li', 'ul', 'ol'] and not tag.get_text(strip=True) and not tag.find('img'):
            tag.decompose()

def break_text_into_lines(soup, max_line_length, list_line_length):
    # Break text into specified line length
    for p in soup.find_all('p'):
        if p.string:
            p.string.replace_with("\n".join(textwrap.wrap(p.text, max_line_length)))
    for li in soup.find_all('li'):
        li_text = li.get_text(separator=' ', strip=True)
        wrapped_text = "\n".join(textwrap.wrap(li_text, list_line_length))
        li.clear()
        li.append(wrapped_text)

def clean_html(soup):
    # Clean up HTML structure
    for br in soup.find_all("br"):
        if not br.next_sibling:
            br.extract()
            
            

