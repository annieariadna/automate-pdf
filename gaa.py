import pdfplumber
import pandas as pd
import re
import traceback
from typing import List, Dict, Any, Optional
import logging

try:
    with pdfplumber.open("test.pdf") as pdf:
        print(f"Procesando PDF con {len(pdf.pages)} páginas...")
        for page_num,page in enumerate(pdf.pages,1):
            text = page.extract_text()
            with open("output.txt", "a", encoding="utf-8") as f:
                f.write(f"--- Página {page_num} ---\n")
                f.write(text if text else "")
                f.write("\n-------------------\n")
except Exception as e:
    logging.error("Error al procesar el PDF:")
    logging.error(traceback.format_exc())