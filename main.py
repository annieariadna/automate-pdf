import fitz  # PyMuPDF
import pandas as pd
import re
from pathlib import Path
import sys

def extract_text_from_pdf(pdf_path):
    """Extrae texto de todas las páginas del PDF"""
    doc = fitz.open(pdf_path)
    pages_text = []
    
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text = page.get_text()
        pages_text.append(text)
    
    doc.close()
    return pages_text

def extract_invoice_data_from_page(page_text):
    """Extrae RUC, RAZÓN SOCIAL y DIRECCIÓN de una página individual"""
    
    # Limpiar el texto para mejorar la extracción
    text = re.sub(r'\s+', ' ', page_text).strip()
    
    data = {'ruc': '', 'razon_social': '', 'direccion': ''}
    
    # Buscar el número de factura
    factura_match = re.search(r'FACTURA ELECTRÓNICA\s*F(\d{3}-\d{8})', text, re.IGNORECASE)
    if factura_match:
        data['numero_factura'] = factura_match.group(1)
    else:
        data['numero_factura'] = 'Sin número'
    
    # Patrón para RUC y RAZÓN SOCIAL (están en secuencia específica)
    # Buscar la sección que contiene RUC RAZÓN SOCIAL seguido de número y nombre
    ruc_razon_pattern = r'RUC\s+RAZÓN SOCIAL\s+(\d{11})\s+([^0-9]+?)(?=\s+AV\.|JR\.|CALLE|CAL\.|MZA\.|PSJ\.|URB\.|DPTO|PISO|\d|\s+Nº\s+GUÍA)'
    
    match = re.search(ruc_razon_pattern, text, re.IGNORECASE | re.DOTALL)
    if match:
        data['ruc'] = match.group(1).strip()
        razon_social = match.group(2).strip()
        # Limpiar la razón social
        razon_social = re.sub(r'\s+', ' ', razon_social).strip()
        data['razon_social'] = razon_social
    
    # Si no encontramos con el patrón anterior, intentar otro método
    if not data['ruc']:
        # Buscar RUC de 11 dígitos que no sea repetitivo
        ruc_matches = re.findall(r'\b(\d{11})\b', text)
        for ruc in ruc_matches:
            # Verificar que no sea repetitivo (como 20100030595 que es del banco emisor)
            if len(set(ruc)) > 4 and ruc != '20100030595':  # Filtrar el RUC del banco emisor
                data['ruc'] = ruc
                break
    
    # Buscar razón social si no la encontramos antes
    if not data['razon_social'] and data['ruc']:
        # Buscar texto después del RUC
        ruc_pos = text.find(data['ruc'])
        if ruc_pos != -1:
            text_after_ruc = text[ruc_pos + 11:]  # Después del RUC
            # Buscar la primera línea que parece ser un nombre de empresa
            razon_match = re.search(r'([A-ZÁÉÍÓÚÑÜ][A-ZÁÉÍÓÚÑÜ\s\-\.&/0-9,]+?)(?=\s+(?:AV\.|JR\.|CALLE|CAL\.|MZA\.|PSJ\.))', text_after_ruc, re.IGNORECASE)
            if razon_match:
                razon_social = razon_match.group(1).strip()
                # Limpiar caracteres extraños al final
                razon_social = re.sub(r'\s+', ' ', razon_social).strip()
                if len(razon_social) > 5:
                    data['razon_social'] = razon_social
    
    # Buscar DIRECCIÓN (incluyendo la parte inicial como AV., JR., etc.)
    direccion_patterns = [
        r'((?:AV\.|AVENIDA|JR\.|JIRON|CALLE|CAL\.|MZA\.|PSJ\.|URB\.)\s*[A-ZÁÉÍÓÚÑÜ0-9\s\.\-/,#]+?)(?=\s+\d{4}-\d{2}-\d{2}|\s+SOLES|\s+DÓLARES|\s+BANCO\s+DE\s+LA|$)',
        r'DIRECCIÓN\s+((?:AV\.|AVENIDA|JR\.|JIRON|CALLE|CAL\.|MZA\.|PSJ\.|URB\.)\s*[A-ZÁÉÍÓÚÑÜ0-9\s\.\-/,#]+?)(?=\s+Nº\s+GUÍA|\s+FORMA\s+PAGO|$)',
        r'DIRECCIÓN\s+([A-ZÁÉÍÓÚÑÜ0-9\s\.\-/,#]+?)(?=\s+Nº\s+GUÍA|\s+FORMA\s+PAGO|$)',
    ]
    
    for pattern in direccion_patterns:
        match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
        if match:
            direccion = match.group(1).strip()
            # Limpiar la dirección
            direccion = re.sub(r'\s+', ' ', direccion).strip()
            # Remover texto que no corresponde a dirección
            direccion = re.sub(r'\s*(FECHA\s+EMISIÓN|MONEDA|FORMA\s+PAGO).*', '', direccion)
            data['direccion'] = direccion

    return data

def process_pdf_invoices(pdf_path, output_excel="PRUEBA_BD.xlsx"):
    """Procesa el PDF página por página y extrae datos de cada factura"""
    
    print(f"Procesando archivo: {pdf_path}")
    
    # Extraer texto de cada página
    pages_text = extract_text_from_pdf(pdf_path)
    
    print(f"Se encontraron {len(pages_text)} páginas en el PDF")
    
    # Extraer datos de cada página
    extracted_data = []
    
    for i, page_text in enumerate(pages_text):
        print(f"Procesando página {i + 1}...")
        
        data = extract_invoice_data_from_page(page_text)
        
        # Solo agregar si encontramos al menos RUC o razón social
        if data.get('ruc') or data.get('razon_social'):
            # Agregar número de página para referencia
            data['pagina'] = i + 1
            extracted_data.append(data)
            print(f"  ✓ Extraído: RUC={data.get('ruc', 'N/A')[:8]}... | Razón={data.get('razon_social', 'N/A')[:20]}...")
        else:
            print(f"  ✗ No se pudieron extraer datos de la página {i + 1}")
    
    # Crear DataFrame
    if extracted_data:
        df = pd.DataFrame(extracted_data)
        
        # Reordenar columnas
        columns_order = ['pagina', 'numero_factura', 'ruc', 'razon_social', 'direccion']
        for col in columns_order:
            if col not in df.columns:
                df[col] = ''
        
        df = df[columns_order]
        
        # Guardar en Excel
        df.to_excel(output_excel, index=False, sheet_name='Facturas')
        
        print(f"\n✅ Archivo Excel creado: {output_excel}")
        print(f"📊 Total de registros extraídos: {len(df)}")
        print("\n📋 Vista previa de los primeros 5 registros:")
        pd.set_option('display.max_columns', None)
        pd.set_option('display.width', None)
        pd.set_option('display.max_colwidth', 30)
        print(df.head().to_string(index=False))
        
        return df
    else:
        print("❌ No se pudieron extraer datos de las facturas")
        return None

def main():
    """Función principal"""
    
    # Verificar si se proporcionó la ruta del PDF
    if len(sys.argv) > 1:
        pdf_path = sys.argv[1]
    else:
        # Solicitar la ruta del archivo
        pdf_path = input("Ingresa la ruta completa del archivo PDF: ").strip().strip('"')
    
    # Verificar que el archivo existe
    if not Path(pdf_path).exists():
        print(f"❌ Error: El archivo {pdf_path} no existe")
        return
    
    # Procesar el PDF
    try:
        result = process_pdf_invoices(pdf_path)
        if result is not None:
            print(f"\n🎉 Proceso completado exitosamente!")
            print(f"📁 Archivo guardado como: PRUEBA_BD.xlsx")
            print(f"📄 Se procesaron {len(result)} facturas de 28 páginas del PDF")
        else:
            print("\n⚠️  No se pudieron extraer datos. Verifica el formato del PDF.")
            
    except Exception as e:
        print(f"❌ Error durante el procesamiento: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()

# INSTRUCCIONES DE INSTALACIÓN:
# pip install PyMuPDF pandas openpyxl

# INSTRUCCIONES DE USO:
# python extractor_facturas.py "ruta/a/tu/archivo.pdf"
# o simplemente ejecutar: python extractor_facturas.py