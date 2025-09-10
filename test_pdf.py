import pdfplumber
import pandas as pd
import re
import traceback
from typing import List, Dict, Any, Optional
import logging
from datetime import datetime
from typing import List, Dict, Any, Optional, Tuple
# Configurar logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class BalanceExtractorEnhanced:
    def __init__(self):
        self.columns = ['CODIGO', 'NOMBRE', 'SALDO ANTERIOR', 'CARGOS', 'ABONOS', 'SALDO ACTUAL']
        self.extracted_date = None
    
    def _extract_date_from_pdf(self, pdf) -> str:
        # Patrón específico para el título del balance
        
        title_date_patterns = [
            # Patrón principal: "BALANCE DE COMPROBACION DIARIO EN MONEDA NACIONAL AL DIA DD/MM/YYYY"
            r'BALANCE\s+DE\s+COMPROBACION\s+DIARIO\s+EN\s+MONEDA\s+NACIONAL\s+AL\s+DIA\s+(\d{1,2})/(\d{1,2})/(\d{4})',
            r'BALANCE\s+DE\s+COMPROBACION\s+.*?AL\s+DIA\s+(\d{1,2})/(\d{1,2})/(\d{4})',
            r'BALANCE\s+DE\s+COMPROBACION\s+.*?AL\s+(\d{1,2})/(\d{1,2})/(\d{4})',
            # Variaciones del patrón
            r'BALANCE.*?COMPROBACION.*?AL\s+DIA\s+(\d{1,2})/(\d{1,2})/(\d{4})',
            r'BALANCE.*?COMPROBACION.*?AL\s+(\d{1,2})/(\d{1,2})/(\d{4})',
        ]
        
        
        # Patrones alternativos con formato de fecha con puntos o guiones
        alternative_patterns = [
            r'BALANCE\s+DE\s+COMPROBACION\s+DIARIO\s+EN\s+MONEDA\s+NACIONAL\s+AL\s+DIA\s+(\d{1,2})\.(\d{1,2})\.(\d{4})',
            r'BALANCE\s+DE\s+COMPROBACION\s+DIARIO\s+EN\s+MONEDA\s+NACIONAL\s+AL\s+DIA\s+(\d{1,2})-(\d{1,2})-(\d{4})',
            r'BALANCE.*?COMPROBACION.*?AL\s+DIA\s+(\d{1,2})\.(\d{1,2})\.(\d{4})',
            r'BALANCE.*?COMPROBACION.*?AL\s+DIA\s+(\d{1,2})-(\d{1,2})-(\d{4})',
        ]
        
        # Combinar todos los patrones
        all_patterns = title_date_patterns + alternative_patterns
        
        # Buscar en las primeras 3 páginas (principalmente la primera)
        for page_num in range(min(3, len(pdf.pages))):
            page = pdf.pages[page_num]
            text = page.extract_text()
            
            if not text:
                continue
            
            # Limpiar texto para mejorar la búsqueda
            clean_text = ' '.join(text.upper().split())
            logger.debug(f"Buscando fecha en página {page_num + 1}")
            logger.debug(f"Texto limpio: {clean_text[:300]}...")
            
            # Buscar el patrón específico del título
            for i, pattern in enumerate(all_patterns):
                matches = re.finditer(pattern, clean_text, re.IGNORECASE)
                
                for match in matches:
                    try:
                        groups = match.groups()
                        
                        if len(groups) == 3:
                            day = groups[0].zfill(2)
                            month = groups[1].zfill(2)
                            year = groups[2]
                            formatted_date = f"{day}/{month}/{year}"
                            
                            logger.info(f"Fecha encontrada en el título (patrón {i+1}): {formatted_date}")
                            logger.info(f"Texto completo del match: {match.group(0)}")
                            return formatted_date
                    
                    except Exception as e:
                        logger.debug(f"Error procesando match de fecha: {e}")
                        continue
        # Si no encuentra la fecha en el título principal, buscar patrones más genéricos pero priorizando fechas con formato DD/MM/YYYY
        logger.warning("No se encontró la fecha en el título principal, buscando patrones alternativos...")
        
        generic_patterns = [
            r'AL\s+(\d{1,2})/(\d{1,2})/(\d{4})',
            r'(\d{1,2})/(\d{1,2})/(\d{4})',
            r'(\d{1,2})\.(\d{1,2})\.(\d{4})',
            r'(\d{1,2})-(\d{1,2})-(\d{4})',
        ]
        
        for page_num in range(min(3, len(pdf.pages))):
            page = pdf.pages[page_num]
            text = page.extract_text()
            
            if not text:
                continue
            
            clean_text = ' '.join(text.upper().split())
            
            for pattern in generic_patterns:
                matches = re.finditer(pattern, clean_text)
                
                for match in matches:
                    try:
                        groups = match.groups()
                        
                        if len(groups) == 3:
                            day = int(groups[0])
                            month = int(groups[1])
                            year = int(groups[2])
                            
                            # Validar que sea una fecha válida
                            if 1 <= day <= 31 and 1 <= month <= 12 and 2020 <= year <= 2030:
                                formatted_date = f"{day:02d}/{month:02d}/{year}"
                                logger.info(f"Fecha alternativa encontrada: {formatted_date}")
                                return formatted_date
                    
                    except Exception as e:
                        logger.debug(f"Error procesando fecha alternativa: {e}")
                        continue
        
        # Si no encuentra fecha, usar fecha actual
        current_date = datetime.now().strftime("%d/%m/%Y")
        logger.warning(f"No se encontró fecha en el PDF, usando fecha actual: {current_date}")
        return current_date
    
    def get_excel_filename(self, original_pdf_path: str = None) -> str:
        """
        Genera el nombre del archivo Excel basado en la fecha extraída
        """
        if self.extracted_date:
            try:
                # Convertir fecha a formato para nombre de archivo (YYYY-MM-DD)
                date_parts = self.extracted_date.split('/')
                if len(date_parts) == 3:
                    day, month, year = date_parts
                    filename_date = f"{year}-{month}-{day}"
                    return f"Balance_Comprobacion_{filename_date}.xlsx"
            except:
                pass
        
        # Fallback si no hay fecha extraída
        return f"Balance_Comprobacion_{datetime.now().strftime('%Y-%m-%d')}.xlsx"
        
    def extract_balance_data(self, pdf_path: str) -> List[Dict[str, Any]]:
        """
        Extrae datos del balance de comprobación desde un PDF
        """
        all_data = []
        
        try:
            with pdfplumber.open(pdf_path) as pdf:
                logger.info(f"Procesando PDF con {len(pdf.pages)} páginas")
                self.extracted_date = self._extract_date_from_pdf(pdf)
                logger.info(f"Fecha extraída del PDF: {self.extracted_date}")
                
                for page_num, page in enumerate(pdf.pages, 1):
                    logger.info(f"Procesando página {page_num}")
                    
                    
                    # Extraer texto de la página
                    text = page.extract_text()
                    if not text:
                        logger.warning(f"No se pudo extraer texto de la página {page_num}")
                        continue
                    
                    # Procesar los datos de esta página
                    page_data = self._parse_page_data(text)
                    all_data.extend(page_data)
                    
                    logger.info(f"Extraídas {len(page_data)} filas de la página {page_num}")
                    
        except Exception as e:
            logger.error(f"Error al procesar el PDF: {e}")
            raise
            
        logger.info(f"Total de filas extraídas: {len(all_data)}")
        return all_data
    
    def _parse_page_data(self, text: str) -> List[Dict[str, Any]]:
        """
        Parsea los datos de una página específica con lógica mejorada
        """
        data_rows = []
        lines = text.split('\n')
        
        # Buscar líneas que contienen datos de cuentas
        for i, line in enumerate(lines):
            line = line.strip()
            if not line:
                continue
            
            # Verificar si la línea contiene datos de cuenta
            if self._is_data_line(line):
                parsed_row = self._parse_data_line_enhanced(line)
                if parsed_row:
                    data_rows.append(parsed_row)
        
        return data_rows
    
    def _is_data_line(self, line: str) -> bool:
        """
        Determina si una línea contiene datos de cuenta
        Versión mejorada para detectar líneas sin nombre
        """
        # Limpiar la línea de espacios extras
        clean_line = ' '.join(line.split())
        
        # La línea debe empezar con dígitos (código de cuenta)
        if not re.match(r'^\d+', clean_line):
            return False
        
        # Buscar patrones de números decimales (montos) incluyendo los que terminan en CR
        decimal_numbers = re.findall(r'\d{1,3}(?:\s?\d{3})*\s?\d{3}\.\d{2}(?:\s*CR)?', clean_line)
        
        # Para que sea una línea válida debe tener:
        # 1. Al menos 2 números decimales (mínimo saldo anterior y saldo actual)
        # 2. Longitud mínima razonable
        if len(decimal_numbers) >= 2 and len(clean_line) > 20:
            logger.debug(f"Línea válida detectada: {clean_line[:50]}...")
            return True
        
        return False
    
    def _parse_data_line_enhanced(self, line: str) -> Optional[Dict[str, Any]]:
        """
        Parsea una línea de datos de cuenta con manejo robusto de casos sin nombre
        """
        try:
            # Limpiar la línea
            clean_line = ' '.join(line.split())
            logger.debug(f"Procesando línea: {clean_line}")
            
            # Extraer código de cuenta (primeros dígitos)
            codigo_match = re.match(r'^(\d+)', clean_line)
            if not codigo_match:
                return None
            
            codigo = codigo_match.group(1)
            
            # Encontrar todos los números con formato de montos
            # Buscar patrones como: 19 380 727 198.64 o 380 727 198.64 CR
            number_pattern = r'(\d{1,3}(?:\s\d{3})*\s\d{3}\.\d{2}(?:\s*CR)?)'
            numbers = re.findall(number_pattern, clean_line)
            
            if not numbers:
                logger.debug(f"No se encontraron números válidos en: {clean_line}")
                return None
            
            # Convertir números a formato string con comas para miles preservando CR
            formatted_numbers = []
            numeric_values = []  # Para cálculos de lógica
            
            for num in numbers:
                # Verificar si tiene CR al final
                has_cr = 'CR' in num
                
                # Limpiar espacios internos y CR para obtener valor numérico
                clean_num = num.replace(' ', '').replace('CR', '')
                try:
                    numeric_val = float(clean_num)
                    # Si tiene CR, es un valor negativo para cálculos
                    if has_cr:
                        numeric_values.append(-numeric_val)
                    else:
                        numeric_values.append(numeric_val)
                    
                    # Formatear como string con comas para miles, preservando CR
                    formatted_str = f"{numeric_val:,.2f}"
                    if has_cr:
                        formatted_str += " CR"
                    formatted_numbers.append(formatted_str)
                except ValueError:
                    logger.debug(f"Error al convertir número: {num}")
                    continue
            
            logger.debug(f"Números formateados: {formatted_numbers}")
            
            # Extraer el nombre (texto entre código y primer número)
            nombre = ""
            if len(formatted_numbers) > 0:
                # Buscar el primer número en la línea original
                first_number_str = numbers[0]
                first_number_pos = clean_line.find(first_number_str)
                
                if first_number_pos > len(codigo):
                    # Extraer texto entre código y primer número
                    nombre_section = clean_line[len(codigo):first_number_pos].strip()
                    # Limpiar caracteres extraños
                    nombre = re.sub(r'[^\w\s\-\.\(\)\/]', ' ', nombre_section).strip()
                    nombre = ' '.join(nombre.split())  # Normalizar espacios
            
            # Si no hay nombre, usar uno descriptivo basado en el código
            if not nombre or len(nombre) < 2:
                if codigo.startswith('1'):
                    nombre = f"-"
                elif codigo.startswith('2'):
                    nombre = f"PASIVO_{codigo}"
                elif codigo.startswith('3'):
                    nombre = f"PATRIMONIO_{codigo}"
                elif codigo.startswith('4'):
                    nombre = f"GASTO_{codigo}"
                elif codigo.startswith('5'):
                    nombre = f"INGRESO_{codigo}"
                else:
                    nombre = f"CUENTA_{codigo}"
            
            # Asignar valores según la cantidad de números encontrados (usando strings formateados)
            saldo_anterior = "0.00"
            cargos = "0.00"
            abonos = "0.00"
            saldo_actual = "0.00"
            
            if len(formatted_numbers) == 1:
                # Solo saldo actual
                saldo_actual = formatted_numbers[0]
            elif len(formatted_numbers) == 2:
                # Saldo anterior y saldo actual (sin movimientos)
                saldo_anterior = formatted_numbers[0]
                saldo_actual = formatted_numbers[1]
            elif len(formatted_numbers) == 3:
                # Saldo anterior, un movimiento (cargo o abono), y saldo actual
                saldo_anterior = formatted_numbers[0]
                # Determinar si es cargo o abono por la diferencia (usando valores numéricos para la lógica)
                movimiento_str = formatted_numbers[1]
                saldo_actual = formatted_numbers[2]
                
                # Si saldo_actual > saldo_anterior, probablemente es un cargo
                if numeric_values[2] > numeric_values[0]:
                    cargos = movimiento_str
                else:
                    abonos = movimiento_str
            elif len(formatted_numbers) >= 4:
                # Formato completo: saldo anterior, cargos, abonos, saldo actual
                saldo_anterior = formatted_numbers[0]
                cargos = formatted_numbers[1]
                abonos = formatted_numbers[2]
                saldo_actual = formatted_numbers[3]
            
            result = {
                'CODIGO': codigo,
                'NOMBRE': nombre,
                'SALDO_ANTERIOR': saldo_anterior,
                'CARGOS': cargos,
                'ABONOS': abonos,
                'SALDO_ACTUAL': saldo_actual
            }
            
            logger.info(f"Línea procesada exitosamente: {codigo} - {nombre} - SA:{saldo_anterior} C:{cargos} A:{abonos} SAct:{saldo_actual}")
            return result
            
        except Exception as e:
            logger.error(f"Error procesando línea: {line[:50]}... - Error: {e}")
            return None
    
    def _extract_account_name(self, line: str, codigo: str, first_number: str) -> str:
        """
        Extrae el nombre de la cuenta entre el código y el primer número
        """
        try:
            # Remover el código del inicio
            after_codigo = line[len(codigo):].strip()
            
            # Encontrar la posición del primer número
            first_num_pattern = re.escape(first_number)
            match = re.search(first_num_pattern, after_codigo)
            
            if match:
                # Extraer texto hasta el primer número
                nombre = after_codigo[:match.start()].strip()
            else:
                # Si no encontramos el número, usar una heurística
                # Buscar donde empiezan los números grandes
                words = after_codigo.split()
                nombre_parts = []
                
                for word in words:
                    # Si encontramos un número grande o con decimales, paramos
                    if (word.isdigit() and len(word) > 4) or re.match(r'[\d,]+\.\d{2}', word):
                        break
                    nombre_parts.append(word)
                
                nombre = ' '.join(nombre_parts).strip()
            
            return nombre
            
        except Exception as e:
            logger.debug(f"Error extrayendo nombre: {e}")
            return ""
    
    def save_to_excel(self, data: List[Dict[str, Any]], output_path: str):
        
        try:
            if not data:
                logger.warning("No hay datos para guardar")
                return
            
            # Crear DataFrame
            df = pd.DataFrame(data)
            
            # Limpiar y validar datos
            df = self._clean_and_validate_data(df)
            
            # Guardar en Excel con formato
            with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
                # Escribir los datos empezando desde la fila 2 (índice 1)
                df.to_excel(writer, sheet_name='Balance_Comprobacion', index=False, startrow=1)
                
                # Obtener workbook y worksheet para formatear
                workbook = writer.book
                worksheet = writer.sheets['Balance_Comprobacion']
                
                if self.extracted_date:
                    # Formato para la celda de fecha combinada
                    date_format = workbook.add_format({
                        'bold': True,
                        'font_size': 14,
                        'bg_color': '#E6F3FF',
                        'border': 1,
                        'align': 'center',
                        'valign': 'vcenter'
                    })
                    
                    # Combinar celdas A1 a F1
                    worksheet.merge_range('A1:F1', f'BALANCE DE COMPROBACIÓN - FECHA: {self.extracted_date}', date_format)
                
                # Formatos para headers (ahora en la fila 2)
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': True,
                    'valign': 'top',
                    'fg_color': '#D7E4BC',
                    'border': 1,
                    'align': 'center'
                })
                
                # Formato para números
                money_format = workbook.add_format({
                    'num_format': '#,##0.00',
                    'align': 'right'
                })
                
                # Escribir headers manualmente en la fila 2 (índice 1) con formato
                headers = ['CODIGO', 'NOMBRE', 'SALDO_ANTERIOR', 'CARGOS', 'ABONOS', 'SALDO_ACTUAL']
                for col_num, header in enumerate(headers):
                    worksheet.write(1, col_num, header, header_format)
                
                # Formatear columnas numéricas (columnas C, D, E, F que corresponden a índices 2, 3, 4, 5)
                for col_num in [2, 3, 4, 5]:  # Columnas de montos
                    col_letter = chr(65 + col_num)  # A=65, B=66, etc.
                    worksheet.set_column(f'{col_letter}:{col_letter}', 15, money_format)
                
                # Ajustar ancho de columnas
                worksheet.set_column('A:A', 12)  # CODIGO
                worksheet.set_column('B:B', 35)  # NOMBRE
                
                # Ajustar altura de la primera fila para que se vea mejor la fecha
                worksheet.set_row(0, 25)  # Fila 1 (índice 0) con altura 25
                
                # Agregar hoja de resumen
                self._add_summary_sheet(writer, df)
            
            logger.info(f"Excel creado exitosamente: {output_path}")
            
            # Mostrar resumen en consola
            print("\n📊 RESUMEN DE EXTRACCIÓN")
            print("=" * 50)
            print(f"Total de cuentas procesadas: {len(df)}")
            
            # Función para sumar strings con formato de montos
            def sum_money_strings(series):
                total = 0.0
                for val in series:
                    try:
                        # Manejar valores con CR (crédito/negativo)
                        if 'CR' in str(val):
                            clean_val = str(val).replace(',', '').replace(' CR', '').replace('CR', '')
                            total -= float(clean_val)
                        else:
                            total += float(str(val).replace(',', ''))
                    except:
                        continue
                return total
            
            suma_sa = sum_money_strings(df['SALDO_ANTERIOR'])
            suma_cargos = sum_money_strings(df['CARGOS'])
            suma_abonos = sum_money_strings(df['ABONOS'])
            suma_sact = sum_money_strings(df['SALDO_ACTUAL'])
            
            print(f"Suma saldos anteriores: {suma_sa:,.2f}")
            print(f"Suma total cargos: {suma_cargos:,.2f}")
            print(f"Suma total abonos: {suma_abonos:,.2f}")
            print(f"Suma saldos actuales: {suma_sact:,.2f}")
            print(f"\n📁 Archivo generado: {output_path}")
            
            # Mostrar muestra de datos
            print(f"\n📋 Muestra de datos extraídos:")
            print(df.head(10).to_string(index=False, max_cols=6))
        except Exception as e:
            logger.error(f"Error al crear Excel: {e}")
            raise
    
    def _clean_and_validate_data(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Limpia y valida los datos extraídos con manejo robusto de datos faltantes
        """
        print(f"\n🔧 Limpiando y validando {len(df)} registros...")
        
        # Remover filas con código vacío
        initial_count = len(df)
        df = df[df['CODIGO'].notna() & (df['CODIGO'] != '')]
        if len(df) < initial_count:
            print(f"   ⚠️ Removidas {initial_count - len(df)} filas sin código")
        
        # Limpiar nombres (remover caracteres extraños)
        df['NOMBRE'] = df['NOMBRE'].astype(str).str.strip()
        df['NOMBRE'] = df['NOMBRE'].str.replace(r'\s+', ' ', regex=True)  # Múltiples espacios a uno
        
        # Detectar y reportar nombres automáticos (cuentas sin nombre)
        auto_names = df[df['NOMBRE'].str.contains(r'^CUENTA_\d+$', na=False)]
        if len(auto_names) > 0:
            print(f"   ⚠️ {len(auto_names)} cuentas sin nombre detectadas (usando placeholders)")
        
        # Limpiar y validar columnas de montos como strings
        numeric_cols = ['SALDO_ANTERIOR', 'CARGOS', 'ABONOS', 'SALDO_ACTUAL']
        for col in numeric_cols:
            # Asegurar que sean strings y limpiar valores problemáticos
            df[col] = df[col].astype(str)
            # Reemplazar 'nan', 'None', etc. con '0.00'
            df[col] = df[col].replace(['nan', 'None', ''], '0.00')
            # Validar formato de montos (incluyendo CR)
            invalid_format = ~df[col].str.match(r'^\d{1,3}(,\d{3})*\.\d{2}(?:\s*CR)?$|^0\.00$')
            if invalid_format.any():
                print(f"   ⚠️ {invalid_format.sum()} valores con formato incorrecto en '{col}' corregidos a '0.00'")
                df.loc[invalid_format, col] = '0.00'
        
        # Detectar cuentas con datos incompletos (todos los valores son "0.00")
        zero_data = df[(df['SALDO_ANTERIOR'] == '0.00') & (df['CARGOS'] == '0.00') & 
                      (df['ABONOS'] == '0.00') & (df['SALDO_ACTUAL'] == '0.00')]
        if len(zero_data) > 0:
            print(f"   ⚠️ {len(zero_data)} cuentas con todos los valores en 0 (posibles datos incompletos)")
        
        # Función auxiliar para convertir string con formato a float para validaciones
        def string_to_float(val_str):
            try:
                # Manejar valores con CR (crédito/negativo)
                if 'CR' in str(val_str):
                    clean_val = str(val_str).replace(',', '').replace(' CR', '').replace('CR', '')
                    return -float(clean_val)
                else:
                    return float(str(val_str).replace(',', ''))
            except:
                return 0.0
        
        # Detectar posibles errores de balance (convertir temporalmente para cálculo)
        df_temp = df.copy()
        for col in numeric_cols:
            df_temp[col + '_num'] = df_temp[col].apply(string_to_float)
        
        balance_errors = df_temp[abs(df_temp['SALDO_ANTERIOR_num'] + df_temp['CARGOS_num'] - 
                                   df_temp['ABONOS_num'] - df_temp['SALDO_ACTUAL_num']) > 0.01]
        if len(balance_errors) > 0:
            print(f"   ⚠️ {len(balance_errors)} cuentas con posibles errores de balance")
        
        # Remover duplicados por código
        duplicates = df.duplicated(subset=['CODIGO'], keep=False)
        if duplicates.any():
            print(f"   ⚠️ {duplicates.sum()} códigos duplicados detectados - manteniendo el primero")
            df = df.drop_duplicates(subset=['CODIGO'], keep='first')
        
        # Ordenar por código
        df = df.sort_values('CODIGO').reset_index(drop=True)
        
        print(f"   ✅ Validación completada: {len(df)} registros válidos")
        return df
    
    def _add_summary_sheet(self, writer: pd.ExcelWriter, df: pd.DataFrame):
        """
        Agrega hoja de resumen con totales y validaciones
        """
        # Función auxiliar para sumar strings con formato de montos
        def sum_money_strings(series):
            total = 0.0
            for val in series:
                try:
                    # Manejar valores con CR (crédito/negativo)
                    if 'CR' in str(val):
                        clean_val = str(val).replace(',', '').replace(' CR', '').replace('CR', '')
                        total -= float(clean_val)
                    else:
                        total += float(str(val).replace(',', ''))
                except:
                    continue
            return total
        
        # Calcular sumas
        suma_sa = sum_money_strings(df['SALDO_ANTERIOR'])
        suma_cargos = sum_money_strings(df['CARGOS'])
        suma_abonos = sum_money_strings(df['ABONOS'])
        suma_sact = sum_money_strings(df['SALDO_ACTUAL'])
        
        # Contar cuentas con saldo mayor a 1M
        cuentas_1m = 0
        for val in df['SALDO_ACTUAL']:
            try:
                # Manejar valores con CR
                if 'CR' in str(val):
                    clean_val = str(val).replace(',', '').replace(' CR', '').replace('CR', '')
                    saldo_val = -float(clean_val)
                else:
                    saldo_val = float(str(val).replace(',', ''))
                if abs(saldo_val) > 1000000:
                    cuentas_1m += 1
            except:
                continue
        
        # Contar cuentas con movimientos
        cuentas_movimientos = 0
        for idx in df.index:
            try:
                # Manejar CR en cargos
                cargos_str = str(df.loc[idx, 'CARGOS'])
                if 'CR' in cargos_str:
                    clean_val = cargos_str.replace(',', '').replace(' CR', '').replace('CR', '')
                    cargos = -float(clean_val)
                else:
                    cargos = float(cargos_str.replace(',', ''))
                
                # Manejar CR en abonos
                abonos_str = str(df.loc[idx, 'ABONOS'])
                if 'CR' in abonos_str:
                    clean_val = abonos_str.replace(',', '').replace(' CR', '').replace('CR', '')
                    abonos = -float(clean_val)
                else:
                    abonos = float(abonos_str.replace(',', ''))
                
                if abs(cargos) > 0 or abs(abonos) > 0:
                    cuentas_movimientos += 1
            except:
                continue
        
        # Crear datos de resumen
        summary_data = {
            'Concepto': [
                'Total de Cuentas',
                'Suma Saldos Anteriores',
                'Suma Total Cargos',
                'Suma Total Abonos',
                'Suma Saldos Actuales',
                'Diferencia (Actual - Anterior)',
                'Validación Balance',
                'Cuentas con Saldo Mayor a 1M',
                'Cuentas con Movimientos'
            ],
            'Valor': [
                len(df),
                f"{suma_sa:,.2f}",
                f"{suma_cargos:,.2f}",
                f"{suma_abonos:,.2f}",
                f"{suma_sact:,.2f}",
                f"{suma_sact - suma_sa:,.2f}",
                'OK' if abs((suma_sa + suma_cargos - suma_abonos) - suma_sact) < 1 else 'REVISAR',
                cuentas_1m,
                cuentas_movimientos
            ]
        }
        
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='Resumen', index=False)

def main():
    """
    Función principal mejorada
    """
    # Configuración
    PDF_PATH = "test.pdf"
    EXCEL_OUTPUT = "test_3.xlsx"
    
    print("🏦 EXTRACTOR MEJORADO - Banco de la Nación")
    print("=" * 55)
    
    try:
        # Crear extractor mejorado
        extractor = BalanceExtractorEnhanced()
        
        # Extraer datos
        print(f"📖 Procesando archivo: {PDF_PATH}")
        data = extractor.extract_balance_data(PDF_PATH)
        
        if not data:
            print("⚠️  No se encontraron datos válidos en el PDF")
            print("   Verifica que el PDF contiene un balance de comprobación válido")
            return
        
        print(f"✅ Extracción completada: {len(data)} registros encontrados")
        
        # Guardar en Excel
        print(f"💾 Generando archivo Excel...")
        extractor.save_to_excel(data, EXCEL_OUTPUT)
        
        print("\n🎉 ¡PROCESO COMPLETADO EXITOSAMENTE!")
        
    except FileNotFoundError:
        print(f"❌ Error: No se encontró el archivo '{PDF_PATH}'")
        print("   📋 Coloca el archivo PDF en la misma carpeta que este script")
    except Exception as e:
        print(f"❌ Error durante el proceso: {e}")
        logger.error(f"Error en main: {e}")
        import traceback
        print(f"   🔧 Detalles técnicos: {traceback.format_exc()}")

if __name__ == "__main__":
    main()