import pdfplumber
import re
import os

class PDFParser:
    def extract_po_from_filename(self, pdf_path):
        """Extrae el número de PO del nombre del archivo"""
        filename = os.path.basename(pdf_path)
        # Buscar patrón "PO #XXXX" en el nombre del archivo
        po_match = re.search(r'PO #(\d+)', filename)
        if po_match:
            return f"PO #{po_match.group(1)}"
        
        # Si no encuentra el patrón, buscar solo números de 4 dígitos
        po_match = re.search(r'(\d{4})', filename)
        if po_match:
            return f"PO #{po_match.group(1)}"
        
        # Si no encuentra nada, devolver el nombre del archivo sin extensión
        return os.path.splitext(filename)[0]
    
    def get_url_data(self, pdf_path):
        """Devuelve un diccionario con el texto del enlace y la ruta del archivo"""
        return {
            'text': self.extract_po_from_filename(pdf_path),
            'path': pdf_path
        }
    
    def extract_data(self, pdf_path):
        rows = []
        print(f"=== PROCESANDO PDF: {pdf_path} ===")

        with pdfplumber.open(pdf_path) as pdf:
            po_no = ""
            for page in pdf.pages:
                text = page.extract_text()
                print(f"--- PÁGINA {pdf.pages.index(page) + 1} ---")
                
                if not po_no:
                    # Buscar PO Number con múltiples patrones
                    po_no_match = re.search(r'Purchase Order No\.\s*(\d+)', text)
                    if not po_no_match:
                        po_no_match = re.search(r'(\d{4})', text)  # Buscar cualquier número de 4 dígitos
                    po_no = po_no_match.group(1) if po_no_match else ""
                    print(f"PO Number encontrado: {po_no}")

                # Buscar solo "Quantity:" (con dos puntos)
                # Buscar cualquier palabra que contenga "tity"
                quantity_pattern = r'\w*tity\w*'
                quantity_matches = list(re.finditer(quantity_pattern, text, re.IGNORECASE))
                
                print(f"Encontradas {len(quantity_matches)} ocurrencias de palabras que contienen 'tity'")
                
                # Si encuentra matches de palabras con "tity", usar la lógica original
                if quantity_matches:
                    for qty_match in quantity_matches:
                        print(f"\n--- PROCESANDO QUANTITY MATCH ---")
                        # Obtener el texto antes de "Quantity"
                        text_before = text[:qty_match.start()]
                        lines_before = text_before.split('\n')
                        
                        # Buscar el estilo en las líneas anteriores
                        estilo = ""
                        color = ""
                        description = ""
                        
                        # Buscar hacia atrás desde "Quantity" hasta encontrar el patrón de estilo
                        for i in range(len(lines_before) - 1, -1, -1):
                            line = lines_before[i].strip()
                            
                            # Buscar patrones de estilo
                            estilo_patterns = [
                                r'(MTJS\d{4})',      # MTJS seguido de 4 números
                                r'(MTS\d{4})',       # MTS seguido de 4 números  
                                r'([A-Z]{3,}\d{3,})' # Patrón general letras+números
                            ]
                            
                            for pattern in estilo_patterns:
                                estilo_match = re.search(pattern, line)
                                if estilo_match:
                                    estilo = estilo_match.group(1)
                                    # El resto de la línea después del estilo es el color
                                    remaining_text = line.replace(estilo, '').strip()
                                    if remaining_text:
                                        color = remaining_text
                                    print(f"ESTILO ENCONTRADO: {estilo}, COLOR: {color}")
                                    break
                            
                            if estilo:
                                # Buscar descripción en líneas cercanas
                                for j in range(max(0, i-2), min(len(lines_before), i+3)):
                                    if j < len(lines_before):
                                        line_desc = lines_before[j].strip()
                                        if ('PurePima' in line_desc or 'Tee' in line_desc or 
                                            'Vintage' in line_desc or 'Pocket' in line_desc):
                                            # Limpiar la descripción
                                            desc_clean = re.sub(r'\d+\.\d+|\d+|Quantity:?|Price:|Unit Price|Discount:', '', line_desc)
                                            desc_clean = desc_clean.strip()
                                            if len(desc_clean) > 5:
                                                description = desc_clean
                                                print(f"DESCRIPCION ENCONTRADA: {description}")
                                                break
                                break
                        
                        print(f"RESULTADO: Estilo='{estilo}', Color='{color}', Description='{description}'")
                        
                        # Solo procesar si encontramos un estilo válido
                        if estilo:
                            print(f"PROCESANDO PRODUCTO CON ESTILO: {estilo}")
                            
                            # Extraer cantidades de la línea después de "Quantity"
                            qty_line_start = qty_match.end()
                            qty_line_end = text.find('\n', qty_line_start)
                            qty_line = text[qty_line_start:qty_line_end] if qty_line_end != -1 else text[qty_line_start:]
                            
                            # Buscar números decimales o enteros
                            qtys = re.findall(r'\d+\.\d+|\d+', qty_line)
                            
                            # Si no encuentra cantidades en la misma línea, buscar en las siguientes
                            if len(qtys) < 4:
                                next_lines = text[qty_line_start:].split('\n')[:3]  # Buscar en las siguientes 3 líneas
                                for next_line in next_lines:
                                    more_qtys = re.findall(r'\d+\.\d+|\d+', next_line)
                                    qtys.extend(more_qtys)
                                    if len(qtys) >= 4:
                                        break
                            
                            print(f"Cantidades encontradas: {qtys}")

                            row = {
                                "PO No": po_no,
                                "Estilo": estilo,
                                "Description": description,
                                "Color": color,
                            }
                            tallas_fijas = ["XS", "S", "M", "L", "XL", "XXL"]
                            total = 0
                            for k, talla in enumerate(tallas_fijas):
                                # Convertir a entero si hay valor, sino dejar vacío
                                if k < len(qtys) and qtys[k]:
                                    try:
                                        valor = int(float(qtys[k]))
                                        row[talla] = valor
                                        total += valor
                                    except ValueError:
                                        row[talla] = qtys[k]
                                else:
                                    row[talla] = ""
                            row["Total"] = total
                            row["URL"] = self.get_url_data(pdf_path)
                            rows.append(row)
                            print(f"ROW AGREGADA: {row}")
                        else:
                            print("NO SE PROCESÓ - Estilo no encontrado")
                
                # Si NO encuentra "Quantity", usar lógica alternativa (buscar directamente por códigos de estilo)
                else:
                    print("=== NO SE ENCONTRÓ 'QUANTITY', USANDO LÓGICA ALTERNATIVA ===")
                    lines = text.split('\n')
                    
                    for i, line in enumerate(lines):
                        # Buscar líneas que contengan códigos de estilo
                        estilo_match = re.search(r'(MTJS\d{4}|MTS\d{4})', line)
                        if estilo_match:
                            estilo = estilo_match.group(1)
                            print(f"ESTILO ENCONTRADO EN LÍNEA {i}: {estilo}")
                            
                            # Extraer color de la misma línea
                            remaining_text = line.replace(estilo, '').strip()
                            color = remaining_text if remaining_text else ""
                            
                            # Buscar descripción en líneas cercanas
                            description = ""
                            for j in range(max(0, i-3), min(len(lines), i+4)):
                                if ('PurePima' in lines[j] or 'Tee' in lines[j] or 
                                    'Vintage' in lines[j] or 'Pocket' in lines[j]):
                                    desc_line = lines[j].strip()
                                    desc_clean = re.sub(r'\d+\.\d+|\d+|Quantity:?|Price:|Unit Price|Discount:', '', desc_line)
                                    desc_clean = desc_clean.strip()
                                    if len(desc_clean) > 5:
                                        description = desc_clean
                                        break
                            
                            # Buscar cantidades en líneas siguientes
                            qtys = []
                            for j in range(i, min(len(lines), i+10)):
                                numbers = re.findall(r'\d+\.\d+', lines[j])
                                if len(numbers) >= 4:
                                    qtys = numbers[:6]
                                    break
                            
                            if qtys:
                                row = {
                                    "PO Number": po_no,
                                    "Estilo": estilo,
                                    "Description": description,
                                    "Color": color,
                                }
                                tallas_fijas = ["XS", "S", "M", "L", "XL", "XXL"]
                                total = 0
                                for k, talla in enumerate(tallas_fijas):
                                    # Convertir a entero si hay valor, sino dejar vacío
                                    if k < len(qtys) and qtys[k]:
                                        try:
                                            valor = int(float(qtys[k]))
                                            row[talla] = valor
                                            total += valor
                                        except ValueError:
                                            row[talla] = qtys[k]
                                    else:
                                        row[talla] = ""
                                row["Total"] = total
                                row["URL"] = self.get_url_data(pdf_path)
                                rows.append(row)
                                print(f"ROW AGREGADA (LÓGICA ALTERNATIVA): {row}")

        print(f"=== TOTAL DE ROWS EXTRAÍDAS: {len(rows)} ===")
        return rows