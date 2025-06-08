import xml.etree.ElementTree as ET
import re
import os
from datetime import datetime
from pathlib import Path
from typing import Optional, Dict, Any, List
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk

try:
    from weasyprint import HTML as WeasyHTML
    WEASYPRINT_AVAILABLE = True
except ImportError:
    WEASYPRINT_AVAILABLE = False

try:
    import barcode
    from barcode.writer import ImageWriter
    import base64
    from io import BytesIO
    BARCODE_AVAILABLE = True
except ImportError:
    BARCODE_AVAILABLE = False

# Configuração do CustomTkinter
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

# Namespace usado no XML da NF-e
NAMESPACES = {'ns': 'http://www.portalfiscal.inf.br/nfe'}

class ProcessadorNFe:
    """Classe responsável pelo processamento de arquivos XML de NF-e"""
    
    def __init__(self):
        self.xml_file_path = None
        self.html_template_path = None
        self.output_directory = None
    
    def get_xml_value(self, element: ET.Element, path: str, default: str = '', 
                     is_currency: bool = False, unit_from_sibling: Optional[str] = None, 
                     sibling_unit_tag: Optional[str] = None) -> str:
        """
        Busca um valor em um elemento XML usando um caminho XPath.
        
        Args:
            element: Elemento XML base
            path: Caminho XPath ou lista de caminhos de fallback
            default: Valor padrão se não encontrado
            is_currency: Se deve formatar como moeda
            unit_from_sibling: Tag do elemento irmão para obter unidade
            sibling_unit_tag: Tag específica da unidade
            
        Returns:
            Valor encontrado ou default
        """
        try:
            if isinstance(path, str):
                path_parts = path.split('/')
                processed_parts = []
                for part in path_parts:
                    if ':' not in part and '@' not in part:
                        processed_parts.append(f'ns:{part}')
                    else:
                        processed_parts.append(part)
                full_path = '/'.join(processed_parts)
            else:
                full_path = []
                for p_str in path:
                    path_parts = p_str.split('/')
                    processed_parts = []
                    for part in path_parts:
                        if ':' not in part and '@' not in part:
                            processed_parts.append(f'ns:{part}')
                        else:
                            processed_parts.append(part)
                    full_path.append('/'.join(processed_parts))

            found_element = None
            if isinstance(full_path, list):
                for p in full_path:
                    found_element = element.find(p, NAMESPACES)
                    if found_element is not None and found_element.text:
                        break
            else:
                found_element = element.find(full_path, NAMESPACES)

            if found_element is not None and found_element.text:
                value = found_element.text.strip()
                
                if is_currency:
                    try:
                        num_value = float(value)
                        value = f'R$ {num_value:,.2f}'.replace(',', 'TEMP').replace('.', ',').replace('TEMP', '.')
                    except ValueError:
                        pass
                
                if unit_from_sibling and sibling_unit_tag:
                    unit_element = element.find(f'ns:{unit_from_sibling}/ns:{sibling_unit_tag}', NAMESPACES)
                    if unit_element is not None and unit_element.text:
                        value = f'{value} {unit_element.text.strip()}'
                elif unit_from_sibling and not sibling_unit_tag:
                    value = f'{value} {unit_from_sibling}'

                return value
            return default
        except Exception:
            return default

    def format_datetime_field(self, dt_str: str, output_format: str = '%d/%m/%Y') -> str:
        """Formata string de data/hora (ISO 8601 com timezone) para o formato desejado."""
        if not dt_str:
            return ''
        try:
            dt_str_no_tz = re.sub(r'[+-]\d{2}:\d{2}$', '', dt_str)
            dt_obj = datetime.fromisoformat(dt_str_no_tz)
            return dt_obj.strftime(output_format)
        except ValueError:
            return dt_str

    def format_items_html(self, root_nfe: ET.Element, max_items_per_page: int = 15) -> str:
        """Formata os itens da NF-e como linhas de tabela HTML com paginação."""
        items_html_list = []
        infNFe = root_nfe.find('ns:infNFe', NAMESPACES)
        if infNFe is None:
            return ''

        det_elements = infNFe.findall('ns:det', NAMESPACES)
        
        # Se há poucos itens, processa normalmente
        if len(det_elements) <= max_items_per_page:
            for det_element in det_elements:
                prod_element = det_element.find('ns:prod', NAMESPACES)
                imposto_element = det_element.find('ns:imposto', NAMESPACES)
                
                if prod_element is None or imposto_element is None:
                    continue

                # Busca valores de ICMS - expandir para incluir todos os CSTs
                v_bc_icms_paths = [
                    'ICMS/ICMS00/vBC', 'ICMS/ICMS10/vBC', 'ICMS/ICMS20/vBC', 'ICMS/ICMS30/vBC',
                    'ICMS/ICMS40/vBC', 'ICMS/ICMS41/vBC', 'ICMS/ICMS50/vBC', 'ICMS/ICMS51/vBC',
                    'ICMS/ICMS60/vBC', 'ICMS/ICMS70/vBC', 'ICMS/ICMS90/vBC'
                ]
                v_bc_icms = 'R$ 0,00'
                for path in v_bc_icms_paths:
                    temp_v_bc_icms = self.get_xml_value(imposto_element, path, is_currency=True)
                    if temp_v_bc_icms:
                        v_bc_icms = temp_v_bc_icms
                        break

                # Percentual ICMS - buscar valor bruto para formatação correta
                p_icms_paths = [
                    'ICMS/ICMS00/pICMS', 'ICMS/ICMS10/pICMS', 'ICMS/ICMS20/pICMS', 'ICMS/ICMS30/pICMS',
                    'ICMS/ICMS40/pICMS', 'ICMS/ICMS41/pICMS', 'ICMS/ICMS50/pICMS', 'ICMS/ICMS51/pICMS',
                    'ICMS/ICMS60/pICMS', 'ICMS/ICMS70/pICMS', 'ICMS/ICMS90/pICMS'
                ]
                p_icms_raw = ''
                for path in p_icms_paths:
                    p_icms_raw = self.get_xml_value(imposto_element, path)
                    if p_icms_raw:
                        break
                
                # Formatar percentual ICMS
                p_icms_formatted = '0%'
                if p_icms_raw:
                    try:
                        perc_value = float(p_icms_raw)
                        p_icms_formatted = f"{perc_value:.1f}%"
                    except ValueError:
                        p_icms_formatted = "0%"

                # Valor ICMS
                v_icms_paths = [
                    'ICMS/ICMS00/vICMS', 'ICMS/ICMS10/vICMS', 'ICMS/ICMS20/vICMS', 'ICMS/ICMS30/vICMS',
                    'ICMS/ICMS40/vICMS', 'ICMS/ICMS41/vICMS', 'ICMS/ICMS50/vICMS', 'ICMS/ICMS51/vICMS',
                    'ICMS/ICMS60/vICMS', 'ICMS/ICMS70/vICMS', 'ICMS/ICMS90/vICMS'
                ]
                v_icms = 'R$ 0,00'
                for path in v_icms_paths:
                    temp_v_icms = self.get_xml_value(imposto_element, path, is_currency=True)
                    if temp_v_icms:
                        v_icms = temp_v_icms
                        break

                v_ipi = self.get_xml_value(imposto_element, 'IPI/IPITrib/vIPI', default='', is_currency=True)
                if not v_ipi:
                    v_ipi = "R$ 0,00"
                
                # Percentual IPI - buscar valor bruto para formatação correta
                p_ipi_raw = self.get_xml_value(imposto_element, 'IPI/IPITrib/pIPI', default='0.00')
                p_ipi_formatted = '0%'
                if p_ipi_raw:
                    try:
                        perc_value = float(p_ipi_raw)
                        p_ipi_formatted = f"{perc_value:.1f}%"
                    except ValueError:
                        p_ipi_formatted = "0%"

                item_row = f"""
                <tr>
                    <td style="text-align:center; padding: 4px; min-height: 25px; white-space: nowrap; vertical-align: middle;">{self.get_xml_value(prod_element, 'cProd')}</td>
                    <td style="padding: 4px; min-height: 25px; vertical-align: middle;">{self.get_xml_value(prod_element, 'xProd')}</td>
                    <td style="text-align:center; padding: 4px; min-height: 25px; white-space: nowrap; vertical-align: middle;">{self.get_xml_value(prod_element, 'NCM')}</td>
                    <td style="text-align:center; padding: 4px; min-height: 25px; white-space: nowrap; vertical-align: middle;">{self.get_xml_value(prod_element, 'CFOP')}</td>
                    <td style="text-align:center; padding: 4px; min-height: 25px; white-space: nowrap; vertical-align: middle;">{self.get_xml_value(prod_element, 'uCom')}</td>
                    <td style="text-align:center; padding: 4px; min-height: 25px; white-space: nowrap; vertical-align: middle;">{float(self.get_xml_value(prod_element, 'qCom', default='0')):,.4f}</td>
                    <td style="text-align:center; padding: 4px; min-height: 25px; white-space: nowrap; vertical-align: middle;">{self.get_xml_value(prod_element, 'vUnCom', is_currency=True)}</td>
                    <td style="text-align:center; padding: 4px; min-height: 25px; white-space: nowrap; vertical-align: middle;">{self.get_xml_value(prod_element, 'vProd', is_currency=True)}</td>
                    <td style="text-align:center; padding: 4px; min-height: 25px; white-space: nowrap; vertical-align: middle;">{v_bc_icms}</td>
                    <td style="text-align:center; padding: 4px; min-height: 25px; white-space: nowrap; vertical-align: middle;">{v_icms}</td>
                    <td style="text-align:center; padding: 4px; min-height: 25px; white-space: nowrap; vertical-align: middle;">{v_ipi}</td>
                    <td style="text-align:center; padding: 4px; min-height: 25px; white-space: nowrap; vertical-align: middle;">{p_icms_formatted}</td>
                    <td style="text-align:center; padding: 4px; min-height: 25px; white-space: nowrap; vertical-align: middle;">{p_ipi_formatted}</td>
                </tr>
                """
                items_html_list.append(item_row)
        else:
            # Para muitos itens, implementar paginação
            return self._format_items_with_pagination(det_elements, max_items_per_page)
        return '\n'.join(items_html_list)

    def _format_items_with_pagination(self, det_elements, max_items_per_page: int = 15) -> str:
        """Formata itens com paginação - retorna apenas os itens para a primeira página."""
        items_html_list = []
        
        # Processa apenas os primeiros itens para a primeira página
        for det_element in det_elements[:max_items_per_page]:
            prod_element = det_element.find('ns:prod', NAMESPACES)
            imposto_element = det_element.find('ns:imposto', NAMESPACES)
            
            if prod_element is None or imposto_element is None:
                continue

            # Busca valores de ICMS - expandir para incluir todos os CSTs
            v_bc_icms_paths = [
                'ICMS/ICMS00/vBC', 'ICMS/ICMS10/vBC', 'ICMS/ICMS20/vBC', 'ICMS/ICMS30/vBC',
                'ICMS/ICMS40/vBC', 'ICMS/ICMS41/vBC', 'ICMS/ICMS50/vBC', 'ICMS/ICMS51/vBC',
                'ICMS/ICMS60/vBC', 'ICMS/ICMS70/vBC', 'ICMS/ICMS90/vBC'
            ]
            v_bc_icms = 'R$ 0,00'
            for path in v_bc_icms_paths:
                temp_v_bc_icms = self.get_xml_value(imposto_element, path, is_currency=True)
                if temp_v_bc_icms:
                    v_bc_icms = temp_v_bc_icms
                    break

            # Percentual ICMS - buscar valor bruto para formatação correta
            p_icms_paths = [
                'ICMS/ICMS00/pICMS', 'ICMS/ICMS10/pICMS', 'ICMS/ICMS20/pICMS', 'ICMS/ICMS30/pICMS',
                'ICMS/ICMS40/pICMS', 'ICMS/ICMS41/pICMS', 'ICMS/ICMS50/pICMS', 'ICMS/ICMS51/pICMS',
                'ICMS/ICMS60/pICMS', 'ICMS/ICMS70/pICMS', 'ICMS/ICMS90/pICMS'
            ]
            p_icms_raw = ''
            for path in p_icms_paths:
                p_icms_raw = self.get_xml_value(imposto_element, path)
                if p_icms_raw:
                    break
            
            # Formatar percentual ICMS
            p_icms_formatted = '0%'
            if p_icms_raw:
                try:
                    perc_value = float(p_icms_raw)
                    p_icms_formatted = f"{perc_value:.1f}%"
                except ValueError:
                    p_icms_formatted = "0%"

            # Valor ICMS
            v_icms_paths = [
                'ICMS/ICMS00/vICMS', 'ICMS/ICMS10/vICMS', 'ICMS/ICMS20/vICMS', 'ICMS/ICMS30/vICMS',
                'ICMS/ICMS40/vICMS', 'ICMS/ICMS41/vICMS', 'ICMS/ICMS50/vICMS', 'ICMS/ICMS51/vICMS',
                'ICMS/ICMS60/vICMS', 'ICMS/ICMS70/vICMS', 'ICMS/ICMS90/vICMS'
            ]
            v_icms = 'R$ 0,00'
            for path in v_icms_paths:
                temp_v_icms = self.get_xml_value(imposto_element, path, is_currency=True)
                if temp_v_icms:
                    v_icms = temp_v_icms
                    break

            v_ipi = self.get_xml_value(imposto_element, 'IPI/IPITrib/vIPI', default='', is_currency=True)
            if not v_ipi:
                v_ipi = "R$ 0,00"
            
            # Percentual IPI - buscar valor bruto para formatação correta
            p_ipi_raw = self.get_xml_value(imposto_element, 'IPI/IPITrib/pIPI', default='0.00')
            p_ipi_formatted = '0%'
            if p_ipi_raw:
                try:
                    perc_value = float(p_ipi_raw)
                    p_ipi_formatted = f"{perc_value:.1f}%"
                except ValueError:
                    p_ipi_formatted = "0%"

            item_row = f"""
            <tr>
                <td style="text-align:center; padding: 4px; min-height: 25px; white-space: nowrap;">{self.get_xml_value(prod_element, 'cProd')}</td>
                <td style="padding: 4px; min-height: 25px;">{self.get_xml_value(prod_element, 'xProd')}</td>
                <td style="text-align:center; padding: 4px; min-height: 25px; white-space: nowrap;">{self.get_xml_value(prod_element, 'NCM')}</td>
                <td style="text-align:center; padding: 4px; min-height: 25px; white-space: nowrap;">{self.get_xml_value(prod_element, 'CFOP')}</td>
                <td style="text-align:center; padding: 4px; min-height: 25px; white-space: nowrap;">{self.get_xml_value(prod_element, 'uCom')}</td>
                <td style="text-align:right; padding: 4px; min-height: 25px; white-space: nowrap;">{float(self.get_xml_value(prod_element, 'qCom', default='0')):,.4f}</td>
                <td style="text-align:right; padding: 4px; min-height: 25px; white-space: nowrap;">{self.get_xml_value(prod_element, 'vUnCom', is_currency=True)}</td>
                <td style="text-align:right; padding: 4px; min-height: 25px; white-space: nowrap;">{self.get_xml_value(prod_element, 'vProd', is_currency=True)}</td>
                <td style="text-align:right; padding: 4px; min-height: 25px; white-space: nowrap;">{v_bc_icms}</td>
                <td style="text-align:right; padding: 4px; min-height: 25px; white-space: nowrap;">{v_icms}</td>
                <td style="text-align:right; padding: 4px; min-height: 25px; white-space: nowrap;">{v_ipi}</td>
                <td style="text-align:right; padding: 4px; min-height: 25px; white-space: nowrap;">{p_icms_formatted}</td>
                <td style="text-align:right; padding: 4px; min-height: 25px; white-space: nowrap;">{p_ipi_formatted}</td>
            </tr>
            """
            items_html_list.append(item_row)
            
        return '\n'.join(items_html_list)

    def create_additional_pages(self, nfe_element: ET.Element, template_content: str, det_elements, max_items_per_page: int = 15) -> List[str]:
        """Cria páginas adicionais para NF-e com muitos itens."""
        additional_pages = []
        total_items = len(det_elements)
        
        # Calcula quantas páginas adicionais são necessárias
        remaining_items = total_items - max_items_per_page
        if remaining_items <= 0:
            return additional_pages
            
        # Processa os itens restantes em páginas
        start_index = max_items_per_page
        page_number = 2
        
        while start_index < total_items:
            end_index = min(start_index + max_items_per_page, total_items)
            page_items = det_elements[start_index:end_index]
            
            # Cria o HTML da página adicional
            page_html = self._create_additional_page_html(nfe_element, template_content, page_items, page_number, total_items)
            additional_pages.append(page_html)
            
            start_index = end_index
            page_number += 1
            
        return additional_pages

    def _create_additional_page_html(self, nfe_element: ET.Element, template_content: str, page_items, page_number: int, total_items: int) -> str:
        """Cria o HTML de uma página adicional."""
        # Extrai apenas o cabeçalho e a tabela de itens do template
        infNFe = nfe_element.find('ns:infNFe', NAMESPACES)
        ide = infNFe.find('ns:ide', NAMESPACES) if infNFe is not None else None
        emit = infNFe.find('ns:emit', NAMESPACES) if infNFe is not None else None
        dest = infNFe.find('ns:dest', NAMESPACES) if infNFe is not None else None
        enderEmit = emit.find('ns:enderEmit', NAMESPACES) if emit is not None else None
        enderDest = dest.find('ns:enderDest', NAMESPACES) if dest is not None else None
        
        # Formata os itens desta página
        items_html = self._format_page_items(page_items)
        
        # Template simplificado para páginas adicionais
        page_html = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <title>DANFE - Página {page_number}</title>
            <style>
                * {{ margin: 0; padding: 0; box-sizing: border-box; }}
                body {{ font-family: 'DejaVu Sans', Arial, sans-serif; font-size: 8pt; }}
                .container {{ width: 100%; max-width: 21cm; margin: 0 auto; }}
                .header {{ border: 1px solid black; margin-bottom: 2mm; padding: 3pt; }}
                .company-info {{ font-weight: bold; text-align: center; margin-bottom: 2pt; }}
                .page-info {{ text-align: center; font-size: 9pt; margin-bottom: 3pt; }}
                table {{ width: 100%; border-collapse: collapse; }}
                td, th {{ border: 1px solid black; padding: 2pt; text-align: left; font-size: 7pt; }}
                th {{ background-color: #f0f0f0; font-weight: bold; text-align: center; }}
                .items-table td {{ padding: 4px; min-height: 25px; }}
            </style>
        </head>
        <body>
            <div class="container">
                <!-- Cabeçalho simplificado -->
                <div class="header">
                    <div class="company-info">{self.get_xml_value(emit, 'xNome')}</div>
                    <div class="page-info">DANFE - Documento Auxiliar da Nota Fiscal Eletrônica - Página {page_number}</div>
                    <div style="text-align: center;">
                        <strong>NF-e Nº:</strong> {self.get_xml_value(ide, 'nNF')} | 
                        <strong>Série:</strong> {self.get_xml_value(ide, 'serie')} | 
                        <strong>Emissão:</strong> {self.format_datetime_field(self.get_xml_value(ide, 'dhEmi'), '%d.%m.%Y')}
                    </div>
                </div>
                
                <!-- Tabela de itens -->
                <p style="font-weight: bold; margin-bottom: 2pt;">Dados do produtos (continuação)</p>
                <table class="items-table">
                    <thead>
                        <tr>
                            <th style="width: 15.5mm">CÓDIGO</th>
                            <th style="width: 66.1mm">DESCRIÇÃO DO PRODUTO</th>
                            <th>NCMSH</th>
                            <th>CFOP</th>
                            <th>UN</th>
                            <th>QTD.</th>
                            <th>VLR.UNIT</th>
                            <th>VLR.TOTAL</th>
                            <th>BC ICMS</th>
                            <th>VLR.ICMS</th>
                            <th>VLR.IPI</th>
                            <th>ALIQ.ICMS</th>
                            <th>ALIQ.IPI</th>
                        </tr>
                    </thead>
                    <tbody>
                        {items_html}
                    </tbody>
                </table>
                
                <div style="text-align: center; margin-top: 5mm; font-size: 7pt;">
                    Página {page_number} - Continuação dos itens da NF-e (Total de itens: {total_items})
                </div>
            </div>
        </body>
        </html>
        """
        
        return page_html

    def _format_page_items(self, page_items) -> str:
        """Formata os itens de uma página específica."""
        items_html_list = []
        
        for det_element in page_items:
            prod_element = det_element.find('ns:prod', NAMESPACES)
            imposto_element = det_element.find('ns:imposto', NAMESPACES)
            
            if prod_element is None or imposto_element is None:
                continue

            # Busca valores de ICMS - expandir para incluir todos os CSTs
            v_bc_icms_paths = [
                'ICMS/ICMS00/vBC', 'ICMS/ICMS10/vBC', 'ICMS/ICMS20/vBC', 'ICMS/ICMS30/vBC',
                'ICMS/ICMS40/vBC', 'ICMS/ICMS41/vBC', 'ICMS/ICMS50/vBC', 'ICMS/ICMS51/vBC',
                'ICMS/ICMS60/vBC', 'ICMS/ICMS70/vBC', 'ICMS/ICMS90/vBC'
            ]
            v_bc_icms = 'R$ 0,00'
            for path in v_bc_icms_paths:
                temp_v_bc_icms = self.get_xml_value(imposto_element, path, is_currency=True)
                if temp_v_bc_icms:
                    v_bc_icms = temp_v_bc_icms
                    break

            # Percentual ICMS
            p_icms_paths = [
                'ICMS/ICMS00/pICMS', 'ICMS/ICMS10/pICMS', 'ICMS/ICMS20/pICMS', 'ICMS/ICMS30/pICMS',
                'ICMS/ICMS40/pICMS', 'ICMS/ICMS41/pICMS', 'ICMS/ICMS50/pICMS', 'ICMS/ICMS51/pICMS',
                'ICMS/ICMS60/pICMS', 'ICMS/ICMS70/pICMS', 'ICMS/ICMS90/pICMS'
            ]
            p_icms_raw = ''
            for path in p_icms_paths:
                p_icms_raw = self.get_xml_value(imposto_element, path)
                if p_icms_raw:
                    break
            
            p_icms_formatted = '0%'
            if p_icms_raw:
                try:
                    perc_value = float(p_icms_raw)
                    p_icms_formatted = f"{perc_value:.1f}%"
                except ValueError:
                    p_icms_formatted = "0%"

            # Valor ICMS
            v_icms_paths = [
                'ICMS/ICMS00/vICMS', 'ICMS/ICMS10/vICMS', 'ICMS/ICMS20/vICMS', 'ICMS/ICMS30/vICMS',
                'ICMS/ICMS40/vICMS', 'ICMS/ICMS41/vICMS', 'ICMS/ICMS50/vICMS', 'ICMS/ICMS51/vICMS',
                'ICMS/ICMS60/vICMS', 'ICMS/ICMS70/vICMS', 'ICMS/ICMS90/vICMS'
            ]
            v_icms = 'R$ 0,00'
            for path in v_icms_paths:
                temp_v_icms = self.get_xml_value(imposto_element, path, is_currency=True)
                if temp_v_icms:
                    v_icms = temp_v_icms
                    break

            v_ipi = self.get_xml_value(imposto_element, 'IPI/IPITrib/vIPI', default='', is_currency=True)
            if not v_ipi:
                v_ipi = "R$ 0,00"
            
            # Percentual IPI
            p_ipi_raw = self.get_xml_value(imposto_element, 'IPI/IPITrib/pIPI', default='0.00')
            p_ipi_formatted = '0%'
            if p_ipi_raw:
                try:
                    perc_value = float(p_ipi_raw)
                    p_ipi_formatted = f"{perc_value:.1f}%"
                except ValueError:
                    p_ipi_formatted = "0%"

            item_row = f"""
            <tr>
                <td style="text-align:center; padding: 4px; min-height: 25px; white-space: nowrap;">{self.get_xml_value(prod_element, 'cProd')}</td>
                <td style="padding: 4px; min-height: 25px;">{self.get_xml_value(prod_element, 'xProd')}</td>
                <td style="text-align:center; padding: 4px; min-height: 25px; white-space: nowrap;">{self.get_xml_value(prod_element, 'NCM')}</td>
                <td style="text-align:center; padding: 4px; min-height: 25px; white-space: nowrap;">{self.get_xml_value(prod_element, 'CFOP')}</td>
                <td style="text-align:center; padding: 4px; min-height: 25px; white-space: nowrap;">{self.get_xml_value(prod_element, 'uCom')}</td>
                <td style="text-align:right; padding: 4px; min-height: 25px; white-space: nowrap;">{float(self.get_xml_value(prod_element, 'qCom', default='0')):,.4f}</td>
                <td style="text-align:right; padding: 4px; min-height: 25px; white-space: nowrap;">{self.get_xml_value(prod_element, 'vUnCom', is_currency=True)}</td>
                <td style="text-align:right; padding: 4px; min-height: 25px; white-space: nowrap;">{self.get_xml_value(prod_element, 'vProd', is_currency=True)}</td>
                <td style="text-align:right; padding: 4px; min-height: 25px; white-space: nowrap;">{v_bc_icms}</td>
                <td style="text-align:right; padding: 4px; min-height: 25px; white-space: nowrap;">{v_icms}</td>
                <td style="text-align:right; padding: 4px; min-height: 25px; white-space: nowrap;">{v_ipi}</td>
                <td style="text-align:right; padding: 4px; min-height: 25px; white-space: nowrap;">{p_icms_formatted}</td>
                <td style="text-align:right; padding: 4px; min-height: 25px; white-space: nowrap;">{p_ipi_formatted}</td>
            </tr>
            """
            items_html_list.append(item_row)
            
        return '\n'.join(items_html_list)

    def format_duplicates_html(self, root_nfe: ET.Element) -> str:
        """Formata as duplicatas (faturas) como HTML."""
        duplicates_html_list = []
        infNFe = root_nfe.find('ns:infNFe', NAMESPACES)
        if infNFe is None:
            return ''
        
        cobr_element = infNFe.find('ns:cobr', NAMESPACES)
        if cobr_element is not None:
            fat_element = cobr_element.find('ns:fat', NAMESPACES)
            n_fat = self.get_xml_value(fat_element, 'nFat') if fat_element is not None else ''
            
            dup_table_start = '<table cellpadding="0" cellspacing="0" border="1" style="width:100%;"><tbody>'
            dup_table_end = '</tbody></table>'

            duplicates_html_list.append(dup_table_start)
            
            # Adicionar cabeçalho "Vencimento NF"
            header_row = f"""
                <tr style="background-color: #f0f0f0;">
                    <td style="text-align:center; width: 33%; font-weight: bold; padding: 4px;">Número</td>
                    <td style="text-align:center; width: 33%; font-weight: bold; padding: 4px;">Vencimento NF</td>
                    <td style="text-align:center; width: 34%; font-weight: bold; padding: 4px;">Valor</td>
                </tr>
                """
            duplicates_html_list.append(header_row)

            for dup_element in cobr_element.findall('ns:dup', NAMESPACES):
                n_dup = self.get_xml_value(dup_element, 'nDup')
                d_venc = self.format_datetime_field(self.get_xml_value(dup_element, 'dVenc'))
                v_dup = self.get_xml_value(dup_element, 'vDup', is_currency=True)
                
                display_dup_num = n_dup if n_dup else n_fat
                if n_fat and n_dup and not n_dup.startswith(n_fat):
                    display_dup_num = f'{n_fat}/{n_dup}'

                dup_row = f"""
                <tr>
                    <td style="text-align:center; width: 33%;">{display_dup_num}</td>
                    <td style="text-align:center; width: 33%;">{d_venc}</td>
                    <td style="text-align:right; width: 34%;">{v_dup}</td>
                </tr>
                """
                duplicates_html_list.append(dup_row)
            
            if not cobr_element.findall('ns:dup', NAMESPACES) and fat_element is not None:
                # Se não há duplicatas mas há fatura, ainda adicionar cabeçalho se não foi adicionado
                if len(duplicates_html_list) == 1:  # Apenas a tag de abertura da tabela
                    header_row = f"""
                    <tr style="background-color: #f0f0f0;">
                        <td style="text-align:center; width: 33%; font-weight: bold; padding: 4px;">Número</td>
                        <td style="text-align:center; width: 33%; font-weight: bold; padding: 4px;">Vencimento NF</td>
                        <td style="text-align:center; width: 34%; font-weight: bold; padding: 4px;">Valor</td>
                    </tr>
                    """
                    duplicates_html_list.append(header_row)
                
                v_liq = self.get_xml_value(fat_element, 'vLiq', is_currency=True)
                dup_row = f"""
                <tr>
                    <td style="text-align:center; width: 33%;">{n_fat}</td>
                    <td style="text-align:center; width: 33%;">-</td>
                    <td style="text-align:right; width: 34%;">{v_liq}</td>
                </tr>
                """
                duplicates_html_list.append(dup_row)

            duplicates_html_list.append(dup_table_end)
        
        if not duplicates_html_list or len(duplicates_html_list) <= 2:
            return '''<table cellpadding="0" cellspacing="0" border="1" style="width:100%;"><tbody>
                <tr style="background-color: #f0f0f0;">
                    <td style="text-align:center; width: 33%; font-weight: bold; padding: 4px;">Número</td>
                    <td style="text-align:center; width: 33%; font-weight: bold; padding: 4px;">Vencimento NF</td>
                    <td style="text-align:center; width: 34%; font-weight: bold; padding: 4px;">Valor</td>
                </tr>
                <tr>
                    <td style="text-align:center;">&nbsp;</td>
                    <td style="text-align:center;">&nbsp;</td>
                    <td style="text-align:right;">&nbsp;</td>
                </tr>
            </tbody></table>'''

        return '\n'.join(duplicates_html_list)

    def prepare_html_for_pdf(self, html_content: str) -> str:
        """
        Prepara o HTML especificamente para conversão em PDF pelo WeasyPrint.
        Remove elementos problemáticos e otimiza a estrutura.
        """
        # Substituir fontes problemáticas por fontes que o WeasyPrint reconhece
        html_content = html_content.replace('"Times New Roman"', '"DejaVu Serif"')
        
        # Remover propriedades CSS que causam problemas no WeasyPrint
        problematic_css = [
            'position: relative;',
            'overflow: hidden;',
            'border-radius: 5px;',
            'margin: -1pt;',
            'width: 100.5%;'
        ]
        
        for css_rule in problematic_css:
            html_content = html_content.replace(css_rule, '')
        
        # Corrigir margens e espaçamentos problemáticos
        html_content = html_content.replace('margin-top: -1px;', 'margin-top: 0;')
        html_content = html_content.replace('height: 5mm;', 'min-height: 5mm;')
        
        # FORÇAR aparição da chave de acesso com estilo inline
        html_content = html_content.replace(
            '<span class="chave-acesso">',
            '<span class="chave-acesso" style="display: block !important; font-size: 8pt !important; font-weight: bold !important; text-align: center !important; letter-spacing: 1pt !important; padding: 3pt 6pt !important; visibility: visible !important; color: black !important; background: white !important; font-family: monospace !important;">'
        )
        
        # Adicionar classes específicas para elementos que precisam de tratamento especial
        html_content = html_content.replace('<table', '<table class="pdf-table"')
        
        return html_content

    def _get_combined_additional_info(self, infAdic_element):
        """Combina informações adicionais do fisco e complementares"""
        if infAdic_element is None:
            return ''
        
        # Capturar ambas as informações
        inf_ad_fisco = self.get_xml_value(infAdic_element, 'infAdFisco')
        inf_cpl = self.get_xml_value(infAdic_element, 'infCpl')
        
        # Combinar as informações
        combined_info = []
        if inf_ad_fisco:
            combined_info.append(inf_ad_fisco)
        if inf_cpl:
            combined_info.append(inf_cpl)
        
        # Retornar combinado ou vazio
        return ' | '.join(combined_info) if combined_info else ''

    def _generate_barcode_base64(self, chave_nfe: str) -> str:
        """
        Gera código de barras da chave de acesso em base64 para incorporar no HTML
        
        Args:
            chave_nfe: Chave de 44 dígitos da NF-e
            
        Returns:
            String base64 da imagem do código de barras
        """
        if not BARCODE_AVAILABLE or not chave_nfe or len(chave_nfe) != 44:
            return ''
        
        try:
            # Usar Code128 que é adequado para a chave da NF-e
            from barcode import Code128
            
            # Configurações para o código de barras
            options = {
                'module_width': 0.2,    # Largura das barras
                'module_height': 8.0,   # Altura das barras
                'quiet_zone': 2.0,      # Zona quieta
                'font_size': 8,         # Tamanho da fonte
                'text_distance': 1.0,   # Distância do texto
                'background': 'white',  # Fundo branco
                'foreground': 'black',  # Barras pretas
            }
            
            # Gerar código de barras
            code = Code128(chave_nfe, writer=ImageWriter())
            
            # Criar buffer de memória
            buffer = BytesIO()
            
            # Salvar imagem no buffer
            code.write(buffer, options=options)
            
            # Converter para base64
            buffer.seek(0)
            img_base64 = base64.b64encode(buffer.read()).decode('utf-8')
            
            return f"data:image/png;base64,{img_base64}"
            
        except Exception as e:
            print(f"Erro ao gerar código de barras: {e}")
            return ''

    def remove_unnecessary_sections(self, html_content: str) -> str:
        """
        Remove seções desnecessárias da DANFE conforme solicitado.
        """
        import re
        
        # Padrões para remover a seção de recebimento
        receipt_patterns = [
            # Padrão específico fornecido pelo usuário
            r'<tbody>\s*<tr>\s*<td colspan="2" class="txt-upper">\s*Recebemos de.*?</tbody>',
            
            # Padrões alternativos caso a estrutura seja diferente
            r'<tr[^>]*>\s*<td[^>]*colspan="2"[^>]*class="txt-upper"[^>]*>\s*Recebemos de.*?</tr>\s*<tr[^>]*>\s*<td[^>]*>\s*<span[^>]*class="nf-label"[^>]*>Data de recebimento.*?</tr>',
            
            # Padrão mais amplo para capturar variações
            r'<tr[^>]*>.*?Recebemos de.*?</tr>\s*<tr[^>]*>.*?Data de recebimento.*?</tr>',
            
            # Padrão pela estrutura da tabela de recebimento
            r'<tbody[^>]*>.*?Recebemos de.*?Identificação de assinatura do Recebedor.*?</tbody>',
        ]
        
        # Padrões para remover apenas o título da seção de transportador
        transport_patterns = [
            # Apenas o título específico mencionado
            r'<p class="area-name">Transportador/volumes transportados</p>',
            
            # Variações do título
            r'<p[^>]*class="area-name"[^>]*>Transportador/volumes transportados</p>',
            
            # Seção 1: QUANTIDADE, ESPÉCIE, MARCA, NUMERAÇÃO, PESO BRUTO, PESO LÍQUIDO (apenas tbody específico)
            r'<tbody>\s*<tr>\s*<td[^>]*class="field quantidade"[^>]*>\s*<span[^>]*class="nf-label"[^>]*>QUANTIDADE</span>.*?<span[^>]*class="nf-label"[^>]*>ESPÉCIE</span>.*?<span[^>]*class="nf-label"[^>]*>MARCA</span>.*?<span[^>]*class="nf-label"[^>]*>NUMERAÇÃO</span>.*?<span[^>]*class="nf-label"[^>]*>PESO BRUTO</span>.*?<span[^>]*class="nf-label"[^>]*>PESO LÍQUIDO</span>.*?</tr>\s*</tbody>',
            
            # Seção 2: Tabela específica com class="no-top" e ENDEREÇO
            r'<table[^>]*cellpadding="0"[^>]*cellspacing="0"[^>]*border="1"[^>]*class="no-top"[^>]*>\s*<tbody>\s*<tr>\s*<td[^>]*class="field endereco"[^>]*>\s*<span[^>]*class="nf-label"[^>]*>ENDEREÇO</span>.*?<span[^>]*class="nf-label"[^>]*>MUNICÍPIO</span>.*?<span[^>]*class="nf-label"[^>]*>UF</span>.*?<span[^>]*class="nf-label"[^>]*>INSC\.\s*ESTADUAL</span>.*?</tr>\s*</tbody>\s*</table>',
            
            # Seção 3: Tabela específica com RAZÃO SOCIAL e FRETE POR CONTA
            r'<table[^>]*cellpadding="0"[^>]*cellspacing="0"[^>]*border="1"[^>]*>\s*<tbody>\s*<tr>\s*<td[^>]*>\s*<span[^>]*class="nf-label"[^>]*>RAZÃO SOCIAL</span>.*?<td[^>]*class="freteConta"[^>]*>\s*<span[^>]*class="nf-label"[^>]*>FRETE POR CONTA</span>.*?<span[^>]*class="nf-label"[^>]*>CÓDIGO ANTT</span>.*?<span[^>]*class="nf-label"[^>]*>PLACA</span>.*?<span[^>]*class="nf-label"[^>]*>CNPJ/CPF</span>.*?</tr>\s*</tbody>\s*</table>',
        ]
        
        # Remover seções de recebimento
        for pattern in receipt_patterns:
            html_content = re.sub(pattern, '', html_content, flags=re.DOTALL | re.IGNORECASE)
        
        # Remover seções de transportador
        for pattern in transport_patterns:
            html_content = re.sub(pattern, '', html_content, flags=re.DOTALL | re.IGNORECASE)
        
        # Limpar possíveis tags vazias que sobraram
        html_content = re.sub(r'<tbody>\s*</tbody>', '', html_content)
        html_content = re.sub(r'<table[^>]*>\s*</table>', '', html_content)
        html_content = re.sub(r'<div[^>]*class="wrapper-table"[^>]*>\s*</div>', '', html_content)
        
        return html_content

    def processar_nfe(self, xml_path: str, template_path: str, output_dir: str, 
                     pdf_filename: str = None, gerar_apenas_pdf: bool = False) -> Dict[str, Any]:
        """
        Processa o arquivo XML da NF-e e gera o DANFE
        
        Args:
            xml_path: Caminho do arquivo XML
            template_path: Caminho do template HTML
            output_dir: Diretório de saída
            pdf_filename: Nome específico para o arquivo PDF
            gerar_apenas_pdf: Se True, gera apenas PDF (não HTML)
            
        Returns:
            Dicionário com status e mensagens
        """
        try:
            tree = ET.parse(xml_path)
            root = tree.getroot()
        except FileNotFoundError:
            return {"success": False, "message": f"Erro: Arquivo XML não encontrado em {xml_path}"}
        except ET.ParseError:
            return {"success": False, "message": f"Erro: Falha ao parsear o arquivo XML {xml_path}"}

        try:
            with open(template_path, 'r', encoding='utf-8') as f:
                html_content = f.read()
        except FileNotFoundError:
            return {"success": False, "message": f"Erro: Arquivo de template HTML não encontrado em {template_path}"}

        # Encontrar elemento NFe
        nfe_element = root.find('ns:NFe', NAMESPACES)
        if nfe_element is None:
            nfe_element = root
            if nfe_element.find('ns:infNFe', NAMESPACES) is None:
                return {"success": False, "message": "Erro: Tag <NFe> ou <infNFe> não encontrada no XML."}

        infNFe = nfe_element.find('ns:infNFe', NAMESPACES)
        if infNFe is None:
            return {"success": False, "message": "Erro: Tag <infNFe> não encontrada no XML."}
            
        # Extrair elementos principais
        ide = infNFe.find('ns:ide', NAMESPACES)
        emit = infNFe.find('ns:emit', NAMESPACES)
        dest = infNFe.find('ns:dest', NAMESPACES)
        total = infNFe.find('ns:total', NAMESPACES)
        ICMSTot = total.find('ns:ICMSTot', NAMESPACES) if total is not None else None
        infAdic = infNFe.find('ns:infAdic', NAMESPACES)
        
        enderEmit = emit.find('ns:enderEmit', NAMESPACES) if emit is not None else None
        enderDest = dest.find('ns:enderDest', NAMESPACES) if dest is not None else None
        # Elementos de transporte e serviços removidos conforme solicitação
        
        # Chave da NF-e
        chNFe_val = ''
        protNFe = root.find('ns:protNFe', NAMESPACES)
        if protNFe is not None:
            infProt = protNFe.find('ns:infProt', NAMESPACES)
            if infProt is not None:
                chNFe_val = self.get_xml_value(infProt, 'chNFe')
        if not chNFe_val and infNFe.get('Id'):
            chNFe_val = infNFe.get('Id')[3:]
        
        # Protocolo
        protocol_label_val = "PROTOCOLO DE AUTORIZAÇÃO DE USO" if chNFe_val else ""
        nProt_val = self.get_xml_value(infProt if protNFe is not None else None, 'nProt')
        dhRecbto_val = self.format_datetime_field(self.get_xml_value(infProt if protNFe is not None else None, 'dhRecbto'), '%d/%m/%Y %H:%M:%S')
        protocol_val = f'{nProt_val} - {dhRecbto_val}' if nProt_val and dhRecbto_val else (nProt_val or dhRecbto_val or '')

        # Mapeamento de placeholders
        replacements = {
            '[ds_company_issuer_name]': self.get_xml_value(emit, 'xNome'),
            '[nl_invoice]': self.get_xml_value(ide, 'nNF'),
            '[ds_invoice_serie]': self.get_xml_value(ide, 'serie'),
            '[url_logo]': '',
            '[ds_company_address]': f"{self.get_xml_value(enderEmit, 'xLgr')} {self.get_xml_value(enderEmit, 'nro')} {self.get_xml_value(enderEmit, 'xCpl')}".strip(),
            '[ds_company_neighborhood]': self.get_xml_value(enderEmit, 'xBairro'),
            '[nu_company_cep]': self.get_xml_value(enderEmit, 'CEP'),
            '[ds_company_city_name]': self.get_xml_value(enderEmit, 'xMun'),
            '[ds_company_uf]': self.get_xml_value(enderEmit, 'UF'),
            '[nl_company_phone_number]': self.get_xml_value(enderEmit, 'fone'),
            '[ds_code_operation_type]': self.get_xml_value(ide, 'tpNF'),
            '[actual_page]': '1',
            '[total_pages]': '1',
            '{BarCode}': chNFe_val,
            '[ds_danfe]': chNFe_val,
            '[_ds_transaction_nature]': self.get_xml_value(ide, 'natOp'),
            '[protocol_label]': protocol_label_val,
            '[ds_protocol]': protocol_val,
            '[nl_company_ie]': self.get_xml_value(emit, 'IE'),
            '[nl_company_ie_st]': self.get_xml_value(emit, 'IEST'),
            '[nl_company_cnpj_cpf]': self.get_xml_value(emit, 'CNPJ'),
            '[ds_client_receiver_name]': self.get_xml_value(dest, 'xNome'),
            '[nl_client_cnpj_cpf]': self.get_xml_value(dest, 'CNPJ', default=self.get_xml_value(dest, 'CPF')),
            '[dt_invoice_issue]': self.format_datetime_field(self.get_xml_value(ide, 'dhEmi'), '%d.%m.%Y'),
            '[ds_client_address]': f"{self.get_xml_value(enderDest, 'xLgr')} {self.get_xml_value(enderDest, 'nro')} {self.get_xml_value(enderDest, 'xCpl')}".strip(),
            '[ds_client_neighborhood]': self.get_xml_value(enderDest, 'xBairro'),
            '[nu_client_cep]': self.get_xml_value(enderDest, 'CEP'),
            '[dt_input_output]': self.format_datetime_field(self.get_xml_value(ide, 'dhSaiEnt', default=self.get_xml_value(ide, 'dhEmi'))),
            '[hr_input_output]': self.format_datetime_field(self.get_xml_value(ide, 'dhSaiEnt', default=self.get_xml_value(ide, 'dhEmi')), '%H:%M:%S'),
            '[ds_client_city_name]': self.get_xml_value(enderDest, 'xMun'),
            '[nl_client_phone_number]': self.get_xml_value(enderDest, 'fone'),
            '[ds_client_uf]': self.get_xml_value(enderDest, 'UF'),
            '[ds_client_ie]': self.get_xml_value(dest, 'IE'),
            '[tot_bc_icms]': self.get_xml_value(ICMSTot, 'vBC', is_currency=True),
            '[tot_icms]': self.get_xml_value(ICMSTot, 'vICMS', is_currency=True),
            '[tot_bc_icms_st]': self.get_xml_value(ICMSTot, 'vBCST', is_currency=True),
            '[tot_icms_st]': self.get_xml_value(ICMSTot, 'vST', is_currency=True),
            '[tot_icms_fcp]': self.get_xml_value(ICMSTot, 'vFCP', is_currency=True),
            '[vl_total_prod]': self.get_xml_value(ICMSTot, 'vProd', is_currency=True),
            '{ApproximateTax}': self.get_xml_value(ICMSTot, 'vTotTrib', is_currency=True),
            '[vl_shipping]': self.get_xml_value(ICMSTot, 'vFrete', is_currency=True),
            '[vl_insurance]': self.get_xml_value(ICMSTot, 'vSeg', is_currency=True),
            '[vl_discount]': self.get_xml_value(ICMSTot, 'vDesc', is_currency=True),
            '[vl_other_expense]': self.get_xml_value(ICMSTot, 'vOutro', is_currency=True),
            '[tot_total_ipi_tax]': self.get_xml_value(ICMSTot, 'vIPI', is_currency=True),
            '[vl_total]': self.get_xml_value(ICMSTot, 'vNF', is_currency=True),
            # Dados de transporte removidos conforme solicitação
            '[ds_transport_carrier_name]': '',
            '[ds_transport_code_shipping_type]': '',
            '[ds_transport_rntc]': '',
            '[ds_transport_vehicle_plate]': '',
            '[ds_transport_vehicle_uf]': '',
            '[nl_transport_cnpj_cpf]': '',
            '[ds_transport_address]': '',
            '[ds_transport_city]': '',
            '[ds_transport_uf]': '',
            '[ds_transport_ie]': '',
            '[nu_transport_amount_transported_volumes]': '',
            '[ds_transport_type_volumes_transported]': '',
            '[ds_transport_mark_volumes_transported]': '',
            '[ds_transport_number_volumes_transported]': '',
            '[vl_transport_gross_weight]': '',
            '[vl_transport_net_weight]': '',
            '[ds_company_im]': self.get_xml_value(emit, 'IM'),
            # Impostos de serviços removidos conforme solicitação
            '[vl_total_serv]': '',
            '[tot_bc_issqn]': '',
            '[tot_issqn]': '',
            '[ds_additional_information]': self._get_combined_additional_info(infAdic),
            '[text_consult_nfe]': 'Consulta de autenticidade no portal nacional da NF-e www.nfe.fazenda.gov.br/portal ou no site da Sefaz Autorizada.',
            '[page-break]': ''
        }

        # Substituir placeholders
        for placeholder, value in replacements.items():
            html_content = html_content.replace(placeholder, str(value if value is not None else ''))

        # Verificar se precisa de paginação
        det_elements = infNFe.findall('ns:det', NAMESPACES)
        max_items_per_page = 15
        needs_pagination = len(det_elements) > max_items_per_page
        
        # Atualizar total de páginas se houver paginação
        if needs_pagination:
            total_pages = (len(det_elements) + max_items_per_page - 1) // max_items_per_page
            html_content = html_content.replace('[total_pages]', str(total_pages))

        # Substituir itens e duplicatas
        items_html = self.format_items_html(nfe_element, max_items_per_page)
        html_content = html_content.replace('[items]', items_html)
        
        duplicates_html_content = self.format_duplicates_html(nfe_element)
        html_content = html_content.replace('[duplicates]', duplicates_html_content)

        # Remover seções desnecessárias conforme solicitado
        html_content = self.remove_unnecessary_sections(html_content)

        # Gerar arquivos de saída
        if pdf_filename:
            output_pdf_path = os.path.join(output_dir, pdf_filename)
            # Nome HTML baseado no PDF mas só será usado se não for apenas PDF
            html_filename = pdf_filename.replace('.pdf', '.html')
            output_html_path = os.path.join(output_dir, html_filename)
        else:
            output_html_path = os.path.join(output_dir, 'danfe_gerada.html')
            output_pdf_path = os.path.join(output_dir, 'danfe_gerada.pdf')

        messages = []
        
        # Gerar HTML apenas se não for modo "apenas PDF"
        if not gerar_apenas_pdf:
            try:
                with open(output_html_path, 'w', encoding='utf-8') as f:
                    f.write(html_content)
                messages.append(f"DANFE HTML gerada com sucesso em: {output_html_path}")
            except IOError as e:
                return {"success": False, "message": f"Erro ao gerar HTML: {e}"}
        
        try:

            # Converter para PDF se WeasyPrint estiver disponível
            if WEASYPRINT_AVAILABLE:
                try:
                    if needs_pagination:
                        # Gerar páginas adicionais
                        additional_pages = self.create_additional_pages(nfe_element, html_content, det_elements, max_items_per_page)
                        
                        # Combinar todas as páginas em um único HTML
                        all_pages_html = html_content
                        
                        for page_html in additional_pages:
                            # Adicionar quebra de página antes de cada página adicional
                            all_pages_html += '<div style="page-break-before: always;"></div>' + page_html
                        
                        html_doc = WeasyHTML(string=all_pages_html, base_url=output_html_path)
                        html_doc.write_pdf(output_pdf_path)
                        messages.append(f"DANFE PDF com {len(additional_pages) + 1} páginas gerada em: {output_pdf_path}")
                    else:
                        # Método simples para NF-e com poucos itens
                        html_doc = WeasyHTML(string=html_content, base_url=output_html_path)
                        html_doc.write_pdf(output_pdf_path)
                        messages.append(f"DANFE PDF gerada com método simples em: {output_pdf_path}")
                    
                except Exception as e:
                    messages.append(f"Erro na conversão simples: {e}")
                    # Tentar com otimizações
                    try:
                        # Abordagem mais agressiva: modificar o HTML especificamente para PDF
                        html_content_pdf = self.prepare_html_for_pdf(html_content)
                        
                        # FORÇAR aparição da chave de acesso com estilo inline após prepare_html_for_pdf
                        html_content_pdf = html_content_pdf.replace(
                            '<span class="chave-acesso">',
                            '<span class="chave-acesso" style="display: block !important; font-size: 8pt !important; font-weight: bold !important; text-align: center !important; letter-spacing: 1pt !important; padding: 3pt 6pt !important; visibility: visible !important; color: black !important; background: white !important; font-family: monospace !important;">'
                        )
                        
                        # ADICIONAR uma segunda chave de acesso em estrutura mais simples
                        # Localizar o final do cabeçalho e adicionar uma linha extra para a chave
                        if 'CHAVE DE ACESSO' in html_content_pdf:
                            # Encontrar a chave no HTML
                            import re
                            chave_match = re.search(r'<span class="chave-acesso"[^>]*>([^<]+)</span>', html_content_pdf)
                            if chave_match:
                                chave_valor = chave_match.group(1)
                                
                                # Adicionar uma versão adicional da chave em local mais visível
                                extra_chave_html = f'''
                                <!-- CHAVE DE ACESSO ADICIONAL PARA PDF -->
                                <div style="width: 100%; background: #f0f0f0; border: 2px solid black; padding: 10pt; margin: 5pt 0; text-align: center;">
                                    <p style="margin: 0; font-size: 6pt; font-weight: bold; color: black;">CHAVE DE ACESSO DA NF-E</p>
                                    <p style="margin: 2pt 0 0 0; font-size: 9pt; font-weight: bold; color: black; font-family: monospace; letter-spacing: 1pt;">{chave_valor}</p>
                                </div>
                                '''
                                
                                # Inserir após o primeiro cabeçalho
                                html_content_pdf = html_content_pdf.replace(
                                    '<!-- Natureza da Operação -->',
                                    extra_chave_html + '\n        <!-- Natureza da Operação -->'
                                )
                        
                        from weasyprint import CSS
                        
                        # CSS otimizado especificamente para WeasyPrint
                        css_weasyprint = CSS(string="""
                            @page {
                                size: A4;
                                margin: 8mm 10mm 8mm 10mm;
                            }
                            
                            /* Reset completo para PDF */
                            * {
                                box-sizing: border-box;
                            }
                            
                            body {
                                font-family: "DejaVu Sans", "Liberation Sans", Arial, sans-serif;
                                margin: 0;
                                padding: 0;
                            }
                            
                            .nfeArea.page {
                                width: 19cm;
                                margin: 0 auto;
                                font-size: 8pt;
                            }
                            
                            .nfeArea table {
                                width: 100%;
                                border-collapse: collapse;
                                table-layout: fixed;
                                margin-bottom: 2pt;
                            }
                            
                            .nfeArea td, .nfeArea th {
                                border: 1px solid #000;
                                padding: 1pt 2pt;
                                vertical-align: top;
                                word-wrap: break-word;
                                font-size: 6pt;
                                line-height: 1.2;
                            }
                            
                            .nfeArea .area-name {
                                font-size: 5pt;
                                font-weight: bold;
                                text-transform: uppercase;
                                margin: 2pt 0 1pt 0;
                            }
                            
                            .nfeArea .info {
                                font-weight: bold;
                                font-size: 7pt;
                            }
                            
                            .nfeArea .font-12 {
                                font-size: 10pt;
                            }
                            
                            .nfeArea .font-8 {
                                font-size: 7pt;
                            }
                            
                            .nfeArea .nf-label {
                                font-size: 5pt;
                                text-transform: uppercase;
                                margin-bottom: 1pt;
                                display: block !important;
                            }
                            
                            .nfeArea .nf-label.label-small {
                                font-size: 4pt;
                            }
                            
                            .nfeArea .tserie {
                                text-align: center;
                                font-weight: bold;
                            }
                            
                            .nfeArea .txt-center {
                                text-align: center;
                            }
                            
                            .nfeArea .txt-right {
                                text-align: right;
                            }
                            
                            .nfeArea .entradaSaida .identificacao {
                                border: 1px solid black;
                                width: 8mm;
                                height: 8mm;
                                text-align: center;
                                line-height: 8mm;
                                float: right;
                                margin: 2pt;
                            }
                            
                            /* CSS MUITO específico para chave de acesso - FORÇAR APARIÇÃO */
                            .chave-acesso,
                            span.chave-acesso,
                            .nfeArea .chave-acesso,
                            .nfeArea span.chave-acesso,
                            table td span.chave-acesso,
                            tbody tr td span.chave-acesso {
                                font-family: "DejaVu Sans Mono", "Liberation Mono", "Courier New", monospace !important;
                                font-size: 8pt !important;
                                font-weight: bold !important;
                                text-align: center !important;
                                letter-spacing: 1pt !important;
                                padding: 3pt 6pt !important;
                                display: block !important;
                                visibility: visible !important;
                                opacity: 1 !important;
                                height: auto !important;
                                width: auto !important;
                                color: #000000 !important;
                                background: #ffffff !important;
                                border: none !important;
                                margin: 2pt 0 !important;
                                line-height: 1.3 !important;
                                word-break: break-all !important;
                                overflow: visible !important;
                            }
                            
                            /* Ajustes específicos para elementos que estavam quebrando */
                            .wrapper-table {
                                margin-bottom: 2pt;
                            }
                            
                            .wrapper-table table {
                                margin: 0;
                            }
                            
                            /* Estilos específicos para tabelas de itens - aumentar tamanho das células */
                            .nfeArea table.items-table td {
                                padding: 6pt 4pt !important;
                                min-height: 30px !important;
                                line-height: 1.3 !important;
                                font-size: 7pt !important;
                                white-space: nowrap !important;
                                vertical-align: middle !important;
                            }
                            
                            /* Células de valores monetários - não quebrar linha */
                            .nfeArea table td[style*="text-align:right"] {
                                white-space: nowrap !important;
                                font-size: 7pt !important;
                                padding: 6pt 4pt !important;
                            }
                            
                            /* Células centralizadas */
                            .nfeArea table td[style*="text-align:center"] {
                                white-space: nowrap !important;
                                font-size: 7pt !important;
                                padding: 6pt 4pt !important;
                            }
                            
                            /* Melhorar legibilidade geral das tabelas */
                            .nfeArea table tr {
                                min-height: 30px !important;
                            }
                        """)
                        
                        html_doc = WeasyHTML(string=html_content_pdf, base_url=output_html_path)
                        html_doc.write_pdf(output_pdf_path, stylesheets=[css_weasyprint])
                        messages.append(f"DANFE PDF gerada com HTML otimizado em: {output_pdf_path}")
                        
                    except Exception as e2:
                        messages.append(f"Erro em todos os métodos: {e2}")
            else:
                return {"success": False, "message": "Biblioteca WeasyPrint não encontrada. Para gerar PDF, instale-a com: pip install WeasyPrint"}

            return {"success": True, "message": "\n".join(messages)}

        except IOError as e:
            return {"success": False, "message": f"Erro: Não foi possível escrever os arquivos de saída: {e}"}

    def processar_nfe_optimized(self, xml_path: str, template_content: str, output_dir: str, pdf_filename: str) -> Dict[str, Any]:
        """
        Versão otimizada para processamento em massa
        - Template já carregado em memória
        - Gera apenas PDF
        - Menos verificações
        - Reutiliza objetos
        """
        try:
            # Parse XML (otimizado - menos verificações)
            tree = ET.parse(xml_path)
            root = tree.getroot()
            
            # Encontrar elementos principais (caminho direto)
            nfe_element = root.find('ns:NFe', NAMESPACES) or root
            infNFe = nfe_element.find('ns:infNFe', NAMESPACES)
            
            if infNFe is None:
                return {"success": False, "message": "infNFe não encontrada"}
            
            # Extrair elementos (cache local)
            ide = infNFe.find('ns:ide', NAMESPACES)
            emit = infNFe.find('ns:emit', NAMESPACES)
            dest = infNFe.find('ns:dest', NAMESPACES)
            total = infNFe.find('ns:total', NAMESPACES)
            ICMSTot = total.find('ns:ICMSTot', NAMESPACES) if total is not None else None
            infAdic = infNFe.find('ns:infAdic', NAMESPACES)
            
            enderEmit = emit.find('ns:enderEmit', NAMESPACES) if emit is not None else None
            enderDest = dest.find('ns:enderDest', NAMESPACES) if dest is not None else None
            
            # Chave da NF-e (otimizado)
            chNFe_val = ''
            protNFe = root.find('ns:protNFe', NAMESPACES)
            if protNFe is not None:
                infProt = protNFe.find('ns:infProt', NAMESPACES)
                if infProt is not None:
                    chNFe_val = self.get_xml_value(infProt, 'chNFe')
            if not chNFe_val and infNFe.get('Id'):
                chNFe_val = infNFe.get('Id')[3:]
            
            # Protocolo (otimizado)
            nProt_val = self.get_xml_value(infProt if protNFe is not None else None, 'nProt')
            dhRecbto_val = self.format_datetime_field(self.get_xml_value(infProt if protNFe is not None else None, 'dhRecbto'), '%d/%m/%Y %H:%M:%S')
            protocol_val = f'{nProt_val} - {dhRecbto_val}' if nProt_val and dhRecbto_val else (nProt_val or dhRecbto_val or '')

            # Replacements (cópia do original mas usando template_content)
            html_content = template_content.replace('[ds_company_issuer_name]', self.get_xml_value(emit, 'xNome'))
            html_content = html_content.replace('[nl_invoice]', self.get_xml_value(ide, 'nNF'))
            html_content = html_content.replace('[ds_invoice_serie]', self.get_xml_value(ide, 'serie'))
            html_content = html_content.replace('[url_logo]', '')
            html_content = html_content.replace('[ds_company_address]', f"{self.get_xml_value(enderEmit, 'xLgr')} {self.get_xml_value(enderEmit, 'nro')} {self.get_xml_value(enderEmit, 'xCpl')}".strip())
            html_content = html_content.replace('[ds_company_neighborhood]', self.get_xml_value(enderEmit, 'xBairro'))
            html_content = html_content.replace('[nu_company_cep]', self.get_xml_value(enderEmit, 'CEP'))
            html_content = html_content.replace('[ds_company_city_name]', self.get_xml_value(enderEmit, 'xMun'))
            html_content = html_content.replace('[ds_company_uf]', self.get_xml_value(enderEmit, 'UF'))
            html_content = html_content.replace('[nl_company_phone_number]', self.get_xml_value(enderEmit, 'fone'))
            html_content = html_content.replace('[ds_code_operation_type]', self.get_xml_value(ide, 'tpNF'))
            html_content = html_content.replace('[actual_page]', '1')
            html_content = html_content.replace('[total_pages]', '1')
            html_content = html_content.replace('{BarCode}', chNFe_val)
            html_content = html_content.replace('[ds_danfe]', chNFe_val)
            html_content = html_content.replace('[_ds_transaction_nature]', self.get_xml_value(ide, 'natOp'))
            html_content = html_content.replace('[protocol_label]', "PROTOCOLO DE AUTORIZAÇÃO DE USO" if chNFe_val else "")
            html_content = html_content.replace('[ds_protocol]', protocol_val)
            html_content = html_content.replace('[nl_company_ie]', self.get_xml_value(emit, 'IE'))
            html_content = html_content.replace('[nl_company_ie_st]', self.get_xml_value(emit, 'IEST'))
            html_content = html_content.replace('[nl_company_cnpj_cpf]', self.get_xml_value(emit, 'CNPJ'))
            html_content = html_content.replace('[ds_client_receiver_name]', self.get_xml_value(dest, 'xNome'))
            html_content = html_content.replace('[nl_client_cnpj_cpf]', self.get_xml_value(dest, 'CNPJ', default=self.get_xml_value(dest, 'CPF')))
            html_content = html_content.replace('[dt_invoice_issue]', self.format_datetime_field(self.get_xml_value(ide, 'dhEmi'), '%d.%m.%Y'))
            html_content = html_content.replace('[ds_client_address]', f"{self.get_xml_value(enderDest, 'xLgr')} {self.get_xml_value(enderDest, 'nro')} {self.get_xml_value(enderDest, 'xCpl')}".strip())
            html_content = html_content.replace('[ds_client_neighborhood]', self.get_xml_value(enderDest, 'xBairro'))
            html_content = html_content.replace('[nu_client_cep]', self.get_xml_value(enderDest, 'CEP'))
            html_content = html_content.replace('[dt_input_output]', self.format_datetime_field(self.get_xml_value(ide, 'dhSaiEnt', default=self.get_xml_value(ide, 'dhEmi'))))
            html_content = html_content.replace('[hr_input_output]', self.format_datetime_field(self.get_xml_value(ide, 'dhSaiEnt', default=self.get_xml_value(ide, 'dhEmi')), '%H:%M:%S'))
            html_content = html_content.replace('[ds_client_city_name]', self.get_xml_value(enderDest, 'xMun'))
            html_content = html_content.replace('[nl_client_phone_number]', self.get_xml_value(enderDest, 'fone'))
            html_content = html_content.replace('[ds_client_uf]', self.get_xml_value(enderDest, 'UF'))
            html_content = html_content.replace('[ds_client_ie]', self.get_xml_value(dest, 'IE'))
            html_content = html_content.replace('[tot_bc_icms]', self.get_xml_value(ICMSTot, 'vBC', is_currency=True))
            html_content = html_content.replace('[tot_icms]', self.get_xml_value(ICMSTot, 'vICMS', is_currency=True))
            html_content = html_content.replace('[tot_bc_icms_st]', self.get_xml_value(ICMSTot, 'vBCST', is_currency=True))
            html_content = html_content.replace('[tot_icms_st]', self.get_xml_value(ICMSTot, 'vST', is_currency=True))
            html_content = html_content.replace('[tot_icms_fcp]', self.get_xml_value(ICMSTot, 'vFCP', is_currency=True))
            html_content = html_content.replace('[vl_total_prod]', self.get_xml_value(ICMSTot, 'vProd', is_currency=True))
            html_content = html_content.replace('{ApproximateTax}', self.get_xml_value(ICMSTot, 'vTotTrib', is_currency=True))
            html_content = html_content.replace('[vl_shipping]', self.get_xml_value(ICMSTot, 'vFrete', is_currency=True))
            html_content = html_content.replace('[vl_insurance]', self.get_xml_value(ICMSTot, 'vSeg', is_currency=True))
            html_content = html_content.replace('[vl_discount]', self.get_xml_value(ICMSTot, 'vDesc', is_currency=True))
            html_content = html_content.replace('[vl_other_expense]', self.get_xml_value(ICMSTot, 'vOutro', is_currency=True))
            html_content = html_content.replace('[tot_total_ipi_tax]', self.get_xml_value(ICMSTot, 'vIPI', is_currency=True))
            html_content = html_content.replace('[vl_total]', self.get_xml_value(ICMSTot, 'vNF', is_currency=True))
            
            # Dados removidos (transporte, impostos serviços)
            html_content = html_content.replace('[ds_transport_carrier_name]', '')
            html_content = html_content.replace('[ds_transport_code_shipping_type]', '')
            html_content = html_content.replace('[ds_transport_rntc]', '')
            html_content = html_content.replace('[ds_transport_vehicle_plate]', '')
            html_content = html_content.replace('[ds_transport_vehicle_uf]', '')
            html_content = html_content.replace('[nl_transport_cnpj_cpf]', '')
            html_content = html_content.replace('[ds_transport_address]', '')
            html_content = html_content.replace('[ds_transport_city]', '')
            html_content = html_content.replace('[ds_transport_uf]', '')
            html_content = html_content.replace('[ds_transport_ie]', '')
            html_content = html_content.replace('[nu_transport_amount_transported_volumes]', '')
            html_content = html_content.replace('[ds_transport_type_volumes_transported]', '')
            html_content = html_content.replace('[ds_transport_mark_volumes_transported]', '')
            html_content = html_content.replace('[ds_transport_number_volumes_transported]', '')
            html_content = html_content.replace('[vl_transport_gross_weight]', '')
            html_content = html_content.replace('[vl_transport_net_weight]', '')
            html_content = html_content.replace('[ds_company_im]', self.get_xml_value(emit, 'IM'))
            html_content = html_content.replace('[vl_total_serv]', '')
            html_content = html_content.replace('[tot_bc_issqn]', '')
            html_content = html_content.replace('[tot_issqn]', '')
            html_content = html_content.replace('[ds_additional_information]', self._get_combined_additional_info(infAdic))
            html_content = html_content.replace('[text_consult_nfe]', 'Consulta de autenticidade no portal nacional da NF-e www.nfe.fazenda.gov.br/portal ou no site da Sefaz Autorizada.')
            html_content = html_content.replace('[page-break]', '')
            
            # ⚡ OTIMIZAÇÃO MASSA: Gerar código de barras
            barcode_html = ''
            if BARCODE_AVAILABLE and chNFe_val:
                barcode_base64 = self._generate_barcode_base64(chNFe_val)
                if barcode_base64:
                    barcode_html = f'<img src="{barcode_base64}" alt="Código de Barras da NF-e" style="max-width: 100%; height: auto;">'
            html_content = html_content.replace('[barcode_image]', barcode_html)

            # Processar itens e duplicatas
            items_html = self.format_items_html(nfe_element)
            html_content = html_content.replace('[items]', items_html)
            
            duplicates_html = self.format_duplicates_html(nfe_element)
            html_content = html_content.replace('[duplicates]', duplicates_html)

            # Remover seções desnecessárias
            html_content = self.remove_unnecessary_sections(html_content)

            # Gerar apenas PDF (otimizado)
            output_pdf_path = os.path.join(output_dir, pdf_filename)
            
            if not WEASYPRINT_AVAILABLE:
                return {"success": False, "message": "WeasyPrint não disponível"}

            try:
                # Conversão direta para PDF
                html_doc = WeasyHTML(string=html_content)
                html_doc.write_pdf(output_pdf_path)
                return {"success": True, "message": f"PDF gerado: {pdf_filename}"}
                
            except Exception as e:
                return {"success": False, "message": f"Erro PDF: {str(e)}"}

        except Exception as e:
            return {"success": False, "message": f"Erro: {str(e)}"}


class AppDanfe(ctk.CTk):
    """Interface gráfica para processamento de NF-e"""
    
    def __init__(self):
        super().__init__()
        
        self.processador = ProcessadorNFe()
        
        # Configurar janela principal
        self.title("Processador de NF-e - Gerador de DANFE")
        self.geometry("800x600")
        self.minsize(600, 400)
        
        # Variáveis
        self.xml_path = tk.StringVar()
        self.template_path = tk.StringVar(value="nfe_vertical.html")
        self.output_path = tk.StringVar(value=os.getcwd())
        
        self.setup_ui()
        
    def setup_ui(self):
        """Configura a interface do usuário"""
        # Frame principal
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Título
        title_label = ctk.CTkLabel(
            main_frame, 
            text="Gerador de DANFE a partir de XML NF-e",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        title_label.pack(pady=(20, 30))
        
        # Frame para seleção de arquivos
        files_frame = ctk.CTkFrame(main_frame)
        files_frame.pack(fill="x", padx=20, pady=10)
        
        # Arquivo XML
        xml_label = ctk.CTkLabel(files_frame, text="Arquivo XML da NF-e:", font=ctk.CTkFont(weight="bold"))
        xml_label.pack(anchor="w", padx=20, pady=(20, 5))
        
        xml_frame = ctk.CTkFrame(files_frame)
        xml_frame.pack(fill="x", padx=20, pady=(0, 10))
        
        self.xml_entry = ctk.CTkEntry(xml_frame, textvariable=self.xml_path, placeholder_text="Selecione o arquivo XML...")
        self.xml_entry.pack(side="left", fill="x", expand=True, padx=(10, 5), pady=10)
        
        xml_button = ctk.CTkButton(xml_frame, text="Procurar", command=self.select_xml_file, width=100)
        xml_button.pack(side="right", padx=(5, 10), pady=10)
        
        # Template HTML
        template_label = ctk.CTkLabel(files_frame, text="Template HTML:", font=ctk.CTkFont(weight="bold"))
        template_label.pack(anchor="w", padx=20, pady=(10, 5))
        
        template_frame = ctk.CTkFrame(files_frame)
        template_frame.pack(fill="x", padx=20, pady=(0, 10))
        
        self.template_entry = ctk.CTkEntry(template_frame, textvariable=self.template_path, placeholder_text="Selecione o template HTML...")
        self.template_entry.pack(side="left", fill="x", expand=True, padx=(10, 5), pady=10)
        
        template_button = ctk.CTkButton(template_frame, text="Procurar", command=self.select_template_file, width=100)
        template_button.pack(side="right", padx=(5, 10), pady=10)
        
        # Diretório de saída
        output_label = ctk.CTkLabel(files_frame, text="Diretório de saída:", font=ctk.CTkFont(weight="bold"))
        output_label.pack(anchor="w", padx=20, pady=(10, 5))
        
        output_frame = ctk.CTkFrame(files_frame)
        output_frame.pack(fill="x", padx=20, pady=(0, 20))
        
        self.output_entry = ctk.CTkEntry(output_frame, textvariable=self.output_path, placeholder_text="Selecione o diretório de saída...")
        self.output_entry.pack(side="left", fill="x", expand=True, padx=(10, 5), pady=10)
        
        output_button = ctk.CTkButton(output_frame, text="Procurar", command=self.select_output_dir, width=100)
        output_button.pack(side="right", padx=(5, 10), pady=10)
        
        # Botão processar
        process_button = ctk.CTkButton(
            main_frame, 
            text="Processar NF-e", 
            command=self.processar_nfe,
            font=ctk.CTkFont(size=16, weight="bold"),
            height=50
        )
        process_button.pack(pady=30)
        
        # Frame para mensagens
        self.messages_frame = ctk.CTkFrame(main_frame)
        self.messages_frame.pack(fill="both", expand=True, padx=20, pady=(10, 20))
        
        messages_label = ctk.CTkLabel(self.messages_frame, text="Mensagens:", font=ctk.CTkFont(weight="bold"))
        messages_label.pack(anchor="w", padx=10, pady=(10, 5))
        
        self.messages_text = ctk.CTkTextbox(self.messages_frame, state="disabled")
        self.messages_text.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
    def select_xml_file(self):
        """Seleciona arquivo XML"""
        filename = filedialog.askopenfilename(
            title="Selecionar arquivo XML da NF-e",
            filetypes=[("Arquivos XML", "*.xml"), ("Todos os arquivos", "*.*")]
        )
        if filename:
            self.xml_path.set(filename)
            
    def select_template_file(self):
        """Seleciona template HTML"""
        filename = filedialog.askopenfilename(
            title="Selecionar template HTML",
            filetypes=[("Arquivos HTML", "*.html"), ("Todos os arquivos", "*.*")]
        )
        if filename:
            self.template_path.set(filename)
            
    def select_output_dir(self):
        """Seleciona diretório de saída"""
        directory = filedialog.askdirectory(title="Selecionar diretório de saída")
        if directory:
            self.output_path.set(directory)
            
    def add_message(self, message: str, is_error: bool = False):
        """Adiciona mensagem ao campo de texto"""
        self.messages_text.configure(state="normal")
        timestamp = datetime.now().strftime("%H:%M:%S")
        prefix = "[ERRO]" if is_error else "[INFO]"
        full_message = f"{timestamp} {prefix} {message}\n"
        self.messages_text.insert("end", full_message)
        self.messages_text.configure(state="disabled")
        self.messages_text.see("end")
        
    def processar_nfe(self):
        """Processa a NF-e"""
        # Validar campos
        if not self.xml_path.get():
            messagebox.showerror("Erro", "Selecione um arquivo XML!")
            return
            
        if not self.template_path.get():
            messagebox.showerror("Erro", "Selecione um template HTML!")
            return
            
        if not self.output_path.get():
            messagebox.showerror("Erro", "Selecione um diretório de saída!")
            return
            
        # Verificar se arquivos existem
        if not os.path.exists(self.xml_path.get()):
            messagebox.showerror("Erro", "Arquivo XML não encontrado!")
            return
            
        if not os.path.exists(self.template_path.get()):
            messagebox.showerror("Erro", "Template HTML não encontrado!")
            return
            
        if not os.path.exists(self.output_path.get()):
            messagebox.showerror("Erro", "Diretório de saída não encontrado!")
            return
        
        # Limpar mensagens
        self.messages_text.configure(state="normal")
        self.messages_text.delete("1.0", "end")
        self.messages_text.configure(state="disabled")
        
        self.add_message("Iniciando processamento da NF-e...")
        
        # Processar
        result = self.processador.processar_nfe(
            self.xml_path.get(),
            self.template_path.get(),
            self.output_path.get()
        )
        
        if result["success"]:
            self.add_message("Processamento concluído com sucesso!")
            for line in result["message"].split("\n"):
                if line.strip():
                    self.add_message(line)
            messagebox.showinfo("Sucesso", "DANFE gerada com sucesso!")
        else:
            self.add_message(result["message"], is_error=True)
            messagebox.showerror("Erro", result["message"])


def main():
    """Função principal"""
    app = AppDanfe()
    app.mainloop()


if __name__ == "__main__":
    main() 