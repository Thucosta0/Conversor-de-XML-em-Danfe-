import xml.etree.ElementTree as ET
import re
from datetime import datetime

try:
    from weasyprint import HTML as WeasyHTML
    WEASYPRINT_AVAILABLE = True
except ImportError:
    WEASYPRINT_AVAILABLE = False

# Namespace usado no XML da NF-e
NAMESPACES = {'ns': 'http://www.portalfiscal.inf.br/nfe'}

def get_xml_value(element, path, default='', is_currency=False, unit_from_sibling=None, sibling_unit_tag=None):
    """
    Busca um valor em um elemento XML usando um caminho XPath.
    Retorna 'default' (string vazia) se não encontrado.
    Formata como moeda ou adiciona unidade se especificado.
    """
    try:
        # Lidar com o namespace nos caminhos
        if isinstance(path, str):
            path_parts = path.split('/')
            processed_parts = []
            for part in path_parts:
                if ':' not in part and '@' not in part: # Não adicionar ns a prefixos de namespace ou atributos
                    processed_parts.append(f'ns:{part}')
                else:
                    processed_parts.append(part)
            full_path = '/'.join(processed_parts)
        else: # path é uma lista (para múltiplos caminhos de fallback)
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
                    # Formata para R$ XX,XX (assumindo que o valor já está com ponto decimal)
                    num_value = float(value)
                    value = f'R$ {num_value:,.2f}'.replace(',', 'TEMP').replace('.', ',').replace('TEMP', '.')
                except ValueError:
                    pass # Mantém o valor original se não puder converter para float
            
            if unit_from_sibling and sibling_unit_tag:
                unit_element = element.find(f'ns:{unit_from_sibling}/ns:{sibling_unit_tag}', NAMESPACES)
                if unit_element is not None and unit_element.text:
                    value = f'{value} {unit_element.text.strip()}'
            elif unit_from_sibling and not sibling_unit_tag: # A unidade é o próprio valor do sibling_unit_tag
                 value = f'{value} {unit_from_sibling}'


            return value
        return default
    except Exception:
        return default

def format_datetime_field(dt_str, output_format='%d/%m/%Y'):
    """Formata string de data/hora (ISO 8601 com timezone) para o formato desejado."""
    if not dt_str:
        return ''
    try:
        # Remove o timezone para parsing, pois datetime.fromisoformat pode ter problemas com ele em algumas versões
        dt_str_no_tz = re.sub(r'[+-]\d{2}:\d{2}$', '', dt_str)
        dt_obj = datetime.fromisoformat(dt_str_no_tz)
        return dt_obj.strftime(output_format)
    except ValueError:
        return dt_str # Retorna original se não puder formatar

def format_items_html(root_nfe):
    """
    Formata os itens da NF-e como linhas de tabela HTML.
    """
    items_html_list = []
    infNFe = root_nfe.find('ns:infNFe', NAMESPACES)
    if infNFe is None:
        return ''

    for item_idx, det_element in enumerate(infNFe.findall('ns:det', NAMESPACES)):
        prod_element = det_element.find('ns:prod', NAMESPACES)
        imposto_element = det_element.find('ns:imposto', NAMESPACES)
        
        if prod_element is None or imposto_element is None:
            continue

        # Extração simplificada, precisará de mais detalhes para CSTs variáveis, etc.
        cst_icms = get_xml_value(imposto_element, 'ICMS/ICMS00/CST', default=get_xml_value(imposto_element, 'ICMS/ICMS10/CST', default=get_xml_value(imposto_element, 'ICMS/ICMS20/CST', default=get_xml_value(imposto_element, 'ICMS/ICMS40/CST', default=get_xml_value(imposto_element, 'ICMS/ICMS60/CST'))))) # Adicionar mais fallbacks se necessário
        v_bc_icms = get_xml_value(imposto_element, 'ICMS/ICMS00/vBC', default=get_xml_value(imposto_element, 'ICMS/ICMS10/vBC', default=get_xml_value(imposto_element, 'ICMS/ICMS20/vBC')), is_currency=True)
        p_icms = get_xml_value(imposto_element, 'ICMS/ICMS00/pICMS', default=get_xml_value(imposto_element, 'ICMS/ICMS10/pICMS', default=get_xml_value(imposto_element, 'ICMS/ICMS20/pICMS')), unit_from_sibling='%') # Adiciona %
        v_icms = get_xml_value(imposto_element, 'ICMS/ICMS00/vICMS', default=get_xml_value(imposto_element, 'ICMS/ICMS10/vICMS', default=get_xml_value(imposto_element, 'ICMS/ICMS20/vICMS')), is_currency=True)
        
        v_ipi = get_xml_value(imposto_element, 'IPI/IPITrib/vIPI', default='0.00', is_currency=True)
        p_ipi = get_xml_value(imposto_element, 'IPI/IPITrib/pIPI', default='0.00', unit_from_sibling='%')


        item_row = f"""
        <tr>
            <td style="text-align:center;">{get_xml_value(prod_element, 'cProd')}</td>
            <td>{get_xml_value(prod_element, 'xProd')}</td>
            <td style="text-align:center;">{get_xml_value(prod_element, 'NCM')}</td>
            <td style="text-align:center;">{cst_icms}</td>
            <td style="text-align:center;">{get_xml_value(prod_element, 'CFOP')}</td>
            <td style="text-align:center;">{get_xml_value(prod_element, 'uCom')}</td>
            <td style="text-align:right;">{float(get_xml_value(prod_element, 'qCom', default='0')):,.4f}</td>
            <td style="text-align:right;">{get_xml_value(prod_element, 'vUnCom', is_currency=True)}</td>
            <td style="text-align:right;">{get_xml_value(prod_element, 'vProd', is_currency=True)}</td>
            <td style="text-align:right;">{v_bc_icms}</td>
            <td style="text-align:right;">{v_icms}</td>
            <td style="text-align:right;">{v_ipi}</td>
            <td style="text-align:right;">{p_icms}</td>
            <td style="text-align:right;">{p_ipi}</td>
        </tr>
        """
        items_html_list.append(item_row)
    return '\n'.join(items_html_list)

def format_duplicates_html(root_nfe):
    """
    Formata as duplicatas (faturas) como HTML.
    """
    duplicates_html_list = []
    infNFe = root_nfe.find('ns:infNFe', NAMESPACES)
    if infNFe is None:
        return ''
    
    cobr_element = infNFe.find('ns:cobr', NAMESPACES)
    if cobr_element:
        fat_element = cobr_element.find('ns:fat', NAMESPACES)
        n_fat = get_xml_value(fat_element, 'nFat') if fat_element is not None else ''
        
        # Estrutura da tabela de duplicatas (pode precisar ajustar colunas/larguras)
        dup_table_start = '<table cellpadding="0" cellspacing="0" border="1" style="width:100%;"><tbody>'
        dup_table_end = '</tbody></table>'
        header_row = '' # O template original não mostra cabeçalho explícito para duplicatas soltas

        duplicates_html_list.append(dup_table_start)
        duplicates_html_list.append(header_row)

        for dup_element in cobr_element.findall('ns:dup', NAMESPACES):
            n_dup = get_xml_value(dup_element, 'nDup')
            d_venc = format_datetime_field(get_xml_value(dup_element, 'dVenc'))
            v_dup = get_xml_value(dup_element, 'vDup', is_currency=True)
            
            # Tenta usar nFat se nDup não estiver presente ou for parte de uma série
            display_dup_num = n_dup if n_dup else n_fat # Ou alguma combinação
            if n_fat and n_dup and not n_dup.startswith(n_fat): # Formato comum: NNFAT/PARCELA
                 display_dup_num = f'{n_fat}/{n_dup}'


            dup_row = f"""
            <tr>
                <td style="text-align:center; width: 33%;">{display_dup_num}</td>
                <td style="text-align:center; width: 33%;">{d_venc}</td>
                <td style="text-align:right; width: 34%;">{v_dup}</td>
            </tr>
            """
            duplicates_html_list.append(dup_row)
        
        if not cobr_element.findall('ns:dup', NAMESPACES) and fat_element is not None: # Caso só tenha <fat> e não <dup>
             v_liq = get_xml_value(fat_element, 'vLiq', is_currency=True)
             dup_row = f"""
             <tr>
                 <td style="text-align:center; width: 33%;">{n_fat}</td>
                 <td style="text-align:center; width: 33%;">-</td>
                 <td style="text-align:right; width: 34%;">{v_liq}</td>
             </tr>
             """
             duplicates_html_list.append(dup_row)


        duplicates_html_list.append(dup_table_end)
    
    if not duplicates_html_list or len(duplicates_html_list) <=2 : # apenas inicio e fim da tabela
        return '<table cellpadding="0" cellspacing="0" border="1" style="width:100%;"><tbody><tr><td style="text-align:center;">&nbsp;</td><td style="text-align:center;">&nbsp;</td><td style="text-align:right;">&nbsp;</td></tr></tbody></table>'


    return '\n'.join(duplicates_html_list)

def main():
    xml_file_path = '26250512420164001048550010003065961741974330-nfe.xml'
    html_template_path = 'nfe_vertical.html'
    output_html_path = 'danfe_gerada.html'
    output_pdf_path = 'danfe_gerada.pdf' # Novo caminho para o PDF

    try:
        tree = ET.parse(xml_file_path)
        root = tree.getroot()
    except FileNotFoundError:
        print(f"Erro: Arquivo XML não encontrado em {xml_file_path}")
        return
    except ET.ParseError:
        print(f"Erro: Falha ao parsear o arquivo XML {xml_file_path}")
        return

    try:
        with open(html_template_path, 'r', encoding='utf-8') as f:
            html_content = f.read()
    except FileNotFoundError:
        print(f"Erro: Arquivo de template HTML não encontrado em {html_template_path}")
        return

    # A raiz pode ser nfeProc ou diretamente NFe, vamos pegar NFe
    nfe_element = root.find('ns:NFe', NAMESPACES)
    if nfe_element is None:
        nfe_element = root # Tenta usar a raiz se NFe não for encontrado diretamente (caso o XML seja só o NFe)
        # Se ainda assim não encontrar infNFe, pode ser um problema.
        if nfe_element.find('ns:infNFe', NAMESPACES) is None:
             print("Erro: Tag <NFe> ou <infNFe> não encontrada no XML.")
             return

    infNFe = nfe_element.find('ns:infNFe', NAMESPACES)
    if infNFe is None:
        print("Erro: Tag <infNFe> não encontrada no XML.")
        return
        
    ide = infNFe.find('ns:ide', NAMESPACES)
    emit = infNFe.find('ns:emit', NAMESPACES)
    dest = infNFe.find('ns:dest', NAMESPACES)
    total = infNFe.find('ns:total', NAMESPACES)
    ICMSTot = total.find('ns:ICMSTot', NAMESPACES) if total is not None else None
    transp = infNFe.find('ns:transp', NAMESPACES)
    infAdic = infNFe.find('ns:infAdic', NAMESPACES)
    
    enderEmit = emit.find('ns:enderEmit', NAMESPACES) if emit is not None else None
    enderDest = dest.find('ns:enderDest', NAMESPACES) if dest is not None else None
    transporta = transp.find('ns:transporta', NAMESPACES) if transp is not None else None
    vol = transp.find('ns:vol', NAMESPACES) if transp is not None else None # pode haver múltiplos 'vol'
    
    # Chave da NF-e (tentar do protNFe primeiro, depois do infNFe Id)
    chNFe_val = ''
    protNFe = root.find('ns:protNFe', NAMESPACES) # Corrigido: buscar protNFe a partir da raiz (root)
    if protNFe is not None:
        infProt = protNFe.find('ns:infProt', NAMESPACES)
        if infProt is not None:
            chNFe_val = get_xml_value(infProt, 'chNFe')
    if not chNFe_val and infNFe.get('Id'):
        chNFe_val = infNFe.get('Id')[3:] # Remove "NFe" do início
    
    # Protocolo
    protocol_label_val = "PROTOCOLO DE AUTORIZAÇÃO DE USO" if chNFe_val else "" # Simples, pode ser melhorado
    nProt_val = get_xml_value(infProt if protNFe else None, 'nProt')
    dhRecbto_val = format_datetime_field(get_xml_value(infProt if protNFe else None, 'dhRecbto'), '%d/%m/%Y %H:%M:%S')
    protocol_val = f'{nProt_val} - {dhRecbto_val}' if nProt_val and dhRecbto_val else (nProt_val or dhRecbto_val or '')


    # Mapeamento de placeholders para valores
    # (Alguns valores podem precisar de concatenação ou lógica mais complexa)
    replacements = {
        '[ds_company_issuer_name]': get_xml_value(emit, 'xNome'),
        '[nl_invoice]': get_xml_value(ide, 'nNF'),
        '[ds_invoice_serie]': get_xml_value(ide, 'serie'),
        '[url_logo]': '', # Manter em branco ou definir um caminho fixo
        '[ds_company_address]': f"{get_xml_value(enderEmit, 'xLgr')} {get_xml_value(enderEmit, 'nro')} {get_xml_value(enderEmit, 'xCpl')}".strip(),
        '[ds_company_neighborhood]': get_xml_value(enderEmit, 'xBairro'),
        '[nu_company_cep]': get_xml_value(enderEmit, 'CEP'),
        '[ds_company_city_name]': get_xml_value(enderEmit, 'xMun'),
        '[ds_company_uf]': get_xml_value(enderEmit, 'UF'),
        '[nl_company_phone_number]': get_xml_value(enderEmit, 'fone'),
        '[ds_code_operation_type]': get_xml_value(ide, 'tpNF'), # 0-Entrada, 1-Saída
        '[actual_page]': '1', # Simples, para uma única página
        '[total_pages]': '1',
        '{BarCode}': chNFe_val, # Usar a chave de acesso para o código de barras
        '[ds_danfe]': chNFe_val,
        '[_ds_transaction_nature]': get_xml_value(ide, 'natOp'),
        '[protocol_label]': protocol_label_val,
        '[ds_protocol]': protocol_val,
        '[nl_company_ie]': get_xml_value(emit, 'IE'),
        '[nl_company_ie_st]': get_xml_value(emit, 'IEST'),
        '[nl_company_cnpj_cpf]': get_xml_value(emit, 'CNPJ'), # Assumindo CNPJ, pode precisar de lógica para CPF

        '[ds_client_receiver_name]': get_xml_value(dest, 'xNome'),
        '[nl_client_cnpj_cpf]': get_xml_value(dest, 'CNPJ', default=get_xml_value(dest, 'CPF')),
        '[dt_invoice_issue]': format_datetime_field(get_xml_value(ide, 'dhEmi')),
        '[ds_client_address]': f"{get_xml_value(enderDest, 'xLgr')} {get_xml_value(enderDest, 'nro')} {get_xml_value(enderDest, 'xCpl')}".strip(),
        '[ds_client_neighborhood]': get_xml_value(enderDest, 'xBairro'),
        '[nu_client_cep]': get_xml_value(enderDest, 'CEP'),
        '[dt_input_output]': format_datetime_field(get_xml_value(ide, 'dhSaiEnt', default=get_xml_value(ide, 'dhEmi'))), # dhSaiEnt ou fallback para dhEmi
        '[hr_input_output]': format_datetime_field(get_xml_value(ide, 'dhSaiEnt', default=get_xml_value(ide, 'dhEmi')), '%H:%M:%S'),
        '[ds_client_city_name]': get_xml_value(enderDest, 'xMun'),
        '[nl_client_phone_number]': get_xml_value(enderDest, 'fone'),
        '[ds_client_uf]': get_xml_value(enderDest, 'UF'),
        '[ds_client_ie]': get_xml_value(dest, 'IE'), # indIEDest pode ser usado para determinar se existe IE

        '[tot_bc_icms]': get_xml_value(ICMSTot, 'vBC', is_currency=True),
        '[tot_icms]': get_xml_value(ICMSTot, 'vICMS', is_currency=True),
        '[tot_bc_icms_st]': get_xml_value(ICMSTot, 'vBCST', is_currency=True),
        '[tot_icms_st]': get_xml_value(ICMSTot, 'vST', is_currency=True),
        '[tot_icms_fcp]': get_xml_value(ICMSTot, 'vFCP', is_currency=True),
        '[vl_total_prod]': get_xml_value(ICMSTot, 'vProd', is_currency=True), # No HTML é vl_total_prod, no XML vProd
        '{ApproximateTax}': get_xml_value(ICMSTot, 'vTotTrib', is_currency=True),
        '[vl_shipping]': get_xml_value(ICMSTot, 'vFrete', is_currency=True),
        '[vl_insurance]': get_xml_value(ICMSTot, 'vSeg', is_currency=True),
        '[vl_discount]': get_xml_value(ICMSTot, 'vDesc', is_currency=True),
        '[vl_other_expense]': get_xml_value(ICMSTot, 'vOutro', is_currency=True),
        '[tot_total_ipi_tax]': get_xml_value(ICMSTot, 'vIPI', is_currency=True),
        '[vl_total]': get_xml_value(ICMSTot, 'vNF', is_currency=True),

        '[ds_transport_carrier_name]': get_xml_value(transporta, 'xNome'),
        '[ds_transport_code_shipping_type]': get_xml_value(transp, 'modFrete'), # 0-Emitente, 1-Destinatário etc.
        '[ds_transport_rntc]': get_xml_value(transporta, 'RNTC'), # ANTT
        '[ds_transport_vehicle_plate]': get_xml_value(transp.find('ns:veicTransp', NAMESPACES) if transp is not None else None, 'placa'),
        '[ds_transport_vehicle_uf]': get_xml_value(transp.find('ns:veicTransp', NAMESPACES) if transp is not None else None, 'UF'),
        '[nl_transport_cnpj_cpf]': get_xml_value(transporta, 'CNPJ', default=get_xml_value(transporta, 'CPF')),
        '[ds_transport_address]': get_xml_value(transporta, 'xEnder'),
        '[ds_transport_city]': get_xml_value(transporta, 'xMun'),
        '[ds_transport_uf]': get_xml_value(transporta, 'UF'),
        '[ds_transport_ie]': get_xml_value(transporta, 'IE'),
        
        # Para 'vol', pode haver múltiplos. O template parece esperar um resumo.
        # Pegando do primeiro 'vol' se existir.
        '[nu_transport_amount_transported_volumes]': get_xml_value(vol, 'qVol'),
        '[ds_transport_type_volumes_transported]': get_xml_value(vol, 'esp'),
        '[ds_transport_mark_volumes_transported]': get_xml_value(vol, 'marca'),
        '[ds_transport_number_volumes_transported]': get_xml_value(vol, 'nVol'),
        '[vl_transport_gross_weight]': get_xml_value(vol, 'pesoB', unit_from_sibling='kg'), # Adiciona kg
        '[vl_transport_net_weight]': get_xml_value(vol, 'pesoL', unit_from_sibling='kg'), # Adiciona kg

        '[ds_company_im]': get_xml_value(emit, 'IM'),
        '[vl_total_serv]': get_xml_value(total.find('ns:ISSQNtot', NAMESPACES) if total is not None else None, 'vServ', is_currency=True),
        '[tot_bc_issqn]': get_xml_value(total.find('ns:ISSQNtot', NAMESPACES) if total is not None else None, 'vBC', is_currency=True),
        '[tot_issqn]': get_xml_value(total.find('ns:ISSQNtot', NAMESPACES) if total is not None else None, 'vISS', is_currency=True),
        
        '[ds_additional_information]': get_xml_value(infAdic, 'infCpl'),
        '[text_consult_nfe]': 'Consulta de autenticidade no portal nacional da NF-e www.nfe.fazenda.gov.br/portal ou no site da Sefaz Autorizada.', # Texto fixo
        '[page-break]': '' # Lógica de quebra de página não implementada
    }

    # Substituir placeholders
    for placeholder, value in replacements.items():
        # Garantir que o valor seja uma string para substituição
        html_content = html_content.replace(placeholder, str(value if value is not None else ''))

    # Substituir [items]
    items_html = format_items_html(nfe_element)
    html_content = html_content.replace('[items]', items_html)
    
    # Substituir [duplicates]
    duplicates_html_content = format_duplicates_html(nfe_element)
    html_content = html_content.replace('[duplicates]', duplicates_html_content)


    try:
        with open(output_html_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        print(f"DANFE HTML gerada com sucesso em: {output_html_path}")

        # Converter para PDF se WeasyPrint estiver disponível
        if WEASYPRINT_AVAILABLE:
            try:
                html_doc = WeasyHTML(string=html_content, base_url=output_html_path) # base_url ajuda a resolver caminhos relativos de CSS/imagens se houver
                html_doc.write_pdf(output_pdf_path)
                print(f"DANFE PDF gerada com sucesso em: {output_pdf_path}")
            except Exception as e:
                print(f"Erro ao converter HTML para PDF com WeasyPrint: {e}")
                print("Certifique-se de que a biblioteca WeasyPrint e suas dependências (Pango, Cairo, GDK-PixBuf) estão instaladas corretamente.")
        else:
            print("Biblioteca WeasyPrint não encontrada. Para gerar PDF, instale-a com: pip install WeasyPrint")

    except IOError:
        print(f"Erro: Não foi possível escrever o arquivo HTML em {output_html_path}")

if __name__ == '__main__':
    main() 