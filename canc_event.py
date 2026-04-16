import xml.etree.ElementTree as ET
import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import zipfile
import logging

def extrair_dados_cte(arquivo_xml):
    tree = ET.parse(arquivo_xml)
    root = tree.getroot()
    ns = {'ns': 'http://www.portalfiscal.inf.br/cte'}
    
    # Se não for CT-e, ignora
    if root.find('.//ns:ide', ns) is None:
        return []

    # ==========================================
    # 1. IDENTIFICAÇÃO DE CANCELAMENTO
    # ==========================================
    cStat = ''
    for elemento in root.iter():
        if elemento.tag.endswith('}cStat'):
            cStat = elemento.text
            
    cancelado = 'SIM' if cStat == '101' else 'NÃO'

    # ==========================================
    # 2. DADOS PRINCIPAIS
    # ==========================================
    chave = ''
    chCTe_tag = root.find('.//ns:protCTe/ns:infProt/ns:chCTe', ns)
    if chCTe_tag is not None:
        chave = chCTe_tag.text
    else:
        infCte_tag = root.find('.//ns:infCte', ns)
        if infCte_tag is not None:
            chave = infCte_tag.attrib.get('Id', '')[3:]

    nCT = root.find('.//ns:ide/ns:nCT', ns).text if root.find('.//ns:ide/ns:nCT', ns) is not None else '0'
    
    dhEmi_tag = root.find('.//ns:ide/ns:dhEmi', ns)
    data_emissao = dhEmi_tag.text[:10] if dhEmi_tag is not None else ''
    
    cfop = root.find('.//ns:ide/ns:CFOP', ns).text if root.find('.//ns:ide/ns:CFOP', ns) is not None else ''
    nat_op = root.find('.//ns:ide/ns:natOp', ns).text if root.find('.//ns:ide/ns:natOp', ns) is not None else ''
    
    uf_origem = root.find('.//ns:ide/ns:UFIni', ns).text if root.find('.//ns:ide/ns:UFIni', ns) is not None else ''
    uf_destino = root.find('.//ns:ide/ns:UFFim', ns).text if root.find('.//ns:ide/ns:UFFim', ns) is not None else ''

    # ==========================================
    # 3. ENVOLVIDOS
    # ==========================================
    # Transportador (Emitente)
    cnpj_transp = root.find('.//ns:emit/ns:CNPJ', ns).text if root.find('.//ns:emit/ns:CNPJ', ns) is not None else ''
    nome_transp = root.find('.//ns:emit/ns:xNome', ns).text if root.find('.//ns:emit/ns:xNome', ns) is not None else ''

    # Remetente
    cnpj_rem_tag = root.find('.//ns:rem/ns:CNPJ', ns)
    if cnpj_rem_tag is None: cnpj_rem_tag = root.find('.//ns:rem/ns:CPF', ns)
    cnpj_rem = cnpj_rem_tag.text if cnpj_rem_tag is not None else ''
    nome_rem = root.find('.//ns:rem/ns:xNome', ns).text if root.find('.//ns:rem/ns:xNome', ns) is not None else ''

    # Destinatário
    doc_dest_tag = root.find('.//ns:dest/ns:CNPJ', ns)
    if doc_dest_tag is None: doc_dest_tag = root.find('.//ns:dest/ns:CPF', ns)
    doc_dest = doc_dest_tag.text if doc_dest_tag is not None else ''
    nome_dest = root.find('.//ns:dest/ns:xNome', ns).text if root.find('.//ns:dest/ns:xNome', ns) is not None else ''

    # Tomador
    tomador_nome = 'NÃO IDENTIFICADO'
    toma3 = root.find('.//ns:ide/ns:toma3/ns:toma', ns)
    toma4 = root.find('.//ns:ide/ns:toma4', ns)
    
    if toma3 is not None:
        tipo_toma = toma3.text
        if tipo_toma == '0': tomador_nome = nome_rem
        elif tipo_toma == '1': 
            exped_tag = root.find('.//ns:exped/ns:xNome', ns)
            tomador_nome = exped_tag.text if exped_tag is not None else ''
        elif tipo_toma == '2':
            receb_tag = root.find('.//ns:receb/ns:xNome', ns)
            tomador_nome = receb_tag.text if receb_tag is not None else ''
        elif tipo_toma == '3': tomador_nome = nome_dest
    elif toma4 is not None:
        toma_nome_tag = toma4.find('.//ns:xNome', ns)
        tomador_nome = toma_nome_tag.text if toma_nome_tag is not None else 'OUTROS'

    # ==========================================
    # 4. VALORES E NOTAS VINCULADAS
    # ==========================================
    vCarga_tag = root.find('.//ns:infCTeNorm/ns:infCarga/ns:vCarga', ns)
    vCarga = float(vCarga_tag.text) if vCarga_tag is not None else 0.0

    vFrete_tag = root.find('.//ns:vPrest/ns:vTPrest', ns)
    vFrete = float(vFrete_tag.text) if vFrete_tag is not None else 0.0

    chaves_tags = root.findall('.//ns:infCTeNorm/ns:infDoc/ns:infNFe/ns:chave', ns)
    if chaves_tags:
        chaves_nfe = ', '.join([c.text for c in chaves_tags])
    else:
        chaves_nfe = 'NENHUMA'

    # ==========================================
    # 5. IMPOSTOS
    # ==========================================
    cst = ''
    vBC = 0.0
    pICMS = 0.0
    vICMS = 0.0
    vICMSST = 0.0
    vICMS_OutraUF = 0.0
    
    icms_base = root.find('.//ns:imp/ns:ICMS', ns)
    if icms_base is not None:
        for tipo_icms in icms_base: 
            cst_tag = tipo_icms.find('ns:CST', ns)
            if cst_tag is not None: cst = cst_tag.text
                
            vbc_tag = tipo_icms.find('ns:vBC', ns)
            if vbc_tag is not None: vBC = float(vbc_tag.text)
                
            picms_tag = tipo_icms.find('ns:pICMS', ns)
            if picms_tag is not None: pICMS = float(picms_tag.text)
                
            vicms_tag = tipo_icms.find('ns:vICMS', ns)
            if vicms_tag is not None: vICMS = float(vicms_tag.text)
                
            vicmsst_tag = tipo_icms.find('ns:vICMSSTRet', ns)
            if vicmsst_tag is None: vicmsst_tag = tipo_icms.find('ns:vICMSST', ns)
            if vicmsst_tag is not None: vICMSST = float(vicmsst_tag.text)

            vicms_outra_uf_tag = tipo_icms.find('ns:vICMSOutraUF', ns)
            if vicms_outra_uf_tag is not None: vICMS_OutraUF += float(vicms_outra_uf_tag.text)

    icms_uf_fim = root.find('.//ns:imp/ns:ICMSUFFim', ns)
    if icms_uf_fim is not None:
        v_icms_uf_fim_tag = icms_uf_fim.find('ns:vICMSUFFim', ns)
        if v_icms_uf_fim_tag is not None: vICMS_OutraUF += float(v_icms_uf_fim_tag.text)

    # Dicionário com a exata ordem do Excel fornecido na imagem
    return [{
        'Chave_CTe': chave,
        'Numero': nCT,
        'Data_Emissao': data_emissao,
        'CFOP': cfop,
        'Nat_Operacao': nat_op,
        'UF_Origem': uf_origem,
        'UF_Destino': uf_destino,
        'CNPJ_Transportador': cnpj_transp,
        'Transportador': nome_transp,
        'CNPJ_Remetente': cnpj_rem,
        'Remetente': nome_rem,
        'CNPJ_CPF_Destinatario': doc_dest,
        'Destinatario': nome_dest,
        'Tomador_Nome': tomador_nome,
        'Valor_Carga': vCarga,
        'Valor_Frete': vFrete,
        'CST_ICMS': cst,
        'Base_Calculo': vBC,
        'Aliquota_ICMS': pICMS,
        'ICMS_Normal': vICMS,
        'ICMS_Retido_ST': vICMSST,
        'ICMS_Outra_UF': vICMS_OutraUF,
        'Chaves_NFe_Vinculadas': chaves_nfe,
        'Cancelado': cancelado  # <--- Único campo novo adicionado no final
    }]

# --- INTERFACE GRÁFICA ---
root = tk.Tk()
root.withdraw()

caminho_zip = filedialog.askopenfilename(
    title="Selecione o arquivo ZIP com os XMLs de CT-e",
    filetypes=[("Arquivos ZIP", "*.zip")]
)

if caminho_zip:
    lista_completa = []
    erros_leitura = 0
    pasta_destino = os.path.dirname(caminho_zip)
    
    # Configuração de Log
    caminho_log = os.path.join(pasta_destino, 'Log_Processamento_CTe.txt')
    logging.basicConfig(
        filename=caminho_log,
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        datefmt='%d/%m/%Y %H:%M:%S',
        filemode='w'
    )
    
    try:
        with zipfile.ZipFile(caminho_zip, 'r') as z:
            arquivos_xml = [f for f in z.namelist() if f.lower().endswith('.xml')]
            total_arquivos = len(arquivos_xml)
            
            if total_arquivos == 0:
                messagebox.showwarning("Aviso", "Não achei nenhum XML válido dentro do ZIP.")
            else:
                janela_progresso = tk.Toplevel()
                janela_progresso.title("Extração Master - CT-e")
                janela_progresso.geometry("450x180")
                
                lbl_status = tk.Label(janela_progresso, text="Iniciando a leitura dos arquivos...", font=("Arial", 10))
                lbl_status.pack(pady=15)
                
                barra = ttk.Progressbar(janela_progresso, orient="horizontal", length=350, mode="determinate")
                barra.pack(pady=5)
                barra["maximum"] = total_arquivos
                
                lbl_porcentagem = tk.Label(janela_progresso, text="0%", font=("Arial", 14, "bold"))
                lbl_porcentagem.pack()
                
                janela_progresso.update()
                
                for indice, nome_arquivo in enumerate(arquivos_xml):
                    try:
                        with z.open(nome_arquivo) as arquivo_xml:
                            dados_extraidos = extrair_dados_cte(arquivo_xml)
                            if dados_extraidos:
                                lista_completa.extend(dados_extraidos)
                    except Exception as e:
                        erros_leitura += 1
                        logging.error(f"Erro no arquivo '{nome_arquivo}': {str(e)}")
                        
                    if (indice + 1) % 50 == 0 or (indice + 1) == total_arquivos:
                        lbl_status.config(text=f"Processando ({indice+1}/{total_arquivos}): {nome_arquivo[-30:]}")
                        barra["value"] = indice + 1
                        porcentagem = int(((indice + 1) / total_arquivos) * 100)
                        lbl_porcentagem.config(text=f"{porcentagem}%")
                        janela_progresso.update()
                
                janela_progresso.destroy()
                
                if lista_completa:
                    df = pd.DataFrame(lista_completa)
                    caminho_excel = os.path.join(pasta_destino, 'Relatorio_CTe_Avancado.xlsx')
                    df.to_excel(caminho_excel, index=False)
                    
                    msg_final = f"Processo 100% Concluído!\n\nLidos {len(lista_completa)} CT-es no total."
                    if erros_leitura > 0:
                        msg_final += f"\nAtenção: {erros_leitura} arquivos falharam. Veja o arquivo de Log para detalhes."
                    msg_final += f"\n\nExcel salvo em:\n{caminho_excel}"
                    
                    messagebox.showinfo("Sucesso!", msg_final)
                else:
                    messagebox.showwarning("Aviso", "Nenhum dado de CT-e encontrado nos XMLs.")
                    
    except Exception as e:
        messagebox.showerror("Erro Crítico", f"Deu erro no processamento do ZIP:\n{e}")
else:
    print("Operação cancelada pelo usuário.")
