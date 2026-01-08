import os
import re
import sys
import pandas as pd
import pdfplumber
import openpyxl
from datetime import datetime, date
from collections import defaultdict
import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
from pathlib import Path
from PIL import Image, ImageTk

# Configuração do tema
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# Função para ler o Excel de contatos
def carregar_contatos_excel(caminho_excel):
    contatos_dict = {}
    wb = openpyxl.load_workbook(caminho_excel)
    sheet = wb.active
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if len(row) >= 4:
            codigo, nome, nome_contato, nome_grupo = row[:4]
            # Converter código para inteiro e depois para string para remover .0
            codigo_limpo = str(int(float(codigo))) if codigo is not None else ""
            contatos_dict[codigo_limpo] = {
                'empresa': nome,
                'contato': nome_contato,
                'grupo': nome_grupo
            }
    return contatos_dict

# Função auxiliar para limpar e padronizar códigos
def limpar_codigo(codigo):
    """Converte código para string limpa, removendo .0 e espaços"""
    if codigo is None or pd.isna(codigo):
        return ""
    try:
        # Se for float com .0, remove o .0
        if isinstance(codigo, float) and codigo.is_integer():
            return str(int(codigo))
        # Se for string, remove espaços e .0 no final
        codigo_str = str(codigo).strip()
        if codigo_str.endswith('.0'):
            codigo_str = codigo_str[:-2]
        return codigo_str
    except:
        return str(codigo).strip()

# Funções de processamento para cada modelo
def processar_one(pasta_pdf, excel_entrada, excel_saida, log_callback, progress_callback):
    codigos_empresas = []
    # Aceita tanto "12-" quanto "12 -"
    padrao = r'^(\d+)\s*-'
    pdf_files = [f for f in os.listdir(pasta_pdf) if f.lower().endswith('.pdf')]
    log_callback(f"Encontrados {len(pdf_files)} arquivos PDF")
    progress_callback(0.2)

    for arquivo in pdf_files:
        match = re.match(padrao, arquivo)
        if match:
            codigo = match.group(1)
            codigos_empresas.append((codigo, arquivo))
            log_callback(f"Código encontrado: {codigo} - {arquivo}")
    
    progress_callback(0.4)
    log_callback("Lendo Excel de Contatos Onvio...")
    df_excel = pd.read_excel(excel_entrada)
    if df_excel.shape[1] < 4:
        raise ValueError("O arquivo Excel deve ter pelo menos 4 colunas (A-D).")
    df_excel.iloc[:, 0] = df_excel.iloc[:, 0].astype(str)
    
    progress_callback(0.6)
    log_callback("Comparando códigos e criando resultados...")
    resultados = []
    for codigo, arquivo_pdf in codigos_empresas:
        resultado = {
            'Código': codigo,
            'Empresa': '',
            'Contato Onvio': '',
            'Grupo Onvio': '',
            'Caminho': os.path.join(pasta_pdf, arquivo_pdf)
        }
        if codigo in df_excel.iloc[:, 0].values:
            linha = df_excel[df_excel.iloc[:, 0] == codigo].iloc[0]
            resultado.update({
                'Empresa': linha.iloc[1],
                'Contato Onvio': linha.iloc[2],
                'Grupo Onvio': linha.iloc[3]
            })
            log_callback(f"Correspondência encontrada para código {codigo}")
        else:
            log_callback(f"Código {codigo} não encontrado no Excel")
        resultados.append(resultado)
    
    progress_callback(0.8)
    log_callback("Salvando arquivo Excel de saída...")
    df_resultado = pd.DataFrame(resultados)
    df_resultado.to_excel(excel_saida, index=False)
    log_callback(f"Arquivo Excel gerado com sucesso: {excel_saida}")
    return len(resultados)

def verifica_certificado_cobranca(data_vencimento):
    hoje = date.today()
    dias_passados = (hoje - data_vencimento).days
    if dias_passados <= 6:
        return 1
    elif dias_passados <= 14:
        return 2
    elif dias_passados <= 19:
        return 3
    elif dias_passados <= 24:
        return 4
    elif dias_passados <= 30:
        return 5
    else:
        return 6

def processar_cobranca(caminho_pdf, excel_entrada, excel_saida, log_callback, progress_callback):
    contatos_dict = carregar_contatos_excel(excel_entrada)
    log_callback("Lendo arquivo PDF...")
    progress_callback(0.2)
    
    with pdfplumber.open(caminho_pdf) as pdf:
        texto_completo = ""
        for pagina in pdf.pages:
            texto_completo += pagina.extract_text()
    
    linhas_texto = texto_completo.split('\n')
    regex_cliente = re.compile(r'Cliente: (\d+)')
    regex_nome = re.compile(r'Nome: (.+)')
    regex_parcela = re.compile(r'(\d{2}/\d{2}/\d{4}) (\d{1,3}(?:\.\d{3})*,\d{2})')
    
    dados = defaultdict(list)
    codigo_atual = None
    empresa_atual = None
    
    progress_callback(0.4)
    log_callback("Extraindo informações do PDF...")
    for linha in linhas_texto:
        match_cliente = regex_cliente.search(linha)
        if match_cliente:
            codigo_atual = limpar_codigo(match_cliente.group(1))  # CORREÇÃO AQUI
            log_callback(f"Debug - Código extraído do PDF: '{codigo_atual}'")
        match_nome = regex_nome.search(linha)
        if match_nome and codigo_atual:
            empresa_atual = match_nome.group(1)
        match_parcela = regex_parcela.search(linha)
        if match_parcela and codigo_atual and empresa_atual:
            data_vencimento = str(match_parcela.group(1))
            valor_parcela = round(float(match_parcela.group(2).replace(".", "").replace(",",".")), 2)
            data_venci = datetime.strptime(data_vencimento, '%d/%m/%Y').date()
            carta = verifica_certificado_cobranca(data_venci)
            
            # CORREÇÃO: Debug para verificar busca no dicionário
            contato_info = contatos_dict.get(codigo_atual, {})
            log_callback(f"Debug - Buscando código '{codigo_atual}' no dicionário: {contato_info}")
            
            contato_individual = contato_info.get('contato', '')
            contato_grupo = contato_info.get('grupo', '')
            
            dados[codigo_atual].append({
                'Código': codigo_atual,
                'Empresa': empresa_atual,
                'Contato Onvio': contato_individual,
                'Grupo Onvio': contato_grupo,
                'Valor da Parcela': valor_parcela,
                'Data de Vencimento': data_vencimento,
                'Carta de Aviso': carta
            })
    
    linhas = []
    for codigo, info_list in dados.items():
        for info in info_list:
            linhas.append(info)
    
    progress_callback(0.8)
    log_callback("Salvando arquivo Excel de saída...")
    df = pd.DataFrame(linhas)
    df.to_excel(excel_saida, index=False)
    log_callback(f"Arquivo Excel gerado com sucesso: {excel_saida}")
    return len(linhas)

def processar_renovacao(excel_base, excel_entrada, excel_saida, log_callback, progress_callback):
    contatos_dict = carregar_contatos_excel(excel_entrada)
    log_callback("Lendo Excel Base...")
    progress_callback(0.2)

    df_comparacao = pd.read_excel(excel_base)
    codigos = df_comparacao.iloc[:, 0]
    pessoas = df_comparacao.iloc[:, 1]
    vencimentos = df_comparacao.iloc[:, 2]

    dados = {}
    hoje = datetime.now()

    progress_callback(0.4)
    log_callback("Comparando códigos e criando resultados...")

    for codigo_atual, pessoa, vencimento in zip(codigos, pessoas, vencimentos):
        codigo_str = limpar_codigo(codigo_atual)

        # Filtrar apenas contratos que ainda não venceram
        if pd.isna(vencimento) or not isinstance(vencimento, (pd.Timestamp, datetime)):
            continue
        if vencimento < hoje:
            continue

        contato_info = contatos_dict.get(codigo_str, {})
        contato_individual = contato_info.get('contato', '')
        contato_grupo = contato_info.get('grupo', '')

        vencimento_str = vencimento.strftime("%d/%m/%Y")

        if codigo_str not in dados:
            dados[codigo_str] = []
        dados[codigo_str].append({
            'Codigo': codigo_str,
            'Contato Onvio': contato_individual,
            'Grupo Onvio': contato_grupo,
            'Nome': pessoa,
            'Vencimento': vencimento_str
        })

    linhas = []
    for codigo, info_list in dados.items():
        for info in info_list:
            linhas.append(info)

    progress_callback(0.8)
    log_callback("Salvando arquivo Excel de saída...")
    df = pd.DataFrame(linhas)
    df.to_excel(excel_saida, index=False)
    log_callback(f"Arquivo Excel gerado com sucesso: {excel_saida}")
    return len(linhas)


def formatar_cnpj(cnpj):
    cnpj_str = re.sub(r'\D', '', str(cnpj))
    if cnpj_str.endswith('.0'):
        cnpj_str = cnpj_str[:-2]
    if len(cnpj_str) == 13:
        cnpj_str = '0' + cnpj_str
    elif len(cnpj_str) == 12:
        cnpj_str = '00' + cnpj_str
    return cnpj_str.zfill(14)

def verifica_certificado_comunicado(data_vencimento):
    hoje = datetime.today()
    dias_restantes = (data_vencimento - hoje).days
    if dias_restantes == 0:
        return 3
    elif 0 < dias_restantes <= 5:
        return 2
    elif dias_restantes > 5:
        return 1
    elif dias_restantes < 0:
        return 4
    else:
        return 0

def processar_comunicado(excel_base, excel_entrada, excel_saida, log_callback, progress_callback):
    contatos_dict = carregar_contatos_excel(excel_entrada)
    log_callback("Lendo Excel Base...")
    progress_callback(0.2)
    
    # CORREÇÃO: Log do dicionário de contatos para debug
    log_callback(f"Debug - Contatos carregados: {len(contatos_dict)} registros")
    log_callback(f"Debug - Primeiros 3 códigos do dicionário: {list(contatos_dict.keys())[:3]}")
    
    df_comparacao = pd.read_excel(excel_base)
    codigos = df_comparacao.iloc[:, 0]
    empresas = df_comparacao.iloc[:, 1]
    cnpjs = df_comparacao.iloc[:, 2]
    vencimentos = df_comparacao.iloc[:, 4]
    situacoes = df_comparacao.iloc[:, 7]
    
    dados = {}
    progress_callback(0.4)
    log_callback("Comparando códigos e criando resultados...")
    for codigo_atual, empresa, cnpj, vencimento, situacao in zip(codigos, empresas, cnpjs, vencimentos, situacoes):
        codigo_atual = limpar_codigo(codigo_atual)  # CORREÇÃO AQUI
        log_callback(f"Debug - Código do Excel Base: '{codigo_atual}' (tipo: {type(codigo_atual)})")
        
        if not pd.isna(cnpj):
            carta = verifica_certificado_comunicado(vencimento)
            cnpj_str = formatar_cnpj(cnpj)
            
            # CORREÇÃO: Debug para verificar busca no dicionário
            contato_info = contatos_dict.get(codigo_atual, {})
            log_callback(f"Debug - Buscando código '{codigo_atual}' no dicionário: {contato_info}")
            
            contato_individual = contato_info.get('contato', '')
            contato_grupo = contato_info.get('grupo', '')
            vencimento_str = vencimento.strftime("%d/%m/%Y") if isinstance(vencimento, pd.Timestamp) else str(vencimento)
            
            if codigo_atual not in dados:
                dados[codigo_atual] = []
            dados[codigo_atual].append({
                'Codigo': codigo_atual,
                'Empresa': empresa,
                'Contato Onvio': contato_individual,
                'Grupo Onvio': contato_grupo,
                'CNPJ': cnpj_str,
                'Vencimento': vencimento_str,
                'Carta de Aviso': carta
            })
    
    linhas = []
    for codigo, info_list in dados.items():
        for info in info_list:
            linhas.append(info)
    
    progress_callback(0.8)
    log_callback("Salvando arquivo Excel de saída...")
    df = pd.DataFrame(linhas)
    df.to_excel(excel_saida, index=False)
    log_callback(f"Arquivo Excel gerado com sucesso: {excel_saida}")
    return len(linhas)


def normalizar_nome(nome):
    """Normaliza nome da empresa para comparação (remove espaços extras, converte para minúsculo)"""
    if nome is None or pd.isna(nome):
        return ""
    return str(nome).strip().lower()


def calcular_similaridade(str1, str2):
    """Calcula a similaridade entre duas strings (0 a 1) usando SequenceMatcher"""
    from difflib import SequenceMatcher
    if not str1 or not str2:
        return 0.0
    return SequenceMatcher(None, str1, str2).ratio()


def buscar_por_similaridade(nome_busca, contatos_por_nome, limite_similaridade=0.8):
    """
    Busca um nome no dicionário de contatos por similaridade.
    Retorna o contato_info se encontrar correspondência >= limite_similaridade, senão None.
    """
    if not nome_busca:
        return None, 0.0

    melhor_match = None
    melhor_similaridade = 0.0

    for nome_contato, contato_info in contatos_por_nome.items():
        similaridade = calcular_similaridade(nome_busca, nome_contato)
        if similaridade >= limite_similaridade and similaridade > melhor_similaridade:
            melhor_similaridade = similaridade
            melhor_match = contato_info

    return melhor_match, melhor_similaridade


def processar_all(excel_origem, excel_contato, excel_saida, log_callback, progress_callback):
    """
    Modelo ALL: Compara Excel de Origem com Excel de Contato.
    Suporta comparação por código (coluna A) OU por nome da empresa (coluna A ou B).
    Mantém todos os registros do Excel de Origem, preenchendo Contato e Grupo quando houver correspondência.
    """
    log_callback("Lendo Excel de Origem...")
    progress_callback(0.2)

    # Ler Excel de Origem
    df_origem = pd.read_excel(excel_origem)
    log_callback(f"Registros no Excel de Origem: {len(df_origem)}")
    log_callback(f"Colunas encontradas: {df_origem.shape[1]}")

    progress_callback(0.4)
    log_callback("Lendo Excel de Contato...")

    # Ler Excel de Contato (4 colunas: código, nome, contato, grupo)
    df_contato = pd.read_excel(excel_contato)
    if df_contato.shape[1] < 4:
        raise ValueError("O Excel de Contato deve ter pelo menos 4 colunas (Código, Nome, Contato, Grupo).")

    log_callback(f"Registros no Excel de Contato: {len(df_contato)}")

    # Criar dicionários de contatos para busca rápida (por código e por nome)
    contatos_por_codigo = {}
    contatos_por_nome = {}

    for _, row in df_contato.iterrows():
        codigo = limpar_codigo(row.iloc[0])
        nome = normalizar_nome(row.iloc[1])
        contato_info = {
            'codigo': row.iloc[0],
            'nome': row.iloc[1] if pd.notna(row.iloc[1]) else '',
            'contato': row.iloc[2] if pd.notna(row.iloc[2]) else '',
            'grupo': row.iloc[3] if pd.notna(row.iloc[3]) else ''
        }

        if codigo:
            contatos_por_codigo[codigo] = contato_info
        if nome:
            contatos_por_nome[nome] = contato_info

    progress_callback(0.6)
    log_callback("Comparando registros e criando resultados...")

    # Obter nomes das colunas originais do Excel de Contato
    col_names = df_contato.columns.tolist()

    # Criar resultado com todos os registros do Excel de Origem
    resultados = []
    correspondencias_codigo = 0
    correspondencias_nome_exato = 0
    correspondencias_nome_similar = 0
    sem_correspondencia = 0

    for _, row in df_origem.iterrows():
        valor_coluna_a = row.iloc[0] if pd.notna(row.iloc[0]) else ''
        valor_coluna_b = row.iloc[1] if df_origem.shape[1] > 1 and pd.notna(row.iloc[1]) else ''

        # Tentar limpar como código
        codigo_limpo = limpar_codigo(valor_coluna_a)
        nome_normalizado_a = normalizar_nome(valor_coluna_a)
        nome_normalizado_b = normalizar_nome(valor_coluna_b)

        contato_info = None

        # 1. Tentar encontrar por código (coluna A)
        if codigo_limpo and codigo_limpo in contatos_por_codigo:
            contato_info = contatos_por_codigo[codigo_limpo]
            correspondencias_codigo += 1

        # 2. Se não encontrou por código, tentar por nome exato (coluna A)
        elif nome_normalizado_a and nome_normalizado_a in contatos_por_nome:
            contato_info = contatos_por_nome[nome_normalizado_a]
            correspondencias_nome_exato += 1

        # 3. Se não encontrou, tentar por nome exato (coluna B)
        elif nome_normalizado_b and nome_normalizado_b in contatos_por_nome:
            contato_info = contatos_por_nome[nome_normalizado_b]
            correspondencias_nome_exato += 1

        # 4. Se não encontrou exato, tentar por similaridade (coluna A) - 80%
        if not contato_info and nome_normalizado_a:
            contato_info, similaridade = buscar_por_similaridade(nome_normalizado_a, contatos_por_nome, 0.8)
            if contato_info:
                correspondencias_nome_similar += 1
                log_callback(f"Similaridade {similaridade:.0%}: '{valor_coluna_a}' -> '{contato_info['nome']}'")

        # 5. Se ainda não encontrou, tentar por similaridade (coluna B) - 80%
        if not contato_info and nome_normalizado_b:
            contato_info, similaridade = buscar_por_similaridade(nome_normalizado_b, contatos_por_nome, 0.8)
            if contato_info:
                correspondencias_nome_similar += 1
                log_callback(f"Similaridade {similaridade:.0%}: '{valor_coluna_b}' -> '{contato_info['nome']}'")

        if contato_info:
            resultados.append({
                col_names[0]: contato_info['codigo'],
                col_names[1]: contato_info['nome'],
                col_names[2]: contato_info['contato'],
                col_names[3]: contato_info['grupo']
            })
        else:
            # Sem correspondência - mantém dados originais com colunas em branco
            sem_correspondencia += 1
            resultados.append({
                col_names[0]: valor_coluna_a,
                col_names[1]: valor_coluna_b if valor_coluna_b else valor_coluna_a,
                col_names[2]: '',
                col_names[3]: ''
            })

    log_callback(f"Correspondências por código: {correspondencias_codigo}")
    log_callback(f"Correspondências por nome exato: {correspondencias_nome_exato}")
    log_callback(f"Correspondências por similaridade (>=80%): {correspondencias_nome_similar}")
    log_callback(f"Sem correspondência (colunas em branco): {sem_correspondencia}")

    progress_callback(0.8)
    log_callback("Salvando arquivo Excel de saída...")
    df_resultado = pd.DataFrame(resultados)
    df_resultado.to_excel(excel_saida, index=False)
    log_callback(f"Arquivo Excel gerado com sucesso: {excel_saida}")
    return len(resultados)


def processar_dombot(excel_base, excel_entrada, excel_saida, log_callback, progress_callback, periodo=""):
    # Nota: Este modelo não usa excel_entrada (Contatos Onvio), pois não utiliza contatos ou grupos
    log_callback("Lendo Excel Base...")
    progress_callback(0.2)
    
    # Ler o Excel base, sheet específica
    df = pd.read_excel(excel_base)
    
    # Renomear colunas para padronização
    df.columns = ['Nº', 'EMPRESAS', 'Tarefa']  # Ignorar a terceira coluna
    
    # Converter 'Nº' para string e limpar
    df['Nº'] = df['Nº'].apply(limpar_codigo)
    
    # Remover duplicatas baseadas em 'Nº' e 'EMPRESAS'
    df = df.drop_duplicates(subset=['Nº', 'EMPRESAS'])
    
    progress_callback(0.4)
    log_callback(f"Registros únicos encontrados: {len(df)}")
    
    # Obter Periodo e Competencia baseados no período fornecido ou data atual
    if periodo:
        try:
            mes, ano = periodo.split('/')
            periodo = f"{mes}/{ano}"
            competencia = f"{mes}{ano}"
            log_callback(f"Usando período customizado: {periodo}")
        except:
            raise ValueError("Formato de período inválido. Use MM/YYYY.")
    else:
        agora = datetime.now()
        periodo = agora.strftime("%m/%Y")
        competencia = agora.strftime("%m%Y")
        log_callback("Usando período atual (fallback)")
    
    # Adicionar colunas
    df['Periodo'] = periodo
    df['Competencia'] = competencia
    df['Salvar Como'] = df['Nº'] + '-' + df['EMPRESAS'] + '-' + df['Competencia']
    df['Caminho'] = df['Salvar Como'].apply(
    lambda x: fr"Z:\Pessoal\2025\GMS\{x}.pdf"
)

    
    # Reordenar colunas conforme especificado
    df = df[['Nº', 'EMPRESAS', 'Periodo', 'Salvar Como', 'Competencia', 'Caminho']]
    
    progress_callback(0.8)
    log_callback("Salvando arquivo Excel de saída...")
    df.to_excel(excel_saida, index=False)
    log_callback(f"Arquivo Excel gerado com sucesso: {excel_saida}")
    return len(df)

# Mapeamento de modelos para funções de processamento
processadores = {
    "ONE": processar_one,
    "Cobranca": processar_cobranca,
    "ProrContrato": processar_renovacao,
    "ComuniCertificado": processar_comunicado,
    "DomBot_GMS": processar_dombot,
    "ALL": processar_all
}

def get_resource_path(relative_path):
        """Retorna o caminho absoluto para arquivos, lidando com PyInstaller"""
        try:
            # PyInstaller cria uma pasta temporária e armazena o caminho em _MEIPASS
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")

        return os.path.join(base_path, relative_path)
    
class ExcelGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("M.E.G_ONE - Main Excel Generator ONE V1.2")
        self.root.geometry("700x500")
        self.root.resizable(False, False)
        
        self.pasta_pdf = ""
        self.excel_base = ""
        self.excel_entrada = ""
        self.excel_saida = ""
        self.modelo = ""
        
        self.setup_ui()
      
    
  
    def load_logo(self):
        """Carrega o logo se existir"""
        try:
            logo_path = get_resource_path("logo.png")  # pode ser .jpg também
            if os.path.exists(logo_path):
                image = Image.open(logo_path)
                image = image.resize((32, 32), Image.Resampling.LANCZOS)
                return ctk.CTkImage(light_image=image, dark_image=image, size=(80, 80))
            else:
                print("Logo não encontrado.")
                return None
        except Exception as e:
            print(f"Erro ao carregar logo: {e}")
            return None

        
    def setup_ui(self):
        # Container principal compacto
        main_frame = ctk.CTkFrame(self.root, corner_radius=10)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Header compacto
        header_frame = ctk.CTkFrame(main_frame, fg_color="transparent", height=50)
        header_frame.pack(fill="x", padx=15, pady=(10, 5))
        header_frame.pack_propagate(False)
        
        # Título com logo (se disponível)
        title_frame = ctk.CTkFrame(header_frame, fg_color="transparent")
        title_frame.pack(expand=True, fill="x")
        
        logo_image = self.load_logo()
        if logo_image:
            logo_label = ctk.CTkLabel(title_frame, image=logo_image, text="")
            logo_label.pack(side="left", padx=(0, 8))
        
        title_label = ctk.CTkLabel(
            title_frame,
            text="M.E.G_ONE",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        title_label.pack(side="left", anchor="w")
        
        # Seleção de modelo
        model_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        model_frame.pack(fill="x", padx=15, pady=5)
        
        ctk.CTkLabel(
            model_frame,
            text="Modelo:",
            font=ctk.CTkFont(size=12, weight="bold")
        ).pack(side="left", padx=(0, 8))
        
        self.modelo_combobox = ctk.CTkComboBox(
            model_frame,
            values=list(processadores.keys()),
            command=self.update_inputs,
            width=200,
            height=28
        )
        self.modelo_combobox.pack(side="left")
        
        # Frame para inputs dinâmicos
        self.inputs_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        self.inputs_frame.pack(fill="x", padx=15, pady=5)
        
        # Controles inferiores
        controls_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        controls_frame.pack(fill="x", padx=15, pady=5)
        
        # Botão processar
        self.process_button = ctk.CTkButton(
            controls_frame,
            text="🚀 Processar Relatórios",
            font=ctk.CTkFont(size=12, weight="bold"),
            height=35,
            command=self.process_files
        )
        self.process_button.pack(fill="x", pady=(0, 5))
        
        # Barra de progresso
        self.progress_bar = ctk.CTkProgressBar(controls_frame, height=8)
        self.progress_bar.pack(fill="x", pady=2)
        self.progress_bar.set(0)
        
        # Status
        self.status_label = ctk.CTkLabel(
            controls_frame,
            text="Selecione um modelo para começar",
            font=ctk.CTkFont(size=10),
            text_color="gray60"
        )
        self.status_label.pack(pady=2)
        
        # Log compacto
        log_frame = ctk.CTkFrame(main_frame, corner_radius=8)
        log_frame.pack(fill="both", expand=True, padx=15, pady=5)
        
        log_header = ctk.CTkFrame(log_frame, fg_color="transparent", height=30)
        log_header.pack(fill="x", padx=10, pady=(8, 0))
        log_header.pack_propagate(False)
        
        ctk.CTkLabel(
            log_header,
            text="📋 Log:",
            font=ctk.CTkFont(size=11, weight="bold")
        ).pack(side="left")
        
        ctk.CTkButton(
            log_header,
            text="Limpar",
            width=60,
            height=24,
            command=self.clear_log
        ).pack(side="right")
        
        # Área de log
        self.log_text = ctk.CTkTextbox(
            log_frame,
            font=ctk.CTkFont(size=9),
            height=100
        )
        self.log_text.pack(fill="both", expand=True, padx=10, pady=(2, 8))
        
        # Rodapé
        footer_label = ctk.CTkLabel(
            main_frame,
            text="© 2025 - Desenvolvido por Hugo",
            font=ctk.CTkFont(size=9),
            text_color="gray50"
        )
        footer_label.pack(pady=5)
        
        # Inicialização
        self.log_message("Sistema inicializado. Selecione um modelo para começar.")
    
    def create_compact_field(self, parent, label_text, button_text, command):
        """Cria um campo de entrada compacto"""
        field_frame = ctk.CTkFrame(parent, fg_color="transparent")
        field_frame.pack(fill="x", pady=2)
        
        # Label
        label = ctk.CTkLabel(
            field_frame,
            text=label_text,
            font=ctk.CTkFont(size=10, weight="bold"),
            width=120,
            anchor="w"
        )
        label.pack(side="left", padx=(0, 5))
        
        # Entry
        entry = ctk.CTkEntry(
            field_frame,
            placeholder_text="Nenhum arquivo selecionado",
            height=26,
            font=ctk.CTkFont(size=9)
        )
        entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        # Button
        button = ctk.CTkButton(
            field_frame,
            text=button_text,
            width=80,
            height=26,
            command=command
        )
        button.pack(side="right")
        
        return entry
    
    def update_inputs(self, choice):
        """Atualiza os campos de entrada baseado no modelo selecionado"""
        self.modelo = choice

        # Limpa campos anteriores
        for widget in self.inputs_frame.winfo_children():
            widget.destroy()

        # Cria campos específicos do modelo
        if choice == "ONE":
            self.pdf_entry = self.create_compact_field(
                self.inputs_frame,
                "📁 Pasta PDF:",
                "Selecionar",
                self.select_pdf_folder
            )
        elif choice == "Cobranca":
            self.pdf_entry = self.create_compact_field(
                self.inputs_frame,
                "📄 Arquivo PDF:",
                "Selecionar",
                self.select_pdf_file
            )
        elif choice == "ALL":
            # Modelo ALL: Excel de Origem e Excel de Contato
            self.excel_base_entry = self.create_compact_field(
                self.inputs_frame,
                "📊 Excel Origem:",
                "Selecionar",
                self.select_excel_base
            )
        else:
            self.excel_base_entry = self.create_compact_field(
                self.inputs_frame,
                "📊 Excel Base:",
                "Selecionar",
                self.select_excel_base
            )

        # Campos comuns
        if choice == "DomBot_GMS":
            # Campo para Período em vez de Contatos Onvio
            periodo_frame = ctk.CTkFrame(self.inputs_frame, fg_color="transparent")
            periodo_frame.pack(fill="x", pady=2)

            label = ctk.CTkLabel(
                periodo_frame,
                text="📅 Período (MM/YYYY):",
                font=ctk.CTkFont(size=10, weight="bold"),
                width=120,
                anchor="w"
            )
            label.pack(side="left", padx=(0, 5))

            self.periodo_entry = ctk.CTkEntry(
                periodo_frame,
                placeholder_text="Ex: 08/2025 (deixe vazio para atual)",
                height=26,
                font=ctk.CTkFont(size=9)
            )
            self.periodo_entry.pack(side="left", fill="x", expand=True)
        elif choice == "ALL":
            # Campo Excel de Contato para modelo ALL
            self.input_entry = self.create_compact_field(
                self.inputs_frame,
                "📋 Excel Contato:",
                "Selecionar",
                self.select_input_excel
            )
        else:
            # Campo normal de Contatos Onvio para outros modelos
            self.input_entry = self.create_compact_field(
                self.inputs_frame,
                "📋 Contatos Onvio:",
                "Selecionar",
                self.select_input_excel
            )
        
        self.output_entry = self.create_compact_field(
            self.inputs_frame, 
            "💾 Saída Excel:", 
            "Definir", 
            self.select_output_excel
        )
        
        self.status_label.configure(text="✅ Pronto para processar")
        self.log_message(f"Modelo selecionado: {choice}")
    
    def clear_log(self):
        """Limpa o log"""
        self.log_text.delete("1.0", "end")
        self.log_message("Log limpo")
    
    def select_pdf_folder(self):
        folder = filedialog.askdirectory(title="Selecionar pasta com arquivos PDF")
        if folder:
            self.pasta_pdf = folder
            self.pdf_entry.delete(0, "end")
            self.pdf_entry.insert(0, os.path.basename(folder))
            self.log_message(f"📁 Pasta selecionada: {folder}")
    
    def select_pdf_file(self):
        file = filedialog.askopenfilename(
            title="Selecionar arquivo PDF",
            filetypes=[("PDF files", "*.pdf")]
        )
        if file:
            self.pasta_pdf = file
            self.pdf_entry.delete(0, "end")
            self.pdf_entry.insert(0, os.path.basename(file))
            self.log_message(f"📄 PDF selecionado: {os.path.basename(file)}")
    
    def select_excel_base(self):
        file = filedialog.askopenfilename(
            title="Selecionar Excel Base",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file:
            self.excel_base = file
            self.excel_base_entry.delete(0, "end")
            self.excel_base_entry.insert(0, os.path.basename(file))
            self.log_message(f"📊 Excel Base: {os.path.basename(file)}")
    
    def select_input_excel(self):
        file = filedialog.askopenfilename(
            title="Selecionar Excel de Contatos Onvio",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file:
            self.excel_entrada = file
            self.input_entry.delete(0, "end")
            self.input_entry.insert(0, os.path.basename(file))
            self.log_message(f"📋 Contatos Onvio: {os.path.basename(file)}")
    
    def select_output_excel(self):
        file = filedialog.asksaveasfilename(
            title="Definir arquivo Excel de saída",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if file:
            self.excel_saida = file
            self.output_entry.delete(0, "end")
            self.output_entry.insert(0, os.path.basename(file))
            self.log_message(f"💾 Saída definida: {os.path.basename(file)}")
    
    def log_message(self, message):
        """Adiciona mensagem ao log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        formatted_message = f"[{timestamp}] {message}\n"
        self.log_text.insert("end", formatted_message)
        self.log_text.see("end")
        self.root.update_idletasks()
    
    def validate_inputs(self):
        """Valida se todos os campos necessários foram preenchidos"""
        if not self.modelo:
            messagebox.showerror("Erro", "Selecione um modelo.")
            return False

        if self.modelo == "ONE" and not self.pasta_pdf:
            messagebox.showerror("Erro", "Selecione a pasta com arquivos PDF.")
            return False

        if self.modelo == "Cobranca" and not self.pasta_pdf:
            messagebox.showerror("Erro", "Selecione o arquivo PDF.")
            return False

        if self.modelo in ["ProrContrato", "ComuniCertificado", "DomBot_GMS"] and not self.excel_base:
            messagebox.showerror("Erro", "Selecione o Excel Base.")
            return False

        if self.modelo == "ALL" and not self.excel_base:
            messagebox.showerror("Erro", "Selecione o Excel de Origem.")
            return False

        if self.modelo == "ALL" and not self.excel_entrada:
            messagebox.showerror("Erro", "Selecione o Excel de Contato.")
            return False

        if self.modelo not in ["DomBot_GMS", "ALL"] and not self.excel_entrada:
            messagebox.showerror("Erro", "Selecione o Excel de Contatos Onvio.")
            return False

        if not self.excel_saida:
            messagebox.showerror("Erro", "Defina o arquivo Excel de saída.")
            return False

        return True
    
    def process_files(self):
        """Inicia o processamento em thread separada"""
        if not self.validate_inputs():
            return
        
        self.process_button.configure(state="disabled")
        thread = threading.Thread(target=self.run_processing)
        thread.daemon = True
        thread.start()
    
    def run_processing(self):
        """Executa o processamento"""
        try:
            self.progress_bar.set(0)
            self.status_label.configure(text="🔄 Processando...")
            self.log_message("🚀 Iniciando processamento...")
            
            processador = processadores.get(self.modelo)
            if not processador:
                raise ValueError(f"Modelo {self.modelo} não encontrado.")
            
            input_file = self.pasta_pdf if self.modelo in ["ONE", "Cobranca"] else self.excel_base
            if self.modelo == "DomBot_GMS":
                periodo = self.periodo_entry.get().strip() if hasattr(self, 'periodo_entry') else ""
                total_registros = processador(
                    input_file, 
                    self.excel_entrada, 
                    self.excel_saida, 
                    self.log_message, 
                    self.progress_bar.set,
                    periodo=periodo
                )
            else:
                total_registros = processador(
                    input_file, 
                    self.excel_entrada, 
                    self.excel_saida, 
                    self.log_message, 
                    self.progress_bar.set
                )
            
            self.progress_bar.set(1.0)
            self.status_label.configure(text="✅ Processamento concluído!")
            self.log_message(f"🎉 Total de registros: {total_registros}")
            self.log_message("✅ Processamento finalizado!")
            
            messagebox.showinfo(
                "Sucesso", 
                f"Processamento concluído!\n\nTotal de registros: {total_registros}\n\nArquivo salvo em:\n{self.excel_saida}"
            )
        
        except Exception as e:
            self.progress_bar.set(0)
            self.status_label.configure(text="❌ Erro no processamento")
            self.log_message(f"❌ ERRO: {str(e)}")
            messagebox.showerror("Erro", f"Erro durante o processamento:\n{str(e)}")
        
        finally:
            self.process_button.configure(state="normal")

def main():
    root = ctk.CTk()
    
     # Adiciona ícone se estiver disponível
    try:
        def get_resource_path(relative_path):
            try:
                return os.path.join(sys._MEIPASS, relative_path)
            except:
                return os.path.join(os.path.abspath("."), relative_path)

        icon_path = get_resource_path("logoIcon.ico")
        if os.path.exists(icon_path):
            root.wm_iconbitmap(icon_path)
    except Exception as e:
        print(f"Erro ao definir ícone da interface: {e}")
        
    app = ExcelGeneratorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()