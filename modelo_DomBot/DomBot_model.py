import pandas as pd

# Parâmetros de entrada (ajuste conforme necessário)
base_excel_path = 'base.xlsx'  # Caminho do Excel base
output_excel_path = 'dombot_principal.xlsx'  # Caminho do Excel de saída
periodo = '07/2025'  # Valor fixo para Periodo (pode ser input do usuário)

# Passo 1: Ler o Excel base e selecionar as duas primeiras colunas
df = pd.read_excel(base_excel_path, usecols=[0, 1])  # Colunas 0 e 1 (A e B)
df.columns = ['Nº', 'EMPRESAS']  # Renomear para clareza

# Converter 'Nº' para string para evitar perda de formatação
df['Nº'] = df['Nº'].astype(str)

# Passo 2: Remover duplicatas baseadas em 'Nº' e 'EMPRESAS'
df = df.drop_duplicates(subset=['Nº', 'EMPRESAS'])

# Passo 3: Adicionar colunas
df['Periodo'] = periodo  # Valor fixo para todas as linhas
df['Competencia'] = periodo.replace('/', '')  # Converter Periodo para Competencia
df['Salvar Como'] = df['Nº'] + '-' + df['EMPRESAS'] + '-' + df['Competencia']  # Fórmula equivalente a =A3&"-"&B3&"-"&E3
df['Caminho'] = r"Z:\Pessoal\2025\GMS\" + df['Salvar Como'] + '.pdf'"  
# Fórmula equivalente a ="Z:\Pessoal\2025\GMS\"&D1&".pdf"

# Reordenar colunas conforme especificado
df = df[['Nº', 'EMPRESAS', 'Periodo', 'Salvar Como', 'Competencia', 'Caminho']]

# Passo 4: Salvar o novo Excel
df.to_excel(output_excel_path, index=False)
print(f'Excel principal gerado com sucesso: {output_excel_path}')