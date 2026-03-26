<p align="center">
  <img src="logo.png" alt="M.E.G ONE Logo" width="120">
</p>

<h1 align="center">M.E.G ONE</h1>
<p align="center"><strong>Main Excel Generator ONE</strong></p>

<p align="center">
  <img src="https://img.shields.io/badge/version-1.0-blue?style=for-the-badge" alt="Version">
  <img src="https://img.shields.io/badge/python-3.10+-3776AB?style=for-the-badge&logo=python&logoColor=white" alt="Python">
  <img src="https://img.shields.io/badge/platform-Windows-0078D6?style=for-the-badge&logo=windows&logoColor=white" alt="Platform">
  <img src="https://img.shields.io/badge/gui-CustomTkinter-2B2B2B?style=for-the-badge" alt="GUI">
  <img src="https://img.shields.io/badge/license-Internal_Use-red?style=for-the-badge" alt="License">
</p>

<p align="center">
  Ferramenta de automacao para geracao de planilhas Excel a partir de PDFs e arquivos <code>.xlsx/.xls</code>,<br>
  projetada para integrar com a automacao <strong>Automessenge ONE</strong>.
</p>

---

## Sobre

O **M.E.G ONE** realiza leitura e analise de arquivos PDF e Excel, gerando automaticamente planilhas consolidadas conforme o modelo selecionado. Ele identifica dados relevantes como codigos de clientes, vencimentos, valores, contatos e funcionarios, cruzando informacoes com bases de referencia (Excel de contatos Onvio, Google Sheets, etc.).

---

## Modelos Suportados

O sistema conta com **9 modelos** de processamento:

| Modelo | Entrada | Saida | Descricao |
|--------|---------|-------|-----------|
| **ONE** | Pasta de PDFs + Excel Contatos | Excel | Cruza codigos dos PDFs com contatos Onvio |
| **Cobranca** | PDF unico + Excel Contatos | Excel | Extrai dados de cobranca (cliente, parcelas, vencimentos) |
| **Contato** | Excel Base + Excel Contatos | Excel | Consolida dados de contatos Onvio |
| **ComuniCertificado** | Excel Base + Excel Contatos | Excel | Gera alertas de vencimento de certificados |
| **DomBot_GMS** | Excel Base (Dominio) | Excel | Gera planilha com periodo, competencia e caminho dos documentos |
| **DomBot_Econsig** | PDF (Emprestimos Consignados) | Excel | Extrai empresas do PDF e gera planilha com datas e caminho dos documentos |
| **DomBot_Admiss** | XLS Empregados + XLS Contrato Exp. | Excel | Classifica tipo de contrato (Experiencia/Indeterminado) e gera planilha |
| **ALL** | Excel Origem + Excel Contato | Excel | Consolida dados de origem com contatos em planilha unificada |
| **ALL_info** | Excel Origem + Excel Contato | Excel | Variante do ALL com informacoes adicionais |

> Os modelos **DomBot** (GMS, Econsig, Admiss) possuem o campo **Pasta Documentos**, que permite definir a pasta onde a automacao principal salvara os arquivos. Esse caminho e escrito na coluna de saida do Excel (Documento / Salvar Como / Caminho).

---

## Interface

<table>
  <tr>
    <td>

A interface grafica foi construida com **CustomTkinter** (tema escuro). Principais recursos:

- Selecao de modelo via dropdown com campos dinamicos
- Campos de entrada especificos por modelo (PDF, Excel, XLS, pastas)
- Campo **Pasta Documentos** nos modelos DomBot para definir caminho de destino
- Campos de data com preenchimento automatico (Econsig)
- Campo de periodo com formato MM/YYYY (GMS)
- Log de eventos em tempo real com timestamps
- Barra de progresso visual
- Validacao de campos obrigatorios antes do processamento
- Processamento em thread separada (interface nao trava)

</td>
  </tr>
</table>

---

## Stack Tecnologica

| Tecnologia | Uso |
|------------|-----|
| **Python 3.10+** | Linguagem principal |
| **pandas** | Manipulacao e analise de dados |
| **openpyxl** | Leitura e escrita de `.xlsx` |
| **pdfplumber** | Extracao de texto de PDFs |
| **customtkinter** | Interface grafica moderna (dark mode) |
| **Pillow (PIL)** | Carregamento de imagens (logo) |
| **tkinter** | Dialogos de selecao de arquivos e pastas |
| **calamine** | Engine para leitura de arquivos `.xls` legados |
| **difflib** | Similaridade de nomes para matching de empresas |
| **urllib** | Download de mapeamento de contatos via Google Sheets |

---

## Estrutura do Projeto

```
MEG_ONE/
├── M.E.G_ONE.py          # Aplicacao principal (~1600 linhas)
├── run_MEG_ONE.bat        # Script de inicializacao rapida
├── logo.png               # Logo exibido na interface
├── logoIcon.ico           # Icone da aplicacao
├── README.md
├── modelo_DomBot/         # Modelos e testes do DomBot
│   ├── DomBot_model.py
│   └── MEG_Test_1.py
└── Versoes Antigas/       # Historico de versoes anteriores
```

---

## Instalacao

### Pre-requisitos

```bash
python --version  # Python 3.10 ou superior
```

### Dependencias

```bash
pip install pandas openpyxl pdfplumber customtkinter pillow python-calamine
```

---

## Como Usar

### Via Script

```bash
python M.E.G_ONE.py
```

### Via Batch (Windows)

```bash
run_MEG_ONE.bat
```

### Passo a Passo

1. Execute a aplicacao
2. Selecione o **modelo de automacao** no dropdown
3. Preencha os campos que aparecerao conforme o modelo:

   | Campo | Quando aparece |
   |-------|---------------|
   | Pasta PDF | Modelo ONE |
   | Arquivo PDF | Modelos Cobranca, DomBot_Econsig |
   | Excel Base / XLS Empregados | Maioria dos modelos |
   | Contatos Onvio / XLS Contrato | Modelos que cruzam dados |
   | Pasta Documentos | Modelos DomBot (GMS, Econsig, Admiss) |
   | Periodo (MM/YYYY) | DomBot_GMS |
   | Data Inicial / Final | DomBot_Econsig |
   | Saida Excel | Todos os modelos |

4. Clique em **"Processar Relatorios"**
5. Acompanhe pelo **log** e pela **barra de progresso**
6. O Excel gerado sera salvo no local definido

---

## Formato dos Arquivos de Entrada

### Excel de Contatos Onvio

| Coluna A | Coluna B | Coluna C | Coluna D |
|----------|----------|----------|----------|
| Codigo | Nome da empresa | Contato individual | Grupo de contato |

### Arquivos PDF

- **ONE**: Nome do PDF deve iniciar com o codigo da empresa (ex: `12345 - relatorio.pdf`)
- **Cobranca**: Conteudo deve conter padroes como `Cliente:`, `Nome:`, `Parcela:`, `Vencimento:`
- **DomBot_Econsig**: Conteudo deve conter linhas `Empresa: CODIGO - NOME`

### Arquivos XLS (DomBot_Admiss)

- **XLS Empregados**: Relatorio "RELACAO DE EMPREGADOS I" exportado do Dominio
- **XLS Contrato Exp.**: Relatorio "Contrato por Prazo Determinado" exportado do Dominio

---

## Observacoes

- Codigos nos arquivos devem estar bem formatados (sem `.0`, espacos extras, etc.)
- Todos os modelos validam o conteudo minimo antes do processamento
- O logo e carregado automaticamente se `logo.png` ou `logo.jpg` estiver na pasta do script
- O mapeamento empresa-codigo no DomBot_Admiss e baixado automaticamente do Google Sheets
- Matching por similaridade (>=80%) e utilizado quando nao ha correspondencia exata de nomes

---

## Licenca

Projeto de uso interno, desenvolvido exclusivamente para automacoes da **Automessenge ONE**.

---

<p align="center">
  Desenvolvido por <strong>Hugo</strong><br>
  Automacao de Processos | Automessenge ONE<br>
  <sub>&copy; 2025</sub>
</p>
