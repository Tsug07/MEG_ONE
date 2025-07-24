# 📊 M.E.G_ONE - Main Excel Generator ONE V1.0

**M.E.G_ONE** é uma ferramenta desenvolvida em Python com interface gráfica para automatizar a geração de planilhas Excel a partir de PDFs e arquivos `.xlsx`, de acordo com os modelos utilizados na automação **Automessenge ONE**.

> Desenvolvido por Hugo - © 2025

---

## 🚀 Funcionalidade

O sistema realiza a leitura e análise de arquivos PDF e Excel e gera automaticamente um arquivo Excel consolidado, conforme o modelo selecionado. Ele identifica dados relevantes, como códigos de clientes, vencimentos, valores e contatos, cruzando informações com um banco de contatos de referência.

---

## 🧩 Modelos Suportados

Atualmente, o M.E.G_ONE suporta os seguintes modelos:

- `ONE`: Lê uma pasta com PDFs e cruza os códigos com um Excel de contatos, gerando uma planilha de controle.
- `Cobranca`: Lê um PDF único com dados de cobrança, extrai cliente, nome, parcelas e vencimentos.
- `ProrContrato`: Realiza comparações entre arquivos Excel para geração de relatórios de renovação contratual.
- `ComuniCertificado`: Compara vencimentos e gera alertas com base no tempo restante ou vencido.

---

## 🖥️ Interface

A interface gráfica (GUI) foi construída com a biblioteca `customtkinter`, com suporte a temas escuros e modernos. Principais recursos:

- Seleção de modelo via dropdown
- Inputs dinâmicos de arquivos conforme o modelo
- Log de eventos em tempo real
- Barra de progresso visual
- Validação de campos obrigatórios

---

## 🛠️ Tecnologias Utilizadas

- Python 3.10+
- `pandas` – Manipulação de dados em Excel
- `openpyxl` – Leitura e escrita de planilhas `.xlsx`
- `pdfplumber` – Extração de texto de arquivos PDF
- `customtkinter` – Interface gráfica moderna
- `PIL` (Pillow) – Carregamento de imagem do logo
- `tkinter` – Seleção de arquivos e pastas

---

## 📂 Estrutura Esperada dos Arquivos

### Excel de Contatos Onvio
- Colunas obrigatórias:
  1. Código
  2. Nome da empresa
  3. Contato individual
  4. Grupo de contato

### Arquivos PDF
- Devem conter no nome ou conteúdo um código identificador (ex: `12345-relatorio.pdf`)
- No caso de cobrança, o conteúdo deve seguir o padrão reconhecido pelo sistema (ex: `Cliente: 12345`, `Nome: Empresa X`, etc.)

---

## ✅ Como Usar

1. Execute o script Python `M.E.G_ONE.py`
2. Selecione o **modelo de automação**
3. Selecione os arquivos conforme o modelo:
   - PDF ou Excel Base
   - Excel de contatos Onvio
   - Caminho para o Excel de saída
4. Clique em **"🚀 Processar Relatórios"**
5. O log e a barra de progresso indicarão o andamento
6. Ao final, o Excel gerado será salvo no local definido

---

## 📌 Observações

- Certifique-se de que os códigos nos arquivos estejam bem formatados (sem `.0`, espaços, etc.)
- Todos os modelos validam o conteúdo mínimo necessário antes do processamento
- O logo do sistema será carregado automaticamente se um dos arquivos `logo.png`, `logo.jpg`, etc., estiver presente na mesma pasta do script

---

## 🧪 Exemplo de Execução
```bash
python M.E.G_ONE.py
```
---

## 📄 Licença

Este projeto é de uso interno e desenvolvido exclusivamente para automações da **Automessenge ONE**.

---

## 👨‍💻 Autor

**Hugo**  
Desenvolvedor Python | Automação de Processos | Contato via Automessenge


