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
- `Contato`: Cruza Excel base com Excel de contatos Onvio para gerar planilha consolidada de contatos.
- `ComuniCertificado`: Compara vencimentos e gera alertas com base no tempo restante ou vencido.
- `DomBot_GMS`: Lê Excel base do Domínio e gera planilha com período, competência e caminho dos documentos. Permite selecionar a pasta onde os documentos serão salvos pela automação.
- `DomBot_Econsig`: Lê PDF de Relação de Empréstimos Consignados, extrai empresas e gera planilha com datas e caminho dos documentos. Permite selecionar a pasta de destino dos documentos.
- `DomBot_Admiss`: Lê XLS de Relação de Empregados (admissões) e XLS de Contrato por Prazo Determinado, classifica tipo de contrato (Experiência/Indeterminado) e gera planilha com caminho dos documentos. Permite selecionar a pasta de destino dos documentos.
- `ALL`: Consolida dados de um Excel de origem com um Excel de contatos, gerando planilha unificada.
- `ALL_info`: Variante do ALL com informações adicionais na planilha de saída.

---

## 🖥️ Interface

A interface gráfica (GUI) foi construída com a biblioteca `customtkinter`, com suporte a temas escuros e modernos. Principais recursos:

- Seleção de modelo via dropdown
- Inputs dinâmicos de arquivos conforme o modelo
- Campo "Pasta Documentos" nos modelos DomBot (GMS, Econsig, Admiss) para definir onde os arquivos da automação serão salvos
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
- `calamine` – Engine para leitura de arquivos XLS legados

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
   - PDF ou Excel Base (dependendo do modelo)
   - Excel de contatos Onvio (quando aplicável)
   - Pasta Documentos (modelos DomBot — define o caminho dos arquivos gerados pela automação)
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


