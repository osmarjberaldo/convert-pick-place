# PCB Pick and Place Converter

**Converte arquivos Altium Pick and Place para formato PC POCISAO JLCSMT STD (XLSX/XML) e BOM (XLSX).**

![License](https://img.shields.io/badge/license-MIT-green)
![Python](https://img.shields.io/badge/python-3.8+-blue)
![Platform](https://img.shields.io/badge/platform-Windows%20|%20Linux-lightgrey)

---

## 📋 Sobre

Aplicativo desktop com **2 abas** para automatizar a conversão de arquivos **Pick and Place** gerados pelo **Altium Designer**:

- **Aba 1 — 📍 Pick Place → POCISAO:** Extrai coordenadas X/Y, Layer e Rotação para planilha de montagem JLCSMT STD
- **Aba 2 — 📋 Pick Place → BOM:** Agrupa componentes por valor para lista de compras (BOM) no formato JLCPCB

### ✨ Funcionalidades

- **Auto-detecção de idioma** — identifica o idioma do sistema operacional (Português → PT-BR, outros → EN)
- **Busca integrada em 7 sites** — Octopart, SnapMagic, Mouser, JLCPCB Parts, DigiKey, SnapEDA e LCSC
- **Tooltips descritivos** em todos os botões
- **Preview em tabela** ordenável com edição inline (duplo-clique)
- **Barra de status** com feedback visual de operações
- **Progresso** com barra indeterminada durante conversões
- **Interface adaptável** — layout compacto com scroll por abas

---

## 🚀 Como Usar

### 1. Instalação

```bash
pip install -r requirements.txt
```

> **Nota:** O aplicativo também tenta instalar automaticamente as dependências ao iniciar.

### 2. Executar

```bash
python converter_app.py
```

O app detecta automaticamente o idioma do sistema:
- Sistema em **Português** → abre em PT-BR 🇧🇷
- Outros idiomas → abre em Inglês 🇺🇸

Para alternar manualmente: menu **Idioma** ou atalho **Ctrl+L**.

### 3. Aba 1 — POCISAO (Coordenadas)

Converte posições dos componentes para planilha de montagem.

**Mapeamento:**

| Pick and Place | POCISAO XLSX | Exemplo |
|---|---|---|
| Designator | Designator | C1, R1, U2 |
| Center-X(mm) | Mid X | 131.8260 mm |
| Center-Y(mm) | Mid Y | 73.3900 mm |
| Layer (TopLayer) | Layer (Top) | Top / Bottom |
| Rotation | Rotation | 0, 90, 180, 270 |

**Arquivo padrão:** `JLCSMT Convert New.xlsx`

**Formatos de saída:** XLSX (Excel) ou XML (com opção de metadados)

### 4. Aba 2 — BOM (Lista de Componentes)

Agrupa componentes com o mesmo valor (ex: todos resistores de 1K juntos).

**Mapeamento:**

| Pick and Place | BOM XLSX | Exemplo |
|---|---|---|
| Description | Comment | "RES 1K 0805" |
| Designator (agrupados) | Designator | "R4, R5, R6, R9, R11" |
| Comment (valor) | Footprint | "1K" |
| — | JLCPCB Part # | (opcional) |

**Arquivo padrão:** `Bom Convert New.xlsx`

### 🔍 Busca de Componentes

Na aba BOM, selecione um componente na tabela e clique em um dos **7 sites** para abrir a busca no navegador. Os botões ficam em uma única linha horizontal com o título acima.

| Site | URL de busca |
|---|---|
| **Octopart** | `octopart.com/search?q=` |
| **SnapMagic** | `snapmagic.com/search?q=` |
| **Mouser (BR)** | `br.mouser.com/c/?q=` |
| **JLCPCB Parts** | `jlcpcb.com/parts/componentSearch?searchTxt=` |
| **DigiKey** | `digikey.com.br/pt/products/result?s=` |
| **SnapEDA** | `snapeda.com/search/?q=` |
| **LCSC** | `lcsc.com/search?q=` |

### 📄 Formatos Suportados (Entrada)

| Aba | TXT | CSV | XLSX | XML |
|---|---|---|---|---|
| 📍 POCISAO | ✅ | ✅ | ✅ | ✅ |
| 📋 BOM | ✅ | ✅ | — | — |

### 📄 Formatos de Saída

| Aba | XLSX | XML |
|---|---|---|
| 📍 POCISAO | ✅ | ✅ |
| 📋 BOM | ✅ | — |

---

## 🖥️ Interface Gráfica

- **Auto-detecção de idioma** — PT-BR / EN conforme o sistema
- **2 abas** independentes com seleção de arquivo e preview
- **Espaçamento arejado** entre seções para melhor legibilidade
- **Tabela ordenável** com edição inline (clique duplo para editar)
- **7 botões de busca** em linha única na aba BOM, com título acima
- **Adicionar/Excluir linhas** na tabela
- **Barra de progresso** durante conversão
- **Barra de status** com feedback de operações
- **Tooltips** em todos os botões para facilitar o uso
- **Abertura automática** do arquivo gerado
- **Atalhos de teclado:**
  - `Ctrl+Q` — Sair
  - `Ctrl+L` — Alternar idioma
- **Menu de contexto** (clique direito) com edição, adicionar e excluir
- **Instalação automática** de dependências ao iniciar

---

## 🛠️ Tecnologias

- **Python 3.8+**
- **Tkinter** — Interface gráfica nativa
- **openpyxl** — Manipulação de arquivos Excel
- **PyInstaller** — Geração de executável (spec incluso)
- **Cross-platform** — Windows e Linux

---

## 👨‍💻 Autor

Criado por **DevOsmar**

![GitHub](https://img.shields.io/badge/GitHub-DevOsmar-181717?logo=github)
![Open Source](https://img.shields.io/badge/Open%20Source-%E2%9D%A4%EF%B8%8F-red)

Este projeto é **Open Source** — contribuições são bem-vindas!

---

## ☕ Doações

Se este projeto foi útil para você, considere fazer uma doação:

### PayPal (Internacional)

[![Doar com PayPal](https://img.shields.io/badge/Doar%20com-PayPal-00457C?logo=paypal)](https://www.paypal.com/donate/?hosted_button_id=HW9QS2TXX9ZYQ)

Aceita cartões de crédito e contas PayPal de qualquer lugar do mundo.

### PIX (Brasil)

Para doadores brasileiros — instantâneo e sem taxas.

**Chave PIX:** `osmarjunioberaldo@hotmail.com`

> 💡 No app, clique no botão "🇧🇷 Doar com PIX" que a chave é copiada automaticamente para a área de transferência.

---

## 📄 Licença

Licenciado sob **MIT** — veja o arquivo [LICENSE](LICENSE).
