# PCB Pick and Place Converter

**Converte arquivos Altium Pick and Place para formatos JLCPCB e PCBWay: PCL, BOM e Centroid/SIMPLE.**

![License](https://img.shields.io/badge/license-MIT-green)
![Python](https://img.shields.io/badge/python-3.8+-blue)
![Platform](https://img.shields.io/badge/platform-Windows%20|%20Linux-lightgrey)

---

## 📋 Sobre

Aplicativo desktop com **4 abas** para automatizar a conversão de arquivos **Pick and Place** gerados pelo **Altium Designer**:

- **Aba 1 — 📍 PCL - JLCPCB:** extrai coordenadas X/Y, Layer e Rotação para planilha de montagem JLCSMT STD
- **Aba 2 — 📋 BOM - JLCPCB:** agrupa componentes por valor para lista de compras (BOM) no formato JLCPCB
- **Aba 3 — 📋 BOM - PCBWAY:** gera BOM PCBWay seguindo o modelo `Sample_BOM_PCBWay.xlsx`
- **Aba 4 — 📍 SIMPLE - PCBWAY:** gera arquivo Centroid/Simple PCBWay seguindo o exemplo oficial da PCBWay

### ✨ Funcionalidades

- **Auto-detecção de idioma** — identifica o idioma do sistema operacional (Português → PT-BR, outros → EN)
- **Busca integrada em 7 sites** — Octopart, SearchEngine, Mouser, JLCPCB Parts, DigiKey, SnapEDA e LCSC
- **Tooltips descritivos** em todos os botões
- **Preview em tabela** ordenável com edição inline (duplo-clique)
- **Campos com seleção controlada** — Layer `Top/Bottom` no JLCPCB e `TB` `T/B` no PCBWay
- **Entrada TXT/CSV/XLSX** nas abas BOM e PCBWay
- **Saída PCBWay em TXT, XLS, XLSX ou CSV**
- **Barra de status** com feedback visual de operações
- **Progresso** com barra indeterminada durante conversões
- **Interface adaptável** — layout compacto com scroll por abas

---

## 🚀 Como Usar

### 1. Instalação

> **Pré-requisito:** tenha o **Python 3** instalado no sistema antes de executar os comandos abaixo.

```bash
git clone https://github.com/osmarjberaldo/convert-pick-place.git
cd convert-pick-place
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

---

## 📍 Aba PCL - JLCPCB

Converte posições dos componentes para planilha de montagem JLCPCB/JLCSMT.

**Mapeamento:**

| Pick and Place | PCL XLSX/XML | Exemplo |
|---|---|---|
| Designator | Designator | C1, R1, U2 |
| Center-X(mm) | Mid X | 131.8260 mm |
| Center-Y(mm) | Mid Y | 73.3900 mm |
| Layer (TopLayer/BottomLayer) | Layer | Top / Bottom |
| Rotation | Rotation | 0, 90, 180, 270 |

**Arquivo padrão:** `JLCSMT Convert New.xlsx`

**Entrada:** TXT, CSV, XLSX ou XML  
**Saída:** XLSX ou XML

---

## 📋 Aba BOM - JLCPCB

Agrupa componentes com o mesmo valor (ex: todos resistores de 1K juntos).

**Mapeamento:**

| Pick and Place | BOM XLSX | Exemplo |
|---|---|---|
| Description | Comment | "RES 1K 0805" |
| Designator (agrupados) | Designator | "R4, R5, R6, R9, R11" |
| Comment (valor) | Footprint | "1K" |
| — | JLCPCB Part # | (opcional) |

**Arquivo padrão:** `Bom Convert New.xlsx`

**Entrada:** TXT, CSV ou XLSX  
**Saída:** XLSX

---

## 📋 Aba BOM - PCBWAY

Agrupa componentes e gera a BOM PCBWay seguindo o modelo `Sample_BOM_PCBWay.xlsx`.

**Mapeamento:**

| Pick and Place | BOM PCBWay | Exemplo |
|---|---|---|
| Ordem gerada | Item # | 1 |
| Designator (agrupados) | *Designator | R1,R2,R3 |
| Quantidade agrupada | *Qty | 3 |
| — | Manufacturer | (em branco) |
| Comment | *Mfg Part # | 10k |
| Description / Comment | Description / Value | Resistor 10k |
| Footprint | *Package/Footprint | 0805 |
| — | Type | SMD |
| — | Your Instructions / Notes | (em branco) |

**Arquivo padrão:** `PCBWay BOM.csv`

**Entrada:** TXT, CSV ou XLSX  
**Saída:** TXT, XLS, XLSX ou CSV

---

## 📍 Aba SIMPLE - PCBWAY

Gera arquivo Centroid/Simple PCBWay seguindo o exemplo oficial:

https://www.pcbway.com/img/images/pcbway/a-sample-of-PCBWay-Centroid-File.txt

**Colunas de saída:**

| PCBWay Centroid | Origem / Regra | Exemplo |
|---|---|---|
| Designator | Designator | R26 |
| Footprint | Footprint | 0603R |
| Mid X | Center-X(mm), convertido para mil | 660.236mil |
| Mid Y | Center-Y(mm), convertido para mil | 1603.937mil |
| Ref X | Mesmo valor de Mid X | 660.236mil |
| Ref Y | Mesmo valor de Mid Y | 1603.937mil |
| Pad X | Mesmo valor de Mid X | 660.236mil |
| Pad Y | Mesmo valor de Mid Y | 1603.937mil |
| TB | Top/Bottom convertido para T/B | T / B |
| Rotation | Rotation com 2 casas decimais | 360.00 |
| Comment | Comment | RES |

> **Observação:** arquivos Pick and Place do Altium normalmente não trazem coordenadas separadas para Ref/Pad. Por isso, a aba usa as coordenadas centrais como equivalente para `Ref X/Y` e `Pad X/Y`.

**Arquivo padrão:** `PCBWay Centroid.txt`

**Entrada:** TXT, CSV ou XLSX  
**Saída:** TXT, XLS, XLSX ou CSV

---

## 🔍 Busca de Componentes

Nas abas BOM/PCBWay, selecione um componente na tabela e clique em um dos **7 sites** para abrir a busca no navegador. Os botões ficam em uma única linha horizontal com o título acima.

| Site | URL de busca |
|---|---|
| **Octopart** | `octopart.com/search?q=` |
| **SearchEngine** | `componentsearchengine.com/search?term=` |
| **Mouser (BR)** | `br.mouser.com/c/?q=` |
| **JLCPCB Parts** | `jlcpcb.com/parts/componentSearch?searchTxt=` |
| **DigiKey** | `digikey.com.br/pt/products/result?s=` |
| **SnapEDA** | `snapeda.com/search/?q=` |
| **LCSC** | `lcsc.com/search?q=` |

---

## 📄 Formatos Suportados (Entrada)

| Aba | TXT | CSV | XLSX | XML |
|---|---|---|---|---|
| 📍 PCL - JLCPCB | ✅ | ✅ | ✅ | ✅ |
| 📋 BOM - JLCPCB | ✅ | ✅ | ✅ | — |
| 📋 BOM - PCBWAY | ✅ | ✅ | ✅ | — |
| 📍 SIMPLE - PCBWAY | ✅ | ✅ | ✅ | — |

## 📄 Formatos de Saída

| Aba | TXT | XLS | XLSX | XML | CSV |
|---|---|---|---|---|---|
| 📍 PCL - JLCPCB | — | — | ✅ | ✅ | — |
| 📋 BOM - JLCPCB | — | — | ✅ | — | — |
| 📋 BOM - PCBWAY | ✅ | ✅ | ✅ | — | ✅ |
| 📍 SIMPLE - PCBWAY | ✅ | ✅ | ✅ | — | ✅ |

---

## 🖥️ Interface Gráfica

- **Auto-detecção de idioma** — PT-BR / EN conforme o sistema
- **4 abas** independentes com seleção de arquivo e preview
- **Espaçamento arejado** entre seções para melhor legibilidade
- **Tabela ordenável** com edição inline (clique duplo para editar)
- **Campos select** para opções controladas como Layer e TB
- **7 botões de busca** em linha única nas abas BOM/PCBWay, com título acima
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
