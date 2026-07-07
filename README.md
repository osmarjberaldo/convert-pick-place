# PCB Pick and Place Converter

**Converte arquivos Altium Pick and Place para formatos de montagem e BOM para JLCPCB e PCBWay.**

![License](https://img.shields.io/badge/license-MIT-green)
![Python](https://img.shields.io/badge/python-3.8+-blue)
![Platform](https://img.shields.io/badge/platform-Windows%20|%20Linux-lightgrey)

---

## Sobre

Aplicativo desktop com **4 abas** para automatizar a conversao de arquivos **Pick and Place** gerados pelo **Altium Designer**:

- **PCL - JLCPCB:** extrai coordenadas X/Y, Layer e Rotacao para planilha de montagem JLCSMT STD.
- **BOM - JLCPCB:** agrupa componentes por valor para lista de compras no formato JLCPCB.
- **BOM - PCBWAY:** gera BOM PCBWay seguindo o modelo `Sample_BOM_PCBWay.xlsx`.
- **SIMPLE - PCBWAY:** gera arquivo centroid/posicao PCBWay seguindo o exemplo oficial da PCBWay.

### Funcionalidades

- Auto-deteccao de idioma: Portugues do Brasil ou Ingles conforme o sistema.
- Busca integrada em 7 sites: Octopart, SnapMagic, Mouser, JLCPCB Parts, DigiKey, SnapEDA e LCSC.
- Preview em tabela ordenavel com edicao inline por duplo clique.
- Campos com selecao controlada: Layer `Top/Bottom` no JLCPCB e `TB` `T/B` no PCBWay.
- Entrada em TXT, CSV e XLSX nas abas BOM/PCBWay.
- Barra de status, barra de progresso e abertura automatica do arquivo gerado.
- Interface adaptavel com scroll por abas.

---

## Como Usar

### 1. Instalacao

```bash
pip install -r requirements.txt
```

> Nota: o aplicativo tambem tenta instalar automaticamente as dependencias ao iniciar.

### 2. Executar

```bash
python converter_app.py
```

O app detecta automaticamente o idioma do sistema:

- Sistema em Portugues: abre em PT-BR.
- Outros idiomas: abre em Ingles.

Para alternar manualmente: menu **Idioma** ou atalho `Ctrl+L`.

---

## Abas

### PCL - JLCPCB

Converte posicoes dos componentes para planilha de montagem JLCPCB/JLCSMT.

| Pick and Place | PCL XLSX/XML | Exemplo |
|---|---|---|
| Designator | Designator | C1, R1, U2 |
| Center-X(mm) | Mid X | 131.8260 mm |
| Center-Y(mm) | Mid Y | 73.3900 mm |
| Layer | Layer | Top / Bottom |
| Rotation | Rotation | 0, 90, 180, 270 |

**Entrada:** TXT, CSV, XLSX ou XML  
**Saida:** XLSX ou XML  
**Arquivo padrao:** `JLCSMT Convert New.xlsx`

### BOM - JLCPCB

Agrupa componentes com o mesmo valor, como todos os resistores de `1K`.

| Pick and Place | BOM XLSX | Exemplo |
|---|---|---|
| Description | Comment | RES 1K 0805 |
| Designator agrupados | Designator | R4, R5, R6 |
| Comment | Footprint | 1K |
| Manual/opcional | JLCPCB Part # | Cxxxx |

**Entrada:** TXT, CSV ou XLSX  
**Saida:** XLSX  
**Arquivo padrao:** `Bom Convert New.xlsx`

### BOM - PCBWAY

Gera BOM PCBWay em CSV no formato do exemplo `Sample_BOM_PCBWay.xlsx`.

| Pick and Place | BOM PCBWay CSV |
|---|---|
| Ordem gerada | Item # |
| Designator agrupados | *Designator |
| Quantidade agrupada | *Qty |
| Vazio | Manufacturer |
| Comment | *Mfg Part # |
| Description ou Comment | Description / Value |
| Footprint | *Package/Footprint |
| Padrao gerado | Type |
| Vazio | Your Instructions / Notes |

**Entrada:** TXT, CSV ou XLSX  
**Saida:** TXT, XLS, XLSX ou CSV  
**Arquivo padrao:** `PCBWay BOM.csv`

### SIMPLE - PCBWAY

Gera arquivo centroid/posicao em TXT seguindo o exemplo oficial da PCBWay:

https://www.pcbway.com/img/images/pcbway/a-sample-of-PCBWay-Centroid-File.txt

| Coluna PCBWay | Origem / Regra |
|---|---|
| Designator | Designator |
| Footprint | Footprint |
| Mid X | Center-X(mm), convertido para mil |
| Mid Y | Center-Y(mm), convertido para mil |
| Ref X | Mesmo valor de Mid X |
| Ref Y | Mesmo valor de Mid Y |
| Pad X | Mesmo valor de Mid X |
| Pad Y | Mesmo valor de Mid Y |
| TB | `T` para Top, `B` para Bottom |
| Rotation | Rotation com 2 casas decimais |
| Comment | Comment |

> Observacao: os arquivos Pick and Place do Altium normalmente nao trazem coordenadas separadas para Ref/Pad. Por isso, a aba usa as coordenadas centrais como equivalente para `Ref X/Y` e `Pad X/Y`.

**Entrada:** TXT, CSV ou XLSX  
**Saida:** TXT, XLS, XLSX ou CSV  
**Arquivo padrao:** `PCBWay Centroid.txt`

---

## Busca de Componentes

Nas abas de BOM/PCBWay, selecione um componente na tabela e clique em um dos sites para abrir a busca no navegador.

| Site | URL de busca |
|---|---|
| Octopart | `octopart.com/search?q=` |
| SnapMagic | `snapmagic.com/search?q=` |
| Mouser (BR) | `br.mouser.com/c/?q=` |
| JLCPCB Parts | `jlcpcb.com/parts/componentSearch?searchTxt=` |
| DigiKey | `digikey.com.br/pt/products/result?s=` |
| SnapEDA | `snapeda.com/search/?q=` |
| LCSC | `lcsc.com/search?q=` |

---

## Formatos Suportados

### Entrada

| Aba | TXT | CSV | XLSX | XML |
|---|---|---|---|---|
| PCL - JLCPCB | Sim | Sim | Sim | Sim |
| BOM - JLCPCB | Sim | Sim | Sim | Nao |
| BOM - PCBWAY | Sim | Sim | Sim | Nao |
| SIMPLE - PCBWAY | Sim | Sim | Sim | Nao |

### Saida

| Aba | TXT | XLS | XLSX | XML | CSV |
|---|---|---|---|---|---|
| PCL - JLCPCB | Nao | Nao | Sim | Sim | Nao |
| BOM - JLCPCB | Nao | Nao | Sim | Nao | Nao |
| BOM - PCBWAY | Sim | Sim | Sim | Nao | Sim |
| SIMPLE - PCBWAY | Sim | Sim | Sim | Nao | Sim |

---

## Interface Grafica

- 4 abas independentes com selecao de arquivo e preview.
- Tabela ordenavel com edicao inline.
- Adicionar e excluir linhas manualmente.
- Menu de contexto com editar, adicionar e excluir.
- Tooltips nos botoes.
- Barra de status com feedback das operacoes.
- Atalhos de teclado:
  - `Ctrl+Q`: sair
  - `Ctrl+L`: alternar idioma

---

## Tecnologias

- Python 3.8+
- Tkinter
- openpyxl
- PyInstaller
- Windows e Linux

---

## Autor

Criado por **DevOsmar**.

Este projeto e Open Source. Contribuicoes sao bem-vindas.

---

## Doacoes

Se este projeto foi util para voce, considere fazer uma doacao.

### PayPal

[Doar com PayPal](https://www.paypal.com/donate/?hosted_button_id=HW9QS2TXX9ZYQ)

### PIX

**Chave PIX:** `osmarjunioberaldo@hotmail.com`

No app, clique no botao de PIX para copiar a chave automaticamente.

---

## Licenca

Licenciado sob MIT. Veja o arquivo [LICENSE](LICENSE).
