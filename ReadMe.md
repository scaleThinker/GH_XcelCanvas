# GH_XcelCanvas 📊

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Platform](https://img.shields.io/badge/Platform-Rhino%207%20%2F%208-black)](https://www.rhino3d.com/)
[![Language](https://img.shields.io/badge/Language-C%23-blue)](https://dotnet.microsoft.com/en-us/languages/csharp)

> **Select Language / Selecione o Idioma:**
> 
> [🇺🇸 English](#-english) | [🇧🇷 Português](#-português)

---

<div id="-english"></div>

## 🇺🇸 English

**GH_XcelCanvas** is a custom plugin for Grasshopper (Rhino 3D) designed to render Excel spreadsheets directly onto the Canvas. 

The main goal is to improve the parametric design workflow by eliminating the need to constantly switch windows (Alt-Tab) to check data. It provides a native "Viewport" for `.xlsx` files, allowing architects and engineers to visualize and select data cells graphically, similar to image viewers within the software.

### Key Features
- [x] **Native Reading:** Direct connection to local `.xlsx` files.
- [ ] **Canvas Rendering:** Visualizes the spreadsheet grid directly in Grasshopper (similar to *LB ImageViewer*).
- [ ] **Interactivity:** Select cells and ranges (e.g., A1:B10) via click-and-drag directly on the component.
- [ ] **Data Mapping:** Automatically outputs selected data as Grasshopper Data Trees.

### Tech Stack
This project is built using **C#** and the .NET Framework, focusing on performance and native integration.
* **Core:** RhinoCommon & Grasshopper SDK.
* **UI/UX:** Eto.Forms (for custom drawing on the canvas).
* **Data:** Microsoft.Office.Interop.Excel (Alpha) / Planned migration to ClosedXML/OpenXML.

### Roadmap
- **Phase 1:** Environment setup and headless data reading (Completed).
- **Phase 2:** Graphic visualization implementation (Custom Attributes) (In Progress).
- **Phase 3:** Mouse interaction logic and cell selection.
- **Phase 4:** Performance optimization (removing Excel dependency).

### Author
Developed by **ScaleThinker** (Brendo Tavares).
*Architect and Developer focused on parametric solutions.*

---

<div id="-português"></div>

## 🇧🇷 Português

**GH_XcelCanvas** é um plugin para Grasshopper (Rhino 3D) desenvolvido para renderizar planilhas do Excel diretamente no Canvas.

O objetivo é otimizar o fluxo de trabalho de design paramétrico, eliminando a necessidade de alternar janelas (Alt-Tab) constantemente para conferir dados. A ferramenta cria uma "Viewport" nativa para arquivos `.xlsx`, permitindo que arquitetos e engenheiros visualizem e selecionem células graficamente, similar a visualizadores de imagem dentro do software.

### Funcionalidades Principais
- [x] **Leitura Nativa:** Conexão direta com arquivos `.xlsx` locais.
- [ ] **Renderização no Canvas:** Visualiza a grade da planilha diretamente no Grasshopper (Estilo *LB ImageViewer*).
- [ ] **Interatividade:** Seleção de células e intervalos (ex: A1:B10) via clique e arraste no próprio componente.
- [ ] **Mapeamento de Dados:** Saída automática dos dados selecionados formatados em Data Trees.

### Tecnologias Utilizadas
Este projeto é desenvolvido em **C#** utilizando o framework .NET, focado em performance e integração nativa.
* **Core:** RhinoCommon & Grasshopper SDK.
* **UI/UX:** Eto.Forms (para desenho customizado no Canvas).
* **Dados:** Microsoft.Office.Interop.Excel (Alpha) / Migração planejada para ClosedXML/OpenXML.

### Roadmap
- **Fase 1:** Configuração do ambiente e leitura de dados "headless" (Concluído).
- **Fase 2:** Implementação da visualização gráfica (Custom Attributes) (Em andamento).
- **Fase 3:** Lógica de interação do mouse e seleção de células.
- **Fase 4:** Otimização de performance (remoção da dependência do Excel instalado).

### Autor
Desenvolvido por **ScaleThinker** (Brendo Tavares).
*Arquiteto e Desenvolvedor focado em soluções paramétricas.*

---
## License
This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
