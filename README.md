<!-- Cabeçalho -->

# Formulário Personalizado para Google Sheets

> [!NOTE]
> Este projeto é Copyright (c) 2025 Murilo Gritti. Todos os direitos reservados.

___

#### Ferramentas Utilizadas

![HTML](https://skillicons.dev/icons?i=html)
![CSS](https://skillicons.dev/icons?i=css)
![JavaScript](https://skillicons.dev/icons?i=javascript)
![App Script](https://skillicons.dev/icons?i=appscript)
![Google Sheets](https://skillicons.dev/icons?i=google-sheets)
___

<!-- Corpo do README -->
## Descrição

Este é um projeto feito sob demanda para uma construtora e terraplanagem e implementa um formulário HTML personalizado que envia dados diretamente para uma planilha do Google Sheets. Ele utiliza Google Apps Script para a lógica do backend (recebimento e processamento dos dados do formulário) e HTML/CSS/JavaScript para a interface do usuário do formulário.

O objetivo é fornecer uma alternativa mais flexível e visualmente customizável aos formulários padrão do Google Forms, permitindo uma integração direta com o Google Sheets para coleta e gerenciamento de dados.

### Funcionalidades
*   Interface de formulário web totalmente personalizável com HTML e CSS.
*   Submissão de dados de forma assíncrona para uma planilha Google Sheets.
*   Lógica de backend gerenciada por Google Apps Script, sem necessidade de servidor externo.
*   Validação de dados (pode ser implementada tanto no lado do cliente com JavaScript quanto no lado do servidor com Apps Script).
*   Exemplo de funções para buscar dados da planilha e popular campos do formulário (como em `anoref.gs`, `periodo.gs`, `pesquisa.gs`).
*   Cálculos e manipulações de dados na planilha via Apps Script (como em `calculoscolunasfinais.gs`).

### Estrutura do Projeto
*   **Arquivos `.gs` (Google Apps Script - Backend):**
    *   `codigo.gs`: Script principal, contém a função `doGet()` para servir a interface HTML e funções para receber os dados do formulário (`salvarDadosNoSheets`, `processarFormulario`, etc.). Também pode conter outras lógicas de negócios.
    *   `anoref.gs`, `periodo.gs`, `pesquisa.gs`: Contêm funções para buscar dados específicos da planilha e enviá-los para o formulário (ex: popular dropdowns).
    *   `consultaconferencia.gs`: Funções relacionadas à consulta e conferência de dados na planilha.
    *   `calculoscolunasfinais.gs`: Funções para realizar cálculos ou processamentos em colunas específicas da planilha após a submissão de dados.
*   **Arquivos `.html` (Frontend):**
    *   `index.html`: Arquivo HTML principal que estrutura o formulário. Utiliza `<?!= include('...'); ?>` para incorporar os outros arquivos HTML.
    *   `script-form.html`: Contém o código JavaScript do lado do cliente. Responsável por manipular o DOM, lidar com eventos do formulário (ex: `submit`), e comunicar-se com o backend do Apps Script usando `google.script.run`.
    *   `style-form.html`: Contém as regras de CSS para estilizar o formulário.

<!-- Corpo do README -->

## Sobre a Backpech:

A **[Backpech](https://www.instagram.com/back.pech/)** é uma empresa que presta serviços de TI particulares, como: manutenção e montagem de computadores, formatação, instalação de drivers e softwares, criação de sites, treinamentos básicos, entre outros. Mais informações podem ser encontradas em nossas redes sociais ou outros canais de contato listados abaixo.

<!-- Contato -->
### Redes de contato:

[![Instagram](https://skillicons.dev/icons?i=instagram)](https://www.instagram.com/back.pech/)
[![Discord](https://skillicons.dev/icons?i=discord)](https://discord.gg/b3zP3ArVJk)
[![E-mail](https://skillicons.dev/icons?i=gmail)](mailto:backpech.ctt@gmail.com)
[![Linkedin](https://skillicons.dev/icons?i=linkedin)](https://www.linkedin.com/in/backpech)
<!-- Contato -->
