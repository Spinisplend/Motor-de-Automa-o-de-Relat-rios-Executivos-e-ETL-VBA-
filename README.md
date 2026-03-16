# Motor de Automação de Relatórios Executivos e ETL (VBA)

> **Nota de Conformidade e Engenharia:** Este repositório documenta a arquitetura de um pipeline de dados e automação de relatórios corporativos desenvolvido para um órgão governamental. Em conformidade com diretrizes de segurança (NDA), o código-fonte não é público.

## Visão Geral
Sistema avançado de automação de *Backoffice* (ETL - Extract, Transform, Load) arquitetado em VBA. O motor extrai grandes volumes de dados não estruturados de planilhas de monitoramento, aplica regras de negócio complexas, calcula indicadores de SLA e gera, de forma 100% autônoma, documentos executivos formatados (Microsoft Word) contendo tabelas analíticas e dashboards visuais.

## Stack Tecnológica
* **Linguagem:** Visual Basic for Applications (VBA)
* **Arquitetura:** Office Interop (Excel para Word/PowerPoint)
* **Processamento de Strings:** Lógica Fuzzy (Algoritmos Nativos)
* **Data Viz:** Instanciação dinâmica de `ChartObjects` (DOM do Excel)

## Arquitetura e Soluções de Engenharia

### 1. Tratamento de Dados Sujos via Lógica Fuzzy (Fuzzy String Matching)
O maior gargalo de automações corporativas é a inconsistência na entrada de dados humanos (nomes de colunas com erros de digitação ou variações de nomenclatura).
* **Solução Algorítmica:** Implementação nativa da **Distância de Levenshtein** e do **Índice de Jaccard** para realizar *Fuzzy Search*. O script não busca colunas por correspondência exata (Hardcoded), mas calcula a porcentagem de similaridade morfológica entre a string desejada e os cabeçalhos existentes, garantindo resiliência contra erros humanos de preenchimento.

### 2. Pipeline de Integração Cross-Application
Orquestração de instâncias fora do ambiente de execução nativo (`CreateObject("Word.Application")`).
* O motor controla o Microsoft Word em segundo plano (Background Process), manipulando o DOM do documento destino para injetar cabeçalhos, rodapés, margin limits e tabelas estilizadas de forma programática.

### 3. Geração Dinâmica de Dashboards (Data Viz)
O script elimina a necessidade de gráficos pré-formatados.
* A partir dos cálculos de rotatividade e percentuais de metas processados em memória (`Porcentagem()`), o motor instancia `ChartObjects` (Gráficos de Rosca e Linhas) em tempo de execução, formata legendas e eixos programaticamente, e realiza a transferência via *Clipboard* para o documento executivo final, garantindo relatórios visuais únicos por contexto.

### 4. Arquitetura Modular Baseada em Estados
O fluxo de geração do documento é fragmentado em uma máquina de casos (`Select Case Lp`), separando a extração de métricas operacionais, segurança e recursos humanos em submódulos independentes, facilitando manutenção estrutural e controle de memória (`Garbage Collection` de instâncias COM).
