# Integrador de Relatórios Técnicos em Excel

Este projeto automatiza a extração e integração de dados técnicos provenientes de relatórios em formato texto (.txt) para planilhas Excel (.xlsm). Ele foi desenvolvido para otimizar o processamento de grandes volumes de dados de simulações, garantindo precisão e agilidade na análise de resultados.

---

## 📂 Estrutura do Projeto

A organização dos arquivos segue um fluxo lógico de entrada, processamento e saída:

- **`Ralatorio_Emerg1` a `Ralatorio_Emerg7`**: Pastas que contêm os arquivos técnicos originais (.txt) fornecidos pelo solicitante.
- **`Resultado`**: Pasta de destino onde os arquivos Excel processados são salvos automaticamente.
- **`ListaBarra.txt`**: Arquivo de configuração que contém a lista de itens (barras) de interesse para a extração de dados.
- **`Valerio Macro.xlsm`**: Ferramenta principal contendo as macros VBA responsáveis pela lógica de integração.
- **`Valerio 1.xlsm`, `Valerio 2.xlsm`**: Modelos e bases de dados utilizados no processo.

---

## 🚀 Como Usar (Guia Passo-a-Passo)

Siga estas etapas para processar novos relatórios:

1.  **Preparação dos Dados**:
    - Coloque os arquivos `.txt` que deseja processar dentro das pastas de relatório correspondentes (ex: `Ralatorio_Emerg1`).
2.  **Configuração do Filtro**:
    - Abra o arquivo `ListaBarra.txt`.
    - Certifique-se de que os nomes das barras que você deseja extrair estão listados corretamente (um por linha).
3.  **Execução da Macro**:
    - Abra o arquivo `Valerio Macro.xlsm`.
    - **Importante**: Caso o Excel mostre um aviso de segurança, clique em **"Habilitar Conteúdo"**.
    - Localize e execute a macro de processamento (geralmente associada a um botão na planilha ou via `Alt + F8`).
4.  **Verificação dos Resultados**:
    - Após a conclusão da execução, verifique a pasta `Resultado`.
    - Novos arquivos `.xlsm` serão gerados contendo os dados extraídos e organizados de forma tabular.

---

## 🛠️ Detalhes Técnicos e Lógica de Processamento

O núcleo da automação reside no arquivo `VBA.bas`, que contém quatro sub-rotinas principais:

### 1. `Sub apaga()`
Uma função de limpeza que limpa todas as células da planilha (do intervalo A3 até o final), garantindo que dados de execuções anteriores não interfiram nos novos resultados.

### 2. `Sub Filtro()`
Aplica um filtro avançado na planilha "Base" utilizando a `ListaBarra.txt` (via `Application.Transpose`). Isso permite que o usuário visualize apenas os dados das barras de interesse definidas previamente.

### 3. `Sub Caso_Base()`
Esta é a rotina de inicialização. Ela:
- Extrai dados brutos de colunas de texto fixas usando a fórmula `EXT.TEXTO` (MID).
- Identifica categorias de linhas através de fórmulas lógicas complexas (`SES`, `DESLOC`).
- Cria as abas **"Base"** (para fluxos e capacidades) e **"Tensao"** (para níveis de tensão).
- Converte os dados em tabelas oficiais do Excel (`ListObjects`), facilitando consultas futuras.

### 4. `Sub Ocorrencia()`
Processa os relatórios subsequentes (contingências) e os integra ao Caso Base:
- Realiza a mesma extração de dados do Caso Base.
- Cria colunas dinâmicas ("Caso 1", "Caso 2", etc.) nas planilhas de destino.
- Utiliza `PROCV` (VLOOKUP) para correlacionar o carregamento e a tensão de cada ocorrência com os itens equivalentes no Caso Base.

---

## 📋 Pré-requisitos e Configuração

- **Microsoft Excel**: Versão compatível com Macros VBA.
- **Configuração Regional**: Atualmente otimizado para Excel em **Português** (devido ao uso de `FormulaLocal`).
- **Padrão de Relatório**: Os arquivos `.txt` devem seguir o padrão de largura fixa identificado na análise inicial (posições 1, 16, 24, 32, etc.).

---

## 📈 Avaliação Técnica e Próximos Passos (Roadmap)

A solução atual é altamente funcional e resolve o problema de processamento manual. As próximas fases planejadas para o projeto incluem:
- **Globalização**: Converter `FormulaLocal` para `Formula` (Inglês) para garantir suporte internacional.
- **Leitura Direta**: Implementar `FileSystemObject` no VBA para ler os arquivos `.txt` diretamente do disco, eliminando a dependência de RPA/Ctrl+C+Ctrl+V.
- **Tratamento de Erros**: Adicionar proteções contra nomes de abas duplicados ou arquivos corrompidos.

---

*Projeto documentado e preparado para versionamento inicial.*
