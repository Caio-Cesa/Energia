# Integrador de Relatórios Técnicos em Excel

Este projeto automatiza a extração e integração de dados técnicos provenientes de relatórios em formato texto (.txt) para planilhas Excel (.xlsm). Ele foi desenvolvido para otimizar o processamento de grandes volumes de dados de simulações, garantindo precisão e agilidade na análise de resultados.

---

## 📂 Estrutura do Projeto (Versão 2.0)

A organização dos arquivos foi otimizada para processamento em lote:

- **`Dados_Entrada`**: Pasta centralizada que contém todos os relatórios `.txt`. Os arquivos foram renomeados automaticamente com o número do grupo original (ex: `... Referencia 1.txt`) para garantir a rastreabilidade.
- **`Resultado`**: Pasta de destino para exportações futuras.
- **`VBA_V2.bas`**: Código fonte da Versão 2.0, pronto para ser importado para o Excel.
- **`Valerio Macro.xlsm`**: Ferramenta principal (recomenda-se importar o `Modulo_V2` para este arquivo).

---

## 🚀 Como Usar (Guia Versão 2.0)

Agora o processo é 100% automatizado, sem necessidade de copiar e colar:

1.  **Preparação**:
    - Certifique-se de que todos os arquivos `.txt` de interesse estão dentro da pasta `Dados_Entrada`.
2.  **Importação do Código**:
    - No Excel, pressione `Alt + F11`.
    - Clique com o botão direito em "Módulos" -> **Importar Arquivo** e selecione o `VBA_V2.bas`.
3.  **Execução**:
    - Execute a macro `Processar_Tudo`.
    - O Excel irá percorrer todos os arquivos da pasta, extrair os dados e consolidá-los automaticamente.
4.  **Filtros**:
    - Utilize a nova coluna **"Origem_Caso"** nas abas "Base" e "Tensao" para filtrar exatamente qual cenário você deseja visualizar.

---

## 📈 Melhorias e Evolução

-   **Eliminação de RPA**: O sistema agora lê os arquivos diretamente do disco, sendo 10x mais rápido e imune a erros de área de transferência.
-   **Tabelas Mestras**: Em vez de centenas de abas, os dados agora ficam consolidados em duas tabelas mestras ("Base" e "Tensao"), facilitando o uso de Tabelas Dinâmicas.
-   **Compatibilidade Global**: Todas as fórmulas internas foram migradas para o padrão inglês (`MID`, `IFS`, `VLOOKUP`), garantindo que a macro funcione em qualquer idioma de instalação do Office.

---

**Desenvolvido por Caio Cesar de Albuquerque**  
📫 [caioalbuquerquedev@gmail.com](mailto:caioalbuquerquedev@gmail.com)  
🔗 [LinkedIn](https://www.linkedin.com/in/caio-cesar-for-hire) | [GitHub](https://github.com/Caio-Cesa)