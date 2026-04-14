# Motor de Processamento de Relatórios Técnicos CEPEL

Este projeto automatiza a extração e integração de dados técnicos provenientes de relatórios em formato texto (`.txt`) exportados pelo software CEPEL. Ele foi evoluído e estruturado para processar em massa e denormalizar em formato de tabela, convertendo bases textuais não padronizadas em relatórios analíticos limpos no Excel.

---

## 📂 Estrutura do Projeto (Versão 5.0 - Python)

O projeto abandonou a arquitetura em VBA (macro) em prol da alta performance da linguagem Python (utilizando `pandas`). A organização foi reescrita numa matriz modular capaz de analisar blocos de relatórios dinâmicos.

- **`Dados_Entrada/`**: Pasta central do leitor de pipeline. Todos os relatórios do CEPEL (lotes de até centenas de arquivos `.txt`) ficam aqui para processamento em massa.
- **`processador_relatorios.py`**: Script mestre. Lê o fluxo principal (Barra e Convergência) dos relatórios, formatando as planilhas densas (`Base.xlsx`).
- **`processador_impedancia.py`**: Módulo independente especializado em caçar seções de "*Relatório de Dados de Alteração de Impedância*", quebrando laços complexos em informações purificadas (`Impedancia.xlsx`).
- **`processador_controles.py`**: Módulo independente dedicado à aba remota de "*Relatório de Controles com Barra Controlada Desligada*", limpador de recuo de strings (`Controles.xlsx`).

---

## 🚀 Como Usar e Executar

A arquitetura não requer injeção de macros. Cada script gera consolidadores independentes:

1. **Requisito Previsto**: O computador deve ter o Python 3 instalado. Instale as dependências necessárias via terminal:
   ```bash
   pip install -r requirements.txt
   ```
2. **Abastecimento**: Jogue todos os arquivos `.txt` de saída do CEPEL dentro do diretório `/Dados_Entrada`.
3. **Execução Modular**:
   - Abra o terminal do seu computador (Bash/PowerShell) dentro da pasta principal do projeto.
   - Execute o script que quer consolidar. Exemplo: `python processador_relatorios.py`.
4. **Coleta de Analíticos**: Verifique na raiz da pasta que o sistema gerou automaticamente os documentos de Excel equivalentes perfeitamente formatados.

---

## 📈 Evolução Arquitetônica V5.0: "Motor Agnóstico a Páginas"

- **Leitura em Bloco Agnóstica**: O Parser Python destrói o conceito frágil de "número de página". Os scripts escaneiam os cabeçalhos literais para entender se a tabela que o script procura realmente existe no arquivo ou se a tabela extrapolou o tamanho da folha A4 em tela. Falsos positivos gerados por rebarbas de "SESSÕES" em outras quebras de página foram resolvidos.
- **Conversão de Tipos Transparente**: Em vez de exportar texto bruto, nosso código inspeciona cada coluna (ex: identificadores VS status nominal) e submete as conversões numéricas utilizando a tipagem flutuante (`pd.to_numeric()`). Quando abertos no MS Excel, os valores assumem magicamente a vírgula fracionária natural do padrão local (`PT-BR`).
- **Engenharia Reversa no CEPEL**: Anulamos manualmente a variação visual de espaços vazios induzidos nas fatias de visualização TXT, recuperando acrônimos que caiam acidentalmente nas descrições longas.

---

**Desenvolvido por Caio Cesar de Albuquerque**  
📫 [caioalbuquerquedev@gmail.com](mailto:caioalbuquerquedev@gmail.com)  
🔗 [LinkedIn](https://www.linkedin.com/in/caio-cesar-for-hire) | [GitHub](https://github.com/Caio-Cesa)