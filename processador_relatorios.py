import pandas as pd
import os

def extrair_valor(linha, start, length):
    """Auxiliar para extração segura de substrings."""
    try:
        return linha[start:start+length].strip()
    except:
        return "-"

def processar_relatorio_consolidado(caminho_arquivo):
    registros = []
    categoria_atual = "Caso Base"
    
    with open(caminho_arquivo, 'r', encoding='windows-1252', errors='replace') as f:
        linhas = f.readlines()
        
        i = 0
        while i < len(linhas):
            linha = linhas[i]
            
            # --- 1. Identificar Categoria (Sessão) ---
            if "  .............." in linha and i + 2 < len(linhas):
                categoria_atual = linhas[i+2][:30].strip()
            elif " X-------------X" in linha and i + 3 < len(linhas):
                categoria_atual = linhas[i+3][:30].strip()

            # --- 2. Lógica de Captura (Esquema Fixo de 2 Linhas) ---
            # Verificamos se a linha atual começa com um número (Possível NUM. da Barra)
            primeiro_campo = extrair_valor(linha, 0, 7)
            
            if primeiro_campo.isdigit():
                # Temos a Parte A! Agora buscamos a Parte B (Linha de baixo)
                if i + 1 < len(linhas):
                    linha_b = linhas[i+1]
                    
                    # Extração Parte A (Números)
                    num = primeiro_campo
                    kv = extrair_valor(linha, 7, 5)
                    tipo = extrair_valor(linha, 12, 5)
                    tensao = extrair_valor(linha, 17, 8)
                    geracao = extrair_valor(linha, 25, 10)
                    
                    # Extração Parte B (Nome e Detalhes)
                    nome = extrair_valor(linha_b, 0, 16)
                    ang = extrair_valor(linha_b, 16, 7)
                    mvar = extrair_valor(linha_b, 23, 9)
                    
                    # Validamos se a Parte B realmente parece um nome (e não uma linha vazia ou 'X--X')
                    if nome != "-" and "---" not in nome:
                        registro = {
                            "NUM": num,
                            "NOME": nome,
                            "KV": kv,
                            "TIPO": tipo,
                            "TENSÃO": tensao,
                            "GERACAO_MW": geracao,
                            "ANGULO": ang,
                            "Mvar": mvar,
                            "SESSÃO": categoria_atual,
                            "ARQUIVO_ORIGEM": os.path.basename(caminho_arquivo)
                        }
                        registros.append(registro)
                        i += 1 # Pula a linha B já processada
                
            i += 1
            
    return registros

if __name__ == "__main__":
    pasta_entrada = "Dados_Entrada"
    base_de_dados_geral = []
    
    print("\n--- INICIANDO PROCESSAMENTO CONSOLIDADO (v4.5) ---")
    
    arquivos_txt = [f for f in os.listdir(pasta_entrada) if f.endswith(".txt")]
    
    if not arquivos_txt:
        print(f"Erro: Nenhum arquivo encontrado em {pasta_entrada}")
    else:
        for idx, arquivo in enumerate(arquivos_txt):
            caminho = os.path.join(pasta_entrada, arquivo)
            print(f"[{idx+1}/{len(arquivos_txt)}] Processando: {arquivo}")
            
            novos_registros = processar_relatorio_consolidado(caminho)
            base_de_dados_geral.extend(novos_registros)
        
        # Criação do DataFrame Mestre
        if base_de_dados_geral:
            df_final = pd.DataFrame(base_de_dados_geral)
            
            # Limpeza e Tipagem: Tentar converter colunas numéricas
            cols_numericas = ["NUM", "TENSÃO", "GERACAO_MW", "ANGULO", "Mvar"]
            for col in cols_numericas:
                df_final[col] = pd.to_numeric(df_final[col], errors='coerce').fillna(0)
            
            # Exportação para Excel
            output_name = "Resultado_Final_Consolidado.xlsx"
            df_final.to_excel(output_name, index=False)
            
            print(f"\n==========================================")
            print(f"SUCESSO! Processamento concluído.")
            print(f"Total de registros extraídos: {len(df_final)}")
            print(f"Arquivo gerado: {output_name}")
            print(f"==========================================\n")
        else:
            print("Nenhum registro técnico foi extraído dos arquivos.")
