import pandas as pd
import os

def extrair(linha, inicio, fim):
    try:
        return linha[inicio:fim].strip()
    except:
        return ""

def processar_arquivo_controles(caminho):
    """Processa a seção de CONTROLES COM BARRA CONTROLADA DESLIGADA (Em qualquer página)."""
    registros = []
    sessao = "-"
    
    achou_tabela = False
    em_dados = False
    skip_header = False
    x_count = 0
    
    with open(caminho, 'r', encoding='windows-1252', errors='replace') as f:
        for linha in f:
            raw = linha.rstrip('\r\n')
            
            if not raw.strip():
                continue
                
            # Extração da SESSAO (Cabeçalho PARPEL em qualquer lugar do arquivo)
            if "PARPEL" in raw and "*" in raw:
                parts = raw.split("*")
                sessao = " ".join(p.strip() for p in parts[1:]) if len(parts) > 1 else "-"
                
            # Controle de seção principal
            if "RELATORIO DE CONTROLES COM BARRA CONTROLADA DESLIGADA" in raw:
                achou_tabela = True
                em_dados = False
                skip_header = True
                x_count = 0
                continue
                
            # Se já encontramos e processamos nossa tabela, ao esbarrar em outro relátorio nós paramos.
            if achou_tabela and "RELATORIO " in raw and "RELATORIO DE CONTROLES COM" not in raw:
                break
                
            # Se ainda não chegou na tabela desejada, ignora todas as linhas
            if not achou_tabela:
                continue
                
            # Identificação do delimitador X
            if skip_header:
                if raw.lstrip().startswith('X') and '---' in raw:
                    x_count += 1
                    if x_count >= 2:
                        em_dados = True
                        skip_header = False
                continue
                
            if not em_dados:
                continue
                
            # Ignoração de lixo após a entrada em 'em_dados' (cabeçalhos extras)
            if 'CEPEL' in raw or 'PAG.' in raw or 'PARPEL' in raw or raw.lstrip().startswith('X'):
                continue
                
            # Valida se a linha tem o padrão esperado (tem TIPO de controle e números)
            # Ex: '  CTAP 44906 PSTNIO-MT138  6940 PSTNIO-MT013  1 CONGELADO     '
            tipo = extrair(raw, 0, 6)
            # Muitas vezes pode ser apenas o complemento se houver espaço vazio na categoria TIPO
            num_de = extrair(raw, 6, 12)
            
            if num_de.isdigit() or tipo.isalpha():
                reg = {
                    'SESSAO': sessao,
                    'ARQUIVO': os.path.basename(caminho),
                    'TIPO': extrair(raw, 0, 6),
                    'NUM_DE': extrair(raw, 6, 12),
                    'NOME_DE': extrair(raw, 12, 26),
                    'NUM_PARA': extrair(raw, 26, 32),
                    'NOME_PARA': extrair(raw, 32, 46),
                    'GC': extrair(raw, 46, 48),
                    'ESTADO': extrair(raw, 48, 65)
                }
                registros.append(reg)
                
    return registros

if __name__ == "__main__":
    pasta = "Dados_Entrada"
    arquivos = sorted([f for f in os.listdir(pasta) if f.endswith(".txt")])
    
    if not arquivos:
        print("Nenhum arquivo .txt encontrado.")
    else:
        print(f"\n{'='*70}")
        print(f"  PROCESSADOR DE CONTROLES - CONSOLIDAÇÃO TOTAL (PAG 2)")
        print(f"  Arquivos encontrados: {len(arquivos)}")
        print(f"{'='*70}\n")
        
        todos_registros = []
        for idx, arq in enumerate(arquivos, 1):
            caminho = os.path.join(pasta, arq)
            print(f"  [{idx}/{len(arquivos)}] {arq}")
            regs = processar_arquivo_controles(caminho)
            todos_registros.extend(regs)
        
        df = pd.DataFrame(todos_registros)
        
        if not df.empty:
            # --- CONVERSÃO NUMÉRICA ---
            cols_numericas = ['NUM_DE', 'NUM_PARA', 'GC']
            
            for col in cols_numericas:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors='coerce')
                    
            output = "Controles.xlsx"
            df.to_excel(output, index=False)
            
            print(f"\n{'='*70}")
            print(f"  CONCLUÍDO!")
            print(f"  Total de registros extraídos: {len(df):,}")
            print(f"  Arquivo gerado: {output}")
            print(f"{'='*70}\n")
        else:
            print("\nNenhum registro de Controles (PAG 2) foi extraído de toda a base.")
