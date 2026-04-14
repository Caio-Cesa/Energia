import pandas as pd
import os

def extrair(linha, inicio, fim):
    """Extrai substring segura."""
    try:
        return linha[inicio:fim].strip()
    except:
        return ""

def processar_arquivo_impedancia(caminho):
    """Processa um arquivo focando na seção de ALTERACAO DE IMPEDANCIA (Em qualquer página)."""
    registros = []
    sessao = "-"
    
    achou_tabela = False
    em_dados = False
    skip_header = False
    x_count = 0
    
    with open(caminho, 'r', encoding='windows-1252', errors='replace') as f:
        for linha in f:
            raw = linha.rstrip('\r\n')
            
            # Pula linhas totalmente vazias
            if not raw.strip():
                continue
                
            # Extrai o nome da sessão (se parece ser o cabeçalho)
            if "PARPEL" in raw and "*" in raw:
                parts = raw.split("*")
                sessao = " ".join(p.strip() for p in parts[1:]) if len(parts) > 1 else "-"
                
            # --- CONTROLE DE SEÇÃO PRINCIPAL ---
            if "RELATORIO DE DADOS DE ALTERACAO DE IMPEDANCIA" in raw:
                achou_tabela = True
                em_dados = False
                skip_header = True
                x_count = 0
                continue
                
            # Se já encontramos e processamos nossa tabela, ao esbarrar em outro relátorio nós paramos.
            if achou_tabela and "RELATORIO " in raw and "RELATORIO DE DADOS DE ALTERACAO" not in raw:
                break
                
            # Se ainda não chegou na tabela desejada, ignora todas as linhas
            if not achou_tabela:
                continue
            
            # Skip de cabeçalho: espera 2 linhas X para entrar nos dados (exatamente como v5.0)
            if skip_header:
                if raw.lstrip().startswith('X') and '---' in raw:
                    x_count += 1
                    if x_count >= 2:
                        em_dados = True
                        skip_header = False
                continue
            
            if not em_dados:
                continue
            
            # --- DENTRO DA SEÇÃO DE DADOS ---
            # Ignora linhas de quebra de página ou lixo no meio dos dados
            if 'CEPEL' in raw or 'PAG.' in raw or 'PARPEL' in raw or raw.lstrip().startswith('X'):
                continue
            
            # As linhas de dados possuem o "NUM" da barra inicial. Verifica se inicia com padding e números.
            campo_num = raw[0:6].strip()
            if campo_num.isdigit():
                reg = {
                    'SESSAO': sessao,
                    'ARQUIVO': os.path.basename(caminho),
                    'DA_BARRA_NUM': campo_num,
                    'DA_BARRA_NOME': extrair(raw, 6, 19),
                    'P_BARRA_NUM': extrair(raw, 19, 24),
                    'P_BARRA_NOME': extrair(raw, 24, 37),
                    'NC': extrair(raw, 37, 40),
                    'ANTIGOS_R_%': extrair(raw, 40, 49),
                    'ANTIGOS_X_%': extrair(raw, 49, 58),
                    'NOVOS_R_%': extrair(raw, 58, 67),
                    'NOVOS_X_%': extrair(raw, 67, 76)
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
        print(f"  PROCESSADOR DE IMPEDÂNCIA - CONSOLIDAÇÃO TOTAL (PAG 1)")
        print(f"  Arquivos encontrados: {len(arquivos)}")
        print(f"{'='*70}\n")
        
        todos_registros = []
        for idx, arq in enumerate(arquivos, 1):
            caminho = os.path.join(pasta, arq)
            print(f"  [{idx}/{len(arquivos)}] {arq}")
            regs = processar_arquivo_impedancia(caminho)
            todos_registros.extend(regs)
        
        df = pd.DataFrame(todos_registros)
        
        if not df.empty:
            # --- CONVERSÃO NUMÉRICA ---
            cols_numericas = [
                'DA_BARRA_NUM', 'P_BARRA_NUM', 'NC', 
                'ANTIGOS_R_%', 'ANTIGOS_X_%', 'NOVOS_R_%', 'NOVOS_X_%'
            ]
            for col in cols_numericas:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors='coerce')
            
            output = "Impedancia.xlsx"
            df.to_excel(output, index=False)
            
            print(f"\n{'='*70}")
            print(f"  CONCLUÍDO!")
            print(f"  Total de registros extraídos: {len(df):,}")
            print(f"  Arquivo gerado: {output}")
            print(f"(Os decimais aparecerão com vírgula no seu Excel BR)")
            print(f"{'='*70}\n")
        else:
            print("\nNenhum registro de impedância foi extraído de toda a base.")
