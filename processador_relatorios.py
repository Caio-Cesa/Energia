import pandas as pd
import os

def extrair(linha, inicio, fim):
    """Extrai substring segura."""
    try:
        return linha[inicio:fim].strip()
    except:
        return ""

def processar_arquivo(caminho):
    """Processa um arquivo focando na seção RELATORIO COMPLETO (PAG 4+)."""
    registros = []
    sessao = "-"
    barra = None
    
    # Estado do parser
    em_dados = False
    skip_header = False
    x_count = 0
    
    with open(caminho, 'r', encoding='windows-1252', errors='replace') as f:
        for linha in f:
            raw = linha.rstrip('\r\n')
            
            # --- CONTROLE DE SEÇÃO ---
            if "RELATORIO COMPLETO DO SISTEMA" in raw:
                parts = raw.split("*")
                sessao = " ".join(p.strip() for p in parts[1:]) if len(parts) > 1 else "-"
                em_dados = False
                skip_header = True
                x_count = 0
                continue
            
            # Skip de cabeçalho: espera 2 linhas X para entrar nos dados
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
            # Quebra de página
            if 'CEPEL' in raw or 'PAG.' in raw:
                em_dados = False
                continue
            
            if not raw.strip():
                continue
            
            # Linha X no meio dos dados (segurança)
            if raw.lstrip().startswith('X') and '---' in raw:
                continue
            
            # Separador de pontos = fim do bloco da barra
            if '........' in raw:
                if barra and not barra.get('_tem_fluxo'):
                    reg = {k: v for k, v in barra.items() if not k.startswith('_')}
                    registros.append(reg)
                barra = None
                continue
            
            # --- CLASSIFICAÇÃO DA LINHA ---
            campo_num = raw[0:7].strip()
            campo_esq = raw[0:23].strip()
            
            # LINHA A: Dados numéricos da barra (NUM no início)
            if campo_num.isdigit():
                # Salvar barra anterior sem fluxos
                if barra and not barra.get('_tem_fluxo'):
                    reg = {k: v for k, v in barra.items() if not k.startswith('_')}
                    registros.append(reg)
                
                barra = {
                    'NUM': campo_num,
                    'KV': extrair(raw, 7, 12),
                    'TIPO': extrair(raw, 12, 15),
                    'TENSAO': extrair(raw, 15, 23),
                    'GERACAO_MW': extrair(raw, 23, 31),
                    'INJ_EQV_MW': extrair(raw, 31, 39),
                    'CARGA_MW': extrair(raw, 39, 47),
                    'ELO_CC_MW': extrair(raw, 47, 59),
                    'SHUNT_Mvar': extrair(raw, 59, 67),
                    'MOTOR_MW': extrair(raw, 67, 75),
                    'SESSAO': sessao,
                    'ARQUIVO': os.path.basename(caminho),
                    '_tem_fluxo': False,
                    '_tem_nome': False
                }
            
            # LINHA B: Nome da barra (segue a Linha A)
            elif barra and not barra.get('_tem_nome') and any(c.isalpha() for c in raw[0:16]):
                barra['NOME'] = extrair(raw, 0, 16)
                barra['ANG'] = extrair(raw, 16, 23)
                barra['GERACAO_Mvar'] = extrair(raw, 23, 31)
                barra['INJ_EQV_Mvar'] = extrair(raw, 31, 39)
                barra['CARGA_Mvar'] = extrair(raw, 39, 47)
                barra['ELO_CC_Mvar'] = extrair(raw, 47, 59)
                barra['EQUIV'] = extrair(raw, 59, 67)
                barra['MOTOR_Mvar'] = extrair(raw, 67, 75)
                barra['_tem_nome'] = True
            
            # LINHA C: Fluxo de circuito (esquerda vazia, dados à direita)
            elif barra and barra.get('_tem_nome') and not campo_esq:
                reg = {k: v for k, v in barra.items() if not k.startswith('_')}
                reg['MVA_NOM'] = extrair(raw, 23, 31)
                reg['MVA_EMR'] = extrair(raw, 31, 39)
                reg['MVA_EQP'] = extrair(raw, 39, 47)
                reg['FLUXO_%'] = extrair(raw, 47, 59)
                reg['SHUNT_L'] = extrair(raw, 59, 67)
                reg['PARA_NUM'] = extrair(raw, 75, 81)
                reg['PARA_NOME'] = extrair(raw, 81, 94)
                reg['NC'] = extrair(raw, 94, 97)
                reg['FLUXO_MW'] = extrair(raw, 97, 105)
                reg['FLUXO_Mvar'] = extrair(raw, 105, 113)
                reg['MVA_Vd'] = extrair(raw, 113, 121)
                reg['TAP'] = extrair(raw, 121, 128)
                reg['DEFAS'] = extrair(raw, 128, 134)
                reg['TIE'] = extrair(raw, 134, 140)
                registros.append(reg)
                barra['_tem_fluxo'] = True
    
    # Última barra sem fluxos
    if barra and not barra.get('_tem_fluxo'):
        reg = {k: v for k, v in barra.items() if not k.startswith('_')}
        registros.append(reg)
    
    return registros


if __name__ == "__main__":
    pasta = "Dados_Entrada"
    arquivo = next((os.path.join(pasta, f) for f in sorted(os.listdir(pasta)) if f.endswith(".txt")), None)
    
    if arquivo:
        print(f"\n{'='*70}")
        print(f"  PROCESSADOR v5.0 - Foco em PAG 4+ (RELATORIO COMPLETO)")
        print(f"  Modo: Validação (1 arquivo)")
        print(f"  Arquivo: {os.path.basename(arquivo)}")
        print(f"{'='*70}\n")
        
        regs = processar_arquivo(arquivo)
        df = pd.DataFrame(regs)
        
        if not df.empty:
            # --- CONVERSÃO NUMÉRICA ---
            # Limpar o "%" de FLUXO_% para virar número
            if 'FLUXO_%' in df.columns:
                df['FLUXO_%'] = df['FLUXO_%'].str.replace('%', '', regex=False)
            
            # Colunas que devem ser numéricas
            cols_numericas = [
                'NUM', 'KV', 'TIPO', 'TENSAO', 'GERACAO_MW', 'INJ_EQV_MW',
                'CARGA_MW', 'ELO_CC_MW', 'SHUNT_Mvar', 'MOTOR_MW',
                'ANG', 'GERACAO_Mvar', 'INJ_EQV_Mvar', 'CARGA_Mvar',
                'ELO_CC_Mvar', 'EQUIV', 'MOTOR_Mvar',
                'MVA_NOM', 'MVA_EMR', 'MVA_EQP', 'FLUXO_%', 'PARA_NUM', 'NC',
                'FLUXO_MW', 'FLUXO_Mvar', 'MVA_Vd'
            ]
            for col in cols_numericas:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors='coerce')
            
            # --- PREVIEW ---
            cols_preview = ['NUM', 'NOME', 'KV', 'TENSAO', 'PARA_NUM', 'PARA_NOME', 'FLUXO_%', 'SESSAO']
            cols_ok = [c for c in cols_preview if c in df.columns]
            
            print(">>> PREVIEW DA TABELA (Top 15):\n")
            print(df[cols_ok].head(15).to_string(index=False))
            print(f"\n>>> Total de registros: {len(df)}")
            print(f">>> Colunas: {list(df.columns)}")
            
            df.to_excel("Resultado_v5_Validacao.xlsx", index=False)
            print(f"\nSalvo em: Resultado_v5_Validacao.xlsx")
            print("(Os decimais aparecerão com vírgula no seu Excel BR)")
        else:
            print("Nenhum registro extraído.")
    else:
        print("Nenhum arquivo .txt encontrado.")
