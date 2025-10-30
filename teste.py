# pip install pandas openpyxl
from __future__ import annotations
import json
from pathlib import Path
from typing import Any, Dict, List, Union

import pandas as pd

# ---------- helpers ----------
def limpar_json_invalido(json_str: str) -> str:
    """
    Remove campos com valores vazios que quebram o JSON.
    Ex: "emprestimosLegados":, vira "emprestimosLegados":null
    """
    import re
    # Substitui ":," por ":null,"
    json_str = re.sub(r':\s*,', ':null,', json_str)
    # Substitui ":}" por ":null}"
    json_str = re.sub(r':\s*}', ':null}', json_str)
    # Substitui ":]" por ":null]"
    json_str = re.sub(r':\s*]', ':null]', json_str)
    return json_str

def extrair_cbo_cnae_de_json(json_str: str) -> List[Dict[str, str]]:
    """
    Extrai todos os CBOs e CNAEs de um JSON (string ou array).
    Retorna lista de dicts com 'CBO' e 'CNAE' formatados.
    """
    resultados = []
    
    try:
        # Limpa JSON inválido antes de fazer parse
        json_limpo = limpar_json_invalido(json_str.strip())
        
        # Remove espaços e faz parse
        json_obj = json.loads(json_limpo)
        
        # Se for lista, processa cada item
        if isinstance(json_obj, list):
            for item in json_obj:
                if isinstance(item, dict):
                    resultado = {}
                    
                    # Extrai CBO
                    if 'cbo' in item and isinstance(item['cbo'], dict):
                        cbo = item['cbo']
                        codigo = cbo.get('codigo', '')
                        descricao = cbo.get('descricao', '')
                        resultado['CBO'] = f"{codigo} - {descricao}" if codigo else ''
                    
                    # Extrai CNAE
                    if 'cnae' in item and isinstance(item['cnae'], dict):
                        cnae = item['cnae']
                        codigo = cnae.get('codigo', '')
                        descricao = cnae.get('descricao', '')
                        resultado['CNAE'] = f"{codigo} - {descricao}" if codigo else ''
                    
                    # Extrai outros campos úteis (nome, cpf, etc)
                    if 'nome' in item:
                        resultado['Nome'] = item['nome']
                    if 'cpf' in item:
                        resultado['CPF'] = str(item['cpf'])
                    if 'nomeEmpregador' in item:
                        resultado['Empregador'] = item['nomeEmpregador']
                    
                    resultados.append(resultado)
        
        # Se for objeto único
        elif isinstance(json_obj, dict):
            resultado = {}
            
            # Extrai CBO
            if 'cbo' in json_obj and isinstance(json_obj['cbo'], dict):
                cbo = json_obj['cbo']
                codigo = cbo.get('codigo', '')
                descricao = cbo.get('descricao', '')
                resultado['CBO'] = f"{codigo} - {descricao}" if codigo else ''
            
            # Extrai CNAE
            if 'cnae' in json_obj and isinstance(json_obj['cnae'], dict):
                cnae = json_obj['cnae']
                codigo = cnae.get('codigo', '')
                descricao = cnae.get('descricao', '')
                resultado['CNAE'] = f"{codigo} - {descricao}" if codigo else ''
            
            # Extrai outros campos úteis
            if 'nome' in json_obj:
                resultado['Nome'] = json_obj['nome']
            if 'cpf' in json_obj:
                resultado['CPF'] = str(json_obj['cpf'])
            if 'nomeEmpregador' in json_obj:
                resultado['Empregador'] = json_obj['nomeEmpregador']
            
            resultados.append(resultado)
    
    except Exception as e:
        # Se falhar, retorna vazio
        pass
    
    return resultados

def excel_extrair_cbo_cnae(
    path_entrada: Union[str, Path],
    path_saida: Union[str, Path],
    sheet: Union[int, str] = 0,
):
    """
    Lê um Excel que contém células com JSON e extrai apenas CBO e CNAE formatados.
    Cada registro JSON vira UMA linha no Excel de saída.
    Processa apenas a coluna 'Margens Prev' que contém os dados completos.
    """
    # Lê o Excel original
    df = pd.read_excel(path_entrada, sheet_name=sheet)
    
    todas_linhas = []
    
    for idx, row in df.iterrows():
        # Procura especificamente pela coluna "Margens Prev"
        if 'Margens Prev' in row and pd.notna(row['Margens Prev']):
            valor = row['Margens Prev']
            
            if isinstance(valor, str):
                valor_strip = valor.strip()
                
                # Se parece JSON
                if (valor_strip.startswith('{') or valor_strip.startswith('[')):
                    # Extrai CBOs e CNAEs
                    registros = extrair_cbo_cnae_de_json(valor_strip)
                    
                    # Adiciona cada registro como uma linha
                    for reg in registros:
                        linha_nova = {}
                        
                        # Adiciona outras colunas da linha original (exceto a coluna de JSON)
                        for col_name, col_val in row.items():
                            if col_name != 'Margens Prev' and pd.notna(col_val):
                                # Não adiciona se for JSON também
                                if not (isinstance(col_val, str) and 
                                       (col_val.strip().startswith('{') or col_val.strip().startswith('['))):
                                    linha_nova[str(col_name)] = col_val
                        
                        # Adiciona os dados extraídos
                        linha_nova.update(reg)
                        todas_linhas.append(linha_nova)
    
    # Se não extraiu nada, retorna o Excel original
    if not todas_linhas:
        df.to_excel(path_saida, index=False, engine='openpyxl')
        return len(df), len(df.columns)
    
    # Cria DataFrame com todas as linhas extraídas
    df_saida = pd.DataFrame(todas_linhas)
    
    # Remove duplicatas baseado no nome (mantém a primeira ocorrência)
    if 'Nome' in df_saida.columns:
        df_saida = df_saida.drop_duplicates(subset=['Nome'], keep='first')
    
    # Reordena colunas para ter CBO e CNAE primeiro
    colunas_desejadas = []
    if 'Nome' in df_saida.columns:
        colunas_desejadas.append('Nome')
    if 'CPF' in df_saida.columns:
        colunas_desejadas.append('CPF')
    if 'CBO' in df_saida.columns:
        colunas_desejadas.append('CBO')
    if 'CNAE' in df_saida.columns:
        colunas_desejadas.append('CNAE')
    if 'Empregador' in df_saida.columns:
        colunas_desejadas.append('Empregador')
    
    # Adiciona colunas restantes
    for col in df_saida.columns:
        if col not in colunas_desejadas:
            colunas_desejadas.append(col)
    
    df_saida = df_saida[colunas_desejadas]
    
    # Salva em Excel
    df_saida.to_excel(path_saida, index=False, engine='openpyxl')
    
    return len(todas_linhas), len(df_saida.columns)

# ---------- uso direto ----------
if __name__ == "__main__":
    import os
    
    # Detecta o diretório do script atual
    script_dir = Path(__file__).parent
    pasta_leiame = script_dir / "leiame"
    
    # Verifica se a pasta existe
    if not pasta_leiame.exists():
        print(f"❌ Pasta 'leiame' não encontrada em: {pasta_leiame}")
        print("Criando pasta 'leiame'...")
        pasta_leiame.mkdir(exist_ok=True)
        print(f"✅ Pasta criada! Coloque seus arquivos Excel (.xlsx, .xls) dentro de: {pasta_leiame}")
        exit(0)
    
    # Busca todos os arquivos Excel na pasta
    arquivos_excel = list(pasta_leiame.glob("*.xlsx")) + list(pasta_leiame.glob("*.xls"))
    
    if not arquivos_excel:
        print(f"❌ Nenhum arquivo Excel encontrado na pasta: {pasta_leiame}")
        print("Coloque arquivos .xlsx ou .xls na pasta 'leiame' e execute novamente.")
        exit(0)
    
    print(f"📂 Encontrados {len(arquivos_excel)} arquivo(s) Excel na pasta 'leiame':\n")
    
    # Processa cada arquivo encontrado
    for arquivo in arquivos_excel:
        print(f"📄 Processando: {arquivo.name}")
        try:
            # Nome do arquivo de saída baseado no arquivo de entrada + timestamp
            import time
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            nome_saida = f"{arquivo.stem}_CBO_CNAE_{timestamp}.xlsx"
            caminho_saida = script_dir / nome_saida
            
            num_linhas, num_colunas = excel_extrair_cbo_cnae(arquivo, caminho_saida, sheet=0)
            
            print(f"   ✅ {num_linhas} linhas processadas")
            print(f"   📊 Extraído CBO e CNAE em '{nome_saida}'\n")
            
        except Exception as e:
            import traceback
            print(f"   ❌ Erro ao processar {arquivo.name}: {e}")
            traceback.print_exc()
            print()
    
    print("🎉 Processamento concluído!")
