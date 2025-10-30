"""processar_propostas.py

Script para processar arquivos Excel com propostas (múltiplas variações de colunas)
e extrair/normalizar os campos aninhados (por exemplo, 'Margens Prev' com JSON).

Coloque os arquivos Excel na pasta `leiame/` ao lado deste repositório e execute:
    python processo_propostas.py

Saída: gera um arquivo Excel com os campos normalizados e colunas extras para CBO/CNAE.
"""
from __future__ import annotations
import json
import os
import re
import shutil
import subprocess
import sys
import time
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union

import pandas as pd


def encontrar_base_dir():
    """
    Procura o diretório base onde o programa está rodando.
    Funciona em qualquer pasta, qualquer PC.
    """
    # Modo executável (PyInstaller) - usa pasta onde o .exe está
    if getattr(sys, 'frozen', False):
        exe_dir = Path(sys.executable).parent
        print(f"✅ Usando pasta do executável: {exe_dir}")
        return exe_dir
    else:
        # Script Python (desenvolvimento) - usa pasta do script
        script_dir = Path(__file__).parent.resolve()
        print(f"✅ Usando pasta do script: {script_dir}")
        return script_dir


def limpar_tela():
    """Limpa a tela do terminal"""
    os.system('cls' if os.name == 'nt' else 'clear')


def exibir_banner():
    """Exibe o banner do programa"""
    print("=" * 70)
    print("║" + " " * 68 + "║")
    print("║" + " " * 15 + "PROCESSADOR DE PROPOSTAS WIIPO" + " " * 23 + "║")
    print("║" + " " * 68 + "║")
    print("=" * 70)
    print()


def exibir_menu():
    """Exibe o menu principal"""
    print("\n┌" + "─" * 68 + "┐")
    print("│  📋 MENU PRINCIPAL" + " " * 49 + "│")
    print("├" + "─" * 68 + "┤")
    print("│                                                                    │")
    print("│  [1] Processar arquivo CSV/Excel                                  │")
    print("│  [2] Ver arquivos processados (pasta 'lidos')                     │")
    print("│  [3] Limpar pasta 'lidos'                                         │")
    print("│  [4] Ver arquivo formatado existente                              │")
    print("│  [0] Sair                                                         │")
    print("│                                                                    │")
    print("└" + "─" * 68 + "┘")
    print()


def aguardar_enter():
    """Aguarda o usuário pressionar Enter"""
    input("\n⏎ Pressione ENTER para continuar...")


def fechar_excel():
    """Verifica se o Excel está aberto e pergunta se quer fechar"""
    try:
        print("🔄 Verificando se o Excel está aberto...")
        result = subprocess.run(['tasklist', '/FI', 'IMAGENAME eq EXCEL.EXE'], 
                              capture_output=True, text=True, shell=True)
        
        if 'EXCEL.EXE' in result.stdout:
            print("\n⚠️  ATENÇÃO: Excel está aberto!")
            print("⚠️  Para evitar erros, é recomendado fechar o Excel antes de continuar.")
            print("⚠️  Todas as planilhas abertas serão fechadas (SALVE SEU TRABALHO!).")
            print("\n💡 OPÇÕES:")
            print("   1. Fechar Excel automaticamente e continuar")
            print("   2. Criar arquivo temporário (abre em nova aba do Excel)")
            print("   3. Cancelar e fechar manualmente")
            
            resposta = input("\n🤔 Escolha uma opção (1/2/3): ").strip()
            
            if resposta == '1':
                print("📊 Fechando Excel...")
                subprocess.run(['taskkill', '/F', '/IM', 'EXCEL.EXE'], 
                             capture_output=True, shell=True)
                time.sleep(2)  # Aguarda o Excel fechar completamente
                print("✅ Excel fechado com sucesso!")
                return 'fechado'
            elif resposta == '2':
                print("📋 Modo arquivo temporário selecionado!")
                return 'temporario'
            else:
                print("❌ Operação cancelada pelo usuário.")
                return 'cancelado'
        else:
            print("✅ Excel não está aberto.")
            return 'livre'
    except Exception as e:
        print(f"⚠️ Erro ao verificar Excel: {e}")
        return 'livre'


def criar_arquivo_temporario(script_dir, arquivo_entrada, pasta_lidos):
    """Cria arquivo temporário para visualização sem afetar arquivos abertos"""
    try:
        from datetime import datetime
        import glob
        
        # Limpa arquivos temporários antigos primeiro
        print(f"\n🧹 Limpando arquivos temporários antigos...")
        temp_files = glob.glob(str(script_dir / "relatorio_propostas_TEMP_*.xlsx"))
        if temp_files:
            for temp_file in temp_files:
                try:
                    os.remove(temp_file)
                    print(f"   🗑️ Removido: {os.path.basename(temp_file)}")
                except Exception as e:
                    print(f"   ⚠️ Não foi possível remover {os.path.basename(temp_file)}: {e}")
            print(f"✅ Limpeza concluída!")
        else:
            print(f"✅ Nenhum arquivo temporário antigo encontrado.")
        
        # Cria arquivo temporário com timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        temp_path = script_dir / f"relatorio_propostas_TEMP_{timestamp}.xlsx"
        
        print(f"\n📊 Criando arquivo temporário: {temp_path.name}")
        print("🔄 Processando...")
        
        # Processa o arquivo
        processar_arquivo(arquivo_entrada, temp_path)
        
        print(f"\n✅ Arquivo temporário criado com sucesso!")
        print(f"📁 Local: {temp_path}")
        print(f"\n💡 IMPORTANTE: Este é um arquivo temporário para visualização.")
        print(f"💡 Ele será automaticamente removido na próxima execução.")
        print(f"💡 Feche o Excel principal e execute a opção [1] novamente para atualizar o arquivo definitivo.")
        
        # Move arquivo original para 'lidos' também no modo temporário
        if not pasta_lidos.exists():
            pasta_lidos.mkdir(exist_ok=True)
        
        destino = pasta_lidos / arquivo_entrada.name
        if destino.exists():
            destino.unlink()
        shutil.move(str(arquivo_entrada), str(destino))
        print(f"📦 Arquivo original movido para: lidos/{arquivo_entrada.name}")
        
        # Pergunta se quer abrir o temporário
        print()
        resposta = input("🤔 Deseja abrir o arquivo temporário agora? (S/N): ").strip().upper()
        
        if resposta in ['S', 'SIM', 'Y', 'YES']:
            os.startfile(str(temp_path))
            print("✅ Arquivo temporário aberto em nova aba do Excel!")
            time.sleep(1)
        else:
            print("⏭️ Arquivo temporário salvo, mas não foi aberto.")
        
        return True
        
    except Exception as e:
        print(f"❌ Erro ao criar arquivo temporário: {e}")
        import traceback
        traceback.print_exc()
        return False


def listar_arquivos_lidos(pasta_lidos):
    """Lista arquivos na pasta lidos"""
    limpar_tela()
    exibir_banner()
    print("📁 ARQUIVOS PROCESSADOS\n")
    
    if not pasta_lidos.exists():
        print("❌ Pasta 'lidos' não existe ainda.")
        aguardar_enter()
        return
    
    arquivos = list(pasta_lidos.glob('*'))
    if not arquivos:
        print("📂 Pasta 'lidos' está vazia.")
    else:
        print(f"Total: {len(arquivos)} arquivo(s)\n")
        for idx, arq in enumerate(arquivos, 1):
            tamanho = arq.stat().st_size / 1024  # KB
            print(f"  {idx}. {arq.name} ({tamanho:.1f} KB)")
    
    aguardar_enter()


def limpar_pasta_lidos(pasta_lidos):
    """Limpa todos os arquivos da pasta lidos"""
    limpar_tela()
    exibir_banner()
    print("🗑️  LIMPAR PASTA 'LIDOS'\n")
    
    if not pasta_lidos.exists():
        print("❌ Pasta 'lidos' não existe ainda.")
        aguardar_enter()
        return
    
    arquivos = list(pasta_lidos.glob('*'))
    if not arquivos:
        print("📂 Pasta 'lidos' já está vazia.")
        aguardar_enter()
        return
    
    print(f"⚠️  Isso irá deletar {len(arquivos)} arquivo(s):")
    for arq in arquivos:
        print(f"  - {arq.name}")
    
    print()
    confirmacao = input("Confirma a exclusão? (S/N): ").strip().upper()
    
    if confirmacao == 'S':
        import shutil
        for arq in arquivos:
            arq.unlink()
        print("\n✅ Arquivos deletados com sucesso!")
    else:
        print("\n❌ Operação cancelada.")
    
    aguardar_enter()


def abrir_arquivo_excel(caminho_arquivo):
    """Abre o arquivo Excel em uma nova instância para não interferir com arquivos abertos"""
    try:
        # Usa os.startfile() que é mais confiável no Windows
        os.startfile(str(caminho_arquivo))
        return True
    except Exception as e:
        print(f"❌ Erro ao abrir arquivo: {e}")
        return False


def ver_arquivo_formatado(script_dir):
    """Visualiza o arquivo formatado existente"""
    limpar_tela()
    exibir_banner()
    print("📄 VER ARQUIVO FORMATADO\n")
    
    saida = script_dir / "relatorio_propostas_formatado.xlsx"
    
    if not saida.exists():
        print("❌ Arquivo formatado não encontrado!")
        print(f"💡 Execute a opção [1] para processar um arquivo primeiro.")
        aguardar_enter()
        return
    
    tamanho = saida.stat().st_size / 1024  # KB
    print(f"📊 Arquivo: {saida.name}")
    print(f"📁 Tamanho: {tamanho:.1f} KB")
    print(f"📍 Local: {saida}")
    print()
    
    abrir = input("Deseja abrir o arquivo? (S/N): ").strip().upper()
    if abrir == 'S':
        if abrir_arquivo_excel(saida):
            print("✅ Arquivo aberto!")
        time.sleep(1)
    else:
        print("❌ Operação cancelada.")
        aguardar_enter()


def processar_arquivo_menu(script_dir, pasta_leiame, pasta_lidos):
    """Processa o arquivo com interface visual"""
    limpar_tela()
    exibir_banner()
    print("🔄 PROCESSAR ARQUIVO\n")
    
    # Verifica se Excel está aberto
    status_excel = fechar_excel()
    if status_excel == 'cancelado':
        print("\n❌ Operação cancelada.")
        aguardar_enter()
        return
    
    print()  # Espaço após verificação do Excel
    
    # Se escolheu modo temporário, usa função específica
    modo_temporario = (status_excel == 'temporario')
    
    arquivos = list(pasta_leiame.glob('*.xlsx')) + list(pasta_leiame.glob('*.xls')) + list(pasta_leiame.glob('*.csv'))
    
    if not arquivos:
        print(f"❌ Nenhum arquivo encontrado em: {pasta_leiame}")
        print("\n💡 Coloque um arquivo CSV ou Excel na pasta 'leiame' e tente novamente.")
        aguardar_enter()
        return
    
    if len(arquivos) > 1:
        print(f'❌ ERRO: Encontrados {len(arquivos)} arquivos!')
        print('⚠️  Por favor, deixe apenas 1 arquivo por vez.\n')
        print('Arquivos encontrados:')
        for arq in arquivos:
            print(f'  - {arq.name}')
        aguardar_enter()
        return
    
    arquivo = arquivos[0]
    print(f"📄 Arquivo encontrado: {arquivo.name}")
    print(f"📊 Tamanho: {arquivo.stat().st_size / 1024:.1f} KB")
    print()
    
    confirmacao = input("Deseja processar este arquivo? (S/N): ").strip().upper()
    
    if confirmacao != 'S':
        print("\n❌ Operação cancelada.")
        aguardar_enter()
        return
    
    print("\n" + "─" * 70)
    print("🔄 PROCESSANDO...")
    print("─" * 70 + "\n")
    
    # Se modo temporário, usa função específica
    if modo_temporario:
        criar_arquivo_temporario(script_dir, arquivo, pasta_lidos)
        aguardar_enter()
        return
    
    # Modo normal - processa arquivo definitivo
    # Nome fixo para o arquivo de saída
    saida = script_dir / "relatorio_propostas_formatado.xlsx"
    pasta_backup = script_dir / 'backup'
    
    # Cria pasta backup se não existir
    if not pasta_backup.exists():
        pasta_backup.mkdir(exist_ok=True)
        print("📁 Pasta de backup criada")
    
    # Se já existe um arquivo formatado, move para backup (substituindo o backup anterior)
    if saida.exists():
        backup_path = pasta_backup / "relatorio_propostas_formatado.xlsx"
        if backup_path.exists():
            backup_path.unlink()  # Remove backup antigo
        shutil.move(str(saida), str(backup_path))
        print("📦 Backup do arquivo anterior criado")
    
    try:
        processar_arquivo(arquivo, saida)
        
        print("\n" + "=" * 70)
        print("✅ Arquivo processado com sucesso!")
        print(f"📁 Salvo como: {saida.name}")
        print(f"📍 Local completo: {saida}")
        print("=" * 70)
        
        # Cria pasta 'lidos' se não existir
        if not pasta_lidos.exists():
            pasta_lidos.mkdir(exist_ok=True)
        
        # Move arquivo para pasta 'lidos'
        destino = pasta_lidos / arquivo.name
        if destino.exists():
            destino.unlink()  # Remove se já existir
        shutil.move(str(arquivo), str(destino))
        print(f"📦 Arquivo original movido para: lidos/{arquivo.name}")
        
        # Pergunta se quer abrir o arquivo
        print()
        abrir = input("Deseja abrir o arquivo formatado? (S/N): ").strip().upper()
        if abrir == 'S':
            if abrir_arquivo_excel(saida):
                print("✅ Arquivo aberto!")
                time.sleep(1)
            
    except Exception as e:
        print(f"\n❌ ERRO ao processar: {e}")
        import traceback
        traceback.print_exc()
    
    aguardar_enter()


def limpar_json_invalido(json_str: str) -> str:
    """Tenta consertar pedaços comuns que tornam o JSON inválido.
    Exemplos: 'emprestimosLegados':,  -> 'emprestimosLegados':null,
    """
    s = json_str
    # substitui :,, :} and :] patterns
    s = re.sub(r':\s*,', ':null,', s)
    s = re.sub(r':\s*}', ':null}', s)
    s = re.sub(r':\s*]', ':null]', s)
    # remove caracteres de controle estranhos (se houver)
    s = s.replace('\x00', '')
    return s


def safe_load_json(s: str) -> Optional[Union[Dict[str, Any], List[Any]]]:
    try:
        return json.loads(s)
    except Exception:
        try:
            s2 = limpar_json_invalido(s)
            return json.loads(s2)
        except Exception:
            return None


def extrair_registros_de_margens(margens_val: Any) -> List[Dict[str, Any]]:
    """Dado o valor da coluna 'Margens Prev' (str com JSON, lista ou dict), retorna lista de registros tratados."""
    if pd.isna(margens_val) or margens_val is None:
        return []
    # se já é lista/dict
    if isinstance(margens_val, (list, dict)):
        parsed = margens_val
    elif isinstance(margens_val, str):
        parsed = safe_load_json(margens_val.strip())
    else:
        parsed = None

    if parsed is None:
        return []

    if isinstance(parsed, dict):
        return [parsed]
    if isinstance(parsed, list):
        # mantemos apenas objetos dict
        return [p for p in parsed if isinstance(p, dict)]
    return []


def mapear_linha(row: pd.Series, margem_item: Dict[str, Any]) -> Dict[str, Any]:
    """Cria um dict plano combinando colunas top-level e campos extraídos do item de margem."""
    out: Dict[str, Any] = {}
    # Copia colunas top-level relevantes (se existirem)
    for col in row.index:
        if col == 'Margens Prev':
            continue
        val = row[col]
        if pd.notna(val):
            out[str(col)] = val

    # Extrai campos do item de margem (cbo, cnae, nome, cpf, empregador, valores)
    cbo = margem_item.get('cbo') if isinstance(margem_item.get('cbo'), dict) else None
    cnae = margem_item.get('cnae') if isinstance(margem_item.get('cnae'), dict) else None

    if cbo:
        out['MargensPrev_CBO_Codigo'] = cbo.get('codigo')
        out['MargensPrev_CBO_Descricao'] = cbo.get('descricao')
    if cnae:
        out['MargensPrev_CNAE_Codigo'] = cnae.get('codigo')
        out['MargensPrev_CNAE_Descricao'] = cnae.get('descricao')

    # Outros campos do registro interno
    for k in ('nome', 'cpf', 'nomeEmpregador', 'valorMargemDisponivel', 'valorBaseMargem'):
        if k in margem_item:
            out[f'MargensPrev_{k}'] = margem_item.get(k)

    return out


def extrair_motivo_inelegibilidade(val: Any) -> Tuple[Optional[int], Optional[str]]:
    if pd.isna(val) or val is None:
        return None, None
    if isinstance(val, str):
        parsed = safe_load_json(val.strip())
        if isinstance(parsed, dict):
            return parsed.get('codigo'), parsed.get('descricao')
        return None, None
    if isinstance(val, dict):
        return val.get('codigo'), val.get('descricao')
    return None, None


def processar_arquivo(path_entrada: Union[str, Path], path_saida: Union[str, Path]):
    # Detecta se é CSV ou Excel
    if str(path_entrada).endswith('.csv'):
        df = pd.read_csv(path_entrada, encoding='utf-8', on_bad_lines='skip')
    else:
        df = pd.read_excel(path_entrada, sheet_name=0)

    linhas_saida: List[Dict[str, Any]] = []

    for _, row in df.iterrows():
        # pega especificamente a coluna 'Margens Prev' (tolerância a variações)
        col_margens = None
        for candidate in ['Margens Prev', 'MargensPrev', 'Margens_Prev']:
            if candidate in row.index:
                col_margens = candidate
                break

        margem_val = row[col_margens] if col_margens else None

        registros = extrair_registros_de_margens(margem_val)
        # extrai motivo inelegibilidade (se houver)
        motivo_codigo, motivo_desc = None, None
        for candidate in ['Motivo Inelegibilidade', 'MotivoInelegibilidade', 'Motivo']:
            if candidate in row.index:
                motivo_codigo, motivo_desc = extrair_motivo_inelegibilidade(row[candidate])
                break

        if registros:
            for reg in registros:
                mapped = mapear_linha(row, reg)
                if motivo_codigo is not None or motivo_desc is not None:
                    mapped['MotivoInelegibilidade_Codigo'] = motivo_codigo
                    mapped['MotivoInelegibilidade_Descricao'] = motivo_desc
                linhas_saida.append(mapped)
        else:
            # nenhuma margem interna: ainda assim mantém a linha (com colunas top-level)
            out = {}
            for col in row.index:
                val = row[col]
                if pd.notna(val) and not (isinstance(val, str) and (val.strip().startswith('{') or val.strip().startswith('['))):
                    out[str(col)] = val
            if motivo_codigo is not None or motivo_desc is not None:
                out['MotivoInelegibilidade_Codigo'] = motivo_codigo
                out['MotivoInelegibilidade_Descricao'] = motivo_desc
            if out:
                linhas_saida.append(out)

    if not linhas_saida:
        print('Nenhum registro extraído — salvando cópia do original.')
        df.to_excel(path_saida, index=False, engine='openpyxl')
        return

    df_out = pd.DataFrame(linhas_saida)

    # Reordena um conjunto de colunas úteis se existirem
    prioridade = ['Parceiro', 'Data de abertura da proposta', 'NÃºmero da Proposta', 'Status da Proposta',
                  'Nome do UsuÃ¡rio', 'CPF do UsuÃ¡rio', 'MargensPrev_CBO_Codigo', 'MargensPrev_CBO_Descricao',
                  'MargensPrev_CNAE_Codigo', 'MargensPrev_CNAE_Descricao', 'MotivoInelegibilidade_Codigo',
                  'MotivoInelegibilidade_Descricao']
    cols = [c for c in prioridade if c in df_out.columns]
    cols += [c for c in df_out.columns if c not in cols]
    df_out = df_out[cols]

    df_out.to_excel(path_saida, index=False, engine='openpyxl')


def main():
    """Função principal com menu interativo"""
    script_dir = encontrar_base_dir()
    pasta_leiame = script_dir / 'leiame'
    pasta_lidos = script_dir / 'lidos'
    
    # Cria pastas se não existirem
    if not pasta_leiame.exists():
        pasta_leiame.mkdir(exist_ok=True)
    
    while True:
        limpar_tela()
        exibir_banner()
        
        # Mostra status das pastas
        print("📂 STATUS DAS PASTAS:")
        print(f"  • leiame: {len(list(pasta_leiame.glob('*')))} arquivo(s)")
        if pasta_lidos.exists():
            print(f"  • lidos: {len(list(pasta_lidos.glob('*')))} arquivo(s)")
        else:
            print(f"  • lidos: (pasta não criada ainda)")
        
        # Verifica se existe arquivo formatado
        saida = script_dir / "relatorio_propostas_formatado.xlsx"
        if saida.exists():
            tamanho = saida.stat().st_size / 1024
            print(f"  • arquivo formatado: relatorio_propostas_formatado.xlsx ({tamanho:.1f} KB)")
        
        exibir_menu()
        
        opcao = input("Escolha uma opção: ").strip()
        
        if opcao == '1':
            processar_arquivo_menu(script_dir, pasta_leiame, pasta_lidos)
        elif opcao == '2':
            listar_arquivos_lidos(pasta_lidos)
        elif opcao == '3':
            limpar_pasta_lidos(pasta_lidos)
        elif opcao == '4':
            ver_arquivo_formatado(script_dir)
        elif opcao == '0':
            limpar_tela()
            exibir_banner()
            print("👋 Encerrando programa...")
            print("\n✅ Até logo!")
            time.sleep(1)
            break
        else:
            print("\n❌ Opção inválida! Escolha entre 0-4.")
            time.sleep(1)
    
    # Pausa final apenas se for executável
    if getattr(sys, 'frozen', False):
        input("\nPressione ENTER para fechar...")


if __name__ == '__main__':
    main()

