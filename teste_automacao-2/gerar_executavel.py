"""
Script para gerar executável do Processador de Propostas
Usa PyInstaller para criar um .exe standalone que funciona em qualquer PC
"""
import subprocess
import sys
import os
from pathlib import Path

def main():
    print("=" * 70)
    print("🔨 GERADOR DE EXECUTÁVEL - PROCESSADOR DE PROPOSTAS WIIPO")
    print("=" * 70)
    print()
    
    # Caminho do Python
    python_path = "C:/Users/KawanWagnnerGonçalve/AppData/Local/Python/pythoncore-3.14-64/python.exe"
    
    # Verifica se Python existe
    if not os.path.exists(python_path):
        print(f"❌ Python não encontrado em: {python_path}")
        print("💡 Ajuste o caminho do Python no script")
        input("\nPressione ENTER para sair...")
        return
    
    print(f"✅ Python encontrado: {python_path}")
    print()
    
    # Caminho do script
    script_path = Path(__file__).parent / "processar_propostas.py"
    
    if not script_path.exists():
        print(f"❌ Script não encontrado: {script_path}")
        input("\nPressione ENTER para sair...")
        return
    
    print(f"✅ Script encontrado: {script_path.name}")
    print()
    
    # Confirmação
    print("🔄 Isso irá:")
    print("   1. Limpar builds anteriores")
    print("   2. Compilar o script Python")
    print("   3. Gerar ProcessadorPropostas.exe na pasta 'dist'")
    print()
    
    resposta = input("Deseja continuar? (S/N): ").strip().upper()
    
    if resposta not in ['S', 'SIM', 'Y', 'YES']:
        print("\n❌ Operação cancelada.")
        input("\nPressione ENTER para sair...")
        return
    
    print("\n" + "─" * 70)
    print("🔨 GERANDO EXECUTÁVEL...")
    print("─" * 70)
    print()
    
    # Comando PyInstaller
    cmd = [
        python_path,
        "-m", "PyInstaller",
        "--onefile",
        "--console",
        str(script_path),
        "--clean",
        "--name", "ProcessadorPropostas"
    ]
    
    try:
        # Executa PyInstaller
        result = subprocess.run(cmd, check=True, capture_output=False)
        
        print("\n" + "=" * 70)
        print("✅ EXECUTÁVEL GERADO COM SUCESSO!")
        print("=" * 70)
        print()
        print(f"📁 Local: dist\\ProcessadorPropostas.exe")
        print()
        print("💡 COMO USAR:")
        print("   1. Copie o arquivo .exe para qualquer pasta")
        print("   2. Execute o .exe")
        print("   3. Ele criará automaticamente as pastas 'leiame' e 'lidos'")
        print("   4. Coloque seus arquivos CSV/Excel na pasta 'leiame'")
        print("   5. Execute o programa e escolha a opção [1]")
        print()
        
    except subprocess.CalledProcessError as e:
        print("\n" + "=" * 70)
        print("❌ ERRO ao gerar executável!")
        print("=" * 70)
        print(f"Código de erro: {e.returncode}")
        print()
    except Exception as e:
        print(f"\n❌ Erro inesperado: {e}")
    
    input("\nPressione ENTER para sair...")

if __name__ == "__main__":
    main()
