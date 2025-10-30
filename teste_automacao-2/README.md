# 📋 PROCESSADOR DE PROPOSTAS WIIPO

## 🚀 Como usar o executável

### 1️⃣ Primeira vez (gerando o .exe)

**Opção A - Automática:**
Execute o arquivo `gerar_executavel.py` (duplo clique ou via Python)

**Opção B - Manual:**

```cmd
C:/Users/KawanWagnnerGonçalve/AppData/Local/Python/pythoncore-3.14-64/python.exe -m PyInstaller --onefile --console processar_propostas.py --clean --name ProcessadorPropostas
```

Isso irá:

- Compilar o script Python
- Gerar o arquivo `ProcessadorPropostas.exe` na pasta `dist/`
- Limpar arquivos temporários

### 2️⃣ Usando o programa

1. **Copie** o arquivo `ProcessadorPropostas.exe` da pasta `dist/` para onde você quiser
2. **Execute** `ProcessadorPropostas.exe` (as pastas são criadas automaticamente)
3. O programa criará automaticamente:
   - `leiame/` (coloque os arquivos aqui)
   - `lidos/` (arquivos processados vão aqui)
   - `backup/` (backup do último processamento)

### 📂 Estrutura de pastas

```
📁 MinhaPasta/
├── ProcessadorPropostas.exe
├── leiame/                    (coloque 1 arquivo CSV/Excel aqui)
├── lidos/                     (arquivos processados)
├── backup/                    (backup do último arquivo formatado)
└── relatorio_propostas_formatado.xlsx  (resultado final)
```

### ✨ Funcionalidades

#### Menu Principal:

- **[1] Processar arquivo**: Processa CSV/Excel e gera arquivo formatado
  - Detecta se Excel está aberto
  - Opção 1: Fecha Excel automaticamente
  - Opção 2: Cria arquivo temporário (abre em nova aba sem mexer no Excel aberto)
  - Opção 3: Cancela operação
- **[2] Ver arquivos processados**: Lista arquivos na pasta 'lidos'
- **[3] Limpar pasta 'lidos'**: Remove todos os arquivos processados
- **[4] Ver arquivo formatado existente**: Abre o último arquivo gerado
- **[0] Sair**: Fecha o programa

### ⚙️ O que o programa faz

✅ **Detecção inteligente de pasta**: Funciona em qualquer PC, em qualquer pasta  
✅ **Proteção contra Excel aberto**: Detecta e oferece opções seguras  
✅ **Modo temporário**: Cria arquivo temporário se Excel estiver aberto  
✅ **Lê arquivos**: CSV ou Excel (.xlsx, .xls)  
✅ **Extrai dados JSON**: Processa colunas 'Margens Prev', 'Motivo Inelegibilidade'  
✅ **Gera Excel formatado**: Colunas organizadas e normalizadas  
✅ **Backup automático**: Mantém backup do último arquivo processado  
✅ **Move originais**: Arquivos processados vão para pasta 'lidos'  
✅ **Validação**: Garante que só há 1 arquivo por vez  
✅ **Limpeza automática**: Remove arquivos temporários antigos

### 📝 Observações importantes

- ⚠️ **Só coloque 1 arquivo por vez** na pasta `leiame/`
- 📁 O arquivo formatado é salvo como `relatorio_propostas_formatado.xlsx`
- 🔄 O arquivo original é **movido** (não copiado) para `lidos/`
- 💾 Backup do último processamento fica em `backup/`
- 🗑️ Arquivos temporários (TEMP\_\*.xlsx) são removidos automaticamente

### 🔧 Recursos avançados

**Modo Temporário (quando Excel está aberto):**

- Cria arquivo `relatorio_propostas_TEMP_[timestamp].xlsx`
- Abre em nova aba do Excel sem conflitos
- Arquivo temporário é removido na próxima execução
- Permite trabalhar com Excel aberto sem riscos

**Sistema de Backup:**

- Mantém sempre o último arquivo formatado em `backup/`
- Substitui backup anterior automaticamente
- Garante que você nunca perca o último processamento

### 🆘 Suporte

Em caso de erro, o programa:

- Mostra mensagem detalhada na tela
- Mantém a janela aberta para você ler o erro
- Oferece opção de visualizar arquivo mesmo com erro

### 🎯 Compatibilidade

✅ Funciona em **qualquer PC** Windows  
✅ Funciona em **qualquer pasta** (Desktop, Downloads, Documentos, etc.)  
✅ **Não precisa** de Python instalado (executável standalone)  
✅ **Cria** todas as pastas automaticamente  
✅ **Detecta** automaticamente onde está rodando
