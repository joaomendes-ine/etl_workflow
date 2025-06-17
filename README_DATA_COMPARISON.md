# Comparação Inteligente de Dados Excel

## Visão Geral

A funcionalidade de **Comparação Inteligente de Dados Excel** (opção 7 do menu principal) é uma ferramenta avançada para validar a integridade de dados entre ficheiros Excel com estruturas de tabela cruzada (crosstab/pivot). Esta ferramenta foi especificamente desenvolvida para comparar ficheiros "publicados" oficiais com ficheiros "recriados" através de processos ETL.

## Características Principais

### 🎯 **Detecção Automática de Estruturas**
- Identifica automaticamente estruturas de tabela cruzada em ficheiros Excel
- Adapta-se a diferentes layouts e posições de dados
- Suporta células mescladas e estruturas complexas

### 📊 **Normalização Avançada de Dados**
- **Valores Numéricos**: Converte formatos portugueses (ex: "30 602,0" → 30602.0)
- **Dimensões**: Aplica regras inteligentes de normalização
  - "(em branco)" → "Total"
  - "De 16 a 24 anos" → "16 - 24 anos"
  - "65 anos ou mais" → "65+"

### 🔍 **Correspondência Difusa**
- Encontra correspondências mesmo quando etiquetas diferem ligeiramente
- Algoritmo de similaridade configurable (limiar padrão: 80%)
- Suporta variações comuns em nomenclaturas

### 📈 **Relatórios Detalhados**
- Relatórios Excel completos com múltiplas folhas
- Estatísticas de precisão por folha
- Listagem detalhada de todas as discrepâncias
- Detalhes técnicos da comparação

## Estrutura de Ficheiros

```
etl_workflow/
├── dataset/
│   └── comparison/           # Ficheiros "publicados" oficiais
│       └── SL_Turismo_7623.xlsx
├── result/
│   ├── validation/          # Ficheiros "recriados" por pasta
│   │   ├── 55/
│   │   │   ├── series/
│   │   │   │   └── 55_series.xlsx
│   │   │   └── quadros/
│   │   └── ...
│   └── comparison/          # Relatórios de comparação gerados
│       └── comparison_report_YYYYMMDD_HHMMSS.xlsx
```

## Como Usar

### 1. **Preparação dos Ficheiros**
1. Coloque os ficheiros **publicados** em `dataset/comparison/`
2. Coloque os ficheiros **recriados** em `result/validation/[número]/`

### 2. **Execução da Comparação**
1. Execute `python main.py`
2. Escolha a opção **7. Validar integridade dos dados**
3. Siga o assistente interativo:
   - Selecione o ficheiro publicado
   - Selecione o ficheiro recriado
   - Escolha as folhas a comparar
   - Confirme a configuração

### 3. **Análise dos Resultados**
O relatório gerado inclui:
- **Resumo**: Estatísticas gerais de precisão
- **Discrepâncias por Folha**: Detalhes de cada diferença encontrada
- **Detalhes Técnicos**: Informações sobre o processo de comparação

## Configurações Avançadas

### Tolerância Numérica
```python
# Padrão: diferenças menores que 0.01 são ignoradas
numeric_tolerance = 0.01
```

### Padrões de Normalização
```python
dimension_patterns = {
    r'\(em\s+branco\)': 'Total',
    r'de\s+(\d+)\s+a\s+(\d+)': r'\1 - \2',
    r'(\d+)\s*anos?\s*ou\s*mais': r'\1+',
    r'menos\s+de\s+(\d+)\s*anos?': r'< \1',
    # ... mais padrões
}
```

### Limiar de Correspondência Difusa
```python
# Padrão: 80% de similaridade para aceitar correspondência
fuzzy_threshold = 0.8
```

## Exemplos de Uso

### Exemplo 1: Comparação Básica
```bash
# Execute a aplicação principal
python main.py

# Escolha opção 7
# Selecione: SL_Turismo_7623.xlsx (publicado)
# Selecione: result/validation/55/series/55_series.xlsx (recriado)
# Escolha: Todas as folhas (T)
```

### Exemplo 2: Demonstração das Capacidades
```bash
# Execute o script de demonstração
python examples/data_comparison_demo.py
```

## Interpretação dos Resultados

### Códigos de Correspondência
- **exact**: Correspondência exata nas dimensões
- **fuzzy**: Correspondência encontrada via algoritmo difuso
- **not_found**: Valor presente apenas no ficheiro recriado

### Níveis de Precisão
- **> 95%**: 🟢 Excelente - Dados altamente consistentes
- **80-95%**: 🟡 Bom - Algumas discrepâncias menores
- **< 80%**: 🔴 Atenção - Requer revisão das diferenças

### Tipos de Discrepâncias
1. **Diferenças Numéricas**: Valores diferentes entre ficheiros
2. **Valores Ausentes**: Dados presentes apenas num ficheiro
3. **Dimensões Incompatíveis**: Estruturas diferentes

## Casos de Uso Comuns

### 1. **Validação Pós-ETL**
Verificar se os dados processados correspondem aos originais após transformações.

### 2. **Auditoria de Dados**
Comparar versões diferentes de ficheiros para identificar alterações.

### 3. **Controlo de Qualidade**
Garantir que relatórios recriados mantêm a fidelidade aos dados publicados.

### 4. **Migração de Dados**
Validar que dados migrados para novos sistemas mantêm integridade.

## Limitações e Considerações

### Limitações Técnicas
- Ficheiros devem ter estruturas de tabela cruzada similares
- Valores comparados devem ser numéricos ou texto
- Performance pode ser afetada com ficheiros muito grandes (>100MB)

### Considerações de Uso
- A correspondência difusa pode ocasionalmente criar falsos positivos
- Diferenças de formatação podem ser interpretadas como discrepâncias
- Recomenda-se revisar manualmente discrepâncias críticas

## Resolução de Problemas

### Erro: "Nenhuma estrutura detectada"
**Causa**: Ficheiro não segue padrão de tabela cruzada
**Solução**: Verifique se o ficheiro tem dimensões claras nas linhas/colunas

### Erro: "Nenhuma correspondência encontrada"
**Causa**: Estruturas muito diferentes entre ficheiros
**Solução**: Ajuste tolerâncias ou verifique alinhamento dos dados

### Performance Lenta
**Causa**: Ficheiros muito grandes ou complexos
**Solução**: Processe folhas individualmente ou otimize estruturas

## API de Programação

### Uso Básico
```python
from src.data_comparator import DataComparator
from src.utils import setup_logging

# Configuração
logger = setup_logging()
comparator = DataComparator(logger)

# Comparação
results = comparator.compare_files(
    published_file="dataset/comparison/dados.xlsx",
    recreated_file="result/validation/55/dados.xlsx",
    sheet_names=["Folha1", "Folha2"]
)

# Relatório
report_file = comparator.generate_comparison_report(results)
```

### Personalização
```python
# Ajustar tolerância
comparator.numeric_tolerance = 0.001

# Adicionar padrões de normalização
comparator.dimension_patterns[r'novo_padrao'] = 'substituicao'
```

## Suporte e Manutenção

Para questões relacionadas com esta funcionalidade:
1. Consulte os logs detalhados em `logs/`
2. Execute `python examples/data_comparison_demo.py` para verificar funcionamento
3. Revise a documentação técnica em `src/data_comparator.py`

---

**Desenvolvido para o Sistema de Workflow ETL**  
Versão: 1.0.0  
Compatibilidade: Python 3.8+, Windows/Linux/macOS 