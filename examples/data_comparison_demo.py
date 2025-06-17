"""
Demonstração da funcionalidade de Comparação Inteligente de Dados Excel

Este script demonstra como usar o módulo data_comparator para comparar
ficheiros Excel com estruturas de tabela cruzada.
"""

import sys
import os

# Adiciona o diretório raiz ao path para importar módulos
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))

from src.data_comparator import DataComparator
from src.utils import setup_logging


def demo_data_comparison():
    """Demonstra o uso básico do comparador de dados."""
    
    # Configura logging
    logger = setup_logging(verbose=True)
    
    print("=" * 80)
    print("DEMONSTRAÇÃO: Comparação Inteligente de Dados Excel")
    print("=" * 80)
    
    # Cria instância do comparador
    comparator = DataComparator(logger)
    
    # Mostra ficheiros disponíveis
    print("\n1. FICHEIROS DISPONÍVEIS")
    print("-" * 40)
    
    published_files, recreated_files = comparator.get_available_files()
    
    print(f"Ficheiros publicados encontrados: {len(published_files)}")
    for file in published_files:
        print(f"  • {file}")
    
    print(f"\nFicheiros recriados encontrados: {len(recreated_files)} pastas")
    for folder, files in list(recreated_files.items())[:5]:  # Mostra apenas as primeiras 5 pastas
        print(f"  • Pasta {folder}: {len(files)} ficheiro(s)")
    
    if len(recreated_files) > 5:
        print(f"  ... mais {len(recreated_files) - 5} pastas")
    
    # Exemplo de normalização de valores
    print("\n2. EXEMPLO DE NORMALIZAÇÃO DE VALORES")
    print("-" * 40)
    
    test_values = [
        "30 602,0",
        "1 234,56",
        "100,00",
        "-50,25",
        None,
        "",
        "N/A",
        123.45
    ]
    
    print("Valores de teste e sua normalização:")
    for value in test_values:
        normalized = comparator.normalize_value(value)
        print(f"  '{value}' → {normalized}")
    
    # Exemplo de normalização de dimensões
    print("\n3. EXEMPLO DE NORMALIZAÇÃO DE DIMENSÕES")
    print("-" * 40)
    
    test_dimensions = [
        "(em branco)",
        "16 - 24 anos",
        "De 16 a 24 anos",
        "65 anos ou mais",
        "menos de 18 anos",
        "n.d.",
        "-"
    ]
    
    print("Dimensões de teste e sua normalização:")
    for dim in test_dimensions:
        normalized = comparator.normalize_dimension_label(dim)
        print(f"  '{dim}' → '{normalized}'")
    
    # Exemplo de correspondência difusa
    print("\n4. EXEMPLO DE CORRESPONDÊNCIA DIFUSA")
    print("-" * 40)
    
    candidates = ["16 - 24 anos", "25 - 34 anos", "35 - 44 anos", "Total"]
    target = "De 16 a 24 anos"
    
    match = comparator.fuzzy_match_dimension(target, candidates)
    print(f"Procurando '{target}' em {candidates}")
    print(f"Melhor correspondência: '{match}'")
    
    print("\n5. ESTRUTURA DE FICHEIROS PARA COMPARAÇÃO")
    print("-" * 40)
    print("Para usar a comparação:")
    print("1. Coloque ficheiros 'publicados' em: dataset/comparison/")
    print("2. Coloque ficheiros 'recriados' em: result/validation/[número]/")
    print("3. Execute a opção 7 no menu principal")
    print("4. O relatório será gerado em: result/comparison/")
    
    print("\n6. CAPACIDADES TÉCNICAS AVANÇADAS")
    print("-" * 40)
    print("• 🎯 Detecção inteligente de tabelas principais (ignora seções 'Filtros')")
    print("• 🎨 Análise de formatação visual (identifica dados vs cabeçalhos)")
    print("• 🔗 Tratamento avançado de células mescladas")
    print("• 📊 Valores apresentados (usa number_format para comparação exata)")
    print("• 🌐 Normalização de valores portugueses ('30 602,0' → 30602.0)")
    print("• 🔍 Correspondência difusa para dimensões similares")
    print("• ⚡ Algoritmo 'andar para cima/esquerda' para mapear dimensões")
    print("• 🎨 Relatórios visuais com destaques amarelos em discrepâncias")
    print("• 📋 Cópia perfeita de formatação original nas folhas de resumo")
    print("• 💬 Comentários automáticos com detalhes das discrepâncias")
    
    print("\n" + "=" * 80)
    print("Demonstração concluída!")
    print("Execute 'python main.py' e escolha a opção 7 para usar interativamente.")
    print("=" * 80)


if __name__ == "__main__":
    demo_data_comparison() 