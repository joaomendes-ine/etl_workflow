"""
Demonstração do Sistema Redesenhado de Comparação Hierárquica de Dados Excel

Este script demonstra as capacidades do novo sistema de comparação:
- Estrutura hierárquica multi-nível
- Equivalência semântica
- Detecção baseada em cores
- Matching inteligente
"""

import sys
import os

# Adiciona o diretório raiz ao path para importar módulos
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))

from src.data_comparator import HierarchicalDataComparator, DataPoint
from src.utils import setup_logging
from colorama import Fore, Style, init

# Inicializa colorama
init()


def demo_semantic_equivalence():
    """Demonstra o sistema de equivalência semântica."""
    print(f"\n{Fore.CYAN}=== DEMONSTRAÇÃO: Equivalência Semântica ==={Style.RESET_ALL}")
    
    logger = setup_logging(verbose=False)
    comparator = HierarchicalDataComparator(logger)
    
    # Exemplos de equivalência
    equivalence_examples = [
        ("****", "4"),
        ("***", "3"),
        ("Total", "(Em branco)"),
        ("(em branco)", "Total"),
        ("**", "2"),
        ("*", "1")
    ]
    
    print("Mapeamentos de equivalência semântica:")
    for original, equivalent in equivalence_examples:
        mapped = comparator.get_semantic_equivalent(original)
        status = "✓" if mapped == equivalent else "✗"
        print(f"  {status} '{original}' → '{mapped}' (esperado: '{equivalent}')")


def demo_coordinate_matching():
    """Demonstra o sistema de correspondência de coordenadas inteligente."""
    print(f"\n{Fore.CYAN}=== DEMONSTRAÇÃO: Correspondência de Coordenadas ==={Style.RESET_ALL}")
    
    logger = setup_logging(verbose=False)
    comparator = HierarchicalDataComparator(logger)
    
    # Exemplos de pontos de dados para teste
    test_cases = [
        # Caso 1: Correspondência exata
        {
            "name": "Correspondência Exata",
            "point1": DataPoint(
                column_level_1="Hóteis", column_level_2="4",
                row_level_1="Estrangeiro", row_level_2="Alemanha",
                value=92767.0
            ),
            "point2": DataPoint(
                column_level_1="Hóteis", column_level_2="4",
                row_level_1="Estrangeiro", row_level_2="Alemanha",
                value=92767.0
            ),
            "expected": True
        },
        
        # Caso 2: Equivalência semântica (****  = 4)
        {
            "name": "Equivalência Semântica (****  = 4)",
            "point1": DataPoint(
                column_level_1="Hóteis", column_level_2="****",
                row_level_1="Estrangeiro", row_level_2="Alemanha",
                value=92767.0
            ),
            "point2": DataPoint(
                column_level_1="Hóteis", column_level_2="4",
                row_level_1="1976", row_level_2="Alemanha",
                value=92767.0
            ),
            "expected": True
        },
        
        # Caso 3: Equivalência de Total
        {
            "name": "Equivalência de Total",
            "point1": DataPoint(
                column_level_1="Total", column_level_2="",
                row_level_1="Total", row_level_2="",
                value=1947611.0
            ),
            "point2": DataPoint(
                column_level_1="(em branco)", column_level_2="(em branco)",
                row_level_1="1976", row_level_2="(em branco)",
                value=1947611.0
            ),
            "expected": True
        },
        
        # Caso 4: Não correspondência (diferentes valores de dimensão)
        {
            "name": "Não Correspondência (dimensões diferentes)",
            "point1": DataPoint(
                column_level_1="1979", column_level_2="",
                row_level_1="Hotelaria", row_level_2="Motéis",
                value=14621.0
            ),
            "point2": DataPoint(
                column_level_1="1979", column_level_2="",
                row_level_1="Hotelaria", row_level_2="Apartamentos",
                value=14621.0
            ),
            "expected": False
        }
    ]
    
    print("Testando correspondência de coordenadas:")
    for case in test_cases:
        result = comparator.smart_coordinate_match(case["point1"], case["point2"])
        status = "✓" if result == case["expected"] else "✗"
        color = Fore.GREEN if result == case["expected"] else Fore.RED
        
        print(f"  {status} {color}{case['name']}{Style.RESET_ALL}")
        print(f"    Ponto 1: {case['point1']}")
        print(f"    Ponto 2: {case['point2']}")
        print(f"    Resultado: {result} (esperado: {case['expected']})")
        print()


def demo_value_normalization():
    """Demonstra a normalização conservadora de valores."""
    print(f"\n{Fore.CYAN}=== DEMONSTRAÇÃO: Normalização Conservadora de Valores ==={Style.RESET_ALL}")
    
    logger = setup_logging(verbose=False)
    comparator = HierarchicalDataComparator(logger)
    
    test_values = [
        # Valores que devem ser normalizados
        ("30 602,0", 30602.0, True),
        ("1 234,56", 1234.56, True),
        ("100,00", 100.0, True),
        ("-50,25", -50.25, True),
        (123.45, 123.45, True),
        ("0", 0.0, True),
        
        # Valores que devem ser rejeitados (anos/dimensões)
        ("1976", None, False),
        ("2023", None, False),
        ("1850", None, False),
        
        # Valores que devem ser rejeitados (texto)
        ("Total", None, False),
        ("N/A", None, False),
        ("(em branco)", None, False),
        ("-", None, False),
        ("", None, False),
        (None, None, False)
    ]
    
    print("Testando normalização de valores:")
    for original, expected, should_normalize in test_values:
        result = comparator.normalize_value_conservative(original)
        
        if should_normalize:
            status = "✓" if result == expected else "✗"
            color = Fore.GREEN if result == expected else Fore.RED
        else:
            status = "✓" if result is None else "✗"
            color = Fore.GREEN if result is None else Fore.RED
        
        print(f"  {status} {color}'{original}' → {result}{Style.RESET_ALL} (esperado: {expected})")


def demo_data_point_creation():
    """Demonstra a criação e manipulação de pontos de dados."""
    print(f"\n{Fore.CYAN}=== DEMONSTRAÇÃO: Estrutura de Pontos de Dados ==={Style.RESET_ALL}")
    
    # Exemplos de pontos de dados hierárquicos
    examples = [
        DataPoint(
            column_level_1="Hóteis",
            column_level_2="4",
            row_level_1="Estrangeiro",
            row_level_2="Alemanha",
            value=92767.0,
            row=15,
            col=8
        ),
        DataPoint(
            column_level_1="Total",
            column_level_2="",
            row_level_1="Total",
            row_level_2="",
            value=1947611.0,
            row=2,
            col=2
        ),
        DataPoint(
            column_level_1="1979",
            column_level_2="",
            row_level_1="Hotelaria",
            row_level_2="Motéis",
            value=14621.0,
            row=25,
            col=12
        )
    ]
    
    print("Exemplos de estrutura de pontos de dados:")
    for i, point in enumerate(examples, 1):
        print(f"  {Fore.YELLOW}Exemplo {i}:{Style.RESET_ALL}")
        print(f"    Coordenadas: {point.get_coordinate_key()}")
        print(f"    String: {point}")
        print(f"    Posição: Linha {point.row}, Coluna {point.col}")
        print()


def demo_comparison_summary():
    """Mostra como seria um resumo de comparação."""
    print(f"\n{Fore.CYAN}=== DEMONSTRAÇÃO: Resumo de Comparação ==={Style.RESET_ALL}")
    
    # Simula resultados de comparação
    mock_results = {
        'published_file': 'dataset/comparison/employment_data.xlsx',
        'recreated_file': 'result/validation/65/65_BD.xlsx',
        'timestamp': '2024-01-15T14:30:00',
        'sheets': {
            'Dados_Principais': {
                'sheet_name': 'Dados_Principais',
                'published_points': 24318,
                'recreated_points': 24318,
                'exact_matches': 22150,
                'semantic_matches': 2100,
                'value_differences': 68,
                'missing_in_published': 0,
                'missing_in_recreated': 0
            }
        },
        'summary': {
            'total_published_points': 24318,
            'total_recreated_points': 24318,
            'total_matches': 24250,
            'total_differences': 68,
            'total_missing_in_published': 0,
            'total_missing_in_recreated': 0,
            'accuracy_percentage': 99.72
        }
    }
    
    print("Exemplo de resumo de comparação:")
    summary = mock_results['summary']
    
    print(f"  [F] Ficheiro Publicado: {mock_results['published_file']}")
    print(f"  [F] Ficheiro Recriado: {mock_results['recreated_file']}")
    print(f"  [D] Data: {mock_results['timestamp']}")
    print()
    
    print(f"  [S] {Fore.GREEN}Estatísticas Gerais:{Style.RESET_ALL}")
    print(f"    • Pontos Publicados: {summary['total_published_points']:,}")
    print(f"    • Pontos Recriados: {summary['total_recreated_points']:,}")
    print(f"    • Correspondências: {summary['total_matches']:,}")
    print(f"    • Diferenças: {summary['total_differences']:,}")
    print(f"    • Precisão: {summary['accuracy_percentage']:.2f}%")
    print()
    
    print(f"  [D] {Fore.BLUE}Detalhes por Folha:{Style.RESET_ALL}")
    for sheet_name, sheet_data in mock_results['sheets'].items():
        accuracy = ((sheet_data['exact_matches'] + sheet_data['semantic_matches']) / 
                   sheet_data['recreated_points']) * 100
        print(f"    • {sheet_name}:")
        print(f"      - Correspondências Exatas: {sheet_data['exact_matches']:,}")
        print(f"      - Correspondências Semânticas: {sheet_data['semantic_matches']:,}")
        print(f"      - Diferenças de Valor: {sheet_data['value_differences']:,}")
        print(f"      - Precisão: {accuracy:.2f}%")


def main():
    """Executa todas as demonstrações."""
    print(f"{Fore.MAGENTA}{'='*80}{Style.RESET_ALL}")
    print(f"{Fore.MAGENTA}  DEMONSTRAÇÃO: Sistema Redesenhado de Comparação Hierárquica{Style.RESET_ALL}")
    print(f"{Fore.MAGENTA}  Soluções para Validação de Dados Estatísticos Portugueses{Style.RESET_ALL}")
    print(f"{Fore.MAGENTA}{'='*80}{Style.RESET_ALL}")
    
    print(f"\n{Fore.WHITE}Este sistema foi redesenhado para resolver os seguintes problemas:{Style.RESET_ALL}")
    print("• [X] Complexidade excessiva de detecção → [OK] Lógica simplificada baseada em cores")
    print("• [X] Falsos positivos em correspondência → [OK] Equivalência semântica precisa")
    print("• [X] Sistema de coordenadas problemático → [OK] Estrutura hierárquica multi-nível")
    print("• [X] Normalização que altera valores → [OK] Preservação conservadora de valores")
    print("• [X] Matching fuzzy excessivo → [OK] Correspondência inteligente controlada")
    
    try:
        demo_semantic_equivalence()
        demo_coordinate_matching()
        demo_value_normalization()
        demo_data_point_creation()
        demo_comparison_summary()
        
        print(f"\n{Fore.GREEN}{'='*80}{Style.RESET_ALL}")
        print(f"{Fore.GREEN}  [OK] DEMONSTRAÇÃO CONCLUÍDA COM SUCESSO{Style.RESET_ALL}")
        print(f"{Fore.GREEN}  O sistema está pronto para validar dados estatísticos reais{Style.RESET_ALL}")
        print(f"{Fore.GREEN}{'='*80}{Style.RESET_ALL}")
        
        print(f"\n{Fore.CYAN}[!] Próximos Passos:{Style.RESET_ALL}")
        print("1. Adicione ficheiros de teste em dataset/comparison/")
        print("2. Execute a opção 7 (Validar dados) no menu principal")
        print("3. Verifique os relatórios gerados em result/comparison/")
        
    except Exception as e:
        print(f"\n{Fore.RED}[X] Erro durante demonstração: {e}{Style.RESET_ALL}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main() 