#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Exemplo de uso da funcionalidade de consolidação inteligente de dimensões.

Este script demonstra como usar o DimensionConsolidator para analisar e consolidar
colunas de dimensão automaticamente, preservando a integridade dos dados.
"""

import os
import sys
import pandas as pd
import logging
from datetime import datetime

# Adiciona o diretório pai ao path para importar os módulos
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from src.dimension_consolidator import DimensionConsolidator
from src.utils import setup_logging

def create_sample_data():
    """
    Cria dados de exemplo para demonstrar a consolidação de dimensões.
    
    Returns:
        str: Caminho do ficheiro criado
    """
    print("Criando dados de exemplo...")
    
    # Dados de exemplo com padrões de consolidação
    sample_data = {
        'indicador': ['Pop_Total', 'Pop_Masculina', 'Pop_Feminina'] * 20,
        'unidade': ['Número', 'Número', 'Número'] * 20,
        'valor': [100000 + i*1000 for i in range(60)],
        'simbologia': ['', '', ''] * 20,
        
        # Dimensões que podem ser consolidadas
        'dim_grupo_etario1': ['0-14', '15-64', '65+'] * 20,
        'dim_grupo_etario2': ['Jovens', 'Adultos', 'Idosos'] * 20,
        'dim_grupo_etario3': ['', '', ''] * 20,  # Coluna vazia
        
        'dim_setor_economia_cae_rev3_ativ_principal': ['Agricultura', 'Indústria', 'Serviços'] * 20,
        'dim_setor_economia_cae_rev3_ativ_secundaria': ['', 'Manufatura', ''] * 20,
        
        'dim_frequencia_trimestral': ['T1', 'T2', 'T3'] * 20,
        'dim_frequencia_mensal': ['Jan', 'Fev', 'Mar'] * 20,
        'dim_frequencia_anual': ['2020', '2021', '2022'] * 20,
        
        # Dimensões únicas (não devem ser consolidadas)
        'dim_regiao': ['Norte', 'Centro', 'Sul'] * 20,
        'dim_ano': [2020, 2021, 2022] * 20,
    }
    
    df = pd.DataFrame(sample_data)
    
    # Cria diretório se não existir
    os.makedirs('dataset/main', exist_ok=True)
    
    # Guarda o ficheiro Excel
    file_path = 'dataset/main/exemplo_consolidacao.xlsx'
    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='dados', index=False)
    
    print(f"Dados de exemplo criados: {file_path}")
    print(f"Dataset: {len(df)} linhas, {len([col for col in df.columns if col.startswith('dim_')])} colunas de dimensão")
    
    return file_path

def example_basic_consolidation():
    """Exemplo básico de consolidação de dimensões."""
    print("=" * 60)
    print("EXEMPLO 1: CONSOLIDAÇÃO BÁSICA DE DIMENSÕES")
    print("=" * 60)
    
    try:
        # 1. Cria dados de exemplo
        input_file = create_sample_data()
        
        # 2. Configura logging
        logger = setup_logging(verbose=True)
        
        # 3. Inicializa o consolidador
        output_dir = 'result/main'
        os.makedirs(output_dir, exist_ok=True)
        
        consolidator = DimensionConsolidator(input_file, output_dir, logger)
        
        print("\n1. Carregando dados originais...")
        original_df = pd.read_excel(input_file, sheet_name='dados')
        original_dim_count = len([col for col in original_df.columns if col.startswith('dim_')])
        print(f"   Colunas de dimensão originais: {original_dim_count}")
        
        print("\n2. Executando consolidação...")
        result_df = consolidator.consolidate(dry_run=False)
        
        print("\n3. Verificando resultados...")
        result_dim_count = len([col for col in result_df.columns if col.startswith('dim_')])
        reduction = original_dim_count - result_dim_count
        
        print(f"   Colunas de dimensão após consolidação: {result_dim_count}")
        print(f"   Redução: {reduction} colunas ({reduction/original_dim_count*100:.1f}%)")
        
        # Mostra consolidações realizadas
        consolidation_mapping = consolidator.get_consolidation_mapping()
        if consolidation_mapping:
            print("\n   Consolidações realizadas:")
            for new_col, original_cols in consolidation_mapping.items():
                print(f"   [OK] {new_col}")
                print(f"        <- {', '.join(original_cols)}")
        else:
            print("\n4. [AVISO] Nenhuma consolidação foi realizada")
            print("   Os dados podem não ter padrões adequados para consolidação automática.")
        
        # 4. Guarda resultados
        print("\n4. Guardando resultados...")
        output_file = consolidator.save_results(format='excel')
        
        print("\n5. Verificação de integridade...")
        integrity_status = consolidator.get_report().get_integrity_status()
        
        # Mostra status das verificações críticas
        print("   Verificações críticas:")
        critical_checks = ['valor_column', 'row_count', 'non_dimension_columns']
        for check in critical_checks:
            if check in integrity_status:
                status = "[OK] PASSOU" if integrity_status[check] else "[ERRO] FALHOU"
                print(f"   {check}: {status}")
        
        print(f"\n[OK] Exemplo concluído com sucesso!")
        print(f"   Ficheiro original: {input_file}")
        print(f"   Ficheiro consolidado: {output_file}")
        
        return consolidator, original_df, result_df
        
    except Exception as e:
        print(f"\n[ERRO] Erro durante o exemplo: {str(e)}")
        logger.error(f"Erro no exemplo básico: {e}", exc_info=True)
        raise

def example_dry_run_simulation():
    """Exemplo de simulação (dry run) antes da consolidação."""
    print("\n" + "=" * 60)
    print("EXEMPLO 2: SIMULAÇÃO ANTES DA CONSOLIDAÇÃO")
    print("=" * 60)
    
    try:
        # 1. Usa os mesmos dados do exemplo anterior
        input_file = 'dataset/main/exemplo_consolidacao.xlsx'
        
        if not os.path.exists(input_file):
            input_file = create_sample_data()
        
        # 2. Configura logging menos verboso para simulação
        logger = setup_logging(verbose=False)
        
        # 3. Inicializa consolidador
        output_dir = 'result/main'
        consolidator = DimensionConsolidator(input_file, output_dir, logger)
        
        print("\n1. Executando simulação (dry run)...")
        result_df = consolidator.consolidate(dry_run=True)
        
        print("\n2. Analisando resultados da simulação...")
        
        # Obtém detalhes das ações planeadas
        planned_actions = consolidator.get_report().get_consolidation_details()
        
        consolidate_actions = [action for action in planned_actions if action['action_type'] == 'consolidate']
        skip_actions = [action for action in planned_actions if action['action_type'] == 'skip']
        
        print(f"   Ações planeadas:")
        print(f"   - Consolidações: {len(consolidate_actions)}")
        print(f"   - Ignoradas: {len(skip_actions)}")
        
        # Mostra consolidações planeadas
        planned_consolidations = consolidator.get_consolidation_mapping()
        if planned_consolidations:
            print(f"\n   [OK] Consolidações planeadas ({len(planned_consolidations)}):")
            for new_col, original_cols in planned_consolidations.items():
                print(f"   {new_col}")
                print(f"      <- {', '.join(original_cols)}")
        
        print("\n3. Exibindo resumo da simulação...")
        consolidator.print_summary()
        
        print("\n4. A simulação não alterou ficheiros!")
        print("   Para aplicar as alterações, execute consolidate(dry_run=False)")
        
        # 5. Agora aplica as alterações realmente
        print("\n5. Aplicando consolidação real...")
        real_result_df = consolidator.consolidate(dry_run=False)
        
        # 6. Guarda resultados
        output_file = consolidator.save_results(format='excel')
        
        print(f"\n[OK] Consolidação aplicada com sucesso!")
        print(f"   Ficheiro guardado: {output_file}")
        
        return consolidator, result_df, real_result_df
        
    except Exception as e:
        print(f"\n[ERRO] Erro durante a simulação: {str(e)}")
        logger.error(f"Erro no exemplo de simulação: {e}", exc_info=True)
        raise

def example_advanced_configuration():
    """Exemplo com configurações avançadas e exclusão de colunas."""
    print("\n" + "=" * 60)
    print("EXEMPLO 3: CONFIGURAÇÃO AVANÇADA")
    print("=" * 60)
    
    try:
        # 1. Cria dados mais complexos
        input_file = create_sample_data()
        
        # 2. Configura logger
        logger = setup_logging(verbose=True)
        
        # 3. Configurações avançadas
        output_dir = 'result/main'
        consolidator = DimensionConsolidator(input_file, output_dir, logger)
        
        print("\n1. Configuração com exclusão de colunas...")
        
        # Lista as colunas de dimensão disponíveis
        df = pd.read_excel(input_file, sheet_name='dados')
        dim_columns = [col for col in df.columns if col.startswith('dim_')]
        
        print(f"   Colunas de dimensão disponíveis: {len(dim_columns)}")
        for col in dim_columns[:5]:  # Mostra apenas as primeiras 5
            print(f"   - {col}")
        if len(dim_columns) > 5:
            print(f"   ... mais {len(dim_columns) - 5} colunas")
        
        # Exclui algumas colunas específicas da consolidação
        exclude_columns = ['dim_ano', 'dim_regiao']
        print(f"\n2. Excluindo colunas da consolidação: {exclude_columns}")
        
        # 3. Executa consolidação com exclusões
        result_df = consolidator.consolidate(
            dry_run=False, 
            exclude_columns=exclude_columns
        )
        
        print("\n3. Verificando configurações aplicadas...")
        
        # Verifica se as colunas excluídas ainda existem
        excluded_still_present = [col for col in exclude_columns if col in result_df.columns]
        print(f"   Colunas excluídas preservadas: {excluded_still_present}")
        
        # 4. Análise detalhada dos resultados
        print("\n4. Análise detalhada...")
        
        report = consolidator.get_report()
        integrity_status = report.get_integrity_status()
        
        print("   Status de integridade:")
        for check_name, status in integrity_status.items():
            icon = "[OK]" if status else "[ERRO]"
            print(f"   {check_name}: {icon}")
        
        # 5. Guarda resultados com metadata
        output_file = consolidator.save_results(format='excel')
        
        # 6. Guarda relatório detalhado
        report_file = os.path.join(output_dir, 'consolidation_advanced_report.json')
        report.save_report(report_file, detailed=True)
        
        print(f"\n[OK] Exemplo avançado concluído!")
        print(f"   Resultado: {output_file}")
        print(f"   Relatório: {report_file}")
        
        return consolidator, result_df
        
    except Exception as e:
        print(f"\n[ERRO] Erro no exemplo avançado: {str(e)}")
        logger.error(f"Erro no exemplo avançado: {e}", exc_info=True)
        raise

def example_batch_processing():
    """Exemplo de processamento em lote de múltiplos ficheiros."""
    print("\n" + "=" * 60)
    print("EXEMPLO 4: PROCESSAMENTO EM LOTE")
    print("=" * 60)
    
    try:
        # 1. Cria múltiplos ficheiros de exemplo
        print("\n1. Criando múltiplos ficheiros de exemplo...")
        
        input_files = []
        for i in range(3):
            # Modifica ligeiramente os dados para cada ficheiro
            sample_data = create_sample_data()
            new_path = f'dataset/main/exemplo_lote_{i+1}.xlsx'
            
            # Copia e renomeia
            import shutil
            shutil.copy2(sample_data, new_path)
            input_files.append(new_path)
            
        print(f"   Criados {len(input_files)} ficheiros para processamento")
        
        # 2. Configurações de processamento
        logger = setup_logging(verbose=False)
        output_dir = 'result/main/batch_processing'
        os.makedirs(output_dir, exist_ok=True)
        
        # 3. Processa cada ficheiro
        print("\n2. Processando ficheiros em lote...")
        
        batch_configs = [
            {'name': 'Configuração Padrão', 'exclude_columns': []},
            {'name': 'Excluindo Ano e Região', 'exclude_columns': ['dim_ano', 'dim_regiao']},
            {'name': 'Apenas Simulação', 'exclude_columns': [], 'dry_run': True}
        ]
        
        results = []
        
        for i, input_file in enumerate(input_files):
            config = batch_configs[i % len(batch_configs)]
            print(f"\n   Processando ficheiro {i+1}: {os.path.basename(input_file)}")
            print(f"   [OK] {config['name']}")
            
            try:
                # Cria subdiretório para cada ficheiro
                file_output_dir = os.path.join(output_dir, f'file_{i+1}')
                os.makedirs(file_output_dir, exist_ok=True)
                
                consolidator = DimensionConsolidator(input_file, file_output_dir, logger)
                
                # Aplica configuração
                result_df = consolidator.consolidate(
                    dry_run=config.get('dry_run', False),
                    exclude_columns=config['exclude_columns']
                )
                
                # Guarda resultados se não for dry_run
                if not config.get('dry_run', False):
                    output_file = consolidator.save_results()
                    print(f"      [OK] Processado: {len(consolidator.get_consolidation_mapping())} consolidações")
                    print(f"      Resultado: {os.path.basename(output_file)}")
                else:
                    print(f"      [SIM] Simulação concluída")
                
                results.append({
                    'input_file': input_file,
                    'config': config,
                    'success': True,
                    'consolidations': len(consolidator.get_consolidation_mapping())
                })
                
            except Exception as e:
                print(f"      [ERRO] Erro: {str(e)}")
                results.append({
                    'input_file': input_file,
                    'config': config,
                    'success': False,
                    'error': str(e)
                })
        
        # 4. Resumo do processamento em lote
        print("\n3. Resumo do processamento em lote:")
        successful = sum(1 for r in results if r['success'])
        print(f"   Ficheiros processados: {successful}/{len(results)}")
        
        total_consolidations = sum(r.get('consolidations', 0) for r in results if r['success'])
        print(f"   Total de consolidações: {total_consolidations}")
        
        print(f"\n[OK] Processamento em lote concluído!")
        print(f"   Resultados em: {output_dir}")
        
        return results
        
    except Exception as e:
        print(f"\n[ERRO] Erro no processamento em lote: {str(e)}")
        raise

def main():
    """Executa todos os exemplos de consolidação de dimensões."""
    print("EXEMPLOS DE CONSOLIDAÇÃO INTELIGENTE DE DIMENSÕES")
    print("=" * 80)
    
    try:
        # Exemplo 1: Consolidação básica
        consolidator1, original_df, result_df = example_basic_consolidation()
        
        # Exemplo 2: Simulação antes da consolidação
        consolidator2, sim_df, real_df = example_dry_run_simulation()
        
        # Exemplo 3: Configuração avançada
        consolidator3, advanced_df = example_advanced_configuration()
        
        # Exemplo 4: Processamento em lote
        batch_results = example_batch_processing()
        
        print("\n" + "=" * 80)
        print("TODOS OS EXEMPLOS CONCLUÍDOS COM SUCESSO!")
        print("=" * 80)
        print("\nPróximos passos:")
        print("1. Verifique os ficheiros gerados em 'result/main/'")
        print("2. Consulte os relatórios JSON para análise detalhada")
        print("3. Adapte os exemplos aos seus dados reais")
        print("4. Use as configurações avançadas conforme necessário")
        
    except Exception as e:
        print(f"\n[ERRO] Erro durante a execução dos exemplos: {str(e)}")
        raise

if __name__ == "__main__":
    main() 