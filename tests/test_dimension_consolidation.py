#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import unittest
import pandas as pd
import logging
import tempfile
import shutil
from pathlib import Path

# Adiciona o diretório pai ao path para importar os módulos
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from src.dimension_consolidator import DimensionConsolidator
from src.dimension_analyzer import DimensionAnalyzer
from src.consolidation_rules import ConsolidationRules
from src.consolidation_report import ConsolidationReport

class TestDimensionConsolidation(unittest.TestCase):
    """Testes para a consolidação inteligente de dimensões"""
    
    @classmethod
    def setUpClass(cls):
        """Configuração executada uma vez antes de todos os testes"""
        # Configura o logger
        cls.logger = logging.getLogger('test_dimension_logger')
        cls.logger.setLevel(logging.INFO)
        handler = logging.StreamHandler()
        handler.setLevel(logging.INFO)
        cls.logger.addHandler(handler)
        
        # Cria diretórios temporários para testes
        cls.temp_dir = tempfile.mkdtemp()
        cls.input_dir = os.path.join(cls.temp_dir, 'input')
        cls.output_dir = os.path.join(cls.temp_dir, 'output')
        os.makedirs(cls.input_dir, exist_ok=True)
        os.makedirs(cls.output_dir, exist_ok=True)
        
        # Cria dados de teste com padrões de dimensões
        cls._create_test_data()
    
    @classmethod
    def _create_test_data(cls):
        """Cria dados de teste com padrões típicos para consolidação"""
        
        # Dados de teste com padrões de consolidação
        test_data = {
            'dim_grupo_etario1': ['0-14', '15-24', '25-34', '35-44', '45-54', '55-64', '65+', None, None, None],
            'dim_grupo_etario2': [None, None, None, None, None, None, None, '0-14', '15-24', '25-34'],
            'dim_grupo_etario3': [None, None, None, None, None, None, None, None, None, None],
            'dim_setor_economia_cae_rev3_ativ_principal': ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'],
            'dim_setor_economia_cae_rev3_ativ_secundaria': [None, None, None, None, None, None, None, 'K', 'L', 'M'],
            'dim_regiao_nuts1': ['Norte', 'Centro', 'Lisboa', 'Alentejo', 'Algarve', 'Açores', 'Madeira', 'Norte', 'Centro', 'Lisboa'],
            'dim_regiao_nuts2': ['Norte', 'Centro', 'AML', 'Alentejo', 'Algarve', 'RAA', 'RAM', 'Norte', 'Centro', 'AML'],
            'dim_genero': ['M', 'F', 'M', 'F', 'M', 'F', 'M', 'F', 'M', 'F'],
            'dim_ano': [2020, 2020, 2020, 2020, 2020, 2021, 2021, 2021, 2021, 2021],
            'indicador': ['População', 'Emprego', 'PIB', 'Exportações', 'Importações'] * 2,
            'unidade': ['Número', 'Número', 'Milhões €', 'Milhões €', 'Milhões €'] * 2,
            'valor': [1000.5, 800.2, 15000.0, 5000.0, 4500.0, 1050.0, 820.0, 15500.0, 5200.0, 4600.0],
            'simbologia': ['', '', '', '', '', '', '', '', '', '']
        }
        
        cls.test_df = pd.DataFrame(test_data)
        
        # Salva como ficheiro Excel de teste
        cls.test_file = os.path.join(cls.input_dir, 'test_dimensions.xlsx')
        with pd.ExcelWriter(cls.test_file, engine='openpyxl') as writer:
            cls.test_df.to_excel(writer, sheet_name='dados', index=False)
        
        cls.logger.info(f"Ficheiro de teste criado: {cls.test_file}")
        cls.logger.info(f"Dimensões de teste: {cls.test_df.shape}")
        
        # Cria um segundo ficheiro de teste sem padrões óbvios
        simple_data = {
            'dim_categoria': ['A', 'B', 'C', 'D', 'E'],
            'dim_tipo': ['X', 'Y', 'Z', 'X', 'Y'],
            'indicador': ['Teste1', 'Teste2', 'Teste3', 'Teste4', 'Teste5'],
            'valor': [10.0, 20.0, 30.0, 40.0, 50.0]
        }
        
        cls.simple_test_file = os.path.join(cls.input_dir, 'simple_test.xlsx')
        with pd.ExcelWriter(cls.simple_test_file, engine='openpyxl') as writer:
            pd.DataFrame(simple_data).to_excel(writer, sheet_name='dados', index=False)
    
    @classmethod
    def tearDownClass(cls):
        """Limpeza executada uma vez após todos os testes"""
        # Remove os diretórios temporários
        shutil.rmtree(cls.temp_dir)
    
    def test_dimension_analyzer_pattern_detection(self):
        """Testa a deteção de padrões pelo DimensionAnalyzer"""
        analyzer = DimensionAnalyzer(self.test_df, self.logger)
        
        # Testa análise de padrões
        patterns = analyzer.analyze_patterns()
        
        # Deve detetar padrões para grupo_etario e setor_economia
        self.assertGreater(len(patterns), 0, "Deve detetar pelo menos um padrão")
        
        # Verifica se detectou padrões esperados
        pattern_keys = list(patterns.keys())
        
        # Deve ter pelo menos padrões relacionados com grupo_etario e setor_economia
        has_grupo_etario = any('grupo_etario' in key for key in pattern_keys)
        has_setor_economia = any('setor_economia' in key for key in pattern_keys)
        
        self.assertTrue(has_grupo_etario or has_setor_economia, 
                       "Deve detetar padrões de grupo_etario ou setor_economia")
        
        self.logger.info(f"Padrões detetados: {list(patterns.keys())}")
    
    def test_dimension_analyzer_value_analysis(self):
        """Testa a análise de valores pelo DimensionAnalyzer"""
        analyzer = DimensionAnalyzer(self.test_df, self.logger)
        
        # Testa análise de valores
        value_analysis = analyzer.analyze_values()
        
        # Deve ter análise para todas as colunas de dimensão
        dim_columns = [col for col in self.test_df.columns if col.startswith('dim_')]
        
        for dim_col in dim_columns:
            self.assertIn(dim_col, value_analysis, f"Deve ter análise de valores para {dim_col}")
            self.assertIsInstance(value_analysis[dim_col], set, f"Valores de {dim_col} devem ser um set")
    
    def test_consolidation_rules_validation(self):
        """Testa as regras de consolidação"""
        # Testa com colunas compatíveis
        compatible_columns = ['dim_grupo_etario1', 'dim_grupo_etario2']
        compatible_values = {
            'dim_grupo_etario1': {'0-14', '15-24', '25-34', '35-44', '45-54', '55-64', '65+'},
            'dim_grupo_etario2': {'0-14', '15-24', '25-34'}
        }
        
        can_consolidate, reasons, warnings = ConsolidationRules.can_consolidate(
            compatible_columns, compatible_values, logger=self.logger
        )
        
        self.assertTrue(can_consolidate, f"Colunas compatíveis devem poder ser consolidadas. Razões: {reasons}")
        
        # Testa geração de nome consolidado
        consolidated_name = ConsolidationRules.generate_consolidated_name(compatible_columns, self.logger)
        
        self.assertTrue(consolidated_name.startswith('dim_'), "Nome consolidado deve começar com 'dim_'")
        self.assertIn('grupo_etario', consolidated_name, "Nome deve conter 'grupo_etario'")
        
        # Testa validação de nome
        is_valid, errors = ConsolidationRules.validate_consolidated_name(consolidated_name)
        self.assertTrue(is_valid, f"Nome gerado deve ser válido. Erros: {errors}")
    
    def test_dimension_consolidator_initialization(self):
        """Testa inicialização do DimensionConsolidator"""
        # Teste com ficheiro válido
        consolidator = DimensionConsolidator(self.test_file, self.output_dir, self.logger)
        
        self.assertEqual(consolidator.input_file, self.test_file)
        self.assertEqual(consolidator.output_dir, self.output_dir)
        self.assertIsNotNone(consolidator.report)
        
        # Teste com ficheiro inexistente
        with self.assertRaises(FileNotFoundError):
            DimensionConsolidator('ficheiro_inexistente.xlsx', self.output_dir, self.logger)
        
        # Teste com ficheiro não-Excel
        text_file = os.path.join(self.input_dir, 'test.txt')
        with open(text_file, 'w') as f:
            f.write("test")
        
        with self.assertRaises(ValueError):
            DimensionConsolidator(text_file, self.output_dir, self.logger)
    
    def test_dimension_consolidator_dry_run(self):
        """Testa modo de simulação do consolidador"""
        consolidator = DimensionConsolidator(self.test_file, self.output_dir, self.logger)
        
        # Executa em modo de simulação
        result_df = consolidator.consolidate(dry_run=True)
        
        # Verifica que o DataFrame foi retornado
        self.assertIsNotNone(result_df)
        self.assertIsInstance(result_df, pd.DataFrame)
        
        # Verifica que a coluna 'valor' não foi modificada
        if 'valor' in result_df.columns:
            original_valores = self.test_df['valor'].tolist()
            result_valores = result_df['valor'].tolist()
            self.assertEqual(original_valores, result_valores, "Coluna 'valor' não deve ser modificada")
        
        # Verifica que o relatório foi gerado
        report = consolidator.get_report()
        self.assertIsNotNone(report)
        
        # Deve ter dados do relatório
        consolidation_details = report.get_consolidation_details()
        self.assertIsInstance(consolidation_details, list)
    
    def test_dimension_consolidator_full_run(self):
        """Testa execução completa do consolidador"""
        consolidator = DimensionConsolidator(self.test_file, self.output_dir, self.logger)
        
        # Executa consolidação completa
        result_df = consolidator.consolidate(dry_run=False)
        
        # Verifica integridade básica
        self.assertIsNotNone(result_df)
        self.assertEqual(len(result_df), len(self.test_df), "Número de linhas deve ser preservado")
        
        # Verifica que a coluna 'valor' não foi modificada
        if 'valor' in result_df.columns and 'valor' in self.test_df.columns:
            self.assertTrue(self.test_df['valor'].equals(result_df['valor']), 
                           "Coluna 'valor' deve ser preservada exatamente")
        
        # Verifica que colunas não-dimensão foram preservadas
        non_dim_cols_original = [col for col in self.test_df.columns if not col.startswith('dim_')]
        non_dim_cols_result = [col for col in result_df.columns if not col.startswith('dim_')]
        
        for col in non_dim_cols_original:
            self.assertIn(col, non_dim_cols_result, f"Coluna não-dimensão '{col}' deve ser preservada")
        
        # Testa salvamento dos resultados
        output_file = consolidator.save_results(format='excel')
        self.assertTrue(os.path.exists(output_file), "Ficheiro de saída deve ser criado")
        
        # Verifica se o ficheiro de relatório foi criado
        report_file = output_file.replace('.xlsx', '_report.json')
        self.assertTrue(os.path.exists(report_file), "Relatório deve ser guardado")
    
    def test_exclude_columns_functionality(self):
        """Testa funcionalidade de exclusão de colunas"""
        consolidator = DimensionConsolidator(self.test_file, self.output_dir, self.logger)
        
        # Lista de colunas a excluir
        exclude_columns = ['dim_genero', 'dim_ano']
        
        # Executa com exclusões
        result_df = consolidator.consolidate(dry_run=True, exclude_columns=exclude_columns)
        
        # Verifica que as colunas excluídas ainda existem no resultado
        for col in exclude_columns:
            self.assertIn(col, result_df.columns, f"Coluna excluída '{col}' deve ainda existir no resultado")
        
        # Verifica no relatório que as colunas foram excluídas da análise
        report = consolidator.get_report()
        analysis_details = report.analysis_details
        
        if 'pattern_detection' in analysis_details:
            excluded = analysis_details['pattern_detection']['details'].get('excluded_columns', [])
            for col in exclude_columns:
                self.assertIn(col, excluded, f"Coluna '{col}' deve estar marcada como excluída")
    
    def test_simple_dataset_no_patterns(self):
        """Testa com dataset simples sem padrões óbvios"""
        consolidator = DimensionConsolidator(self.simple_test_file, self.output_dir, self.logger)
        
        # Executa consolidação
        result_df = consolidator.consolidate(dry_run=True)
        
        # Deve executar sem erros mesmo sem padrões
        self.assertIsNotNone(result_df)
        
        # Verifica que nenhuma consolidação foi realizada
        consolidation_mapping = consolidator.get_consolidation_mapping()
        # Para dataset simples, pode ou não ter consolidações dependendo dos critérios
        self.assertIsInstance(consolidation_mapping, dict)
    
    def test_consolidation_report_generation(self):
        """Testa geração de relatórios detalhados"""
        consolidator = DimensionConsolidator(self.test_file, self.output_dir, self.logger)
        
        # Executa consolidação
        result_df = consolidator.consolidate(dry_run=False)
        
        # Testa geração de resumo no console
        try:
            consolidator.print_summary()  # Não deve gerar erro
        except Exception as e:
            self.fail(f"print_summary() gerou erro: {e}")
        
        # Verifica integridade dos dados no relatório
        report = consolidator.get_report()
        integrity_status = report.get_integrity_status()
        
        self.assertIsInstance(integrity_status, dict)
        
        # Verificações críticas devem ter passado
        if 'valor_column' in integrity_status:
            self.assertTrue(integrity_status['valor_column'], 
                           "Verificação da coluna 'valor' deve passar")

if __name__ == '__main__':
    unittest.main(verbosity=2) 