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

from src.excel_converter import ExcelConverter
from src.data_validator import DataValidator
from src.utils import calculate_dataframe_hash, validate_dataframe_integrity

class TestExcelConversion(unittest.TestCase):
    """Testes para o conversor de Excel"""
    
    @classmethod
    def setUpClass(cls):
        """Configuração executada uma vez antes de todos os testes"""
        # Configura o logger
        cls.logger = logging.getLogger('test_logger')
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
        
        # Verifica se o arquivo de teste existe na pasta principal ou na pasta de validação
        test_file_paths = [
            os.path.abspath(os.path.join('dataset', 'main', '55_BD.xlsx')),
            os.path.abspath(os.path.join('dataset', '55_BD.xlsx')),
            os.path.abspath(os.path.join('dataset', 'validation', '55', 'quadros', 'Q1.xlsx')),
            os.path.abspath(os.path.join('dataset', 'validation', '55', 'series', 'I1.xlsx'))
        ]
        
        # Busca pelo primeiro arquivo de teste que existe
        cls.test_file = None
        for test_path in test_file_paths:
            if os.path.exists(test_path):
                cls.test_file = test_path
                break
                
        if cls.test_file is None:
            # Se nenhum arquivo existir, cria um arquivo de teste simples para executar os testes
            cls.logger.warning("Nenhum arquivo de teste encontrado. Criando arquivo de teste temporário.")
            
            # Cria um DataFrame simples para o teste
            test_df = pd.DataFrame({
                'indicador': ['Teste1', 'Teste2', 'Teste3'],
                'valor': [10.5, 20.3, 30.7]
            })
            
            # Salva o DataFrame como Excel em um arquivo temporário
            cls.test_file = os.path.join(cls.temp_dir, '55_BD.xlsx')
            with pd.ExcelWriter(cls.test_file, engine='openpyxl') as writer:
                test_df.to_excel(writer, sheet_name='dados', index=False)
                
            cls.logger.info(f"Arquivo de teste temporário criado: {cls.test_file}")
            
        # Copia o arquivo para o diretório temporário
        shutil.copy(cls.test_file, cls.input_dir)
        cls.temp_test_file = os.path.join(cls.input_dir, os.path.basename(cls.test_file))
        
        # Cria instâncias para teste
        cls.converter = ExcelConverter(cls.input_dir, cls.output_dir, cls.logger)
        cls.validator = DataValidator(cls.logger)
        
        # Lê o arquivo original para comparação (tenta primeiro a planilha 'dados', se não existir usa a primeira planilha)
        try:
            cls.original_df = pd.read_excel(cls.temp_test_file, sheet_name='dados')
        except:
            # Se a planilha 'dados' não existir, usa a primeira planilha disponível
            cls.original_df = pd.read_excel(cls.temp_test_file, sheet_name=0)
    
    @classmethod
    def tearDownClass(cls):
        """Limpeza executada uma vez após todos os testes"""
        # Remove os diretórios temporários
        shutil.rmtree(cls.temp_dir)
    
    def test_csv_conversion(self):
        """Testa a conversão para CSV"""
        # Processa o arquivo para CSV
        result = self.converter.process_excel_file(self.temp_test_file, 'csv')
        
        # Verifica se o processamento foi bem-sucedido
        self.assertNotIn('error', result, "Erro durante a conversão para CSV")
        self.assertTrue(result['validation']['is_valid'], "Validação da conversão para CSV falhou")
        
        # Verifica se o arquivo foi criado
        output_file = result['output_file']
        self.assertTrue(os.path.exists(output_file), f"Arquivo de saída não encontrado: {output_file}")
        
        # Verifica se a estrutura do DataFrame foi preservada
        csv_df = pd.read_csv(output_file, low_memory=False)
        self.assertEqual(self.original_df.shape[0], csv_df.shape[0], "Número de linhas não corresponde")
        self.assertEqual(self.original_df.shape[1], csv_df.shape[1], "Número de colunas não corresponde")
        
        # Verifica os valores de algumas colunas (se existirem)
        if 'valor' in self.original_df.columns and 'valor' in csv_df.columns:
            # Verifica se a coluna 'valor' manteve os valores originais (amostragem)
            # Usa uma amostra para tornar o teste mais rápido
            sample_rows = min(100, self.original_df.shape[0])
            for i in range(sample_rows):
                original_value = self.original_df['valor'].iloc[i]
                csv_value = csv_df['valor'].iloc[i]
                self.assertAlmostEqual(
                    original_value, 
                    csv_value, 
                    places=10, 
                    msg=f"Valor na linha {i} não corresponde: original={original_value}, csv={csv_value}"
                )
    
    def test_json_conversion(self):
        """Testa a conversão para JSON"""
        # Processa o arquivo para JSON
        result = self.converter.process_excel_file(self.temp_test_file, 'json')
        
        # Verifica se o processamento foi bem-sucedido
        self.assertNotIn('error', result, "Erro durante a conversão para JSON")
        self.assertTrue(result['validation']['is_valid'], "Validação da conversão para JSON falhou")
        
        # Verifica se o arquivo foi criado
        output_file = result['output_file']
        self.assertTrue(os.path.exists(output_file), f"Arquivo de saída não encontrado: {output_file}")
        
        # Carrega o arquivo JSON convertido
        json_df = pd.read_json(output_file)
        
        # Verifica o número de linhas e colunas
        self.assertEqual(self.original_df.shape[0], json_df.shape[0], "Número de linhas não corresponde")
        self.assertEqual(self.original_df.shape[1], json_df.shape[1], "Número de colunas não corresponde")
        
        # Verifica os valores de algumas colunas (se existirem)
        if 'valor' in self.original_df.columns and 'valor' in json_df.columns:
            # Verifica se a coluna 'valor' manteve os valores originais (amostragem)
            # Usa uma amostra para tornar o teste mais rápido
            sample_rows = min(100, self.original_df.shape[0])
            for i in range(sample_rows):
                original_value = self.original_df['valor'].iloc[i]
                json_value = json_df['valor'].iloc[i]
                self.assertAlmostEqual(
                    original_value, 
                    json_value, 
                    places=10, 
                    msg=f"Valor na linha {i} não corresponde: original={original_value}, json={json_value}"
                )
            
    def test_data_integrity(self):
        """Testa a função de validação de integridade de dados"""
        # Cria uma cópia do DataFrame original
        df_copy = self.original_df.copy()
        
        # Verifica se a integridade é mantida quando não há alterações
        is_valid, details = validate_dataframe_integrity(self.original_df, df_copy)
        self.assertTrue(is_valid, "A validação de integridade falhou para DataFrames idênticos")
        
        # Modifica um valor e verifica se a integridade é quebrada
        df_modified = df_copy.copy()
        
        # Escolhe uma coluna para modificar (preferencialmente 'valor' se existir)
        if 'valor' in df_modified.columns:
            col_to_modify = 'valor'
        else:
            # Usa a primeira coluna numérica disponível
            numeric_cols = df_modified.select_dtypes(include=['number']).columns
            if len(numeric_cols) > 0:
                col_to_modify = numeric_cols[0]
            else:
                # Se não houver colunas numéricas, usa a primeira coluna
                col_to_modify = df_modified.columns[0]
        
        if len(df_modified) > 0:
            df_modified.loc[0, col_to_modify] = 999.99  # Modifica um valor
            is_valid, details = validate_dataframe_integrity(self.original_df, df_modified)
            self.assertFalse(is_valid, "A validação de integridade não detectou a modificação")

if __name__ == '__main__':
    unittest.main() 