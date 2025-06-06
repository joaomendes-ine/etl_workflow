import pandas as pd
import numpy as np
import os
import json
from typing import Dict, List, Any, Tuple, Optional
from src.utils import calculate_dataframe_hash, validate_dataframe_integrity

class CustomJSONEncoder(json.JSONEncoder):
    """
    Encoder JSON personalizado para lidar com tipos não serializáveis nativamente.
    """
    def default(self, obj):
        # Tratamento de tipos numéricos NumPy
        if isinstance(obj, (np.int8, np.int16, np.int32, np.int64, 
                          np.uint8, np.uint16, np.uint32, np.uint64)):
            return int(obj)
        elif isinstance(obj, (np.float16, np.float32, np.float64)):
            return float(obj)
        elif isinstance(obj, np.bool_):
            return bool(obj)
        elif isinstance(obj, np.ndarray):
            return obj.tolist()
        # Tratamento para sets (convertendo para listas)
        elif isinstance(obj, set):
            return list(obj)
        # Tratamento para objetos pandas e datetime
        elif pd.isna(obj):
            return None
        return super(CustomJSONEncoder, self).default(obj)

class DataValidator:
    """
    Classe responsável por validar a integridade dos dados durante o processo de ETL.
    Implementa múltiplas verificações para garantir que nenhum dado seja perdido ou modificado.
    """
    
    def __init__(self, logger):
        """
        Inicializa o validador de dados.
        
        Args:
            logger: Logger configurado
        """
        self.logger = logger
        
    def validate_excel_read(self, file_path: str, df: pd.DataFrame) -> bool:
        """
        Valida se o arquivo Excel foi lido corretamente.
        
        Args:
            file_path: Caminho do arquivo Excel
            df: DataFrame com os dados lidos
            
        Returns:
            True se o arquivo foi lido com sucesso, False caso contrário
        """
        if df is None or df.empty:
            self.logger.error(f"Falha na leitura do arquivo {file_path}: DataFrame vazio ou nulo")
            return False
        
        # Removendo a verificação rígida de colunas específicas para permitir
        # a conversão de qualquer tipo de arquivo Excel
        # As colunas 'indicador' e 'valor' eram exigidas anteriormente, mas isso
        # estava impedindo a conversão de arquivos Excel com estruturas diferentes.
        
        # Agora, apenas verificamos se o DataFrame tem pelo menos uma coluna
        if df.columns.size == 0:
            self.logger.error(f"O arquivo {file_path} não contém colunas")
            return False
        
        # Verifica se há pelo menos uma linha de dados
        if df.shape[0] == 0:
            self.logger.warning(f"O arquivo {file_path} não contém linhas de dados")
            # Continuamos mesmo sem dados, já que pode ser um arquivo vazio válido
        
        self.logger.info(f"Validação de leitura bem-sucedida para {file_path}: {df.shape[0]} linhas, {df.shape[1]} colunas")
        return True
    
    def validate_conversion(self, original_df: pd.DataFrame, 
                           output_path: str, 
                           format_type: str) -> Tuple[bool, Dict[str, Any]]:
        """
        Valida se a conversão foi realizada corretamente, comparando com os dados originais.
        
        Args:
            original_df: DataFrame original
            output_path: Caminho do arquivo de saída
            format_type: Tipo de formato do arquivo convertido (csv ou json)
            
        Returns:
            Tupla com booleano (True se íntegro) e dicionário com detalhes de validação
        """
        try:
            # Verifica se o formato é JSON para aplicar regras específicas
            is_json = format_type.lower() == 'json'
            
            # Carrega o arquivo convertido de volta para um DataFrame para comparação
            if format_type.lower() == 'csv':
                converted_df = pd.read_csv(output_path, low_memory=False)
            elif format_type.lower() == 'json':
                converted_df = pd.read_json(output_path)
            else:
                self.logger.error(f"Formato não suportado para validação: {format_type}")
                return False, {"error": f"Formato não suportado: {format_type}"}
            
            # Valida a integridade dos dados com flag específica para JSON
            is_valid, validation_details = validate_dataframe_integrity(
                original_df, converted_df, is_json=is_json
            )
            
            # Convertemos os resultados para tipos Python nativos para evitar problemas de serialização
            serializable_details = self._convert_to_serializable(validation_details)
            
            if is_valid:
                self.logger.info(f"Validação de conversão bem-sucedida para {output_path}")
                self.logger.debug("Detalhes da validação: " + json.dumps(serializable_details, indent=2))
            else:
                self.logger.error(f"Falha na validação de conversão para {output_path}")
                self.logger.error("Detalhes da falha: " + json.dumps(serializable_details, indent=2))
                
            return is_valid, serializable_details
            
        except Exception as e:
            self.logger.error(f"Erro durante a validação de conversão: {str(e)}")
            return False, {"error": str(e)}
    
    def _convert_to_serializable(self, data):
        """
        Converte valores NumPy e pandas para tipos Python nativos para permitir serialização JSON.
        
        Args:
            data: Dicionário ou valor a ser convertido
            
        Returns:
            Dados convertidos para tipos serializáveis
        """
        if isinstance(data, dict):
            return {k: self._convert_to_serializable(v) for k, v in data.items()}
        elif isinstance(data, list):
            return [self._convert_to_serializable(v) for v in data]
        elif isinstance(data, set):
            return list(data)
        elif isinstance(data, (np.int8, np.int16, np.int32, np.int64, np.uint8, np.uint16, np.uint32, np.uint64)):
            return int(data)
        elif isinstance(data, (np.float16, np.float32, np.float64)):
            return float(data)
        elif isinstance(data, np.bool_):
            return bool(data)
        elif isinstance(data, np.ndarray):
            return data.tolist()
        elif pd.isna(data):
            return None
        return data
    
    def check_numerical_precision(self, original_df: pd.DataFrame, 
                                 converted_df: pd.DataFrame) -> Dict[str, Any]:
        """
        Verifica a precisão numérica em colunas numéricas.
        
        Args:
            original_df: DataFrame original
            converted_df: DataFrame convertido
            
        Returns:
            Dicionário com resultados da verificação
        """
        results = {}
        
        # Verifica apenas colunas numéricas
        numeric_cols = original_df.select_dtypes(include=['number']).columns
        
        for col in numeric_cols:
            if col in converted_df.columns:
                # Calcula a diferença absoluta máxima
                max_diff = np.max(np.abs(original_df[col] - converted_df[col]))
                results[col] = {
                    "max_absolute_difference": float(max_diff),
                    "is_exact": float(max_diff) == 0
                }
            else:
                results[col] = {"error": "Coluna não encontrada no DataFrame convertido"}
                
        return results
    
    def check_special_characters(self, original_df: pd.DataFrame, 
                               converted_df: pd.DataFrame) -> Dict[str, List[str]]:
        """
        Verifica se caracteres especiais foram preservados em colunas de texto.
        
        Args:
            original_df: DataFrame original
            converted_df: DataFrame convertido
            
        Returns:
            Dicionário com resultados da verificação
        """
        results = {}
        
        # Verifica apenas colunas de texto
        text_cols = original_df.select_dtypes(include=['object']).columns
        
        for col in text_cols:
            if col in converted_df.columns:
                mismatched_rows = []
                
                # Compara os valores de texto
                for i, (orig_val, conv_val) in enumerate(zip(original_df[col], converted_df[col])):
                    if orig_val != conv_val and not (pd.isna(orig_val) and pd.isna(conv_val)):
                        mismatched_rows.append(i)
                
                results[col] = mismatched_rows
                
        return results
    
    def generate_validation_report(self, file_path: str, 
                                  output_path: str, 
                                  validation_results: Dict[str, Any]) -> Dict[str, Any]:
        """
        Gera um relatório de validação completo.
        
        Args:
            file_path: Caminho do arquivo original
            output_path: Caminho do arquivo convertido
            validation_results: Resultados da validação
            
        Returns:
            Dicionário com o relatório de validação
        """
        report = {
            "file_info": {
                "original_file": file_path,
                "converted_file": output_path,
                "validation_timestamp": pd.Timestamp.now().isoformat()
            },
            "validation_results": validation_results,
            "status": "Aprovado" if validation_results.get("is_valid", False) else "Falhou"
        }
        
        return report 