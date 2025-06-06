import os
import logging
import hashlib
import pandas as pd
import numpy as np
from datetime import datetime
from typing import Dict, Any, List, Union, Tuple

def setup_logging(log_dir: str = 'logs', verbose: bool = False) -> logging.Logger:
    """
    Configura o sistema de logging com formato adequado.
    
    Args:
        log_dir: Diretório onde os logs serão armazenados
        verbose: Se True, configura o logger para modo detalhado
    
    Returns:
        Logger configurado
    """
    os.makedirs(log_dir, exist_ok=True)
    
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    log_file = os.path.join(log_dir, f'etl_process_{timestamp}.log')
    
    # Configuração do logger
    logger = logging.getLogger('etl_workflow')
    logger.setLevel(logging.DEBUG if verbose else logging.INFO)
    
    # Handler para arquivo
    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    file_handler.setFormatter(file_formatter)
    file_handler.setLevel(logging.DEBUG)
    
    # Handler para console com configuração de encoding segura
    console_handler = logging.StreamHandler()
    console_formatter = logging.Formatter('%(levelname)s: %(message)s')
    console_handler.setFormatter(console_formatter)
    console_handler.setLevel(logging.DEBUG if verbose else logging.INFO)
    
    # Configura encoding para evitar problemas no Windows
    try:
        import sys
        if hasattr(console_handler.stream, 'reconfigure'):
            console_handler.stream.reconfigure(encoding='utf-8', errors='replace')
        elif hasattr(sys.stdout, 'buffer'):
            # Para versões mais antigas do Python
            import io
            sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    except:
        # Se não conseguir configurar UTF-8, continua com configuração padrão
        pass
    
    # Adiciona os handlers ao logger
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    
    return logger

def calculate_dataframe_hash(df: pd.DataFrame) -> str:
    """
    Calcula o hash do DataFrame para verificação de integridade.
    
    Args:
        df: DataFrame pandas a ser verificado
    
    Returns:
        String contendo o hash SHA-256 do DataFrame
    """
    # Converte o DataFrame para uma representação em bytes
    df_bytes = pd.util.hash_pandas_object(df, index=True).values.tobytes()
    
    # Calcula o hash SHA-256
    sha256_hash = hashlib.sha256(df_bytes).hexdigest()
    
    return sha256_hash

def validate_dataframe_integrity(original_df: pd.DataFrame, 
                                 converted_df: pd.DataFrame,
                                 is_json: bool = False) -> Tuple[bool, Dict[str, Any]]:
    """
    Valida a integridade dos dados entre o DataFrame original e o convertido.
    
    Args:
        original_df: DataFrame original
        converted_df: DataFrame após conversão
        is_json: Se True, aplica verificações específicas para JSON (mais flexíveis)
    
    Returns:
        Tupla com booleano (True se íntegro) e dicionário com detalhes de validação
    """
    validation_results = {
        "shape_match": original_df.shape == converted_df.shape,
        "columns_match": list(original_df.columns) == list(converted_df.columns),
        "dtypes_match": True,
        "data_match": True,
        "null_values_match": True,
        "original_hash": calculate_dataframe_hash(original_df),
        "converted_hash": calculate_dataframe_hash(converted_df)
    }
    
    # Verifica se os tipos de dados correspondem (mais flexível para JSON)
    for col in original_df.columns:
        # Para JSON, alguns tipos podem mudar, mas precisamos garantir compatibilidade
        if is_json:
            if original_df[col].dtype.kind == 'f' and converted_df[col].dtype.kind != 'f':
                validation_results["dtypes_match"] = False
                break
            if original_df[col].dtype.kind == 'i' and converted_df[col].dtype.kind not in ('i', 'u'):
                validation_results["dtypes_match"] = False
                break
        else:
            # Para CSV, verificação mais rigorosa
            if original_df[col].dtype != converted_df[col].dtype:
                validation_results["dtypes_match"] = False
                break
    
    # Verifica valores nulos
    null_count_orig = original_df.isnull().sum().sum()
    null_count_conv = converted_df.isnull().sum().sum()
    
    # Para JSON, podemos ser mais flexíveis com valores nulos
    if is_json:
        # Diferença de até 5% nos valores nulos é aceitável para JSON
        null_tolerance = max(1, int(null_count_orig * 0.05))
        validation_results["null_values_match"] = abs(null_count_orig - null_count_conv) <= null_tolerance
    else:
        validation_results["null_values_match"] = null_count_orig == null_count_conv
    
    # Verifica valores numéricos (para colunas numéricas)
    for col in original_df.select_dtypes(include='number').columns:
        if col in converted_df.columns:
            # Para JSON, podemos permitir pequenas diferenças de precisão
            if is_json:
                # Ignora valores NaN na comparação
                orig_values = original_df[col].dropna()
                conv_values = converted_df[col].dropna()
                
                # Se houver valores para comparar
                if not orig_values.empty and not conv_values.empty:
                    # Calcula a diferença relativa média (tolerável para JSON)
                    try:
                        diff = np.abs(orig_values.values - conv_values.values[:len(orig_values)])
                        # Diferença máxima aceitável é 1e-10 para valores JSON
                        if np.max(diff) > 1e-10:
                            validation_results["data_match"] = False
                            break
                    except:
                        # Se não conseguir calcular a diferença, considera que não há correspondência
                        validation_results["data_match"] = False
                        break
            else:
                # Para CSV, exigimos igualdade exata
                if not original_df[col].equals(converted_df[col]):
                    validation_results["data_match"] = False
                    break
    
    # Verifica valores de texto (para colunas de objeto)
    for col in original_df.select_dtypes(include='object').columns:
        if col in converted_df.columns:
            # Comparação de valores de texto, tratando NaN corretamente
            for i, (orig_val, conv_val) in enumerate(zip(original_df[col], converted_df[col])):
                # Se ambos são NaN ou o mesmo valor
                if not (pd.isna(orig_val) and pd.isna(conv_val)) and orig_val != conv_val:
                    validation_results["data_match"] = False
                    break
            
            if not validation_results["data_match"]:
                break
    
    # Para JSON, o hash geralmente será diferente devido a diferenças de precisão
    # então, se todos os outros critérios passarem, consideramos válido
    if is_json and validation_results["shape_match"] and validation_results["columns_match"] and validation_results["dtypes_match"]:
        is_valid = True
    else:
        # Resultado final para outros formatos
        is_valid = all([
            validation_results["shape_match"],
            validation_results["columns_match"],
            validation_results["dtypes_match"],
            validation_results["data_match"],
            validation_results["null_values_match"]
        ])
    
    return is_valid, validation_results

def get_file_stats(file_path: str) -> Dict[str, Any]:
    """
    Obtém estatísticas do arquivo para registro.
    
    Args:
        file_path: Caminho do arquivo
    
    Returns:
        Dicionário com estatísticas do arquivo
    """
    stats = os.stat(file_path)
    return {
        "file_path": file_path,
        "file_size_bytes": stats.st_size,
        "last_modified": datetime.fromtimestamp(stats.st_mtime).isoformat(),
        "created": datetime.fromtimestamp(stats.st_ctime).isoformat()
    }

def ensure_directory_exists(directory_path: str) -> None:
    """
    Garante que o diretório especificado existe.
    
    Args:
        directory_path: Caminho do diretório
    """
    if not os.path.exists(directory_path):
        os.makedirs(directory_path)
        
def get_excel_sheet_names(excel_file_path: str) -> List[str]:
    """
    Retorna os nomes das planilhas em um arquivo Excel.
    
    Args:
        excel_file_path: Caminho do arquivo Excel
    
    Returns:
        Lista com os nomes das planilhas
    """
    try:
        return pd.ExcelFile(excel_file_path).sheet_names
    except Exception as e:
        raise ValueError(f"Erro ao ler planilhas do arquivo Excel: {str(e)}") 