import os
import json
import pandas as pd
from typing import Dict, List, Any, Tuple, Optional, Union
from datetime import datetime
from src.utils import ensure_directory_exists, get_file_stats, get_excel_sheet_names
from src.data_validator import DataValidator, CustomJSONEncoder

class ExcelConverter:
    """
    Classe responsável pela conversão de arquivos Excel para CSV ou JSON,
    garantindo a integridade absoluta dos dados durante o processo.
    """
    
    def __init__(self, input_dir: str, output_dir: str, logger):
        """
        Inicializa o conversor de Excel.
        
        Args:
            input_dir: Diretório de entrada dos arquivos Excel
            output_dir: Diretório de saída para os arquivos convertidos
            logger: Logger configurado
        """
        self.input_dir = input_dir
        self.output_dir = output_dir
        self.logger = logger
        self.validator = DataValidator(logger)
        
        # Garante que os diretórios existam
        ensure_directory_exists(self.input_dir)
        ensure_directory_exists(self.output_dir)
        
    def list_excel_files(self) -> List[str]:
        """
        Lista todos os arquivos Excel no diretório de entrada.
        
        Returns:
            Lista de caminhos de arquivos Excel
        """
        excel_files = []
        
        for file in os.listdir(self.input_dir):
            if file.endswith(('.xlsx', '.xls')):
                excel_files.append(os.path.join(self.input_dir, file))
                
        return excel_files
    
    def read_excel_file(self, file_path: str, sheet_name: Optional[str] = None) -> pd.DataFrame:
        """
        Lê um arquivo Excel preservando todos os tipos de dados.
        
        Args:
            file_path: Caminho do arquivo Excel
            sheet_name: Nome da planilha a ser lida (se None, tenta ler 'dados')
            
        Returns:
            DataFrame pandas com os dados do arquivo Excel
        """
        try:
            self.logger.info(f"Lendo arquivo Excel: {file_path}")
            
            # Se sheet_name não foi especificado, tenta usar 'dados' ou a primeira planilha
            if sheet_name is None:
                try:
                    sheets = get_excel_sheet_names(file_path)
                    if 'dados' in sheets:
                        sheet_name = 'dados'
                    else:
                        sheet_name = sheets[0]
                        self.logger.info(f"Planilha 'dados' não encontrada, usando a primeira planilha: {sheet_name}")
                except Exception as e:
                    self.logger.error(f"Erro ao obter nomes das planilhas: {str(e)}")
                    sheet_name = 0  # Usa a primeira planilha como fallback
                    self.logger.info(f"Usando índice 0 como fallback para a planilha")
            
            # Lê o arquivo Excel com configurações para preservar todos os dados
            try:
                df = pd.read_excel(
                    file_path,
                    sheet_name=sheet_name,
                    keep_default_na=True,
                    na_values=[''],  # Define quais valores serão considerados como NaN
                    engine='openpyxl'  # Usa o engine openpyxl para melhor compatibilidade
                )
            except Exception as e:
                self.logger.warning(f"Erro ao ler planilha específica '{sheet_name}': {str(e)}")
                self.logger.info(f"Tentando ler a primeira planilha como fallback")
                # Tenta ler a primeira planilha como fallback
                df = pd.read_excel(
                    file_path,
                    sheet_name=0,
                    keep_default_na=True,
                    na_values=[''],
                    engine='openpyxl'
                )
            
            # Valida a leitura do arquivo
            if not self.validator.validate_excel_read(file_path, df):
                raise ValueError(f"Falha na validação da leitura do arquivo {file_path}")
                
            # Exibe informações sobre o arquivo lido
            self.logger.info(f"Arquivo lido com sucesso: {df.shape[0]} linhas, {df.shape[1]} colunas")
            self.logger.debug(f"Colunas: {df.columns.tolist()}")
            self.logger.debug(f"Tipos de dados: {df.dtypes}")
            
            return df
            
        except Exception as e:
            self.logger.error(f"Erro ao ler o arquivo Excel {file_path}: {str(e)}")
            raise
    
    def convert_to_csv(self, df: pd.DataFrame, output_path: str) -> str:
        """
        Converte um DataFrame para CSV preservando todos os dados.
        
        Args:
            df: DataFrame a ser convertido
            output_path: Caminho do arquivo de saída
            
        Returns:
            Caminho do arquivo CSV gerado
        """
        try:
            self.logger.info(f"Convertendo para CSV: {output_path}")
            
            # Configura opções para preservar a integridade dos dados
            df.to_csv(
                output_path,
                index=False,  # Não inclui o índice
                na_rep='',  # Representa valores NaN como string vazia
                float_format='%.15g',  # Preserva a precisão de valores flutuantes
                encoding='utf-8'  # Usa UTF-8 para suportar caracteres especiais
            )
            
            self.logger.info(f"Arquivo CSV gerado: {output_path}")
            return output_path
            
        except Exception as e:
            self.logger.error(f"Erro ao converter para CSV: {str(e)}")
            raise
    
    def convert_to_json(self, df: pd.DataFrame, output_path: str) -> str:
        """
        Converte um DataFrame para JSON preservando todos os dados.
        
        Args:
            df: DataFrame a ser convertido
            output_path: Caminho do arquivo de saída
            
        Returns:
            Caminho do arquivo JSON gerado
        """
        try:
            self.logger.info(f"Convertendo para JSON: {output_path}")
            
            # Configura opções para preservar a integridade dos dados
            df.to_json(
                output_path,
                orient='records',  # Formato de registros (lista de objetos)
                date_format='iso',  # Formato ISO para datas
                double_precision=15,  # Preserva precisão de valores flutuantes
                force_ascii=False,  # Permite caracteres não-ASCII
                indent=2  # Formata JSON para melhor legibilidade
            )
            
            self.logger.info(f"Arquivo JSON gerado: {output_path}")
            return output_path
            
        except Exception as e:
            self.logger.error(f"Erro ao converter para JSON: {str(e)}")
            raise
    
    def process_excel_file(self, file_path: str, output_format: str, 
                          sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """
        Processa um arquivo Excel, convertendo-o para o formato especificado.
        
        Args:
            file_path: Caminho do arquivo Excel
            output_format: Formato de saída ('csv' ou 'json')
            sheet_name: Nome da planilha a ser processada
            
        Returns:
            Dicionário com informações sobre o processamento
        """
        try:
            # Registra o início do processamento
            start_time = datetime.now()
            self.logger.info(f"Iniciando processamento do arquivo: {file_path}")
            self.logger.info(f"Formato de saída: {output_format}")
            
            # Verificação básica do arquivo
            if not os.path.exists(file_path):
                raise FileNotFoundError(f"Arquivo não encontrado: {file_path}")
            
            if not os.path.isfile(file_path):
                raise ValueError(f"O caminho não é um arquivo: {file_path}")
            
            # Obtém estatísticas do arquivo original
            try:
                original_stats = get_file_stats(file_path)
                self.logger.debug(f"Estatísticas do arquivo original: tamanho={original_stats['file_size_bytes']} bytes")
            except Exception as stats_err:
                self.logger.warning(f"Erro ao obter estatísticas do arquivo: {str(stats_err)}")
                # Continuamos mesmo sem estatísticas
                original_stats = {"error": str(stats_err)}
            
            # Lê o arquivo Excel
            try:
                df = self.read_excel_file(file_path, sheet_name)
                self.logger.info(f"Arquivo lido com sucesso: {df.shape[0]} linhas x {df.shape[1]} colunas")
            except Exception as read_err:
                self.logger.error(f"Erro ao ler arquivo Excel: {str(read_err)}")
                raise ValueError(f"Falha na leitura do arquivo: {str(read_err)}")
            
            # Gera o nome do arquivo de saída
            file_name = os.path.splitext(os.path.basename(file_path))[0]
            output_path = os.path.join(
                self.output_dir, 
                f"{file_name}.{output_format.lower()}"
            )
            self.logger.info(f"Arquivo de saída será: {output_path}")
            
            # Converte para o formato especificado
            try:
                if output_format.lower() == 'csv':
                    self.logger.info(f"Convertendo para CSV: {output_path}")
                    output_file = self.convert_to_csv(df, output_path)
                elif output_format.lower() == 'json':
                    self.logger.info(f"Convertendo para JSON: {output_path}")
                    output_file = self.convert_to_json(df, output_path)
                else:
                    raise ValueError(f"Formato de saída não suportado: {output_format}")
                    
                self.logger.info(f"Conversão concluída com sucesso: {output_file}")
            except Exception as conv_err:
                self.logger.error(f"Erro durante a conversão: {str(conv_err)}")
                raise ValueError(f"Falha na conversão: {str(conv_err)}")
            
            # Valida a conversão
            try:
                self.logger.info(f"Validando a conversão: {output_file}")
                is_valid, validation_details = self.validator.validate_conversion(
                    df, output_file, output_format.lower()
                )
                self.logger.info(f"Validação concluída: {'aprovado' if is_valid else 'falhou'}")
            except Exception as val_err:
                self.logger.warning(f"Erro durante a validação: {str(val_err)}")
                # Continuamos mesmo com falha na validação, mas registramos
                is_valid = False
                validation_details = {"error": str(val_err)}
            
            # Obtém estatísticas do arquivo de saída
            try:
                output_stats = get_file_stats(output_file)
                self.logger.debug(f"Estatísticas do arquivo de saída: tamanho={output_stats['file_size_bytes']} bytes")
            except Exception as out_stats_err:
                self.logger.warning(f"Erro ao obter estatísticas do arquivo de saída: {str(out_stats_err)}")
                # Continuamos mesmo sem estatísticas
                output_stats = {"error": str(out_stats_err)}
            
            # Registra o tempo de processamento
            end_time = datetime.now()
            processing_time = (end_time - start_time).total_seconds()
            self.logger.info(f"Tempo de processamento: {processing_time:.2f} segundos")
            
            # Compila o resultado
            result = {
                "file_path": file_path,
                "output_file": output_file,
                "format": output_format.lower(),
                "processing_time_seconds": processing_time,
                "rows_processed": df.shape[0],
                "columns_processed": df.shape[1],
                "original_file_stats": original_stats,
                "output_file_stats": output_stats,
                "validation": {
                    "is_valid": is_valid
                }
            }
            
            if is_valid:
                self.logger.info(f"Processamento bem-sucedido: {file_path} -> {output_file}")
            else:
                self.logger.warning(f"Validação falhou após conversão: {file_path} -> {output_file}")
                
            return result
            
        except Exception as e:
            self.logger.error(f"Erro no processamento do arquivo {file_path}: {str(e)}", exc_info=True)
            return {
                "file_path": file_path,
                "error": str(e),
                "timestamp": datetime.now().isoformat()
            }
    
    def process_all_files(self, output_format: str) -> List[Dict[str, Any]]:
        """
        Processa todos os arquivos Excel no diretório de entrada.
        
        Args:
            output_format: Formato de saída ('csv' ou 'json')
            
        Returns:
            Lista de resultados de processamento
        """
        results = []
        files = self.list_excel_files()
        
        self.logger.info(f"Encontrados {len(files)} arquivos Excel para processamento")
        self.logger.debug(f"Lista de arquivos a processar: {[os.path.basename(f) for f in files]}")
        
        if not files:
            self.logger.warning("Nenhum arquivo Excel encontrado no diretório de entrada")
            return results
        
        for file_path in files:
            self.logger.info(f"Iniciando processamento do arquivo: {os.path.basename(file_path)}")
            try:
                result = self.process_excel_file(file_path, output_format)
                results.append(result)
                if "error" in result:
                    self.logger.error(f"Erro ao processar {os.path.basename(file_path)}: {result['error']}")
                else:
                    self.logger.info(f"Arquivo {os.path.basename(file_path)} processado com sucesso para {output_format}")
            except Exception as e:
                self.logger.error(f"Exceção não tratada ao processar {os.path.basename(file_path)}: {str(e)}", exc_info=True)
                results.append({
                    "file_path": file_path,
                    "error": f"Exceção não tratada: {str(e)}",
                    "timestamp": datetime.now().isoformat()
                })
            
        # Resumo do processamento
        successful = sum(1 for r in results if "error" not in r)
        self.logger.info(f"Processamento concluído: {successful}/{len(results)} arquivos processados com sucesso")
        
        # Detalhes dos resultados
        for idx, result in enumerate(results):
            if "error" not in result:
                self.logger.debug(f"[{idx+1}] Sucesso: {os.path.basename(result['file_path'])} -> {os.path.basename(result['output_file'])}")
            else:
                self.logger.debug(f"[{idx+1}] Falha: {os.path.basename(result['file_path'])} - {result['error']}")
        
        return results 