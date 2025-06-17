#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import argparse
import json
from datetime import datetime
from typing import Dict, Any, List

from src.utils import setup_logging, ensure_directory_exists
from src.excel_converter import ExcelConverter
from src.data_validator import CustomJSONEncoder

def parse_arguments():
    """
    Analisa os argumentos de linha de comando.
    
    Returns:
        Argumentos analisados
    """
    parser = argparse.ArgumentParser(
        description="Ferramenta de conversão de Excel para CSV/JSON com garantia de integridade de dados",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter
    )
    
    parser.add_argument(
        "--format", "-f",
        choices=["csv", "json"],
        default=None,
        help="Formato de saída (csv ou json)"
    )
    
    parser.add_argument(
        "--input-dir", "-i",
        default="dataset/main",
        help="Diretório de entrada com arquivos Excel"
    )
    
    parser.add_argument(
        "--output-dir", "-o",
        default="result/main",
        help="Diretório de saída para arquivos convertidos"
    )
    
    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        help="Ativa modo detalhado de logs"
    )
    
    parser.add_argument(
        "--file", 
        help="Processa apenas um arquivo específico (opcional)"
    )
    
    parser.add_argument(
        "--sheet",
        help="Nome da planilha a ser processada (padrão: 'dados')"
    )
    
    parser.add_argument(
        "--summary",
        action="store_true",
        help="Gera um resumo do processamento em JSON"
    )
    
    return parser.parse_args()

def prompt_for_format() -> str:
    """
    Solicita ao usuário o formato de saída desejado.
    
    Returns:
        Formato escolhido ('csv' ou 'json')
    """
    while True:
        choice = input("\nEscolha o formato de saída (csv/json): ").strip().lower()
        if choice in ["csv", "json"]:
            return choice
        print("Opção inválida. Por favor, escolha 'csv' ou 'json'.")

def save_summary(results: List[Dict[str, Any]], output_dir: str) -> None:
    """
    Salva um resumo do processamento em um arquivo JSON.
    
    Args:
        results: Lista de resultados do processamento
        output_dir: Diretório de saída
    """
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    summary_path = os.path.join(output_dir, f"conversion_summary_{timestamp}.json")
    
    summary = {
        "timestamp": datetime.now().isoformat(),
        "total_files": len(results),
        "successful": sum(1 for r in results if "error" not in r),
        "failed": sum(1 for r in results if "error" in r),
        "results": results
    }
    
    with open(summary_path, 'w', encoding='utf-8') as f:
        json.dump(summary, f, indent=2, ensure_ascii=False, cls=CustomJSONEncoder)
        
    print(f"\nResumo salvo em: {summary_path}")

def display_welcome_message():
    """Exibe mensagem de boas-vindas"""
    print("\n" + "="*80)
    print("  SISTEMA DE CONVERSÃO DE EXCEL - GARANTIA DE INTEGRIDADE DE DADOS")
    print("  Desenvolvido para o Workflow ETL com Apache Airflow")
    print("="*80)
    print("\nEste sistema converte arquivos Excel para CSV ou JSON preservando 100% da integridade dos dados.")
    print("Os arquivos convertidos serão salvos no diretório de saída especificado.")

def main():
    """Função principal do programa"""
    # Exibe mensagem de boas-vindas
    display_welcome_message()
    
    # Analisa argumentos
    args = parse_arguments()
    
    # Configura o logger
    logger = setup_logging(verbose=args.verbose)
    logger.info("Iniciando processo de conversão de Excel")
    
    # Garante que os diretórios existam
    ensure_directory_exists(args.input_dir)
    ensure_directory_exists(args.output_dir)
    
    # Se o formato não foi especificado, solicita ao usuário
    output_format = args.format
    if output_format is None:
        output_format = prompt_for_format()
        logger.info(f"Formato escolhido pelo usuário: {output_format}")
    
    # Inicializa o conversor
    converter = ExcelConverter(args.input_dir, args.output_dir, logger)
    
    results = []
    try:
        # Processa arquivos
        if args.file:
            # Processa apenas um arquivo específico
            file_path = os.path.join(args.input_dir, args.file)
            if os.path.exists(file_path):
                logger.info(f"Processando arquivo único: {file_path}")
                result = converter.process_excel_file(file_path, output_format, args.sheet)
                results.append(result)
            else:
                logger.error(f"Arquivo não encontrado: {file_path}")
                print(f"ERRO: Arquivo não encontrado: {file_path}")
                sys.exit(1)
        else:
            # Processa todos os arquivos Excel no diretório
            logger.info(f"Processando todos os arquivos no diretório: {args.input_dir}")
            results = converter.process_all_files(output_format)
        
        # Gera resumo se solicitado
        if args.summary or len(results) > 1:
            save_summary(results, args.output_dir)
            
        # Exibe mensagem de conclusão
        successful = sum(1 for r in results if "error" not in r)
        print(f"\nProcessamento concluído: {successful}/{len(results)} arquivos processados com sucesso.")
        print(f"Arquivos convertidos disponíveis em: {os.path.abspath(args.output_dir)}")
        
    except Exception as e:
        logger.error(f"Erro durante a execução: {str(e)}")
        print(f"\nERRO: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main() 
    