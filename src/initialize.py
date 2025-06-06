#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import logging
from src.utils import setup_logging, ensure_directory_exists

def initialize_project_structure():
    """
    Inicializa a estrutura de pastas do projeto.
    
    Cria as seguintes pastas se elas não existirem:
    - dataset/validation/1 a 96/quadros
    - dataset/validation/1 a 96/series
    - dataset/main
    - result/validation/1 a 96/quadros
    - result/validation/1 a 96/series
    - result/main
    - logs
    """
    logger = setup_logging(verbose=True)
    logger.info("Inicializando estrutura de pastas do projeto")
    
    # Lista para armazenar todas as pastas criadas
    folders = []
    
    # Cria as pastas de 1 a 96 com suas subpastas
    for folder_num in range(1, 97):
        # Pastas de entrada para consolidação (opções 1 e 2)
        input_quadros = os.path.join("dataset", "validation", str(folder_num), "quadros")
        input_series = os.path.join("dataset", "validation", str(folder_num), "series")
        
        # Pastas de saída para consolidação (opções 1 e 2)
        output_quadros = os.path.join("result", "validation", str(folder_num), "quadros")
        output_series = os.path.join("result", "validation", str(folder_num), "series")
        
        folders.extend([input_quadros, input_series, output_quadros, output_series])
    
    # Adiciona as outras pastas principais
    folders.extend([
        # Pasta de entrada para conversão (opções 3 e 4)
        os.path.join("dataset", "main"),
        
        # Pasta de saída para conversão (opções 3 e 4)
        os.path.join("result", "main"),
        
        # Pasta para logs
        os.path.join("logs")
    ])
    
    # Cria cada pasta
    created_count = 0
    for folder in folders:
        ensure_directory_exists(folder)
        created_count += 1
        # Loga apenas algumas pastas para não poluir o log
        if created_count % 20 == 0 or folder_num <= 5 or folder_num >= 92:
            logger.info(f"Pasta criada/verificada: {folder}")
    
    logger.info(f"Estrutura de pastas inicializada com sucesso: {len(folders)} pastas criadas/verificadas")
    return folders

if __name__ == "__main__":
    """
    Executa a inicialização da estrutura de pastas quando o script é executado diretamente.
    Uso: python -m src.initialize
    """
    folders = initialize_project_structure()
    
    print("\n=== Estrutura de pastas do projeto ===")
    print(f"Foram criadas/verificadas {len(folders)} pastas, incluindo:")
    
    # Exibe apenas algumas pastas de exemplo para não sobrecarregar o console
    sample_folders = [
        os.path.join("dataset", "validation", "1", "quadros"),
        os.path.join("dataset", "validation", "1", "series"),
        os.path.join("dataset", "validation", "55", "quadros"),
        os.path.join("dataset", "validation", "55", "series"),
        os.path.join("dataset", "validation", "96", "quadros"),
        os.path.join("dataset", "validation", "96", "series"),
        os.path.join("result", "validation", "1", "quadros"),
        os.path.join("result", "validation", "96", "series"),
        os.path.join("dataset", "main"),
        os.path.join("result", "main"),
        os.path.join("logs")
    ]
    
    for folder in sample_folders:
        print(f" - {folder}")
    
    print(f" - ... e mais {len(folders) - len(sample_folders)} pastas\n")
    
    print("O projeto está pronto para uso!")
    print("Execute 'python main.py' para iniciar o programa principal.") 