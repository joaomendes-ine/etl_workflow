"""
Módulo de Comparação Inteligente de Dados Excel - Versão Avançada
Compara ficheiros Excel com estruturas de tabela cruzada (crosstab/pivot)
Suporta formatação visual, células mescladas, valores apresentados e relatórios destacados.
"""

import os
import pandas as pd
import numpy as np
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter, range_boundaries
from openpyxl.formatting.rule import FormulaRule
import re
from typing import Dict, List, Tuple, Any, Optional, Set
from datetime import datetime
import logging
from difflib import SequenceMatcher
from copy import copy
from src.utils import ensure_directory_exists


class DataComparator:
    """
    Classe principal para comparação inteligente de ficheiros Excel.
    Especializada em estruturas de tabela cruzada com normalização avançada.
    """
    
    def __init__(self, logger: logging.Logger):
        """
        Inicializa o comparador de dados.
        
        Args:
            logger: Logger configurado para registo de operações
        """
        self.logger = logger
        self.comparison_results = []
        self.summary_stats = {}
        self.dimension_mapping = {}
        
        # Configurações de tolerância para comparação numérica
        self.numeric_tolerance = 1.0  # Tolerância de ±1 para números publicados (arredondados)
        self.percentage_tolerance = 0.001
        
        # Padrões para normalização de dimensões
        self.dimension_patterns = {
            r'\(em\s+branco\)': 'Total',
            r'de\s+(\d+)\s+a\s+(\d+)': r'\1 - \2',
            r'(\d+)\s*anos?\s*ou\s*mais': r'\1+',
            r'menos\s+de\s+(\d+)\s*anos?': r'< \1',
            r'^\s*-\s*$': 'Não especificado',
            r'n\.?\s*d\.?\s*': 'Não disponível'
        }
    
    def get_available_files(self) -> Tuple[List[str], Dict[str, List[str]]]:
        """
        Obtém lista de ficheiros disponíveis para comparação.
        
        Returns:
            Tupla com (ficheiros publicados, ficheiros recriados por pasta)
        """
        published_files = []
        recreated_files = {}
        
        # Procura ficheiros publicados em dataset/comparison/
        comparison_dir = "dataset/comparison"
        if os.path.exists(comparison_dir):
            for file in os.listdir(comparison_dir):
                if file.lower().endswith(('.xlsx', '.xls')) and not file.startswith('~$'):
                    published_files.append(file)
        
        # Procura ficheiros recriados em result/validation/
        validation_dir = "result/validation"
        if os.path.exists(validation_dir):
            for folder in os.listdir(validation_dir):
                folder_path = os.path.join(validation_dir, folder)
                if os.path.isdir(folder_path):
                    folder_files = []
                    
                    # Verifica subpastas series e quadros
                    for subfolder in ['series', 'quadros']:
                        subfolder_path = os.path.join(folder_path, subfolder)
                        if os.path.exists(subfolder_path):
                            for file in os.listdir(subfolder_path):
                                if file.lower().endswith(('.xlsx', '.xls')) and not file.startswith('~$'):
                                    folder_files.append(os.path.join(subfolder, file))
                    
                    # Verifica ficheiros na pasta principal
                    for file in os.listdir(folder_path):
                        file_path = os.path.join(folder_path, file)
                        if os.path.isfile(file_path) and file.lower().endswith(('.xlsx', '.xls')) and not file.startswith('~$'):
                            folder_files.append(file)
                    
                    if folder_files:
                        recreated_files[folder] = folder_files
        
        return published_files, recreated_files
    
    def select_files_interactively(self) -> Tuple[Optional[str], Optional[str]]:
        """
        Interface interativa para seleção de ficheiros.
        
        Returns:
            Tupla com (caminho_ficheiro_publicado, caminho_ficheiro_recriado)
        """
        from colorama import Fore, Style
        
        published_files, recreated_files = self.get_available_files()
        
        if not published_files:
            print(f"{Fore.RED}Nenhum ficheiro encontrado em dataset/comparison/{Style.RESET_ALL}")
            return None, None
        
        if not recreated_files:
            print(f"{Fore.RED}Nenhum ficheiro encontrado em result/validation/{Style.RESET_ALL}")
            return None, None
        
        # Seleção do ficheiro publicado
        print(f"\n{Fore.GREEN}[Seleção de Ficheiro Publicado]{Style.RESET_ALL}")
        print(f"Ficheiros disponíveis em dataset/comparison/:")
        for i, file in enumerate(published_files, 1):
            print(f"  {Fore.WHITE}{i}.{Style.RESET_ALL} {file}")
        
        while True:
            try:
                choice = input(f"\n{Fore.GREEN}>>{Style.RESET_ALL} Escolha o ficheiro publicado (1-{len(published_files)}): ")
                if choice == '0':
                    return None, None
                
                pub_idx = int(choice) - 1
                if 0 <= pub_idx < len(published_files):
                    published_file = os.path.join("dataset/comparison", published_files[pub_idx])
                    break
                else:
                    print(f"{Fore.RED}Número inválido. Tente novamente.{Style.RESET_ALL}")
            except ValueError:
                print(f"{Fore.RED}Entrada inválida. Digite um número.{Style.RESET_ALL}")
        
        # Seleção do ficheiro recriado
        print(f"\n{Fore.GREEN}[Seleção de Ficheiro Recriado]{Style.RESET_ALL}")
        print(f"Pastas disponíveis em result/validation/:")
        
        folder_list = sorted(recreated_files.keys(), key=lambda x: int(x) if x.isdigit() else float('inf'))
        for i, folder in enumerate(folder_list, 1):
            file_count = len(recreated_files[folder])
            print(f"  {Fore.WHITE}{i}.{Style.RESET_ALL} Pasta {folder} ({file_count} ficheiro{'s' if file_count != 1 else ''})")
        
        while True:
            try:
                choice = input(f"\n{Fore.GREEN}>>{Style.RESET_ALL} Escolha a pasta (1-{len(folder_list)}): ")
                if choice == '0':
                    return None, None
                
                folder_idx = int(choice) - 1
                if 0 <= folder_idx < len(folder_list):
                    selected_folder = folder_list[folder_idx]
                    break
                else:
                    print(f"{Fore.RED}Número inválido. Tente novamente.{Style.RESET_ALL}")
            except ValueError:
                print(f"{Fore.RED}Entrada inválida. Digite um número.{Style.RESET_ALL}")
        
        # Seleção do ficheiro específico na pasta
        folder_files = recreated_files[selected_folder]
        if len(folder_files) == 1:
            recreated_file = os.path.join("result/validation", selected_folder, folder_files[0])
        else:
            print(f"\n{Fore.GREEN}[Seleção de Ficheiro na Pasta {selected_folder}]{Style.RESET_ALL}")
            for i, file in enumerate(folder_files, 1):
                print(f"  {Fore.WHITE}{i}.{Style.RESET_ALL} {file}")
            
            while True:
                try:
                    choice = input(f"\n{Fore.GREEN}>>{Style.RESET_ALL} Escolha o ficheiro (1-{len(folder_files)}): ")
                    if choice == '0':
                        return None, None
                    
                    file_idx = int(choice) - 1
                    if 0 <= file_idx < len(folder_files):
                        recreated_file = os.path.join("result/validation", selected_folder, folder_files[file_idx])
                        break
                    else:
                        print(f"{Fore.RED}Número inválido. Tente novamente.{Style.RESET_ALL}")
                except ValueError:
                    print(f"{Fore.RED}Entrada inválida. Digite um número.{Style.RESET_ALL}")
        
        self.logger.info(f"Ficheiros selecionados: {published_file} vs {recreated_file}")
        return published_file, recreated_file
    
    def get_sheet_names(self, file_path: str) -> List[str]:
        """
        Obtém nomes das folhas de um ficheiro Excel.
        
        Args:
            file_path: Caminho do ficheiro Excel
            
        Returns:
            Lista com nomes das folhas
        """
        try:
            wb = load_workbook(file_path, read_only=True)
            sheet_names = wb.sheetnames
            wb.close()
            return sheet_names
        except Exception as e:
            self.logger.error(f"Erro ao ler folhas de {file_path}: {e}")
            return []
    
    def select_sheets_interactively(self, published_file: str, recreated_file: str) -> List[str]:
        """
        Interface para seleção de folhas a comparar.
        
        Args:
            published_file: Caminho do ficheiro publicado
            recreated_file: Caminho do ficheiro recriado
            
        Returns:
            Lista com nomes das folhas a comparar
        """
        from colorama import Fore, Style
        
        pub_sheets = set(self.get_sheet_names(published_file))
        rec_sheets = set(self.get_sheet_names(recreated_file))
        
        common_sheets = sorted(pub_sheets.intersection(rec_sheets))
        
        if not common_sheets:
            print(f"{Fore.RED}Nenhuma folha comum encontrada entre os ficheiros.{Style.RESET_ALL}")
            return []
        
        print(f"\n{Fore.GREEN}[Seleção de Folhas para Comparação]{Style.RESET_ALL}")
        print(f"Folhas comuns encontradas:")
        for i, sheet in enumerate(common_sheets, 1):
            print(f"  {Fore.WHITE}{i}.{Style.RESET_ALL} {sheet}")
        
        print(f"  {Fore.WHITE}T.{Style.RESET_ALL} Todas as folhas")
        print(f"  {Fore.WHITE}0.{Style.RESET_ALL} Cancelar")
        
        while True:
            choice = input(f"\n{Fore.GREEN}>>{Style.RESET_ALL} Escolha as folhas (ex: 1,3,5 ou T para todas): ").strip().lower()
            
            if choice == '0':
                return []
            elif choice == 't':
                return common_sheets
            else:
                try:
                    indices = [int(x.strip()) - 1 for x in choice.split(',') if x.strip()]
                    selected_sheets = [common_sheets[i] for i in indices if 0 <= i < len(common_sheets)]
                    
                    if selected_sheets:
                        return selected_sheets
                    else:
                        print(f"{Fore.RED}Nenhuma folha válida selecionada. Tente novamente.{Style.RESET_ALL}")
                except (ValueError, IndexError):
                    print(f"{Fore.RED}Entrada inválida. Use números separados por vírgula.{Style.RESET_ALL}")
    
    def normalize_value(self, value: Any) -> Optional[float]:
        """
        Normaliza valores numéricos com tratamento robusto para diferentes formatos.
        
        Args:
            value: Valor a normalizar
            
        Returns:
            Valor numérico normalizado ou None se não for número
        """
        if pd.isna(value) or value is None:
            return None
        
        # Se já é número
        if isinstance(value, (int, float)):
            return float(value)
        
        # Converte para string e remove espaços
        str_value = str(value).strip()
        
        # Verifica se é vazio
        if not str_value or str_value in ['-', '']:
            return None
        
        # Preserva o valor original para debug
        original_value = str_value
        
        try:
            # Remove espaços de todos os tipos (incluindo não-quebráveis)
            str_value = re.sub(r'\s+', '', str_value)
            
            # Detecta formato baseado na posição de vírgulas e pontos
            comma_count = str_value.count(',')
            dot_count = str_value.count('.')
            
            if comma_count == 0 and dot_count <= 1:
                # Formato simples: só pontos como decimal
                pass
            elif comma_count == 1 and dot_count == 0:
                # Formato português: vírgula como decimal
                str_value = str_value.replace(',', '.')
            elif comma_count > 0 and dot_count > 0:
                # Formato misto - determinar qual é decimal
                last_comma = str_value.rfind(',')
                last_dot = str_value.rfind('.')
                
                if last_comma > last_dot:
                    # Vírgula é decimal, pontos são separadores de milhares
                    str_value = str_value.replace('.', '').replace(',', '.')
                else:
                    # Ponto é decimal, vírgulas são separadores de milhares
                    str_value = str_value.replace(',', '')
            elif comma_count > 1:
                # Múltiplas vírgulas como separadores de milhares
                str_value = str_value.replace(',', '')
            elif dot_count > 1:
                # Múltiplos pontos como separadores de milhares
                # Se o último segmento tem 1-2 dígitos, trata como decimal
                parts = str_value.split('.')
                if len(parts[-1]) <= 2:
                    str_value = ''.join(parts[:-1]) + '.' + parts[-1]
                else:
                    str_value = str_value.replace('.', '')
            
            # Remove caracteres não numéricos restantes (preserva sinal e decimal)
            str_value = re.sub(r'[^\d.\-+]', '', str_value)
            
            result = float(str_value)
            return result
        except (ValueError, TypeError):
            self.logger.debug(f"Valor não numérico ignorado: '{original_value}' -> '{str_value}'")
            return None
    
    def normalize_dimension_label(self, label: Any) -> str:
        """
        Normaliza etiquetas de dimensão para facilitar correspondência.
        
        Args:
            label: Etiqueta a normalizar
            
        Returns:
            Etiqueta normalizada
        """
        if pd.isna(label) or label is None:
            return 'Total'
        
        str_label = str(label).strip()
        
        # Trata casos especiais antes dos padrões - mais abrangente
        special_cases = [
            '(em branco)', 'em branco', 'blank', '(blank)', 'vazio', '(vazio)', '',
            'total', 'todos', 'geral', 'sum', 'soma', 'all', 'conjunto',
            '(total)', '(todos)', '(geral)', 'totais'
        ]
        if str_label.lower() in special_cases:
            return 'Total'
        
        # Aplica padrões de normalização
        for pattern, replacement in self.dimension_patterns.items():
            str_label = re.sub(pattern, replacement, str_label, flags=re.IGNORECASE)
        
        # Normalização adicional
        str_label = re.sub(r'\s+', ' ', str_label)  # Remove espaços múltiplos
        str_label = str_label.strip()
        
        return str_label
    
    def fuzzy_match_dimension(self, target: str, candidates: List[str], threshold: float = 0.8) -> Optional[str]:
        """
        Encontra correspondência difusa para uma dimensão.
        
        Args:
            target: Dimensão a procurar
            candidates: Lista de candidatos
            threshold: Limiar de similaridade (0-1)
            
        Returns:
            Melhor correspondência ou None se não encontrada
        """
        target_norm = self.normalize_dimension_label(target).lower()
        best_match = None
        best_score = 0
        
        for candidate in candidates:
            candidate_norm = self.normalize_dimension_label(candidate).lower()
            
            # Correspondência exata
            if target_norm == candidate_norm:
                return candidate
            
            # Correspondência difusa
            score = SequenceMatcher(None, target_norm, candidate_norm).ratio()
            if score > best_score and score >= threshold:
                best_score = score
                best_match = candidate
        
        return best_match
    
    def has_background_color(self, cell) -> bool:
        """
        Verifica se uma célula tem cor de fundo de forma robusta.
        
        Args:
            cell: Célula openpyxl
            
        Returns:
            True se tem cor de fundo, False caso contrário
        """
        try:
            if not cell.fill:
                return False
                
            # Verifica tipo de padrão
            if cell.fill.patternType and cell.fill.patternType != 'none':
                # Verifica cores RGB
                if hasattr(cell.fill, 'fgColor') and cell.fill.fgColor:
                    if hasattr(cell.fill.fgColor, 'rgb') and cell.fill.fgColor.rgb:
                        rgb = cell.fill.fgColor.rgb
                        # Cores que consideramos "sem fundo"
                        if rgb not in ['00000000', 'FFFFFFFF', None, 'FFFFFF', '000000']:
                            return True
                    
                    # Verifica theme colors
                    if hasattr(cell.fill.fgColor, 'theme') and cell.fill.fgColor.theme is not None:
                        return True
                        
                    # Verifica indexed colors
                    if hasattr(cell.fill.fgColor, 'indexed') and cell.fill.fgColor.indexed is not None:
                        # Índices 64 e 65 são geralmente automático/sem cor
                        return cell.fill.fgColor.indexed not in [64, 65]
                        
                return True  # Se tem padrão mas não conseguiu verificar cor, assume que tem cor
                
            return False
        except Exception as e:
            # Em caso de erro, assume que não tem cor (mais seguro)
            self.logger.debug(f"Erro ao verificar cor de fundo: {e}")
            return False
    
    def find_data_table(self, ws) -> Tuple[int, int, int, int]:
        """
        Encontra os limites da tabela principal de dados com múltiplas estratégias.
        
        Args:
            ws: Worksheet openpyxl
            
        Returns:
            Tupla com (min_row, max_row, min_col, max_col) da área de dados
        """
        data_cells = []
        
        # Estratégia 1: Procura células numéricas sem cor de fundo (ideal)
        for row in range(1, ws.max_row + 1):
            for col in range(1, ws.max_column + 1):
                cell = ws.cell(row=row, column=col)
                
                if cell.value is not None:
                    normalized_value = self.normalize_value(cell.value)
                    if normalized_value is not None and not self.has_background_color(cell):
                        data_cells.append((row, col))
        
        # Estratégia 2: Se não encontrou dados, relaxa critério (ignora cor de fundo)
        if not data_cells:
            self.logger.info("Estratégia 1 falhou, tentando estratégia 2: ignorando cores de fundo")
            for row in range(1, ws.max_row + 1):
                for col in range(1, ws.max_column + 1):
                    cell = ws.cell(row=row, column=col)
                    
                    if cell.value is not None:
                        normalized_value = self.normalize_value(cell.value)
                        if normalized_value is not None:
                            data_cells.append((row, col))
        
        # Estratégia 3: Se ainda não encontrou, procura por qualquer valor numérico
        if not data_cells:
            self.logger.info("Estratégia 2 falhou, tentando estratégia 3: qualquer valor numérico")
            for row in range(1, ws.max_row + 1):
                for col in range(1, ws.max_column + 1):
                    cell = ws.cell(row=row, column=col)
                    
                    if isinstance(cell.value, (int, float)) and cell.value != 0:
                        data_cells.append((row, col))
        
        # Estratégia 4: Fallback - ignora seções com texto "filtro" ou similar
        if not data_cells:
            self.logger.info("Estratégia 3 falhou, tentando estratégia 4: evita seções 'filtros'")
            skip_rows = set()
            
            # Identifica linhas a evitar
            for row in range(1, min(20, ws.max_row + 1)):  # Verifica primeiras 20 linhas
                for col in range(1, min(10, ws.max_column + 1)):  # Primeiras 10 colunas
                    cell = ws.cell(row=row, column=col)
                    if cell.value and isinstance(cell.value, str):
                        if any(word in str(cell.value).lower() for word in ['filtro', 'filter', 'nota', 'fonte']):
                            skip_rows.add(row)
            
            # Procura dados fora das linhas a evitar
            for row in range(1, ws.max_row + 1):
                if row in skip_rows:
                    continue
                for col in range(1, ws.max_column + 1):
                    cell = ws.cell(row=row, column=col)
                    if cell.value is not None:
                        normalized_value = self.normalize_value(cell.value)
                        if normalized_value is not None:
                            data_cells.append((row, col))
        
        if not data_cells:
            self.logger.warning("Nenhuma célula de dados encontrada com todas as estratégias")
            return (1, ws.max_row, 1, ws.max_column)
        
        # Determina limites da área de dados
        min_row = min(cell[0] for cell in data_cells)
        max_row = max(cell[0] for cell in data_cells)
        min_col = min(cell[1] for cell in data_cells)
        max_col = max(cell[1] for cell in data_cells)
        
        # Expande para incluir cabeçalhos (até 5 linhas/colunas antes dos dados)
        header_buffer = 5
        final_min_row = max(1, min_row - header_buffer)
        final_min_col = max(1, min_col - header_buffer)
        
        self.logger.info(f"Área de dados detectada: linhas {final_min_row}-{max_row}, colunas {final_min_col}-{max_col} ({len(data_cells)} células)")
        return (final_min_row, max_row, final_min_col, max_col)
    
    def get_merged_cell_value(self, ws, row: int, col: int) -> Any:
        """
        Obtém o valor de uma célula considerando células mescladas.
        
        Args:
            ws: Worksheet openpyxl
            row: Número da linha
            col: Número da coluna
            
        Returns:
            Valor da célula ou da célula mesclada principal
        """
        cell = ws.cell(row=row, column=col)
        
        # Se a célula tem valor, retorna diretamente
        if cell.value is not None:
            return cell.value
        
        # Verifica se está numa área mesclada
        for merged_range in ws.merged_cells.ranges:
            if merged_range.min_row <= row <= merged_range.max_row and \
               merged_range.min_col <= col <= merged_range.max_col:
                # Retorna o valor da célula principal (top-left da área mesclada)
                main_cell = ws.cell(row=merged_range.min_row, column=merged_range.min_col)
                return main_cell.value
        
        return None
    
    def get_cell_dimensions(self, ws, data_row: int, data_col: int, 
                           table_bounds: Tuple[int, int, int, int]) -> Dict[str, Any]:
        """
        Obtém as dimensões de uma célula de dados com múltiplas estratégias robustas.
        Para cada valor, encontra:
        - Valor da COLUNA: olha para cima na mesma coluna 
        - Valor da LINHA: olha para esquerda na mesma linha
        
        Args:
            ws: Worksheet openpyxl
            data_row: Linha da célula de dados
            data_col: Coluna da célula de dados
            table_bounds: Limites da tabela (min_row, max_row, min_col, max_col)
            
        Returns:
            Dicionário com dimensões da célula {nome_dimensão: valor_dimensão}
        """
        dimensions = {}
        min_row, max_row, min_col, max_col = table_bounds
        
        # ESTRATÉGIA 1: Busca direta (original)
        col_value = self._find_column_dimension(ws, data_row, data_col, min_row)
        row_value = self._find_row_dimension(ws, data_row, data_col, min_col)
        
        # ESTRATÉGIA 2: Se não encontrou coluna, busca mais ampla
        if not col_value:
            col_value = self._find_column_dimension_extended(ws, data_row, data_col, min_row, max_col)
        
        # ESTRATÉGIA 3: Se não encontrou linha, busca mais ampla  
        if not row_value:
            row_value = self._find_row_dimension_extended(ws, data_row, data_col, min_col, max_row)
        
        # ESTRATÉGIA 4: Fallback com busca em área expandida
        if not col_value or not row_value:
            fallback_col, fallback_row = self._fallback_dimension_search(ws, data_row, data_col, table_bounds)
            if not col_value:
                col_value = fallback_col
            if not row_value:
                row_value = fallback_row
        
        # Adiciona dimensões encontradas com nomes inteligentes
        if col_value:
            normalized_col = self.normalize_value(col_value)
            if normalized_col is not None and 1900 <= normalized_col <= 2030:
                dimensions["Anos"] = col_value
            else:
                dimensions["Coluna"] = col_value
        
        if row_value:
            is_month = any(month in str(row_value).lower() for month in 
                         ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho',
                          'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro',
                          'jan', 'fev', 'mar', 'abr', 'mai', 'jun',
                          'jul', 'ago', 'set', 'out', 'nov', 'dez', 'total'])
            if is_month:
                dimensions["Mês"] = row_value
            else:
                dimensions["Linha"] = row_value
        
        # Debug logging detalhado
        if dimensions:
            dim_str = ", ".join([f"{name}:{value}" for name, value in dimensions.items()])
            self.logger.debug(f"Célula ({data_row},{data_col}): {dim_str}")
        else:
            self.logger.debug(f"Célula ({data_row},{data_col}): NENHUMA DIMENSÃO - col_value='{col_value}', row_value='{row_value}'")
        
        return dimensions
    
    def _find_column_dimension(self, ws, data_row: int, data_col: int, min_row: int) -> str:
        """Busca valor de dimensão na coluna (estratégia principal)"""
        for search_row in range(data_row - 1, min_row - 1, -1):
            cell_value = self.get_merged_cell_value(ws, search_row, data_col)
            if cell_value is not None:
                cell_str = str(cell_value).strip()
                
                if not cell_str or cell_str.lower() in ['', 'anos', 'mes', 'mês', 'unidade: n.º']:
                    continue
                
                if cell_str.lower() in ['total', 'todos', 'geral', '(em branco)', 'em branco']:
                    return 'Total'
                
                normalized = self.normalize_value(cell_value)
                is_year = normalized is not None and 1900 <= normalized <= 2030
                is_month = any(month in cell_str.lower() for month in 
                             ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho',
                              'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro',
                              'jan', 'fev', 'mar', 'abr', 'mai', 'jun',
                              'jul', 'ago', 'set', 'out', 'nov', 'dez'])
                is_text = isinstance(cell_value, str) and len(cell_str) > 0
                
                if is_year or is_month or is_text:
                    return self.normalize_dimension_label(cell_str)
        return None
    
    def _find_row_dimension(self, ws, data_row: int, data_col: int, min_col: int) -> str:
        """Busca valor de dimensão na linha (estratégia principal)"""
        for search_col in range(data_col - 1, min_col - 1, -1):
            cell_value = self.get_merged_cell_value(ws, data_row, search_col)
            if cell_value is not None:
                cell_str = str(cell_value).strip()
                
                if not cell_str or cell_str.lower() in ['', 'mes', 'mês', 'anos']:
                    continue
                
                if cell_str.lower() in ['total', 'todos', 'geral', '(em branco)', 'em branco']:
                    return 'Total'
                
                normalized = self.normalize_value(cell_value)
                is_year = normalized is not None and 1900 <= normalized <= 2030
                is_month = any(month in cell_str.lower() for month in 
                             ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho',
                              'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro',
                              'jan', 'fev', 'mar', 'abr', 'mai', 'jun',
                              'jul', 'ago', 'set', 'out', 'nov', 'dez'])
                is_text = isinstance(cell_value, str) and len(cell_str) > 0
                
                if is_month or (is_text and not is_year):
                    return self.normalize_dimension_label(cell_str)
        return None
    
    def _find_column_dimension_extended(self, ws, data_row: int, data_col: int, min_row: int, max_col: int) -> str:
        """Busca estendida de dimensão de coluna (colunas adjacentes)"""
        # Busca em colunas adjacentes também
        for col_offset in range(-2, 3):  # -2, -1, 0, 1, 2
            search_col = data_col + col_offset
            if search_col < 1 or search_col > max_col:
                continue
                
            for search_row in range(data_row - 1, max(1, min_row - 5), -1):
                cell_value = self.get_merged_cell_value(ws, search_row, search_col)
                if cell_value is not None:
                    cell_str = str(cell_value).strip()
                    
                    if cell_str.lower() in ['total', 'todos', 'geral', '(em branco)', 'em branco']:
                        return 'Total'
                    
                    normalized = self.normalize_value(cell_value)
                    if normalized is not None and 1900 <= normalized <= 2030:
                        return self.normalize_dimension_label(cell_str)
                    
                    if isinstance(cell_value, str) and len(cell_str) > 1 and cell_str.lower() not in ['anos', 'mes', 'mês']:
                        return self.normalize_dimension_label(cell_str)
        return None
    
    def _find_row_dimension_extended(self, ws, data_row: int, data_col: int, min_col: int, max_row: int) -> str:
        """Busca estendida de dimensão de linha (linhas adjacentes)"""
        # Busca em linhas adjacentes também
        for row_offset in range(-2, 3):  # -2, -1, 0, 1, 2
            search_row = data_row + row_offset
            if search_row < 1 or search_row > max_row:
                continue
                
            for search_col in range(data_col - 1, max(1, min_col - 5), -1):
                cell_value = self.get_merged_cell_value(ws, search_row, search_col)
                if cell_value is not None:
                    cell_str = str(cell_value).strip()
                    
                    if cell_str.lower() in ['total', 'todos', 'geral', '(em branco)', 'em branco']:
                        return 'Total'
                    
                    is_month = any(month in cell_str.lower() for month in 
                                 ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho',
                                  'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro',
                                  'jan', 'fev', 'mar', 'abr', 'mai', 'jun',
                                  'jul', 'ago', 'set', 'out', 'nov', 'dez'])
                    
                    if is_month:
                        return self.normalize_dimension_label(cell_str)
                    
                    normalized = self.normalize_value(cell_value)
                    if isinstance(cell_value, str) and len(cell_str) > 1 and cell_str.lower() not in ['anos', 'mes', 'mês'] and not (normalized and 1900 <= normalized <= 2030):
                        return self.normalize_dimension_label(cell_str)
        return None
    
    def _fallback_dimension_search(self, ws, data_row: int, data_col: int, table_bounds) -> tuple:
        """Busca de último recurso usando coordenadas fixas"""
        min_row, max_row, min_col, max_col = table_bounds
        
        col_value = None
        row_value = None
        
        # Fallback: usa primeira linha não vazia da coluna
        if not col_value:
            for search_row in range(min_row, data_row):
                cell_value = self.get_merged_cell_value(ws, search_row, data_col)
                if cell_value is not None and str(cell_value).strip():
                    col_value = f"Col_{data_col}_{self.normalize_dimension_label(str(cell_value))}"
                    break
            
            # Se ainda não encontrou, usa coordenada pura
            if not col_value:
                col_value = f"Col_{data_col}"
        
        # Fallback: usa primeira coluna não vazia da linha
        if not row_value:
            for search_col in range(min_col, data_col):
                cell_value = self.get_merged_cell_value(ws, data_row, search_col)
                if cell_value is not None and str(cell_value).strip():
                    row_value = f"Row_{data_row}_{self.normalize_dimension_label(str(cell_value))}"
                    break
            
            # Se ainda não encontrou, usa coordenada pura
            if not row_value:
                row_value = f"Row_{data_row}"
        
        return col_value, row_value
    
    def get_displayed_value(self, cell) -> Optional[float]:
        """
        Obtém o valor apresentado de uma célula, usando abordagem mais inclusiva.
        
        Args:
            cell: Célula openpyxl
            
        Returns:
            Valor numérico conforme apresentado na célula, ou None se for cabeçalho
        """
        if cell.value is None:
            return None
        
        # Primeiro tenta normalizar qualquer valor (número ou texto)
        normalized_value = None
        
        if isinstance(cell.value, (int, float)):
            normalized_value = float(cell.value)
        else:
            # Tenta normalizar texto para número
            normalized_value = self.normalize_value(cell.value)
            
            # Se é texto e não se converte para número, verifica se é cabeçalho óbvio
            if normalized_value is None and isinstance(cell.value, str):
                str_value = str(cell.value).strip().lower()
                # Lista mais específica de padrões que SÃO cabeçalhos
                explicit_headers = [
                    'janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho',
                    'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro',
                    'january', 'february', 'march', 'april', 'may', 'june',
                    'july', 'august', 'september', 'october', 'november', 'december',
                    'indicador', 'unidade', 'fonte', 'nota', 'total', 'jan', 'fev', 'mar',
                    'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez',
                    'mês', 'mes', 'month', 'ano', 'anos', 'year', 'years'
                ]
                
                # Se contém cabeçalho explícito, não é valor
                if any(header == str_value or header in str_value.split() for header in explicit_headers):
                    return None
                
                # Se não conseguiu normalizar e não é cabeçalho explícito, também não é valor
                return None
        
        # Se conseguiu normalizar para número
        if normalized_value is not None:
            # FILTRO RIGOROSO: Anos típicos (1900-2030) NÃO são valores para comparar
            # São sempre dimensões, nunca dados estatísticos
            if (1900 <= normalized_value <= 2030 and 
                normalized_value == int(normalized_value)):
                # Anos são SEMPRE dimensões, nunca valores para comparar
                return None
            
            # Também rejeita outros valores pequenos que podem ser anos
            if (1800 <= normalized_value <= 2100 and 
                normalized_value == int(normalized_value) and
                normalized_value < 10000):
                # Estes são provavelmente anos ou identificadores, não dados estatísticos
                return None
            
            # Aceita todos os outros valores numéricos
            # Aplica formatação se necessário
            displayed_value = normalized_value
            
            if hasattr(cell, 'number_format') and cell.number_format and cell.number_format != 'General':
                # Para formatos com separadores de milhares ou decimais específicos
                if '0.0' in cell.number_format or '0.00' in cell.number_format:
                    # Arredonda conforme a formatação
                    try:
                        decimal_places = cell.number_format.count('0') - len(cell.number_format.split('.')[0].replace('#', '').replace('0', ''))
                        if decimal_places > 0:
                            displayed_value = round(float(normalized_value), decimal_places)
                    except:
                        pass  # Se der erro, usa valor original
            
            return displayed_value
        
        return None
    
    def detect_crosstab_structure(self, file_path: str, sheet_name: str) -> Dict[str, Any]:
        """
        Deteta automaticamente a estrutura de tabela cruzada numa folha com análise avançada.
        
        Args:
            file_path: Caminho do ficheiro
            sheet_name: Nome da folha
            
        Returns:
            Dicionário com informação da estrutura
        """
        try:
            # Lê folha com openpyxl para preservar formatação
            wb = load_workbook(file_path, data_only=False)
            ws = wb[sheet_name]
            
            # Encontra a área principal de dados
            table_bounds = self.find_data_table(ws)
            min_row, max_row, min_col, max_col = table_bounds
            
            # Procura células de dados na área identificada com múltiplas estratégias
            data_cells = []
            
            # Estratégia 1: Células numéricas sem cor de fundo
            for row in range(min_row, max_row + 1):
                for col in range(min_col, max_col + 1):
                    cell = ws.cell(row=row, column=col)
                    if cell.value is not None:
                        displayed_value = self.get_displayed_value(cell)
                        if displayed_value is not None and not self.has_background_color(cell):
                            data_cells.append({
                                'row': row,
                                'col': col,
                                'value': displayed_value,
                                'cell': cell
                            })
            
            # Estratégia 2: Se não encontrou dados, ignora cor de fundo
            if not data_cells:
                self.logger.info(f"Estratégia 1 falhou para {sheet_name}, tentando estratégia 2: ignorando cores de fundo")
                for row in range(min_row, max_row + 1):
                    for col in range(min_col, max_col + 1):
                        cell = ws.cell(row=row, column=col)
                        if cell.value is not None:
                            displayed_value = self.get_displayed_value(cell)
                            if displayed_value is not None:
                                data_cells.append({
                                    'row': row,
                                    'col': col,
                                    'value': displayed_value,
                                    'cell': cell
                                })
            
            # Estratégia 3: Se ainda não encontrou, procura qualquer valor numérico
            if not data_cells:
                self.logger.info(f"Estratégia 2 falhou para {sheet_name}, tentando estratégia 3: qualquer valor numérico")
                for row in range(min_row, max_row + 1):
                    for col in range(min_col, max_col + 1):
                        cell = ws.cell(row=row, column=col)
                        if isinstance(cell.value, (int, float)) and cell.value != 0:
                            data_cells.append({
                                'row': row,
                                'col': col,
                                'value': float(cell.value),
                                'cell': cell
                            })
            
            if not data_cells:
                wb.close()
                return {'error': 'Nenhuma célula de dados encontrada com todas as estratégias'}
            
            structure = {
                'worksheet': ws,
                'workbook': wb,
                'table_bounds': table_bounds,
                'data_cells': data_cells,
                'total_data_points': len(data_cells)
            }
            
            self.logger.info(f"Estrutura avançada detectada em {sheet_name}: {len(data_cells)} células de dados")
            return structure
            
        except Exception as e:
            self.logger.error(f"Erro ao detectar estrutura avançada em {file_path}[{sheet_name}]: {e}")
            return {'error': str(e)}
    
    def extract_crosstab_data(self, structure: Dict[str, Any]) -> Dict[Tuple, Any]:
        """
        Extrai dados da estrutura de tabela cruzada avançada.
        
        Args:
            structure: Estrutura detectada pela função detect_crosstab_structure
            
        Returns:
            Dicionário mapeando coordenadas de dimensão para dados da célula
        """
        if 'error' in structure:
            return {}
        
        ws = structure['worksheet']
        table_bounds = structure['table_bounds']
        data_cells = structure['data_cells']
        data_map = {}
        
        # Processa cada célula de dados
        for data_cell in data_cells:
            row = data_cell['row']
            col = data_cell['col']
            value = data_cell['value']
            cell = data_cell['cell']
            
            # Obtém dimensões da célula
            dimensions = self.get_cell_dimensions(ws, row, col, table_bounds)
            
            # Debug específico para valores "Total"
            if dimensions and any('Total' in str(dim_val) for dim_val in dimensions.values()):
                self.logger.info(f"[TOTAL DEBUG] Célula ({row},{col}) valor={value}: {dimensions}")
            
            # Cria chave única baseada nas dimensões
            if dimensions:
                # Ordena as dimensões para criar chave consistente
                sorted_dims = []
                for dim_name in sorted(dimensions.keys()):
                    dim_value = dimensions[dim_name]
                    sorted_dims.append(f"{dim_name}:{dim_value}")
                
                coords_key = tuple(sorted_dims)
                
                data_map[coords_key] = {
                    'value': value,
                    'row': row,
                    'col': col,
                    'cell': cell,
                    'dimensions': dimensions
                }
            else:
                # Fallback para células sem dimensões claras
                simple_key = (f"row_{row}", f"col_{col}")
                data_map[simple_key] = {
                    'value': value,
                    'row': row,
                    'col': col,
                    'cell': cell,
                    'dimensions': {}
                }
        
        return data_map
    
    def compare_data_maps(self, published_map: Dict[Tuple, Any], 
                         recreated_map: Dict[Tuple, Any], 
                         sheet_name: str) -> List[Dict[str, Any]]:
        """
        Compara dois mapas de dados extraídos com estrutura avançada.
        
        Args:
            published_map: Dados do ficheiro publicado
            recreated_map: Dados do ficheiro recriado
            sheet_name: Nome da folha
            
        Returns:
            Lista de discrepâncias encontradas
        """
        discrepancies = []
        
        # Verifica valores no ficheiro recriado
        for coords, recreated_data in recreated_map.items():
            recreated_value = recreated_data['value']
            recreated_row = recreated_data['row']
            recreated_col = recreated_data['col']
            recreated_cell = recreated_data['cell']
            
            # Procura correspondência exata
            if coords in published_map:
                published_data = published_map[coords]
                published_value = published_data['value']
                
                # Compara valores com tolerância
                if abs(recreated_value - published_value) > self.numeric_tolerance:
                    discrepancies.append({
                        'sheet': sheet_name,
                        'coordinates': coords,
                        'recreated_value': recreated_value,
                        'published_value': published_value,
                        'difference': recreated_value - published_value,
                        'recreated_row': recreated_row,
                        'recreated_col': recreated_col,
                        'recreated_cell': recreated_cell,
                        'match_type': 'exact'
                    })
            else:
                # Procura correspondência difusa nas coordenadas
                best_match = None
                best_match_coords = None
                best_score = 0
                
                for pub_coords, pub_data in published_map.items():
                    if len(pub_coords) == len(coords):
                        # Calcula similaridade entre coordenadas
                        matches = 0
                        total_dims = len(coords)
                        
                        for rec_dim, pub_dim in zip(coords, pub_coords):
                            if rec_dim == pub_dim:
                                matches += 1
                            else:
                                # Verifica correspondência difusa individual
                                rec_parts = str(rec_dim).split(':')
                                pub_parts = str(pub_dim).split(':')
                                
                                if len(rec_parts) == 2 and len(pub_parts) == 2:
                                    rec_name, rec_value = rec_parts
                                    pub_name, pub_value = pub_parts
                                    
                                    # Se os nomes das dimensões correspondem
                                    if rec_name == pub_name:
                                        # Correspondência difusa mais permissiva
                                        if self.fuzzy_match_dimension(rec_value, [pub_value], 0.6):
                                            matches += 0.9  # Correspondência quase total
                                        elif rec_name == pub_name:  # Mesmo nome de dimensão
                                            matches += 0.5  # Correspondência parcial pelo nome
                        
                        score = matches / total_dims
                        if score > best_score and score >= 0.6:  # Threshold mais baixo para mais matches
                            best_score = score
                            best_match = matches
                            best_match_coords = pub_coords
                
                if best_match_coords:
                    published_data = published_map[best_match_coords]
                    published_value = published_data['value']
                    
                    if abs(recreated_value - published_value) > self.numeric_tolerance:
                        discrepancies.append({
                            'sheet': sheet_name,
                            'coordinates': coords,
                            'matched_coordinates': best_match_coords,
                            'recreated_value': recreated_value,
                            'published_value': published_value,
                            'difference': recreated_value - published_value,
                            'recreated_row': recreated_row,
                            'recreated_col': recreated_col,
                            'recreated_cell': recreated_cell,
                            'match_type': 'fuzzy',
                            'match_score': best_score
                        })
                else:
                    # Valor não encontrado
                    discrepancies.append({
                        'sheet': sheet_name,
                        'coordinates': coords,
                        'recreated_value': recreated_value,
                        'published_value': None,
                        'difference': None,
                        'recreated_row': recreated_row,
                        'recreated_col': recreated_col,
                        'recreated_cell': recreated_cell,
                        'match_type': 'not_found'
                    })
        
        return discrepancies
    
    def compare_files(self, published_file: str, recreated_file: str, 
                     sheet_names: List[str]) -> Dict[str, Any]:
        """
        Compara dois ficheiros Excel nas folhas especificadas.
        
        Args:
            published_file: Caminho do ficheiro publicado
            recreated_file: Caminho do ficheiro recriado
            sheet_names: Lista de folhas a comparar
            
        Returns:
            Resultados da comparação
        """
        self.logger.info(f"Iniciando comparação entre {published_file} e {recreated_file}")
        
        comparison_results = {
            'published_file': published_file,
            'recreated_file': recreated_file,
            'sheets_compared': sheet_names,
            'total_discrepancies': 0,
            'sheet_results': {},
            'summary': {}
        }
        
        total_discrepancies = 0
        
        for sheet_name in sheet_names:
            self.logger.info(f"Comparando folha: {sheet_name}")
            
            # Deteta estruturas
            pub_structure = self.detect_crosstab_structure(published_file, sheet_name)
            rec_structure = self.detect_crosstab_structure(recreated_file, sheet_name)
            
            if 'error' in pub_structure or 'error' in rec_structure:
                self.logger.error(f"Erro na estrutura da folha {sheet_name}")
                comparison_results['sheet_results'][sheet_name] = {
                    'error': f"Erro na estrutura: {pub_structure.get('error', '')} | {rec_structure.get('error', '')}"
                }
                # Limpa recursos se possível
                if 'workbook' in pub_structure:
                    pub_structure['workbook'].close()
                if 'workbook' in rec_structure:
                    rec_structure['workbook'].close()
                continue
            
            # Extrai dados
            self.logger.info(f"[{sheet_name}] Extraindo dados do arquivo PUBLICADO...")
            pub_data = self.extract_crosstab_data(pub_structure)
            self.logger.info(f"[{sheet_name}] Arquivo PUBLICADO: {len(pub_data)} pontos de dados")
            
            self.logger.info(f"[{sheet_name}] Extraindo dados do arquivo RECRIADO...")
            rec_data = self.extract_crosstab_data(rec_structure)
            self.logger.info(f"[{sheet_name}] Arquivo RECRIADO: {len(rec_data)} pontos de dados")
            
            # Debug comparativo de coordenadas "Total"
            pub_total_coords = [coords for coords in pub_data.keys() if 'Total' in str(coords)]
            rec_total_coords = [coords for coords in rec_data.keys() if 'Total' in str(coords)]
            self.logger.info(f"[{sheet_name}] Coordenadas 'Total' - Publicado: {len(pub_total_coords)}, Recriado: {len(rec_total_coords)}")
            
            if pub_total_coords:
                self.logger.info(f"[{sheet_name}] Primeiras coordenadas 'Total' no PUBLICADO: {pub_total_coords[:3]}")
            if rec_total_coords:
                self.logger.info(f"[{sheet_name}] Primeiras coordenadas 'Total' no RECRIADO: {rec_total_coords[:3]}")
            
            # Compara dados
            discrepancies = self.compare_data_maps(pub_data, rec_data, sheet_name)
            
            total_discrepancies += len(discrepancies)
            
            comparison_results['sheet_results'][sheet_name] = {
                'published_data_points': len(pub_data),
                'recreated_data_points': len(rec_data),
                'discrepancies': discrepancies,
                'discrepancy_count': len(discrepancies),
                'pub_structure': pub_structure,  # Mantém para criar relatório visual
                'rec_structure': rec_structure
            }
            
            self.logger.info(f"Folha {sheet_name}: {len(pub_data)} vs {len(rec_data)} pontos, {len(discrepancies)} discrepâncias")
        
        comparison_results['total_discrepancies'] = total_discrepancies
        
        # Calcula estatísticas de resumo
        total_published_points = sum(r.get('published_data_points', 0) for r in comparison_results['sheet_results'].values())
        total_recreated_points = sum(r.get('recreated_data_points', 0) for r in comparison_results['sheet_results'].values())
        
        comparison_results['summary'] = {
            'total_published_points': total_published_points,
            'total_recreated_points': total_recreated_points,
            'total_discrepancies': total_discrepancies,
            'accuracy_percentage': (max(0, total_recreated_points - total_discrepancies) / max(1, total_recreated_points)) * 100
        }
        
        self.comparison_results = comparison_results
        return comparison_results
    
    def copy_worksheet_with_formatting(self, source_ws, target_wb, target_name: str):
        """
        Copia uma folha preservando toda a formatação original.
        
        Args:
            source_ws: Worksheet origem
            target_wb: Workbook destino
            target_name: Nome da folha destino
            
        Returns:
            Worksheet copiada
        """
        target_ws = target_wb.create_sheet(target_name)
        
        # Copia dimensões das colunas
        for col_letter in source_ws.column_dimensions:
            if source_ws.column_dimensions[col_letter].width:
                target_ws.column_dimensions[col_letter].width = source_ws.column_dimensions[col_letter].width
        
        # Copia dimensões das linhas
        for row_num in source_ws.row_dimensions:
            if source_ws.row_dimensions[row_num].height:
                target_ws.row_dimensions[row_num].height = source_ws.row_dimensions[row_num].height
        
        # Copia células e formatação
        for row in source_ws.iter_rows():
            for cell in row:
                target_cell = target_ws.cell(row=cell.row, column=cell.column)
                
                # Copia valor
                target_cell.value = cell.value
                
                # Copia formatação
                if cell.has_style:
                    target_cell.font = copy(cell.font)
                    target_cell.border = copy(cell.border)
                    target_cell.fill = copy(cell.fill)
                    target_cell.number_format = cell.number_format
                    target_cell.protection = copy(cell.protection)
                    target_cell.alignment = copy(cell.alignment)
        
        # Copia células mescladas
        for merged_range in source_ws.merged_cells.ranges:
            target_ws.merge_cells(str(merged_range))
        
        return target_ws
    
    def create_highlighted_report_sheet(self, source_ws, target_wb, sheet_name: str, 
                                      discrepancies: List[Dict[str, Any]]):
        """
        Cria folha de relatório com discrepâncias destacadas em amarelo.
        
        Args:
            source_ws: Worksheet origem (recriada)
            target_wb: Workbook destino
            sheet_name: Nome da folha
            discrepancies: Lista de discrepâncias
            
        Returns:
            Worksheet com destaques
        """
        # Copia folha com formatação original
        target_name = f"Resumo_{sheet_name}"[:31]
        target_ws = self.copy_worksheet_with_formatting(source_ws, target_wb, target_name)
        
        # Define cor amarela para destacar discrepâncias
        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        
        # Aplica destaque nas células com discrepâncias
        for discrepancy in discrepancies:
            if 'recreated_row' in discrepancy and 'recreated_col' in discrepancy:
                row = discrepancy['recreated_row']
                col = discrepancy['recreated_col']
                
                target_cell = target_ws.cell(row=row, column=col)
                target_cell.fill = yellow_fill
                
                # Adiciona comentário com detalhes da discrepância
                comment_text = f"DISCREPÂNCIA DETECTADA\n"
                comment_text += f"Valor recriado: {discrepancy['recreated_value']}\n"
                
                if discrepancy.get('published_value') is not None:
                    comment_text += f"Valor publicado: {discrepancy['published_value']}\n"
                    comment_text += f"Diferença: {discrepancy.get('difference', 'N/A')}\n"
                else:
                    comment_text += "Valor não encontrado no ficheiro publicado\n"
                
                comment_text += f"Tipo: {discrepancy['match_type']}"
                
                # Adiciona comentário (se suportado)
                try:
                    target_cell.comment = comment_text
                except:
                    # Se não conseguir adicionar comentário, continua
                    pass
        
        return target_ws
    
    def generate_comparison_report(self, results: Dict[str, Any], output_dir: str = "result/comparison") -> str:
        """
        Gera relatório Excel visual detalhado da comparação com destaques.
        
        Args:
            results: Resultados da comparação
            output_dir: Diretório de saída
            
        Returns:
            Caminho do ficheiro de relatório gerado
        """
        ensure_directory_exists(output_dir)
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        report_file = os.path.join(output_dir, f"visual_comparison_report_{timestamp}.xlsx")
        
        wb = Workbook()
        
        # Remove folha padrão
        wb.remove(wb.active)
        
        # Folha de resumo geral
        summary_ws = wb.create_sheet("Resumo_Geral")
        self._create_summary_sheet(summary_ws, results)
        
        # Cria folhas visuais para cada folha comparada
        for sheet_name, sheet_results in results['sheet_results'].items():
            if 'discrepancies' in sheet_results and 'rec_structure' in sheet_results:
                rec_structure = sheet_results['rec_structure']
                
                if 'worksheet' in rec_structure:
                    source_ws = rec_structure['worksheet']
                    discrepancies = sheet_results['discrepancies']
                    
                    # Cria folha visual com destaques
                    self.create_highlighted_report_sheet(
                        source_ws, wb, sheet_name, discrepancies
                    )
        
        # Folha de discrepâncias detalhadas
        for sheet_name, sheet_results in results['sheet_results'].items():
            if 'discrepancies' in sheet_results and sheet_results['discrepancies']:
                disc_ws = wb.create_sheet(f"Detalhes_{sheet_name}"[:31])
                self._create_discrepancy_sheet(disc_ws, sheet_results, sheet_name)
        
        # Folha de detalhes técnicos
        details_ws = wb.create_sheet("Info_Tecnica")
        self._create_technical_details_sheet(details_ws, results)
        
        # Limpa recursos das estruturas
        for sheet_results in results['sheet_results'].values():
            if 'pub_structure' in sheet_results and 'workbook' in sheet_results['pub_structure']:
                sheet_results['pub_structure']['workbook'].close()
            if 'rec_structure' in sheet_results and 'workbook' in sheet_results['rec_structure']:
                sheet_results['rec_structure']['workbook'].close()
        
        wb.save(report_file)
        self.logger.info(f"Relatório visual gerado: {report_file}")
        
        return report_file
    
    def _create_summary_sheet(self, ws, results: Dict[str, Any]):
        """Cria folha de resumo do relatório."""
        # Cabeçalho
        ws['A1'] = "RELATÓRIO DE COMPARAÇÃO DE DADOS"
        ws['A1'].font = Font(bold=True, size=16)
        
        row = 3
        ws[f'A{row}'] = "Ficheiro Publicado:"
        ws[f'B{row}'] = results['published_file']
        row += 1
        
        ws[f'A{row}'] = "Ficheiro Recriado:"
        ws[f'B{row}'] = results['recreated_file']
        row += 1
        
        ws[f'A{row}'] = "Data da Comparação:"
        ws[f'B{row}'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        row += 2
        
        # Estatísticas gerais
        ws[f'A{row}'] = "ESTATÍSTICAS GERAIS"
        ws[f'A{row}'].font = Font(bold=True)
        row += 1
        
        summary = results['summary']
        ws[f'A{row}'] = "Total de pontos publicados:"
        ws[f'B{row}'] = summary['total_published_points']
        row += 1
        
        ws[f'A{row}'] = "Total de pontos recriados:"
        ws[f'B{row}'] = summary['total_recreated_points']
        row += 1
        
        ws[f'A{row}'] = "Total de discrepâncias:"
        ws[f'B{row}'] = summary['total_discrepancies']
        row += 1
        
        ws[f'A{row}'] = "Precisão:"
        ws[f'B{row}'] = f"{summary['accuracy_percentage']:.2f}%"
        row += 2
        
        # Resumo por folha
        ws[f'A{row}'] = "RESUMO POR FOLHA"
        ws[f'A{row}'].font = Font(bold=True)
        row += 1
        
        ws[f'A{row}'] = "Folha"
        ws[f'B{row}'] = "Pontos Publicados"
        ws[f'C{row}'] = "Pontos Recriados"
        ws[f'D{row}'] = "Discrepâncias"
        ws[f'E{row}'] = "Precisão (%)"
        
        # Formatação cabeçalho
        for col in ['A', 'B', 'C', 'D', 'E']:
            ws[f'{col}{row}'].font = Font(bold=True)
            ws[f'{col}{row}'].fill = PatternFill(start_color="CCCCCC", end_color="CCCCCC", fill_type="solid")
        
        row += 1
        
        for sheet_name, sheet_results in results['sheet_results'].items():
            if 'error' not in sheet_results:
                pub_points = sheet_results.get('published_data_points', 0)
                rec_points = sheet_results.get('recreated_data_points', 0)
                discrepancies = sheet_results.get('discrepancy_count', 0)
                accuracy = (max(0, rec_points - discrepancies) / max(1, rec_points)) * 100
                
                ws[f'A{row}'] = sheet_name
                ws[f'B{row}'] = pub_points
                ws[f'C{row}'] = rec_points
                ws[f'D{row}'] = discrepancies
                ws[f'E{row}'] = f"{accuracy:.2f}%"
                row += 1
        
        # Ajusta largura das colunas
        for col in ['A', 'B', 'C', 'D', 'E']:
            ws.column_dimensions[col].width = 20
    
    def _create_discrepancy_sheet(self, ws, sheet_results: Dict[str, Any], sheet_name: str):
        """Cria folha de discrepâncias para uma folha específica."""
        ws['A1'] = f"DISCREPÂNCIAS - {sheet_name}"
        ws['A1'].font = Font(bold=True, size=14)
        
        discrepancies = sheet_results['discrepancies']
        
        if not discrepancies:
            ws['A3'] = "Nenhuma discrepância encontrada nesta folha."
            return
        
        # Cabeçalhos
        headers = ['Coordenadas', 'Valor Recriado', 'Valor Publicado', 'Diferença', 'Tipo Correspondência']
        for i, header in enumerate(headers, 1):
            cell = ws.cell(row=3, column=i, value=header)
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="CCCCCC", end_color="CCCCCC", fill_type="solid")
        
        # Dados das discrepâncias
        for i, disc in enumerate(discrepancies, 4):
            ws.cell(row=i, column=1, value=str(disc['coordinates']))
            ws.cell(row=i, column=2, value=disc['recreated_value'])
            ws.cell(row=i, column=3, value=disc.get('published_value', 'N/A'))
            ws.cell(row=i, column=4, value=disc.get('difference', 'N/A'))
            ws.cell(row=i, column=5, value=disc['match_type'])
            
            # Destaca discrepâncias significativas
            if disc.get('difference') and abs(disc['difference']) > self.numeric_tolerance * 10:
                for col in range(1, 6):
                    ws.cell(row=i, column=col).fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
        
        # Ajusta largura das colunas
        for col in range(1, 6):
            ws.column_dimensions[chr(64 + col)].width = 25
    
    def _create_technical_details_sheet(self, ws, results: Dict[str, Any]):
        """Cria folha com detalhes técnicos da comparação."""
        ws['A1'] = "DETALHES TÉCNICOS DA COMPARAÇÃO"
        ws['A1'].font = Font(bold=True, size=14)
        
        row = 3
        
        # Parâmetros de comparação
        ws[f'A{row}'] = "PARÂMETROS DE COMPARAÇÃO"
        ws[f'A{row}'].font = Font(bold=True)
        row += 1
        
        ws[f'A{row}'] = "Tolerância numérica:"
        ws[f'B{row}'] = self.numeric_tolerance
        row += 1
        
        ws[f'A{row}'] = "Limiar correspondência difusa:"
        ws[f'B{row}'] = "0.8"
        row += 2
        
        # Informações das folhas
        ws[f'A{row}'] = "INFORMAÇÕES DAS FOLHAS"
        ws[f'A{row}'].font = Font(bold=True)
        row += 1
        
        for sheet_name, sheet_results in results['sheet_results'].items():
            ws[f'A{row}'] = f"Folha: {sheet_name}"
            ws[f'A{row}'].font = Font(bold=True)
            row += 1
            
            if 'error' in sheet_results:
                ws[f'B{row}'] = f"Erro: {sheet_results['error']}"
                row += 1
            else:
                ws[f'B{row}'] = f"Pontos de dados publicados: {sheet_results.get('published_data_points', 0)}"
                row += 1
                ws[f'B{row}'] = f"Pontos de dados recriados: {sheet_results.get('recreated_data_points', 0)}"
                row += 1
                ws[f'B{row}'] = f"Discrepâncias encontradas: {sheet_results.get('discrepancy_count', 0)}"
                row += 1
            row += 1
        
        # Ajusta largura das colunas
        ws.column_dimensions['A'].width = 30
        ws.column_dimensions['B'].width = 40


def run_interactive_comparison(logger: logging.Logger):
    """
    Executa o processo interativo de comparação de dados.
    
    Args:
        logger: Logger configurado
    """
    from colorama import Fore, Style
    
    try:
        # Cria instância do comparador
        comparator = DataComparator(logger)
        
        print(f"\n{Fore.GREEN}[Comparação Inteligente de Dados Excel]{Style.RESET_ALL}")
        print("Esta funcionalidade compara ficheiros Excel com estruturas de tabela cruzada,")
        print("aplicando normalização avançada e correspondência difusa de dimensões.\n")
        
        # Seleção interativa de ficheiros
        published_file, recreated_file = comparator.select_files_interactively()
        
        if not published_file or not recreated_file:
            print(f"{Fore.YELLOW}Operação cancelada.{Style.RESET_ALL}")
            return
        
        # Seleção de folhas
        selected_sheets = comparator.select_sheets_interactively(published_file, recreated_file)
        
        if not selected_sheets:
            print(f"{Fore.YELLOW}Nenhuma folha selecionada para comparação.{Style.RESET_ALL}")
            return
        
        # Confirmação
        print(f"\n{Fore.CYAN}Configuração da Comparação:{Style.RESET_ALL}")
        print(f"Ficheiro publicado: {published_file}")
        print(f"Ficheiro recriado: {recreated_file}")
        print(f"Folhas a comparar: {', '.join(selected_sheets)}")
        
        confirm = input(f"\n{Fore.GREEN}Continuar com a comparação? (s/N):{Style.RESET_ALL} ").strip().lower()
        if confirm not in ['s', 'sim', 'y', 'yes']:
            print(f"{Fore.YELLOW}Operação cancelada.{Style.RESET_ALL}")
            return
        
        # Executa comparação
        print(f"\n{Fore.CYAN}Iniciando comparação...{Style.RESET_ALL}")
        results = comparator.compare_files(published_file, recreated_file, selected_sheets)
        
        # Apresenta resumo
        print(f"\n{Fore.GREEN}Comparação concluída!{Style.RESET_ALL}")
        print(f"\n{Fore.CYAN}Resumo dos Resultados:{Style.RESET_ALL}")
        
        summary = results['summary']
        print(f"• Total de pontos publicados: {summary['total_published_points']}")
        print(f"• Total de pontos recriados: {summary['total_recreated_points']}")
        print(f"• Total de discrepâncias: {summary['total_discrepancies']}")
        print(f"• Precisão geral: {summary['accuracy_percentage']:.2f}%")
        
        # Resumo por folha
        if results['sheet_results']:
            print(f"\n{Fore.CYAN}Resumo por Folha:{Style.RESET_ALL}")
            for sheet_name, sheet_results in results['sheet_results'].items():
                if 'error' not in sheet_results:
                    discrepancies = sheet_results.get('discrepancy_count', 0)
                    rec_points = sheet_results.get('recreated_data_points', 0)
                    accuracy = (max(0, rec_points - discrepancies) / max(1, rec_points)) * 100
                    
                    status_color = Fore.GREEN if accuracy > 95 else Fore.YELLOW if accuracy > 80 else Fore.RED
                    print(f"• {sheet_name}: {status_color}{accuracy:.2f}%{Style.RESET_ALL} "
                          f"({discrepancies} discrepâncias em {rec_points} pontos)")
                else:
                    print(f"• {sheet_name}: {Fore.RED}Erro{Style.RESET_ALL} - {sheet_results['error']}")
        
        # Gera relatório
        print(f"\n{Fore.CYAN}Gerando relatório detalhado...{Style.RESET_ALL}")
        report_file = comparator.generate_comparison_report(results)
        
        print(f"\n{Fore.GREEN}✅ Relatório gerado com sucesso!{Style.RESET_ALL}")
        print(f"Ficheiro: {report_file}")
        
        # Mostra próximos passos
        if summary['total_discrepancies'] > 0:
            print(f"\n{Fore.YELLOW}Próximos Passos:{Style.RESET_ALL}")
            print("1. Revise o relatório Excel gerado para detalhes das discrepâncias")
            print("2. Verifique se as diferenças são aceitáveis ou requerem correção")
            print("3. Ajuste os dados recriados conforme necessário")
        else:
            print(f"\n{Fore.GREEN}✅ Perfeito!{Style.RESET_ALL} Nenhuma discrepância encontrada.")
            print("Os dados recriados correspondem exatamente aos publicados.")
        
    except Exception as e:
        logger.error(f"Erro durante comparação: {e}", exc_info=True)
        print(f"\n{Fore.RED}Erro crítico durante a comparação:{Style.RESET_ALL} {str(e)}")
        print("Consulte os logs para detalhes técnicos.")