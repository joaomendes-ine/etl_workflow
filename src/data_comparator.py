"""
Módulo de Comparação Inteligente de Dados Excel - Versão Redesenhada
Sistema simplificado para comparação de ficheiros Excel com estruturas hierárquicas.
Foca em reconhecimento visual da estrutura e equivalência semântica.
"""

import os
import pandas as pd
import numpy as np
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import re
from typing import Dict, List, Tuple, Any, Optional, Set
from datetime import datetime
import logging
from dataclasses import dataclass
from src.utils import ensure_directory_exists


@dataclass
class DataPoint:
    """Representa um ponto de dados com coordenadas hierárquicas."""
    column_level_1: str = ""
    column_level_2: str = ""
    row_level_1: str = ""
    row_level_2: str = ""
    value: float = 0.0
    row: int = 0
    col: int = 0
    
    def get_coordinate_key(self) -> tuple:
        """Retorna chave de coordenadas para matching."""
        return (self.column_level_1, self.column_level_2, self.row_level_1, self.row_level_2)
    
    def __str__(self):
        return f"({self.column_level_1}|{self.column_level_2}) x ({self.row_level_1}|{self.row_level_2}) = {self.value}"


class HierarchicalDataComparator:
    """
    Comparador de dados hierárquicos simplificado e robusto.
    Foca na estrutura visual e equivalência semântica.
    """
    
    def __init__(self, logger: logging.Logger):
        """
        Inicializa o comparador hierárquico.
        
        Args:
            logger: Logger configurado
        """
        self.logger = logger
        self.comparison_results = []
        
        # Tolerância numérica conservadora mas realista
        self.numeric_tolerance = 0.001  # Muito baixa para evitar falsos positivos
        
        # Mapeamento de equivalência semântica expandido
        self.semantic_equivalence = {
            # Equivalências de asteriscos
            "****": "4",
            "***": "3", 
            "**": "2",
            "*": "1",
            
            # Equivalências de Total/Em branco (bi-direcionais)
            "Total": "(Em branco)",
            "(Em branco)": "Total",
            "(em branco)": "Total", 
            "total": "Total",
            "TOTAL": "Total",
            
            # Normalizações de formatação
            "em branco": "(Em branco)",
            "EM BRANCO": "(Em branco)",
            "Em Branco": "(Em branco)",
            
            # Normalizações de idade
            "16 - 24 anos": "De 16 a 24 anos",
            "25 - 34 anos": "De 25 a 34 anos", 
            "35 - 44 anos": "De 35 a 44 anos",
            "45 - 54 anos": "De 45 a 54 anos",
            "55 - 64 anos": "De 55 a 64 anos",
            "65 anos ou mais": "65 ou mais anos",
            "menos de 18 anos": "Menos de 18 anos",
            
            # Normalizações de dados ausentes
            "n.d.": "Não disponível",
            "N.D.": "Não disponível",
            "-": "",
            "...": "",
            
            # Normalizações de género
            "Homens": "Masculino",
            "Mulheres": "Feminino",
            "H": "Masculino",
            "M": "Feminino",
        }
        
        # Cores para identificação de cabeçalhos
        self.header_colors = {
            'blue_dark': ['FF0070C0', 'FF002060', 'FF1F4E79'],  # Publicados
            'blue_light': ['FFB8CCE4', 'FFDCE6F1', 'FFADD8E6'], # Recriados
        }
    
    def normalize_value_conservative(self, value: Any) -> Optional[float]:
        """
        Normalização conservadora de valores - preserva exatidão.
        
        Args:
            value: Valor a normalizar
            
        Returns:
            Valor normalizado ou None se não for numérico
        """
        if value is None or value == "":
            return None
            
        # Se já é número, usa diretamente
        if isinstance(value, (int, float)):
            if pd.isna(value) or np.isinf(value):
                return None
            return float(value)
        
        # Se é string, tenta converter
        if isinstance(value, str):
            # Remove espaços
            value = value.strip()
            
            # Rejeita strings claramente não numéricas
            if any(char in value.lower() for char in ['total', 'não', 'n/a', 'nd']):
                return None
            
            # Verifica se é apenas um hífen (significa valor em falta)
            if value == '-':
                return None
            
            # Substitui vírgula por ponto (formato português)
            value = value.replace(',', '.')
            
            # Remove separadores de milhares (espaços)
            value = value.replace(' ', '')
            
            try:
                num_value = float(value)
                
                # Rejeita anos (identificados como dimensões, não valores)
                if 1800 <= num_value <= 2100 and num_value == int(num_value):
                    return None
                    
                return num_value
            except (ValueError, TypeError):
                return None
        
        return None

    def find_data_table(self, ws) -> Tuple[int, int, int, int]:
        """
        Identifica automaticamente a tabela principal de dados numa folha Excel.
        Ignora seções como "Filtros" e foca na área principal de dados.
        
        Args:
            ws: Worksheet do openpyxl
            
        Returns:
            Tupla (min_row, max_row, min_col, max_col) da tabela principal
        """
        data_cells = []
        
        # Procura células com valores numéricos sem cor de fundo
        for row in range(1, min(ws.max_row + 1, 200)):  # Limita busca
            for col in range(1, min(ws.max_column + 1, 50)):
                cell = ws.cell(row=row, column=col)
                
                # Verifica se é um valor de dados válido
                if (cell.value is not None and 
                    self.normalize_value_conservative(cell.value) is not None and
                    not self.has_background_color(cell)):
                    data_cells.append((row, col))
        
        if not data_cells:
            # Fallback: usa toda a área de dados
            return (1, ws.max_row, 1, ws.max_column)
        
        # Define limites da tabela principal
        min_row = min(cell[0] for cell in data_cells)
        max_row = max(cell[0] for cell in data_cells)
        min_col = min(cell[1] for cell in data_cells)
        max_col = max(cell[1] for cell in data_cells)
        
        # Expande para incluir cabeçalhos (5 linhas/colunas de buffer)
        min_row = max(1, min_row - 5)
        min_col = max(1, min_col - 5)
        max_row = min(ws.max_row, max_row + 2)
        max_col = min(ws.max_column, max_col + 2)
        
        return (min_row, max_row, min_col, max_col)

    def has_background_color(self, cell) -> bool:
        """
        Verifica se uma célula tem cor de fundo.
        
        Args:
            cell: Célula do openpyxl
            
        Returns:
            True se tem cor de fundo
        """
        if not hasattr(cell, 'fill') or not cell.fill:
            return False
        
        # Verifica se tem cor RGB definida
        if hasattr(cell.fill, 'start_color') and cell.fill.start_color:
            rgb = cell.fill.start_color.rgb
            # Ignora branco e None
            return rgb and rgb not in ['FFFFFFFF', 'FF000000', '00000000']
        
        return False

    def get_merged_cell_value(self, ws, row: int, col: int) -> Any:
        """
        Obtém o valor de uma célula considerando células mescladas.
        
        Args:
            ws: Worksheet
            row: Linha
            col: Coluna
            
        Returns:
            Valor da célula ou célula principal se mesclada
        """
        cell = ws.cell(row=row, column=col)
        
        # Se tem valor direto, retorna
        if cell.value is not None:
            return cell.value
        
        # Verifica se está numa área mesclada
        for merged_range in ws.merged_cells.ranges:
            if (merged_range.min_row <= row <= merged_range.max_row and
                merged_range.min_col <= col <= merged_range.max_col):
                # Retorna valor da célula principal (top-left)
                main_cell = ws.cell(row=merged_range.min_row, column=merged_range.min_col)
                return main_cell.value
        
        return cell.value

    def is_likely_header(self, value: Any) -> bool:
        """
        Verifica se um valor é provável de ser um cabeçalho (não um valor de dados).
        
        Args:
            value: Valor a verificar
            
        Returns:
            True se parecer ser cabeçalho
        """
        if not value:
            return False
        
        value_str = str(value).strip()
        
        # Valores claramente numéricos puros não são cabeçalhos
        try:
            float_val = float(value_str.replace(' ', '').replace(',', '.'))
            # Se é um número grande (não ano), provavelmente é dado
            if float_val > 2100 or (float_val > 0 and float_val < 1800):
                return False
        except:
            pass
        
        # Padrões típicos de cabeçalhos
        header_patterns = [
            'anos', 'ano', 'categoria', 'subcategoria', 'estabelecimentos',
            'unidade', 'total', 'em branco', 'região', 'distrito', 'concelho'
        ]
        
        value_lower = value_str.lower()
        if any(pattern in value_lower for pattern in header_patterns):
            return True
        
        # Anos (1800-2100) podem ser cabeçalhos
        try:
            year = int(value_str)
            if 1800 <= year <= 2100:
                return True
        except:
            pass
        
        # Strings curtas com letras (códigos, siglas) podem ser cabeçalhos
        if len(value_str) <= 10 and any(c.isalpha() for c in value_str):
            return True
        
        # Texto longo é geralmente cabeçalho
        if len(value_str) > 10 and not value_str.replace(' ', '').replace(',', '').replace('.', '').isdigit():
            return True
        
        return False

    def get_cell_dimensions(self, ws, data_row: int, data_col: int, 
                           table_bounds: Tuple[int, int, int, int]) -> Dict[str, Any]:
        """
        Mapeia as dimensões (cabeçalhos de linha/coluna) para uma célula de dados.
        Algoritmo inteligente que distingue cabeçalhos de valores.
        
        Args:
            ws: Worksheet
            data_row: Linha da célula de dados
            data_col: Coluna da célula de dados
            table_bounds: Limites da tabela (min_row, max_row, min_col, max_col)
            
        Returns:
            Dicionário com dimensões mapeadas
        """
        min_row, max_row, min_col, max_col = table_bounds
        
        dimensions = {
            'column_headers': [],
            'row_headers': [],
            'column_coords': f"{get_column_letter(data_col)}",
            'row_coords': str(data_row)
        }
        
        # Anda para cima na coluna para encontrar cabeçalhos verdadeiros
        for check_row in range(data_row - 1, max(1, min_row - 5), -1):  # Vai até antes da tabela
            cell_value = self.get_merged_cell_value(ws, check_row, data_col)
            
            if cell_value and str(cell_value).strip():
                value_str = str(cell_value).strip()
                
                # Só aceita se parecer realmente um cabeçalho
                if self.is_likely_header(value_str):
                    header_value = self.get_semantic_equivalent(value_str)
                    
                    if header_value not in dimensions['column_headers']:
                        dimensions['column_headers'].append(header_value)
                    
                    # Para de procurar se encontrou cabeçalhos suficientes
                    if len(dimensions['column_headers']) >= 2:
                        break
        
        # Anda para a esquerda na linha para encontrar rótulos verdadeiros
        for check_col in range(data_col - 1, max(1, min_col - 5), -1):  # Vai até antes da tabela
            cell_value = self.get_merged_cell_value(ws, data_row, check_col)
            
            if cell_value and str(cell_value).strip():
                value_str = str(cell_value).strip()
                
                # Só aceita se parecer realmente um rótulo/cabeçalho
                if self.is_likely_header(value_str):
                    row_value = self.get_semantic_equivalent(value_str)
                    
                    if row_value not in dimensions['row_headers']:
                        dimensions['row_headers'].append(row_value)
                    
                    # Para de procurar se encontrou rótulos suficientes
                    if len(dimensions['row_headers']) >= 2:
                        break
        
        # Garante pelo menos 2 níveis (preenche com vazios se necessário)
        while len(dimensions['column_headers']) < 2:
            dimensions['column_headers'].append('')
        while len(dimensions['row_headers']) < 2:
            dimensions['row_headers'].append('')
        
        return dimensions

    def get_displayed_value(self, cell) -> Optional[float]:
        """
        Obtém o valor como é apresentado ao utilizador (considerando formatação).
        
        Args:
            cell: Célula do openpyxl
            
        Returns:
            Valor numérico apresentado ou None
        """
        if cell.value is None:
            return None
        
        # Se tem number_format específico, usa formatação
        if hasattr(cell, 'number_format') and cell.number_format:
            try:
                # Para formatos portugueses com separadores
                if isinstance(cell.value, (int, float)):
                    return float(cell.value)
            except:
                pass
        
        # Fallback para normalização conservadora
        return self.normalize_value_conservative(cell.value)
    
    def get_semantic_equivalent(self, text: str) -> str:
        """
        Obtém equivalente semântico de um texto.
        
        Args:
            text: Texto a verificar
            
        Returns:
            Equivalente semântico ou texto original
        """
        text_clean = str(text).strip()
        return self.semantic_equivalence.get(text_clean, text_clean)
    
    def is_header_cell(self, cell, file_type: str = 'published') -> bool:
        """
        Verifica se célula é cabeçalho baseado na cor de fundo.
        
        Args:
            cell: Célula openpyxl
            file_type: 'published' ou 'recreated'
            
        Returns:
            True se for cabeçalho
        """
        if not hasattr(cell, 'fill') or not cell.fill:
            return False
        
        fill_color = None
        if hasattr(cell.fill, 'start_color') and cell.fill.start_color:
            fill_color = cell.fill.start_color.rgb
        
        if not fill_color or fill_color == 'FF000000':  # Sem cor ou preto
            return False
        
        # Verifica cores de cabeçalho
        if file_type == 'published':
            return any(fill_color.startswith(color) for color in self.header_colors['blue_dark'])
        else:  # recreated
            return any(fill_color.startswith(color) for color in self.header_colors['blue_light'])
    
    def detect_spacing_hierarchy(self, cell_value: str, ws, row: int, col: int) -> int:
        """
        Detecta nível hierárquico baseado no espaçamento visual.
        
        Args:
            cell_value: Valor da célula
            ws: Worksheet
            row: Linha
            col: Coluna
            
        Returns:
            Nível hierárquico (0 = top level, 1+ = sub-levels)
        """
        # Verifica se há indentação no valor
        if isinstance(cell_value, str):
            leading_spaces = len(cell_value) - len(cell_value.lstrip())
            if leading_spaces > 0:
                return 1  # Sub-nível
        
        # Verifica posição relativa (células mescladas indicam hierarquia)
        for merged_range in ws.merged_cells.ranges:
            if (merged_range.min_row <= row <= merged_range.max_row and
                merged_range.min_col <= col <= merged_range.max_col):
                # Se está numa área mesclada, é provável que seja nível superior
                return 0
        
        return 0  # Nível superior por padrão
    
    def extract_hierarchical_coordinates(self, ws, data_row: int, data_col: int, 
                                       file_type: str = 'published') -> Dict[str, str]:
        """
        Extrai coordenadas hierárquicas de uma célula de dados.
        
        Args:
            ws: Worksheet
            data_row: Linha da célula de dados
            data_col: Coluna da célula de dados
            file_type: Tipo do ficheiro
            
        Returns:
            Dicionário com coordenadas hierárquicas
        """
        coordinates = {
            'column_level_1': '',
            'column_level_2': '',
            'row_level_1': '',
            'row_level_2': ''
        }
        
        # Extrai coordenadas de coluna (olha para cima)
        col_headers = []
        for check_row in range(max(1, data_row - 10), data_row):
            cell = ws.cell(row=check_row, column=data_col)
            if cell.value and self.is_header_cell(cell, file_type):
                hierarchy_level = self.detect_spacing_hierarchy(str(cell.value), ws, check_row, data_col)
                col_headers.append((hierarchy_level, str(cell.value).strip()))
        
        # Ordena por nível hierárquico
        col_headers.sort(key=lambda x: x[0])
        if len(col_headers) >= 1:
            coordinates['column_level_1'] = col_headers[0][1]
        if len(col_headers) >= 2:
            coordinates['column_level_2'] = col_headers[1][1]
        
        # Extrai coordenadas de linha (olha para esquerda)
        row_headers = []
        for check_col in range(max(1, data_col - 10), data_col):
            cell = ws.cell(row=data_row, column=check_col)
            if cell.value and self.is_header_cell(cell, file_type):
                hierarchy_level = self.detect_spacing_hierarchy(str(cell.value), ws, data_row, check_col)
                row_headers.append((hierarchy_level, str(cell.value).strip()))
        
        # Ordena por nível hierárquico
        row_headers.sort(key=lambda x: x[0])
        if len(row_headers) >= 1:
            coordinates['row_level_1'] = row_headers[0][1]
        if len(row_headers) >= 2:
            coordinates['row_level_2'] = row_headers[1][1]
        
        return coordinates
    
    def extract_data_points(self, file_path: str, sheet_name: str, 
                           file_type: str = 'published') -> List[DataPoint]:
        """
        Extrai pontos de dados de uma folha Excel com coordenadas hierárquicas.
        Usa detecção inteligente da tabela principal e mapeamento avançado.
        
        Args:
            file_path: Caminho do ficheiro
            sheet_name: Nome da folha
            file_type: Tipo do ficheiro
            
        Returns:
            Lista de pontos de dados
        """
        data_points = []
        
        try:
            wb = load_workbook(file_path, data_only=True)
            ws = wb[sheet_name]
            
            # Estratégia 1: Detecção inteligente da tabela principal
            table_bounds = self.find_data_table(ws)
            min_row, max_row, min_col, max_col = table_bounds
            
            self.logger.debug(f"Tabela detectada em {file_path}[{sheet_name}]: "
                            f"linhas {min_row}-{max_row}, colunas {min_col}-{max_col}")
            
            # Extrai dados apenas da área da tabela principal
            for row in range(min_row, max_row + 1):
                for col in range(min_col, max_col + 1):
                    cell = ws.cell(row=row, column=col)
                    
                    # Usa valor apresentado (com formatação)
                    displayed_value = self.get_displayed_value(cell)
                    
                    if displayed_value is not None:
                        # Mapeia dimensões usando algoritmo "andar para cima/esquerda"
                        dimensions = self.get_cell_dimensions(ws, row, col, table_bounds)
                        
                        # Aplica equivalência semântica aos cabeçalhos
                        col_headers = [self.get_semantic_equivalent(h) for h in dimensions['column_headers']]
                        row_headers = [self.get_semantic_equivalent(h) for h in dimensions['row_headers']]
                        
                        data_point = DataPoint(
                            column_level_1=col_headers[0] if len(col_headers) > 0 else "",
                            column_level_2=col_headers[1] if len(col_headers) > 1 else "",
                            row_level_1=row_headers[0] if len(row_headers) > 0 else "",
                            row_level_2=row_headers[1] if len(row_headers) > 1 else "",
                            value=displayed_value,
                            row=row,
                            col=col
                        )
                        
                        data_points.append(data_point)
            
            # Estratégia 2: Fallback para varredura completa se poucos dados encontrados
            if len(data_points) < 10:
                self.logger.warning(f"Poucos dados encontrados com detecção inteligente ({len(data_points)}), "
                                  f"usando varredura completa como fallback")
                
                data_points = []  # Reset
                
                for row in range(1, min(ws.max_row + 1, 100)):
                    for col in range(1, min(ws.max_column + 1, 30)):
                        cell = ws.cell(row=row, column=col)
                        
                        normalized_value = self.normalize_value_conservative(cell.value)
                        
                        if normalized_value is not None:
                            # Usa extração hierárquica simples como fallback
                            coordinates = self.extract_hierarchical_coordinates(ws, row, col, file_type)
                            
                            data_point = DataPoint(
                                column_level_1=self.get_semantic_equivalent(coordinates['column_level_1']),
                                column_level_2=self.get_semantic_equivalent(coordinates['column_level_2']),
                                row_level_1=self.get_semantic_equivalent(coordinates['row_level_1']),
                                row_level_2=self.get_semantic_equivalent(coordinates['row_level_2']),
                                value=normalized_value,
                                row=row,
                                col=col
                            )
                            
                            data_points.append(data_point)
            
            wb.close()
            
        except Exception as e:
            self.logger.error(f"Erro ao extrair dados de {file_path}[{sheet_name}]: {e}")
        
        self.logger.info(f"Extraídos {len(data_points)} pontos de dados de {file_path}[{sheet_name}] ({file_type})")
        return data_points
    
    def fuzzy_match_dimension(self, target: str, candidates: List[str], threshold: float = 0.85) -> Optional[str]:
        """
        Encontra a melhor correspondência difusa para uma dimensão.
        
        Args:
            target: Dimensão alvo
            candidates: Lista de candidatos
            threshold: Limiar mínimo de similaridade
            
        Returns:
            Melhor correspondência ou None
        """
        from difflib import SequenceMatcher
        
        if not target.strip():
            return None
        
        best_match = None
        best_score = 0
        
        target_norm = target.strip().lower()
        
        for candidate in candidates:
            if not candidate.strip():
                continue
            
            candidate_norm = candidate.strip().lower()
            
            # Correspondência exata tem prioridade máxima
            if target_norm == candidate_norm:
                return candidate
            
            # Calcula similaridade
            score = SequenceMatcher(None, target_norm, candidate_norm).ratio()
            
            if score > best_score and score >= threshold:
                best_score = score
                best_match = candidate
        
        return best_match

    def smart_coordinate_match(self, point1: DataPoint, point2: DataPoint) -> bool:
        """
        Verifica se duas coordenadas correspondem semanticamente.
        Versão simplificada e robusta focada no que funciona.
        
        Args:
            point1: Primeiro ponto de dados
            point2: Segundo ponto de dados
            
        Returns:
            True se as coordenadas correspondem
        """
        # Função para verificar se qualquer dimensão match
        def dimensions_match(dims1, dims2):
            # Aplica equivalência semântica em todas as dimensões
            all_dims1 = []
            all_dims2 = []
            
            for dim in dims1:
                if dim and str(dim).strip():
                    all_dims1.append(self.get_semantic_equivalent(str(dim).strip()).lower())
            
            for dim in dims2:
                if dim and str(dim).strip():
                    all_dims2.append(self.get_semantic_equivalent(str(dim).strip()).lower())
            
            # Se não tem dimensões, considera match (dados sem cabeçalhos)
            if not all_dims1 and not all_dims2:
                return True
            
            # Verifica se qualquer dimensão corresponde
            for d1 in all_dims1:
                for d2 in all_dims2:
                    # Correspondência exata
                    if d1 == d2:
                        return True
                    
                    # Correspondência fuzzy para casos similares
                    from difflib import SequenceMatcher
                    similarity = SequenceMatcher(None, d1, d2).ratio()
                    if similarity >= 0.85:  # 85% de similaridade
                        return True
            
            # Casos especiais para totais/vazios
            total_keywords = ['total', '(em branco)', 'em branco', '']
            has_total1 = any(d in total_keywords for d in all_dims1)
            has_total2 = any(d in total_keywords for d in all_dims2)
            
            if has_total1 and has_total2:
                return True
            
            return False
        
        # Verifica correspondência de colunas
        col_dims1 = [point1.column_level_1, point1.column_level_2]
        col_dims2 = [point2.column_level_1, point2.column_level_2]
        col_match = dimensions_match(col_dims1, col_dims2)
        
        # Verifica correspondência de linhas
        row_dims1 = [point1.row_level_1, point1.row_level_2]
        row_dims2 = [point2.row_level_1, point2.row_level_2]
        row_match = dimensions_match(row_dims1, row_dims2)
        
        return col_match and row_match
    
    def compare_data_points(self, published_points: List[DataPoint], 
                           recreated_points: List[DataPoint], 
                           sheet_name: str) -> Dict[str, Any]:
        """
        Compara listas de pontos de dados com lógica inteligente.
        
        Args:
            published_points: Pontos do ficheiro publicado
            recreated_points: Pontos do ficheiro recriado
            sheet_name: Nome da folha
            
        Returns:
            Resultados da comparação
        """
        # Converte para mapas para busca eficiente
        published_map = {point.get_coordinate_key(): point for point in published_points}
        recreated_map = {point.get_coordinate_key(): point for point in recreated_points}
        
        matches = []
        value_differences = []
        missing_in_published = []
        missing_in_recreated = []
        
        # Compara pontos recriados com publicados
        for rec_point in recreated_points:
            # Procura correspondência exata primeiro
            exact_match = published_map.get(rec_point.get_coordinate_key())
            
            if exact_match:
                # Correspondência exata encontrada
                value_diff = abs(rec_point.value - exact_match.value)
                
                if value_diff <= self.numeric_tolerance:
                    matches.append({
                        'recreated': rec_point,
                        'published': exact_match,
                        'match_type': 'exact',
                        'value_difference': value_diff
                    })
                else:
                    value_differences.append({
                        'recreated': rec_point,
                        'published': exact_match,
                        'match_type': 'exact',
                        'value_difference': value_diff
                    })
            else:
                # Procura correspondência semântica
                semantic_match = None
                for pub_point in published_points:
                    if self.smart_coordinate_match(rec_point, pub_point):
                        semantic_match = pub_point
                        break
                
                if semantic_match:
                    value_diff = abs(rec_point.value - semantic_match.value)
                    
                    if value_diff <= self.numeric_tolerance:
                        matches.append({
                            'recreated': rec_point,
                            'published': semantic_match,
                            'match_type': 'semantic',
                            'value_difference': value_diff
                        })
                    else:
                        value_differences.append({
                            'recreated': rec_point,
                            'published': semantic_match,
                            'match_type': 'semantic',
                            'value_difference': value_diff
                        })
                else:
                    missing_in_published.append(rec_point)
        
        # Procura valores publicados não encontrados nos recriados
        for pub_point in published_points:
            found = False
            
            # Verifica correspondência exata
            if recreated_map.get(pub_point.get_coordinate_key()):
                found = True
            else:
                # Verifica correspondência semântica
                for rec_point in recreated_points:
                    if self.smart_coordinate_match(pub_point, rec_point):
                        found = True
                        break
            
            if not found:
                missing_in_recreated.append(pub_point)
        
        results = {
            'sheet_name': sheet_name,
            'published_points': len(published_points),
            'recreated_points': len(recreated_points),
            'exact_matches': len([m for m in matches if m['match_type'] == 'exact']),
            'semantic_matches': len([m for m in matches if m['match_type'] == 'semantic']),
            'value_differences': len(value_differences),
            'missing_in_published': len(missing_in_published),
            'missing_in_recreated': len(missing_in_recreated),
            'matches': matches,
            'differences': value_differences,
            'missing_published': missing_in_published,
            'missing_recreated': missing_in_recreated
        }
        
        return results
    
    def compare_files(self, published_file: str, recreated_file: str, 
                     sheet_names: List[str]) -> Dict[str, Any]:
        """
        Compara dois ficheiros Excel simplificado.
        
        Args:
            published_file: Caminho do ficheiro publicado
            recreated_file: Caminho do ficheiro recriado
            sheet_names: Lista de folhas a comparar
            
        Returns:
            Resultados da comparação
        """
        self.logger.info(f"Iniciando comparação hierárquica: {published_file} vs {recreated_file}")
        
        results = {
            'published_file': published_file,
            'recreated_file': recreated_file,
            'timestamp': datetime.now().isoformat(),
            'sheets': {},
            'summary': {}
        }
        
        total_published = 0
        total_recreated = 0
        total_matches = 0
        total_differences = 0
        total_missing_pub = 0
        total_missing_rec = 0
        
        for sheet_name in sheet_names:
            self.logger.info(f"Comparando folha: {sheet_name}")
            
            # Extrai pontos de dados
            published_points = self.extract_data_points(published_file, sheet_name, 'published')
            recreated_points = self.extract_data_points(recreated_file, sheet_name, 'recreated')
            
            # Compara pontos
            sheet_results = self.compare_data_points(published_points, recreated_points, sheet_name)
            
            results['sheets'][sheet_name] = sheet_results
            
            # Atualiza totais
            total_published += sheet_results['published_points']
            total_recreated += sheet_results['recreated_points']
            total_matches += sheet_results['exact_matches'] + sheet_results['semantic_matches']
            total_differences += sheet_results['value_differences']
            total_missing_pub += sheet_results['missing_in_published']
            total_missing_rec += sheet_results['missing_in_recreated']
            
            self.logger.info(f"Folha {sheet_name}: {sheet_results['published_points']} pub, "
                           f"{sheet_results['recreated_points']} rec, "
                           f"{total_matches} matches, {total_differences} diffs")
        
        # Calcula resumo
        accuracy = (total_matches / max(1, total_recreated)) * 100 if total_recreated > 0 else 0
        
        results['summary'] = {
            'total_published_points': total_published,
            'total_recreated_points': total_recreated,
            'total_matches': total_matches,
            'total_differences': total_differences,
            'total_missing_in_published': total_missing_pub,
            'total_missing_in_recreated': total_missing_rec,
            'accuracy_percentage': accuracy
        }
        
        return results
    
    def copy_worksheet_with_formatting(self, source_ws, target_wb, target_name: str):
        """
        Copia uma worksheet preservando toda a formatação original.
        
        Args:
            source_ws: Worksheet fonte
            target_wb: Workbook destino
            target_name: Nome da nova worksheet
            
        Returns:
            Nova worksheet copiada
        """
        from copy import copy
        
        # Cria nova worksheet
        target_ws = target_wb.create_sheet(title=target_name)
        
        # Copia todas as células com formatação
        for row in source_ws.iter_rows():
            for cell in row:
                new_cell = target_ws.cell(row=cell.row, column=cell.column, value=cell.value)
                
                # Copia formatação se existir
                if cell.has_style:
                    new_cell.font = copy(cell.font)
                    new_cell.border = copy(cell.border)
                    new_cell.fill = copy(cell.fill)
                    new_cell.number_format = cell.number_format
                    new_cell.protection = copy(cell.protection)
                    new_cell.alignment = copy(cell.alignment)
        
        # Copia células mescladas
        for merged_range in source_ws.merged_cells.ranges:
            target_ws.merge_cells(str(merged_range))
        
        # Copia dimensões de colunas
        for col_letter, col_dimension in source_ws.column_dimensions.items():
            target_ws.column_dimensions[col_letter].width = col_dimension.width
        
        # Copia dimensões de linhas
        for row_num, row_dimension in source_ws.row_dimensions.items():
            target_ws.row_dimensions[row_num].height = row_dimension.height
        
        return target_ws

    def create_highlighted_report_sheet(self, source_file: str, source_sheet: str, 
                                      target_wb, sheet_name: str, 
                                      discrepancies: List[Dict[str, Any]]):
        """
        Cria folha de relatório com destaques visuais das discrepâncias.
        
        Args:
            source_file: Ficheiro fonte
            source_sheet: Nome da folha fonte
            target_wb: Workbook destino
            sheet_name: Nome da nova folha
            discrepancies: Lista de discrepâncias a destacar
        """
        # Carrega ficheiro fonte
        source_wb = load_workbook(source_file)
        source_ws = source_wb[source_sheet]
        
        # Copia com formatação
        target_ws = self.copy_worksheet_with_formatting(source_ws, target_wb, sheet_name)
        
        # Destaca discrepâncias em amarelo
        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        
        for discrepancy in discrepancies:
            if 'recreated' in discrepancy:
                row = discrepancy['recreated'].row
                col = discrepancy['recreated'].col
                
                # Aplica destaque amarelo
                cell = target_ws.cell(row=row, column=col)
                cell.fill = yellow_fill
                
                # Adiciona comentário explicativo
                comment_text = (
                    f"DISCREPÂNCIA DETECTADA\n"
                    f"Valor recriado: {discrepancy['recreated'].value}\n"
                    f"Valor publicado: {discrepancy.get('published', {}).get('value', 'N/A')}\n"
                    f"Diferença: {discrepancy.get('value_difference', 'N/A')}\n"
                    f"Tipo: {discrepancy.get('match_type', 'N/A')}"
                )
                
                from openpyxl.comments import Comment
                cell.comment = Comment(comment_text, "Sistema Comparação")
        
        source_wb.close()
        return target_ws

    def generate_report(self, results: Dict[str, Any], output_dir: str = "result/comparison") -> str:
        """
        Gera relatório completo de comparação com destaques visuais.
        
        Args:
            results: Resultados da comparação
            output_dir: Diretório de saída
            
        Returns:
            Caminho do ficheiro gerado
        """
        ensure_directory_exists(output_dir)
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        report_file = os.path.join(output_dir, f"visual_comparison_report_{timestamp}.xlsx")
        
        wb = Workbook()
        
        # Folha de resumo
        ws = wb.active
        ws.title = "Resumo_Geral"
        
        # Cabeçalho
        ws['A1'] = "RELATÓRIO VISUAL DE COMPARAÇÃO DE DADOS"
        ws['A1'].font = Font(bold=True, size=16, color="0070C0")
        ws.merge_cells('A1:F1')
        
        row = 3
        
        # Informações dos ficheiros
        info_style = Font(bold=True, color="2F5597")
        ws[f'A{row}'] = "📁 INFORMAÇÕES DOS FICHEIROS"
        ws[f'A{row}'].font = info_style
        row += 1
        
        ws[f'A{row}'] = "Ficheiro Publicado:"
        ws[f'B{row}'] = os.path.basename(results['published_file'])
        row += 1
        
        ws[f'A{row}'] = "Ficheiro Recriado:"
        ws[f'B{row}'] = os.path.basename(results['recreated_file'])
        row += 1
        
        ws[f'A{row}'] = "Data da Comparação:"
        ws[f'B{row}'] = results['timestamp']
        row += 2
        
        # Estatísticas principais
        ws[f'A{row}'] = "📊 ESTATÍSTICAS GERAIS"
        ws[f'A{row}'].font = info_style
        row += 1
        
        summary = results['summary']
        
        # Cria tabela de estatísticas
        stats_headers = ['Métrica', 'Valor', 'Descrição']
        for i, header in enumerate(stats_headers, 1):
            cell = ws.cell(row=row, column=i, value=header)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        row += 1
        
        stats_data = [
            ("Pontos Publicados", summary['total_published_points'], "Total de valores no ficheiro original"),
            ("Pontos Recriados", summary['total_recreated_points'], "Total de valores no ficheiro recriado"),
            ("Correspondências", summary['total_matches'], "Valores que correspondem exactamente"),
            ("Diferenças", summary['total_differences'], "Valores com discrepâncias"),
            ("Faltam no Publicado", summary['total_missing_in_published'], "Valores só no ficheiro recriado"),
            ("Faltam no Recriado", summary['total_missing_in_recreated'], "Valores só no ficheiro publicado"),
            ("Precisão (%)", f"{summary['accuracy_percentage']:.2f}%", "Percentagem de valores correctos")
        ]
        
        for metric, value, description in stats_data:
            ws.cell(row=row, column=1, value=metric)
            ws.cell(row=row, column=2, value=value)
            ws.cell(row=row, column=3, value=description)
            
            # Destaca linha de precisão
            if "Precisão" in metric:
                for col in range(1, 4):
                    cell = ws.cell(row=row, column=col)
                    cell.fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
                    cell.font = Font(bold=True)
            
            row += 1
        
        row += 1
        
        # Resumo por folha
        ws[f'A{row}'] = "📄 RESUMO POR FOLHA"
        ws[f'A{row}'].font = info_style
        row += 1
        
        sheet_headers = ['Folha', 'Publicados', 'Recriados', 'Matches', 'Diferenças', 'Precisão (%)']
        for i, header in enumerate(sheet_headers, 1):
            cell = ws.cell(row=row, column=i, value=header)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
        row += 1
        
        for sheet_name, sheet_data in results['sheets'].items():
            sheet_matches = sheet_data['exact_matches'] + sheet_data['semantic_matches']
            sheet_accuracy = (sheet_matches / max(1, sheet_data['recreated_points'])) * 100
            
            values = [
                sheet_name,
                sheet_data['published_points'],
                sheet_data['recreated_points'],
                sheet_matches,
                sheet_data['value_differences'],
                f"{sheet_accuracy:.2f}%"
            ]
            
            for i, value in enumerate(values, 1):
                cell = ws.cell(row=row, column=i, value=value)
                
                # Destaca folhas com problemas
                if i == 6 and sheet_accuracy < 95:  # Precisão baixa
                    cell.fill = PatternFill(start_color="FFD9D9", end_color="FFD9D9", fill_type="solid")
                    cell.font = Font(bold=True, color="C5504B")
            
            row += 1
        
        # Ajusta larguras das colunas
        column_widths = [20, 15, 15, 12, 12, 15]
        for i, width in enumerate(column_widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = width
        
        # Cria folhas de detalhes por folha
        for sheet_name, sheet_data in results['sheets'].items():
            if sheet_data['value_differences'] > 0:
                # Cria folha destacada com discrepâncias
                highlighted_sheet_name = f"Resumo_{sheet_name}"
                self.create_highlighted_report_sheet(
                    results['recreated_file'], sheet_name, wb, highlighted_sheet_name,
                    sheet_data['differences']
                )
                
                # Cria folha de detalhes das discrepâncias
                details_sheet_name = f"Detalhes_{sheet_name}"
                details_ws = wb.create_sheet(title=details_sheet_name)
                
                # Cabeçalho da folha de detalhes
                details_ws['A1'] = f"DETALHES DAS DISCREPÂNCIAS - {sheet_name}"
                details_ws['A1'].font = Font(bold=True, size=14)
                details_ws.merge_cells('A1:H1')
                
                # Cabeçalhos da tabela
                headers = ['Linha', 'Coluna', 'Valor Recriado', 'Valor Publicado', 'Diferença', 
                          'Coord. Coluna', 'Coord. Linha', 'Tipo Match']
                
                for i, header in enumerate(headers, 1):
                    cell = details_ws.cell(row=3, column=i, value=header)
                    cell.font = Font(bold=True)
                    cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
                
                # Dados das discrepâncias
                row = 4
                for diff in sheet_data['differences']:
                    rec_point = diff['recreated']
                    pub_point = diff.get('published')
                    
                    values = [
                        rec_point.row,
                        get_column_letter(rec_point.col),
                        rec_point.value,
                        pub_point.value if pub_point else "N/A",
                        diff.get('value_difference', 'N/A'),
                        f"{rec_point.column_level_1}|{rec_point.column_level_2}",
                        f"{rec_point.row_level_1}|{rec_point.row_level_2}",
                        diff.get('match_type', 'N/A')
                    ]
                    
                    for i, value in enumerate(values, 1):
                        details_ws.cell(row=row, column=i, value=value)
                    
                    row += 1
                
                # Ajusta larguras
                for i in range(1, 9):
                    details_ws.column_dimensions[get_column_letter(i)].width = 15
        
        # Folha de informações técnicas
        tech_ws = wb.create_sheet(title="Info_Tecnica")
        tech_ws['A1'] = "INFORMAÇÕES TÉCNICAS DA COMPARAÇÃO"
        tech_ws['A1'].font = Font(bold=True, size=14)
        
        tech_info = [
            ("Tolerância Numérica", self.numeric_tolerance),
            ("Equivalências Semânticas", len(self.semantic_equivalence)),
            ("Algoritmo de Detecção", "Hierárquico com Múltiplas Estratégias"),
            ("Versão do Sistema", "2.0 - Visual Enhanced"),
            ("Capacidades", "Detecção Visual, Múltiplos Fallbacks, Relatórios Interactivos")
        ]
        
        row = 3
        for label, value in tech_info:
            tech_ws.cell(row=row, column=1, value=label)
            tech_ws.cell(row=row, column=2, value=value)
            row += 1
        
        wb.save(report_file)
        self.logger.info(f"Relatório visual completo gerado: {report_file}")
        
        return report_file
    
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
                        print(f"{Fore.RED}Nenhuma folha válida selecionada.{Style.RESET_ALL}")
                except ValueError:
                    print(f"{Fore.RED}Formato inválido. Use números separados por vírgula.{Style.RESET_ALL}")


# Mantém compatibilidade com código existente
DataComparator = HierarchicalDataComparator


def run_interactive_comparison(logger: logging.Logger):
    """
    Executa comparação interativa com o novo sistema hierárquico.
    
    Args:
        logger: Logger configurado
    """
    from colorama import Fore, Style
    
    comparator = HierarchicalDataComparator(logger)
    
    print(f"\n{Fore.GREEN}[Comparação Hierárquica de Dados Excel]{Style.RESET_ALL}")
    print("Sistema redesenhado para estruturas hierárquicas com equivalência semântica.\n")
    
    # Seleção de ficheiros
    published_files, recreated_files = comparator.get_available_files()
    
    if not published_files or not recreated_files:
        print(f"{Fore.RED}Ficheiros insuficientes para comparação.{Style.RESET_ALL}")
        return
    
    # Interface simplificada para seleção
    published_file, recreated_file = comparator.select_files_interactively()
    
    if not published_file or not recreated_file:
        return
    
    # Seleção de folhas
    sheet_names = comparator.select_sheets_interactively(published_file, recreated_file)
    
    if not sheet_names:
        return
    
    # Executa comparação
    print(f"\n{Fore.CYAN}Executando comparação hierárquica...{Style.RESET_ALL}")
    
    results = comparator.compare_files(published_file, recreated_file, sheet_names)
    
    # Mostra resumo
    summary = results['summary']
    print(f"\n{Fore.GREEN}RESULTADOS DA COMPARAÇÃO:{Style.RESET_ALL}")
    print(f"Pontos publicados: {summary['total_published_points']}")
    print(f"Pontos recriados: {summary['total_recreated_points']}")
    print(f"Correspondências: {summary['total_matches']}")
    print(f"Diferenças: {summary['total_differences']}")
    print(f"Precisão: {summary['accuracy_percentage']:.2f}%")
    
    # Gera relatório
    report_file = comparator.generate_report(results)
    print(f"\n{Fore.CYAN}Relatório gerado: {report_file}{Style.RESET_ALL}")
    
    input("\nPressione Enter para continuar...")