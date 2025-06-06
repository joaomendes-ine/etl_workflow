#!/usr/bin/env python3
"""
Módulo de Consolidação Interativa de Dimensões

Este módulo implementa a funcionalidade interativa para consolidação de colunas
de dimensão com controlo total do utilizador sobre o processo de consolidação.

Características principais:
- Interface interativa com sintaxe de pipe (|)
- Preservação da lógica "último valor" hierárquica
- Geração inteligente de nomes de colunas
- Validação completa de integridade de dados
- Compatibilidade com qualquer ficheiro Excel com estrutura standard
"""

import pandas as pd
import numpy as np
import re
from typing import List, Dict, Set, Tuple, Any
import os
from datetime import datetime
import logging
from colorama import Fore, Style

# Importa funções do módulo consolidate_dimensions.py
from src.consolidate_dimensions import (
    clean_total_values,
    remove_exact_duplicates,
    validate_valor_integrity,
    generate_summary_report
)

class InteractiveConsolidator:
    """
    Classe responsável pela consolidação interativa de dimensões.
    Permite ao utilizador controlar totalmente o processo de consolidação.
    """
    
    def __init__(self, input_file: str, output_dir: str, logger: logging.Logger = None):
        """
        Inicializa o consolidador interativo.
        
        Args:
            input_file: Caminho do ficheiro Excel de entrada
            output_dir: Diretório de saída
            logger: Logger configurado
        """
        self.input_file = input_file
        self.output_dir = output_dir
        self.logger = logger or logging.getLogger(__name__)
        self.df_original = None
        self.df_consolidated = None
        self.consolidation_mapping = {}
        self.consolidation_plan = []
        self.final_column_order = []
        
        # Valida inputs
        self._validate_inputs()
    
    def _validate_inputs(self):
        """Valida os inputs fornecidos"""
        if not os.path.exists(self.input_file):
            raise FileNotFoundError(f"Ficheiro de entrada não encontrado: {self.input_file}")
        
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)
            self.logger.info(f"Diretório de saída criado: {self.output_dir}")
    
    def load_and_analyze_data(self) -> bool:
        """
        Carrega e analisa o ficheiro Excel.
        
        Returns:
            True se o carregamento foi bem-sucedido
        """
        try:
            self.logger.info(f"Carregando dados de: {self.input_file}")
            self.df_original = pd.read_excel(self.input_file, engine='openpyxl')
            
            if self.df_original.empty:
                print(f"{Fore.RED}Erro:{Style.RESET_ALL} O ficheiro Excel está vazio")
                return False
            
            self.logger.info(f"Dados carregados: {len(self.df_original)} linhas, {len(self.df_original.columns)} colunas")
            return True
            
        except Exception as e:
            self.logger.error(f"Erro ao carregar ficheiro: {str(e)}")
            print(f"{Fore.RED}Erro ao carregar ficheiro:{Style.RESET_ALL} {str(e)}")
            return False
    
    def display_dimensions(self) -> List[str]:
        """
        Apresenta as dimensões disponíveis e retorna a lista.
        
        Returns:
            Lista de colunas de dimensão
        """
        dim_columns = [col for col in self.df_original.columns if col.startswith('dim_')]
        
        if not dim_columns:
            print(f"{Fore.YELLOW}Aviso:{Style.RESET_ALL} Nenhuma coluna de dimensão encontrada (colunas que começam com 'dim_')")
            return []
        
        print(f"\n{Fore.CYAN}=== DIMENSÕES DISPONÍVEIS ==={Style.RESET_ALL}")
        
        for i, dim in enumerate(dim_columns, 1):
            # Conta valores únicos (excluindo vazios e 'Total')
            unique_values = self.df_original[dim].dropna()
            unique_values = unique_values[unique_values.astype(str).str.lower() != 'total']
            unique_count = unique_values.nunique()
            
            # Mostra amostra de valores
            sample_values = unique_values.unique()[:3]
            sample_str = ', '.join(str(v) for v in sample_values)
            if len(sample_values) < unique_count:
                sample_str += "..."
            
            print(f"{Fore.WHITE}{i:2d}.{Style.RESET_ALL} {Fore.GREEN}{dim}{Style.RESET_ALL} "
                  f"({unique_count} valores únicos)")
            if sample_str:
                print(f"     Exemplos: {sample_str}")
        
        return dim_columns
    
    def display_consolidation_instructions(self):
        """Apresenta as instruções de consolidação"""
        print(f"\n{Fore.CYAN}=== INSTRUÇÕES DE CONSOLIDAÇÃO ==={Style.RESET_ALL}")
        print("Use a sintaxe de pipe (|) para definir grupos de consolidação:")
        print("• Use | para separar grupos de consolidação")
        print("• Use , para agrupar dimensões que serão consolidadas")
        print("• Números individuais permanecem como colunas separadas")
        print("\nExemplos:")
        print(f"  {Fore.YELLOW}1 | 2 | 3,4,5{Style.RESET_ALL}      → dim_1 separada | dim_2 separada | dim_3,4,5 consolidadas")
        print(f"  {Fore.YELLOW}1,2 | 3 | 4,5,6{Style.RESET_ALL}    → dim_1,2 consolidadas | dim_3 separada | dim_4,5,6 consolidadas")
        print(f"  {Fore.YELLOW}1 | 2 | 3 | 4{Style.RESET_ALL}      → todas permanecem separadas")
        print(f"  {Fore.YELLOW}auto{Style.RESET_ALL}              → detecção automática de padrões")
        print("\nA ordem dos grupos define a ordem final das colunas.")
    
    def get_user_consolidation_input(self, dim_columns: List[str]) -> str:
        """
        Obtém a entrada do utilizador para consolidação.
        
        Args:
            dim_columns: Lista de colunas de dimensão
            
        Returns:
            String com a escolha do utilizador
        """
        while True:
            user_input = input(f"\n{Fore.GREEN}>>{Style.RESET_ALL} Digite sua escolha de consolidação: ").strip()
            
            if not user_input:
                print(f"{Fore.YELLOW}Entrada vazia. Tente novamente.{Style.RESET_ALL}")
                continue
            
            if user_input.lower() == 'auto':
                return user_input
            
            # Validação básica da sintaxe
            if self._validate_input_syntax(user_input, len(dim_columns)):
                return user_input
            else:
                print(f"{Fore.RED}Sintaxe inválida. Tente novamente.{Style.RESET_ALL}")
                continue
    
    def _validate_input_syntax(self, user_input: str, max_dim: int) -> bool:
        """
        Valida a sintaxe da entrada do utilizador.
        
        Args:
            user_input: Entrada do utilizador
            max_dim: Número máximo de dimensões
            
        Returns:
            True se a sintaxe está correta
        """
        try:
            # Divide por pipes
            groups = user_input.split('|')
            
            used_indices = set()
            
            for group_str in groups:
                group_str = group_str.strip()
                if not group_str:
                    continue
                
                # Parse números no grupo
                for num_str in group_str.split(','):
                    num = int(num_str.strip())
                    if num < 1 or num > max_dim:
                        print(f"Número {num} fora do intervalo (1-{max_dim})")
                        return False
                    
                    if num in used_indices:
                        print(f"Número {num} usado mais de uma vez")
                        return False
                    
                    used_indices.add(num)
            
            return True
            
        except ValueError:
            print("Formato inválido. Use apenas números, vírgulas e pipes.")
            return False
    
    def parse_consolidation_input(self, user_input: str, dim_columns: List[str]) -> Tuple[List[Dict], List[str]]:
        """
        Analisa a entrada do utilizador e cria o plano de consolidação.
        
        Args:
            user_input: Entrada do utilizador
            dim_columns: Lista de colunas de dimensão
            
        Returns:
            Tupla com (plano_consolidacao, ordem_colunas_final)
        """
        if user_input.lower() == 'auto':
            return self._auto_detect_consolidation(dim_columns)
        
        consolidation_plan = []
        final_column_order = []
        
        # Divide por pipes para obter grupos
        groups = user_input.split('|')
        
        print(f"\n{Fore.CYAN}=== ANÁLISE DO PLANO DE CONSOLIDAÇÃO ==={Style.RESET_ALL}")
        
        for group_str in groups:
            group_str = group_str.strip()
            if not group_str:
                continue
            
            # Parse números no grupo
            try:
                indices = []
                for num_str in group_str.split(','):
                    num = int(num_str.strip()) - 1  # Converte para índice 0-based
                    if 0 <= num < len(dim_columns):
                        indices.append(num)
                
                if not indices:
                    continue
                
                # Obtém nomes das dimensões
                selected_dims = [dim_columns[i] for i in indices]
                
                if len(selected_dims) == 1:
                    # Dimensão única - mantém como está
                    final_column_order.append(selected_dims[0])
                    print(f"✓ Mantendo '{Fore.GREEN}{selected_dims[0]}{Style.RESET_ALL}' sem consolidação")
                else:
                    # Múltiplas dimensões - consolida
                    consolidated_name = self._generate_smart_dimension_name(selected_dims)
                    
                    # Remove duplicatas preservando ordem
                    unique_dims = list(dict.fromkeys(selected_dims))
                    
                    if len(unique_dims) < len(selected_dims):
                        print(f"⚠️ Dimensões duplicadas removidas no grupo: {selected_dims}")
                        selected_dims = unique_dims
                    
                    consolidation_plan.append({
                        'source_columns': selected_dims,
                        'target_column': consolidated_name
                    })
                    final_column_order.append(consolidated_name)
                    
                    print(f"✓ Consolidando {len(selected_dims)} dimensões → '{Fore.YELLOW}{consolidated_name}{Style.RESET_ALL}'")
                    print(f"  Dimensões: {', '.join(selected_dims)}")
            
            except ValueError as e:
                print(f"❌ Erro ao processar grupo '{group_str}': entrada inválida")
                continue
        
        return consolidation_plan, final_column_order
    
    def _generate_smart_dimension_name(self, columns: List[str]) -> str:
        """
        Gera nome inteligente para coluna consolidada baseado nos padrões das colunas.
        
        Args:
            columns: Lista de colunas a consolidar
            
        Returns:
            Nome limpo e único para a coluna consolidada
        """
        if not columns:
            return "dim_consolidada"
        
        if len(columns) == 1:
            return self._clean_column_name(columns[0])
        
        # Estratégia 1: Verifica se há padrão óbvio com base no mapeamento predefinido
        predefined_mapping = {
            'grupo_etario': ['grupo_etario', 'etario', 'idade'],
            'setor_atividade_economica': ['setor', 'economia', 'cae', 'atividade'],
            'profissao': ['profissao', 'cpp', 'cnp'],
            'situacao_profissional': ['situacao_profissao', 'situacao_profiss'],
            'condicao_trabalho': ['condicao_trabalho', 'trabalho_inativo'],
            'trabalho_casa': ['trabalho_casa'],
            'educacao_formacao': ['educacao', 'formacao'],
            'estado_saude': ['estado_saude', 'limitacoes_saude', 'problemas_saude'],
            'tempo_tarefas_trabalho': ['leitura_manuais', 'calculos', 'trab_arduo', 'tarefas_destreza', 'interacao'],
            'autonomia_trabalho': ['grau_autonomia', 'autonomia_decidir']
        }
        
        # Verifica correspondências com o mapeamento predefinido
        for target_name, keywords in predefined_mapping.items():
            all_columns_text = ' '.join(columns).lower()
            keyword_matches = sum(1 for keyword in keywords if keyword in all_columns_text)
            
            if keyword_matches >= 2:  # Pelo menos 2 palavras-chave correspondem
                clean_name = f"dim_{target_name}"
                self.logger.debug(f"Nome baseado em mapeamento predefinido: '{clean_name}' para {columns}")
                return clean_name
        
        # Estratégia 2: Encontra prefixo comum mais longo
        common_prefix = self._find_longest_common_prefix(columns)
        
        if common_prefix and len(common_prefix) > 8:  # Pelo menos "dim_xxxx"
            clean_name = self._clean_column_name(common_prefix)
            self.logger.debug(f"Nome baseado em prefixo comum: '{clean_name}' para {columns}")
            return clean_name
        
        # Estratégia 3: Extrai palavras-chave mais frequentes
        all_words = []
        for col in columns:
            # Remove 'dim_' e divide em palavras
            base_name = col[4:] if col.startswith('dim_') else col
            words = re.findall(r'\w+', base_name.lower())
            # Filtra palavras muito genéricas
            filtered_words = [w for w in words if w not in ['dim', 'de', 'do', 'da', 'e', 'ou'] and len(w) > 2]
            all_words.extend(filtered_words)
        
        # Conta frequência
        word_counts = {}
        for word in all_words:
            word_counts[word] = word_counts.get(word, 0) + 1
        
        # Pega as 2-3 palavras mais comuns
        common_words = sorted(word_counts.items(), key=lambda x: x[1], reverse=True)[:3]
        
        if common_words:
            keywords = [word for word, count in common_words if count > 1]  # Aparecem em múltiplas colunas
            if keywords:
                clean_name = f"dim_{'_'.join(keywords[:2])}"  # Máximo 2 palavras
                self.logger.debug(f"Nome baseado em palavras-chave: '{clean_name}' para {columns}")
                return self._clean_column_name(clean_name)
        
        # Estratégia 4: Fallback - usa primeira coluna com sufixo consolidada
        base_name = columns[0][4:] if columns[0].startswith('dim_') else columns[0]
        # Remove números do final
        base_name = re.sub(r'\d+$', '', base_name).rstrip('_')
        
        if base_name:
            fallback_name = f"dim_{base_name}"
        else:
            fallback_name = f"dim_consolidada_{len(columns)}"
        
        self.logger.debug(f"Nome fallback: '{fallback_name}' para {columns}")
        return self._clean_column_name(fallback_name)
    
    def _find_longest_common_prefix(self, names: List[str]) -> str:
        """Encontra o prefixo comum mais longo entre os nomes"""
        if not names:
            return ""
        
        prefix = names[0]
        for name in names[1:]:
            common = ""
            for c1, c2 in zip(prefix, name):
                if c1 == c2:
                    common += c1
                else:
                    break
            prefix = common
        
        return prefix.rstrip('_')
    
    def _clean_column_name(self, name: str) -> str:
        """
        Limpa e normaliza o nome da coluna.
        
        Args:
            name: Nome original
            
        Returns:
            Nome limpo
        """
        # Converte para minúsculas
        name = name.lower()
        
        # Remove caracteres especiais, mantém apenas letras, números e underscores
        name = re.sub(r'[^a-z0-9_]', '_', name)
        
        # Remove underscores múltiplos
        name = re.sub(r'_+', '_', name)
        
        # Remove underscores no início e fim
        name = name.strip('_')
        
        # Garante que começa com 'dim_'
        if not name.startswith('dim_'):
            name = f'dim_{name}'
        
        # Remove 'dim_dim_' redundante
        name = re.sub(r'^dim_dim_', 'dim_', name)
        
        # Trunca se muito longo
        if len(name) > 50:
            name = name[:47] + '...'
        
        return name
    
    def _auto_detect_consolidation(self, dim_columns: List[str]) -> Tuple[List[Dict], List[str]]:
        """
        Detecção automática de padrões de consolidação.
        
        Args:
            dim_columns: Lista de colunas de dimensão
            
        Returns:
            Tupla com (plano_consolidacao, ordem_colunas_final)
        """
        print(f"\n{Fore.CYAN}=== DETECÇÃO AUTOMÁTICA DE PADRÕES ==={Style.RESET_ALL}")
        
        # Usa a lógica do DimensionAnalyzer para detectar padrões
        from src.dimension_analyzer import DimensionAnalyzer
        
        analyzer = DimensionAnalyzer(self.df_original, self.logger)
        patterns = analyzer.analyze_patterns()
        
        consolidation_plan = []
        final_column_order = []
        used_columns = set()
        
        if patterns:
            print(f"Padrões detectados: {len(patterns)}")
            
            for pattern_name, columns in patterns.items():
                if len(columns) > 1:
                    # Filtra colunas já usadas
                    available_columns = [col for col in columns if col not in used_columns]
                    
                    if len(available_columns) > 1:
                        consolidated_name = self._generate_smart_dimension_name(available_columns)
                        
                        consolidation_plan.append({
                            'source_columns': available_columns,
                            'target_column': consolidated_name
                        })
                        final_column_order.append(consolidated_name)
                        used_columns.update(available_columns)
                        
                        print(f"✓ Padrão '{pattern_name}': {len(available_columns)} colunas → '{consolidated_name}'")
        
        # Adiciona colunas restantes como individuais
        for col in dim_columns:
            if col not in used_columns:
                final_column_order.append(col)
                print(f"✓ Mantendo '{col}' individual")
        
        return consolidation_plan, final_column_order
    
    def display_consolidation_summary(self, consolidation_plan: List[Dict], final_column_order: List[str], dim_columns: List[str]):
        """
        Apresenta o resumo do plano de consolidação.
        
        Args:
            consolidation_plan: Plano de consolidação
            final_column_order: Ordem final das colunas
            dim_columns: Colunas de dimensão originais
        """
        print(f"\n{Fore.CYAN}=== RESUMO DA CONSOLIDAÇÃO ==={Style.RESET_ALL}")
        print(f"Dimensões originais: {Fore.WHITE}{len(dim_columns)}{Style.RESET_ALL}")
        print(f"Grupos de consolidação: {Fore.WHITE}{len(consolidation_plan)}{Style.RESET_ALL}")
        print(f"Dimensões finais: {Fore.WHITE}{len(final_column_order)}{Style.RESET_ALL}")
        
        if consolidation_plan:
            print(f"\n{Fore.YELLOW}Consolidações a realizar:{Style.RESET_ALL}")
            for group in consolidation_plan:
                print(f"\n• {', '.join(group['source_columns'])}")
                print(f"  → {Fore.GREEN}{group['target_column']}{Style.RESET_ALL}")
        
        # Calcula valores únicos a preservar
        total_unique_values = self._count_all_unique_dimension_values(dim_columns)
        print(f"\n{Fore.CYAN}Valores únicos totais a preservar:{Style.RESET_ALL} {total_unique_values}")
    
    def _count_all_unique_dimension_values(self, dim_columns: List[str]) -> int:
        """
        Conta todos os valores únicos nas dimensões.
        
        Args:
            dim_columns: Lista de colunas de dimensão
            
        Returns:
            Número total de valores únicos
        """
        all_values = set()
        
        for col in dim_columns:
            if col in self.df_original.columns:
                values = self.df_original[col].dropna()
                values = values[values.astype(str).str.lower() != 'total']
                # Filtra valores vazios e 'nan'
                clean_values = values.astype(str)
                clean_values = clean_values[clean_values.str.strip() != '']
                clean_values = clean_values[clean_values.str.lower() != 'nan']
                all_values.update(clean_values.unique())
        
        return len(all_values)
    
    def confirm_consolidation_plan(self) -> bool:
        """
        Confirma o plano de consolidação com o utilizador.
        
        Returns:
            True se o utilizador confirmar
        """
        while True:
            confirm = input(f"\n{Fore.GREEN}>>{Style.RESET_ALL} Confirmar este plano de consolidação? (s/n/r para refazer): ").lower().strip()
            
            if confirm == 's' or confirm == 'sim':
                return True
            elif confirm == 'n' or confirm == 'nao' or confirm == 'não':
                print("Consolidação cancelada.")
                return False
            elif confirm == 'r' or confirm == 'refazer':
                return None  # Indica para refazer
            else:
                print(f"{Fore.YELLOW}Responda com 's' (sim), 'n' (não) ou 'r' (refazer){Style.RESET_ALL}")
    
    def apply_consolidation(self) -> bool:
        """
        Aplica o plano de consolidação com lógica "último valor".
        
        Returns:
            True se a consolidação foi bem-sucedida
        """
        try:
            print(f"\n{Fore.CYAN}=== APLICANDO CONSOLIDAÇÃO ==={Style.RESET_ALL}")
            
            # Cria cópia dos dados originais
            self.df_consolidated = self.df_original.copy()
            
            # Passo 1: Limpa valores 'Total' (reutiliza função existente)
            print("Limpando valores 'Total'...")
            self.df_consolidated = clean_total_values(self.df_consolidated)
            
            # Passo 2: Aplica cada grupo de consolidação
            successfully_consolidated = []
            
            for group_idx, group in enumerate(self.consolidation_plan):
                source_cols = group['source_columns']
                target_col = group['target_column']
                
                try:
                    print(f"\nConsolidando {len(source_cols)} colunas em '{target_col}'...")
                    
                    # Verifica se todas as colunas de origem ainda existem
                    existing_source_cols = [col for col in source_cols if col in self.df_consolidated.columns]
                    
                    if not existing_source_cols:
                        print(f"  ⚠️  Pulando '{target_col}': nenhuma coluna de origem encontrada")
                        continue
                    
                    if len(existing_source_cols) < len(source_cols):
                        missing_cols = [col for col in source_cols if col not in existing_source_cols]
                        print(f"  ⚠️  Algumas colunas de origem não encontradas: {missing_cols}")
                        print(f"  Procedendo com {len(existing_source_cols)} colunas disponíveis")
                    
                    # Verifica se o nome da coluna de destino já existe nas colunas de origem
                    if target_col in existing_source_cols:
                        # Se o nome de destino é uma das colunas de origem, gera nome único
                        base_target = target_col
                        counter = 1
                        while target_col in self.df_consolidated.columns:
                            target_col = f"{base_target}_consolidado_{counter}"
                            counter += 1
                        print(f"  ℹ️  Nome de destino '{base_target}' já existe nas colunas de origem, usando '{target_col}'")
                    elif target_col in self.df_consolidated.columns and target_col not in existing_source_cols:
                        # Gera nome alternativo se existe mas não é uma coluna de origem
                        counter = 1
                        original_target = target_col
                        while target_col in self.df_consolidated.columns:
                            target_col = f"{original_target}_{counter}"
                            counter += 1
                        print(f"  ℹ️  Nome de coluna '{original_target}' já existe, usando '{target_col}'")
                    
                    # Aplica lógica "último valor" hierárquica
                    consolidated_values = []
                    
                    for idx, row in self.df_consolidated.iterrows():
                        # Varre colunas da esquerda para a direita (preserva hierarquia)
                        # Toma o ÚLTIMO valor não-vazio (mais à direita)
                        final_value = ''
                        
                        for col in existing_source_cols:
                            if col in row and pd.notna(row[col]):
                                val = str(row[col]).strip()
                                if val != '' and val != 'nan' and val.lower() != 'total':
                                    final_value = val  # Continua a atualizar - último ganha
                        
                        consolidated_values.append(final_value)
                    
                    # Adiciona nova coluna
                    self.df_consolidated[target_col] = consolidated_values
                    
                    # Remove colunas de origem (apenas as que existem)
                    self.df_consolidated = self.df_consolidated.drop(columns=existing_source_cols)
                    
                    # Regista mapeamento
                    self.consolidation_mapping[target_col] = existing_source_cols
                    
                    # Log de resultados
                    unique_count = self.df_consolidated[target_col].nunique()
                    print(f"✅ Consolidado: {unique_count} valores únicos em '{target_col}'")
                    
                    successfully_consolidated.append(group_idx)
                    
                except Exception as group_error:
                    error_msg = f"Erro na consolidação do grupo '{target_col}': {str(group_error)}"
                    self.logger.error(error_msg)
                    print(f"❌ {error_msg}")
                    print(f"  Colunas de origem: {source_cols}")
                    # Continua com os outros grupos
                    continue
            
            if not successfully_consolidated:
                print(f"{Fore.RED}❌ Nenhuma consolidação foi bem-sucedida{Style.RESET_ALL}")
                return False
            
            print(f"\n✅ {len(successfully_consolidated)}/{len(self.consolidation_plan)} consolidações bem-sucedidas")
            
            # Passo 3: Reordena colunas conforme preferência do utilizador
            self.df_consolidated = self._reorder_columns_by_user_preference()
            
            # Passo 4: Remove duplicatas exatas (reutiliza função existente)
            print("\nRemoção de duplicatas...")
            self.df_consolidated = remove_exact_duplicates(self.df_consolidated)
            
            # Passo 5: Valida integridade (reutiliza função existente)
            print("\nValidação de integridade...")
            try:
                valor_ok = validate_valor_integrity(self.df_original, self.df_consolidated)
                if not valor_ok:
                    print(f"{Fore.YELLOW}Aviso de integridade na coluna 'valor' - revisar recomendado{Style.RESET_ALL}")
            except Exception as integrity_error:
                self.logger.warning(f"Erro na validação de integridade: {str(integrity_error)}")
                print(f"{Fore.YELLOW}⚠️  Erro na validação de integridade: {str(integrity_error)}{Style.RESET_ALL}")
            
            return True
            
        except Exception as e:
            self.logger.error(f"Erro durante consolidação: {str(e)}")
            print(f"{Fore.RED}Erro durante consolidação:{Style.RESET_ALL} {str(e)}")
            print(f"{Fore.RED}Detalhes do erro:{Style.RESET_ALL}")
            import traceback
            traceback.print_exc()
            return False
    
    def _reorder_columns_by_user_preference(self) -> pd.DataFrame:
        """
        Reordena colunas respeitando a preferência do utilizador.
        
        Returns:
            DataFrame com colunas reordenadas
        """
        print("Reordenando colunas conforme preferência do utilizador...")
        
        # Separa tipos de colunas
        current_dims = [col for col in self.df_consolidated.columns if col.startswith('dim_')]
        non_dims = [col for col in self.df_consolidated.columns if not col.startswith('dim_')]
        
        # Constrói ordem final
        final_order = []
        
        # 1. Adiciona dimensões do utilizador na sua ordem
        for dim in self.final_column_order:
            if dim in self.df_consolidated.columns:
                final_order.append(dim)
        
        # 2. Adiciona dimensões restantes
        for dim in current_dims:
            if dim not in final_order:
                final_order.append(dim)
        
        # 3. Adiciona colunas não-dimensão (prioriza importantes primeiro)
        priority_non_dims = ['indicador', 'unidade', 'valor', 'simbologia', 'estado_valor', 'coeficiente_variacao']
        for col in priority_non_dims:
            if col in non_dims:
                final_order.append(col)
                non_dims.remove(col)
        
        # Adiciona restantes
        final_order.extend(non_dims)
        
        print(f"  Reordenadas {len(final_order)} colunas")
        return self.df_consolidated[final_order]
    
    def save_results(self, format_type: str = 'excel', filename: str = None) -> str:
        """
        Guarda os resultados consolidados.
        
        Args:
            format_type: Tipo de formato ('excel', 'csv', 'json')
            filename: Nome do ficheiro (opcional)
            
        Returns:
            Caminho do ficheiro guardado
        """
        if self.df_consolidated is None:
            raise ValueError("Nenhum resultado para guardar. Execute a consolidação primeiro.")
        
        # Gera nome do ficheiro se não fornecido
        if filename is None:
            base_name = os.path.splitext(os.path.basename(self.input_file))[0]
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"{base_name}_consolidado_{timestamp}"
        
        # Define extensão baseada no formato
        if format_type.lower() == 'excel':
            output_file = os.path.join(self.output_dir, f"{filename}.xlsx")
            
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                self.df_consolidated.to_excel(writer, sheet_name='Dados_Consolidados', index=False)
                
                # Adiciona sheet com resumo
                summary_report = self._generate_summary_report()
                summary_df = pd.DataFrame({'Resumo': summary_report.split('\n')})
                summary_df.to_excel(writer, sheet_name='Resumo_Consolidacao', index=False)
                
        elif format_type.lower() == 'csv':
            output_file = os.path.join(self.output_dir, f"{filename}.csv")
            self.df_consolidated.to_csv(output_file, index=False, encoding='utf-8-sig')
            
        elif format_type.lower() == 'json':
            output_file = os.path.join(self.output_dir, f"{filename}.json")
            self.df_consolidated.to_json(output_file, orient='records', indent=2, force_ascii=False)
            
        else:
            raise ValueError(f"Formato não suportado: {format_type}")
        
        self.logger.info(f"Resultados guardados: {output_file}")
        return output_file
    
    def _generate_summary_report(self) -> str:
        """
        Gera relatório de resumo da consolidação.
        
        Returns:
            String com o relatório
        """
        # Reutiliza função existente adaptada
        return generate_summary_report(self.df_original, self.df_consolidated)
    
    def print_summary(self):
        """Imprime resumo da consolidação no console"""
        if self.df_consolidated is None:
            print(f"{Fore.YELLOW}Nenhuma consolidação realizada ainda{Style.RESET_ALL}")
            return
        
        print(f"\n{Fore.CYAN}{'='*60}{Style.RESET_ALL}")
        print(f"{Fore.CYAN}           RESUMO DA CONSOLIDAÇÃO INTERATIVA{Style.RESET_ALL}")
        print(f"{Fore.CYAN}{'='*60}{Style.RESET_ALL}")
        
        original_dims = len([c for c in self.df_original.columns if c.startswith('dim_')])
        final_dims = len([c for c in self.df_consolidated.columns if c.startswith('dim_')])
        reduction = original_dims - final_dims
        reduction_pct = (reduction / original_dims * 100) if original_dims > 0 else 0
        
        print(f"Dimensões originais:     {original_dims}")
        print(f"Dimensões consolidadas:  {final_dims}")
        print(f"Redução:                 {reduction} ({reduction_pct:.1f}%)")
        print(f"Linhas processadas:      {len(self.df_consolidated):,}")
        
        if self.consolidation_mapping:
            print(f"\n{Fore.YELLOW}Consolidações realizadas:{Style.RESET_ALL}")
            for new_col, original_cols in self.consolidation_mapping.items():
                print(f"  • {Fore.GREEN}{new_col}{Style.RESET_ALL} ← {', '.join(original_cols)}")
        
        print(f"\n{Fore.GREEN}✅ Consolidação concluída com sucesso!{Style.RESET_ALL}")
        print(f"{Fore.CYAN}{'='*60}{Style.RESET_ALL}")


def get_input_file() -> str:
    """
    Obtém o ficheiro de entrada do utilizador.
    
    Returns:
        Caminho do ficheiro selecionado
    """
    print(f"\n{Fore.GREEN}[Seleção de Ficheiro]{Style.RESET_ALL} Escolha o ficheiro Excel a processar:\n")
    
    # Opção 1: Ficheiro da pasta principal
    main_files = []
    if os.path.exists("dataset/main"):
        main_files = [f for f in os.listdir("dataset/main") if f.lower().endswith(('.xlsx', '.xls'))]
    
    if main_files:
        print(f"  {Fore.WHITE}1.{Style.RESET_ALL} Ficheiro da pasta principal:")
        for i, file in enumerate(main_files[:5]):  # Mostra até 5 ficheiros
            print(f"     {i+1}. {file}")
        if len(main_files) > 5:
            print(f"     ... e mais {len(main_files) - 5} ficheiros")
    
    print(f"  {Fore.WHITE}2.{Style.RESET_ALL} Outro ficheiro (introduzir caminho)")
    print(f"  {Fore.WHITE}0.{Style.RESET_ALL} Voltar ao menu principal")
    
    choice = input(f"\n{Fore.GREEN}>>{Style.RESET_ALL} Digite a sua escolha: ").strip()
    
    if choice == "0":
        return None
    elif choice == "1" and main_files:
        if len(main_files) == 1:
            return os.path.join("dataset/main", main_files[0])
        else:
            # Se há múltiplos ficheiros, permite escolher
            print(f"\n{Fore.GREEN}Selecione o ficheiro:{Style.RESET_ALL}")
            for i, file in enumerate(main_files, 1):
                print(f"  {i}. {file}")
            
            while True:
                file_choice = input(f"\n{Fore.GREEN}>>{Style.RESET_ALL} Digite o número do ficheiro: ").strip()
                try:
                    file_idx = int(file_choice) - 1
                    if 0 <= file_idx < len(main_files):
                        return os.path.join("dataset/main", main_files[file_idx])
                    else:
                        print(f"{Fore.RED}Número inválido.{Style.RESET_ALL} Tente novamente.")
                except ValueError:
                    print(f"{Fore.RED}Digite um número válido.{Style.RESET_ALL}")
    elif choice == "2":
        custom_path = input(f"{Fore.GREEN}>>{Style.RESET_ALL} Introduza o caminho completo para o ficheiro Excel: ").strip()
        
        if not os.path.exists(custom_path):
            print(f"{Fore.RED}Erro:{Style.RESET_ALL} Ficheiro não encontrado: {custom_path}")
            return None
        
        if not custom_path.lower().endswith(('.xlsx', '.xls')):
            print(f"{Fore.RED}Erro:{Style.RESET_ALL} O ficheiro deve ser Excel (.xlsx ou .xls)")
            return None
        
        return custom_path
    else:
        print(f"{Fore.RED}Opção inválida.{Style.RESET_ALL}")
        return None 