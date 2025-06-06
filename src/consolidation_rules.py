import re
import unicodedata
from typing import Dict, List, Set, Tuple
import logging

class ConsolidationRules:
    """
    Classe com regras e validadores para consolidação de colunas de dimensão.
    Define critérios de compatibilidade e convenções de nomeação.
    """
    
    # Thresholds configuráveis
    MIN_SIMILARITY_THRESHOLD = 0.3
    MAX_OVERLAP_RATIO = 0.8
    MAX_UNIQUE_VALUES = 1000
    
    @staticmethod
    def can_consolidate(columns: List[str], value_sets: Dict[str, Set[str]], 
                       similarity_scores: Dict[Tuple[str, str], float] = None,
                       logger: logging.Logger = None) -> Tuple[bool, List[str], List[str]]:
        """
        Determina se as colunas podem ser consolidadas com segurança.
        
        Args:
            columns: Lista de colunas a consolidar
            value_sets: Dicionário com sets de valores únicos para cada coluna
            similarity_scores: Scores de similaridade entre pares de colunas
            logger: Logger configurado
            
        Returns:
            Tupla com (pode_consolidar, lista_de_razões, lista_de_avisos)
        """
        if logger is None:
            logger = logging.getLogger(__name__)
        
        reasons = []
        warnings = []
        can_consolidate = True
        
        if len(columns) <= 1:
            reasons.append("Necessário pelo menos 2 colunas para consolidação")
            return False, reasons, warnings
        
        # Regra 1: Verificar compatibilidade de valores
        compatibility_check = ConsolidationRules._check_value_compatibility(value_sets, logger)
        if not compatibility_check['compatible']:
            can_consolidate = False
            reasons.extend(compatibility_check['reasons'])
        warnings.extend(compatibility_check['warnings'])
        
        # Regra 2: Verificar similaridade mínima (se disponível)
        if similarity_scores:
            similarity_check = ConsolidationRules._check_similarity_threshold(
                columns, similarity_scores, logger
            )
            if not similarity_check['meets_threshold']:
                can_consolidate = False
                reasons.extend(similarity_check['reasons'])
            warnings.extend(similarity_check['warnings'])
        
        # Regra 3: Verificar se não há conflitos de significado
        semantic_check = ConsolidationRules._check_semantic_conflicts(columns, value_sets, logger)
        if not semantic_check['no_conflicts']:
            warnings.extend(semantic_check['warnings'])
            # Conflitos semânticos são avisos, não impedem consolidação
        
        # Regra 4: Verificar tamanho resultante
        size_check = ConsolidationRules._check_result_size(value_sets, logger)
        if not size_check['acceptable_size']:
            warnings.extend(size_check['warnings'])
            # Tamanho grande é aviso, não impedimento
        
        logger.info(f"Verificação de consolidação para {columns}: {'APROVADA' if can_consolidate else 'REJEITADA'}")
        if reasons:
            logger.info(f"Razões de rejeição: {reasons}")
        if warnings:
            logger.info(f"Avisos: {warnings}")
        
        return can_consolidate, reasons, warnings
    
    @staticmethod
    def _check_value_compatibility(value_sets: Dict[str, Set[str]], logger: logging.Logger) -> Dict[str, any]:
        """Verifica se os valores das colunas são compatíveis para consolidação"""
        all_values = list(value_sets.values())
        
        if not all_values:
            return {
                'compatible': False,
                'reasons': ['Nenhum valor encontrado nas colunas'],
                'warnings': []
            }
        
        # Verifica se há colunas completamente vazias
        empty_columns = [col for col, values in value_sets.items() if not values]
        if empty_columns:
            logger.warning(f"Colunas vazias encontradas: {empty_columns}")
        
        # Calcula sobreposições
        overlaps = []
        total_unique = len(set().union(*all_values)) if all_values else 0
        
        for i, set1 in enumerate(all_values):
            for set2 in all_values[i+1:]:
                if set1 and set2:  # Ignora sets vazios
                    overlap = len(set1.intersection(set2))
                    overlaps.append(overlap)
        
        max_overlap = max(overlaps) if overlaps else 0
        overlap_ratio = max_overlap / total_unique if total_unique > 0 else 0
        
        reasons = []
        warnings = []
        compatible = True
        
        # Verifica sobreposição excessiva
        if overlap_ratio > ConsolidationRules.MAX_OVERLAP_RATIO:
            warnings.append(
                f"Sobreposição alta de valores ({overlap_ratio:.1%}) - "
                "pode indicar colunas com significados diferentes"
            )
        
        # Verifica tipos de dados incompatíveis
        type_conflicts = ConsolidationRules._detect_type_conflicts(all_values)
        if type_conflicts:
            warnings.extend(type_conflicts)
        
        return {
            'compatible': compatible,
            'reasons': reasons,
            'warnings': warnings,
            'overlap_ratio': overlap_ratio,
            'total_unique_values': total_unique
        }
    
    @staticmethod
    def _detect_type_conflicts(value_sets: List[Set[str]]) -> List[str]:
        """Deteta conflitos de tipos de dados entre conjuntos de valores"""
        warnings = []
        
        # Analisa padrões de dados em cada conjunto
        patterns = []
        for values in value_sets:
            if not values:
                patterns.append('empty')
                continue
            
            numeric_count = sum(1 for v in values if v.replace('.', '').replace('-', '').isdigit())
            numeric_ratio = numeric_count / len(values)
            
            if numeric_ratio > 0.8:
                patterns.append('numeric')
            elif numeric_ratio < 0.2:
                patterns.append('text')
            else:
                patterns.append('mixed')
        
        # Verifica se há conflitos significativos
        unique_patterns = set(p for p in patterns if p != 'empty')
        if len(unique_patterns) > 1 and 'mixed' not in unique_patterns:
            warnings.append(
                f"Tipos de dados diferentes detetados: {unique_patterns} - "
                "verificar se as colunas têm significados compatíveis"
            )
        
        return warnings
    
    @staticmethod
    def _check_similarity_threshold(columns: List[str], 
                                   similarity_scores: Dict[Tuple[str, str], float],
                                   logger: logging.Logger) -> Dict[str, any]:
        """Verifica se as colunas atendem ao threshold mínimo de similaridade"""
        relevant_scores = []
        
        for i, col1 in enumerate(columns):
            for col2 in columns[i+1:]:
                score = similarity_scores.get((col1, col2)) or similarity_scores.get((col2, col1))
                if score is not None:
                    relevant_scores.append(score)
        
        if not relevant_scores:
            return {
                'meets_threshold': True,  # Sem scores, não pode avaliar
                'reasons': [],
                'warnings': ['Scores de similaridade não disponíveis']
            }
        
        avg_similarity = sum(relevant_scores) / len(relevant_scores)
        min_similarity = min(relevant_scores)
        
        reasons = []
        warnings = []
        meets_threshold = True
        
        if avg_similarity < ConsolidationRules.MIN_SIMILARITY_THRESHOLD:
            meets_threshold = False
            reasons.append(
                f"Similaridade média muito baixa ({avg_similarity:.3f} < {ConsolidationRules.MIN_SIMILARITY_THRESHOLD})"
            )
        
        if min_similarity < 0.1:
            warnings.append(
                f"Algumas colunas têm similaridade muito baixa (mín: {min_similarity:.3f})"
            )
        
        return {
            'meets_threshold': meets_threshold,
            'reasons': reasons,
            'warnings': warnings,
            'avg_similarity': avg_similarity,
            'min_similarity': min_similarity
        }
    
    @staticmethod
    def _check_semantic_conflicts(columns: List[str], value_sets: Dict[str, Set[str]], 
                                 logger: logging.Logger) -> Dict[str, any]:
        """Verifica conflitos semânticos baseados nos nomes das colunas"""
        warnings = []
        
        # Palavras que podem indicar significados conflitantes
        conflict_indicators = [
            ('principal', 'secundario'), ('principal', 'secundária'),
            ('primario', 'secundario'), ('primário', 'secundário'),
            ('atual', 'anterior'), ('novo', 'antigo'),
            ('entrada', 'saida'), ('entrada', 'saída'),
            ('origem', 'destino'), ('inicial', 'final')
        ]
        
        # Verifica se há indicadores de conflito nos nomes
        column_words = []
        for col in columns:
            words = re.findall(r'\w+', col.lower())
            column_words.append(words)
        
        for word_set1 in conflict_indicators:
            for word_set2 in conflict_indicators:
                if word_set1 != word_set2:
                    continue
                
                found_word1 = any(word_set1[0] in words for words in column_words)
                found_word2 = any(word_set1[1] in words for words in column_words)
                
                if found_word1 and found_word2:
                    warnings.append(
                        f"Possível conflito semântico detetado: '{word_set1[0]}' vs '{word_set1[1]}' - "
                        "verificar se as colunas devem realmente ser consolidadas"
                    )
        
        return {
            'no_conflicts': len(warnings) == 0,
            'warnings': warnings
        }
    
    @staticmethod
    def _check_result_size(value_sets: Dict[str, Set[str]], logger: logging.Logger) -> Dict[str, any]:
        """Verifica se o resultado da consolidação não será excessivamente grande"""
        all_values = list(value_sets.values())
        total_unique = len(set().union(*all_values)) if all_values else 0
        
        warnings = []
        acceptable = True
        
        if total_unique > ConsolidationRules.MAX_UNIQUE_VALUES:
            warnings.append(
                f"Muitos valores únicos no resultado ({total_unique}) - "
                "consolidação pode resultar numa coluna muito diversa"
            )
        
        # Calcula estatísticas adicionais
        avg_values_per_column = sum(len(vs) for vs in all_values) / len(all_values) if all_values else 0
        
        if avg_values_per_column > 500:
            warnings.append(
                f"Média alta de valores por coluna ({avg_values_per_column:.0f}) - "
                "verificar se as colunas são realmente relacionadas"
            )
        
        return {
            'acceptable_size': acceptable,
            'warnings': warnings,
            'total_unique_values': total_unique,
            'avg_values_per_column': avg_values_per_column
        }
    
    @staticmethod
    def generate_consolidated_name(columns: List[str], logger: logging.Logger = None) -> str:
        """
        Gera nome apropriado para a coluna consolidada.
        
        Args:
            columns: Lista de colunas a consolidar
            logger: Logger configurado
            
        Returns:
            Nome limpo para a coluna consolidada
        """
        if logger is None:
            logger = logging.getLogger(__name__)
        
        if not columns:
            return "dim_consolidada"
        
        if len(columns) == 1:
            return ConsolidationRules._clean_column_name(columns[0])
        
        # Método 1: Encontrar prefixo comum
        common_prefix = ConsolidationRules._find_longest_common_prefix(columns)
        
        if common_prefix and len(common_prefix) > 6:  # Pelo menos "dim_xx"
            clean_name = ConsolidationRules._clean_column_name(common_prefix)
            logger.debug(f"Usando prefixo comum para nome: '{clean_name}'")
            return clean_name
        
        # Método 2: Extrair padrão base removendo sufixos
        base_name = ConsolidationRules._extract_base_pattern(columns)
        
        if base_name:
            clean_name = ConsolidationRules._clean_column_name(base_name)
            logger.debug(f"Usando padrão base para nome: '{clean_name}'")
            return clean_name
        
        # Método 3: Fallback - criar nome descritivo
        fallback_name = ConsolidationRules._generate_fallback_name(columns)
        clean_name = ConsolidationRules._clean_column_name(fallback_name)
        logger.debug(f"Usando nome fallback: '{clean_name}'")
        
        return clean_name
    
    @staticmethod
    def _find_longest_common_prefix(columns: List[str]) -> str:
        """Encontra o prefixo comum mais longo entre as colunas"""
        if not columns:
            return ""
        
        prefix = columns[0]
        for col in columns[1:]:
            common = ""
            for c1, c2 in zip(prefix, col):
                if c1 == c2:
                    common += c1
                else:
                    break
            prefix = common
        
        # Remove trailing underscore
        return prefix.rstrip('_')
    
    @staticmethod
    def _extract_base_pattern(columns: List[str]) -> str:
        """Extrai padrão base removendo sufixos numéricos e variações"""
        patterns = []
        
        for col in columns:
            # MELHORIA: Detecção específica para padrões numéricos óbvios
            # Remove sufixos numéricos do final
            match = re.match(r'^(.+?)(\d+)$', col)
            if match:
                base_part = match.group(1).rstrip('_')
                patterns.append(base_part)
                continue
            
            # MELHORIA: Detecção para números no meio
            # Para casos como dim_grupo1_etario, dim_grupo2_etario
            match = re.match(r'^(.+?)(\d+)(.+)$', col)
            if match:
                prefix = match.group(1).rstrip('_')
                suffix = match.group(3).lstrip('_')
                if suffix:
                    combined_pattern = f"{prefix}_{suffix}"
                    patterns.append(combined_pattern)
                else:
                    patterns.append(prefix)
                continue
            
            # Fallback: Remove último componente após underscore
            parts = col.split('_')
            if len(parts) > 2:
                patterns.append('_'.join(parts[:-1]))
            else:
                patterns.append(col)
        
        # MELHORIA: Encontra padrão mais comum com prioridade para padrões limpos
        if patterns:
            pattern_counts = {}
            for pattern in patterns:
                pattern_counts[pattern] = pattern_counts.get(pattern, 0) + 1
            
            # Prioriza padrões que aparecem em todas as colunas
            max_count = max(pattern_counts.values())
            best_patterns = [pattern for pattern, count in pattern_counts.items() if count == max_count]
            
            # Se há um padrão que cobre todas as colunas, usa esse
            if max_count == len(columns) and best_patterns:
                return best_patterns[0]
            
            # Senão, retorna o padrão mais frequente
            return max(pattern_counts.items(), key=lambda x: x[1])[0]
        
        return ""
    
    @staticmethod
    def _generate_fallback_name(columns: List[str]) -> str:
        """Gera nome de fallback quando outros métodos falham"""
        # Extrai palavras-chave dos nomes das colunas
        all_words = []
        for col in columns:
            words = re.findall(r'\w+', col.lower())
            # Remove 'dim' e números
            filtered_words = [w for w in words if w != 'dim' and not w.isdigit()]
            all_words.extend(filtered_words)
        
        # Conta frequência das palavras
        word_counts = {}
        for word in all_words:
            word_counts[word] = word_counts.get(word, 0) + 1
        
        # Pega as palavras mais comuns (máximo 3)
        common_words = sorted(word_counts.items(), key=lambda x: x[1], reverse=True)[:3]
        
        if common_words:
            keywords = [word for word, count in common_words]
            return f"dim_{'_'.join(keywords)}"
        else:
            return f"dim_consolidada_{len(columns)}_colunas"
    
    @staticmethod
    def _clean_column_name(name: str) -> str:
        """
        Aplica convenções de nomeação para nomes de colunas.
        
        Args:
            name: Nome original da coluna
            
        Returns:
            Nome limpo seguindo as convenções
        """
        # Remove acentos e caracteres especiais
        name = unicodedata.normalize('NFKD', name)
        name = ''.join(c for c in name if not unicodedata.combining(c))
        
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
        
        # Trunca se muito longo (limite do Excel é 255 caracteres, usamos 50 para segurança)
        if len(name) > 50:
            name = name[:47] + '...'
        
        return name
    
    @staticmethod
    def validate_consolidated_name(name: str) -> Tuple[bool, List[str]]:
        """
        Valida se o nome consolidado segue as convenções.
        
        Args:
            name: Nome a validar
            
        Returns:
            Tupla com (é_válido, lista_de_erros)
        """
        errors = []
        
        # Verifica se começa com 'dim_'
        if not name.startswith('dim_'):
            errors.append("Nome deve começar com 'dim_'")
        
        # Verifica caracteres válidos
        if not re.match(r'^[a-z0-9_]+$', name):
            errors.append("Nome deve conter apenas letras minúsculas, números e underscores")
        
        # Verifica comprimento
        if len(name) < 5:
            errors.append("Nome muito curto (mínimo 5 caracteres)")
        
        if len(name) > 50:
            errors.append("Nome muito longo (máximo 50 caracteres)")
        
        # Verifica se não termina com underscore
        if name.endswith('_'):
            errors.append("Nome não deve terminar com underscore")
        
        # Verifica se não tem underscores múltiplos
        if '__' in name:
            errors.append("Nome não deve ter underscores múltiplos")
        
        return len(errors) == 0, errors 