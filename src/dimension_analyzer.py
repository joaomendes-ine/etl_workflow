import pandas as pd
import re
from typing import Dict, List, Set, Tuple
from collections import defaultdict
import logging
from difflib import SequenceMatcher

class DimensionAnalyzer:
    """
    Classe responsável pela análise de colunas de dimensão para deteção de padrões
    e análise de valores para consolidação inteligente.
    """
    
    def __init__(self, df: pd.DataFrame, logger: logging.Logger = None):
        """
        Inicializa o analisador de dimensões.
        
        Args:
            df: DataFrame com os dados a analisar
            logger: Logger configurado
        """
        self.df = df
        self.logger = logger or logging.getLogger(__name__)
        self.dimension_columns = [col for col in df.columns if col.startswith('dim_')]
        self.patterns = {}
        self.value_mappings = {}
        self.similarity_matrix = {}
        
        self.logger.info(f"Inicializado DimensionAnalyzer com {len(self.dimension_columns)} colunas de dimensão")
        self.logger.debug(f"Colunas de dimensão encontradas: {self.dimension_columns}")
    
    def analyze_patterns(self) -> Dict[str, List[str]]:
        """
        Deteta padrões de nomeação em colunas de dimensão usando regex e análise de strings.
        VERSÃO MELHORADA: Mais agressiva na detecção de dimensões relacionadas.
        
        Returns:
            Dicionário com padrão base como chave e lista de colunas como valor
        """
        if not self.dimension_columns:
            self.logger.warning("Nenhuma coluna de dimensão encontrada para análise de padrões")
            return {}
        
        patterns = defaultdict(list)
        
        self.logger.info("🔍 Análise agressiva de padrões de dimensões relacionadas...")
        
        # Método 1: Sufixos numéricos (MELHORADO)
        self._detect_enhanced_numeric_patterns(patterns)
        
        # Método 2: Palavras-chave comuns (NOVO)
        self._detect_keyword_patterns(patterns)
        
        # Método 3: Similaridade semântica agressiva (MELHORADO)
        self._detect_semantic_similarity_patterns(patterns)
        
        # Método 4: Prefixos longos comuns (MELHORADO)
        self._detect_enhanced_prefix_patterns(patterns)
        
        # Método 5: Padrões de classificação (NOVO - para CAE, CPP, etc.)
        self._detect_classification_patterns(patterns)
        
        # Remove duplicatas e filtra grupos pequenos
        cleaned_patterns = self._clean_and_merge_patterns(patterns)
        
        self.patterns = cleaned_patterns
        self.logger.info(f"✅ Padrões agressivos detetados: {len(self.patterns)}")
        for pattern, cols in self.patterns.items():
            self.logger.info(f"   📊 '{pattern}': {len(cols)} colunas → {cols}")
        
        return self.patterns
    
    def _detect_enhanced_numeric_patterns(self, patterns: defaultdict):
        """Deteta padrões numéricos melhorados com estratégia mais agressiva"""
        self.logger.debug("🔢 Detectando padrões numéricos agressivos...")
        
        # Estratégia 1: Prefixos exatos com sufixos numéricos
        prefix_groups = defaultdict(list)
        
        for col in self.dimension_columns:
            # Remove 'dim_' para análise
            base_name = col[4:] if col.startswith('dim_') else col
            self.logger.debug(f"   Analisando coluna '{col}' -> base_name='{base_name}'")
            
            # Procura por padrões como 'grupo_etario1', 'grupo_etario2'
            # Regex para capturar base + número no final
            match = re.match(r'^(.+?)(\d+)$', base_name)
            if match:
                base_part = match.group(1).rstrip('_')
                number = int(match.group(2))
                prefix_groups[base_part].append(col)
                self.logger.debug(f"   ✅ Padrão numérico encontrado: '{col}' -> base='{base_part}', num={number}")
                continue
            
            # Estratégia 2: Procura por números no meio (mais flexível)
            # Para casos como 'dim_setor1_economia', 'dim_setor2_economia'
            match = re.match(r'^(.+?)(\d+)(.*)$', base_name)
            if match:
                prefix = match.group(1).rstrip('_')
                suffix = match.group(3).lstrip('_')
                combined_base = f"{prefix}_{suffix}".strip('_') if suffix else prefix
                prefix_groups[combined_base].append(col)
                self.logger.debug(f"   ✅ Padrão numérico flexível: '{col}' -> base='{combined_base}'")
                continue
                
            self.logger.debug(f"   ❌ Nenhum padrão numérico encontrado para '{col}'")
        
        self.logger.debug(f"   📊 Grupos de prefixos encontrados: {dict(prefix_groups)}")
        
        # Estratégia 3: Análise de similaridade de nomes mais agressiva
        # Para casos onde os nomes são muito similares mas não seguem padrão numérico exato
        for i, col1 in enumerate(self.dimension_columns):
            for col2 in self.dimension_columns[i+1:]:
                # Calcula similaridade de nome
                similarity = self._calculate_aggressive_name_similarity(col1, col2)
                
                if similarity >= 0.7:  # Threshold mais baixo (era 0.85)
                    base_key = self._extract_aggressive_common_base([col1, col2])
                    if base_key and len(base_key) > 3:  # Base deve ter pelo menos 4 caracteres
                        prefix_groups[base_key].extend([col1, col2])
                        self.logger.debug(f"   ✅ Similaridade alta detectada: '{col1}' + '{col2}' -> base='{base_key}' (sim={similarity:.2f})")
        
        # Remove duplicatas e adiciona aos padrões
        for base, columns in prefix_groups.items():
            # Remove duplicatas mantendo ordem
            unique_columns = []
            seen = set()
            for col in columns:
                if col not in seen:
                    unique_columns.append(col)
                    seen.add(col)
            
            if len(unique_columns) >= 2:
                pattern_key = f"numeric_{base}"
                patterns[pattern_key].extend(unique_columns)
                self.logger.info(f"   🎯 PADRÃO CONFIRMADO: '{pattern_key}' com {len(unique_columns)} colunas: {unique_columns}")
            else:
                self.logger.debug(f"   ⚠️ Grupo '{base}' rejeitado: apenas {len(unique_columns)} coluna(s)")
    
    def _calculate_aggressive_name_similarity(self, col1: str, col2: str) -> float:
        """Calcula similaridade de nome de forma mais agressiva"""
        # Remove prefixo 'dim_' 
        name1 = col1[4:] if col1.startswith('dim_') else col1
        name2 = col2[4:] if col2.startswith('dim_') else col2
        
        # Remove números do final para comparação
        clean_name1 = re.sub(r'\d+$', '', name1).rstrip('_')
        clean_name2 = re.sub(r'\d+$', '', name2).rstrip('_')
        
        # Se os nomes limpos são iguais, similaridade máxima
        if clean_name1 == clean_name2 and clean_name1:
            return 1.0
        
        # Calcula similaridade usando diferentes métodos
        similarities = []
        
        # Método 1: Prefixo comum
        common_prefix = self._find_common_prefix(clean_name1, clean_name2)
        if common_prefix:
            prefix_sim = len(common_prefix) / max(len(clean_name1), len(clean_name2))
            similarities.append(prefix_sim)
        
        # Método 2: Jaro-Winkler aproximado
        sequence_sim = SequenceMatcher(None, clean_name1, clean_name2).ratio()
        similarities.append(sequence_sim)
        
        # Método 3: Palavras em comum
        words1 = set(clean_name1.split('_'))
        words2 = set(clean_name2.split('_'))
        if words1 and words2:
            common_words = words1 & words2
            word_sim = len(common_words) / len(words1 | words2)
            similarities.append(word_sim)
        
        # Retorna a maior similaridade
        return max(similarities) if similarities else 0.0
    
    def _extract_aggressive_common_base(self, columns: List[str]) -> str:
        """Extrai base comum de forma mais agressiva"""
        if not columns:
            return ""
        
        # Remove 'dim_' de todas
        clean_names = [col[4:] if col.startswith('dim_') else col for col in columns]
        
        # Remove números do final
        base_names = [re.sub(r'\d+$', '', name).rstrip('_') for name in clean_names]
        
        # Se todos têm a mesma base após limpeza, usa essa base
        if len(set(base_names)) == 1 and base_names[0]:
            return base_names[0]
        
        # Senão, encontra o prefixo comum mais longo
        if len(clean_names) < 2:
            return clean_names[0] if clean_names else ""
        
        common = clean_names[0]
        for name in clean_names[1:]:
            # Encontra prefixo comum
            new_common = ""
            for i, (c1, c2) in enumerate(zip(common, name)):
                if c1 == c2:
                    new_common += c1
                else:
                    break
            common = new_common.rstrip('_')
        
        return common if len(common) > 3 else ""
    
    def _detect_keyword_patterns(self, patterns: defaultdict):
        """NOVO: Deteta padrões baseados em palavras-chave semânticas"""
        self.logger.debug("🔤 Detectando padrões por palavras-chave...")
        
        # Palavras-chave que indicam conceitos relacionados
        keyword_groups = {
            'grupo_etario': ['grupo_etario', 'idade', 'etario', 'faixa_etaria'],
            'setor_economia': ['setor', 'economia', 'cae', 'atividade', 'ativ'],
            'frequencia': ['frequencia', 'freq', 'trimestral', 'mensal', 'anual'],
            'condicao_trabalho': ['condicao', 'trabalho', 'emprego', 'profiss'],
            'geografia': ['regiao', 'local', 'area', 'territorio', 'nuts'],
            'exercicio': ['exercicio', 'ano', 'periodo', 'data'],
            'nivel_ensino': ['ensino', 'educacao', 'escolar', 'nivel'],
            'situacao': ['situacao', 'estado', 'status', 'condicao'],
        }
        
        for keyword_base, keywords in keyword_groups.items():
            matching_columns = []
            
            for col in self.dimension_columns:
                col_lower = col.lower()
                if any(keyword in col_lower for keyword in keywords):
                    matching_columns.append(col)
                    self.logger.debug(f"      Coluna '{col}' corresponde a keyword '{keyword_base}'")
                    
            if len(matching_columns) > 1:
                patterns[f"keyword_{keyword_base}"].extend(matching_columns)
                self.logger.info(f"   🎯 PADRÃO KEYWORD CONFIRMADO: '{keyword_base}' com {len(matching_columns)} colunas: {matching_columns}")
            else:
                self.logger.debug(f"   Grupo keyword '{keyword_base}' rejeitado: apenas {len(matching_columns)} coluna(s)")
    
    def _detect_semantic_similarity_patterns(self, patterns: defaultdict):
        """MELHORADO: Detecção semântica mais agressiva"""
        self.logger.debug("   Detectando similaridade semântica agressiva...")
        
        processed_columns = set()
        
        for i, col1 in enumerate(self.dimension_columns):
            if col1 in processed_columns:
                continue
                
            similar_group = [col1]
            
            for col2 in self.dimension_columns[i+1:]:
                if col2 in processed_columns:
                    continue
                    
                # Calcula similaridade de nome
                name_similarity = self._calculate_enhanced_name_similarity(col1, col2)
                
                if name_similarity > 0.6:  # Threshold mais baixo = mais agressivo
                    similar_group.append(col2)
                    processed_columns.add(col2)
            
            if len(similar_group) > 1:
                # Gera nome base comum
                common_base = self._extract_common_semantic_base(similar_group)
                patterns[f"semantic_{common_base}"].extend(similar_group)
                processed_columns.add(col1)
                
                self.logger.debug(f"      Grupo semântico '{common_base}': {similar_group}")
    
    def _detect_enhanced_prefix_patterns(self, patterns: defaultdict):
        """MELHORADO: Detecção de prefixos mais sofisticada"""
        self.logger.debug("   Detectando prefixos comuns melhorados...")
        
        # Agrupa por prefixos progressivamente menores
        prefix_groups = defaultdict(list)
        
        for col in self.dimension_columns:
            parts = col.split('_')
            
            # Testa diferentes tamanhos de prefixo
            for prefix_size in range(2, len(parts)):
                prefix = '_'.join(parts[:prefix_size])
                
                if len(prefix) >= 8:  # Prefixo mínimo significativo
                    prefix_groups[prefix].append(col)
        
        # Filtra apenas grupos com múltiplas colunas
        for prefix, cols in prefix_groups.items():
            if len(cols) > 1:
                patterns[f"prefix_{prefix}"].extend(cols)
                self.logger.debug(f"      Prefixo '{prefix}': {cols}")
    
    def _detect_classification_patterns(self, patterns: defaultdict):
        """NOVO: Deteta padrões de sistemas de classificação (CAE, CPP, etc.)"""
        self.logger.debug("   Detectando padrões de classificação...")
        
        classification_indicators = {
            'cae': ['cae', 'atividade', 'setor'],
            'cpp': ['cpp', 'profiss', 'ocupacao'],
            'nuts': ['nuts', 'regiao', 'territorio'],
            'cpc': ['cpc', 'produto', 'consumo'],
            'nace': ['nace', 'economic', 'activity']
        }
        
        for class_type, indicators in classification_indicators.items():
            matching_cols = []
            
            for col in self.dimension_columns:
                col_lower = col.lower()
                if any(indicator in col_lower for indicator in indicators):
                    matching_cols.append(col)
            
            if len(matching_cols) > 1:
                patterns[f"classification_{class_type}"].extend(matching_cols)
                self.logger.debug(f"      Classificação '{class_type}': {matching_cols}")
    
    def _calculate_enhanced_name_similarity(self, col1: str, col2: str) -> float:
        """MELHORADO: Calcula similaridade mais sofisticada"""
        # Similaridade básica
        basic_similarity = SequenceMatcher(None, col1, col2).ratio()
        
        # Similaridade de palavras (ignora ordem)
        words1 = set(col1.lower().split('_'))
        words2 = set(col2.lower().split('_'))
        
        if words1 and words2:
            word_intersection = len(words1.intersection(words2))
            word_union = len(words1.union(words2))
            word_similarity = word_intersection / word_union
        else:
            word_similarity = 0.0
        
        # Combina ambas as métricas
        combined = (basic_similarity * 0.4) + (word_similarity * 0.6)
        
        return combined
    
    def _extract_common_semantic_base(self, columns: List[str]) -> str:
        """Extrai base semântica comum de um grupo de colunas"""
        if not columns:
            return "unknown"
        
        # Encontra palavras comuns
        word_sets = [set(col.lower().split('_')) for col in columns]
        common_words = set.intersection(*word_sets) if word_sets else set()
        
        # Remove palavras muito genéricas
        generic_words = {'dim', 'de', 'do', 'da', 'dos', 'das', 'e', 'ou', 'com', 'sem'}
        meaningful_words = common_words - generic_words
        
        if meaningful_words:
            # Ordena por aparição na primeira coluna
            first_col_words = columns[0].lower().split('_')
            ordered_words = [w for w in first_col_words if w in meaningful_words]
            return '_'.join(ordered_words[:3])  # Máximo 3 palavras
        else:
            # Fallback: usa prefixo comum
            common_prefix = self._find_common_prefix(columns[0], columns[1])
            return common_prefix.replace('dim_', '') or 'related'
    
    def _clean_and_merge_patterns(self, patterns: defaultdict) -> Dict[str, List[str]]:
        """Limpa e mescla padrões sobrepostos"""
        self.logger.debug("🧹 Limpando e mesclando padrões...")
        
        # Converte para dicionário normal e remove duplicatas
        clean_patterns = {}
        all_assigned_columns = set()
        
        # Prioriza padrões por qualidade (mais específicos primeiro)
        pattern_priority = ['keyword_', 'classification_', 'numeric_', 'semantic_', 'prefix_']
        
        self.logger.debug(f"   Padrões antes da limpeza: {dict(patterns)}")
        
        for priority_type in pattern_priority:
            self.logger.debug(f"   Processando padrões do tipo '{priority_type}'")
            
            for pattern_name, columns in patterns.items():
                if not any(pt in pattern_name for pt in [priority_type]) and len(columns) > 1:
                    continue
                
                if priority_type in pattern_name and len(columns) > 1:
                    # Remove colunas já atribuídas
                    available_columns = [col for col in columns if col not in all_assigned_columns]
                    
                    self.logger.debug(f"      Padrão '{pattern_name}': {len(columns)} colunas originais, {len(available_columns)} disponíveis")
                    
                    if len(available_columns) > 1:
                        # Remove duplicatas mantendo ordem
                        unique_columns = []
                        seen = set()
                        for col in available_columns:
                            if col not in seen:
                                unique_columns.append(col)
                                seen.add(col)
                        
                        if len(unique_columns) > 1:
                            clean_patterns[pattern_name] = unique_columns
                            all_assigned_columns.update(unique_columns)
                            self.logger.info(f"   ✅ PADRÃO APROVADO '{pattern_name}': {unique_columns}")
                        else:
                            self.logger.debug(f"      Padrão '{pattern_name}' rejeitado após remoção de duplicatas: {len(unique_columns)} coluna(s)")
                    else:
                        self.logger.debug(f"      Padrão '{pattern_name}' rejeitado: colunas já atribuídas")
        
        # BACKUP: Se nenhum padrão foi detectado, adiciona TODOS os padrões válidos (modo agressivo)
        if not clean_patterns:
            self.logger.warning("🚨 NENHUM PADRÃO DETECTADO! Ativando modo backup agressivo...")
            
            for pattern_name, columns in patterns.items():
                if len(columns) > 1:
                    # Remove duplicatas
                    unique_columns = list(dict.fromkeys(columns))  # Preserva ordem
                    
                    if len(unique_columns) > 1:
                        clean_patterns[pattern_name] = unique_columns
                        self.logger.warning(f"   🆘 BACKUP: Padrão '{pattern_name}' adicionado: {unique_columns}")
        
        self.logger.info(f"   📊 PADRÕES FINAIS: {len(clean_patterns)} grupos")
        for pattern_name, columns in clean_patterns.items():
            self.logger.info(f"      '{pattern_name}': {columns}")
        
        return clean_patterns
    
    def _find_common_prefix(self, col1: str, col2: str) -> str:
        """Encontra o prefixo comum entre duas strings"""
        common = ''
        for c1, c2 in zip(col1, col2):
            if c1 == c2:
                common += c1
            else:
                break
        # Remove trailing underscore
        return common.rstrip('_')
    
    def analyze_values(self, columns: List[str] = None) -> Dict[str, Set[str]]:
        """
        Extrai valores únicos para cada coluna especificada.
        
        Args:
            columns: Lista de colunas a analisar. Se None, analisa todas as colunas de dimensão.
            
        Returns:
            Dicionário com nome da coluna como chave e set de valores únicos como valor
        """
        if columns is None:
            columns = self.dimension_columns
        
        value_mappings = {}
        
        for col in columns:
            if col in self.df.columns:
                # Remove valores nulos e converte para string para comparação
                unique_values = set(str(val) for val in self.df[col].dropna().unique() if pd.notna(val))
                value_mappings[col] = unique_values
                self.logger.debug(f"Coluna '{col}': {len(unique_values)} valores únicos")
            else:
                self.logger.warning(f"Coluna '{col}' não encontrada no DataFrame")
                value_mappings[col] = set()
        
        self.value_mappings.update(value_mappings)
        return value_mappings
    
    def calculate_similarity(self, col1: str, col2: str) -> float:
        """
        Calcula score de similaridade entre duas colunas baseado nos valores.
        
        Args:
            col1: Nome da primeira coluna
            col2: Nome da segunda coluna
            
        Returns:
            Score de similaridade entre 0 e 1
        """
        if col1 not in self.value_mappings:
            self.analyze_values([col1])
        if col2 not in self.value_mappings:
            self.analyze_values([col2])
        
        values1 = self.value_mappings.get(col1, set())
        values2 = self.value_mappings.get(col2, set())
        
        if not values1 and not values2:
            return 1.0  # Ambas vazias
        if not values1 or not values2:
            return 0.0  # Uma vazia, outra não
        
        # Calcula Jaccard similarity
        intersection = len(values1.intersection(values2))
        union = len(values1.union(values2))
        
        jaccard_similarity = intersection / union if union > 0 else 0.0
        
        # Considera também a similaridade estrutural dos valores
        structural_similarity = self._calculate_structural_similarity(values1, values2)
        
        # Combina ambas as métricas
        combined_similarity = (jaccard_similarity * 0.7) + (structural_similarity * 0.3)
        
        # Cache do resultado
        self.similarity_matrix[(col1, col2)] = combined_similarity
        self.similarity_matrix[(col2, col1)] = combined_similarity
        
        self.logger.debug(f"Similaridade entre '{col1}' e '{col2}': {combined_similarity:.3f}")
        
        return combined_similarity
    
    def _calculate_structural_similarity(self, values1: Set[str], values2: Set[str]) -> float:
        """Calcula similaridade estrutural entre dois conjuntos de valores"""
        # Compara tipos de dados, comprimentos médios, padrões, etc.
        
        # Similaridade de tamanho dos valores
        avg_len1 = sum(len(v) for v in values1) / len(values1) if values1 else 0
        avg_len2 = sum(len(v) for v in values2) / len(values2) if values2 else 0
        
        len_similarity = 1 - abs(avg_len1 - avg_len2) / max(avg_len1, avg_len2, 1)
        
        # Similaridade de padrões (numérico vs texto)
        numeric1 = sum(1 for v in values1 if v.replace('.', '').replace('-', '').isdigit())
        numeric2 = sum(1 for v in values2 if v.replace('.', '').replace('-', '').isdigit())
        
        numeric_ratio1 = numeric1 / len(values1) if values1 else 0
        numeric_ratio2 = numeric2 / len(values2) if values2 else 0
        
        pattern_similarity = 1 - abs(numeric_ratio1 - numeric_ratio2)
        
        return (len_similarity + pattern_similarity) / 2
    
    def get_consolidation_candidates(self) -> Dict[str, Dict[str, any]]:
        """
        Identifica candidatos para consolidação baseado em padrões e similaridade de valores.
        
        Returns:
            Dicionário com informações detalhadas sobre candidatos para consolidação
        """
        candidates = {}
        
        # Analisa padrões se ainda não foi feito
        if not self.patterns:
            self.analyze_patterns()
        
        for pattern, columns in self.patterns.items():
            if len(columns) <= 1:
                continue
            
            # Analisa valores para as colunas do padrão
            value_analysis = self.analyze_values(columns)
            
            # Calcula similaridades entre todas as colunas do grupo
            similarities = {}
            for i, col1 in enumerate(columns):
                for col2 in columns[i+1:]:
                    sim_score = self.calculate_similarity(col1, col2)
                    similarities[(col1, col2)] = sim_score
            
            # Determina se são bons candidatos
            avg_similarity = sum(similarities.values()) / len(similarities) if similarities else 0
            
            candidates[pattern] = {
                'columns': columns,
                'value_sets': value_analysis,
                'similarities': similarities,
                'avg_similarity': avg_similarity,
                'total_unique_values': len(set().union(*value_analysis.values())),
                'can_consolidate': self._assess_consolidation_feasibility(columns, value_analysis, avg_similarity)
            }
            
            self.logger.info(f"Candidato '{pattern}': {len(columns)} colunas, similaridade média: {avg_similarity:.3f}")
        
        return candidates
    
    def _assess_consolidation_feasibility(self, columns: List[str], value_sets: Dict[str, Set[str]], avg_similarity: float) -> Dict[str, any]:
        """Avalia a viabilidade de consolidação para um grupo de colunas"""
        # Verifica sobreposições de valores
        all_values = list(value_sets.values())
        overlaps = []
        
        for i, set1 in enumerate(all_values):
            for set2 in all_values[i+1:]:
                overlap = len(set1.intersection(set2))
                overlaps.append(overlap)
        
        max_overlap = max(overlaps) if overlaps else 0
        total_values = len(set().union(*all_values)) if all_values else 0
        
        # Critérios de viabilidade
        feasible = True
        reasons = []
        warnings = []
        
        # Critério 1: Similaridade mínima
        if avg_similarity < 0.3:
            feasible = False
            reasons.append(f"Similaridade muito baixa ({avg_similarity:.3f})")
        
        # Critério 2: Sobreposições excessivas podem indicar problemas
        if max_overlap > total_values * 0.8:
            warnings.append("Sobreposição alta de valores - verificar se têm significados diferentes")
        
        # Critério 3: Muitos valores únicos pode indicar colunas muito diferentes
        if total_values > 1000:
            warnings.append("Muitos valores únicos - consolidação pode resultar em coluna muito diversa")
        
        return {
            'feasible': feasible,
            'reasons': reasons,
            'warnings': warnings,
            'max_overlap': max_overlap,
            'total_unique_values': total_values
        } 