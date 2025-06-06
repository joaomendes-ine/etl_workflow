import json
from datetime import datetime
from typing import Dict, List, Any
import pandas as pd
import logging
from src.data_validator import CustomJSONEncoder

class ConsolidationReport:
    """
    Classe responsável pela geração de relatórios detalhados sobre 
    o processo de consolidação de dimensões.
    """
    
    def __init__(self, logger: logging.Logger = None):
        """
        Inicializa o gerador de relatórios.
        
        Args:
            logger: Logger configurado
        """
        self.logger = logger or logging.getLogger(__name__)
        self.consolidation_log = []
        self.analysis_details = {}
        self.integrity_checks = {}
        self.performance_metrics = {}
        self.start_time = None
        self.end_time = None
    
    def start_timing(self):
        """Inicia o cronómetro do processo"""
        self.start_time = datetime.now()
        self.logger.debug("Cronómetro de consolidação iniciado")
    
    def end_timing(self):
        """Termina o cronómetro do processo"""
        self.end_time = datetime.now()
        if self.start_time:
            duration = (self.end_time - self.start_time).total_seconds()
            self.performance_metrics['total_duration_seconds'] = duration
            self.logger.debug(f"Cronómetro de consolidação terminado: {duration:.2f}s")
    
    def log_analysis_phase(self, phase_name: str, details: Dict[str, Any]):
        """
        Regista detalhes de uma fase da análise.
        
        Args:
            phase_name: Nome da fase
            details: Detalhes da fase
        """
        self.analysis_details[phase_name] = {
            'timestamp': datetime.now().isoformat(),
            'details': details
        }
        self.logger.debug(f"Fase '{phase_name}' registada no relatório")
    
    def log_consolidation_action(self, action_type: str, source_columns: List[str], 
                                target_column: str, success: bool, details: Dict[str, Any] = None):
        """
        Regista uma ação de consolidação.
        
        Args:
            action_type: Tipo de ação ('consolidate', 'skip', 'error')
            source_columns: Colunas de origem
            target_column: Coluna de destino
            success: Se a ação foi bem-sucedida
            details: Detalhes adicionais
        """
        action_record = {
            'timestamp': datetime.now().isoformat(),
            'action_type': action_type,
            'source_columns': source_columns,
            'target_column': target_column,
            'success': success,
            'details': details or {}
        }
        
        self.consolidation_log.append(action_record)
        
        action_desc = f"'{action_type}' para {len(source_columns)} colunas -> '{target_column}'"
        if success:
            self.logger.info(f"Ação registada: {action_desc}")
        else:
            self.logger.warning(f"Ação falhada: {action_desc}")
    
    def log_integrity_check(self, check_name: str, passed: bool, details: Dict[str, Any]):
        """
        Regista resultado de verificação de integridade.
        
        Args:
            check_name: Nome da verificação
            passed: Se passou na verificação
            details: Detalhes da verificação
        """
        self.integrity_checks[check_name] = {
            'timestamp': datetime.now().isoformat(),
            'passed': passed,
            'details': details
        }
        
        status = "PASSOU" if passed else "FALHOU"
        self.logger.info(f"Verificação de integridade '{check_name}': {status}")
    
    def generate_summary(self, original_df: pd.DataFrame, consolidated_df: pd.DataFrame) -> Dict[str, Any]:
        """
        Gera resumo das operações de consolidação mostrando números originais.
        
        Args:
            original_df: DataFrame original (após limpeza de dimensões vazias)
            consolidated_df: DataFrame consolidado
            
        Returns:
            Dicionário com resumo das operações
        """
        # Obtém contagem ORIGINAL de dimensões (antes da limpeza)
        data_loading_info = self.analysis_details.get('data_loading', {}).get('details', {})
        empty_removal_info = self.analysis_details.get('empty_dimensions_removal', {}).get('details', {})
        
        # Número original de dimensões (antes da remoção de vazias)
        original_dim_count = data_loading_info.get('original_dimension_columns', 
                                                 len([col for col in original_df.columns if col.startswith('dim_')]))
        
        # Número atual de dimensões (após limpeza e consolidação)
        consolidated_dim_count = len([col for col in consolidated_df.columns if col.startswith('dim_')])
        
        # Dimensões removidas por estarem vazias
        removed_empty_dimensions = empty_removal_info.get('removed_empty_dimensions', [])
        empty_removed_count = len(removed_empty_dimensions)
        
        # Cálculo da redução baseado no número original
        total_reduction = original_dim_count - consolidated_dim_count
        reduction_percentage = (total_reduction / original_dim_count * 100) if original_dim_count > 0 else 0
        
        # Contabiliza consolidações realizadas
        consolidations_performed = len([action for action in self.consolidation_log 
                                      if action['action_type'] == 'consolidate' and action['success']])
        
        # Contabiliza erros
        errors_occurred = len([action for action in self.consolidation_log 
                             if action['action_type'] == 'error' or not action['success']])
        
        # Status de integridade
        integrity_status = self.get_integrity_status()
        critical_checks = ['row_count', 'valor_column', 'non_dimension_columns', 'absolute_value_preservation']
        critical_passed = all(integrity_status.get(check, False) for check in critical_checks)
        
        summary = {
            'original_dimensions': original_dim_count,  # NÚMERO ORIGINAL
            'consolidated_dimensions': consolidated_dim_count,
            'total_reduction': total_reduction,
            'reduction_percentage': round(reduction_percentage, 1),
            'empty_dimensions_removed': empty_removed_count,
            'removed_empty_dimension_names': removed_empty_dimensions,
            'consolidations_performed': consolidations_performed,
            'errors_occurred': errors_occurred,
            'rows_processed': len(consolidated_df),
            'critical_integrity_maintained': critical_passed,
            'execution_time_seconds': round(self.performance_metrics.get('total_duration_seconds', 0), 2),
            'actions_summary': {
                'consolidate': consolidations_performed,
                'error': errors_occurred
            }
        }
        
        self.logger.info(f"Resumo gerado: {total_reduction} dimensões reduzidas ({reduction_percentage:.1f}%)")
        
        return summary
    
    def generate_detailed_report(self, original_df: pd.DataFrame, 
                                consolidated_df: pd.DataFrame) -> Dict[str, Any]:
        """
        Gera relatório detalhado completo.
        
        Args:
            original_df: DataFrame original
            consolidated_df: DataFrame consolidado
            
        Returns:
            Relatório completo com todos os detalhes
        """
        summary = self.generate_summary(original_df, consolidated_df)
        
        detailed_report = {
            'report_metadata': {
                'generated_at': datetime.now().isoformat(),
                'report_version': '1.0',
                'generator': 'ConsolidationReport'
            },
            'summary': summary,
            'analysis_phases': self.analysis_details,
            'consolidation_actions': self.consolidation_log,
            'integrity_checks': self.integrity_checks,
            'performance_details': self._generate_performance_details(),
            'warnings_and_recommendations': self._generate_recommendations()
        }
        
        self.logger.info("Relatório detalhado gerado com sucesso")
        
        return detailed_report
    
    def _generate_performance_details(self) -> Dict[str, Any]:
        """Gera detalhes de performance"""
        details = self.performance_metrics.copy()
        
        if self.start_time and self.end_time:
            details['start_time'] = self.start_time.isoformat()
            details['end_time'] = self.end_time.isoformat()
        
        # Calcula métricas derivadas
        if 'total_duration_seconds' in details:
            duration = details['total_duration_seconds']
            details['duration_formatted'] = f"{duration:.2f} segundos"
            
            if duration > 300:  # 5 minutos
                details['performance_rating'] = 'slow'
            elif duration > 60:  # 1 minuto
                details['performance_rating'] = 'moderate'
            else:
                details['performance_rating'] = 'fast'
        
        return details
    
    def _generate_recommendations(self) -> List[str]:
        """Gera recomendações baseadas no processo"""
        recommendations = []
        
        # Baseado em verificações de integridade
        failed_checks = [name for name, check in self.integrity_checks.items() if not check['passed']]
        if failed_checks:
            recommendations.append(
                f"Verificar problemas de integridade em: {', '.join(failed_checks)}"
            )
        
        # Baseado em ações falhadas
        failed_actions = [action for action in self.consolidation_log if not action['success']]
        if failed_actions:
            recommendations.append(
                f"Revisar {len(failed_actions)} ações de consolidação que falharam"
            )
        
        # Baseado em performance
        if self.performance_metrics.get('total_duration_seconds', 0) > 300:
            recommendations.append(
                "Considerar otimizar o processo para datasets grandes - tempo de execução elevado"
            )
        
        # Baseado em padrões de análise
        if 'pattern_detection' in self.analysis_details:
            patterns_found = len(self.analysis_details['pattern_detection']['details'].get('patterns', {}))
            if patterns_found == 0:
                recommendations.append(
                    "Nenhum padrão de consolidação detetado - verificar estrutura dos dados"
                )
        
        return recommendations
    
    def save_report(self, file_path: str, detailed: bool = True) -> str:
        """
        Guarda o relatório num ficheiro.
        
        Args:
            file_path: Caminho do ficheiro de saída
            detailed: Se True, guarda relatório detalhado; caso contrário, apenas resumo
            
        Returns:
            Caminho do ficheiro guardado
        """
        try:
            if detailed:
                report_data = self.generate_detailed_report(pd.DataFrame(), pd.DataFrame())
            else:
                report_data = {
                    'summary': self.generate_summary(pd.DataFrame(), pd.DataFrame()),
                    'generated_at': datetime.now().isoformat()
                }
            
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(report_data, f, indent=2, ensure_ascii=False, cls=CustomJSONEncoder)
            
            self.logger.info(f"Relatório guardado em: {file_path}")
            return file_path
            
        except Exception as e:
            self.logger.error(f"Erro ao guardar relatório: {str(e)}")
            raise
    
    def print_summary(self, original_df: pd.DataFrame = None, consolidated_df: pd.DataFrame = None):
        """Imprime resumo formatado no console"""
        if original_df is not None and consolidated_df is not None:
            summary = self.generate_summary(original_df, consolidated_df)
        else:
            # Usa dados cached se disponível
            summary = getattr(self, '_cached_summary', {})
        
        if not summary:
            print("Nenhum resumo disponível")
            return
        
        print("\n============================================================")
        print("           RELATÓRIO DE CONSOLIDAÇÃO DE DIMENSÕES")
        print("============================================================")
        print(f"Dimensões originais:     {summary['original_dimensions']}")
        print(f"Dimensões consolidadas:  {summary['consolidated_dimensions']}")
        print(f"Redução:                 {summary['total_reduction']} ({summary['reduction_percentage']:.1f}%)")
        print(f"Linhas processadas:      {summary['rows_processed']:,}")
        
        # Informações sobre dimensões vazias removidas
        empty_removed = summary.get('empty_dimensions_removed', 0)
        if empty_removed > 0:
            print(f"\nDimensões vazias removidas: {empty_removed}")
            removed_names = summary.get('removed_empty_dimension_names', [])
            if removed_names:
                print("Motivo: Dimensões completamente vazias (sem valores)")
                if len(removed_names) <= 5:
                    for name in removed_names:
                        print(f"  • {name}")
                else:
                    for name in removed_names[:5]:
                        print(f"  • {name}")
                    print(f"  • ... mais {len(removed_names) - 5} dimensões")
        
        # Informações sobre valores "Total" removidos
        total_removal_info = self.analysis_details.get('total_values_removal', {}).get('details', {})
        total_removed = total_removal_info.get('total_values_removed', 0)
        if total_removed > 0:
            print(f"\nValores 'Total' removidos: {total_removed}")
            print("Motivo: Remoção de valores 'Total' solicitada pelo utilizador")
        
        print(f"\nAções executadas:")
        actions_summary = summary.get('actions_summary', {})
        for action_type, count in actions_summary.items():
            print(f"  - {action_type.capitalize()}: {count}")
        
        if summary.get('critical_integrity_maintained'):
            print("\n[OK] Integridade essencial mantida (algumas verificações não-críticas podem ter falhado)")
            print("    [INFO] Falhas não-críticas são esperadas durante consolidação de dimensões")
        else:
            print("\n[AVISO] Algumas verificações de integridade falharam")
        
        if summary.get('execution_time_seconds'):
            duration = summary['execution_time_seconds']
            print(f"\nTempo de execução: {duration:.2f} segundos")
        
        print("============================================================")
        
        # Cache o summary para uso futuro
        self._cached_summary = summary
    
    def get_consolidation_details(self) -> List[Dict[str, Any]]:
        """
        Retorna detalhes de todas as ações de consolidação.
        
        Returns:
            Lista de ações de consolidação
        """
        return self.consolidation_log.copy()
    
    def get_failed_actions(self) -> List[Dict[str, Any]]:
        """
        Retorna apenas as ações que falharam.
        
        Returns:
            Lista de ações falhadas
        """
        return [action for action in self.consolidation_log if not action['success']]
    
    def get_integrity_status(self) -> Dict[str, bool]:
        """
        Retorna o status de todas as verificações de integridade.
        
        Returns:
            Dicionário com nome da verificação como chave e boolean como valor
        """
        return {check_name: check['passed'] for check_name, check in self.integrity_checks.items()}
    
    def get_phase_details(self, phase_name: str) -> Dict[str, Any]:
        """
        Retorna detalhes de uma fase específica da análise.
        
        Args:
            phase_name: Nome da fase a consultar
            
        Returns:
            Dicionário com detalhes da fase ou None se não encontrada
        """
        phase_info = self.analysis_details.get(phase_name)
        if phase_info:
            return phase_info.get('details', {})
        return None
    
    def reset(self):
        """Limpa todos os dados do relatório para nova execução"""
        self.consolidation_log = []
        self.analysis_details = {}
        self.integrity_checks = {}
        self.performance_metrics = {}
        self.start_time = None
        self.end_time = None
        self.logger.debug("Dados do relatório limpos para nova execução") 