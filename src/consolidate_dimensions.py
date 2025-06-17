#!/usr/bin/env python3
"""
Employment and Unemployment Data Dimension Consolidation Script - COMPREHENSIVE VERSION

This script consolidates multiple dimension columns using a comprehensive mapping
that includes ALL values from original dimensions, with hierarchical value selection.

Key Features:
- Uses comprehensive dimension mapping with ALL original values
- Implements "last value" selection for conflicting dimensions
- Preserves exact row order from original file
- Removes 'dim_geografia' as it's redundant (single value)
- Maintains data integrity with validation

Requirements:
- pandas
- openpyxl (for Excel support)
- numpy
"""

import pandas as pd
import numpy as np
import re
from typing import List, Dict, Set, Tuple
import sys
import os
from datetime import datetime

def get_comprehensive_dimension_groups() -> Dict[str, Dict]:
    """
    Define dimension groups with ALL values from original dimensions.
    Only the 10 dimensions mapped by the user.
    Each group contains:
    - 'columns': list of original column names to consolidate
    - 'values': ALL unique values from those columns (excluding 'Total')
    """
    return {
        'dim_grupo_etario': {
            'columns': [
                'dim_grupo_etario1', 'dim_grupo_etario2', 'dim_grupo_etario3',
                'dim_grupo_etario4', 'dim_grupo_etario_jovens', 'dim_grupo_etario1a',
                'dim_grupo_etario2a', 'dim_grupo_etario3_1'
            ],
            'values': [
                "15 - 24 anos", "15 - 64 anos", "15 e mais anos", "16 - 19 anos",
                "16 - 24 anos", "16 - 34 anos", "16 - 64 anos", "16 - 74 anos",
                "16 - 89 anos", "16 e mais anos", "20 - 24 anos", "25 - 34 anos",
                "35 - 44 anos", "45 - 54 anos", "55 - 64 anos", "55 - 74 anos",
                "65 - 89 anos", "65 e mais anos", "Menos de 15 anos", "Menos de 16 anos"
            ]
        },
        'dim_setor_atividade_economica': {
            'columns': [
                'dim_setor_economia_cae_rev3_ativ_principal',
                'dim_atividade_economia_cae_rev3_ativ_principal',
                'dim_setor_economia_cae_rev3_ativ_secundaria',
                'dim_setor_economia_cae_rev21_ativ_principal',
                'dim_atividade_economia_cae_rev21_ativ_principal',
                'dim_setor_atividade92'
            ],
            'values': [
                "A - Agricultura, produção animal, caça, floresta e pesca",
                "A a B: Agricultura, silvicultura e pesca",
                "A: Agricultura, produção animal, caça, floresta e pesca",
                "B - Indústrias extrativas",
                "B a F: Indústria, construção, energia e água",
                "C - Indústrias transformadoras",
                "C a F: Indústria, construção, energia e água",
                "C: Indústrias transformadoras",
                "D - Eletricidade, gás, vapor, água quente e fria e ar frio",
                "E - Captação, tratamento e distribuição de água; saneamento, gestão de resíduos e despoluição",
                "F - Construção",
                "F: Construção",
                "G - Comércio por grosso e a retalho; reparação de veículos automóveis e motociclos",
                "G a Q: Serviços",
                "G a U: Serviços",
                "G: Comércio por grosso e a retalho",
                "H - Transportes e armazenagem",
                "H: Alojamento e restauração",
                "I - Alojamento, restauração e similares",
                "I: Transportes, armazenagem e comunicações",
                "J - Atividades de informação e de comunicação",
                "K - Atividades financeiras e de seguros",
                "L - Atividades imobiliárias",
                "L: Administração Pública, defesa e Segurança Social obrigatória",
                "M - Atividades de consultoria, científica, técnicas e similares",
                "M: Educação",
                "N - Atividades administrativas e dos serviços de apoio",
                "N: Saúde e acção social",
                "O - Administração Pública e Defesa; Segurança Social Obrigatória",
                "O: Outras actividades de serviços colectivos, sociais e pessoais",
                "P - Educação",
                "Primário",
                "Q - Atividades de saúde humana e apoio social",
                "R - Atividades artísticas, de espetáculos, desportivas e recreativas",
                "S a U - Outros serviços",
                "Secundário",
                "Terciário"
            ]
        },
        'dim_profissao': {
            'columns': [
                'dim_profissao_principal_CPP10',
                'dim_profissao_principal_CNP94'
            ],
            'values': [
                "1  Rep. poder legisl. e órg. execut., dirig., diret. e gest. executivos",
                "1: Quadros superiores da administração pública, dirigentes e quadros superiores de empresa",
                "2  Especialistas das atividades intelectuais e científicas",
                "2: Especialistas das profissões intelectuais e científicas",
                "3  Técnicos e profissões de nível intermédio",
                "3: Técnicos e profissionais de nível intermédio",
                "4  Pessoal administrativo",
                "4: Pessoal administrativo e similares",
                "5  Trab. serv. pessoais, proteção e segurança e vendedores",
                "5: Pessoal dos serviços e vendedores",
                "6  Agricultores e trab. qualif. da agric., da pesca e da floresta",
                "6: Agricultores e trabalhadores qualificados da agricultura e pescas",
                "7  Trabalhadores qualif. da indústria, construção e artífices",
                "7: Operários, artífices e trabalhadores similares",
                "8  Oper. de instalações e máquinas e trab. da montagem",
                "8: Operadores de instalações e máquinas e trabalhadores da montagem",
                "9  Trabalhadores não qualificados",
                "9: Trabalhadores não qualificados"
            ]
        },
        'dim_situacao_profissional': {
            'columns': [
                'dim_situacao_profissao_principal',
                'dim_situacao_profissao_principal_98',
                'dim_situacao_profissao_principal_TCP98'
            ],
            'values': [
                "Trabalhador por conta de outrem",
                "Trabalhador por conta propria",
                "Trabalhador por conta propria como isolado",
                "Trabalhador por conta própria como isolado",
                "Trabalhador por conta propria como empregador"
            ]
        },
        'dim_condicao_trabalho': {
            'columns': [
                'dim_condicao_trabalho',
                'dim_condicao_trabalho_inativo',
                'dim_condicao_trabalho_inativo98',
                'dim_condicao_trabalho_modulo2020'
            ],
            'values': [
                "Desempregado",
                "Doméstico",
                "Doméstico  (dos 16 aos 89 anos)",
                "Empregado",
                "Estudante",
                "Estudante (dos 16 aos 89 anos)",
                "Inativo",
                "Outro inativo",
                "Outro inativo (16 e mais anos)",
                "Outros inativos",
                "Reformado",
                "Reformado (dos 16 aos 89 anos)"
            ]
        },
        'dim_trabalho_casa': {
            'columns': [
                'dim_trabalho_casa_20',
                'dim_trabalho_casa_22',
                'dim_trabalho_casa_22_TC',
                'dim_trabalho_casa_TIC',
                'dim_trabalho_casa_equipamento'
            ],
            'values': [
                "Apenas de computador",
                "Computador e smartphone",
                "Não sabe",
                "Não trabalhou em casa",
                "Não trabalhou em casa ou não trabalhou sempre ou quase sempre em casa",
                "Não utilizou TIC ou não sabe",
                "O trabalho em casa foi realizado fora do horário de trabalho",
                "Trabalhou em casa",
                "Trabalhou em casa pontualmente",
                "Trabalhou em casa regularmente mediante um sistema que concilia trabalho presencial e em casa",
                "Trabalhou sempre em casa",
                "Trabalhou sempre ou quase sempre em casa",
                "Utilizou TIC"
            ]
        },
        'dim_educacao_formacao': {
            'columns': [
                'dim_pop1674_educacao_formacao',
                'dim_pop1674_educacao_formal',
                'dim_pop1674_educacao_nao_formal'
            ],
            'values': [
                "Frequentou atividades de educação e formação",
                "Frequentou educação formal",
                "Frequentou educação não-formal",
                "Não frequentou atividades de educação e formação",
                "Não frequentou educação formal",
                "Não frequentou educação não-formal"
            ]
        },
        'dim_estado_saude': {
            'columns': [
                'dim_pop1689_estado_saude',
                'dim_pop1689_limitacoes_saude',
                'dim_problemas_saude',
                'dim_limita_problema'
            ],
            'values': [
                "Bom",
                "Dois ou mais problemas de saúde",
                "Limita consideravelmente",
                "Limita em certa medida",
                "Limitado, mas não severamente",
                "Mau",
                "Muito bom",
                "Muito mau",
                "Nada limitado",
                "Não limita",
                "Razoável",
                "Severamente limitado",
                "Um problema de saúde"
            ]
        },
        'dim_tempo_tarefas_trabalho': {
            'columns': [
                'dim_leitura_manuais',
                'dim_calculos',
                'dim_trab_arduo',
                'dim_tarefas_destreza_manual',
                'dim_interacao_organizacao',
                'dim_interacao_fora_organizacao',
                'dim_aconselhar_ensinar',
                'dim_trab_dispositivos_digitais'
            ],
            'values': [
                "Metade ou mais  do tempo de trabalho",
                "Nenhum tempo de trabalho",
                "Não resposta",
                "Pouco ou parte do tempo de trabalho"
            ]
        },
        'dim_autonomia_trabalho': {
            'columns': [
                'dim_grau_autonomia',
                'dim_autonomia_decidir'
            ],
            'values': [
                "Alguma autonomia para decidir a ordem das tarefas/trabalhos, mas pouca ou nenhuma autonomia para decidir sobre a sua execução",
                "Alguma autonomia para decidir a ordem das tarefas/trabalhos, mas total ou muita autonomia para decidir sobre a sua execução",
                "Alguma autonomia para decidir a ordem e a execução das tarefas/trabalhos",
                "Alguma autonomia para decidir a ordem e alguma autonomia para decidir o conteúdo das tarefas",
                "Alguma autonomia para decidir a ordem e o conteúdo das tarefas",
                "Alguma autonomia para decidir a ordem e pouca ou nenhuma autonomia para decidir o conteúdo das tarefas",
                "Alguma autonomia para decidir a ordem e total ou muita autonomia para decidir o conteúdo das tarefas",
                "Alguma autonomia para decidir sobre a execução das tarefas/trabalhos, mas pouca ou nenhuma autonomia para decidir sobre a sua ordem",
                "Não resposta",
                "Pouca ou nenhuma autonomia para decidir a ordem e alguma autonomia para decidir o conteúdo das tarefas",
                "Pouca ou nenhuma autonomia para decidir a ordem e o conteúdo das tarefas",
                "Pouca ou nenhuma autonomia para decidir a ordem e total ou muita autonomia para decidir o conteúdo das tarefas",
                "Pouca ou nenhuma autonomia para decidir sobre a ordem e a execução das tarefas/trabalhos",
                "Total ou muita autonomia para decidir a ordem das tarefas/trabalhos, mas apenas alguma autonomia para decidir sobre a sua execução",
                "Total ou muita autonomia para decidir a ordem das tarefas/trabalhos, mas pouca ou nenhuma autonomia para decidir sobre a sua execução",
                "Total ou muita autonomia para decidir a ordem e a execução das tarefas/trabalhos",
                "Total ou muita autonomia para decidir a ordem e alguma autonomia para decidir o conteúdo das tarefas",
                "Total ou muita autonomia para decidir a ordem e o conteúdo das tarefas",
                "Total ou muita autonomia para decidir a ordem e pouca ou nenhuma autonomia para decidir o conteúdo das tarefas",
                "Total ou muita autonomia para decidir sobre a execução das tarefas/trabalhos, mas pouca ou nenhuma autonomia para decidir sobre a sua ordem"
            ]
        }
    }

def load_data(file_path: str) -> pd.DataFrame:
    """Load Excel file into pandas DataFrame."""
    print(f"Loading data from {file_path}...")
    try:
        df = pd.read_excel(file_path, engine='openpyxl')
        print(f"Loaded {len(df)} rows and {len(df.columns)} columns")
        return df
    except Exception as e:
        print(f"Error loading file: {e}")
        sys.exit(1)

def clean_total_values(df: pd.DataFrame) -> pd.DataFrame:
    """Replace all 'Total' or 'total' values with empty string."""
    print("\nCleaning 'Total' values...")
    
    total_count = 0
    for col in df.columns:
        if col.startswith('dim_'):
            mask = df[col].astype(str).str.lower() == 'total'
            total_count += mask.sum()
    
    print(f"Found {total_count} 'Total' values to clean")
    
    for col in df.columns:
        if col.startswith('dim_'):
            df[col] = df[col].astype(str).replace(['Total', 'total', 'TOTAL'], '', regex=False)
            df[col] = df[col].replace(['nan', 'None', 'NaN'], '', regex=False)
    
    return df

def remove_geografia_column(df: pd.DataFrame) -> pd.DataFrame:
    """Remove dim_geografia column as it's redundant (single value)."""
    print("\nRemoving dim_geografia column (redundant)...")
    
    if 'dim_geografia' in df.columns:
        unique_values = df['dim_geografia'].dropna().unique()
        print(f"  dim_geografia has {len(unique_values)} unique values: {unique_values}")
        df = df.drop(columns=['dim_geografia'])
        print("  Removed dim_geografia column")
    else:
        print("  dim_geografia not found in columns")
    
    return df

def consolidate_dimensions_hierarchical(df: pd.DataFrame) -> pd.DataFrame:
    """
    Consolidate dimensions using hierarchical "last value" logic.
    
    Key Rules:
    1. For each row, scan dimension columns from left to right (order in original file)
    2. Take the LAST non-empty value found (rightmost in hierarchy)
    3. This preserves hierarchical order from original file
    4. NEVER duplicates rows or values
    """
    print("\nConsolidating dimensions with hierarchical logic...")
    
    dimension_groups = get_comprehensive_dimension_groups()
    
    for new_dim, config in dimension_groups.items():
        existing_cols = [col for col in config['columns'] if col in df.columns]
        
        if not existing_cols:
            continue
            
        print(f"\nConsolidating {len(existing_cols)} columns into '{new_dim}':")
        for col in existing_cols:
            print(f"  - {col}")
        
        print(f"  Expected values: {len(config['values'])} unique values")
        
        # Apply hierarchical "last value" logic
        consolidated_values = []
        
        for idx, row in df.iterrows():
            # Scan columns from left to right (preserving original hierarchy)
            # Take the LAST (rightmost) non-empty value
            final_value = ''
            
            for col in existing_cols:  # Already in correct order from mapping
                if col in row and pd.notna(row[col]):
                    val = str(row[col]).strip()
                    if val != '' and val != 'nan' and val.lower() != 'total':
                        final_value = val  # Keep updating - last one wins
            
            consolidated_values.append(final_value)
        
        # Add the new consolidated column
        df[new_dim] = consolidated_values
        
        # Remove the old columns (but NOT the new consolidated column if it has the same name)
        cols_to_remove = [col for col in existing_cols if col != new_dim]
        if cols_to_remove:
            df = df.drop(columns=cols_to_remove)
        
        # Report consolidation results
        if new_dim in df.columns:
            unique_count = df[new_dim].astype(str).str.strip()
            unique_count = unique_count[unique_count != ''].nunique()
            print(f"  → Consolidated to {unique_count} unique values in {new_dim}")
        else:
            print(f"  ⚠️  Error: {new_dim} column was not created properly")
    
    return df

def validate_expected_values(df: pd.DataFrame) -> bool:
    """Validate that consolidated dimensions contain expected values."""
    print("\nValidating consolidated dimension values...")
    
    dimension_groups = get_comprehensive_dimension_groups()
    all_valid = True
    
    for new_dim, config in dimension_groups.items():
        if new_dim not in df.columns:
            continue
            
        expected_values = set(config['values'])
        actual_values = set(df[new_dim].dropna().astype(str).str.strip())
        actual_values.discard('')  # Remove empty strings
        
        # Check for unexpected values
        unexpected = actual_values - expected_values
        missing = expected_values - actual_values
        
        print(f"\n  {new_dim}:")
        print(f"    Expected: {len(expected_values)} values")
        print(f"    Found: {len(actual_values)} values")
        
        if unexpected:
            print(f"    ⚠️  Unexpected values ({len(unexpected)}): {list(unexpected)[:5]}...")
            all_valid = False
        
        if len(missing) > len(expected_values) * 0.5:  # More than 50% missing
            print(f"    ⚠️  Many expected values missing ({len(missing)}/{len(expected_values)})")
            all_valid = False
        elif missing:
            print(f"    ℹ️  Some expected values not found ({len(missing)})")
    
    if all_valid:
        print("\n✅ All dimension values are within expected ranges")
    else:
        print("\n⚠️  Some dimension values need review")
    
    return all_valid

def preserve_column_order(df: pd.DataFrame) -> pd.DataFrame:
    """Reorder columns to maintain logical hierarchy with exactly 13 dimensions."""
    print("\nPreserving column order...")
    
    # Define the desired order for the 13 final dimensions only
    dimension_order = [
        # Original dimensions (3)
        'dim_ano', 'dim_trimestre', 'dim_sexo',
        # Consolidated dimensions (10) 
        'dim_grupo_etario', 'dim_setor_atividade_economica', 
        'dim_profissao', 'dim_situacao_profissional', 
        'dim_condicao_trabalho', 'dim_trabalho_casa', 
        'dim_educacao_formacao', 'dim_estado_saude',
        'dim_tempo_tarefas_trabalho', 'dim_autonomia_trabalho'
    ]
    
    # Get existing dimension columns in preferred order
    ordered_dims = []
    for dim in dimension_order:
        if dim in df.columns:
            ordered_dims.append(dim)
    
    # Add any remaining dimension columns (should not happen if unwanted dims are removed)
    remaining_dims = [col for col in df.columns if col.startswith('dim_') and col not in ordered_dims]
    if remaining_dims:
        print(f"  ⚠️  Unexpected remaining dimensions: {remaining_dims}")
        ordered_dims.extend(remaining_dims)
    
    # Add non-dimension columns
    non_dim_cols = [col for col in df.columns if not col.startswith('dim_')]
    
    # Priority for non-dimension columns
    priority_cols = ['indicador', 'unidade', 'valor', 'simbologia', 'estado_valor', 'coeficiente_variacao']
    ordered_non_dims = []
    
    for col in priority_cols:
        if col in non_dim_cols:
            ordered_non_dims.append(col)
            non_dim_cols.remove(col)
    
    # Add remaining non-dimension columns
    ordered_non_dims.extend(non_dim_cols)
    
    # Final column order
    final_order = ordered_dims + ordered_non_dims
    
    # Reorder dataframe
    df = df[final_order]
    
    print(f"  Reordered {len(ordered_dims)} dimensions + {len(ordered_non_dims)} other columns")
    return df

def remove_exact_duplicates(df: pd.DataFrame) -> pd.DataFrame:
    """Remove truly identical rows after consolidation."""
    print("\nRemoving exact duplicates...")
    
    initial_rows = len(df)
    
    # Check for complete duplicates across all columns
    df = df.drop_duplicates(keep='first')
    
    removed_rows = initial_rows - len(df)
    print(f"  Removed {removed_rows} exact duplicate rows")
    print(f"  Remaining rows: {len(df)}")
    
    return df

def validate_valor_integrity(df_original: pd.DataFrame, df_final: pd.DataFrame) -> bool:
    """Validate that valor column hasn't been corrupted."""
    print("\nValidating valor column integrity...")
    
    if 'valor' not in df_original.columns or 'valor' not in df_final.columns:
        print("  Warning: 'valor' column not found")
        return True
    
    original_valores = pd.to_numeric(df_original['valor'], errors='coerce').fillna(0)
    final_valores = pd.to_numeric(df_final['valor'], errors='coerce').fillna(0)
    
    original_sum = original_valores.sum()
    final_sum = final_valores.sum()
    
    print(f"  Original: {len(df_original):,} rows, sum: {original_sum:,.2f}")
    print(f"  Final: {len(df_final):,} rows, sum: {final_sum:,.2f}")
    
    difference = final_sum - original_sum
    tolerance = abs(original_sum * 0.001)  # 0.1% tolerance
    
    if abs(difference) <= tolerance:
        print(f"  ✅ Valor integrity maintained (diff: {difference:,.2f})")
        return True
    else:
        print(f"  ⚠️  WARNING: Valor sum changed by {difference:,.2f}")
        return False

def generate_summary_report(df_original: pd.DataFrame, df_final: pd.DataFrame) -> str:
    """Generate comprehensive summary report."""
    report = []
    report.append("=" * 80)
    report.append("COMPREHENSIVE DIMENSION CONSOLIDATION SUMMARY")
    report.append("=" * 80)
    report.append(f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    report.append("")
    
    report.append("ORIGINAL DATA:")
    report.append(f"  - Total rows: {len(df_original):,}")
    report.append(f"  - Total columns: {len(df_original.columns)}")
    original_dims = len([c for c in df_original.columns if c.startswith('dim_')])
    report.append(f"  - Dimension columns: {original_dims}")
    report.append("")
    
    report.append("CONSOLIDATED DATA:")
    report.append(f"  - Total rows: {len(df_final):,}")
    report.append(f"  - Total columns: {len(df_final.columns)}")
    final_dims = len([c for c in df_final.columns if c.startswith('dim_')])
    report.append(f"  - Dimension columns: {final_dims}")
    report.append("")
    
    report.append("CONSOLIDATION RESULTS:")
    report.append(f"  - Rows change: {len(df_final) - len(df_original):+,}")
    report.append(f"  - Dimensions reduced: {original_dims - final_dims} ({original_dims} → {final_dims})")
    report.append(f"  - Reduction rate: {((original_dims - final_dims) / original_dims * 100):.1f}%")
    report.append("")
    
    report.append("CONSOLIDATED DIMENSION SUMMARY:")
    dim_cols = sorted([c for c in df_final.columns if c.startswith('dim_')])
    for col in dim_cols:
        unique_values = df_final[col].astype(str).str.strip()
        unique_values = unique_values[unique_values != ''].nunique()
        report.append(f"  - {col}: {unique_values} unique values")
    
    report.append("")
    report.append("=" * 80)
    
    return "\n".join(report)

def remove_unwanted_dimensions(df: pd.DataFrame) -> pd.DataFrame:
    """
    Remove all dimension columns that are NOT in the final desired set.
    Keep only: dim_ano, dim_trimestre, dim_sexo + 10 consolidated dimensions = 13 total
    """
    print("\nRemoving unwanted dimension columns...")
    
    # Define the ONLY dimensions we want to keep (13 total)
    desired_dimensions = {
        # Original dimensions to keep (3)
        'dim_ano', 'dim_trimestre', 'dim_sexo',
        # Consolidated dimensions (10 total)
        'dim_grupo_etario', 'dim_setor_atividade_economica', 
        'dim_profissao', 'dim_situacao_profissional', 
        'dim_condicao_trabalho', 'dim_trabalho_casa', 
        'dim_educacao_formacao', 'dim_estado_saude',
        'dim_tempo_tarefas_trabalho', 'dim_autonomia_trabalho'
    }
    
    # Find all current dimension columns
    current_dims = [col for col in df.columns if col.startswith('dim_')]
    
    # Find dimensions to remove
    dims_to_remove = [col for col in current_dims if col not in desired_dimensions]
    
    if dims_to_remove:
        print(f"  Removing {len(dims_to_remove)} unwanted dimension columns:")
        for col in dims_to_remove:
            print(f"    - {col}")
        
        # Remove unwanted dimensions
        df = df.drop(columns=dims_to_remove)
    else:
        print("  No unwanted dimensions found")
    
    # Verify we have exactly 13 dimensions
    remaining_dims = [col for col in df.columns if col.startswith('dim_')]
    print(f"  Final dimension count: {len(remaining_dims)} (target: 13)")
    
    if len(remaining_dims) != 13:
        print(f"  ⚠️  Warning: Expected 13 dimensions, found {len(remaining_dims)}")
        print(f"  Dimensions found: {remaining_dims}")
    
    return df

def main():
    """Main execution function."""
    # Configuration
    dataset_dir = "dataset"
    result_dir = "result"
    input_file = os.path.join(dataset_dir, '65_BD.xlsx')
    output_file = os.path.join(result_dir, '65_BD_consolidated.xlsx')
    
    print("=" * 80)
    print("COMPREHENSIVE EMPLOYMENT DATA DIMENSION CONSOLIDATION")
    print("=" * 80)
    
    # Check directories
    if not os.path.exists(dataset_dir):
        print(f"Error: Dataset directory '{dataset_dir}' not found!")
        sys.exit(1)
    
    if not os.path.exists(result_dir):
        os.makedirs(result_dir)
        print(f"Created result directory: {result_dir}")
    
    if not os.path.exists(input_file):
        print(f"Error: Input file '{input_file}' not found!")
        sys.exit(1)
    
    # Load and process data
    df = load_data(input_file)
    df_original = df.copy()
    
    # Step 1: Clean Total values
    df = clean_total_values(df)
    
    # Step 2: Remove redundant geografia column
    df = remove_geografia_column(df)
    
    # Step 3: Consolidate dimensions with hierarchical logic
    df = consolidate_dimensions_hierarchical(df)
    
    # Step 4: Validate consolidated values
    validate_expected_values(df)
    
    # Step 5: Preserve column order
    df = preserve_column_order(df)
    
    # Step 6: Remove exact duplicates
    df = remove_exact_duplicates(df)
    
    # Step 7: Validate integrity
    valor_ok = validate_valor_integrity(df_original, df)
    if not valor_ok:
        print("\n⚠️  Valor integrity warning - review recommended")
    
    # Step 8: Remove unwanted dimensions
    df = remove_unwanted_dimensions(df)
    
    # Generate summary
    summary = generate_summary_report(df_original, df)
    print("\n" + summary)
    
    # Save results
    print(f"\nSaving consolidated data to {output_file}...")
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Consolidated_Data', index=False)
        
        summary_df = pd.DataFrame({'Summary': summary.split('\n')})
        summary_df.to_excel(writer, sheet_name='Summary_Report', index=False)
    
    print("✅ Excel file saved with UTF-8 encoding")
    
    # Save CSV
    csv_file = os.path.join(result_dir, '65_Emprego_e_desemprego_BD_consolidated.csv')
    df.to_csv(csv_file, index=False, encoding='utf-8-sig')
    print(f"✅ CSV file saved: {csv_file}")
    
    # File size validation
    for filepath in [output_file, csv_file]:
        if os.path.exists(filepath):
            size_mb = os.path.getsize(filepath) / (1024 * 1024)
            print(f"   {os.path.basename(filepath)}: {size_mb:.2f} MB")
    
    print(f"\n🎉 Comprehensive consolidation complete! Files saved in '{result_dir}' folder.")

if __name__ == "__main__":
    main()