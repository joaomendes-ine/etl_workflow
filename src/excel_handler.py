import os
import pandas as pd
import re
import logging
import xlwings as xw  # Alterado de openpyxl para xlwings
from typing import List, Dict, Any, Tuple
from src.utils import ensure_directory_exists
import time  # Adicionado para pausas entre operações

# Obter o logger configurado centralmente.
# Assume-se que setup_logging() é chamado no ponto de entrada da aplicação (ex: main.py).
logger = logging.getLogger('etl_workflow')


def parse_sheet_name(file_name: str) -> Tuple[str, List[int]]:
    """
    Analisa o nome da planilha/arquivo para extrair prefixo e valores numéricos.
    Lida com formatos como Q.1, I.10, II.2, III.2.1, etc.
    
    Args:
        file_name: Nome do arquivo/planilha
        
    Returns:
        Tupla com (prefixo, lista de números)
    """
    # Remove a extensão do arquivo se presente
    base_name = os.path.splitext(file_name)[0]
    
    # Ordem de prefixos para classificação
    prefix_order = {
        'Q': 1,
        'I': 2,
        'II': 3, 
        'III': 4,
        'IV': 5,
        'V': 6
    }
    
    # Extrai o prefixo (Q, I, II, III, etc.)
    prefix_match = re.match(r'^([A-Z]+)', base_name)
    prefix = prefix_match.group(1) if prefix_match else ""
    
    # Extrai todos os números no nome (para lidar com casos como III.2.1)
    numbers = []
    for num_str in re.findall(r'(\d+)', base_name):
        try:
            numbers.append(int(num_str))
        except ValueError:
            numbers.append(0)
    
    # Se não tem números, coloca um valor alto para ordenar por último
    if not numbers:
        numbers = [float('inf')]
    
    # Completa a lista de números com zeros para garantir comparação correta
    # para casos como III.2 vs III.2.1
    while len(numbers) < 3:
        numbers.append(0)
        
    return prefix, numbers


def sort_excel_files(file_list: List[str]) -> List[str]:
    """
    Ordena os ficheiros Excel de acordo com uma estrutura específica:
    Primeiro os arquivos Q.x, depois I.x, II.x, III.x, etc.
    Respeita subníveis como III.2.1, III.2.2, etc.
    
    Args:
        file_list: Lista de nomes de arquivos Excel
        
    Returns:
        Lista ordenada de nomes de arquivos
    """
    # Ordem de prefixos para classificação
    prefix_order = {
        'Q': 1,
        'I': 2,
        'II': 3, 
        'III': 4,
        'IV': 5,
        'V': 6
    }
    
    # Lista para armazenar arquivos com seus metadados
    files_with_metadata = []
    
    for f_name in file_list:
        if not (f_name.lower().endswith('.xlsx') or f_name.lower().endswith('.xls')):
            continue  # Ignorar ficheiros não Excel
        
        # Obter prefixo e números do nome do arquivo
        prefix, numbers = parse_sheet_name(f_name)
        
        # Atribuir valor de ordem do prefixo, se conhecido
        prefix_value = prefix_order.get(prefix, 999)  # Default alto para desconhecidos
        
        # Adicionar à lista com os metadados necessários para ordenação
        files_with_metadata.append({
            'name': f_name,
            'prefix': prefix,
            'prefix_value': prefix_value,
            'numbers': numbers,
            'original_base': os.path.splitext(f_name)[0]
        })
    
    # Ordenar a lista com base nos critérios:
    # 1. Valor do prefixo (Q antes de I antes de II, etc.)
    # 2. Números na sequência (1.1 antes de 1.2, etc.)
    # 3. Nome original para desempate, se necessário
    files_with_metadata.sort(key=lambda x: (
        x['prefix_value'],
        x['numbers'][0], 
        x['numbers'][1], 
        x['numbers'][2],
        x['original_base']
    ))
    
    # Retornar apenas os nomes dos arquivos, agora ordenados
    return [f['name'] for f in files_with_metadata]


def copy_sheet_properties(source_sheet, target_sheet):
    """
    Copia propriedades importantes de uma planilha para outra, incluindo:
    - Largura das colunas
    - Altura das linhas
    - Formatação de células
    
    Args:
        source_sheet: Planilha de origem (xlwings)
        target_sheet: Planilha de destino (xlwings)
    """
    try:
        # Tentamos copiar as larguras das colunas
        last_column = source_sheet.used_range.last_cell.column
        for col_idx in range(1, last_column + 1):
            # Obter a largura da coluna na planilha de origem
            column_width = source_sheet.api.Columns(col_idx).ColumnWidth
            # Aplicar a mesma largura na planilha de destino
            target_sheet.api.Columns(col_idx).ColumnWidth = column_width
        
        # Tentamos copiar as alturas das linhas
        last_row = source_sheet.used_range.last_cell.row
        for row_idx in range(1, last_row + 1):
            # Obter a altura da linha na planilha de origem
            row_height = source_sheet.api.Rows(row_idx).RowHeight
            # Aplicar a mesma altura na planilha de destino
            target_sheet.api.Rows(row_idx).RowHeight = row_height
        
        # Copiamos também a cor de fundo das células, bordas e outras propriedades
        used_range = source_sheet.used_range
        used_range.copy()
        target_sheet.range('A1').paste('formats')
        
        # Tentamos também copiar outras propriedades gerais da planilha
        target_sheet.api.StandardWidth = source_sheet.api.StandardWidth
        
        logger.info("Propriedades da planilha copiadas com sucesso (larguras, alturas e formatação)")
    except Exception as e:
        logger.warning(f"Erro ao copiar propriedades da planilha: {e}")


def consolidate_excel_files(source_folder: str, output_folder: str, output_file_name_base: str) -> bool:
    """
    Consolida ficheiros Excel de uma pasta de origem para um único ficheiro Excel na pasta de destino.
    Cada ficheiro de origem torna-se uma folha no ficheiro de destino, preservando toda a formatação original.
    A ordem das folhas segue uma estrutura específica: Q.x, I.x, II.x, III.x, etc.
    
    Utiliza xlwings, que é baseado no próprio Excel para garantir a preservação exata da formatação.

    Args:
        source_folder: Caminho para a pasta contendo os ficheiros Excel de origem.
        output_folder: Caminho para a pasta onde o ficheiro consolidado será guardado.
        output_file_name_base: Base para o nome do ficheiro de saída (ex: "55_quadros").

    Returns:
        True se a consolidação for bem-sucedida e ficheiros forem processados, False caso contrário.
    """
    ensure_directory_exists(output_folder) # Garante que a pasta de destino existe

    try:
        excel_files = [f for f in os.listdir(source_folder)
                      if os.path.isfile(os.path.join(source_folder, f)) and
                      not f.startswith('~$') and # Ignora arquivos temporários do Excel
                      (f.lower().endswith('.xlsx') or f.lower().endswith('.xls'))]
    except FileNotFoundError:
        logger.error(f"Pasta de origem '{source_folder}' não encontrada.")
        return False
    except Exception as e:
        logger.error(f"Erro ao listar ficheiros em '{source_folder}': {e}")
        return False

    if not excel_files:
        logger.warning(f"Nenhum ficheiro Excel (.xlsx, .xls) válido encontrado em '{source_folder}'.")
        return False

    sorted_files = sort_excel_files(excel_files)
    logger.info(f"Ficheiros Excel encontrados e ordenados para processamento: {sorted_files}")

    # Define o nome do arquivo de saída (sem o sufixo _BD)
    output_file_path = os.path.join(output_folder, f"{output_file_name_base}.xlsx")
    
    # Remove o arquivo de saída anterior se existir
    if os.path.exists(output_file_path):
        try:
            os.remove(output_file_path)
            logger.info(f"Arquivo de saída anterior removido: '{output_file_path}'")
        except OSError as e:
            logger.error(f"Erro ao remover arquivo de saída anterior '{output_file_path}': {e}. Certifique-se de que o ficheiro não está aberto.")
            return False # Retorna False se não conseguir remover o arquivo
    
    # Também remova qualquer arquivo com o sufixo _BD que possa existir (da versão anterior do código)
    old_bd_file_path = os.path.join(output_folder, f"{output_file_name_base}_BD.xlsx")
    if os.path.exists(old_bd_file_path):
        try:
            os.remove(old_bd_file_path)
            logger.info(f"Arquivo antigo com sufixo _BD removido: '{old_bd_file_path}'")
        except OSError:
            # Se não conseguir remover, apenas logue e continue
            logger.warning(f"Não foi possível remover o arquivo antigo com sufixo _BD: '{old_bd_file_path}'")

    # Inicia o Excel em segundo plano
    app = None
    try:
        app = xw.App(visible=False)
        app.display_alerts = False  # Desativa alertas do Excel
        
        processed_any_file = False
        
        # Cria um novo workbook
        target_workbook = app.books.add()
        
        # Obtém a planilha inicial do workbook (geralmente "Sheet1")
        default_sheet = target_workbook.sheets[0]
        
        for i, file_name in enumerate(sorted_files):
            file_path = os.path.join(source_folder, file_name)
            # Extrair o nome da planilha do nome do arquivo (exemplo: Q1, Q2, etc.)
            sheet_name_original = os.path.splitext(file_name)[0]
            # Trunca o nome da folha se for maior que 31 caracteres (limite do Excel)
            sheet_name = sheet_name_original[:31]
            
            try:
                # Abre o workbook de origem
                source_workbook = app.books.open(file_path)
                
                # Verifica se tem pelo menos uma folha
                if len(source_workbook.sheets) == 0:
                    logger.warning(f"Ficheiro '{file_name}' não contém folhas. Ignorando.")
                    source_workbook.close()
                    continue
                
                # Obtém a primeira folha
                source_sheet = source_workbook.sheets[0]
                
                if i == 0:
                    # Para o primeiro arquivo, renomeamos a planilha padrão (Sheet1) para o nome correto
                    default_sheet.name = sheet_name
                    
                    # Copiamos o conteúdo da primeira planilha do arquivo fonte para a Sheet1 renomeada
                    # Copiar células, não a planilha inteira
                    used_range = source_sheet.used_range
                    if used_range:
                        data_to_copy = used_range.value
                        if data_to_copy:
                            default_sheet.range('A1').value = data_to_copy
                    
                    # Copiamos toda a formatação e propriedades da planilha
                    copy_sheet_properties(source_sheet, default_sheet)
                else:
                    # Para os arquivos subsequentes, adicionamos uma nova planilha
                    new_sheet = target_workbook.sheets.add(after=target_workbook.sheets[-1])
                    new_sheet.name = sheet_name
                    
                    # Copiamos o conteúdo
                    used_range = source_sheet.used_range
                    if used_range:
                        data_to_copy = used_range.value
                        if data_to_copy:
                            new_sheet.range('A1').value = data_to_copy
                    
                    # Copiamos toda a formatação e propriedades da planilha
                    copy_sheet_properties(source_sheet, new_sheet)
                
                # Fecha o workbook de origem sem salvar
                source_workbook.close()
                
                logger.info(f"Ficheiro '{file_name}' (folha '{sheet_name}') adicionado com sucesso.")
                processed_any_file = True
                
                # Pequena pausa para dar tempo ao Excel processar
                time.sleep(0.5)
                
            except Exception as e:
                logger.error(f"Erro ao processar ficheiro '{file_name}': {e}", exc_info=True)
                # Tenta fechar o workbook de origem se aberto
                if 'source_workbook' in locals():
                    try:
                        source_workbook.close()
                    except:
                        pass
        
        if not processed_any_file:
            logger.warning(f"Nenhum ficheiro Excel foi processado com sucesso para '{output_file_path}'.")
            try:
                target_workbook.close()
            except:
                pass
            return False
        
        # Salva o workbook consolidado (agora com o nome original sem sufixo _BD)
        try:
            target_workbook.save(output_file_path)
            logger.info(f"Ficheiros Excel de '{source_folder}' consolidados com sucesso em '{output_file_path}'.")
            target_workbook.close()
            return True
        except Exception as e:
            logger.error(f"Erro ao salvar ficheiro consolidado '{output_file_path}': {e}", exc_info=True)
            return False
            
    except Exception as e:
        logger.error(f"Falha crítica na consolidação de arquivos Excel: {e}", exc_info=True)
        return False
        
    finally:
        # Garante que o Excel seja fechado corretamente
        if app:
            try:
                app.quit()
            except:
                pass 