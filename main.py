import os
import sys
import logging
import json
from datetime import datetime
from colorama import init, Fore, Style # Para cores no terminal
from src.utils import setup_logging # Para configurar o logging centralmente
from src.excel_handler import consolidate_excel_files
from src.excel_converter import ExcelConverter # Importação da classe ExcelConverter
from src.data_validator import CustomJSONEncoder
from src.dimension_consolidator import DimensionConsolidator  # Nova importação

# Inicializa colorama para cores no terminal
init()

# Configurar o logger principal para toda a aplicação
# verbose=True pode ser ajustado conforme necessário ou tornado configurável
logger = setup_logging(verbose=True)

# Definir caminhos base relativos à raiz do projeto
# Estrutura de pastas:
# dataset/validation/1-96/quadros (entrada para opção 1)
# dataset/validation/1-96/series (entrada para opção 2)
# dataset/main (entrada para opções 3 e 4)
# result/validation/1-96/quadros (saída para opção 1)
# result/validation/1-96/series (saída para opção 2)
# result/main (saída para opções 3 e 4)
BASE_DATASET_PATH = os.path.join("dataset", "validation")
BASE_CONVERTED_PATH = os.path.join("result", "validation")
BASE_DATASET_MAIN_PATH = os.path.join("dataset", "main")
BASE_RESULT_MAIN_PATH = os.path.join("result", "main")

def ensure_all_directories():
    """Garante que todos os diretórios necessários existam."""
    from src.utils import ensure_directory_exists
    
    # Garante apenas os diretórios mínimos necessários
    # Para criar toda a estrutura, use python -m src.initialize
    ensure_directory_exists(BASE_DATASET_MAIN_PATH)
    ensure_directory_exists(BASE_RESULT_MAIN_PATH)
    ensure_directory_exists("logs")
    
    # Cria as pastas de 1 a 96 (exceto a 55 que já existe) com subpastas
    for folder_num in range(1, 97):
        # Pula a pasta 55 que já existe
        if folder_num == 55:
            continue
            
        # Cria as subpastas quadros e series para dataset (entrada)
        input_quadros = os.path.join(BASE_DATASET_PATH, str(folder_num), "quadros")
        input_series = os.path.join(BASE_DATASET_PATH, str(folder_num), "series")
        
        # Cria as subpastas quadros e series para result (saída)
        output_quadros = os.path.join(BASE_CONVERTED_PATH, str(folder_num), "quadros")
        output_series = os.path.join(BASE_CONVERTED_PATH, str(folder_num), "series")
        
        # Cria as pastas
        ensure_directory_exists(input_quadros)
        ensure_directory_exists(input_series)
        ensure_directory_exists(output_quadros)
        ensure_directory_exists(output_series)
    
    logger.info("Diretórios básicos criados. Para criar toda a estrutura, execute 'python -m src.initialize'")

def display_header():
    """Exibe o cabeçalho do programa com formatação moderna."""
    os.system('cls' if os.name == 'nt' else 'clear')  # Limpa a tela
    print(f"\n{Fore.CYAN}{'='*80}{Style.RESET_ALL}")
    print(f"{Fore.CYAN}  SISTEMA DE WORKFLOW ETL - PROCESSAMENTO DE DADOS EXCEL{Style.RESET_ALL}")
    print(f"{Fore.CYAN}  Consolidação e Conversão com Garantia de Integridade{Style.RESET_ALL}")
    print(f"{Fore.CYAN}{'='*80}{Style.RESET_ALL}\n")

def display_menu():
    """Apresenta o menu de opções ao utilizador com formatação moderna."""
    display_header()
    print(f"{Fore.GREEN}[Menu Principal]{Style.RESET_ALL} Por favor, escolha uma das seguintes opções:\n")
    
    print(f"{Fore.YELLOW}Conversão de Formatos:{Style.RESET_ALL}")
    print(f"  {Fore.WHITE}1.{Style.RESET_ALL} Converter ficheiros Excel para CSV (em {BASE_DATASET_MAIN_PATH})")
    print(f"  {Fore.WHITE}2.{Style.RESET_ALL} Converter ficheiros Excel para JSON (em {BASE_DATASET_MAIN_PATH})")
    
    print(f"\n{Fore.YELLOW}Consolidação de Ficheiros Excel:{Style.RESET_ALL}")
    print(f"  {Fore.WHITE}3.{Style.RESET_ALL} Fundir os ficheiros dos quadros")
    print(f"  {Fore.WHITE}4.{Style.RESET_ALL} Fundir os ficheiros das series")
    
    print(f"\n{Fore.YELLOW}Consolidação Inteligente de Dimensões:{Style.RESET_ALL}")
    print(f"  {Fore.WHITE}5.{Style.RESET_ALL} Consolidar colunas de dimensão automaticamente")
    print(f"  {Fore.WHITE}6.{Style.RESET_ALL} Consolidar colunas de dimensão interativamente")
    
    print(f"\n{Fore.YELLOW}Validação de Dados:{Style.RESET_ALL}")
    print(f"  {Fore.WHITE}7.{Style.RESET_ALL} Validar integridade dos dados")
    print(f"  {Fore.WHITE}8.{Style.RESET_ALL} Analisar valores em falta")
    
    print(f"\n{Fore.YELLOW}Sistema:{Style.RESET_ALL}")
    print(f"  {Fore.WHITE}0.{Style.RESET_ALL} Sair do programa")
    
    print(f"\n{Fore.CYAN}{'='*80}{Style.RESET_ALL}")

def select_folder_number():
    """
    Apresenta um menu para o usuário selecionar um número de pasta (1 a 96).
    Mostra apenas pastas que contêm arquivos Excel (.xlsx ou .xls).
    
    Returns:
        str: O número da pasta selecionada ou None se cancelado
    """
    display_header()
    print(f"{Fore.GREEN}[Seleção de Pasta]{Style.RESET_ALL} Escolha uma pasta (1 a 96):\n")
    
    # Verifica quais pastas contêm arquivos Excel
    folders_with_excel = []
    
    # Define se é "quadros" ou "series" com base no contexto do logger
    if "quadros" in str(logger.getChild('context')):
        sub_folder = "quadros"
    else:
        sub_folder = "series"
    
    # Verifica cada pasta de 1 a 96
    for folder_num in range(1, 97):
        folder_path = os.path.join(BASE_DATASET_PATH, str(folder_num), sub_folder)
        
        # Verifica se a pasta existe
        if os.path.exists(folder_path):
            # Verifica se há arquivos Excel na pasta
            excel_files = [f for f in os.listdir(folder_path) 
                          if f.lower().endswith(('.xlsx', '.xls')) and
                          os.path.isfile(os.path.join(folder_path, f))]
            
            if excel_files:
                folders_with_excel.append(folder_num)
    
    # Se não encontrar nenhuma pasta com arquivos Excel
    if not folders_with_excel:
        print(f"{Fore.YELLOW}Nenhuma pasta contém arquivos Excel (.xlsx ou .xls).{Style.RESET_ALL}")
        print("Por favor, adicione arquivos Excel antes de continuar.")
        input("\nPressione Enter para voltar ao menu principal...")
        return None
    
    # Divide a visualização em linhas com 10 números cada para melhor visualização
    current_row = []
    for i in range(0, 100, 10):
        row = []
        for j in range(i, min(i+10, 97)):
            # Se a pasta contém arquivos Excel, mostra o número
            if j in folders_with_excel:
                # Destacar alguns números importantes como exemplo
                if j in [1, 55, 96]:
                    row.append(f"{Fore.YELLOW}{j}{Style.RESET_ALL}")
                else:
                    row.append(f"{j}")
            else:
                # Pasta sem arquivos Excel, mostra em cinza
                row.append(f"{Fore.BLACK}{j}{Style.RESET_ALL}")
        
        # Adiciona a linha apenas se tiver algum número visível
        if any(j in folders_with_excel for j in range(i, min(i+10, 97))):
            print("  " + " ".join(row))
    
    print(f"\n  {Fore.WHITE}0.{Style.RESET_ALL} Voltar ao menu principal")
    
    choice = input(f"\n{Fore.GREEN}>>{Style.RESET_ALL} Digite o número da pasta (1-96): ")
    logger.info(f"O utilizador escolheu a pasta: '{choice}'")
    
    if choice == "0":
        logger.info("O utilizador optou por voltar ao menu principal")
        return None
    
    try:
        folder_num = int(choice)
        if 1 <= folder_num <= 96:
            # Verifica se a pasta escolhida contém arquivos Excel
            if folder_num in folders_with_excel:
                return str(folder_num)
            else:
                print(f"\n{Fore.RED}Erro:{Style.RESET_ALL} A pasta {folder_num} não contém arquivos Excel.")
                input("\nPressione Enter para tentar novamente...")
                return select_folder_number()
        else:
            print(f"\n{Fore.RED}Erro:{Style.RESET_ALL} O número deve estar entre 1 e 96.")
            input("\nPressione Enter para tentar novamente...")
            return select_folder_number()
    except ValueError:
        print(f"\n{Fore.RED}Erro:{Style.RESET_ALL} Por favor, digite um número válido.")
        input("\nPressione Enter para tentar novamente...")
        return select_folder_number()

def prompt_for_folder(for_conversion=False):
    """
    Solicita ao utilizador para escolher a pasta a processar.
    
    Args:
        for_conversion: Se True, mostra opções para conversão CSV/JSON,
                        caso contrário, mostra opções para consolidação
    
    Returns:
        Tuple com (input_dir, output_dir, folder_id) ou (None, None, None) se cancelado
    """
    if for_conversion:
        # Para conversão CSV/JSON (opções 3 e 4)
        print(f"\n{Fore.GREEN}[Seleção de Pasta]{Style.RESET_ALL} Escolha a pasta a processar:\n")
        print(f"  {Fore.WHITE}1.{Style.RESET_ALL} Pasta principal (em {BASE_DATASET_MAIN_PATH})")
        print(f"  {Fore.WHITE}2.{Style.RESET_ALL} Outra pasta (introduzir caminho)")
        print(f"  {Fore.WHITE}0.{Style.RESET_ALL} Voltar ao menu principal")
        
        choice = input(f"\n{Fore.GREEN}>>{Style.RESET_ALL} Digite o número da sua escolha: ")
        logger.info(f"Utilizador escolheu a opção: '{choice}' para pasta a processar na conversão")
        
        if choice == "0":
            logger.info("Utilizador optou por voltar ao menu principal")
            return None, None, None
            
        if choice == "1":
            logger.info("Utilizador escolheu processar a pasta principal")
            input_dir = BASE_DATASET_MAIN_PATH
            output_dir = BASE_RESULT_MAIN_PATH
            folder_id = "main"  # Identificador para main
        elif choice == "2":
            logger.info("Utilizador escolheu processar uma pasta personalizada")
            custom_path = input(f"{Fore.GREEN}>>{Style.RESET_ALL} Introduza o caminho completo para a pasta: ").strip()
            logger.info(f"Caminho personalizado informado: '{custom_path}'")
            
            if not os.path.isdir(custom_path):
                logger.error(f"A pasta '{custom_path}' não existe ou não é um diretório.")
                print(f"\n{Fore.RED}ERRO:{Style.RESET_ALL} A pasta '{custom_path}' não foi encontrada. Verifique o caminho.")
                return None, None, None
                
            input_dir = custom_path
            # Pasta de saída é baseada no nome da pasta personalizada
            folder_name = os.path.basename(os.path.normpath(custom_path))
            output_dir = os.path.join("result", folder_name)
            folder_id = folder_name
            logger.info(f"Diretório de saída definido como: '{output_dir}'")
        else:
            logger.warning(f"Opção inválida selecionada: '{choice}'")
            print(f"\n{Fore.RED}ERRO:{Style.RESET_ALL} Opção '{choice}' é inválida. Por favor, tente novamente.")
            return None, None, None
    else:
        # Para consolidação (opções 1 e 2)
        # Primeiro, solicita o número da pasta (1 a 96)
        folder_num = select_folder_number()
        if folder_num is None:
            return None, None, None
            
        # Agora, determina se é quadros ou series baseado no botão pressionado anteriormente
        if "quadros" in str(logger.getChild('context')):
            sub_folder = "quadros"
        else:
            sub_folder = "series"
            
        input_dir = os.path.join(BASE_DATASET_PATH, folder_num, sub_folder)
        output_dir = os.path.join(BASE_CONVERTED_PATH, folder_num, sub_folder)
        folder_id = folder_num
        
        logger.info(f"Diretórios configurados - Entrada: '{input_dir}', Saída: '{output_dir}'")
    
    # Verifica se o diretório de entrada existe
    if not os.path.isdir(input_dir):
        logger.error(f"A pasta de origem '{input_dir}' não existe ou não é um diretório.")
        print(f"\n{Fore.RED}ERRO:{Style.RESET_ALL} A pasta de origem '{input_dir}' não foi encontrada.")
        print(f"Execute 'python -m src.initialize' para criar todas as pastas necessárias.")
        input("\nPressione Enter para continuar...")
        return None, None, None
        
    return input_dir, output_dir, folder_id

def handle_merge_files(file_type: str):
    """
    Executa o processo de fusão (consolidação) para um tipo específico (quadros ou series).
    
    Args:
        file_type: Tipo de arquivo a processar ("quadros" ou "series")
    """
    # Cria um contexto temporário para o logger saber qual tipo de arquivo estamos processando
    logger.getChild(f'context.{file_type}')
    
    display_header()
    print(f"{Fore.GREEN}[Fusão de Ficheiros]{Style.RESET_ALL} {file_type.capitalize()}\n")
    
    # Solicita ao utilizador para escolher a pasta (1 a 96)
    input_dir, output_dir, folder_id = prompt_for_folder(for_conversion=False)
    if input_dir is None or output_dir is None or folder_id is None:
        return
    
    output_file_base = f"{folder_id}_{file_type}"
    output_file_full_path = os.path.join(output_dir, f"{output_file_base}.xlsx")

    logger.info(f"A iniciar processo de fusão para a pasta {folder_id}/{file_type}")
    logger.info(f"Pasta de origem dos ficheiros: '{input_dir}'")
    logger.info(f"Pasta de destino para o ficheiro consolidado: '{output_dir}'")
    logger.info(f"Nome base do ficheiro de saída: '{output_file_base}.xlsx'")

    print(f"{Fore.CYAN}Origem:{Style.RESET_ALL} {input_dir}")
    print(f"{Fore.CYAN}Destino:{Style.RESET_ALL} {output_file_full_path}\n")

    try:
        print(f"{Fore.YELLOW}Processando...{Style.RESET_ALL}")
        success = consolidate_excel_files(input_dir, output_dir, output_file_base)
        if success:
            print(f"\n{Fore.GREEN}Fusão concluída com sucesso!{Style.RESET_ALL}")
            print(f"Ficheiro guardado em: '{output_file_full_path}'")
        else:
            print(f"\n{Fore.RED}A fusão encontrou problemas.{Style.RESET_ALL} Verifique os logs para mais detalhes.")
    except Exception as e:
        logger.error(f"Erro inesperado durante a fusão para '{file_type}' na pasta {folder_id}: {e}", exc_info=True)
        print(f"\n{Fore.RED}Ocorreu um erro crítico:{Style.RESET_ALL} {str(e)}")
        print("Consulte os logs para detalhes técnicos.")
    
    input("\nPressione Enter para continuar...")

def prompt_for_file_selection(excel_files_in_folder: list[str], input_dir: str) -> list[str] | None:
    """
    Apresenta um menu para o utilizador selecionar ficheiros específicos de uma lista.

    Args:
        excel_files_in_folder: Lista de nomes de ficheiros Excel.
        input_dir: Diretório de entrada onde os ficheiros estão localizados.

    Returns:
        Lista de caminhos completos dos ficheiros selecionados ou None se cancelado.
    """
    if not excel_files_in_folder:
        logger.warning("Nenhum ficheiro Excel para selecionar.")
        print(f"{Fore.YELLOW}Nenhum ficheiro Excel encontrado na pasta selecionada.{Style.RESET_ALL}")
        return None

    selected_file_paths = []
    while True:
        display_header()
        print(f"{Fore.GREEN}[Seleção de Ficheiros para Conversão]{Style.RESET_ALL}")
        print(f"Ficheiros disponíveis em: {Fore.CYAN}{input_dir}{Style.RESET_ALL}")

        for idx, file_name in enumerate(excel_files_in_folder):
            print(f"  {Fore.WHITE}{idx + 1}.{Style.RESET_ALL} {file_name}")

        print(f"  {Fore.WHITE}T.{Style.RESET_ALL} Selecionar todos os ficheiros")
        print(f"  {Fore.WHITE}0.{Style.RESET_ALL} Voltar ao menu anterior")

        choice = input(f"{Fore.GREEN}>>{Style.RESET_ALL} Digite os números dos ficheiros (ex: 1,3,5), 'T' para todos, ou '0' para voltar: ").strip().lower()
        logger.info(f"Utilizador escolheu os ficheiros: '{choice}' para conversão.")

        if choice == '0':
            logger.info("Utilizador optou por voltar.")
            return None
        elif choice == 't':
            logger.info("Utilizador selecionou todos os ficheiros.")
            selected_file_paths = [os.path.join(input_dir, f) for f in excel_files_in_folder]
            break
        else:
            try:
                chosen_indices = []
                parts = choice.split(',')
                valid_selection = True
                for part in parts:
                    part = part.strip()
                    if not part: continue # Ignora partes vazias (ex: 1,,2)
                    
                    idx = int(part) - 1
                    if 0 <= idx < len(excel_files_in_folder):
                        if idx not in chosen_indices: # Evita duplicados
                            chosen_indices.append(idx)
                        else:
                            logger.warning(f"Número de ficheiro duplicado '{part}' na seleção.")
                            # Não é um erro fatal, apenas um aviso.
                    else:
                        print(f"{Fore.RED}Erro:{Style.RESET_ALL} Número de ficheiro '{part}' é inválido.")
                        logger.warning(f"Seleção de ficheiro inválida: '{part}'. Índice fora do intervalo.")
                        valid_selection = False
                        break
                
                if valid_selection and chosen_indices:
                    selected_file_paths = [os.path.join(input_dir, excel_files_in_folder[i]) for i in chosen_indices]
                    # Ordenar com base na ordem original da lista de ficheiros
                    selected_file_paths.sort(key=lambda x: excel_files_in_folder.index(os.path.basename(x)))
                    logger.info(f"Ficheiros selecionados para processamento: {selected_file_paths}")
                    break
                elif not chosen_indices and valid_selection: # Se o input foi apenas vírgulas ou espaços
                     print(f"{Fore.RED}Erro:{Style.RESET_ALL} Nenhuma seleção válida foi feita.")
                     logger.warning("Nenhum ficheiro válido selecionado após o parse.")
                elif not valid_selection: # Erro já foi impresso
                    pass


            except ValueError:
                print(f"{Fore.RED}Erro:{Style.RESET_ALL} Entrada inválida. Por favor, use números separados por vírgula, 'T' ou '0'.")
                logger.warning(f"Entrada inválida para seleção de ficheiros: '{choice}'.")
            
            input("Pressione Enter para tentar novamente...")

    return selected_file_paths

def handle_conversion(output_format: str):
    """
    Executa o processo de conversão dos ficheiros Excel para CSV ou JSON.
    
    Args:
        output_format: O formato de saída ('csv' ou 'json')
    """
    display_header()
    print(f"{Fore.GREEN}[Conversão para {output_format.upper()}]{Style.RESET_ALL}\n")
    
    # Solicita ao utilizador para escolher a pasta (usando a função atualizada com for_conversion=True)
    input_dir, output_dir, _ = prompt_for_folder(for_conversion=True)
    if input_dir is None or output_dir is None:
        return
    
    # Inicializa o conversor e executa a conversão
    try:
        print(f"\n{Fore.CYAN}Informações:{Style.RESET_ALL}")
        print(f"Formato de saída: {output_format.upper()}")
        print(f"Pasta de origem: {input_dir}")
        print(f"Pasta de destino: {output_dir}")
        
        # Lista os arquivos Excel na pasta de entrada para verificar
        excel_files_in_folder = [f for f in os.listdir(input_dir) 
                                 if f.lower().endswith(('.xlsx', '.xls')) and
                                 os.path.isfile(os.path.join(input_dir, f))]
        
        if not excel_files_in_folder:
            logger.warning(f"Nenhum arquivo Excel (.xlsx, .xls) encontrado em '{input_dir}'")
            print(f"{Fore.YELLOW}AVISO:{Style.RESET_ALL} Nenhum arquivo Excel foi encontrado na pasta de origem.")
            print("Verifique se há arquivos para processar.")
            input("Pressione Enter para continuar...")
            return

        # Solicita ao utilizador para selecionar os ficheiros
        selected_file_paths = prompt_for_file_selection(excel_files_in_folder, input_dir)

        if not selected_file_paths:
            logger.info("Nenhum ficheiro selecionado para conversão ou utilizador cancelou.")
            # A mensagem de cancelamento/nenhum ficheiro já é tratada em prompt_for_file_selection ou no fluxo normal
            # se prompt_for_file_selection retornar None.
            # Se prompt_for_file_selection retorna None, significa que o utilizador escolheu voltar.
            # Se retorna uma lista vazia (que não deve acontecer com a lógica atual), seria um erro.
            # A função prompt_for_file_selection já lida com o input para voltar ao menu.
            return

        logger.info(f"Ficheiros selecionados para conversão: {selected_file_paths}")
        
        print(f"{Fore.CYAN}Ficheiros selecionados para conversão:{Style.RESET_ALL} {len(selected_file_paths)}")
        for idx, file_path in enumerate(selected_file_paths[:5], 1):
            print(f"  {idx}. {os.path.basename(file_path)}")
        if len(selected_file_paths) > 5:
            print(f"  ... mais {len(selected_file_paths) - 5} ficheiros")
            
        print(f"{Fore.YELLOW}Iniciando conversão...{Style.RESET_ALL}")
        
        # Cria uma instância do ExcelConverter
        logger.info(f"Criando instância do ExcelConverter para '{input_dir}' -> '{output_dir}'")
        converter = ExcelConverter(input_dir, output_dir, logger)
        
        # Processa os ficheiros selecionados individualmente
        results = []
        logger.info(f"Iniciando processamento dos {len(selected_file_paths)} ficheiros selecionados para {output_format}")
        
        for file_path_to_process in selected_file_paths:
            logger.info(f"Processando ficheiro individual: {file_path_to_process}")
            # A função process_excel_file espera o caminho completo do ficheiro.
            # sheet_name pode ser deixado como None para que ExcelConverter decida (primeira ou 'dados').
            result = converter.process_excel_file(file_path_to_process, output_format)
            results.append(result)
            if "error" in result:
                logger.error(f"Erro ao converter {os.path.basename(file_path_to_process)}: {result['error']}")
            else:
                logger.info(f"Ficheiro {os.path.basename(file_path_to_process)} convertido com sucesso.")

        logger.info(f"Processamento de conversão concluído, {len(results)} ficheiros tentados.")
        
        # Exibe um resumo dos resultados
        successful = sum(1 for r in results if "error" not in r)
        logger.info(f"Processamento bem-sucedido para {successful}/{len(results)} arquivos")
        
        print(f"\n{Fore.GREEN}Processamento concluído!{Style.RESET_ALL}")
        print(f"Resultado: {successful}/{len(results)} ficheiros convertidos com sucesso.")
        
        if successful > 0:
            print(f"\n{Fore.CYAN}Ficheiros convertidos:{Style.RESET_ALL}")
            for result in results[:5]:
                if "error" not in result:
                    print(f"  • {os.path.basename(result['file_path'])} -> {os.path.basename(result['output_file'])}")
            if successful > 5:
                print(f"  ... mais {successful - 5} ficheiros")
        
        if successful < len(results):
            print(f"\n{Fore.RED}Erros encontrados:{Style.RESET_ALL}")
            error_count = 0
            for result in results:
                if "error" in result:
                    print(f"  • {os.path.basename(result['file_path'])}: {result['error']}")
                    error_count += 1
                    if error_count >= 5 and len(results) - successful > 5:
                        print(f"  ... mais {len(results) - successful - 5} erros")
                        break
        
        # Salva um resumo JSON se houver muitos resultados
        if len(results) > 5:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            summary_path = os.path.join(output_dir, f"conversion_summary_{timestamp}.json")
            
            summary = {
                "timestamp": datetime.now().isoformat(),
                "total_files": len(results),
                "successful": successful,
                "failed": len(results) - successful,
                "format": output_format,
                "results": results
            }
            
            with open(summary_path, 'w', encoding='utf-8') as f:
                json.dump(summary, f, indent=2, ensure_ascii=False, cls=CustomJSONEncoder)
                
            print(f"\n{Fore.CYAN}Resumo detalhado salvo em:{Style.RESET_ALL} {summary_path}")
        
        print(f"\n{Fore.CYAN}Pasta de destino:{Style.RESET_ALL} {output_dir}")
        
    except Exception as e:
        logger.error(f"Erro durante a conversão para {output_format}: {e}", exc_info=True)
        print(f"\n{Fore.RED}Ocorreu um erro crítico:{Style.RESET_ALL} {str(e)}")
        print("Consulte os logs para detalhes técnicos.")
    
    input("\nPressione Enter para continuar...")

def handle_dimension_consolidation():
    """
    Executa o processo de consolidação inteligente de dimensões.
    """
    display_header()
    print(f"{Fore.GREEN}[Consolidação Inteligente de Dimensões]{Style.RESET_ALL}\n")
    print("Esta funcionalidade analisa e consolida automaticamente colunas de dimensão similares,")
    print("preservando TODOS os valores exatamente como aparecem nos dados originais.")
    print(f"{Fore.CYAN}COMPATÍVEL COM FERRAMENTA OLAP:{Style.RESET_ALL} Mantém todos os valores para recriação de relatórios.\n")
    
    # Solicita ao utilizador para escolher o ficheiro
    input_dir, output_dir, _ = prompt_for_folder(for_conversion=True)
    if input_dir is None or output_dir is None:
        return
    
    # Lista ficheiros Excel na pasta de entrada
    try:
        excel_files_in_folder = [f for f in os.listdir(input_dir) 
                                 if f.lower().endswith(('.xlsx', '.xls')) and
                                 os.path.isfile(os.path.join(input_dir, f))]
    except Exception as e:
        logger.error(f"Erro ao listar ficheiros em '{input_dir}': {e}")
        print(f"\n{Fore.RED}Erro:{Style.RESET_ALL} Não foi possível acessar a pasta '{input_dir}'.")
        input("\nPressione Enter para continuar...")
        return
    
    if not excel_files_in_folder:
        logger.warning(f"Nenhum arquivo Excel (.xlsx, .xls) encontrado em '{input_dir}'")
        print(f"{Fore.YELLOW}AVISO:{Style.RESET_ALL} Nenhum arquivo Excel foi encontrado na pasta selecionada.")
        print("Adicione ficheiros Excel antes de continuar.")
        input("\nPressione Enter para continuar...")
        return
    
    # Permite seleção de ficheiro específico
    selected_file_paths = prompt_for_file_selection(excel_files_in_folder, input_dir)
    
    if not selected_file_paths:
        logger.info("Nenhum ficheiro selecionado para consolidação de dimensões.")
        return
    
    # Para consolidação de dimensões, processamos apenas um ficheiro de cada vez
    if len(selected_file_paths) > 1:
        print(f"\n{Fore.YELLOW}Aviso:{Style.RESET_ALL} A consolidação de dimensões processa um ficheiro de cada vez.")
        print(f"Será processado o primeiro ficheiro selecionado: {os.path.basename(selected_file_paths[0])}")
        print("Para processar outros ficheiros, execute esta opção novamente.")
        input("\nPressione Enter para continuar...")
    
    input_file = selected_file_paths[0]
    
    # Opções de consolidação
    print(f"\n{Fore.GREEN}[Opções de Consolidação]{Style.RESET_ALL}")
    print(f"Ficheiro a processar: {Fore.CYAN}{os.path.basename(input_file)}{Style.RESET_ALL}")
    print(f"Pasta de saída: {Fore.CYAN}{output_dir}{Style.RESET_ALL}\n")
    
    print("Escolha o modo de execução:")
    print(f"  {Fore.WHITE}1.{Style.RESET_ALL} Consolidação completa (aplica alterações)")
    print(f"  {Fore.WHITE}2.{Style.RESET_ALL} Modo de simulação (apenas visualiza o que seria feito)")
    print(f"  {Fore.WHITE}0.{Style.RESET_ALL} Voltar ao menu anterior")
    
    mode_choice = input(f"\n{Fore.GREEN}>>{Style.RESET_ALL} Digite a sua escolha: ").strip()
    logger.info(f"Modo de consolidação escolhido: '{mode_choice}'")
    
    if mode_choice == "0":
        logger.info("Utilizador optou por voltar ao menu anterior")
        return
    elif mode_choice == "1":
        dry_run = False
        mode_desc = "consolidação completa"
    elif mode_choice == "2":
        dry_run = True
        mode_desc = "simulação"
    else:
        print(f"\n{Fore.RED}Erro:{Style.RESET_ALL} Opção inválida.")
        input("\nPressione Enter para continuar...")
        return
    
    # Configurações avançadas (opcional)
    exclude_columns = []
    print(f"\n{Fore.YELLOW}Configurações Avançadas (opcional):{Style.RESET_ALL}")
    exclude_input = input("Colunas a excluir da consolidação (separadas por vírgula, ou Enter para nenhuma): ").strip()
    
    if exclude_input:
        exclude_columns = [col.strip() for col in exclude_input.split(',') if col.strip()]
        logger.info(f"Colunas a excluir: {exclude_columns}")
    
    try:
        print(f"\n{Fore.CYAN}Informações:{Style.RESET_ALL}")
        print(f"Modo: {mode_desc}")
        print(f"Ficheiro: {os.path.basename(input_file)}")
        print(f"Pasta de saída: {output_dir}")
        if exclude_columns:
            print(f"Colunas excluídas: {', '.join(exclude_columns)}")
        
        print(f"\n{Fore.YELLOW}Iniciando {mode_desc}...{Style.RESET_ALL}")
        
        # Cria instância do consolidador
        logger.info(f"Criando instância do DimensionConsolidator para '{input_file}' -> '{output_dir}'")
        consolidator = DimensionConsolidator(input_file, output_dir, logger)
        
        # Executa consolidação
        logger.info(f"Iniciando consolidação de dimensões ({'simulação' if dry_run else 'execução real'})")
        result_df = consolidator.consolidate(dry_run=dry_run, exclude_columns=exclude_columns)
        
        # Exibe resumo
        print(f"\n{Fore.GREEN}Processamento concluído!{Style.RESET_ALL}")
        consolidator.print_summary()
        
        # Guarda resultados se não for simulação
        if not dry_run:
            print(f"{Fore.CYAN}Guardando resultados...{Style.RESET_ALL}")
            
            # Pergunta formato de saída
            print("Escolha o formato de saída:")
            print(f"  {Fore.WHITE}1.{Style.RESET_ALL} Excel (.xlsx)")
            print(f"  {Fore.WHITE}2.{Style.RESET_ALL} CSV (.csv)")
            print(f"  {Fore.WHITE}3.{Style.RESET_ALL} JSON (.json)")
            
            format_choice = input(f"\n{Fore.GREEN}>>{Style.RESET_ALL} Digite a sua escolha (padrão: Excel): ").strip()
            
            format_map = {'1': 'excel', '2': 'csv', '3': 'json', '': 'excel'}
            output_format = format_map.get(format_choice, 'excel')
            
            try:
                output_file = consolidator.save_results(format_type=output_format)
                print(f"\n{Fore.GREEN}Resultados guardados em:{Style.RESET_ALL} {output_file}")
                
                # Guarda relatório específico de preservação de valores para Ferramenta OLAP
                preservation_report_file = consolidator.save_value_preservation_report()
                print(f"{Fore.CYAN}Relatório de preservação (Ferramenta OLAP):{Style.RESET_ALL} {preservation_report_file}")
                
                # Mostra mapeamento de consolidações realizadas
                consolidation_mapping = consolidator.get_consolidation_mapping()
                if consolidation_mapping:
                    print(f"\n{Fore.CYAN}Consolidações realizadas:{Style.RESET_ALL}")
                    for new_col, original_cols in consolidation_mapping.items():
                        print(f"  • {Fore.YELLOW}{new_col}{Style.RESET_ALL} ← {', '.join(original_cols)}")
                else:
                    print(f"\n{Fore.YELLOW}Nenhuma consolidação foi realizada.{Style.RESET_ALL}")
                    print("Os dados não continham padrões adequados para consolidação automática.")
                
                # Mostra informações de preservação de valores
                preservation_report = consolidator.get_value_preservation_report()
                preservation_summary = preservation_report.get('preservation_summary', {})
                
                if preservation_summary.get('total_values_preserved', 0) > 0:
                    print(f"\n{Fore.GREEN}[OK] Preservação de Valores:{Style.RESET_ALL}")
                    print(f"  • {preservation_summary['total_values_preserved']} valores únicos preservados")
                    print(f"  • {preservation_summary['total_consolidations']} consolidações realizadas")
                    print(f"  • {Fore.GREEN}Compatível com Ferramenta OLAP{Style.RESET_ALL} para recriação de relatórios")
                
            except Exception as save_error:
                logger.error(f"Erro ao guardar resultados: {save_error}")
                print(f"\n{Fore.RED}Erro ao guardar resultados:{Style.RESET_ALL} {str(save_error)}")
        else:
            print(f"\n{Fore.YELLOW}Modo de simulação:{Style.RESET_ALL} Nenhum ficheiro foi alterado.")
            print("Execute novamente no modo de consolidação completa para aplicar as alterações.")
        
    except Exception as e:
        logger.error(f"Erro durante consolidação de dimensões: {e}", exc_info=True)
        print(f"\n{Fore.RED}Ocorreu um erro crítico:{Style.RESET_ALL} {str(e)}")
        print("Consulte os logs para detalhes técnicos.")
    
    input("\nPressione Enter para continuar...")

def handle_validation():
    """Executa o processo de validação de dados através de comparação inteligente."""
    display_header()
    print(f"{Fore.GREEN}[Validação de Integridade de Dados]{Style.RESET_ALL}\n")
    
    try:
        # Importa o módulo de comparação de dados
        from src.data_comparator import run_interactive_comparison
        
        # Executa a comparação interativa
        run_interactive_comparison(logger)
        
    except Exception as e:
        logger.error(f"Erro durante validação de dados: {e}", exc_info=True)
        print(f"\n{Fore.RED}Ocorreu um erro crítico:{Style.RESET_ALL} {str(e)}")
        print("Consulte os logs para detalhes técnicos.")
    
    input("\nPressione Enter para continuar...")

def handle_missing_values_analysis():
    """Executa o processo de análise de valores em falta."""
    display_header()
    print(f"{Fore.GREEN}[Análise de Valores em Falta]{Style.RESET_ALL}\n")
    
    try:
        # Importa o módulo de análise de valores em falta
        from src.missing_values_analyzer import run_missing_values_analysis
        
        # Executa a análise interativa
        run_missing_values_analysis(logger)
        
    except Exception as e:
        logger.error(f"Erro durante análise de valores em falta: {e}", exc_info=True)
        print(f"\n{Fore.RED}Ocorreu um erro crítico:{Style.RESET_ALL} {str(e)}")
        print("Consulte os logs para detalhes técnicos.")
    
    input("\nPressione Enter para continuar...")

def consolidate_dimensions_interactive():
    """Executa o processo interativo de consolidação de dimensões."""
    try:
        # Importa o módulo de consolidação interativa
        from src.interactive_consolidation import InteractiveConsolidator, get_input_file
        from src.utils import setup_logging
        
        # Setup do logger
        logger = setup_logging(verbose=True)
        
        display_header()
        print(f"{Fore.GREEN}[Consolidação Interativa de Dimensões]{Style.RESET_ALL}\n")
        print("Esta funcionalidade permite consolidar colunas de dimensão com controlo total sobre o processo.")
        print("Compatível com qualquer ficheiro Excel seguindo a estrutura: dim_1 | dim_2 | ... | indicador | unidade | valor\n")
        
        # Passo 1: Obtém ficheiro de entrada
        input_file = get_input_file()
        if not input_file:
            print("Operação cancelada.")
            input("\nPressione Enter para continuar...")
            return
        
        # Passo 2: Define diretório de saída
        output_dir = "result/interactive_consolidation"
        
        # Passo 3: Inicializa o consolidador
        print(f"\n{Fore.CYAN}Inicializando consolidador interativo...{Style.RESET_ALL}")
        consolidator = InteractiveConsolidator(input_file, output_dir, logger)
        
        # Passo 4: Carrega e analisa dados
        if not consolidator.load_and_analyze_data():
            input("\nPressione Enter para continuar...")
            return
        
        # Loop principal de consolidação
        while True:
            # Passo 5: Apresenta dimensões disponíveis
            dim_columns = consolidator.display_dimensions()
            
            if not dim_columns:
                print(f"\n{Fore.YELLOW}Nenhuma dimensão encontrada para consolidar.{Style.RESET_ALL}")
                input("\nPressione Enter para continuar...")
                return
            
            # Passo 6: Apresenta instruções
            consolidator.display_consolidation_instructions()
            
            # Passo 7: Obtém entrada do utilizador
            user_input = consolidator.get_user_consolidation_input(dim_columns)
            
            # Passo 8: Analisa entrada e cria plano de consolidação
            consolidation_plan, final_column_order = consolidator.parse_consolidation_input(user_input, dim_columns)
            
            if not consolidation_plan and not final_column_order:
                print(f"\n{Fore.YELLOW}Nenhuma consolidação válida especificada. Tente novamente.{Style.RESET_ALL}")
                continue
            
            # Passo 9: Apresenta resumo do plano
            consolidator.display_consolidation_summary(consolidation_plan, final_column_order, dim_columns)
            
            # Passo 10: Confirma plano com utilizador
            confirmation = consolidator.confirm_consolidation_plan()
            
            if confirmation is None:  # Refazer
                continue
            elif not confirmation:  # Cancelar
                return
            else:  # Confirmar
                break
        
        # Passo 11: Guarda plano no consolidador e aplica
        consolidator.consolidation_plan = consolidation_plan
        consolidator.final_column_order = final_column_order
        
        print(f"\n{Fore.CYAN}Aplicando consolidação...{Style.RESET_ALL}")
        
        if not consolidator.apply_consolidation():
            print(f"\n{Fore.RED}Falha na aplicação da consolidação.{Style.RESET_ALL}")
            input("\nPressione Enter para continuar...")
            return
        
        # Passo 12: Apresenta resumo final
        consolidator.print_summary()
        
        # Passo 13: Pergunta formato de saída
        print(f"\n{Fore.CYAN}Guardando resultados...{Style.RESET_ALL}")
        print("Escolha o formato de saída:")
        print(f"  {Fore.WHITE}1.{Style.RESET_ALL} Excel (.xlsx)")
        print(f"  {Fore.WHITE}2.{Style.RESET_ALL} CSV (.csv)")
        print(f"  {Fore.WHITE}3.{Style.RESET_ALL} JSON (.json)")
        
        format_choice = input(f"\n{Fore.GREEN}>>{Style.RESET_ALL} Digite a sua escolha (padrão: Excel): ").strip()
        
        format_map = {'1': 'excel', '2': 'csv', '3': 'json', '': 'excel'}
        output_format = format_map.get(format_choice, 'excel')
        
        # Passo 14: Guarda resultados
        try:
            output_file = consolidator.save_results(format_type=output_format)
            print(f"\n{Fore.GREEN}✅ Resultados guardados em:{Style.RESET_ALL} {output_file}")
            
            # Mostra mapeamento de consolidações realizadas
            if consolidator.consolidation_mapping:
                print(f"\n{Fore.CYAN}Consolidações realizadas:{Style.RESET_ALL}")
                for new_col, original_cols in consolidator.consolidation_mapping.items():
                    print(f"  • {Fore.YELLOW}{new_col}{Style.RESET_ALL} ← {', '.join(original_cols)}")
            else:
                print(f"\n{Fore.YELLOW}Nenhuma consolidação foi necessária.{Style.RESET_ALL}")
                print("Todas as dimensões permaneceram separadas conforme especificado.")
            
        except Exception as save_error:
            logger.error(f"Erro ao guardar resultados: {save_error}")
            print(f"\n{Fore.RED}Erro ao guardar resultados:{Style.RESET_ALL} {str(save_error)}")
        
    except Exception as e:
        logger.error(f"Erro durante consolidação interativa: {e}", exc_info=True)
        print(f"\n{Fore.RED}Ocorreu um erro crítico:{Style.RESET_ALL} {str(e)}")
        print("Consulte os logs para detalhes técnicos.")
    
    input("\nPressione Enter para continuar...")

def main():
    """Função principal que executa o menu e interage com o utilizador."""
    logger.info("Aplicação ETL iniciada.")
    try:
        # Adiciona colorama para cores no terminal
        init()
        
        # Garante que os diretórios básicos existam
        ensure_all_directories()
        
        while True:
            display_menu()
            choice = input(f"\n{Fore.GREEN}>>{Style.RESET_ALL} Digite o número da sua escolha: ")
            logger.debug(f"Utilizador escolheu a opção: '{choice}'")

            if choice == '1':
                logger.info("Opção '1' selecionada: Converter para CSV")
                handle_conversion("csv")
            elif choice == '2':
                logger.info("Opção '2' selecionada: Converter para JSON")
                handle_conversion("json")
            elif choice == '3':
                logger.info("Opção '3' selecionada: Fundir os ficheiros dos quadros")
                handle_merge_files("quadros")
            elif choice == '4':
                logger.info("Opção '4' selecionada: Fundir os ficheiros das series")
                handle_merge_files("series")
            elif choice == '5':
                logger.info("Opção '5' selecionada: Consolidação inteligente de dimensões")
                handle_dimension_consolidation()
            elif choice == '6':
                logger.info("Opção '6' selecionada: Consolidar colunas de dimensão interativamente")
                consolidate_dimensions_interactive()
            elif choice == '7':
                logger.info("Opção '7' selecionada: Validar dados")
                handle_validation()
            elif choice == '8':
                logger.info("Opção '8' selecionada: Analisar valores em falta")
                handle_missing_values_analysis()
            elif choice == '0':
                logger.info("Opção '0' selecionada: Sair do programa.")
                display_header()
                print(f"{Fore.GREEN}Obrigado por utilizar o Sistema de Workflow ETL!{Style.RESET_ALL}")
                print("A aplicação será encerrada.")
                break
            else:
                logger.warning(f"Escolha inválida ('{choice}') feita pelo utilizador.")
                print(f"\n{Fore.RED}Opção '{choice}' é inválida.{Style.RESET_ALL} Por favor, escolha um número do menu.")
                input("\nPressione Enter para continuar...")
    except KeyboardInterrupt:
        logger.info("Programa interrompido pelo utilizador (Ctrl+C)")
        print(f"\n\n{Fore.YELLOW}Programa interrompido.{Style.RESET_ALL} Obrigado por utilizar o Sistema de Workflow ETL!")
    except Exception as e:
        logger.error(f"Erro não tratado na execução principal: {str(e)}", exc_info=True)
        print(f"\n{Fore.RED}Ocorreu um erro crítico:{Style.RESET_ALL} {str(e)}")
        print("Consulte os logs para detalhes técnicos.")
        
    print("\nA aplicação será encerrada em breve.")
    # Pequena pausa para garantir que as mensagens sejam exibidas
    import time
    time.sleep(1)

if __name__ == "__main__":
    main() 