import sys
from typing import List, Dict, Tuple
from logic import read_meds, gen_matrix, export_excel, convert_xlsx_to_csv

def main() -> None:
    """
    Orquestra a execução do sistema de interações medicamentosas.

    Esta função serve como o ponto de entrada principal, coordenando a
    leitura do ficheiro de texto, a geração da matriz de interações,
    a exportação dos resultados para um ficheiro Excel e a sua conversão
    para o formato CSV. Em caso de erro, imprime uma mensagem clara e
    termina a execução de forma segura.

    Returns
    -------
    None
        Esta função não retorna nenhum valor.
    """
    input_file: str = "medicamentos.txt"
    output_excel: str = "Interacoes_medicamentosas.xlsx"
    output_csv: str = "Interacoes_medicamentosas.csv"

    print("A iniciar o sistema de orquestração das interações medicamentosas...")

    try:
        print(f"A ler a lista de medicamentos a partir de '{input_file}'...")
        medications_list: List[str] = read_meds(input_file)
    except FileNotFoundError as file_error:
        print(f"Erro ao ler os medicamentos: {file_error}")
        sys.exit(1)
    except OSError as os_error:
        print(f"Erro de sistema ao abrir os medicamentos: {os_error}")
        sys.exit(1)

    try:
        print("Leitura concluída com sucesso. A processar e gerar a matriz de interações...")
        interactions_matrix: Dict[Tuple[str, str], int] = gen_matrix(medications_list)
    except Exception as matrix_error:
        print(f"Erro crítico ao gerar a matriz: {matrix_error}")
        sys.exit(1)

    try:
        print(f"Matriz de interações gerada. A exportar os dados para '{output_excel}'...")
        export_excel(interactions_matrix, medications_list, output_excel)
    except TypeError as type_error:
        print(f"Erro nos dados durante a exportação: {type_error}")
        sys.exit(1)
    except OSError as os_error:
        print(f"Erro de entrada/saída ao guardar o Excel: {os_error}")
        sys.exit(1)

    try:
        print(f"Processo de exportação concluído com sucesso. A iniciar a conversão para CSV: '{output_csv}'...")
        convert_xlsx_to_csv(output_excel, output_csv)
        print("Conversão concluída com sucesso. A terminar o sistema.")
    except FileNotFoundError as file_error:
        print(f"Erro ao localizar o ficheiro Excel para conversão: {file_error}")
        sys.exit(1)
    except OSError as io_error:
        print(f"Erro de leitura/escrita na conversão para CSV: {io_error}")
        sys.exit(1)
    except Exception as critical_error:
        print(f"Erro crítico e inesperado no processo final: {critical_error}")
        print("A encerrar o sistema por motivos de segurança.")
        sys.exit(1)

if __name__ == "__main__":
    main()

