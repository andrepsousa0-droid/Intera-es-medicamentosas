import sys
from typing import List, Dict, Tuple
from Interacoes_medicamentosas import read_meds, gen_matrix, export_excel

def main() -> None:
    """
    Orquestra a execução do sistema de interações medicamentosas.

    Esta função serve como o ponto de entrada principal, coordenando a
    leitura do ficheiro de texto, a geração da matriz de interações e
    a exportação dos resultados para um ficheiro Excel. Em caso de erro,
    imprime uma mensagem clara e termina a execução de forma segura.

    Returns
    -------
    None
        Esta função não retorna nenhum valor.
    """
    input_file: str = "medicamentos.txt"
    output_file: str = "Interacoes_medicamentosas.xlsx"

    try:
        print("A iniciar o sistema de orquestração das interações medicamentosas...")
        
        print(f"A ler a lista de medicamentos a partir de '{input_file}'...")
        medications_list: List[str] = read_meds(input_file)
        
        print("Leitura concluída com sucesso. A processar e gerar a matriz de interações...")
        interactions_matrix: Dict[Tuple[str, str], int] = gen_matrix(medications_list)
        
        print(f"Matriz de interações gerada. A exportar os dados para '{output_file}'...")
        export_excel(interactions_matrix, medications_list, output_file)
        
        print("Processo de exportação concluído com sucesso. A terminar o sistema.")

    except Exception as critical_error:
        print(f"Erro crítico detetado na execução do programa: {critical_error}")
        print("A encerrar o sistema por motivos de segurança.")
        sys.exit(1)

if __name__ == "__main__":
    main()