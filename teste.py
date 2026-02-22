# 1. Importação dos módulos necessários (Objetivo 2)
import openpyxl
import random
import sys

# 2. Criação da primeira função: extrair dados do ficheiro de texto (Objetivo 1)
def read_meds(filename: str) -> list:
    # Tratamento de exceções caso o ficheiro não exista (Objetivo 3)
    try:
        with open(filename, 'r', encoding='utf-8') as f:
            return [line.replace("", "").strip() for line in f if line.strip()]
    except Exception as e:
        print(f"Error reading file: {e}")
        sys.exit(1)

# 3. Criação da segunda função: gerar matriz de interações (Objetivo 4 e 6)
# Usa estruturas mutáveis (dict) e imutáveis (tuplos) com tipificação
def gen_matrix(meds: list) -> dict:
    # Gera valores de 1 a 6, e coloca 0 quando cruza com o próprio medicamento
    return {(x, y): (0 if x == y else random.randint(1, 6)) for x in meds for y in meds}

# 4. Criação da terceira função: exportar dados para Excel (Objetivo 1)
def export_excel(matrix: dict, meds: list, filename: str) -> None:
    # Tratamento de exceções caso o Excel esteja aberto e bloqueie a gravação (Objetivo 3)
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        
        # Preenchimento das linhas e colunas
        for i, x in enumerate(meds, start=2):
            ws.cell(row=1, column=i, value=x) 
            ws.cell(row=i, column=1, value=x) 
            
            for j, y in enumerate(meds, start=2):
                ws.cell(row=i, column=j, value=matrix[(x, y)]) 
                
        wb.save(filename)
    except Exception as e:
        print(f"Error saving Excel: {e}")
        sys.exit(1)

# 5. Função principal: estruturação por decomposição funcional (Objetivo 5)
def main() -> None:
    meds: list = read_meds('medicamentos.txt')
    matrix: dict = gen_matrix(meds)
    export_excel(matrix, meds, 'Interacoes_medicamentosas.xlsx')

# Executa o código
if __name__ == "__main__":
    main()