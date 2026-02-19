import pandas as pd
import random

# Primeira Função
def generate_matriz_interactions(list_meds: list) -> dict:
   

    matriz_effects = {}
    
    # Percorre a lista para as linhas
    for med_line in list_meds:
        # Percorre a mesma lista para as colunas
        for med_column in list_meds:
            
            # Se não forem o mesmo medicamento, gera o valor
            if med_line != med_column:
                random_effect = random.randint(1, 6)
                matriz_effects[(med_line, med_column)] = random_effect
                
    
    return matriz_effects

# --- 2. Ler os Dados de Excel ---
lista_final = pd.read_excel('Dados.xlsx').iloc[:, 0].dropna().tolist()

# --- 3. Gerar o Dicionário de Interações ---
tabela_de_interacoes = generate_matriz_interactions(lista_final)



# Cria a grelha onde as linhas e as colunas são ambas a 'lista_final'
df_exportar = pd.DataFrame(index=lista_final, columns=lista_final)

# Preenche as "coordenadas" da grelha com os valores do dicionário
for (med_linha, med_coluna), efeito in tabela_de_interacoes.items():
    df_exportar.at[med_linha, med_coluna] = efeito

# Substitui os espaços vazios por um traço
df_exportar = df_exportar.fillna('0')

# Guarda o ficheiro no computador.
nome_ficheiro = 'Matriz_Completa_Resultados.xlsx'
df_exportar.to_excel(nome_ficheiro)