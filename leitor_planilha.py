import pandas as pd

# Carregando a planilha
planilha = pd.read_excel('planilha.xlsx')

# Preenchendo os valores NaN com zero
planilha.fillna(0, inplace=True)

# Obtendo os meses
meses = planilha.iloc[:, 0].tolist()

# Obtendo os anos
anos = list(planilha.columns)[1:]

# Obtendo os valores das células
valores_celulas = []

# Iterar sobre as colunas da célula B até a coluna AU
for coluna in planilha.columns[1:47]:  # De B até AU são 46 colunas 
    # Lendo os valores da célula atual (B2 até AU13)
    valores = planilha.loc[0:12, coluna].tolist()
    valores_celulas.extend(valores)  # Adicionando os valores à lista

# Gerando o SQL
sql = ''
try:
    index_valores = 0  # Índice para iterar sobre valores_celulas
    for i, ano in enumerate(anos):
        sql += f"--- INCP {ano}\n\n"
        for j, mes in enumerate(meses):
            # Verificar se o índice está dentro dos limites de valores_celulas
            if index_valores < len(valores_celulas):
                # Utilizando f-strings para formatar a data corretamente
                data = f"{ano}-{meses[j]:02d}-01 03:00:00.000"
                # Adicionando o valor correspondente ao mês e ano ao SQL
                sql += f"INSERT INTO VARIACAO_INDICE\n"
                sql += f"(ID, INDICE_ID, DATA, VALOR, CREATED_BY, CREATED_DATE, MODIFIED_BY, MODIFIED_DATE, VERSION)\n"
                sql += f"VALUES((SELECT MAX(ID) + 1 FROM VARIACAO_INDICE), 1, '{data}', {valores_celulas[index_valores]}, 'ATTUS', CURRENT_TIMESTAMP, 'ATTUS', CURRENT_TIMESTAMP, 0);\n\n"
                index_valores += 1  # Avançar para o próximo valor em valores_celulas
            else:
                print("Índices fora dos limites da lista!")

    # Salvando o SQL em um arquivo TXT
    with open('sql_gerado.txt', 'w') as f:
        f.write(sql)

    print('SQL gerado com sucesso!')

except Exception as e:
    # Salvando o erro no arquivo TXT
    with open('erro.txt', 'w') as f:
        f.write(f"Ocorreu um erro: {str(e)}")
