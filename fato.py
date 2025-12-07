import pandas as pd
import os

file_path_origem = "CONTROLE MENSAL 2024.xlsx"
file_path_destino = "TRATAMENTO DE DADOS.xlsx"

def limpar_valor(valor):
    """Converte valores para float, transformando erros ou vazios em 0.0"""
    try:
        if pd.isna(valor) or str(valor).strip() == '':
            return 0.0
        return float(valor)
    except:
        return 0.0

try:
    print(f"Lendo dados de: {file_path_origem}...")
    

    df_macro = pd.read_excel(file_path_origem, sheet_name='LEITURA VOL. MACRO_m³', header=None)
    df_micro = pd.read_excel(file_path_origem, sheet_name='LEITURA VOL. MICRO_m³', header=None)


    idx_macro = 10   # Linha 11 do Excel
    col_nome_sistema = 1 # Coluna B
    col_inicio_macro = 10 # Coluna K

    idx_micro = 10  # Linha 11 do Excel 
    col_inicio_micro = 4 # Coluna E

    lista_meses = [
        "jan 2024", "fev 2024", "mar 2024", "abr 2024", "mai 2024", "jun 2024",
        "jul 2024", "ago 2024", "set 2024", "out 2024", "nov 2024", "dez 2024"
    ]

    dados_coletados = []

    print("Processando linhas e realizando cálculos...")

    while True:
        # Condição de parada (fim das linhas)
        if idx_macro >= len(df_macro) or idx_micro >= len(df_micro):
            break

        # Nome do Sistema
        sistema = df_macro.iloc[idx_micro, col_nome_sistema]

        # Condição de parada (celula vazia ou TOTAL)
        if pd.isna(sistema) or str(sistema).strip().upper().startswith('TOTAL'):
            break

        # Loop dos 12 meses
        for i, mes_nome in enumerate(lista_meses):

            indice_coluna_macro = col_inicio_macro + (i * 2)
            
            # Para o micro, mantém o padrão (coluna a coluna, ou ajuste se precisar)
            indice_coluna_micro = col_inicio_micro + i
            
            # Valores Macro e Micro
            val_macro = limpar_valor(df_macro.iloc[idx_macro, indice_coluna_macro])
            val_micro = limpar_valor(df_micro.iloc[idx_micro, indice_coluna_micro])

            # Cálculos
            vol_perdido = val_macro - val_micro
            
            # IDF (Evita divisão por zero)
            if val_macro > 0:
                idf = vol_perdido / val_macro
            else:
                idf = 0.0

            dados_coletados.append({
                'SISTEMAS': sistema,
                'VOL. MACROMED': val_macro,
                'VOL. MICROMED': val_micro,
                'VOL. PERDIDO': vol_perdido,
                'DATA': mes_nome,
                'IDF': idf
            })

        # Próximo sistema
        idx_macro += 1
        idx_micro += 1

    # Cria o DataFrame final
    df_final = pd.DataFrame(dados_coletados)
    
    print("\n--- Processamento Concluído ---")
    print(f"Total de registros gerados: {len(df_final)}")
    print(df_final.head(3).to_string(index=False)) # Mostra prévia

    
    if os.path.exists(file_path_destino):
        # Se o arquivo já existe, usamos o modo 'a' (append) com replace na aba específica
        # Engine openpyxl é necessária para manipular arquivos existentes sem corrompê-los
        with pd.ExcelWriter(file_path_destino, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
            df_final.to_excel(writer, sheet_name='FATO', index=False)
            print("Aba 'FATO' atualizada com sucesso no arquivo existente.")
    else:
        # Se o arquivo não existe, criamos um novo
        df_final.to_excel(file_path_destino, sheet_name='FATO', index=False)
        print("Arquivo criado e aba 'FATO' salva com sucesso.")

except FileNotFoundError as e:
    print(f"ERRO DE ARQUIVO: {e}")
    print("Verifique se o nome do arquivo de origem está correto e na mesma pasta.")
except PermissionError:
    print(f"ERRO DE PERMISSÃO: Não foi possível salvar em '{file_path_destino}'.")
    print("Feche o arquivo Excel se ele estiver aberto e tente novamente.")
except Exception as e:
    print(f"Ocorreu um erro inesperado: {e}")
