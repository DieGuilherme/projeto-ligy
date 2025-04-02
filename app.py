##################################
#### PROJETO FATURAMENTO LIGY ####
##################################


import pandas as pd
import numpy as np
import os

# Caminho do arquivo
arquivo_xlsx = r"C:\Users\br0661399023\OneDrive - Enel Spa\1. Diego\1. E2E_B2C\7. Python\Workspace\Ligy\simulado.xlsx"

def carregar_planilhas():
    """Carregar as planilhas necessárias"""
    df_faturamento = pd.read_excel(arquivo_xlsx, sheet_name="Faturamento", header=1)
    df_temp = pd.read_excel(arquivo_xlsx, sheet_name="temp", header=0)
    df_aux = pd.read_excel(arquivo_xlsx, sheet_name="auxiliar", header=[0])
    return df_faturamento, df_temp, df_aux

def criar_identificador(df_temp):
    """Criar identificador único para cada cliente"""
    if "nome" in df_temp.columns and "ref_fat_cli" in df_temp.columns:
        df_temp["cliente_ref"] = df_temp["nome"].astype(str) + " | " + pd.to_datetime(df_temp["ref_fat_cli"]).dt.strftime('%Y_%m')
    else:
        print("Erro: Colunas 'nome' e/ou 'ref_fat_cli' não foram encontradas na aba 'temp'.")
        exit()

def calcular_limite_comp(df_temp, df_aux):
    """Calcular o limite de consumo"""
    if "consumo (kWh)" in df_temp.columns and "tipo_forn" in df_temp.columns:
        # Transpor e mapear os valores de custo
        df_aux_transposed = df_aux.set_index(df_aux.columns[0]).T
        df_aux_transposed.columns = ["custo_disp"]
        disponibilidade_dict = df_aux_transposed["custo_disp"].to_dict()

        df_temp["tipo_forn"] = df_temp["tipo_forn"].str.strip()  # Limpar a colunax'
        df_temp["custo_disp"] = df_temp["tipo_forn"].map(disponibilidade_dict)
        df_temp["limite_comp"] = df_temp["consumo (kWh)"] - df_temp["custo_disp"]
    else:
        print("Erro: Algumas das colunas necessárias para calcular 'limite_comp' não foram encontradas.")

def calcular_rateio(df_temp):
    """Calcular rateio"""
    df_temp["rateio"] = round((df_temp["geracao_usina (kWh)"] * df_temp["rateio_cliente (%)"]), 2)

def calcular_limite2(df_temp):
    """Definir o valor devido de energia a faturar"""
    df_temp["limite2"] = np.round(df_temp["limite_comp"] - df_temp["credito_acum (kWh)"], 2)
    return df_temp

def calcular_cred_a_acumular(df_temp):
    """Definir o valor de credito a acumular para os proximos meses"""
    df_temp["cred_a_acumular"] = np.round(np.where((df_temp["limite2"] - df_temp["rateio"])>0,
             df_temp["limite2"] - df_temp["rateio"],
              df_temp["rateio"]) - df_temp["limite2"], 2) ### aqui pode ter um problema de arredondamento ###

def calcular_cons_faturado(df_temp):
    """Definir o valor devido para faturamento"""
    df_temp["cons_faturado"] = (df_temp["credito_acum (kWh)"] + df_temp["limite2"])

def calcular_valores_adicionais(df_temp):
    """Calcular valores adicionais"""
    df_temp["val_add"] = round(df_temp["tx_ip ($$)"] + df_temp["cob_des_add"], 2)

def calcular_fatura_enel(df_temp):
    """Calcular dados para a Fatura Enel."""
    df_temp["credito"] = df_temp["credito_acum (kWh)"] * df_temp["credito_acum (kWh)"]
    df_temp["limite3"] = df_temp["limite2"] - df_temp["credito_acum (kWh)"]
    df_temp["cred_rateio"] = np.round(np.where(df_temp["rateio"] <= df_temp["limite2"],
             df_temp["rateio"],
             df_temp["limite2"]), 2)
    df_temp["consumo_final"] = df_temp["limite2"] - df_temp["cred_rateio"]
    df_temp["val_consumo"] = df_temp["consumo (kWh)"] * df_temp["tarifa_conv ($$)"]
    df_temp["val_credito"] = df_temp["credito_acum (kWh)"] * df_temp["tarifa_cred_acum ($$)"]
    df_temp["val_rateio"] = df_temp["cred_rateio"] * df_temp["tarifa_gd"]
    df_temp["val_cons_final"] = df_temp["val_consumo"] - df_temp["val_credito"] - df_temp["val_rateio"] + df_temp["val_add"]

def calcular_valor_ligy(df_temp):
    """Calcular o Valor_Ligy"""
    df_temp["Valor_Ligy"] = round(df_temp["cons_faturado"] * df_temp["tarifa_gd"] * 0.8, 2)
    
def calcular_dados_fatura_ligy(df_temp):
    """Calcular dados para a Fatura Ligy."""
    df_temp["sub_energia_s_gd"] = round(df_temp["consumo (kWh)"] * df_temp["tarifa_conv ($$)"],2)
    df_temp["fat_enel_s_gd"] = df_temp["sub_energia_s_gd"] + df_temp["val_add"]
    df_temp["sub_energia_gd"] = round(df_temp["cons_faturado"] * df_temp["tarifa_gd"],2)
    df_temp["dif_cons_injec"] = round(df_temp["sub_energia_s_gd"] - df_temp["sub_energia_gd"],2)
    df_temp["benef_gd"] = round(df_temp["fat_enel_s_gd"] - df_temp["val_cons_final"],2) 
    df_temp["benef_ligy"] = round((df_temp["benef_gd"] * 0.2),2)
    df_temp["s_ligy"] = df_temp["fat_enel_s_gd"]
    df_temp["c_ligy"] = round(df_temp["val_cons_final"] + df_temp["Valor_Ligy"],2) 
    df_temp["economia_real"] = round(df_temp["s_ligy"] - df_temp["c_ligy"],2)
    df_temp["economia_percebida"] = np.where(df_temp["s_ligy"] == 0,
                                            0, df_temp["economia_real"]/df_temp["s_ligy"])
    df_temp["carbono"] = round((df_temp["consumo (kWh)"]*0.09) - (df_temp["consumo (kWh)"]*0.0305),2)
    df_temp["fatura_ligy"] = round(df_temp["benef_gd"] - df_temp["benef_ligy"],2)
    df_temp["dif"] = df_temp["fatura_enel_real"] - df_temp["val_cons_final"] 
    df_temp["farol"] = np.where((df_temp["dif"].abs() > 0.4) | df_temp["val_cons_final"].isna(), "NOK", "OK")

def main():
    # Carregar dados
    df_faturamento, df_temp, df_aux = carregar_planilhas()

    # Passos do cálculo
    criar_identificador(df_temp)
    calcular_limite_comp(df_temp, df_aux)
    calcular_rateio(df_temp)
    df_temp = calcular_limite2(df_temp)
    calcular_cred_a_acumular(df_temp)
    calcular_cons_faturado(df_temp)
    calcular_valor_ligy(df_temp)
    calcular_valores_adicionais(df_temp)
    calcular_fatura_enel (df_temp)
    calcular_dados_fatura_ligy(df_temp)
    df_temp.to_excel("debug_completo.xlsx", index=False)

    # Exibir e salvar resultados
    print("\n🔹 **Tabela Check de Fatura**")
    print(df_temp[["cliente_ref", "Valor_Ligy", "consumo (kWh)", "tipo_forn", "custo_disp", "limite_comp", "limite2", "cred_a_acumular", "cons_faturado", "val_add", "val_cons_final"]].head().to_string(index=False))
    
    # Criar DataFrames para cada tabela
    df_check_fatura = df_temp[["cliente_ref", "Valor_Ligy", "consumo (kWh)", "tipo_forn", "custo_disp", "limite_comp", "cons_faturado", "val_add" ,"val_cons_final"]]
    df_fatura_ligy = df_temp[["cliente_ref", "sub_energia_s_gd", "fat_enel_s_gd", "sub_energia_gd", "dif_cons_injec","tx_ip ($$)", "cob_des_add" , "benef_gd", "benef_ligy", "s_ligy", "c_ligy", "economia_real", "economia_percebida", "fatura_ligy", "carbono", "val_cons_final", "fatura_enel_real", "farol"]]
    
    # Salvar em arquivo Excel com duas guias
    with pd.ExcelWriter('resultados_faturamento.xlsx') as writer:
        df_check_fatura.to_excel(writer, sheet_name='Check de Fatura', index=False)
        df_fatura_ligy.to_excel(writer, sheet_name='Fatura Ligy', index=False)

    print("\n🔹 **Tabela para montar Fatura Ligy**")
    print(df_temp[["cliente_ref", "sub_energia_s_gd", "fat_enel_s_gd", "sub_energia_gd", "dif_cons_injec","tx_ip ($$)", 
                             "cob_des_add", "val_cons_final", "fatura_enel_real", "benef_gd", "benef_ligy", "s_ligy", "c_ligy", 
                            "economia_real", "economia_percebida", "fatura_ligy", "carbono", "farol"]].head().to_string(index=False))
    
# Executar o script
if __name__ == "__main__":
    main()
