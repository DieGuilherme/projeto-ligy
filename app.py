import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(page_title="Faturamento Ligy", layout="wide")
st.title("📊 Projeto Faturamento Ligy")

st.markdown("Envie o arquivo Excel com as abas `Faturamento`, `temp` e `auxiliar` para gerar os cálculos.")

uploaded_file = st.file_uploader("Selecione o arquivo .xlsx", type=["xlsx"])

if uploaded_file:
    # Carregar planilhas
    df_faturamento = pd.read_excel(uploaded_file, sheet_name="Faturamento", header=1)
    df_temp = pd.read_excel(uploaded_file, sheet_name="temp", header=0)
    df_aux = pd.read_excel(uploaded_file, sheet_name="auxiliar", header=0)

    # Etapas do processamento
    df_temp["cliente_ref"] = df_temp["nome"].astype(str) + " | " + pd.to_datetime(df_temp["ref_fat_cli"]).dt.strftime('%Y_%m')

    df_aux_transposed = df_aux.set_index(df_aux.columns[0]).T
    df_aux_transposed.columns = ["custo_disp"]
    disponibilidade_dict = df_aux_transposed["custo_disp"].to_dict()
    df_temp["tipo_forn"] = df_temp["tipo_forn"].str.strip()
    df_temp["custo_disp"] = df_temp["tipo_forn"].map(disponibilidade_dict)
    df_temp["limite_comp"] = df_temp["consumo (kWh)"] - df_temp["custo_disp"]

    df_temp["rateio"] = round(df_temp["geracao_usina (kWh)"] * df_temp["rateio_cliente (%)"], 2)
    df_temp["limite2"] = np.round(df_temp["limite_comp"] - df_temp["credito_acum (kWh)"], 2)

    df_temp["cred_a_acumular"] = np.round(np.where((df_temp["limite2"] - df_temp["rateio"])>0,
             df_temp["limite2"] - df_temp["rateio"],
              df_temp["rateio"]) - df_temp["limite2"], 2)

    df_temp["cons_faturado"] = df_temp["credito_acum (kWh)"] + df_temp["limite2"]
    df_temp["val_add"] = round(df_temp["tx_ip ($$)"] + df_temp["cob_des_add"], 2)
    df_temp["Valor_Ligy"] = round(df_temp["cons_faturado"] * df_temp["tarifa_gd"] * 0.8, 2)

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

    st.success("✅ Processamento concluído!")

    st.subheader("🔹 Tabela Check de Fatura")
    st.dataframe(df_temp[["cliente_ref", "Valor_Ligy", "consumo (kWh)", "tipo_forn", "custo_disp", "limite_comp", "limite2", "cred_a_acumular", "cons_faturado", "val_add", "val_cons_final"]])

    st.subheader("🔹 Tabela Fatura Ligy")
    st.dataframe(df_temp[["cliente_ref", "sub_energia_s_gd", "fat_enel_s_gd", "sub_energia_gd", "dif_cons_injec", "tx_ip ($$)", "cob_des_add", "val_cons_final", "fatura_enel_real", "benef_gd", "benef_ligy", "s_ligy", "c_ligy", "economia_real", "economia_percebida", "fatura_ligy", "carbono", "farol"]])

    # Download dos resultados
check_fatura_cols = [
    "cliente_ref", "Valor_Ligy", "consumo (kWh)", "tipo_forn",
    "custo_disp", "limite_comp", "cons_faturado", "val_add", "val_cons_final"
]

fatura_ligy_cols = [
    "cliente_ref", "sub_energia_s_gd", "fat_enel_s_gd", "sub_energia_gd", "dif_cons_injec",
    "tx_ip ($$)", "cob_des_add", "benef_gd", "benef_ligy", "s_ligy", "c_ligy",
    "economia_real", "economia_percebida", "fatura_ligy", "carbono",
    "val_cons_final", "fatura_enel_real", "farol"
]

with pd.ExcelWriter("resultados_faturamento_streamlit.xlsx") as writer:
    df_temp[check_fatura_cols].to_excel(writer, sheet_name='Check de Fatura', index=False)
    df_temp[fatura_ligy_cols].to_excel(writer, sheet_name='Fatura Ligy', index=False)
    
with open("resultados_faturamento_streamlit.xlsx", "rb") as f:
        st.download_button("📥 Baixar resultados em Excel", f, file_name="resultado_faturamento_ligy.xlsx")
