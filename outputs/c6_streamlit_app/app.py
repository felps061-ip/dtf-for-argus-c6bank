from __future__ import annotations

from datetime import date, datetime

import pandas as pd
import streamlit as st

from c6_processor import (
    FilterOptions,
    build_filtered_from_csv,
    dataframe_to_xlsx_bytes,
    parse_faltam_values,
)


st.set_page_config(page_title="C6 para XLSX", layout="wide")

st.title("C6 para XLSX")

with st.sidebar:
    prazo_text = st.text_input("FALTAM", value="")
    faltam_modo_label = st.selectbox(
        "Regra FALTAM",
        ["Maior que o prazo", "Até o prazo", "Exato"],
        index=0,
    )
    taxa_min = st.number_input("TAXA mínima", min_value=0.0, value=0.0, step=0.01, format="%.2f")
    troco_min = st.number_input("TROCO maior que", value=0.0, step=100.0, format="%.2f")
    data_inicial = st.date_input("DATA DE NASC inicial", value=date(1954, 1, 1), format="DD/MM/YYYY")
    data_final = st.date_input("DATA DE NASC final", value=date(1995, 12, 31), format="DD/MM/YYYY")

uploaded = st.file_uploader("CSV C6", type=["csv"])

if uploaded is None:
    st.info("Envie o CSV extraído do C6.")
else:
    try:
        options = FilterOptions(
            faltam=parse_faltam_values(prazo_text),
            faltam_modo={
                "Maior que o prazo": "maior",
                "Até o prazo": "ate",
                "Exato": "exato",
            }[faltam_modo_label],
            taxa_min=taxa_min,
            troco_min=troco_min,
            data_inicial=data_inicial,
            data_final=data_final,
        )
        transformed, filtered = build_filtered_from_csv(uploaded, options)
        xlsx_bytes = dataframe_to_xlsx_bytes(filtered)
    except Exception as exc:
        st.error(str(exc))
    else:
        col1, col2, col3 = st.columns(3)
        col1.metric("Linhas no CSV", f"{len(transformed):,}".replace(",", "."))
        col2.metric("Linhas filtradas", f"{len(filtered):,}".replace(",", "."))
        col3.metric("Removidas", f"{len(transformed) - len(filtered):,}".replace(",", "."))

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        st.download_button(
            "Baixar XLSX",
            data=xlsx_bytes,
            file_name=f"c6_filtrado_{timestamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

        preview = filtered.head(200).copy()
        if not preview.empty:
            preview["DATA DE NASC"] = pd.to_datetime(preview["DATA DE NASC"]).dt.strftime("%d/%m/%Y")
        st.dataframe(preview, use_container_width=True, hide_index=True)
