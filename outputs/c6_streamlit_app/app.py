from __future__ import annotations

from datetime import date, datetime

import pandas as pd
import streamlit as st

from c6_processor import (
    DEFAULT_PHONE_BASE_DIR,
    FilterOptions,
    build_filtered_from_csv,
    dataframe_to_xlsx_bytes,
    missing_phone_cpfs_to_csv_bytes,
    parse_faltam_values,
)


APP_NAME = "Auto C6BANK - Datafast for Argus"

st.set_page_config(page_title=APP_NAME, layout="wide")

st.title(APP_NAME)

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

uploaded = st.file_uploader("Extração DTF CSV", type=["csv"])

if uploaded is None:
    st.info("Envie o CSV extraído do C6.")
else:
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

    tab_cpfs, tab_xlsx = st.tabs(["1. CPFs para Lemiti", "2. Retorno Lemiti e XLSX"])

    with tab_cpfs:
        try:
            transformed, filtered = build_filtered_from_csv(
                uploaded, options, DEFAULT_PHONE_BASE_DIR
            )
            missing_csv_bytes = missing_phone_cpfs_to_csv_bytes(filtered)
        except Exception as exc:
            st.error(str(exc))
        else:
            missing_count = len(filtered[filtered["TELEFONE BASE"].fillna("").astype(str).str.strip() == ""])
            col1, col2, col3 = st.columns(3)
            col1.metric("Linhas no CSV", f"{len(transformed):,}".replace(",", "."))
            col2.metric("Linhas filtradas", f"{len(filtered):,}".replace(",", "."))
            col3.metric("Sem telefone", f"{missing_count:,}".replace(",", "."))

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            st.download_button(
                "Baixar CPFs sem telefone",
                data=missing_csv_bytes,
                file_name=f"cpfs_sem_telefone_{timestamp}.csv",
                mime="text/csv",
                type="primary",
            )

    with tab_xlsx:
        lemit_uploaded = st.file_uploader(
            "Retorno Lemiti", type=["csv", "xlsx", "xls"], key="lemit_upload"
        )
        if lemit_uploaded is None:
            st.info("Suba o retorno do Lemiti para gerar o XLSX final.")
        else:
            try:
                transformed, filtered = build_filtered_from_csv(
                    uploaded, options, DEFAULT_PHONE_BASE_DIR, lemit_uploaded
                )
                xlsx_bytes = dataframe_to_xlsx_bytes(filtered)
                missing_csv_bytes = missing_phone_cpfs_to_csv_bytes(filtered)
            except Exception as exc:
                st.error(str(exc))
            else:
                missing_count = len(filtered[filtered["TELEFONE BASE"].fillna("").astype(str).str.strip() == ""])
                col1, col2, col3 = st.columns(3)
                col1.metric("Linhas no CSV", f"{len(transformed):,}".replace(",", "."))
                col2.metric("Linhas filtradas", f"{len(filtered):,}".replace(",", "."))
                col3.metric("Ainda sem telefone", f"{missing_count:,}".replace(",", "."))

                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                st.download_button(
                    "Baixar XLSX",
                    data=xlsx_bytes,
                    file_name=f"auto_c6bank_{timestamp}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                )
                st.download_button(
                    "Baixar CPFs ainda sem telefone",
                    data=missing_csv_bytes,
                    file_name=f"cpfs_ainda_sem_telefone_{timestamp}.csv",
                    mime="text/csv",
                )

                preview = filtered.head(200).copy()
                if not preview.empty:
                    preview["DATA DE NASC"] = pd.to_datetime(preview["DATA DE NASC"]).dt.strftime("%d/%m/%Y")
                st.dataframe(preview, use_container_width=True, hide_index=True)
