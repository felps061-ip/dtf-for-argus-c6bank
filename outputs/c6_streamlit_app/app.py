from __future__ import annotations

from datetime import date, datetime

import pandas as pd
import streamlit as st

from argus_api import DEFAULT_SKILL_REFERENCE, DEFAULT_TOKEN_SIGNATURE, upload_mailing_csv_bytes

from c6_processor import (
    DEFAULT_PHONE_BASE_DIRS,
    FilterOptions,
    argus_discadora_csv_bytes,
    build_argus_campaign_name,
    build_argus_output_filename,
    build_lemit_processed_stem,
    build_filtered_from_csv,
    build_sms_enviados_output_filename,
    build_sms_output_filename,
    dataframe_to_xlsx_bytes,
    missing_phone_cpfs_to_csv_bytes,
    parse_date_value,
    parse_faltam_values,
    sms_enviados_xlsx_bytes,
    sms_xlsx_bytes,
)


APP_NAME = "Auto C6BANK - Datafast for Argus"
FOOTER_TEXT = "v1.0 by: Equipe TI Grupo Império"

st.set_page_config(page_title=APP_NAME, layout="wide")

st.title(APP_NAME)
st.markdown(
    f"""
    <style>
    .block-container {{
        padding-bottom: 3.5rem;
    }}
    .app-footer {{
        position: fixed;
        left: 0;
        bottom: 0;
        width: 100%;
        padding: 0.55rem 1rem;
        text-align: center;
        font-size: 0.82rem;
        color: #5f6368;
        background: rgba(255, 255, 255, 0.94);
        border-top: 1px solid #e5e7eb;
        z-index: 999;
    }}
    </style>
    <div class="app-footer">{FOOTER_TEXT}</div>
    """,
    unsafe_allow_html=True,
)

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
    telefone_data_text = st.text_input("Período Telefone", value="01/01/2026")

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
        telefone_data_inicial=parse_date_value(telefone_data_text),
    )

    tab_cpfs, tab_xlsx, tab_sms = st.tabs(
        ["1. CPFs para Lemiti", "2. Retorno Lemiti, XLSX e Argus", "3. SMS"]
    )

    with tab_cpfs:
        try:
            transformed, filtered = build_filtered_from_csv(
                uploaded, options, DEFAULT_PHONE_BASE_DIRS
            )
            missing_csv_bytes = missing_phone_cpfs_to_csv_bytes(filtered)
            audit_xlsx_bytes = dataframe_to_xlsx_bytes(filtered)
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
            st.download_button(
                "Baixar planilha de conferencia",
                data=audit_xlsx_bytes,
                file_name=f"conferencia_c6_telefones_{timestamp}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
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
                    uploaded, options, DEFAULT_PHONE_BASE_DIRS, lemit_uploaded
                )
                xlsx_bytes = dataframe_to_xlsx_bytes(filtered)
                xlsx_filename = f"{build_lemit_processed_stem(uploaded, len(filtered))}.xlsx"
                suggested_campaign = build_argus_campaign_name(xlsx_filename)
                campanha = st.text_input("Campanha Argus", value=suggested_campaign)
                argus_csv = argus_discadora_csv_bytes(filtered, campanha)
                argus_filename = build_argus_output_filename(xlsx_filename)
                missing_csv_bytes = missing_phone_cpfs_to_csv_bytes(filtered)
            except Exception as exc:
                st.error(str(exc))
            else:
                missing_count = len(filtered[filtered["TELEFONE BASE"].fillna("").astype(str).str.strip() == ""])
                missing_ddb_count = len(filtered[filtered["BDD"].fillna("").astype(str).str.strip() == ""])
                col1, col2, col3, col4 = st.columns(4)
                col1.metric("Linhas no CSV", f"{len(transformed):,}".replace(",", "."))
                col2.metric("Linhas filtradas", f"{len(filtered):,}".replace(",", "."))
                col3.metric("Ainda sem telefone", f"{missing_count:,}".replace(",", "."))
                col4.metric("Sem BDD", f"{missing_ddb_count:,}".replace(",", "."))

                st.download_button(
                    "Baixar XLSX",
                    data=xlsx_bytes,
                    file_name=xlsx_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                )
                st.download_button(
                    "Baixar CSV Discadora Argus",
                    data=argus_csv,
                    file_name=argus_filename,
                    mime="text/csv",
                    type="primary",
                )
                with st.expander("Enviar este CSV para a Argus"):
                    st.caption(
                        "Este botão realiza um envio real para o endpoint configurado na Argus."
                    )
                    argus_token = st.text_input(
                        "Token-Signature",
                        value=DEFAULT_TOKEN_SIGNATURE,
                        type="password",
                        key="argus_token_signature",
                    )
                    if st.button("Enviar agora", type="primary", key="send_argus_csv"):
                        try:
                            result = upload_mailing_csv_bytes(
                                argus_csv,
                                argus_filename,
                                argus_token,
                                DEFAULT_SKILL_REFERENCE,
                            )
                        except Exception as exc:
                            st.error(str(exc))
                        else:
                            st.success(
                                "Mailing enviado. "
                                f"idArquivo: {result.id_arquivo} | "
                                f"codStatus: {result.cod_status} | "
                                f"descStatus: {result.desc_status}"
                            )
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
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

    with tab_sms:
        sms_lemit_uploaded = st.file_uploader(
            "Retorno Lemiti para SMS", type=["csv", "xlsx", "xls"], key="sms_lemit_upload"
        )
        if sms_lemit_uploaded is None:
            st.info("Suba o retorno do Lemiti para gerar o XLSX de SMS.")
        else:
            try:
                transformed, filtered = build_filtered_from_csv(
                    uploaded, options, DEFAULT_PHONE_BASE_DIRS, sms_lemit_uploaded
                )
                sms_bytes = sms_xlsx_bytes(filtered)
                enviados_bytes = sms_enviados_xlsx_bytes(filtered)
                sms_filename = build_sms_output_filename(uploaded, len(filtered))
                enviados_filename = build_sms_enviados_output_filename(uploaded, len(filtered))
            except Exception as exc:
                st.error(str(exc))
            else:
                st.metric("Linhas base para SMS", f"{len(filtered):,}".replace(",", "."))
                st.download_button(
                    "Baixar XLSX SMS",
                    data=sms_bytes,
                    file_name=sms_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                )
                st.download_button(
                    "Baixar XLSX Enviados",
                    data=enviados_bytes,
                    file_name=enviados_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
