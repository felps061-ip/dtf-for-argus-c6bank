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
    extract_uf_from_filename,
    format_bank_info,
    identify_bank_from_filename,
    missing_phone_cpfs_to_csv_bytes,
    parse_date_value,
    parse_faltam_values,
    read_c6_csv,
    sms_enviados_xlsx_bytes,
    sms_xlsx_bytes,
)


APP_NAME = "Auto C6BANK - Datafast for Argus"
FOOTER_TEXT = "v1.0 by: Equipe TI Grupo Império"

st.set_page_config(page_title=APP_NAME, layout="wide")

st.markdown(
    f"""
    <style>
    .block-container {{
        padding-top: 1rem;
        padding-bottom: 3.5rem;
        background: #F5F3EE;
    }}
    .hero {{
        background: transparent;
        color: #ffffff;
        padding: 0 0 0.55rem 0;
        margin-bottom: 0.35rem;
    }}
    .hero h1 {{
        margin: 0;
        color: #111111;
        font-size: 1.35rem;
        letter-spacing: 0;
    }}
    .hero p {{
        margin: 0.2rem 0 0;
        color: #555555;
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
    <div class="hero">
      <h1>Auto Datafast para Argus</h1>
      <p>Prepare bases C6 e Digio automaticamente.</p>
    </div>
    <div class="app-footer">{FOOTER_TEXT}</div>
    """,
    unsafe_allow_html=True,
)


def _format_count(value: int) -> str:
    return f"{int(value):,}".replace(",", ".")


def _shorten(name: str, limit: int = 42) -> str:
    return name if len(name) <= limit else f"{name[: limit - 3]}..."


def _friendly_error(error: object) -> str:
    text = str(error)
    if "Banco nao identificado" in text:
        return 'Arquivo incorreto. Selecione um arquivo que contenha "C6" ou "Digio" no nome.'
    if "Cabe" in text and "ausentes" in text:
        columns = text.split(":", 1)[-1].strip()
        return f"Arquivo com colunas faltando: {columns}. Confira se selecionou a exportacao DTF correta."
    return text


def _summary_card(title: str, value: str) -> None:
    st.markdown(
        f'<div class="step-card"><div class="step-title">{title}</div><div class="step-value">{value}</div></div>',
        unsafe_allow_html=True,
    )

with st.sidebar:
    st.subheader("2. Conferir filtros")
    prazo_text = st.text_input("Prazos FALTAM", value="", help="Pode ficar vazio ou receber varios prazos, como 67,72.")
    faltam_modo_label = st.selectbox(
        "Regra dos prazos",
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
    st.info('Envie o CSV extraido do C6 ou Digio.')
else:
    try:
        bank = identify_bank_from_filename(uploaded)
        uf = extract_uf_from_filename(uploaded.name)
        raw_rows = len(read_c6_csv(uploaded))
    except Exception as exc:
        st.error(_friendly_error(exc))
        st.stop()
    else:
        st.success(
            "✓ Base carregada  "
            f"Banco: {bank.display_name} | UF: {uf} | "
            f"Registros: {_format_count(raw_rows)} | "
            f"Coeficiente: {str(bank.coefficient).replace('.', ',')}"
        )

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
        ["Preparar para Lemiti", "Finalizar Base e Argus", "Gerar SMS"]
    )

    with tab_cpfs:
        try:
            transformed, filtered = build_filtered_from_csv(
                uploaded, options, DEFAULT_PHONE_BASE_DIRS
            )
            missing_csv_bytes = missing_phone_cpfs_to_csv_bytes(filtered)
            audit_xlsx_bytes = dataframe_to_xlsx_bytes(filtered)
        except Exception as exc:
            st.error(_friendly_error(exc))
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
                st.error(_friendly_error(exc))
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
                    st.write(f"Banco: {bank.display_name}")
                    st.write(f"UF: {uf}")
                    st.write(f"Campanha: {campanha}")
                    st.write(f"Arquivo: {argus_filename}")
                    argus_token = st.text_input(
                        "Token-Signature",
                        value=DEFAULT_TOKEN_SIGNATURE,
                        type="password",
                        key="argus_token_signature",
                    )
                    confirm_argus = st.checkbox(
                        "Conferi banco, UF, campanha e arquivo. Quero enviar para a Argus.",
                        key="confirm_argus_send",
                    )
                    if st.button(
                        "Enviar para Argus",
                        type="primary",
                        key="send_argus_csv",
                        disabled=not confirm_argus,
                    ):
                        try:
                            result = upload_mailing_csv_bytes(
                                argus_csv,
                                argus_filename,
                                argus_token,
                                DEFAULT_SKILL_REFERENCE,
                            )
                        except Exception as exc:
                            st.error(_friendly_error(exc))
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

                preview_columns = [
                    column
                    for column in ["CPF", "NOME", "BENEFÍCIO", "TELEFONE BASE", "BDD"]
                    if column in filtered.columns
                ]
                preview = filtered[preview_columns].head(20).copy()
                if not preview.empty:
                    preview["Banco"] = bank.display_name
                    preview["UF"] = uf
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
                st.error(_friendly_error(exc))
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
