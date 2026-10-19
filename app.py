import pandas as pd
import streamlit as st
from model import (
    calculate_kfi_data,
    parse_financial_statements,
    update_excel_kfi_sheet,
)

st.set_page_config(
    page_title="KFI Sheet Generator", page_icon="📑", layout="wide"
)

# Стилі для відтворення таблиці 4-го аркуша KFI
st.markdown(
    """
    <style>
    .kfi-container {
        display: flex;
        gap: 30px;
        margin-top: 15px;
    }
    .kfi-main-table {
        width: 60%;
        border-collapse: collapse;
        font-family: inherit;
        font-size: 14px;
    }
    .kfi-side-table {
        width: 35%;
        border-collapse: collapse;
        font-family: inherit;
        font-size: 14px;
        height: fit-content;
    }
    .kfi-main-table td, .kfi-side-table td {
        padding: 5px 12px;
        border: 1px solid #3e4249;
    }
    .kfi-val {
        text-align: right;
        font-family: monospace;
        font-size: 14px;
    }
    .row_bold { font-weight: bold; }
    .row_gray { background-color: rgba(255, 255, 255, 0.08); }
    .row_gray_bold { background-color: rgba(255, 255, 255, 0.12); font-weight: bold; }
    .kfi-header {
        background-color: #2b313a;
        color: #00d26a;
        font-weight: bold;
        text-align: center;
        padding: 8px;
    }
    </style>
""",
    unsafe_allow_html=True,
)

st.title("📑 Key Financial Indicators (KFI) — Заповнення 4-го аркуша")
st.caption(
    "Завантажте звіт `raport_finansowy...xlsx` із 4 аркушами для автоматичного розрахунку KFI"
)

col1, col2 = st.columns([3, 1])
with col1:
    uploaded_file = st.file_uploader(
        "Оберіть Excel-файл звітності", type=["xlsx"]
    )
with col2:
    eur_rate = st.number_input(
        "Exchange rate UAH/EUR (input)",
        min_value=1.0,
        value=48.44,
        step=0.01,
        format="%.2f",
    )

if uploaded_file is not None:
    file_bytes = uploaded_file.read()

    try:
        with st.spinner("Аналіз звітності та розрахунок показників KFI..."):
            f1_data, f2_data = parse_financial_statements(file_bytes)
            kfi_res = calculate_kfi_data(f1_data, f2_data, eur_rate=eur_rate)

        st.success("4-й аркуш KFI успішно сформовано!")

        # Відображення структури KFI
        st.subheader("Key Financial Indicators (KFI)")

        top_box_html = ['<table class="kfi-side-table">']
        top_box_html.append(
            '<tr><th colspan="2" class="kfi-header">Параметри та індикатори EUR</th></tr>'
        )
        for label, val in kfi_res["top_box"]:
            top_box_html.append(
                f'<tr><td>{label}</td><td class="kfi-val"><b>{val}</b></td></tr>'
            )
        top_box_html.append("</table>")

        main_table_html = ['<table class="kfi-main-table">']
        main_table_html.append(
            '<tr><th class="kfi-header">P&L BRIDGE / Стаття</th><th class="kfi-header">Значення</th></tr>'
        )
        for label, val, cls in kfi_res["bridge_table"]:
            main_table_html.append(
                f'<tr class="{cls}"><td>{label}</td><td class="kfi-val">{val}</td></tr>'
            )
        main_table_html.append("</table>")

        st.markdown(
            f'<div class="kfi-container">{"".join(main_table_html)}{"".join(top_box_html)}</div>',
            unsafe_allow_html=True,
        )

        st.write("")
        # Кнопка збереження оригінального файлу із заповненим KFI
        updated_excel_bytes = update_excel_kfi_sheet(file_bytes, kfi_res)

        st.download_button(
            label="📥 Завантажити оновлений Excel (із заповненим 4-м аркушем KFI)",
            data=updated_excel_bytes,
            file_name=uploaded_file.name.replace(".xlsx", "_KFI_filled.xlsx"),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        st.error(f"Помилка при обробці файлу: {str(e)}")