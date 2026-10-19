import io
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Key Financial Indicators (KFI)", page_icon="📊", layout="wide"
)

st.title("📊 Key Financial Indicators (KFI) — Заповнення 4-го аркуша")
st.caption(
    "Завантажте фінансовий звіт компанії (Форма № 1 та Форма № 2), щоб сформувати показники KFI"
)

# Панель вводу курсу та завантаження файлу
col_file, col_rate = st.columns([3, 1])
with col_file:
    uploaded_file = st.file_uploader(
        "Оберіть вхідний файл фінансового звіту (.xlsx)", type=["xlsx"]
    )
with col_rate:
    eur_rate = st.number_input(
        "Exchange rate UAH/EUR (input)",
        min_value=1.0,
        value=48.44,
        step=0.01,
        format="%.2f",
    )


def parse_and_calculate_kfi(file_bytes: bytes, eur_rate: float):
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    sheets = xl.sheet_names

    f1_name = next(
        (s for s in sheets if "Форма № 1" in s or "Баланс" in s), None
    )
    f2_name = next((s for s in sheets if "Форма № 2" in s or "Звіт про" in s), None)

    if not f1_name or not f2_name:
        raise ValueError(
            "У файлі обов'язково мають бути аркуші 'Форма № 1' та 'Форма № 2'."
        )

    f1_raw = pd.read_excel(xl, sheet_name=f1_name)
    f2_raw = pd.read_excel(xl, sheet_name=f2_name)

    def extract_codes(df):
        matches = df[
            df.apply(lambda r: r.astype(str).str.contains("Код").any(), axis=1)
        ]
        if matches.empty:
            raise ValueError("Не знайдено колонку 'Код' у таблиці звіту.")
        h_idx = matches.index[0]
        cleaned = df.iloc[h_idx + 1 :].copy()
        cleaned.columns = df.iloc[h_idx]
        cleaned["Код"] = pd.to_numeric(cleaned["Код"], errors="coerce")
        cleaned = cleaned.dropna(subset=["Код"])
        cleaned["Код"] = cleaned["Код"].astype(int)

        cols = cleaned.columns.tolist()
        curr_col = next(
            (
                c
                for c in cols
                if any(k in str(c).lower() for k in ["кінець", "поточний", "звітний"])
            ),
            cols[-1],
        )
        prev_col = next(
            (
                c
                for c in cols
                if any(k in str(c).lower() for k in ["початок", "попередній"])
            ),
            cols[-2],
        )

        data = {}
        for _, row in cleaned.iterrows():
            code = row["Код"]
            prev_raw = pd.to_numeric(row[prev_col], errors="coerce")
            curr_raw = pd.to_numeric(row[curr_col], errors="coerce")
            # Запобігання проникненню NaN: порожні комірки суворо стають 0.0
            data[code] = {
                "prev": 0.0 if pd.isna(prev_raw) else float(prev_raw),
                "curr": 0.0 if pd.isna(curr_raw) else float(curr_raw),
            }
        return data

    f1 = extract_codes(f1_raw)
    f2 = extract_codes(f2_raw)

    def g1(code, p="curr"):
        val = f1.get(code, {}).get(p, 0.0)
        return 0.0 if (val is None or pd.isna(val)) else float(val)

    def g2(code, p="curr"):
        val = f2.get(code, {}).get(p, 0.0)
        return 0.0 if (val is None or pd.isna(val)) else float(val)

    # 1. Верхній блок EUR
    fixed_assets_eur = (
        (g1(1010, "curr") / eur_rate) / 1000.0 if eur_rate else 0.0
    )
    tot_assets_curr = g1(1300, "curr")
    fixed_assets_pct = (
        (g1(1010, "curr") / tot_assets_curr) if tot_assets_curr != 0 else 0.0
    )

    turnover_prev = g2(2000, "prev")
    turnover_curr = g2(2000, "curr")
    turnover_growth = (
        ((turnover_curr / turnover_prev) - 1.0) if turnover_prev != 0 else 0.0
    )

    inv_real_estate = (
        ((g1(1016, "curr") - g1(1017, "curr")) / eur_rate) / 1000.0
        if eur_rate
        else 0.0
    )

    # 2. P&L BRIDGE (Mio. UAH)
    turnover = turnover_curr / 1000.0
    profit_after_tax = (g2(2350, "curr") - g2(2355, "curr")) / 1000.0
    tax_charge = -g2(2300, "curr") / 1000.0
    profit_before_tax = profit_after_tax + tax_charge

    extra_writeoffs = -g2(2240, "curr") / 1000.0
    extra_other = g2(2270, "curr") / 1000.0
    extra_items = extra_writeoffs + extra_other

    bank_int_income = -g2(2220, "curr") / 1000.0
    bank_int_expense = g2(2250, "curr") / 1000.0
    net_bank_interest = bank_int_income + bank_int_expense

    ebit = profit_before_tax + extra_items + net_bank_interest
    ebit_pct = (
        (ebit / turnover) if (turnover and not pd.isna(turnover)) else 0.0
    )

    depreciation = g2(2515, "curr") / 1000.0
    ebitda = ebit + depreciation
    ebitda_pct = (
        (ebitda / turnover) if (turnover and not pd.isna(turnover)) else 0.0
    )

    ebda = ebitda - tax_charge - net_bank_interest
    ebda_pct = (
        (ebda / turnover) if (turnover and not pd.isna(turnover)) else 0.0
    )

    new_inv_tangible = (
        g1(1011, "curr") - g1(1011, "prev")
    ) / 1000.0 + max(0.0, (g1(1005, "curr") - g1(1005, "prev")) / 1000.0)

    # 3. Gearing та активи
    equity = g1(1495, "curr") / 1000.0
    bank_loans = (
        (
            (g1(1510, "curr") + g1(1600, "curr"))
            if (g1(1510, "curr") + g1(1600, "curr")) != 0
            else g1(1595, "curr")
        )
        / 1000.0
    )
    other_loans = (g1(1515, "curr") + g1(1610, "curr")) / 1000.0
    interest_debt = bank_loans + other_loans
    cash_in_bank = g1(1165, "curr") / 1000.0
    total_assets = tot_assets_curr / 1000.0
    equity_to_assets = (equity / total_assets) if total_assets != 0 else 0.0

    trade_creditors = g1(1615, "curr") / 1000.0
    trade_debtors = g1(1125, "curr") / 1000.0
    inventory = g1(1100, "curr") / 1000.0

    current_assets = g1(1195, "curr") / 1000.0
    current_liab = g1(1695, "curr") / 1000.0
    payroll_cost = (g2(2505, "curr") + g2(2510, "curr")) / 1000.0

    # Функція форматування: виключає появу NaN та виводить прочерки
    def fmt_val(v, is_pct=False, force_dash=False):
        if pd.isna(v) or v is None:
            return "-"
        if force_dash and abs(v) < 1e-9:
            return "-"
        if is_pct:
            return f"{v * 100:.1f}%"
        return f"{v:.2f}"

    df_eur = pd.DataFrame(
        [
            {
                "Параметр": "Exchange rate UAH/EUR (input)",
                "Значення": f"{eur_rate:.2f}",
            },
            {
                "Параметр": "Fixed assets, MIO EUR",
                "Значення": fmt_val(fixed_assets_eur),
            },
            {
                "Параметр": "Fixed assets % in total ASSETS",
                "Значення": fmt_val(fixed_assets_pct, is_pct=True),
            },
            {
                "Параметр": "Turnover growth YoY",
                "Значення": fmt_val(turnover_growth, is_pct=True),
            },
            {
                "Параметр": "Investment real estate, MIO EUR",
                "Значення": fmt_val(inv_real_estate, force_dash=True),
            },
        ]
    )

    df_bridge = pd.DataFrame(
        [
            {"Стаття KFI": "TURNOVER", "Значення": fmt_val(turnover)},
            {
                "Стаття KFI": "Profit (+)/Loss (-) AFTER tax",
                "Значення": fmt_val(profit_after_tax),
            },
            {
                "Стаття KFI": "Add Back Tax Charge (+)/Credit (-)",
                "Значення": fmt_val(tax_charge, force_dash=True),
            },
            {
                "Стаття KFI": "Profit (+)/Loss (-) before tax",
                "Значення": fmt_val(profit_before_tax),
            },
            {
                "Стаття KFI": "Extraordinary items",
                "Значення": fmt_val(extra_items, force_dash=True),
            },
            {"Стаття KFI": "Currency translation difference", "Значення": "-"},
            {
                "Стаття KFI": "Net interest & similar charges - Bank interest",
                "Значення": fmt_val(net_bank_interest, force_dash=True),
            },
            {"Стаття KFI": "EBIT", "Значення": fmt_val(ebit)},
            {
                "Стаття KFI": "  - As percentage of turnover",
                "Значення": fmt_val(ebit_pct, is_pct=True),
            },
            {
                "Стаття KFI": "Depreciation and Amortisation",
                "Значення": fmt_val(depreciation),
            },
            {
                "Стаття KFI": "Depreciation, Amortisation & Grants",
                "Значення": fmt_val(depreciation),
            },
            {"Стаття KFI": "EBITDA", "Значення": fmt_val(ebitda)},
            {
                "Стаття KFI": "  - As percentage of turnover",
                "Значення": fmt_val(ebitda_pct, is_pct=True),
            },
            {
                "Стаття KFI": "Less Tax Charge (+)/Credit (-)",
                "Значення": fmt_val(tax_charge, force_dash=True),
            },
            {"Стаття KFI": "EBDA", "Значення": fmt_val(ebda)},
            {
                "Стаття KFI": "  - As percentage of turnover",
                "Значення": fmt_val(ebda_pct, is_pct=True),
            },
            {
                "Стаття KFI": "New investments - tangible",
                "Значення": fmt_val(new_inv_tangible),
            },
            {
                "Стаття KFI": "New investments",
                "Значення": fmt_val(new_inv_tangible),
            },
            {"Стаття KFI": "Equity", "Значення": fmt_val(equity)},
            {
                "Стаття KFI": "Equity (total shareholders funds)",
                "Значення": fmt_val(equity),
            },
            {
                "Стаття KFI": "Bank loans, borrowings, mortgages",
                "Значення": fmt_val(bank_loans, force_dash=True),
            },
            {
                "Стаття KFI": "Interest bearing debt",
                "Значення": fmt_val(interest_debt, force_dash=True),
            },
            {
                "Стаття KFI": "cash in bank (+)",
                "Значення": fmt_val(cash_in_bank),
            },
            {"Стаття KFI": "Total Assets", "Значення": fmt_val(total_assets)},
            {
                "Стаття KFI": "Equity / Total assets",
                "Значення": fmt_val(equity_to_assets, is_pct=True),
            },
            {
                "Стаття KFI": "Trade Creditors",
                "Значення": fmt_val(trade_creditors),
            },
            {"Стаття KFI": "Trade Debtors", "Значення": fmt_val(trade_debtors)},
            {"Стаття KFI": "Inventory", "Значення": fmt_val(inventory)},
            {
                "Стаття KFI": "Current Assets (from financial statements)",
                "Значення": fmt_val(current_assets),
            },
            {
                "Стаття KFI": "Current Assets",
                "Значення": fmt_val(current_assets),
            },
            {
                "Стаття KFI": "Current Liabilities (from financial statements)",
                "Значення": fmt_val(current_liab),
            },
            {
                "Стаття KFI": "Current Liabilities",
                "Значення": fmt_val(current_liab),
            },
            {"Стаття KFI": "Payroll cost", "Значення": fmt_val(payroll_cost)},
        ]
    )

    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
    if "KFI" in wb.sheetnames:
        ws = wb["KFI"]
        ws["H1"] = eur_rate
    output = io.BytesIO()
    wb.save(output)

    return df_eur, df_bridge, output.getvalue()


if uploaded_file is not None:
    try:
        bytes_data = uploaded_file.read()
        with st.spinner("Обробка фінансової звітності..."):
            df_eur, df_bridge, file_ready = parse_and_calculate_kfi(
                bytes_data, eur_rate
            )

        st.success("4-й аркуш KFI розраховано успішно!")

        col_left, col_right = st.columns([3, 2])
        with col_left:
            st.subheader("P&L BRIDGE та статті KFI (Mio. UAH / %)")
            st.dataframe(df_bridge, use_container_width=True, hide_index=True)

        with col_right:
            st.subheader("Показники в EUR")
            st.dataframe(df_eur, use_container_width=True, hide_index=True)

            st.write("---")
            st.download_button(
                label="📥 Завантажити оновлений Excel (із заповненим KFI)",
                data=file_ready,
                file_name=uploaded_file.name.replace(".xlsx", "_KFI_filled.xlsx"),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    except Exception as e:
        st.error(f"Помилка розрахунку: {str(e)}")
