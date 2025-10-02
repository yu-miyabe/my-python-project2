import streamlit as st
import pandas as pd
import openpyxl
import io

st.title("スタンド・ガイド正規出図工数表作成ツール")

# プルダウンで選択（A〜L + 空白）
options = [chr(i) for i in range(ord('A'), ord('L') + 1)] + ["（空白）"]
selected_letter = st.selectbox(" 出荷区分を選択", options)

# 協力会社リスト（新しい内容）
company_options = [
    "PAP",
    "Y・Gテック",
    "ヒラテ技研（近江八幡メンバー）",
    "ヒラテ技研（請負メンバー）",
    "DSE",
    "ユニテツク",
    "タイガ設計",
    "中央エンジ"
]

# 設計会社の選択
stand_company = st.selectbox(" スタンド設計会社を選択", company_options)
guide_company = st.selectbox(" ガイド設計会社を選択", company_options)

# Excelファイルのアップロード
uploaded_file = st.file_uploader("📁 仕様一覧表(xlsm)をアップロードしてください", type=["xlsx", "xlsm"])

# 書き込み対象ファイルのパス（固定）
target_path = r"C:\Users\200804\home\スタンドガイド工数計画_v11.XLSX"

if uploaded_file:
    try:
        # 集計処理
        df = pd.read_excel(uploaded_file, sheet_name="仕様一覧表", engine="openpyxl")
        n_col = df.columns[13]
        ac_col = df.columns[28]

        if selected_letter == "（空白）":
            filtered_df = df[df[n_col].isna() | (df[n_col].astype(str).str.strip() == "")]
        else:
            filtered_df = df[df[n_col] == selected_letter]

        ac_cleaned = (
            filtered_df[ac_col]
            .astype(str)
            .str.replace(r"[^\d.]", "", regex=True)
            .replace("", pd.NA)
            .dropna()
            .astype(float)
        )

        total_ac = ac_cleaned.sum()
        total_distance = total_ac / 1000
        count = len(filtered_df)

        if count == 0:
            st.warning(f"⚠️ N列に '{selected_letter}' が含まれる行は見つかりませんでした。")
        else:
            label = "空白" if selected_letter == "（空白）" else selected_letter
            st.success(f"✅ 出荷区分が '{label}' （{count} 機番）の総機長は：{total_distance:.2f}m")
            st.dataframe(filtered_df)

            # Excelファイル読み込み＆書き込み
            wb = openpyxl.load_workbook(target_path)
            for sheet_name, company_value in zip(
                ["スタンド正規出図", "ガイド正規出図"],
                [stand_company, guide_company]
            ):
                ws = wb[sheet_name]
                for row in ws.iter_rows():
                    for cell in row:
                        if cell.value == "総距離(m)":
                            ws.cell(row=cell.row + 1, column=cell.column).value = total_distance
                        if cell.value == "計算式":
                            ws.cell(row=cell.row - 1, column=cell.column).value = company_value

            # メモリ上に保存してダウンロード
            output = io.BytesIO()
            wb.save(output)
            output.seek(0)

            st.download_button(
                label="📥 スタンド・ガイド工数表をダウンロード",
                data=output,
                file_name="スタンドガイド工数計画_更新済.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"❌ エラーが発生しました：{str(e)}")
else:
    st.info("👆 上のボックスにExcelファイルをドラッグ＆ドロップしてください")