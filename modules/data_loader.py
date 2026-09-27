"""
data_loader.py - データ読み込みモジュール

各種データファイル（CSV, Excel, TXT, PDF）を読み込み、
AI解析用のテキスト形式に変換する。
"""

import pandas as pd
import pdfplumber
import io


class DataLoader:
    @staticmethod
    def _df_to_text(df):
        """
        DataFrameを安全にテキスト化する（tabulateがなくても動作する）
        """
        try:
            return df.to_markdown(index=False)
        except Exception:
            # tabulateがない場合はto_stringを使用
            return df.to_string(index=False)

    @staticmethod
    def load_data(uploaded_file):
        """
        アップロードされたデータファイルを読み込み、テキストとして返す
        :param uploaded_file: アップロードされたファイルオブジェクト
        :return: (content: str, summary: str)
        """
        filename = getattr(uploaded_file, "name", "unknown")
        ext = filename.split(".")[-1].lower()

        content = ""
        summary = ""

        try:
            if ext in ["csv"]:
                # エンコーディングの自動判別を試みる
                try:
                    df = pd.read_csv(uploaded_file, encoding="utf-8")
                except UnicodeDecodeError:
                    if hasattr(uploaded_file, "seek"):
                        uploaded_file.seek(0)
                    df = pd.read_csv(uploaded_file, encoding="cp932")

                content = DataLoader._df_to_text(df)
                summary = f"CSVデータ: {df.shape[0]}行 × {df.shape[1]}列\n" + str(df.head())

            elif ext in ["xlsx", "xls"]:
                df = pd.read_excel(uploaded_file)
                content = DataLoader._df_to_text(df)
                summary = f"Excelデータ: {df.shape[0]}行 × {df.shape[1]}列\n" + str(df.head())

            elif ext in ["txt"]:
                raw_bytes = uploaded_file.getvalue() if hasattr(uploaded_file, "getvalue") else uploaded_file.read()
                try:
                    content = raw_bytes.decode("utf-8")
                except UnicodeDecodeError:
                    content = raw_bytes.decode("cp932", errors="replace")
                summary = f"テキストデータ (先頭100文字): {content[:100]}..."

            elif ext in ["pdf"]:
                with pdfplumber.open(uploaded_file) as pdf:
                    pages = [page.extract_text() for page in pdf.pages if page.extract_text()]
                content = "\n".join(pages)
                summary = f"PDFデータ: 全{len(pages)}ページ (抽出完了)"

            else:
                return "未対応のファイル形式です。", "エラー: 未対応の形式です"

            # 長すぎるデータへの対処 (AIプロンプトのトークン制限防止)
            if len(content) > 15000:
                content = content[:15000] + "\n... (データ省略: 長すぎるため一部を抜粋)"

            return content, summary

        except Exception as e:
            return str(e), f"読み込みエラー: {str(e)}"
