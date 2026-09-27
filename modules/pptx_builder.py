"""
pptx_builder.py - PPTX生成モジュール（トンマナ適用版）

AIが出力したJSON構成と、元のPPTXから抽出したトンマナ情報を使って
新しいPPTXファイルを生成する。
- 元PPTXのスライドマスター/レイアウトを再利用（背景・ロゴ自動継承）
- フォント名・サイズ・色をトンマナに合わせて適用
- グラフ・テーブル・箇条書きの生成
- ロゴの自動配置
"""

import io
import copy
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor


class PPTXBuilder:
    # グラフタイプのマッピング
    CHART_TYPE_MAP = {
        "bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
        "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
        "stacked_bar": XL_CHART_TYPE.COLUMN_STACKED,
        "line": XL_CHART_TYPE.LINE,
        "line_markers": XL_CHART_TYPE.LINE_MARKERS,
        "pie": XL_CHART_TYPE.PIE,
        "doughnut": XL_CHART_TYPE.DOUGHNUT,
        "area": XL_CHART_TYPE.AREA,
    }

    def __init__(self, base_pptx_file, tone_and_manner=None):
        """
        ベースとなるPPTXファイルをロードする
        :param base_pptx_file: BytesIO等のファイルオブジェクト
        :param tone_and_manner: PPTXAnalyzer.extract_all()の戻り値（トンマナ情報）
        """
        self.prs = Presentation(base_pptx_file)
        self.tone = tone_and_manner or {}

        # 既存スライドを削除（テンプレートのサンプルスライドを消す）
        # ※ マスターとレイアウトは残る
        self._remove_existing_slides()

        # トンマナからフォント情報を取得
        font_styles = self.tone.get("font_styles", {})
        title_fonts = font_styles.get("title", [])
        body_fonts = font_styles.get("body", [])

        self.title_font_name = title_fonts[0].get("font_name") if title_fonts else None
        self.title_font_size = title_fonts[0].get("font_size_pt") if title_fonts else None
        self.title_font_bold = title_fonts[0].get("bold", True) if title_fonts else True
        self.title_font_color = title_fonts[0].get("color_rgb") if title_fonts else None

        self.body_font_name = body_fonts[0].get("font_name") if body_fonts else None
        self.body_font_size = body_fonts[0].get("font_size_pt") if body_fonts else None
        self.body_font_bold = body_fonts[0].get("bold", False) if body_fonts else False
        self.body_font_color = body_fonts[0].get("color_rgb") if body_fonts else None

        # ロゴ情報
        self._logos = self.tone.get("logos", [])

    def _remove_existing_slides(self):
        """
        テンプレートの既存スライドを全て削除する
        （マスターとレイアウトは保持される）
        """
        while len(self.prs.slides) > 0:
            rId = self.prs.slides._sldIdLst[0].get(
                "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
            )
            if rId is None:
                # 別の方法で削除
                slide_id = self.prs.slides._sldIdLst[0]
                self.prs.slides._sldIdLst.remove(slide_id)
            else:
                self.prs.part.drop_rel(rId)
                slide_id = self.prs.slides._sldIdLst[0]
                self.prs.slides._sldIdLst.remove(slide_id)

    def build(self, slide_data_json):
        """
        JSONデータを元にスライドを生成・追加する
        :param slide_data_json: AI生成のスライド構成JSON
        """
        slides_data = slide_data_json.get("slides", [])

        for slide_info in slides_data:
            layout_index = slide_info.get("layout_index", 0)

            # 安全にレイアウトを取得
            if layout_index >= len(self.prs.slide_layouts):
                layout_index = 0

            layout = self.prs.slide_layouts[layout_index]
            slide = self.prs.slides.add_slide(layout)

            # タイトルの設定
            title_text = slide_info.get("title", "")
            if title_text and slide.shapes.title:
                self._set_title(slide.shapes.title, title_text)

            # サブタイトルの設定（タイトルスライド用）
            subtitle_text = slide_info.get("subtitle", "")
            if subtitle_text:
                self._set_subtitle(slide, subtitle_text)

            # 本文（箇条書き）の設定
            body_text = slide_info.get("body", [])
            if body_text:
                self._set_body(slide, body_text)

            # テーブルの追加
            table_info = slide_info.get("table")
            if table_info:
                self._add_table(slide, table_info)

            # グラフの追加
            chart_info = slide_info.get("chart")
            if chart_info:
                self._add_chart(slide, chart_info)

    def _set_title(self, title_shape, text):
        """
        タイトルにテキストを設定し、トンマナのフォントスタイルを適用する
        """
        title_shape.text = text
        for para in title_shape.text_frame.paragraphs:
            for run in para.runs:
                self._apply_font_style(run, is_title=True)

    def _set_subtitle(self, slide, text):
        """
        サブタイトルプレースホルダーにテキストを設定する
        """
        for shape in slide.placeholders:
            # サブタイトルプレースホルダー (idx=1 がサブタイトルの場合が多い)
            if shape.placeholder_format.idx == 1:
                shape.text = text
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        self._apply_font_style(run, is_title=False)
                return

    def _set_body(self, slide, body_items):
        """
        本文（箇条書き）をプレースホルダーに設定する
        タイトル以外の最初のテキストプレースホルダーを使用
        """
        body_shape = None
        for shape in slide.placeholders:
            # タイトル（idx=0）以外のテキスト枠を探す
            if shape.placeholder_format.idx != 0 and shape.has_text_frame:
                body_shape = shape
                break

        if not body_shape:
            return

        tf = body_shape.text_frame
        tf.clear()

        for i, point in enumerate(body_items):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = str(point)
            p.level = 0  # インデントレベル

            # フォントスタイル適用
            for run in p.runs:
                self._apply_font_style(run, is_title=False)

    def _apply_font_style(self, run, is_title=False):
        """
        テキストランにトンマナのフォントスタイルを適用する
        """
        font = run.font

        if is_title:
            if self.title_font_name:
                font.name = self.title_font_name
            if self.title_font_size:
                font.size = Pt(self.title_font_size)
            if self.title_font_bold is not None:
                font.bold = self.title_font_bold
            if self.title_font_color:
                try:
                    font.color.rgb = RGBColor.from_string(self.title_font_color)
                except Exception:
                    pass
        else:
            if self.body_font_name:
                font.name = self.body_font_name
            if self.body_font_size:
                font.size = Pt(self.body_font_size)
            if self.body_font_bold is not None:
                font.bold = self.body_font_bold
            if self.body_font_color:
                try:
                    font.color.rgb = RGBColor.from_string(self.body_font_color)
                except Exception:
                    pass

    def _add_table(self, slide, table_info):
        """
        テーブルをスライドに追加する
        table_info: {"headers": ["列1", "列2"], "rows": [["値1", "値2"], ...]}
        """
        try:
            headers = table_info.get("headers", [])
            rows = table_info.get("rows", [])
            if not headers or not rows:
                return

            num_rows = len(rows) + 1  # ヘッダー行を含む
            num_cols = len(headers)

            # テーブルの位置とサイズ
            left = Inches(1)
            top = Inches(2.5)
            width = Inches(8)
            height = Inches(0.5 * num_rows)

            table_shape = slide.shapes.add_table(
                num_rows, num_cols, left, top, width, height
            )
            table = table_shape.table

            # ヘッダー行の設定
            for col_idx, header in enumerate(headers):
                cell = table.cell(0, col_idx)
                cell.text = str(header)
                # ヘッダーのフォントスタイル
                for para in cell.text_frame.paragraphs:
                    para.alignment = PP_ALIGN.CENTER
                    for run in para.runs:
                        run.font.bold = True
                        if self.body_font_name:
                            run.font.name = self.body_font_name
                        if self.body_font_size:
                            run.font.size = Pt(self.body_font_size - 2)

            # データ行の設定
            for row_idx, row_data in enumerate(rows):
                for col_idx, value in enumerate(row_data):
                    if col_idx < num_cols:
                        cell = table.cell(row_idx + 1, col_idx)
                        cell.text = str(value)
                        for para in cell.text_frame.paragraphs:
                            for run in para.runs:
                                if self.body_font_name:
                                    run.font.name = self.body_font_name
                                if self.body_font_size:
                                    run.font.size = Pt(self.body_font_size - 2)

        except Exception as e:
            print(f"テーブル生成エラー: {e}")

    def _add_chart(self, slide, chart_info):
        """
        グラフをスライドに追加する
        chart_info: {
            "type": "bar"|"line"|"pie"|...,
            "categories": ["A", "B"],
            "series": [{"name": "売上", "values": [100, 200]}]
        }
        """
        try:
            chart_type_str = chart_info.get("type", "bar").lower()
            categories = chart_info.get("categories", [])
            series_list = chart_info.get("series", [])

            if not categories or not series_list:
                return

            chart_data = CategoryChartData()
            chart_data.categories = categories

            for series in series_list:
                values = series.get("values", [])
                # 数値でない値を0に変換
                clean_values = []
                for v in values:
                    try:
                        clean_values.append(float(v))
                    except (ValueError, TypeError):
                        clean_values.append(0)
                chart_data.add_series(
                    series.get("name", "Series"),
                    clean_values
                )

            # グラフタイプの取得
            chart_type = self.CHART_TYPE_MAP.get(
                chart_type_str, XL_CHART_TYPE.COLUMN_CLUSTERED
            )

            # グラフの位置とサイズ（スライドの中央下部に配置）
            x, y = Inches(1.5), Inches(2.8)
            cx, cy = Inches(7), Inches(4.2)

            chart_shape = slide.shapes.add_chart(
                chart_type, x, y, cx, cy, chart_data
            )

            # グラフのスタイル調整
            chart = chart_shape.chart
            chart.has_legend = len(series_list) > 1

            # テーマカラーをグラフに適用
            theme_colors = self.tone.get("theme_colors", [])
            if theme_colors:
                try:
                    for i, series_obj in enumerate(chart.series):
                        if i < len(theme_colors):
                            color = RGBColor.from_string(theme_colors[i])
                            series_obj.format.fill.solid()
                            series_obj.format.fill.fore_color.rgb = color
                except Exception:
                    pass  # グラフの色設定に失敗しても続行

        except Exception as e:
            print(f"グラフ生成エラー: {e}")

    def get_output_io(self):
        """
        生成したPPTXをBytesIOに保存して返す
        """
        output = io.BytesIO()
        self.prs.save(output)
        output.seek(0)
        return output
