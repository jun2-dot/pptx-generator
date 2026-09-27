"""
pptx_analyzer.py - PPTXトンマナ抽出モジュール

ベースとなるPPTXファイルから以下のデザイン要素を抽出する:
- スライドレイアウト名とプレースホルダー配置
- カラーテーマ（RGBカラーパレット）
- フォント名・フォントサイズ（タイトル用・本文用）
- 背景色・背景画像
- ロゴ画像（固定位置の画像を検出）
- テキストスタイル（太字・色・配置）
"""

import io
import os
from pptx import Presentation
from pptx.util import Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE


class PPTXAnalyzer:
    def __init__(self, pptx_file):
        """
        アップロードされたPPTXファイルを読み込み解析する
        :param pptx_file: ファイルライクオブジェクトまたはパス
        """
        self.prs = Presentation(pptx_file)
        self._theme_colors = []
        self._font_styles = {"title": [], "body": []}
        self._logos = []
        self._backgrounds = []

    def extract_all(self):
        """
        全てのトンマナ情報を一括で抽出して返す
        """
        return {
            "layouts": self.extract_layouts(),
            "theme_colors": self.extract_theme_colors(),
            "font_styles": self.extract_font_styles(),
            "backgrounds": self.extract_backgrounds(),
            "logos": self.extract_logos(),
            "slide_size": self.extract_slide_size(),
        }

    def extract_layouts(self):
        """
        スライドマスターから利用可能なレイアウトの情報を抽出する
        :return: レイアウト情報のリスト
        """
        layouts_info = []
        for i, layout in enumerate(self.prs.slide_layouts):
            placeholders = []
            for shape in layout.placeholders:
                ph_info = {
                    "id": shape.placeholder_format.idx,
                    "name": shape.name,
                    "type": str(shape.placeholder_format.type),
                    "left": shape.left,
                    "top": shape.top,
                    "width": shape.width,
                    "height": shape.height,
                }
                placeholders.append(ph_info)

            layouts_info.append({
                "index": i,
                "name": layout.name,
                "placeholders": placeholders
            })
        return layouts_info

    def extract_theme_colors(self):
        """
        スライド内で使用されている色をRGBで収集する
        """
        colors = set()
        # 既存スライドからテキスト色を収集
        for slide in self.prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for para in shape.text_frame.paragraphs:
                        for run in para.runs:
                            try:
                                if run.font.color and run.font.color.rgb:
                                    colors.add(str(run.font.color.rgb))
                            except (AttributeError, TypeError):
                                pass
        # レイアウトからも収集
        for layout in self.prs.slide_layouts:
            for shape in layout.placeholders:
                if shape.has_text_frame:
                    for para in shape.text_frame.paragraphs:
                        for run in para.runs:
                            try:
                                if run.font.color and run.font.color.rgb:
                                    colors.add(str(run.font.color.rgb))
                            except (AttributeError, TypeError):
                                pass
        self._theme_colors = list(colors)
        return self._theme_colors

    def extract_font_styles(self):
        """
        タイトル用・本文用のフォントスタイルを抽出する
        戻り値: {"title": [...], "body": [...]}
        各要素は {"font_name": str, "font_size_pt": float, "bold": bool, "color_rgb": str}
        """
        title_fonts = []
        body_fonts = []

        # 既存スライドからフォント情報を収集
        for slide in self.prs.slides:
            for shape in slide.shapes:
                if not shape.has_text_frame:
                    continue
                is_title = (
                    hasattr(shape, "placeholder_format")
                    and shape.placeholder_format is not None
                    and shape.placeholder_format.idx == 0  # タイトルプレースホルダー
                )
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        font_info = self._extract_run_font(run)
                        if font_info:
                            if is_title:
                                title_fonts.append(font_info)
                            else:
                                body_fonts.append(font_info)

        # レイアウトからも収集（スライドがない場合のフォールバック）
        if not title_fonts and not body_fonts:
            for layout in self.prs.slide_layouts:
                for shape in layout.placeholders:
                    if not shape.has_text_frame:
                        continue
                    is_title = shape.placeholder_format.idx == 0
                    for para in shape.text_frame.paragraphs:
                        for run in para.runs:
                            font_info = self._extract_run_font(run)
                            if font_info:
                                if is_title:
                                    title_fonts.append(font_info)
                                else:
                                    body_fonts.append(font_info)

        # 重複を除いて代表値を決定
        self._font_styles = {
            "title": self._deduplicate_fonts(title_fonts),
            "body": self._deduplicate_fonts(body_fonts),
        }
        return self._font_styles

    def _extract_run_font(self, run):
        """
        テキストランからフォント情報を抽出する
        """
        try:
            font = run.font
            font_name = font.name
            font_size = font.size
            bold = font.bold
            color_rgb = None
            try:
                if font.color and font.color.rgb:
                    color_rgb = str(font.color.rgb)
            except (AttributeError, TypeError):
                pass

            return {
                "font_name": font_name,
                "font_size_pt": round(font_size / Pt(1), 1) if font_size else None,
                "bold": bold,
                "color_rgb": color_rgb,
            }
        except Exception:
            return None

    def _deduplicate_fonts(self, font_list):
        """
        フォント情報の重複を除去する
        """
        seen = set()
        unique = []
        for f in font_list:
            key = (f.get("font_name"), f.get("font_size_pt"), f.get("bold"))
            if key not in seen:
                seen.add(key)
                unique.append(f)
        return unique

    def extract_backgrounds(self):
        """
        スライドの背景情報を抽出する
        """
        backgrounds = []
        for i, slide in enumerate(self.prs.slides):
            bg_info = {"slide_index": i, "type": "unknown", "color_rgb": None}
            try:
                bg = slide.background
                fill = bg.fill
                if fill.type is not None:
                    bg_info["fill_type"] = str(fill.type)
                    if fill.type == 1:  # Solid fill
                        try:
                            bg_info["type"] = "solid"
                            bg_info["color_rgb"] = str(fill.fore_color.rgb)
                        except Exception:
                            pass
                    elif fill.type == 2:  # Gradient
                        bg_info["type"] = "gradient"
                    elif fill.type == 6:  # Picture
                        bg_info["type"] = "picture"
            except Exception:
                pass
            backgrounds.append(bg_info)

        # レイアウトの背景も収集（スライドがない場合）
        if not backgrounds:
            for i, layout in enumerate(self.prs.slide_layouts):
                bg_info = {"layout_index": i, "type": "unknown"}
                try:
                    bg = layout.background
                    fill = bg.fill
                    if fill.type is not None:
                        bg_info["fill_type"] = str(fill.type)
                except Exception:
                    pass
                backgrounds.append(bg_info)

        self._backgrounds = backgrounds
        return self._backgrounds

    def extract_logos(self):
        """
        スライド内の固定位置にある画像（ロゴ候補）を検出する
        小さめの画像（スライドの25%未満の面積）をロゴと判定
        """
        logos = []
        slide_width = self.prs.slide_width
        slide_height = self.prs.slide_height
        slide_area = slide_width * slide_height

        for slide in self.prs.slides:
            for shape in slide.shapes:
                if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    shape_area = shape.width * shape.height
                    # スライド面積の25%未満の画像はロゴ候補
                    if shape_area < slide_area * 0.25:
                        logo_info = {
                            "name": shape.name,
                            "left": shape.left,
                            "top": shape.top,
                            "width": shape.width,
                            "height": shape.height,
                            "image_content_type": None,
                        }
                        try:
                            logo_info["image_content_type"] = shape.image.content_type
                        except Exception:
                            pass
                        logos.append(logo_info)

        # 重複位置のロゴを統合（複数スライドで同じ位置にあるもの）
        self._logos = self._deduplicate_logos(logos)
        return self._logos

    def _deduplicate_logos(self, logos):
        """
        同じ位置にあるロゴを統合する
        """
        seen = set()
        unique = []
        for logo in logos:
            # 位置で判定（EMU単位で近い位置は同じと見なす、100000EMU≈1cm の許容範囲）
            key = (
                round(logo["left"] / 100000),
                round(logo["top"] / 100000),
                round(logo["width"] / 100000),
                round(logo["height"] / 100000),
            )
            if key not in seen:
                seen.add(key)
                unique.append(logo)
        return unique

    def extract_slide_size(self):
        """
        スライドのサイズを返す
        """
        return {
            "width_emu": self.prs.slide_width,
            "height_emu": self.prs.slide_height,
            "width_inches": round(self.prs.slide_width / Emu(914400), 2),
            "height_inches": round(self.prs.slide_height / Emu(914400), 2),
        }

    def get_summary(self):
        """
        抽出したトンマナ・レイアウトのサマリーを返す（人間が読みやすい形式）
        """
        info = self.extract_all()
        lines = []
        lines.append(f"📐 スライドサイズ: {info['slide_size']['width_inches']}\" x {info['slide_size']['height_inches']}\"")
        lines.append(f"📑 レイアウト数: {len(info['layouts'])}")
        for layout in info['layouts']:
            lines.append(f"  - [{layout['index']}] {layout['name']} (プレースホルダー: {len(layout['placeholders'])}個)")
        lines.append(f"🎨 検出されたカラー: {', '.join(info['theme_colors']) if info['theme_colors'] else '(なし)'}")

        if info['font_styles']['title']:
            lines.append("🔤 タイトルフォント:")
            for f in info['font_styles']['title']:
                lines.append(f"  - {f.get('font_name', '不明')} / {f.get('font_size_pt', '?')}pt / 太字={f.get('bold')}")
        if info['font_styles']['body']:
            lines.append("🔤 本文フォント:")
            for f in info['font_styles']['body']:
                lines.append(f"  - {f.get('font_name', '不明')} / {f.get('font_size_pt', '?')}pt / 太字={f.get('bold')}")
        if info['logos']:
            lines.append(f"🏷️ ロゴ候補: {len(info['logos'])}個検出")
        if info['backgrounds']:
            bg_types = set(bg.get('type', '不明') for bg in info['backgrounds'])
            lines.append(f"🖼️ 背景: {', '.join(bg_types)}")

        return "\n".join(lines)
