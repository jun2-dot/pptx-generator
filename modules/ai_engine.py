"""
ai_engine.py - AI解析エンジンモジュール

Google Gemini API / OpenAI API を使用して、
データとユーザー指示からスライド構成をJSON形式で生成する。

対応出力:
- タイトルスライド (title + subtitle)
- コンテンツスライド (title + body箇条書き)
- グラフスライド (title + chart)
- テーブルスライド (title + table)
- 複合スライド (title + body + chart or table)
"""

import json
import os
import google.generativeai as genai
from openai import OpenAI


class AIEngine:
    def __init__(self, api_type, api_key):
        """
        AI APIクライアントを初期化する
        :param api_type: "OpenAI" または "Gemini"
        :param api_key: APIキー
        """
        self.api_type = api_type
        self.api_key = api_key

        if self.api_type == "OpenAI":
            self.client = OpenAI(api_key=self.api_key)
        elif self.api_type == "Gemini":
            genai.configure(api_key=self.api_key)
            self.gemini_model = genai.GenerativeModel('gemini-1.5-pro')

    def generate_slides_structure(self, data_text, instructions, layout_info,
                                    num_slides=None, approved_outline=None):
        """
        AIに指示とデータ、レイアウト情報を渡し、スライド構成のJSONを生成する

        :param data_text: データの内容（テキスト形式）
        :param instructions: ユーザーからの指示テキスト
        :param layout_info: PPTXAnalyzerから抽出されたレイアウト情報
        :param num_slides: スライド枚数の指定（Noneの場合はAIに任せる）
        :param approved_outline: 承認済みの構成案（これに基づいて詳細JSONを生成）
        :return: スライド構成のdict
        """
        system_prompt = self._build_system_prompt(layout_info, num_slides)

        user_prompt = f"【提供データ】\n{data_text}\n\n【ユーザーの指示】\n{instructions}"

        # 承認済み構成案がある場合、それに従うよう指示を追加
        if approved_outline:
            outline_text = json.dumps(approved_outline, ensure_ascii=False, indent=2)
            user_prompt += f"\n\n【承認済みスライド構成案（この構成に厳密に従ってください）】\n{outline_text}"

        try:
            if self.api_type == "OpenAI":
                result_text = self._call_openai(system_prompt, user_prompt)
            elif self.api_type == "Gemini":
                result_text = self._call_gemini(system_prompt, user_prompt)
            else:
                raise ValueError("サポートされていないAPIタイプです")

            return self._parse_json_response(result_text)

        except json.JSONDecodeError as e:
            raise Exception(f"AIの応答をJSONとして解析できませんでした: {str(e)}")
        except Exception as e:
            raise Exception(f"AI処理中にエラーが発生しました: {str(e)}")

    def generate_outline(self, data_text, instructions, layout_info, num_slides):
        """
        スライドの構成案（アウトライン）を生成する。
        各スライドのタイトル・概要・使用する要素を提示する。

        :param data_text: データの内容（テキスト形式）
        :param instructions: ユーザーからの指示テキスト
        :param layout_info: レイアウト情報
        :param num_slides: スライド枚数
        :return: 構成案のdict
        """
        system_prompt = self._build_outline_prompt(layout_info, num_slides)
        user_prompt = f"【提供データ】\n{data_text}\n\n【ユーザーの指示】\n{instructions}"

        try:
            if self.api_type == "OpenAI":
                result_text = self._call_openai(system_prompt, user_prompt)
            elif self.api_type == "Gemini":
                result_text = self._call_gemini(system_prompt, user_prompt)
            else:
                raise ValueError("サポートされていないAPIタイプです")

            return self._parse_json_response(result_text)

        except json.JSONDecodeError as e:
            raise Exception(f"構成案のJSON解析に失敗しました: {str(e)}")
        except Exception as e:
            raise Exception(f"構成案の生成中にエラーが発生しました: {str(e)}")

    def revise_outline(self, data_text, instructions, layout_info,
                       num_slides, current_outline, revision_request):
        """
        ユーザーの修正指示に基づいて構成案を修正する。

        :param data_text: データの内容
        :param instructions: 元の指示テキスト
        :param layout_info: レイアウト情報
        :param num_slides: スライド枚数
        :param current_outline: 現在の構成案
        :param revision_request: ユーザーの修正指示
        :return: 修正された構成案のdict
        """
        system_prompt = self._build_outline_prompt(layout_info, num_slides)

        current_outline_text = json.dumps(current_outline, ensure_ascii=False, indent=2)
        user_prompt = (
            f"【提供データ】\n{data_text}\n\n"
            f"【元の指示】\n{instructions}\n\n"
            f"【現在の構成案】\n{current_outline_text}\n\n"
            f"【修正指示】\n{revision_request}\n\n"
            f"上記の修正指示に従って構成案を修正してください。"
            f"修正が必要ない部分はそのまま維持してください。"
        )

        try:
            if self.api_type == "OpenAI":
                result_text = self._call_openai(system_prompt, user_prompt)
            elif self.api_type == "Gemini":
                result_text = self._call_gemini(system_prompt, user_prompt)
            else:
                raise ValueError("サポートされていないAPIタイプです")

            return self._parse_json_response(result_text)

        except json.JSONDecodeError as e:
            raise Exception(f"修正案のJSON解析に失敗しました: {str(e)}")
        except Exception as e:
            raise Exception(f"構成案の修正中にエラーが発生しました: {str(e)}")

    def _build_outline_prompt(self, layout_info, num_slides):
        """
        構成案（アウトライン）生成用のシステムプロンプトを構築する
        """
        layout_summary = json.dumps(layout_info, ensure_ascii=False, indent=2)

        return f"""あなたはプロのプレゼンテーションクリエイターです。
提供されたデータとユーザーの指示に基づき、PowerPointスライドの構成案（アウトライン）をJSON形式で提案してください。

【スライド枚数】
必ず {num_slides} 枚のスライドで構成してください。

【利用可能なレイアウト】
{layout_summary}

【出力ルール】
1. 必ず以下のJSON形式で出力してください。マークダウンの```json等は含めないでください。
2. 各スライドには slide_number, title, description, elements, layout_index を含めてください。
3. descriptionには、そのスライドで伝えたい内容の概要を2〜3文で書いてください。
4. elementsには、使用する要素（箇条書き、グラフ、テーブルなど）をリストで記載してください。
5. 日本語で出力してください。

【JSON形式】
{{
  "outline": [
    {{
      "slide_number": 1,
      "title": "スライドのタイトル",
      "description": "このスライドの概要説明。何を伝えるか、どのようにまとめるか。",
      "elements": ["タイトルとサブタイトル"],
      "layout_index": 0
    }},
    {{
      "slide_number": 2,
      "title": "売上推移",
      "description": "過去3年間の売上推移を棒グラフで可視化し、成長トレンドを示す。",
      "elements": ["棒グラフ", "簡単な説明テキスト"],
      "layout_index": 1
    }}
  ]
}}

【elementsに使える要素の例】
- "タイトルとサブタイトル": タイトルスライド用
- "箇条書き": 要点をリスト形式で表示
- "棒グラフ" / "折れ線グラフ" / "円グラフ" / "ドーナツグラフ" / "面グラフ": データの可視化
- "テーブル": 表形式でデータを表示
- "説明テキスト": 文章による説明

【注意事項】
- スライド枚数は必ず {num_slides} 枚にしてください。
- 各スライドの内容が重複しないようにしてください。
- データの特徴を活かした構成を提案してください。
- 最初のスライドはタイトルスライドにすることを推奨します。
- 最後のスライドはまとめ・結論にすることを推奨します。"""

    def _build_system_prompt(self, layout_info, num_slides=None):
        """
        AIへのシステムプロンプトを構築する
        """
        # レイアウト情報を簡潔にまとめる
        layout_summary = json.dumps(layout_info, ensure_ascii=False, indent=2)

        num_slides_instruction = ""
        if num_slides:
            num_slides_instruction = f"\n7. スライドは必ず {num_slides} 枚で構成してください。"

        return f"""あなたはプロのプレゼンテーションクリエイターです。
提供されたデータとユーザーの指示に基づき、PowerPointスライドの構成をJSON形式で生成してください。

【利用可能なレイアウト】
{layout_summary}

【出力ルール】
1. 必ず以下のJSON形式で出力してください。マークダウンの```json等は含めないでください。
2. layout_indexは上記レイアウト一覧のindexから最適なものを選んでください。
3. 各スライドにはtitleを必ず含めてください。
4. body（箇条書き）、chart（グラフ）、table（テーブル）は必要に応じて含めてください。
5. データの数値はそのまま正確に使用してください。
6. 日本語で出力してください。{num_slides_instruction}

【JSON形式】
{{
  "slides": [
    {{
      "layout_index": 0,
      "title": "スライドタイトル",
      "subtitle": "サブタイトル（タイトルスライドの場合のみ）",
      "body": ["箇条書き項目1", "箇条書き項目2", "箇条書き項目3"],
      "chart": {{
        "type": "bar",
        "categories": ["カテゴリA", "カテゴリB", "カテゴリC"],
        "series": [
          {{"name": "系列名", "values": [100, 200, 300]}}
        ]
      }},
      "table": {{
        "headers": ["列1", "列2", "列3"],
        "rows": [
          ["値1", "値2", "値3"],
          ["値4", "値5", "値6"]
        ]
      }}
    }}
  ]
}}

【chartのtype一覧】
- "bar": 棒グラフ（縦棒）
- "stacked_bar": 積み上げ棒グラフ
- "line": 折れ線グラフ
- "line_markers": マーカー付き折れ線グラフ
- "pie": 円グラフ
- "doughnut": ドーナツグラフ
- "area": 面グラフ

【注意事項】
- chartとtableが不要なスライドでは、そのキーを含めないでください。
- 1つのスライドにchartとtableの両方を含めることもできます。
- bodyの各項目は簡潔で分かりやすい表現にしてください。
- データに基づく分析や洞察も適宜追加してください。
- 承認済みスライド構成案が提供された場合は、その構成に厳密に従ってください。"""

    def _call_openai(self, system_prompt, user_prompt):
        """
        OpenAI APIを呼び出す
        """
        response = self.client.chat.completions.create(
            model="gpt-4o",
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            temperature=0.7,
            max_tokens=4096,
        )
        return response.choices[0].message.content

    def _call_gemini(self, system_prompt, user_prompt):
        """
        Google Gemini APIを呼び出す
        """
        prompt = system_prompt + "\n\n" + user_prompt
        response = self.gemini_model.generate_content(
            prompt,
            generation_config={
                "response_mime_type": "application/json",
                "temperature": 0.7,
                "max_output_tokens": 4096,
            }
        )
        return response.text

    def _parse_json_response(self, result_text):
        """
        AIの応答テキストをJSONとしてパースする
        マークダウンブロックの除去などのクリーンアップも行う
        """
        result_text = result_text.strip()

        # マークダウンのコードブロックを除去
        if result_text.startswith("```json"):
            result_text = result_text[7:]
        elif result_text.startswith("```"):
            result_text = result_text[3:]

        if result_text.endswith("```"):
            result_text = result_text[:-3]

        result_text = result_text.strip()

        parsed = json.loads(result_text)

        # バリデーション: slidesキーが存在するか確認
        if "slides" not in parsed:
            # slidesキーがない場合、全体をslidesで包む
            if isinstance(parsed, list):
                parsed = {"slides": parsed}
            else:
                parsed = {"slides": [parsed]}

        return parsed
