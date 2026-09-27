# 📊 PPTX Generator - AI スライド自動生成

AIを活用して、既存のPowerPointファイルのデザイン（トンマナ）を踏襲し、データに基づく新しいプレゼンテーション資料を自動生成するWebアプリケーションです。

## ✨ 主な機能

- 🎨 **トンマナ自動抽出**: ベースPPTXからカラーパレット、フォント、レイアウト、ロゴを自動抽出
- 📂 **多形式データ対応**: Excel (.xlsx)、CSV、テキスト (.txt)、PDF に対応
- 🤖 **AI自動生成**: OpenAI (GPT-4o) / Google Gemini に対応（選択可能）
- 📊 **グラフ自動生成**: 棒グラフ、折れ線、円グラフ、ドーナツ、面グラフなど
- 📋 **テーブル自動生成**: データをテーブル形式でスライドに配置
- 🔤 **フォント・色の継承**: 元資料のデザインを忠実に再現

## 🚀 セットアップ

### 1. 依存関係のインストール
```bash
pip install -r requirements.txt
```

### 2. 環境変数（オプション）
`.env.example` をコピーして `.env` を作成し、デフォルトのAPIキーを設定できます。
```bash
copy .env.example .env
```

### 3. アプリの起動
```bash
streamlit run app.py
```

ブラウザで `http://localhost:8501` が自動的に開きます。

## 📖 使い方

### Step 1: ベースPPTXをアップロード
テンプレートとなるPowerPointファイルをアップロードします。
→ カラー、フォント、レイアウト、ロゴ情報が自動抽出されプレビューされます。

### Step 2: データファイルをアップロード
スライドの内容となるデータをアップロードします。
- Excel: 表形式データ
- CSV: カンマ区切りデータ
- TXT: テキストデータ
- PDF: PDFテキスト

### Step 3: 生成指示を入力
AIに対して、どのようなスライドを作成したいかを自由記述で指示します。

**指示の例:**
```
- 売上データの推移を棒グラフで1枚
- 地域別の比較をテーブルで1枚
- 今後の課題と対策を箇条書きで1枚
- 合計3枚のスライドにまとめてください
```

### Step 4: 生成 & ダウンロード
「PPTXを自動生成する」ボタンを押して、完成した資料をダウンロードします。

## 📁 ファイル構成

```
pptx-generator/
├── app.py                    # Streamlit メインアプリ
├── requirements.txt          # 依存パッケージ
├── .env.example              # APIキー設定例
├── README.md                 # このファイル
└── modules/
    ├── __init__.py
    ├── pptx_analyzer.py      # トンマナ抽出（色・フォント・レイアウト・ロゴ）
    ├── data_loader.py        # データ読み込み（Excel/CSV/TXT/PDF）
    ├── ai_engine.py          # AI解析エンジン（Gemini/OpenAI）
    └── pptx_builder.py       # PPTX生成（トンマナ適用）
```

## ⚙️ 技術スタック

| カテゴリ | 技術 |
|---------|------|
| UI | Streamlit |
| PPTX操作 | python-pptx |
| データ処理 | pandas, openpyxl |
| PDF読取 | pdfplumber |
| AI | OpenAI API, Google Gemini API |

streamlit>=1.30.0
python-pptx>=0.6.22
pandas>=2.0.0
openpyxl>=3.1.2
tabulate>=0.9.0
pdfplumber>=0.10.3
openai>=1.10.0
google-generativeai>=0.3.2
python-dotenv>=1.0.0
pydantic>=2.0.0

__pycache__/
*.pyc
.env
.venv/
venv/
test_*.py

"""
app.py - PPTXジェネレーター メインアプリケーション

Streamlit Webアプリとして動作する。
5つのステップでパワーポイント資料を自動生成する:
  Step 1: ベースPPTXアップロード → トンマナ抽出・プレビュー
  Step 2: データファイルアップロード → データプレビュー
  Step 3: スライド枚数 & 生成指示入力
  Step 4: 構成案の確認 → 承認 or 修正指示
  Step 5: PPTX自動生成 & ダウンロード
"""

import streamlit as st
import os
import io

from modules.pptx_analyzer import PPTXAnalyzer
from modules.data_loader import DataLoader
from modules.ai_engine import AIEngine
from modules.pptx_builder import PPTXBuilder

# ──────────────────────────────────────────────
# ページ設定
# ──────────────────────────────────────────────
st.set_page_config(
    page_title="PPTX Generator - AI スライド自動生成",
    page_icon="📊",
    layout="wide",
)

st.title("📊 PPTXジェネレーター")
st.caption("AIを活用して、デザインを踏襲したパワーポイント資料を自動生成します")

# ──────────────────────────────────────────────
# サイドバー: API設定
# ──────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ AI設定")
    api_type = st.radio(
        "AIエンジンの選択",
        ["OpenAI", "Gemini"],
        help="OpenAI: GPT-4o使用、Gemini: Gemini 1.5 Pro使用"
    )

    # 環境変数からのデフォルト値取得
    env_key_name = "OPENAI_API_KEY" if api_type == "OpenAI" else "GEMINI_API_KEY"
    default_key = os.environ.get(env_key_name, "")
    api_key = st.text_input(
        f"{api_type} API Key",
        value=default_key,
        type="password",
        help="APIキーはセッション内のみで使用され、保存されません"
    )

    if not api_key:
        st.warning("⚠️ APIキーを入力してください。")

    st.divider()
    st.markdown("### 📖 使い方")
    st.markdown("""
    1. **Step 1**: テンプレートPPTXをアップロード
    2. **Step 2**: データファイルをアップロード
    3. **Step 3**: スライド枚数と生成指示を入力
    4. **Step 4**: 構成案を確認 → 承認 or 修正
    5. **Step 5**: PPTXを生成してダウンロード
    """)

# ──────────────────────────────────────────────
# セッションステートの初期化
# ──────────────────────────────────────────────
defaults = {
    "tone_and_manner": None,
    "layout_info": None,
    "base_pptx_bytes": None,
    "data_text": "",
    "data_summary": "",
    "outline": None,           # 現在の構成案
    "outline_approved": False,  # 構成案が承認されたか
    "revision_count": 0,        # 修正回数（UIの再描画制御用）
}
for key, default_val in defaults.items():
    if key not in st.session_state:
        st.session_state[key] = default_val

# ──────────────────────────────────────────────
# Step 1: ベースPPTXのアップロード
# ──────────────────────────────────────────────
st.header("📑 Step 1: ベースPPTXのアップロード")
st.markdown("テンプレートとなるPowerPointファイルをアップロードしてください。デザイン（トンマナ）を自動抽出します。")

base_pptx = st.file_uploader(
    "PPTXファイルを選択",
    type=["pptx"],
    key="base_pptx_uploader"
)

if base_pptx:
    try:
        st.session_state.base_pptx_bytes = base_pptx.read()
        analyzer = PPTXAnalyzer(io.BytesIO(st.session_state.base_pptx_bytes))
        st.session_state.tone_and_manner = analyzer.extract_all()
        st.session_state.layout_info = st.session_state.tone_and_manner["layouts"]

        st.success("✅ トンマナの抽出が完了しました！")

        with st.expander("🎨 抽出されたトンマナ情報を確認", expanded=True):
            col1, col2 = st.columns(2)

            with col1:
                st.subheader("📐 レイアウト一覧")
                for layout in st.session_state.layout_info:
                    ph_count = len(layout['placeholders'])
                    st.markdown(f"- **[{layout['index']}]** {layout['name']} ({ph_count}個のプレースホルダー)")

                size = st.session_state.tone_and_manner.get("slide_size", {})
                st.markdown(f"**スライドサイズ**: {size.get('width_inches', '?')}\" × {size.get('height_inches', '?')}\"")

            with col2:
                st.subheader("🎨 検出されたカラー")
                colors = st.session_state.tone_and_manner.get("theme_colors", [])
                if colors:
                    color_html = " ".join(
                        [f'<span style="display:inline-block;width:30px;height:30px;background:#{c};border:1px solid #ccc;margin:2px;border-radius:4px;" title="#{c}"></span>'
                         for c in colors]
                    )
                    st.markdown(color_html, unsafe_allow_html=True)
                    st.caption(", ".join([f"#{c}" for c in colors]))
                else:
                    st.info("カラー情報が検出されませんでした")

                st.subheader("🔤 フォント情報")
                font_styles = st.session_state.tone_and_manner.get("font_styles", {})
                title_fonts = font_styles.get("title", [])
                body_fonts = font_styles.get("body", [])

                if title_fonts:
                    st.markdown("**タイトル:**")
                    for f in title_fonts:
                        st.markdown(f"  `{f.get('font_name', '不明')}` / {f.get('font_size_pt', '?')}pt / 太字={'✅' if f.get('bold') else '❌'}")
                if body_fonts:
                    st.markdown("**本文:**")
                    for f in body_fonts:
                        st.markdown(f"  `{f.get('font_name', '不明')}` / {f.get('font_size_pt', '?')}pt / 太字={'✅' if f.get('bold') else '❌'}")

                logos = st.session_state.tone_and_manner.get("logos", [])
                if logos:
                    st.subheader("🏷️ ロゴ候補")
                    st.markdown(f"{len(logos)}個のロゴ画像を検出しました")

    except Exception as e:
        st.error(f"❌ PPTXの解析に失敗しました: {e}")

st.divider()

# ──────────────────────────────────────────────
# Step 2: データファイルのアップロード
# ──────────────────────────────────────────────
st.header("📂 Step 2: データファイルのアップロード")
st.markdown("スライドの内容となるデータファイルをアップロードしてください。")

data_file = st.file_uploader(
    "データファイルを選択 (Excel, CSV, TXT, PDF)",
    type=["csv", "xlsx", "xls", "txt", "pdf"],
    key="data_file_uploader"
)

if data_file:
    content, summary = DataLoader.load_data(data_file)
    if "エラー" not in summary:
        st.session_state.data_text = content
        st.session_state.data_summary = summary
        st.success("✅ データの読み込みが完了しました。")
        with st.expander("📋 データのプレビュー", expanded=False):
            st.text(summary)
            st.divider()
            st.text(content[:3000] + ("..." if len(content) > 3000 else ""))
    else:
        st.error(f"❌ {summary}")

st.divider()

# ──────────────────────────────────────────────
# Step 3: スライド枚数 & 生成指示
# ──────────────────────────────────────────────
st.header("✍️ Step 3: スライド枚数と生成指示")

col_left, col_right = st.columns([1, 3])

with col_left:
    st.subheader("📄 スライド枚数")
    num_slides = st.number_input(
        "生成するスライドの枚数",
        min_value=1,
        max_value=30,
        value=5,
        step=1,
        help="タイトルスライドとまとめスライドを含めた総枚数を指定してください"
    )
    st.caption(f"💡 {num_slides}枚のスライドを生成します")

with col_right:
    st.subheader("📝 生成指示")
    instructions = st.text_area(
        "AIへの指示（どのようなスライドを作成したいか）",
        placeholder="例:\n- 売上データの推移を棒グラフで表示\n- 地域別の比較をテーブルで表示\n- 今後の課題と対策を箇条書きでまとめる\n- 最後にまとめと提案を入れる",
        height=150,
        help="具体的に指示するほど精度が上がります。グラフの種類、まとめ方なども記載してください。"
    )

st.divider()

# ──────────────────────────────────────────────
# Step 4: 構成案の確認 → 承認 or 修正
# ──────────────────────────────────────────────
st.header("📋 Step 4: スライド構成案の確認")
st.markdown("AIがスライドの構成案を提案します。内容を確認して承認するか、修正を指示してください。")

# 入力状態のチェック
input_ready = (
    api_key
    and st.session_state.base_pptx_bytes
    and st.session_state.data_text
    and instructions
)

if not input_ready:
    # 不足している項目を表示
    missing = []
    if not api_key:
        missing.append("APIキー")
    if not st.session_state.base_pptx_bytes:
        missing.append("ベースPPTX")
    if not st.session_state.data_text:
        missing.append("データファイル")
    if not instructions:
        missing.append("生成指示")
    st.info(f"⏳ 以下を入力してください: {', '.join(missing)}")

# ── 構成案を生成するボタン ──
if st.button(
    "💡 構成案を生成する",
    type="secondary",
    disabled=not input_ready
):
    with st.spinner("🤖 AIがスライド構成案を作成中..."):
        try:
            ai_engine = AIEngine(api_type, api_key)
            outline = ai_engine.generate_outline(
                data_text=st.session_state.data_text,
                instructions=instructions,
                layout_info=st.session_state.layout_info,
                num_slides=num_slides
            )
            st.session_state.outline = outline
            st.session_state.outline_approved = False
            st.session_state.revision_count = 0
            st.rerun()
        except Exception as e:
            st.error(f"❌ 構成案の生成に失敗しました: {str(e)}")

# ── 構成案が存在する場合、表示する ──
if st.session_state.outline:
    outline_items = st.session_state.outline.get("outline", [])

    if not st.session_state.outline_approved:
        st.success(f"📄 {len(outline_items)} 枚のスライド構成案が生成されました")

        # 構成案をカード形式で表示
        for item in outline_items:
            slide_num = item.get("slide_number", "?")
            title = item.get("title", "無題")
            description = item.get("description", "")
            elements = item.get("elements", [])
            layout_idx = item.get("layout_index", 0)

            with st.container():
                st.markdown(f"""
                <div style="border:1px solid #ddd; border-radius:8px; padding:16px; margin-bottom:12px; background:#fafafa;">
                    <h4 style="margin:0 0 8px 0;">📄 スライド {slide_num}: {title}</h4>
                    <p style="color:#555; margin:0 0 8px 0;">{description}</p>
                    <p style="margin:0;"><strong>使用要素:</strong> {' / '.join([f'<code>{e}</code>' for e in elements])}</p>
                    <p style="margin:0; color:#888; font-size:0.85em;">レイアウト: [{layout_idx}]</p>
                </div>
                """, unsafe_allow_html=True)

        # JSON詳細の展開表示
        with st.expander("🔍 構成案の詳細データ (JSON)"):
            st.json(st.session_state.outline)

        st.markdown("---")

        # ── 承認 or 修正 の選択 ──
        st.subheader("✅ この構成案でよろしいですか？")

        col_approve, col_revise = st.columns(2)

        with col_approve:
            if st.button(
                "✅ この構成で作成する",
                type="primary",
                use_container_width=True,
            ):
                st.session_state.outline_approved = True
                st.rerun()

        with col_revise:
            st.markdown("**修正したい場合は下に指示を入力してください:**")

        # 修正指示入力エリア
        revision_input = st.text_area(
            "修正指示",
            placeholder="例:\n- 2枚目のグラフを折れ線グラフに変更して\n- 3枚目と4枚目の順番を入れ替えて\n- 最後にQ&Aスライドを追加して",
            height=100,
            key=f"revision_input_{st.session_state.revision_count}",
            help="修正したい部分を具体的に指示してください。"
        )

        if st.button(
            "🔄 修正して再提案する",
            disabled=not revision_input,
            use_container_width=True,
        ):
            with st.spinner("🤖 AIが構成案を修正中..."):
                try:
                    ai_engine = AIEngine(api_type, api_key)
                    revised = ai_engine.revise_outline(
                        data_text=st.session_state.data_text,
                        instructions=instructions,
                        layout_info=st.session_state.layout_info,
                        num_slides=num_slides,
                        current_outline=st.session_state.outline,
                        revision_request=revision_input
                    )
                    st.session_state.outline = revised
                    st.session_state.revision_count += 1
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ 修正に失敗しました: {str(e)}")

    else:
        # 承認済みの場合
        st.success("✅ 構成案が承認されました！ Step 5 でPPTXを生成してください。")
        with st.expander("📋 承認済みの構成案を確認"):
            for item in outline_items:
                slide_num = item.get("slide_number", "?")
                title = item.get("title", "無題")
                elements = item.get("elements", [])
                st.markdown(f"**スライド {slide_num}**: {title} — {', '.join(elements)}")

        # 構成案をリセットするボタン
        if st.button("🔙 構成案をやり直す"):
            st.session_state.outline = None
            st.session_state.outline_approved = False
            st.session_state.revision_count = 0
            st.rerun()

st.divider()

# ──────────────────────────────────────────────
# Step 5: PPTXの生成
# ──────────────────────────────────────────────
st.header("🚀 Step 5: PPTXの生成")

# 入力状態のチェック表示
col1, col2, col3, col4, col5 = st.columns(5)
with col1:
    if api_key:
        st.success("✅ APIキー")
    else:
        st.error("❌ APIキー")
with col2:
    if st.session_state.base_pptx_bytes:
        st.success("✅ ベースPPTX")
    else:
        st.error("❌ ベースPPTX")
with col3:
    if st.session_state.data_text:
        st.success("✅ データ")
    else:
        st.error("❌ データ")
with col4:
    if instructions:
        st.success("✅ 生成指示")
    else:
        st.error("❌ 生成指示")
with col5:
    if st.session_state.outline_approved:
        st.success("✅ 構成案承認済")
    else:
        st.error("❌ 構成案未承認")

# 全条件が揃っているか確認
all_ready = (
    api_key
    and st.session_state.base_pptx_bytes
    and st.session_state.data_text
    and instructions
    and st.session_state.outline_approved
)

if st.button("🎯 PPTXを自動生成する", type="primary", disabled=not all_ready):
    progress = st.progress(0, text="準備中...")

    try:
        # ステップ1: 承認済み構成案をもとにAIで詳細JSON生成
        progress.progress(10, text="🤖 承認済み構成案をもとに詳細スライドデータを生成中...")
        ai_engine = AIEngine(api_type, api_key)
        slide_json = ai_engine.generate_slides_structure(
            data_text=st.session_state.data_text,
            instructions=instructions,
            layout_info=st.session_state.layout_info,
            num_slides=num_slides,
            approved_outline=st.session_state.outline
        )

        progress.progress(50, text="📊 スライドデータの生成が完了しました")

        # AI生成結果のプレビュー
        with st.expander("🔍 AI生成結果 (JSON構成)を確認"):
            st.json(slide_json)

        generated_count = len(slide_json.get("slides", []))
        st.info(f"📄 {generated_count} 枚のスライドを生成します")

        # ステップ2: PPTX生成
        progress.progress(60, text="📝 PPTXファイルを生成中...")
        builder = PPTXBuilder(
            io.BytesIO(st.session_state.base_pptx_bytes),
            tone_and_manner=st.session_state.tone_and_manner
        )
        builder.build(slide_json)

        progress.progress(90, text="📦 ファイルを準備中...")
        output_io = builder.get_output_io()

        progress.progress(100, text="✨ 完了！")

        st.success(f"✨ {generated_count} 枚のスライドを生成しました！")

        # ダウンロードボタン
        st.download_button(
            label="📥 生成されたPPTXをダウンロード",
            data=output_io,
            file_name="generated_presentation.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            type="primary",
        )

    except Exception as e:
        progress.empty()
        st.error(f"❌ エラーが発生しました: {str(e)}")
        with st.expander("エラー詳細"):
            import traceback
            st.code(traceback.format_exc())

# ──────────────────────────────────────────────
# フッター
# ──────────────────────────────────────────────
st.divider()
st.caption("PPTXジェネレーター v1.1 | Powered by AI + python-pptx")
