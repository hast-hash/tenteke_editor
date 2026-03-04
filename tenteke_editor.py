"""
Tenteke Editor  ―  高機能テキストエディタ + TEI/XML エディタ
  ・タブで複数ファイルを同時開き (中ボタンクリックで閉じる)
  ・ドラッグ＆ドロップでファイルを開く (pip install tkinterdnd2 が必要)
  ・常に画面下まで行番号を表示
  ・正規表現検索・置換  (Ctrl+H)
  ・自動保存 (バックグラウンドスレッド)
  ・マクロ記録・管理 (Python スクリプト)
  ── TEI/XML 機能 ──────────────────────────────────────────────────
  ・XML/TEI シンタックスハイライト (タグ・属性・コメント・CDATAなど色分け)
  ・タグ自動閉じ (> 入力時に </tag> を自動挿入してカーソルを中へ)
  ・XML バリデーション / インデント整形 (Ctrl+Shift+F)
  ・TEI 構造ツリーパネル (F6 で表示/非表示・要素クリックでジャンプ)
  ・TEI スニペットバー (F10 で表示/非表示)
      → テキスト選択 → ボタンクリック でタグを自動挿入・折り返し
      → 選択なしでクリック でカーソル位置にテンプレート挿入
  ・XPath 検索ダイアログ (Ctrl+Shift+X)
  ・TEI 新規ドキュメントテンプレート
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from tkinter import font as tkfont
import os, re, json, time, threading, bisect
import xml.etree.ElementTree as ET
import xml.sax
import xml.sax.handler
import io

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    HAS_DND = True
except ImportError:
    HAS_DND = False

# ===========================================================
# 定数
# ===========================================================
AUTOSAVE_DIR = os.path.join(os.path.expanduser("~"), ".tenteke_autosave")
MACRO_FILE   = os.path.join(os.path.expanduser("~"), ".tenteke_macros.json")
EDITOR_FONT  = ("Courier New", 11)
LINE_FONT    = ("Courier New", 11)

LIGHT = dict(
    text_bg="#ffffff", text_fg="#003366", insert="black",
    linenum_bg="#f0f0f0", linenum_fg="#999999",
    ruler_bg="#dddddd", ruler_fg="#555555",
)
DARK = dict(
    text_bg="#1e1e1e", text_fg="#ccddff", insert="white",
    linenum_bg="#252526", linenum_fg="#888888",
    ruler_bg="#333333", ruler_fg="#aaaaaa",
)

XML_COLORS_LIGHT = {
    "xml_tag":     "#0000bb",
    "xml_endtag":  "#0000bb",
    "xml_attr":    "#7f0055",
    "xml_value":   "#1a00cc",
    "xml_ns":      "#9f009f",
    "xml_comment": "#3f7f5f",
    "xml_cdata":   "#888888",
    "xml_pi":      "#808000",
    "xml_doctype": "#808000",
    "xml_error":   "#cc0000",
}
XML_COLORS_DARK = {
    "xml_tag":     "#88aaff",
    "xml_endtag":  "#88aaff",
    "xml_attr":    "#dd88ff",
    "xml_value":   "#ffcc77",
    "xml_ns":      "#cc88dd",
    "xml_comment": "#77cc88",
    "xml_cdata":   "#aaaaaa",
    "xml_pi":      "#eeee66",
    "xml_doctype": "#eeee66",
    "xml_error":   "#ff8888",
}

XML_EXTENSIONS    = {'.xml', '.tei', '.odd', '.rng', '.xsl', '.xslt', '.svg'}
TEI_VOID_ELEMENTS = {
    'lb', 'pb', 'cb', 'milestone', 'gap', 'space', 'graphic',
    'ptr', 'link', 'br', 'hr', 'img', 'input', 'meta',
}

# ===========================================================
# TEI スニペットバー: カテゴリ別アクション
#   (label, wrap_template)
#   wrap_template の {} に選択テキストが入る
#   {} が無い場合は void 扱いでそのまま挿入
# ===========================================================
TEI_BAR_ACTIONS = {
    "固有名詞": [
        ("人名",   "persName",  '<persName key="">{}</persName>'),
        ("地名",   "placeName", '<placeName key="">{}</placeName>'),
        ("組織名", "orgName",   '<orgName key="">{}</orgName>'),
        ("参照",   "rs",        '<rs type="" key="">{}</rs>'),
        ("日付",   "date",      '<date when="YYYY-MM-DD">{}</date>'),
        ("作品名", "title",     '<title level="m">{}</title>'),
        ("外国語", "foreign",   '<foreign xml:lang="">{}</foreign>'),
    ],
    "マーカー": [
        ("太字",   "hi-bold",   '<hi rend="bold">{}</hi>'),
        ("斜体",   "hi-italic", '<hi rend="italic">{}</hi>'),
        ("注釈",   "note",      '<note place="foot">{}</note>'),
        ("参照先", "ref",       '<ref target="">{}</ref>'),
        ("引用",   "q",         '<q>{}</q>'),
        ("引用文", "quote",     '<quote>\n  <p>{}</p>\n</quote>'),
    ],
    "本文構造": [
        ("段落",   "p",         '<p>{}</p>'),
        ("見出し", "head",      '<head>{}</head>'),
        ("章節",   "div",       '<div>\n  <head></head>\n  <p>{}</p>\n</div>'),
        ("詩行",   "l",         '<l>{}</l>'),
        ("詩行群", "lg",        '<lg>\n  <l>{}</l>\n</lg>'),
        ("台詞",   "sp",        '<sp>\n  <speaker></speaker>\n  <p>{}</p>\n</sp>'),
    ],
    "改行・丁": [
        ("改行",   "lb",        '<lb/>'),
        ("改丁",   "pb",        '<pb n=""/>'),
        ("改段",   "cb",        '<cb n=""/>'),
        ("区切",   "ms",        '<milestone unit="" n=""/>'),
    ],
    "校異・訂正": [
        ("校異",   "app",       '<app>\n  <lem wit="#A">{}</lem>\n  <rdg wit="#B"></rdg>\n</app>'),
        ("訂正",   "sic-corr",  '<choice>\n  <sic>{}</sic>\n  <corr></corr>\n</choice>'),
        ("正規化", "orig-reg",  '<choice>\n  <orig>{}</orig>\n  <reg></reg>\n</choice>'),
        ("欠損",   "gap",       '<gap reason="illegible" extent="1" unit="char"/>'),
        ("不明瞭", "unclear",   '<unclear reason="">{}</unclear>'),
        ("削除",   "del",       '<del rend="strikethrough">{}</del>'),
        ("追加",   "add",       '<add place="above">{}</add>'),
        ("代替",   "subst",     '<subst>\n  <del>{}</del>\n  <add></add>\n</subst>'),
    ],
    "図表・書誌": [
        ("図",     "figure",    '<figure>\n  <graphic url=""/>\n  <figDesc>{}</figDesc>\n</figure>'),
        ("表",     "table",     '<table>\n  <row>\n    <cell>{}</cell>\n  </row>\n</table>'),
        ("書誌",   "bibl",      '<bibl>\n  <author></author>\n  <title>{}</title>\n</bibl>'),
        ("人物",   "person",    '<person xml:id="">\n  <persName>{}</persName>\n</person>'),
        ("場所",   "place",     '<place xml:id="">\n  <placeName>{}</placeName>\n</place>'),
    ],
}

# ===========================================================
# TEI スニペット (メニュー用フルテンプレート)
# ===========================================================
TEI_SNIPPETS = {
    "── ドキュメント ──": None,
    "TEI 新規ドキュメント": """\
<?xml version="1.0" encoding="UTF-8"?>
<TEI xmlns="http://www.tei-c.org/ns/1.0">
  <teiHeader>
    <fileDesc>
      <titleStmt>
        <title>タイトル</title>
      </titleStmt>
      <publicationStmt>
        <p>未公刊</p>
      </publicationStmt>
      <sourceDesc>
        <p>テキストの出所を記述</p>
      </sourceDesc>
    </fileDesc>
  </teiHeader>
  <text>
    <body>
      <div>
        <p>本文をここに入力</p>
      </div>
    </body>
  </text>
</TEI>""",
    "teiHeader (完全版)": """\
<teiHeader>
  <fileDesc>
    <titleStmt>
      <title></title>
      <author></author>
    </titleStmt>
    <editionStmt>
      <edition></edition>
    </editionStmt>
    <publicationStmt>
      <publisher></publisher>
      <date when=""></date>
      <availability>
        <licence target=""></licence>
      </availability>
    </publicationStmt>
    <sourceDesc>
      <bibl>
        <title></title>
        <author></author>
        <date when=""></date>
      </bibl>
    </sourceDesc>
  </fileDesc>
  <encodingDesc>
    <projectDesc><p></p></projectDesc>
  </encodingDesc>
  <profileDesc>
    <langUsage>
      <language ident="ja">日本語</language>
    </langUsage>
  </profileDesc>
</teiHeader>""",
    "── 本文構造 ──": None,
    "div (章節)": "<div>\n  <head></head>\n  <p></p>\n</div>",
    "div type付き": '<div type="chapter" n="">\n  <head></head>\n  <p></p>\n</div>',
    "p (段落)": "<p></p>",
    "head (見出し)": "<head></head>",
    "ab (匿名ブロック)": "<ab></ab>",
    "── 改行・改丁 ──": None,
    "lb (改行)": "<lb/>",
    "pb (改丁)": '<pb n=""/>',
    "cb (改段)": '<cb n=""/>',
    "milestone": '<milestone unit="" n=""/>',
    "── インライン要素 ──": None,
    "hi (強調 bold)": '<hi rend="bold"></hi>',
    "hi (斜体 italic)": '<hi rend="italic"></hi>',
    "note (脚注)": '<note place="foot"></note>',
    "note (インライン)": '<note type="editorial"></note>',
    "ref (参照リンク)": '<ref target=""></ref>',
    "q (引用符付き引用)": "<q></q>",
    "quote (ブロック引用)": "<quote>\n  <p></p>\n</quote>",
    "foreign (外国語)": '<foreign xml:lang=""></foreign>',
    "── 固有名詞 ──": None,
    "persName (人名)": '<persName key=""></persName>',
    "placeName (地名)": '<placeName key=""></placeName>',
    "orgName (組織名)": '<orgName key=""></orgName>',
    "rs (一般参照文字列)": '<rs type="" key=""></rs>',
    "date (日付)": '<date when="YYYY-MM-DD"></date>',
    "title (作品名)": '<title level="m"></title>',
    "── 詩・演劇 ──": None,
    "lg / l (詩行群)": "<lg>\n  <l></l>\n  <l></l>\n</lg>",
    "sp (台詞)": "<sp>\n  <speaker></speaker>\n  <p></p>\n</sp>",
    "stage (舞台指示)": '<stage type="entrance"></stage>',
    "── 表・図 ──": None,
    "table (表)": '<table>\n  <row role="head">\n    <cell></cell>\n    <cell></cell>\n  </row>\n  <row>\n    <cell></cell>\n    <cell></cell>\n  </row>\n</table>',
    "figure (図)": '<figure>\n  <graphic url=""/>\n  <figDesc></figDesc>\n</figure>',
    "── 校異・訂正 ──": None,
    "app / lem / rdg (校異)": '<app>\n  <lem wit="#A"></lem>\n  <rdg wit="#B"></rdg>\n</app>',
    "choice / sic / corr (訂正)": "<choice>\n  <sic></sic>\n  <corr></corr>\n</choice>",
    "choice / orig / reg (正規化)": "<choice>\n  <orig></orig>\n  <reg></reg>\n</choice>",
    "gap (欠損)": '<gap reason="illegible" extent="1" unit="char"/>',
    "unclear (不明瞭)": '<unclear reason=""></unclear>',
    "del / add (削除・追加)": '<del rend="strikethrough"></del>\n<add place="above"></add>',
    "subst (代替)": "<subst>\n  <del></del>\n  <add></add>\n</subst>",
    "── 書誌・リスト ──": None,
    "bibl (書誌情報)": '<bibl>\n  <author></author>\n  <title></title>\n  <date when=""></date>\n</bibl>',
    "listPerson (人物リスト)": '<listPerson>\n  <person xml:id="">\n    <persName></persName>\n    <birth when=""/>\n    <death when=""/>\n  </person>\n</listPerson>',
    "listPlace (地名リスト)": '<listPlace>\n  <place xml:id="">\n    <placeName></placeName>\n    <location><geo></geo></location>\n  </place>\n</listPlace>',
}

# ===========================================================
# ユーティリティ
# ===========================================================

def _build_line_starts(content: str) -> list:
    starts = [0]
    for i, ch in enumerate(content):
        if ch == '\n':
            starts.append(i + 1)
    return starts


def _offset_to_linecol(starts: list, idx: int) -> str:
    line_idx = bisect.bisect_right(starts, idx) - 1
    col = idx - starts[line_idx]
    return f"{line_idx + 1}.{col}"


_RE_COMMENT = re.compile(r'<!--.*?-->', re.DOTALL)
_RE_CDATA   = re.compile(r'<!\[CDATA\[.*?\]\]>', re.DOTALL)
_RE_PI      = re.compile(r'<\?.*?\?>', re.DOTALL)
_RE_DOCTYPE = re.compile(r'<!DOCTYPE\b[^>]*>', re.DOTALL)
_RE_TAG     = re.compile(r'</?[\w:.-]+(?:\s[^<>]*)?\s*/?>', re.DOTALL)
_RE_ATTR    = re.compile(r'\s([\w:.-]+)\s*=\s*("[^"]*"|\'[^\']*\')')
_RE_NS_ATTR = re.compile(r'\s(xmlns(?::[\w.-]+)?)\s*=')


# ===========================================================
# TEI 要素属性定義 (コンテキストパネル用)
#   構造: { "要素名": { "desc": "説明", "attrs": [(名前, ラベル, 種別, [候補値]), ...] } }
#   種別: "text" | "select" | "buttons"
# ===========================================================
TEI_ELEMENT_ATTRS = {
    "TEI": {
        "desc": "TEI ドキュメントのルート要素",
        "attrs": [
            ("xmlns",    "名前空間",   "text",   ["http://www.tei-c.org/ns/1.0"]),
            ("xml:id",   "ID",         "text",   []),
            ("xml:lang", "言語",       "select", ["ja", "en", "zh", "ko", "de", "fr"]),
            ("version",  "バージョン", "text",   ["3.3.0"]),
        ]
    },
    "div": {
        "desc": "テキストの分割単位（章・節など）",
        "attrs": [
            ("type",     "種別", "select", ["chapter", "section", "subsection", "part", "volume", "poem", "play", "letter", "entry"]),
            ("n",        "番号", "text",   []),
            ("xml:id",   "ID",   "text",   []),
            ("xml:lang", "言語", "select", ["ja", "en", "zh", "ko"]),
            ("subtype",  "サブタイプ", "text", []),
        ]
    },
    "p": {
        "desc": "段落",
        "attrs": [
            ("xml:id",   "ID",   "text",    []),
            ("n",        "番号", "text",    []),
            ("rend",     "表示", "buttons", ["indent", "noindent", "center", "right"]),
            ("xml:lang", "言語", "select",  ["ja", "en", "zh", "ko"]),
        ]
    },
    "head": {
        "desc": "見出し",
        "attrs": [
            ("type", "種別", "select",  ["main", "sub", "chapter", "section"]),
            ("n",    "番号", "text",    []),
            ("rend", "表示", "buttons", ["center", "right", "bold"]),
        ]
    },
    "persName": {
        "desc": "人名 ― 人物への参照を含むフレーズ",
        "attrs": [
            ("key",      "人物識別子", "text",   []),
            ("ref",      "参照先URI",  "text",   []),
            ("type",     "種別",       "select", ["forename", "surname", "full", "alias", "pseudonym"]),
            ("role",     "役割",       "text",   []),
            ("xml:lang", "言語",       "select", ["ja", "en", "zh"]),
        ]
    },
    "placeName": {
        "desc": "地名 ― 場所への参照を含むフレーズ",
        "attrs": [
            ("key",      "場所識別子", "text",   []),
            ("ref",      "参照先URI",  "text",   []),
            ("type",     "種別",       "select", ["country", "region", "settlement", "street", "building", "mountain", "river", "sea"]),
            ("xml:lang", "言語",       "select", ["ja", "en", "zh"]),
        ]
    },
    "orgName": {
        "desc": "組織名",
        "attrs": [
            ("key",  "組織識別子", "text",   []),
            ("ref",  "参照先URI",  "text",   []),
            ("type", "種別",       "select", ["government", "military", "religious", "commercial", "academic"]),
        ]
    },
    "rs": {
        "desc": "一般参照文字列",
        "attrs": [
            ("type", "種別",       "text", ["person", "place", "org", "thing", "event"]),
            ("key",  "識別子",     "text", []),
            ("ref",  "参照先URI",  "text", []),
        ]
    },
    "date": {
        "desc": "日付",
        "attrs": [
            ("when",      "日付 (YYYY-MM-DD)", "text",   []),
            ("from",      "開始日",            "text",   []),
            ("to",        "終了日",            "text",   []),
            ("notBefore", "以前でない",         "text",   []),
            ("notAfter",  "以後でない",         "text",   []),
            ("type",      "種別",              "select", ["publication", "creation", "birth", "death", "event"]),
            ("calendar",  "暦法",              "select", ["#Gregorian", "#Julian", "#Japanese", "#Chinese"]),
        ]
    },
    "title": {
        "desc": "作品名・書名",
        "attrs": [
            ("level",    "レベル", "select", ["m", "s", "j", "a", "u"]),
            ("type",     "種別",   "select", ["main", "sub", "series", "alt"]),
            ("key",      "識別子", "text",   []),
            ("ref",      "参照先URI", "text", []),
            ("xml:lang", "言語",   "select", ["ja", "en", "zh"]),
        ]
    },
    "hi": {
        "desc": "強調 ― 特定の書体・表示方法を示す",
        "attrs": [
            ("rend", "書体指定", "buttons", ["bold", "italic", "underline", "strikethrough", "superscript", "subscript", "ruby", "kenten", "large", "small"]),
        ]
    },
    "note": {
        "desc": "注釈・注記",
        "attrs": [
            ("place",  "配置場所", "select", ["foot", "end", "margin", "inline", "above", "below"]),
            ("type",   "種別",     "select", ["editorial", "authorial", "critical", "gloss", "translation"]),
            ("n",      "番号",     "text",   []),
            ("xml:id", "ID",       "text",   []),
            ("resp",   "担当者",   "text",   []),
        ]
    },
    "ref": {
        "desc": "参照リンク",
        "attrs": [
            ("target", "参照先URI", "text",   []),
            ("type",   "種別",      "select", ["internal", "external", "bibl", "person", "place"]),
            ("xml:id", "ID",        "text",   []),
        ]
    },
    "foreign": {
        "desc": "外国語テキスト",
        "attrs": [
            ("xml:lang", "言語コード", "select", ["en", "zh", "ko", "de", "fr", "it", "la", "el"]),
        ]
    },
    "q": {
        "desc": "引用符付き引用",
        "attrs": [
            ("who",  "発話者", "text",   []),
            ("type", "種別",   "select", ["spoken", "written", "thought", "direct", "indirect"]),
        ]
    },
    "quote": {
        "desc": "ブロック引用",
        "attrs": [
            ("source", "出典", "text",   []),
            ("type",   "種別", "select", ["block", "inline"]),
        ]
    },
    "lb": {
        "desc": "改行",
        "attrs": [
            ("n",     "番号",     "text",   []),
            ("break", "改行種別", "select", ["yes", "no", "maybe"]),
            ("type",  "種別",     "text",   []),
        ]
    },
    "pb": {
        "desc": "改丁（ページ）",
        "attrs": [
            ("n",      "丁番号",   "text",   []),
            ("facs",   "画像URI",  "text",   []),
            ("xml:id", "ID",       "text",   []),
            ("type",   "種別",     "select", ["recto", "verso"]),
        ]
    },
    "cb": {
        "desc": "改段（欄）",
        "attrs": [
            ("n",    "段番号", "text", []),
            ("type", "種別",   "text", []),
        ]
    },
    "milestone": {
        "desc": "区切りマーカー",
        "attrs": [
            ("unit", "単位", "select", ["section", "paragraph", "line", "page", "column", "vol"]),
            ("n",    "番号", "text",   []),
            ("type", "種別", "text",   []),
            ("ed",   "版",   "text",   []),
        ]
    },
    "gap": {
        "desc": "欠損・不明部分",
        "attrs": [
            ("reason",   "理由", "select", ["illegible", "damage", "missing", "deleted", "omitted"]),
            ("extent",   "範囲", "text",   []),
            ("unit",     "単位", "select", ["char", "word", "line", "page"]),
            ("quantity", "数量", "text",   []),
            ("resp",     "担当者", "text", []),
        ]
    },
    "space": {
        "desc": "空白",
        "attrs": [
            ("dim",      "方向", "select", ["horizontal", "vertical"]),
            ("extent",   "範囲", "text",   []),
            ("unit",     "単位", "select", ["char", "word", "line"]),
            ("quantity", "数量", "text",   []),
        ]
    },
    "unclear": {
        "desc": "不明瞭な箇所",
        "attrs": [
            ("reason", "理由",  "select", ["illegible", "damage", "ambiguous", "background_noise"]),
            ("cert",   "確信度", "select", ["high", "medium", "low", "unknown"]),
            ("resp",   "担当者", "text",   []),
        ]
    },
    "del": {
        "desc": "削除された文字",
        "attrs": [
            ("rend", "削除方法", "buttons", ["strikethrough", "overwritten", "erased", "overstrike", "underline"]),
            ("type", "種別",     "select",  ["erased", "overwritten"]),
            ("hand", "筆跡",     "text",    []),
        ]
    },
    "add": {
        "desc": "追加された文字",
        "attrs": [
            ("place", "場所", "select", ["above", "below", "margin", "inline", "interlinear"]),
            ("hand",  "筆跡", "text",   []),
            ("type",  "種別", "select", ["insertion", "correction"]),
        ]
    },
    "subst": {
        "desc": "代替（del + add の組み合わせ）",
        "attrs": [
            ("hand", "筆跡", "text", []),
        ]
    },
    "choice": {
        "desc": "選択肢（sic/corr、orig/reg の親）",
        "attrs": []
    },
    "sic": {
        "desc": "原文のまま（誤記）",
        "attrs": [
            ("cert", "確信度", "select", ["high", "medium", "low"]),
        ]
    },
    "corr": {
        "desc": "訂正後",
        "attrs": [
            ("resp", "担当者", "text",   []),
            ("cert", "確信度", "select", ["high", "medium", "low"]),
        ]
    },
    "orig": {"desc": "原形（正規化前）", "attrs": []},
    "reg":  {
        "desc": "正規化形",
        "attrs": [("resp", "担当者", "text", [])],
    },
    "app": {
        "desc": "校異群",
        "attrs": [
            ("type",   "種別", "select", ["substantive", "orthographic", "punctuation"]),
            ("xml:id", "ID",   "text",   []),
        ]
    },
    "lem": {
        "desc": "底本テキスト",
        "attrs": [
            ("wit",    "証拠資料 (#sigla)", "text", []),
            ("source", "出典",              "text", []),
            ("resp",   "担当者",            "text", []),
        ]
    },
    "rdg": {
        "desc": "異読テキスト",
        "attrs": [
            ("wit",  "証拠資料 (#sigla)", "text",   []),
            ("type", "種別",              "select", ["substantive", "orthographic", "transposition", "addition", "omission"]),
        ]
    },
    "l": {
        "desc": "詩行",
        "attrs": [
            ("n",      "行番号", "text", []),
            ("xml:id", "ID",     "text", []),
            ("rhyme",  "韻",     "text", []),
            ("met",    "韻律",   "text", []),
        ]
    },
    "lg": {
        "desc": "詩行群（連・節など）",
        "attrs": [
            ("type",   "種別", "select", ["stanza", "refrain", "couplet", "quatrain", "sestet", "coda"]),
            ("n",      "番号", "text",   []),
            ("xml:id", "ID",   "text",   []),
            ("met",    "韻律", "text",   []),
        ]
    },
    "sp": {
        "desc": "台詞（演劇テキスト）",
        "attrs": [
            ("who",    "発話者 (#id)", "text", []),
            ("xml:id", "ID",           "text", []),
        ]
    },
    "stage": {
        "desc": "舞台指示",
        "attrs": [
            ("type", "種別", "select", ["entrance", "exit", "business", "setting", "modifier"]),
        ]
    },
    "figure": {
        "desc": "図（graphic の親）",
        "attrs": [
            ("type",   "種別", "text", []),
            ("n",      "番号", "text", []),
            ("xml:id", "ID",   "text", []),
        ]
    },
    "graphic": {
        "desc": "画像ファイルへの参照",
        "attrs": [
            ("url",      "画像URL",    "text",   []),
            ("width",    "幅",         "text",   []),
            ("height",   "高さ",       "text",   []),
            ("mimeType", "MIMEタイプ", "select", ["image/jpeg", "image/png", "image/tiff", "image/svg+xml"]),
        ]
    },
    "table": {
        "desc": "表",
        "attrs": [
            ("rows",   "行数", "text", []),
            ("cols",   "列数", "text", []),
            ("xml:id", "ID",   "text", []),
        ]
    },
    "row": {
        "desc": "表の行",
        "attrs": [
            ("role", "役割", "select", ["head", "data", "label", "sum"]),
            ("n",    "番号", "text",   []),
        ]
    },
    "cell": {
        "desc": "表のセル",
        "attrs": [
            ("role", "役割",     "select", ["head", "data", "label", "sum"]),
            ("cols", "結合列数", "text",   []),
            ("rows", "結合行数", "text",   []),
        ]
    },
    "bibl": {
        "desc": "書誌情報",
        "attrs": [
            ("type",   "種別",   "select", ["primary", "secondary", "manuscript", "edition"]),
            ("xml:id", "ID",     "text",   []),
            ("key",    "識別子", "text",   []),
        ]
    },
    "teiHeader":      {"desc": "TEI ヘッダー",       "attrs": [("type", "種別", "select", ["text", "corpus"])]},
    "fileDesc":       {"desc": "ファイル記述",        "attrs": []},
    "titleStmt":      {"desc": "タイトル記述",        "attrs": []},
    "author":         {"desc": "著者",                "attrs": [("xml:id", "ID", "text", []), ("key", "識別子", "text", []), ("ref", "参照先URI", "text", [])]},
    "editor":         {"desc": "編者",                "attrs": [("role", "役割", "select", ["principal", "associate", "technical"]), ("xml:id", "ID", "text", [])]},
    "publicationStmt":{"desc": "公刊記述",            "attrs": []},
    "sourceDesc":     {"desc": "出所記述",            "attrs": []},
    "encodingDesc":   {"desc": "エンコーディング記述","attrs": []},
    "profileDesc":    {"desc": "プロファイル記述",    "attrs": []},
    "text": {
        "desc": "本文",
        "attrs": [
            ("xml:lang", "言語", "select", ["ja", "en", "zh", "ko"]),
            ("type",     "種別", "text",   []),
        ]
    },
    "body": {"desc": "本文の主要部分", "attrs": []},
    "ab": {
        "desc": "匿名ブロック",
        "attrs": [
            ("type",   "種別", "text", []),
            ("xml:id", "ID",   "text", []),
        ]
    },
    "seg": {
        "desc": "テキストセグメント",
        "attrs": [
            ("type",     "種別", "text", []),
            ("xml:id",   "ID",   "text", []),
            ("function", "機能", "text", []),
        ]
    },
}


# ===========================================================
# ツールチップ (ホバー表示)
# ===========================================================
class _Tooltip:
    """ウィジェットにホバーツールチップを付与"""

    def __init__(self, widget: tk.Widget, text: str):
        self._widget = widget
        self._text   = text
        self._tip    = None
        widget.bind("<Enter>", self._show)
        widget.bind("<Leave>", self._hide)

    def _show(self, _=None):
        x = self._widget.winfo_rootx() + 10
        y = self._widget.winfo_rooty() + self._widget.winfo_height() + 4
        self._tip = tk.Toplevel(self._widget)
        self._tip.wm_overrideredirect(True)
        self._tip.wm_geometry(f"+{x}+{y}")
        tk.Label(
            self._tip, text=self._text,
            background="#ffffcc", relief="solid", borderwidth=1,
            font=("", 8), padx=4, pady=2
        ).pack()

    def _hide(self, _=None):
        if self._tip:
            self._tip.destroy()
            self._tip = None


# ===========================================================
# XML シンタックスハイライター
# ===========================================================
class XMLHighlighter:

    MAX_CHARS   = 300_000
    DEBOUNCE_MS = 400

    def __init__(self, text_widget: tk.Text):
        self.w    = text_widget
        self._job = None
        self._configure_tags(dark=False)

    def _configure_tags(self, dark: bool):
        colors = XML_COLORS_DARK if dark else XML_COLORS_LIGHT
        for name, color in colors.items():
            self.w.tag_configure(name, foreground=color)
        for t in ["xml_tag", "xml_endtag", "xml_ns", "xml_attr", "xml_value",
                  "xml_pi", "xml_doctype", "xml_cdata", "xml_comment"]:
            self.w.tag_raise(t)
        self.w.tag_raise("sel")

    def set_theme(self, dark: bool):
        self._configure_tags(dark)

    def schedule(self):
        if self._job:
            self.w.after_cancel(self._job)
        self._job = self.w.after(self.DEBOUNCE_MS, self.highlight)

    def clear(self):
        for tag in XML_COLORS_LIGHT:
            self.w.tag_remove(tag, "1.0", "end")

    def highlight(self):
        self._job = None
        content = self.w.get("1.0", "end-1c")
        if len(content) > self.MAX_CHARS:
            return

        starts = _build_line_starts(content)

        def pos(idx):
            return _offset_to_linecol(starts, idx)

        self.clear()

        protected = []
        for pat, tag in [
            (_RE_COMMENT, "xml_comment"),
            (_RE_CDATA,   "xml_cdata"),
            (_RE_PI,      "xml_pi"),
            (_RE_DOCTYPE, "xml_doctype"),
        ]:
            for m in pat.finditer(content):
                s, e = m.start(), m.end()
                protected.append((s, e))
                self.w.tag_add(tag, pos(s), pos(e))

        def in_protected(s, e):
            return any(ps < e and pe > s for ps, pe in protected)

        for m in _RE_TAG.finditer(content):
            s, e = m.start(), m.end()
            if in_protected(s, e):
                continue
            tag_text = m.group()
            tag_kind = "xml_endtag" if tag_text.startswith("</") else "xml_tag"
            self.w.tag_add(tag_kind, pos(s), pos(e))

            for am in _RE_NS_ATTR.finditer(tag_text):
                self.w.tag_add("xml_ns",
                                pos(s + am.start(1)), pos(s + am.end(1)))
            for am in _RE_ATTR.finditer(tag_text):
                self.w.tag_add("xml_attr",
                                pos(s + am.start(1)), pos(s + am.end(1)))
                val_start = am.start(0) + am.group(0).index('=') + 1
                self.w.tag_add("xml_value",
                                pos(s + val_start), pos(s + am.end(0)))


# ===========================================================
# SAX ハンドラ
# ===========================================================
class _SAXTracker(xml.sax.handler.ContentHandler):

    def __init__(self):
        self.elements = []
        self._stack   = []
        self._loc     = None

    def setDocumentLocator(self, locator):
        self._loc = locator

    def startElement(self, name, attrs):
        line = self._loc.getLineNumber() if self._loc else 0
        self.elements.append((len(self._stack), name, line, dict(attrs)))
        self._stack.append(name)

    def endElement(self, _):
        if self._stack:
            self._stack.pop()


# ===========================================================
# TEI 構造ツリーパネル
# ===========================================================
class TEITreePanel(tk.Frame):

    def __init__(self, master, app):
        super().__init__(master)
        self.app = app
        self._iid_line = {}
        self._build()

    def _build(self):
        hdr = tk.Frame(self, bg="#e4e4e4")
        hdr.pack(fill="x")
        tk.Label(hdr, text="📋 TEI/XML 構造ツリー", bg="#e4e4e4",
                 font=("", 9, "bold"), anchor="w").pack(side="left", padx=6, pady=3)
        tk.Button(hdr, text="↻", command=self.refresh, relief="flat",
                  bg="#e4e4e4", cursor="hand2", padx=6, pady=1,
                  font=("", 9)).pack(side="right", padx=2)

        frm = tk.Frame(self)
        frm.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(frm, show="tree", selectmode="browse")
        vsb = ttk.Scrollbar(frm, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both", expand=True)
        hsb.pack(fill="x")

        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        self._stv = tk.StringVar()
        tk.Label(self, textvariable=self._stv, fg="#555555",
                 font=("", 8), anchor="w").pack(fill="x", padx=4, pady=2)

    def refresh(self, pane=None):
        self.tree.delete(*self.tree.get_children())
        self._iid_line.clear()
        p = pane or self.app.current_pane()
        if not p:
            return
        content = p.text.get("1.0", "end-1c").strip()
        if not content:
            self._stv.set("内容が空です")
            return
        try:
            handler = _SAXTracker()
            xml.sax.parseString(content.encode("utf-8"), handler)
            self._build_tree(handler.elements)
            self._stv.set(f"{len(handler.elements)} 要素")
        except xml.sax.SAXException as e:
            self.tree.insert("", "end", text=f"⚠ XML エラー: {e}")
            self._stv.set("パース失敗")

    def _build_tree(self, elements):
        stack = [("", -1)]
        for depth, tag, line, attrs in elements:
            if '}' in tag:
                tag = tag.split('}', 1)[1]
            while len(stack) > 1 and stack[-1][1] >= depth:
                stack.pop()
            parent_iid = stack[-1][0]

            if "xml:id" in attrs:
                attr_str = f'  @id="{attrs["xml:id"]}"'
            elif "n" in attrs:
                attr_str = f'  @n="{attrs["n"]}"'
            elif attrs:
                k = next(iter(attrs))
                attr_str = f'  @{k}="{attrs[k]}"'
            else:
                attr_str = ""
            label = f"<{tag}>{attr_str}"

            iid = self.tree.insert(parent_iid, "end", text=label)
            self._iid_line[iid] = line
            stack.append((iid, depth))

    def _on_select(self, _=None):
        sel = self.tree.selection()
        if not sel:
            return
        line = self._iid_line.get(sel[0])
        if not line:
            return
        p = self.app.current_pane()
        if not p:
            return
        p.text.mark_set("insert", f"{line}.0")
        p.text.see(f"{line}.0")
        p.text.focus_set()
        self.app._update_statusbar(p)


# ===========================================================
# XPath 検索ダイアログ
# ===========================================================
class XPathDialog(tk.Toplevel):

    def __init__(self, master, app):
        super().__init__(master)
        self.app = app
        self.title("XPath 検索")
        self.geometry("540x420")
        self.attributes("-topmost", True)
        self._build()

    def _build(self):
        tk.Label(self, text="XPath 式:", anchor="w").pack(fill="x", padx=8, pady=(8, 0))
        self.xv = tk.StringVar()
        tk.Entry(self, textvariable=self.xv,
                 font=("Courier New", 10)).pack(fill="x", padx=8, pady=2)

        tk.Label(self, text="名前空間 (prefix=URI, カンマ区切り):",
                 anchor="w").pack(fill="x", padx=8)
        self.nsv = tk.StringVar(value="tei=http://www.tei-c.org/ns/1.0")
        tk.Entry(self, textvariable=self.nsv).pack(fill="x", padx=8, pady=2)

        bf = tk.Frame(self)
        bf.pack(pady=4)
        tk.Button(bf, text="検索", command=self._search, width=10).pack(
            side="left", padx=4)
        tk.Button(bf, text="クリア", command=self._clear, width=10).pack(
            side="left", padx=4)

        rf = tk.Frame(self)
        rf.pack(fill="both", expand=True, padx=8, pady=(0, 4))
        self.result = tk.Text(rf, font=("Courier New", 9),
                              wrap="none", state="disabled")
        vsb = tk.Scrollbar(rf, command=self.result.yview)
        hsb = tk.Scrollbar(self, orient="horizontal", command=self.result.xview)
        self.result.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side="right", fill="y")
        self.result.pack(fill="both", expand=True)
        hsb.pack(fill="x", padx=8)

        self.stv = tk.StringVar()
        tk.Label(self, textvariable=self.stv, fg="navy",
                 anchor="w").pack(fill="x", padx=8, pady=2)

    def _parse_ns(self) -> dict:
        ns = {}
        for part in self.nsv.get().split(','):
            part = part.strip()
            if '=' in part:
                k, v = part.split('=', 1)
                ns[k.strip()] = v.strip()
        return ns

    def _search(self):
        p = self.app.current_pane()
        if not p:
            return
        xpath = self.xv.get().strip()
        if not xpath:
            return
        content = p.text.get("1.0", "end-1c")
        try:
            root = ET.fromstring(content)
            ns = self._parse_ns()
            matches = root.findall(xpath, ns)
            self.result.config(state="normal")
            self.result.delete("1.0", "end")
            for i, elem in enumerate(matches, 1):
                tag = elem.tag
                if '}' in tag:
                    tag = tag.split('}', 1)[1]
                txt = (elem.text or "").strip()[:60]
                attrs = " ".join(
                    f'{k}="{v}"' for k, v in list(elem.attrib.items())[:3])
                self.result.insert(
                    "end",
                    f"{i:4}. <{tag}{' ' + attrs if attrs else ''}>  {txt}\n")
            self.result.config(state="disabled")
            self.stv.set(f"{len(matches)} 件見つかりました")
        except ET.ParseError as e:
            self.stv.set(f"XML パースエラー: {e}")
        except Exception as e:
            self.stv.set(f"エラー: {e}")

    def _clear(self):
        self.result.config(state="normal")
        self.result.delete("1.0", "end")
        self.result.config(state="disabled")
        self.stv.set("")


# ===========================================================
# TEI スニペットバー
#   ツールバー直下の 2 行パネル
#   上段: カテゴリ選択ボタン
#   下段: そのカテゴリのスニペットボタン
#   ・テキスト選択 → ボタン: タグで折り返し (wrap)
#   ・選択なし     → ボタン: テンプレートをカーソルへ挿入
# ===========================================================
class TEISnippetBar(tk.Frame):

    # カテゴリ表示色
    CAT_COLORS = {
        "固有名詞":   "#d4e8ff",
        "マーカー":   "#ffe4d4",
        "本文構造":   "#d4ffd4",
        "改行・丁":   "#f0d4ff",
        "校異・訂正": "#fff4d4",
        "図表・書誌": "#d4fff4",
    }

    def __init__(self, master, app):
        super().__init__(master, relief="groove", bd=1, bg="#eeeeee")
        self.app = app
        self._current_cat = list(TEI_BAR_ACTIONS.keys())[0]
        self._cat_btns    = {}
        self._build()

    # ── UI 構築 ─────────────────────────────────────────────
    def _build(self):
        # ── 上段: カテゴリセレクタ ──────────────────────────
        top = tk.Frame(self, bg="#dddddd")
        top.pack(fill="x")

        tk.Label(top, text="TEI スニペット:", font=("", 8, "bold"),
                 bg="#dddddd", fg="#333333", padx=6).pack(side="left")

        for cat_name in TEI_BAR_ACTIONS:
            color = self.CAT_COLORS.get(cat_name, "#dddddd")
            btn = tk.Button(
                top, text=cat_name, font=("", 8),
                relief="flat", bg="#dddddd", padx=8, pady=1,
                cursor="hand2",
                command=lambda c=cat_name: self._show_cat(c),
            )
            btn.pack(side="left", padx=1, pady=2)
            self._cat_btns[cat_name] = (btn, color)

        # 操作説明
        tk.Label(top,
                 text="← 選択→ボタン: タグ折り返し挿入 / 選択なし→ボタン: テンプレート挿入",
                 font=("", 7), fg="#666666", bg="#dddddd").pack(
            side="right", padx=8)

        # ── 下段: スニペットボタン ──────────────────────────
        self._btn_row = tk.Frame(self, bg="#eeeeee")
        self._btn_row.pack(fill="x", pady=2)

        self._show_cat(self._current_cat)

    # ── カテゴリ切り替え ─────────────────────────────────────
    def _show_cat(self, cat_name: str):
        self._current_cat = cat_name

        # カテゴリボタンの見た目を更新
        for name, (btn, color) in self._cat_btns.items():
            if name == cat_name:
                btn.config(relief="sunken", bg=color, font=("", 8, "bold"))
            else:
                btn.config(relief="flat",   bg="#dddddd", font=("", 8))

        # スニペットボタンを再生成
        for w in self._btn_row.winfo_children():
            w.destroy()

        color = self.CAT_COLORS.get(cat_name, "#eeeeee")
        for label, tag_id, template in TEI_BAR_ACTIONS[cat_name]:
            self._make_snippet_btn(label, tag_id, template, color)

    # ── スニペットボタン 1 個 ─────────────────────────────────
    def _make_snippet_btn(self, label: str, tag_id: str,
                          template: str, bg_color: str):
        btn = tk.Button(
            self._btn_row,
            text=label,
            font=("", 9),
            relief="raised",
            bg=bg_color,
            padx=10, pady=2,
            cursor="hand2",
            command=lambda t=template: self.app._apply_snippet_action(t),
        )
        btn.pack(side="left", padx=3, pady=2)

        # ツールチップ: タグ名を表示
        tooltip_text = f"<{tag_id}>  {template[:50].replace(chr(10),'↵')}"
        _Tooltip(btn, tooltip_text)

        # ホバーでステータスバーに詳細表示
        def on_enter(e, t=template):
            preview = t.replace('\n', '↵')[:80]
            self.app._st_xml.config(text=preview)

        def on_leave(e):
            self.app._update_xml_status()

        btn.bind("<Enter>", on_enter)
        btn.bind("<Leave>", on_leave)


# ===========================================================
# TEI コンテキストパネル (右サイドパネル)
#   カーソル位置の TEI 要素を検出し、属性キー・値をリアルタイム表示・編集
# ===========================================================
class TEIContextPanel(tk.Frame):

    DEBOUNCE_MS = 250
    WIDTH       = 270

    def __init__(self, master, app):
        super().__init__(master, bg="#f5f5f5", relief="groove", bd=1)
        self.app        = app
        self._job       = None
        self._cur_tag   = None   # 現在表示中のタグ名
        self._cur_start = None   # 現在タグの tk index 文字列
        self._build()

    # ── UI 構築 ─────────────────────────────────────────────
    def _build(self):
        # ヘッダー
        hdr = tk.Frame(self, bg="#336699")
        hdr.pack(fill="x")
        tk.Label(hdr, text="🏷 TEI 要素属性", bg="#336699", fg="white",
                 font=("", 9, "bold"), anchor="w").pack(side="left", padx=6, pady=4)
        tk.Button(hdr, text="✕", bg="#336699", fg="white",
                  activebackground="#225588", activeforeground="white",
                  relief="flat", cursor="hand2", font=("", 9),
                  command=lambda: self.app.toggle_context_panel()).pack(
            side="right", padx=4)

        # 要素名・説明
        self._el_frame = tk.Frame(self, bg="#ddeeff", pady=4)
        self._el_frame.pack(fill="x")
        self._el_label = tk.Label(
            self._el_frame, text="（カーソルを置いてください）",
            bg="#ddeeff", fg="#003366", font=("", 10, "bold"),
            wraplength=240, justify="left", anchor="w")
        self._el_label.pack(fill="x", padx=8, pady=(2, 0))
        self._desc_label = tk.Label(
            self._el_frame, text="",
            bg="#ddeeff", fg="#555555", font=("", 8),
            wraplength=240, justify="left", anchor="w")
        self._desc_label.pack(fill="x", padx=8, pady=(0, 2))

        # 属性エリア (Canvas でスクロール可能)
        wrap = tk.Frame(self, bg="#f5f5f5")
        wrap.pack(fill="both", expand=True)
        self._canvas = tk.Canvas(wrap, bg="#f5f5f5", highlightthickness=0)
        self._vsb    = ttk.Scrollbar(wrap, orient="vertical",
                                     command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=self._vsb.set)
        self._vsb.pack(side="right", fill="y")
        self._canvas.pack(side="left", fill="both", expand=True)

        self._attr_frame = tk.Frame(self._canvas, bg="#f5f5f5")
        self._canvas_win = self._canvas.create_window(
            (0, 0), window=self._attr_frame, anchor="nw")
        self._attr_frame.bind("<Configure>", self._on_frame_cfg)
        self._canvas.bind("<Configure>",     self._on_canvas_cfg)
        self._canvas.bind("<MouseWheel>",    self._on_wheel)

    def _on_frame_cfg(self, _=None):
        self._canvas.configure(scrollregion=self._canvas.bbox("all"))

    def _on_canvas_cfg(self, event):
        self._canvas.itemconfig(self._canvas_win, width=event.width)

    def _on_wheel(self, event):
        self._canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    # ── デバウンス更新 ──────────────────────────────────────
    def schedule_update(self, pane=None):
        if self._job:
            self.after_cancel(self._job)
        self._job = self.after(self.DEBOUNCE_MS, lambda: self.update(pane))

    def update(self, pane=None):
        self._job = None
        p = pane or self.app.current_pane()
        if not p or not p.is_xml_mode:
            self._el_label.config(text="（XML/TEI モードではありません）")
            self._desc_label.config(text="F7 で XML ハイライトを有効にしてください")
            self._clear_attrs()
            self._cur_tag   = None
            self._cur_start = None
            return

        tag_name, attrs_dict, tag_idx = self._detect_element(p)
        if not tag_name:
            self._el_label.config(text="（TEI タグが見つかりません）")
            self._desc_label.config(text="カーソルを TEI タグの近くに置いてください")
            self._clear_attrs()
            self._cur_tag   = None
            self._cur_start = None
            return

        # 同じ要素・位置なら属性値ウィジェットを再構築せず維持
        if tag_name == self._cur_tag and tag_idx == self._cur_start:
            return

        self._cur_tag   = tag_name
        self._cur_start = tag_idx

        # ローカル名 (名前空間プレフィックスや {} を除去)
        local = tag_name.split(':')[-1].split('}')[-1]
        info  = TEI_ELEMENT_ATTRS.get(local, {})
        desc  = info.get("desc", "")
        self._el_label.config(text=f"<{local}>  ← 最近傍タグ")
        self._desc_label.config(text=desc if desc else "（TEI_ELEMENT_ATTRS に定義なし）")

        self._clear_attrs()
        self._build_attrs(p, tag_name, tag_idx, attrs_dict, info.get("attrs", []))
        self._canvas.yview_moveto(0)   # スクロールをトップへ

    # ── カーソルに最も近い TEI 要素を検出 ───────────────────────
    def _detect_element(self, pane):
        """
        カーソルに最も近い開始タグ（または自己閉じタグ）を検出する。

        アルゴリズム:
          ① カーソル前後 SEARCH_RADIUS 文字内の全開始/自己閉じタグを列挙し、
            カーソルとの文字距離が最小のタグを選択する。
            ・カーソルがタグ記号 < ... > の内側にある場合は距離 = 0
            ・lb / pb / gap など void 要素もカーソル隣接で検出可能
          ② SEARCH_RADIUS 内に何もなければスタックで囲みタグを検索

        戻り値: (tag_name, attrs_dict, tag_tk_index) または (None, {}, None)
        """
        SEARCH_RADIUS = 400   # カーソル前後を探す文字数

        TAG_RE  = re.compile(r'<(!--.*?--|!\[CDATA\[.*?\]\]>|[^>]*)>', re.DOTALL)
        OPEN_RE = re.compile(r'^([\w:.-]+)((?:\s[^>]*)?)(\s*/)?$')
        END_RE  = re.compile(r'^/([\w:.-]+)')

        try:
            full    = pane.text.get("1.0", "end-1c")
            before  = pane.text.get("1.0", "insert")
            cur_off = len(before)
        except Exception:
            return None, {}, None

        starts_full = _build_line_starts(full)

        def parse_attrs(attr_str):
            return {am.group(1): am.group(2).strip('"\'')
                    for am in _RE_ATTR.finditer(attr_str or "")}

        # ① カーソル前後 SEARCH_RADIUS 内で最も近い開始/自己閉じタグを探す
        s = max(0, cur_off - SEARCH_RADIUS)
        e = min(len(full), cur_off + SEARCH_RADIUS)

        best_name  = None
        best_attrs = {}
        best_idx   = None
        best_dist  = float('inf')

        for m in TAG_RE.finditer(full, s, e):
            inner = m.group(1).strip()
            if not inner or inner.startswith('!') or inner.startswith('?'):
                continue
            if END_RE.match(inner):          # 終了タグはスキップ
                continue
            op_m = OPEN_RE.match(inner)
            if not op_m:
                continue

            ts, te = m.start(), m.end()
            # カーソルがタグ記号内にある場合は距離 0
            dist = (0 if ts <= cur_off <= te
                    else min(abs(cur_off - ts), abs(cur_off - te)))

            if dist < best_dist:
                best_dist  = dist
                best_name  = op_m.group(1)
                best_attrs = parse_attrs(op_m.group(2))
                best_idx   = _offset_to_linecol(starts_full, ts)

        if best_name:
            return best_name, best_attrs, best_idx

        # ② フォールバック: カーソル以前のスタックで囲みタグを検索
        stack      = []
        starts_bef = _build_line_starts(before)

        for m in TAG_RE.finditer(before):
            inner = m.group(1).strip()
            if not inner or inner.startswith('!'):
                continue
            end_m = END_RE.match(inner)
            if end_m:
                name = end_m.group(1)
                for i in range(len(stack) - 1, -1, -1):
                    sn = stack[i][0].split(':')[-1].split('}')[-1]
                    if sn == name or stack[i][0] == name:
                        stack = stack[:i]
                        break
                continue
            op_m = OPEN_RE.match(inner)
            if op_m and not op_m.group(3):   # 開始タグのみ
                stack.append((op_m.group(1), m.start(),
                               parse_attrs(op_m.group(2))))

        if stack:
            tag_name, offset, attrs = stack[-1]
            return tag_name, attrs, _offset_to_linecol(starts_bef, offset)

        return None, {}, None

    # ── 属性 UI の構築 ─────────────────────────────────────────
    def _clear_attrs(self):
        for w in self._attr_frame.winfo_children():
            w.destroy()

    def _build_attrs(self, pane, tag_name, tag_idx,
                     current_attrs: dict, attr_defs: list):
        """
        attr_defs: [(attr_name, label, kind, suggestions), ...]
        kind: "text" | "select" | "buttons"
        """
        if not attr_defs and not current_attrs:
            tk.Label(self._attr_frame, text="（属性定義なし）",
                     bg="#f5f5f5", fg="#888888", font=("", 8)
                     ).pack(padx=6, pady=8)
            return

        defined_names = {a[0] for a in attr_defs}

        for attr_name, label, kind, suggestions in attr_defs:
            cur_val = current_attrs.get(attr_name, "")
            self._make_attr_row(pane, tag_idx, attr_name, label, kind,
                                suggestions, cur_val)

        # 定義外の属性も表示
        extra = {k: v for k, v in current_attrs.items()
                 if k not in defined_names}
        if extra:
            tk.Label(self._attr_frame, text="── その他の属性 ──",
                     bg="#f5f5f5", fg="#888888", font=("", 7)
                     ).pack(fill="x", padx=6, pady=(6, 0))
            for k, v in extra.items():
                self._make_attr_row(pane, tag_idx, k,
                                    f"@{k}", "text", [], v)

    def _make_attr_row(self, pane, tag_idx, attr_name, label,
                       kind, suggestions, cur_val):
        row = tk.Frame(self._attr_frame, bg="#f5f5f5")
        row.pack(fill="x", padx=6, pady=3)

        tk.Label(row, text=label + ":", bg="#f5f5f5", fg="#333333",
                 font=("", 8, "bold"), anchor="w").pack(fill="x")
        tk.Label(row, text=f"  @{attr_name}", bg="#f5f5f5", fg="#7f0055",
                 font=("Courier New", 8), anchor="w").pack(fill="x")

        if kind == "buttons":
            bf = tk.Frame(row, bg="#f5f5f5")
            bf.pack(fill="x", pady=2)
            for sugg in suggestions:
                active = (cur_val == sugg)
                tk.Button(
                    bf, text=sugg, font=("", 7), padx=4, pady=1,
                    relief="sunken" if active else "raised",
                    bg="#336699" if active else "#e8e8e8",
                    fg="white" if active else "#333333",
                    cursor="hand2",
                    command=lambda an=attr_name, sv=sugg, ti=tag_idx, p=pane:
                        self._apply_attr(p, an, sv, ti)
                ).pack(side="left", padx=1, pady=1)
        else:
            var = tk.StringVar(value=cur_val)
            if kind == "select" and suggestions:
                w = ttk.Combobox(row, textvariable=var,
                                 values=suggestions,
                                 font=("Courier New", 9),
                                 width=24, state="normal")
            else:
                w = tk.Entry(row, textvariable=var,
                             font=("Courier New", 9),
                             width=24, bg="white",
                             relief="solid", bd=1)
            w.pack(fill="x", pady=1)
            tk.Button(
                row, text="適用 ✔", font=("", 8), padx=6, pady=1,
                relief="flat", bg="#336699", fg="white", cursor="hand2",
                command=lambda an=attr_name, v=var, ti=tag_idx, p=pane:
                    self._apply_attr(p, an, v.get(), ti)
            ).pack(anchor="e", pady=1)

        # 区切り線
        tk.Frame(self._attr_frame, bg="#cccccc", height=1).pack(
            fill="x", padx=4, pady=2)

    # ── 属性値の適用 ─────────────────────────────────────────
    def _apply_attr(self, pane, attr_name: str,
                    new_value: str, tag_idx: str):
        """
        tag_idx (tk index) に始まる開始タグ内の attr_name 属性値を書き換える。
        属性がなければ新規追加する。
        """
        if not tag_idx:
            return
        try:
            # タグ全体を取得 (最大 500 文字)
            tag_text = pane.text.get(tag_idx, f"{tag_idx} + 500c")
            m = re.match(r'<[^>]+>', tag_text, re.DOTALL)
            if not m:
                return
            old_tag = m.group(0)
            tag_end = f"{tag_idx} + {len(old_tag)}c"

            # 属性を置換または追加
            attr_re = re.compile(
                r'(\s' + re.escape(attr_name) + r'\s*=\s*)("([^"]*)"|\'([^\']*)\')')
            if attr_re.search(old_tag):
                new_tag = attr_re.sub(
                    lambda mo: f'{mo.group(1)}"{new_value}"',
                    old_tag, count=1)
            else:
                # 属性が存在しない → タグ末尾 / > の前に挿入
                if old_tag.endswith('/>'):
                    new_tag = old_tag[:-2] + f' {attr_name}="{new_value}"/>'
                else:
                    new_tag = old_tag[:-1] + f' {attr_name}="{new_value}">'

            if new_tag == old_tag:
                return

            pane.text.delete(tag_idx, tag_end)
            pane.text.insert(tag_idx, new_tag)
            pane.modified = True
            self.app._update_tab_title(pane)
            if pane.is_xml_mode and pane._highlighter:
                pane._highlighter.schedule()
            # パネルを更新してボタン状態を反映
            self._cur_tag   = None   # 強制再描画
            self._cur_start = None
            self.schedule_update(pane)
        except Exception:
            pass   # 編集中など例外は無視


# ===========================================================
# エディタペイン
# ===========================================================
class EditorPane(tk.Frame):

    def __init__(self, master, app, filepath=None):
        super().__init__(master)
        self.app        = app
        self.filepath   = filepath
        self.modified   = False
        self._highlighter = None
        self._xml_mode  = False
        self._build()
        if filepath:
            self._load_file(filepath)
        else:
            self.after(100, self._fill_line_numbers)

    def _build(self):
        self.ruler = tk.Label(
            self, anchor="w", font=("Courier", 10),
            bg=LIGHT["ruler_bg"], fg=LIGHT["ruler_fg"],
            text="".join(
                str((i // 10) % 10) if i % 10 == 0 else "·"
                for i in range(1, 201)),
        )
        self.ruler.pack(fill="x")

        body = tk.Frame(self)
        body.pack(fill="both", expand=True)

        self.line_numbers = tk.Text(
            body, width=5, padx=4, takefocus=0, border=0,
            bg=LIGHT["linenum_bg"], fg=LIGHT["linenum_fg"],
            font=LINE_FONT, state="disabled", cursor="arrow",
            selectbackground=LIGHT["linenum_bg"],
        )
        self.line_numbers.pack(side="left", fill="y")

        self.text = tk.Text(
            body, wrap="none", undo=True,
            bg=LIGHT["text_bg"], fg=LIGHT["text_fg"],
            insertbackground=LIGHT["insert"],
            font=EDITOR_FONT,
            selectbackground="#3390ff", selectforeground="white",
        )
        self.text.pack(side="left", fill="both", expand=True)

        vsb = tk.Scrollbar(body, orient="vertical", command=self._yscroll)
        vsb.pack(side="right", fill="y")
        self.text.config(yscrollcommand=vsb.set)

        hsb = tk.Scrollbar(self, orient="horizontal", command=self.text.xview)
        hsb.pack(side="bottom", fill="x")
        self.text.config(xscrollcommand=hsb.set)

        self.text.bind("<KeyRelease>",      self._on_key)
        self.text.bind("<ButtonRelease-1>", self._on_click)
        self.text.bind("<MouseWheel>",      self._on_wheel)
        self.text.bind("<Configure>",
                       lambda e: self.after_idle(self._fill_line_numbers))
        self.line_numbers.bind("<MouseWheel>",
            lambda e: self.text.event_generate("<MouseWheel>", delta=e.delta))
        self.text.bind(">", self._on_greater_than)

    # ── XML モード ─────────────────────────────────────────────
    @property
    def is_xml_mode(self) -> bool:
        return self._xml_mode

    def enable_xml_mode(self, dark: bool = False):
        if self._highlighter is None:
            self._highlighter = XMLHighlighter(self.text)
        self._highlighter.set_theme(dark)
        self._xml_mode = True
        self._highlighter.schedule()

    def disable_xml_mode(self):
        if self._highlighter:
            self._highlighter.clear()
        self._xml_mode = False

    def _detect_xml(self) -> bool:
        if self.filepath:
            ext = os.path.splitext(self.filepath)[1].lower()
            if ext in XML_EXTENSIONS:
                return True
        first = self.text.get("1.0", "3.0").lstrip()
        return (first.startswith("<?xml") or first.startswith("<TEI") or
                first.startswith("<tei"))

    # ── タグ自動閉じ ───────────────────────────────────────────
    def _on_greater_than(self, event):
        self.text.insert("insert", ">")

        if not self._xml_mode:
            return "break"

        content = self.text.get("1.0", "insert")

        if content.endswith("/>") or re.search(r'</', content[-80:]):
            return "break"

        m = re.search(r'<([\w:.-]+)(?:\s[^<>]*)?>$', content)
        if m:
            tag_name = m.group(1)
            local = tag_name.split(':')[-1].lower()
            if local not in TEI_VOID_ELEMENTS:
                closing = f"</{tag_name}>"
                self.text.insert("insert", closing)
                self.text.mark_set("insert", f"insert - {len(closing)}c")

        return "break"

    # ── ファイル ───────────────────────────────────────────────
    def _load_file(self, path):
        try:
            with open(path, encoding="utf-8", errors="ignore") as f:
                content = f.read()
            self.text.delete("1.0", "end")
            self.text.insert("1.0", content)
            self.filepath = path
            self.modified = False
            self.after(50, self._fill_line_numbers)
            if self._detect_xml():
                self.enable_xml_mode(self.app.dark_mode)
        except Exception as e:
            messagebox.showerror("エラー", f"開けません:\n{e}")

    def save(self):
        if not self.filepath:
            return self.save_as()
        return self._write(self.filepath)

    def save_as(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[
                ("XML/TEI", "*.xml *.tei *.odd"),
                ("テキスト", "*.txt"),
                ("Python",   "*.py"),
                ("すべて",   "*.*"),
            ])
        if path:
            self.filepath = path
            return self._write(path)
        return False

    def _write(self, path):
        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write(self.text.get("1.0", "end-1c"))
            self.modified = False
            self.app._update_tab_title(self)
            return True
        except Exception as e:
            messagebox.showerror("保存エラー", str(e))
            return False

    # ── 行番号 ─────────────────────────────────────────────────
    def _fill_line_numbers(self, _=None):
        total   = int(self.text.index("end-1c").split(".")[0])
        visible = self._visible_lines()
        count   = max(total, visible)
        self.line_numbers.config(state="normal")
        self.line_numbers.delete("1.0", "end")
        self.line_numbers.insert("1.0",
            "\n".join(str(i) for i in range(1, count + 1)))
        self.line_numbers.config(state="disabled")
        self.line_numbers.yview_moveto(self.text.yview()[0])

    def _visible_lines(self):
        h = self.text.winfo_height()
        if h <= 1:
            return 60
        try:
            f  = tkfont.Font(font=self.text.cget("font"))
            lh = f.metrics("linespace")
            return max(1, h // max(lh, 1)) + 3
        except Exception:
            return 60

    def _yscroll(self, *args):
        self.text.yview(*args)
        self.line_numbers.yview(*args)

    def _on_wheel(self, _=None):
        self.after_idle(
            lambda: self.line_numbers.yview_moveto(self.text.yview()[0]))

    def _on_key(self, event=None):
        self._fill_line_numbers()
        self.app._update_statusbar(self)
        if not self.modified:
            self.modified = True
            self.app._update_tab_title(self)
        if self._xml_mode and self._highlighter:
            self._highlighter.schedule()
        char = event.char if event else ""
        if (self.app._recording and char and len(char) == 1
                and char.isprintable() and char != ' '):
            self.app._macro_buffer.append(char)

    def _on_click(self, _=None):
        self.app._update_statusbar(self)
        self.after_idle(
            lambda: self.line_numbers.yview_moveto(self.text.yview()[0]))

    def apply_theme(self, th, dark: bool = False):
        self.text.config(bg=th["text_bg"], fg=th["text_fg"],
                         insertbackground=th["insert"])
        self.line_numbers.config(
            bg=th["linenum_bg"], fg=th["linenum_fg"],
            selectbackground=th["linenum_bg"])
        self.ruler.config(bg=th["ruler_bg"], fg=th["ruler_fg"])
        if self._highlighter:
            self._highlighter.set_theme(dark)
            self._highlighter.schedule()

    def autosave(self):
        content = self.text.get("1.0", "end-1c")
        if not content.strip():
            return None
        os.makedirs(AUTOSAVE_DIR, exist_ok=True)
        base = os.path.splitext(os.path.basename(self.filepath))[0] \
               if self.filepath else "untitled"
        path = os.path.join(
            AUTOSAVE_DIR, f"{base}_{time.strftime('%Y%m%d_%H%M%S')}.bak")
        with open(path, "w", encoding="utf-8") as f:
            f.write(content)
        baks = sorted(
            [x for x in os.listdir(AUTOSAVE_DIR) if x.startswith(base)],
            reverse=True)
        for old in baks[10:]:
            try:
                os.remove(os.path.join(AUTOSAVE_DIR, old))
            except Exception:
                pass
        return path


# ===========================================================
# 検索・置換ダイアログ
# ===========================================================
class SearchReplaceDialog(tk.Toplevel):

    def __init__(self, master, app):
        super().__init__(master)
        self.app = app
        self.title("検索と置換")
        self.resizable(False, False)
        self.attributes("-topmost", True)
        self._ranges  = []
        self._current = -1
        self._build()
        self.protocol("WM_DELETE_WINDOW", self._close)

    def _build(self):
        p = dict(padx=6, pady=3)
        tk.Label(self, text="検索:").grid(row=0, column=0, sticky="e", **p)
        self.sv = tk.StringVar()
        tk.Entry(self, textvariable=self.sv, width=36).grid(
            row=0, column=1, columnspan=3, **p)

        tk.Label(self, text="置換:").grid(row=1, column=0, sticky="e", **p)
        self.rv = tk.StringVar()
        tk.Entry(self, textvariable=self.rv, width=36).grid(
            row=1, column=1, columnspan=3, **p)

        self.re_v   = tk.BooleanVar()
        self.case_v = tk.BooleanVar()
        self.wrap_v = tk.BooleanVar(value=True)
        tk.Checkbutton(self, text="正規表現",     variable=self.re_v
                       ).grid(row=2, column=0, **p)
        tk.Checkbutton(self, text="大文字小文字", variable=self.case_v
                       ).grid(row=2, column=1, **p)
        tk.Checkbutton(self, text="折り返し検索", variable=self.wrap_v
                       ).grid(row=2, column=2, **p)

        bf = tk.Frame(self)
        bf.grid(row=3, column=0, columnspan=4, pady=6)
        for label, cmd in [
            ("次を検索",     self.find_next),
            ("前を検索",     self.find_prev),
            ("置換",         self.replace_one),
            ("すべて置換",   self.replace_all),
            ("全ハイライト", self.highlight_all),
        ]:
            tk.Button(bf, text=label, width=11, command=cmd).pack(
                side="left", padx=2)

        self.stv = tk.StringVar()
        tk.Label(self, textvariable=self.stv, fg="navy").grid(
            row=4, column=0, columnspan=4, padx=6, pady=3)

        self.sv.trace_add("write", lambda *_: self._clear())

    @property
    def _text(self):
        p = self.app.current_pane()
        return p.text if p else None

    def _pattern(self):
        raw = self.sv.get()
        if not raw or not self._text:
            return None
        flags = 0 if self.case_v.get() else re.IGNORECASE
        try:
            return re.compile(
                raw if self.re_v.get() else re.escape(raw), flags)
        except re.error as e:
            messagebox.showerror("正規表現エラー", str(e), parent=self)
            return None

    def _ofs(self, n):
        content = self._text.get("1.0", "end-1c")
        line = content[:n].count("\n") + 1
        col  = n - (content[:n].rfind("\n") + 1)
        return f"{line}.{col}"

    def _all_matches(self):
        pat = self._pattern()
        if not pat or not self._text:
            return []
        content = self._text.get("1.0", "end-1c")
        return [(self._ofs(m.start()), self._ofs(m.end()))
                for m in pat.finditer(content)]

    def highlight_all(self):
        if not self._text:
            return
        self._clear()
        self._ranges = self._all_matches()
        for s, e in self._ranges:
            self._text.tag_add("hl", s, e)
        self._text.tag_config("hl", background="#ffff00", foreground="black")
        self.stv.set(f"{len(self._ranges)} 件見つかりました")

    def _clear(self):
        if self._text:
            self._text.tag_remove("hl",  "1.0", "end")
            self._text.tag_remove("cur", "1.0", "end")
        self._ranges  = []
        self._current = -1

    def find_next(self): self._find(True)
    def find_prev(self): self._find(False)

    def _find(self, forward):
        self._ranges = self._all_matches()
        if not self._ranges:
            self.stv.set("見つかりません")
            return
        cursor = self._text.index("insert")
        if forward:
            for i, (s, _) in enumerate(self._ranges):
                if self._text.compare(s, ">=", cursor):
                    self._current = i
                    break
            else:
                if self.wrap_v.get():
                    self._current = 0
                else:
                    self.stv.set("これ以上ありません")
                    return
        else:
            found = False
            for i in range(len(self._ranges) - 1, -1, -1):
                if self._text.compare(self._ranges[i][0], "<", cursor):
                    self._current = i
                    found = True
                    break
            if not found:
                if self.wrap_v.get():
                    self._current = len(self._ranges) - 1
                else:
                    self.stv.set("これ以上ありません")
                    return
        self._jump(self._current)

    def _jump(self, idx):
        self._text.tag_remove("hl",  "1.0", "end")
        self._text.tag_remove("cur", "1.0", "end")
        for i, (s, e) in enumerate(self._ranges):
            self._text.tag_add("cur" if i == idx else "hl", s, e)
        self._text.tag_config("hl",  background="#ffff00", foreground="black")
        self._text.tag_config("cur", background="#ff6600", foreground="white")
        s, _ = self._ranges[idx]
        self._text.mark_set("insert", s)
        self._text.see(s)
        self.stv.set(f"{idx + 1} / {len(self._ranges)} 件")

    def replace_one(self):
        self._ranges = self._all_matches()
        if not self._ranges:
            self.stv.set("見つかりません")
            return
        if self._current < 0 or self._current >= len(self._ranges):
            self._find(True)
            return
        s, e = self._ranges[self._current]
        self._text.delete(s, e)
        self._text.insert(s, self.rv.get())
        self._ranges = self._all_matches()
        self._find(True)

    def replace_all(self):
        matches = self._all_matches()
        if not matches:
            self.stv.set("見つかりません")
            return
        for s, e in reversed(matches):
            self._text.delete(s, e)
            self._text.insert(s, self.rv.get())
        self.stv.set(f"{len(matches)} 件置換しました")
        self._clear()

    def _close(self):
        self._clear()
        self.destroy()


# ===========================================================
# マクロ管理ダイアログ
# ===========================================================
class MacroDialog(tk.Toplevel):

    def __init__(self, master, app):
        super().__init__(master)
        self.app = app
        self.title("マクロ管理")
        self.geometry("580x440")
        self.attributes("-topmost", True)
        self.macros = self._load()
        self._build()

    def _load(self):
        if os.path.exists(MACRO_FILE):
            try:
                with open(MACRO_FILE, encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
        return {}

    def _save_disk(self):
        with open(MACRO_FILE, "w", encoding="utf-8") as f:
            json.dump(self.macros, f, ensure_ascii=False, indent=2)

    def _build(self):
        lf = tk.Frame(self)
        lf.pack(side="left", fill="both", padx=6, pady=6)
        tk.Label(lf, text="マクロ一覧").pack()
        self.lb = tk.Listbox(lf, width=22)
        self.lb.pack(fill="both", expand=True)
        self.lb.bind("<<ListboxSelect>>", self._select)
        self._refresh()

        ef = tk.Frame(self)
        ef.pack(side="left", fill="both", expand=True, padx=6, pady=6)
        tk.Label(ef, text="名前:").pack(anchor="w")
        self.nv = tk.StringVar()
        tk.Entry(ef, textvariable=self.nv).pack(fill="x")
        tk.Label(ef, text="スクリプト (Python):  変数: editor / text / re").pack(
            anchor="w", pady=(6, 0))
        self.sc = tk.Text(ef, font=("Courier New", 10), height=13)
        self.sc.pack(fill="both", expand=True)
        self.sc.insert("1.0",
            "# 例: 選択テキストを大文字に変換\n"
            "try:\n"
            "    sel = editor.get('sel.first', 'sel.last')\n"
            "    editor.delete('sel.first', 'sel.last')\n"
            "    editor.insert('insert', sel.upper())\n"
            "except Exception:\n"
            "    pass\n")
        bf = tk.Frame(ef)
        bf.pack(pady=4)
        for label, cmd in [
            ("保存", self._save), ("実行", self._run),
            ("削除", self._delete), ("新規", self._new),
        ]:
            tk.Button(bf, text=label, width=8, command=cmd).pack(
                side="left", padx=2)
        self.sv = tk.StringVar()
        tk.Label(ef, textvariable=self.sv, fg="navy", wraplength=320).pack()

    def _refresh(self):
        self.lb.delete(0, "end")
        for n in self.macros:
            self.lb.insert("end", n)

    def _select(self, _=None):
        s = self.lb.curselection()
        if not s:
            return
        n = self.lb.get(s[0])
        self.nv.set(n)
        self.sc.delete("1.0", "end")
        self.sc.insert("1.0", self.macros[n])

    def _save(self):
        n = self.nv.get().strip()
        if not n:
            self.sv.set("名前を入力してください")
            return
        self.macros[n] = self.sc.get("1.0", "end-1c")
        self._save_disk()
        self._refresh()
        self.sv.set(f"「{n}」保存済み")
        self.app._load_macros()

    def _run(self):
        p = self.app.current_pane()
        if not p:
            return
        script = self.sc.get("1.0", "end-1c")
        try:
            exec(script, {"editor": p.text,  # noqa: S102
                          "text": p.text.get("1.0", "end-1c"),
                          "re": re, "tk": tk})
            self.sv.set("実行完了")
        except Exception as e:
            self.sv.set(f"エラー: {e}")

    def _delete(self):
        n = self.nv.get().strip()
        if n in self.macros:
            del self.macros[n]
            self._save_disk()
            self._refresh()
            self.nv.set("")
            self.sc.delete("1.0", "end")
            self.sv.set(f"「{n}」削除")
            self.app._load_macros()

    def _new(self):
        self.nv.set("")
        self.sc.delete("1.0", "end")
        self.sc.insert("1.0", "# 新しいマクロをここに記述\n")


# ===========================================================
# メインアプリケーション
# ===========================================================
_BaseApp = TkinterDnD.Tk if HAS_DND else tk.Tk


class TentekeEditor(_BaseApp):

    FRAME_COLORS = ["#ccffff", "#ffccff", "#ffffcc", "#ccffcc"]

    def __init__(self):
        super().__init__()
        self.title("Tenteke Editor  [TEI/XML 対応]")
        self.geometry("1380x900")
        self.minsize(900, 560)

        self.dark_mode          = False
        self._frame_idx         = 0
        self._autosave_stop     = threading.Event()
        self.autosave_enabled   = tk.BooleanVar(value=True)
        self.autosave_interval  = tk.IntVar(value=60)
        self._recording         = False
        self._macro_buffer      = []
        self._search_dlg        = None
        self._macro_dlg         = None
        self._xpath_dlg         = None
        self._tree_visible      = False
        self._snip_bar_visible  = True   # スニペットバーは初期表示
        self._ctx_panel_visible = True   # コンテキストパネルは初期表示

        os.makedirs(AUTOSAVE_DIR, exist_ok=True)

        self._build_ui()
        self._load_macros()
        self._start_autosave()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.new_tab()

    # ── UI 構築 ────────────────────────────────────────────────
    def _build_ui(self):
        self.outer = tk.Frame(self, bg=self.FRAME_COLORS[0], bd=8)
        self.outer.pack(fill="both", expand=True)

        self._create_menu()
        self._create_toolbar()

        # TEI スニペットバー (ツールバー直下・初期表示)
        self._snippet_bar = TEISnippetBar(self.outer, self)
        self._snippet_bar.pack(fill="x")

        # メインエリア: ツリーパネル + ノートブック
        self._main_frame = tk.Frame(self.outer)
        self._main_frame.pack(fill="both", expand=True)

        self._tree_panel = TEITreePanel(self._main_frame, self)

        self.notebook = ttk.Notebook(self._main_frame)
        self.notebook.pack(side="left", fill="both", expand=True)
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)
        self.notebook.bind("<Button-2>", self._close_tab_middle)

        # TEI コンテキストパネル (右サイド、初期表示)
        self._ctx_panel = TEIContextPanel(self._main_frame, self)
        self._ctx_panel.pack(side="right", fill="y")
        self._ctx_panel.configure(width=TEIContextPanel.WIDTH)
        self._ctx_panel.pack_propagate(False)

        self._create_statusbar()

        if HAS_DND:
            self.drop_target_register(DND_FILES)
            self.dnd_bind("<<Drop>>", self._on_drop)

    def _create_menu(self):
        mb = tk.Menu(self)

        # ── ファイル ──
        fm = tk.Menu(mb, tearoff=0)
        fm.add_command(label="新規タブ\tCtrl+T",        command=self.new_tab)
        fm.add_command(label="新規 TEI ドキュメント",   command=self.new_tei_document)
        fm.add_separator()
        fm.add_command(label="開く\tCtrl+O",             command=self.open_file)
        fm.add_command(label="保存\tCtrl+S",             command=self.save_file)
        fm.add_command(label="名前を付けて保存",         command=self.save_file_as)
        fm.add_command(label="タブを閉じる\tCtrl+W",    command=self.close_tab)
        fm.add_separator()
        fm.add_command(label="自動保存設定...",          command=self._autosave_settings)
        fm.add_separator()
        fm.add_command(label="終了",                     command=self._on_close)
        mb.add_cascade(label="ファイル", menu=fm)

        # ── 編集 ──
        em = tk.Menu(mb, tearoff=0)
        em.add_command(label="元に戻す\tCtrl+Z",
            command=lambda: self.current_pane() and
                            self.current_pane().text.edit_undo())
        em.add_command(label="やり直す\tCtrl+Y",
            command=lambda: self.current_pane() and
                            self.current_pane().text.edit_redo())
        em.add_separator()
        for lbl, ev in [("切り取り", "<<Cut>>"), ("コピー", "<<Copy>>"),
                        ("貼り付け", "<<Paste>>")]:
            em.add_command(label=lbl,
                command=lambda e=ev: self.current_pane() and
                                     self.current_pane().text.event_generate(e))
        em.add_separator()
        em.add_command(label="すべて選択\tCtrl+A",
            command=lambda: self.current_pane() and
                            self.current_pane().text.tag_add(
                                "sel", "1.0", "end"))
        mb.add_cascade(label="編集", menu=em)

        # ── 検索 ──
        sm = tk.Menu(mb, tearoff=0)
        sm.add_command(label="検索・置換\tCtrl+H", command=self.open_search)
        sm.add_command(label="行ジャンプ\tCtrl+G", command=self.goto_line)
        mb.add_cascade(label="検索", menu=sm)

        # ── TEI/XML ──
        xm = tk.Menu(mb, tearoff=0)
        xm.add_command(label="スニペットバー 表示/非表示\tF10",
                        command=self.toggle_snippet_bar)
        xm.add_command(label="属性パネル 表示/非表示\tF11",
                        command=self.toggle_context_panel)
        xm.add_separator()
        xm.add_command(label="XML ハイライト ON/OFF\tF7",
                        command=self.toggle_xml_mode)
        xm.add_command(label="XML バリデーション\tF8",
                        command=self.validate_xml)
        xm.add_command(label="XML 整形 (インデント)\tCtrl+Shift+F",
                        command=self.format_xml)
        xm.add_separator()
        xm.add_command(label="構造ツリー 表示/非表示\tF6",
                        command=self.toggle_tree_panel)
        xm.add_command(label="XPath 検索\tCtrl+Shift+X",
                        command=self.open_xpath_dialog)
        xm.add_separator()

        # スニペットサブメニュー (メニューからも挿入可)
        snip_menu = tk.Menu(xm, tearoff=0)
        for name, snippet in TEI_SNIPPETS.items():
            if snippet is None:
                snip_menu.add_separator()
                snip_menu.add_command(label=name, state="disabled")
            else:
                snip_menu.add_command(
                    label=name,
                    command=lambda s=snippet: self.insert_snippet(s))
        xm.add_cascade(label="TEI スニペット挿入 (フルテンプレート)", menu=snip_menu)
        mb.add_cascade(label="TEI/XML", menu=xm)

        # ── マクロ ──
        self.macro_menu = tk.Menu(mb, tearoff=0)
        self.macro_menu.add_command(label="マクロ管理...",
                                     command=self.open_macro_dialog)
        self.macro_menu.add_command(label="記録開始/停止\tF9",
                                     command=self.toggle_record)
        self.macro_menu.add_separator()
        mb.add_cascade(label="マクロ", menu=self.macro_menu)

        # ── 表示 ──
        vm = tk.Menu(mb, tearoff=0)
        vm.add_command(label="ダークモード\tF5",   command=self.toggle_dark)
        vm.add_command(label="外枠色切替",          command=self.toggle_frame)
        vm.add_separator()
        vm.add_command(label="フォントサイズ変更",  command=self.change_font)
        mb.add_cascade(label="表示", menu=vm)

        self.config(menu=mb)

        # キーバインド
        self.bind_all("<Control-t>", lambda e: self.new_tab())
        self.bind_all("<Control-o>", lambda e: self.open_file())
        self.bind_all("<Control-s>", lambda e: self.save_file())
        self.bind_all("<Control-w>", lambda e: self.close_tab())
        self.bind_all("<Control-h>", lambda e: self.open_search())
        self.bind_all("<Control-g>", lambda e: self.goto_line())
        self.bind_all("<F5>",        lambda e: self.toggle_dark())
        self.bind_all("<F6>",        lambda e: self.toggle_tree_panel())
        self.bind_all("<F7>",        lambda e: self.toggle_xml_mode())
        self.bind_all("<F8>",        lambda e: self.validate_xml())
        self.bind_all("<F9>",        lambda e: self.toggle_record())
        self.bind_all("<F10>",       lambda e: self.toggle_snippet_bar())
        self.bind_all("<F11>",       lambda e: self.toggle_context_panel())
        self.bind_all("<Control-F>", lambda e: self.format_xml())
        self.bind_all("<Control-X>", lambda e: self.open_xpath_dialog())

    def _create_toolbar(self):
        tb = tk.Frame(self.outer, bd=1, relief="raised")
        tb.pack(fill="x")

        def B(text, cmd):
            b = tk.Button(tb, text=text, command=cmd,
                          relief="flat", padx=6, pady=2, cursor="hand2")
            b.pack(side="left", padx=1, pady=2)
            return b

        def S():
            tk.Frame(tb, width=1, bg="#aaaaaa").pack(
                side="left", fill="y", padx=5, pady=3)

        B("📄 新規",       self.new_tab)
        B("📂 開く",       self.open_file)
        B("💾 保存",       self.save_file)
        S()
        B("🔍 検索・置換", self.open_search)
        B("⤵ 行ジャンプ", self.goto_line)
        S()
        B("🌲 ツリー",    self.toggle_tree_panel)
        B("✔ XML検証",   self.validate_xml)
        B("⇌ 整形",      self.format_xml)
        B("🔎 XPath",     self.open_xpath_dialog)
        B("✦ TEI新規",   self.new_tei_document)
        S()
        self.snip_btn = B("📌 スニペットバー", self.toggle_snippet_bar)
        self.ctx_btn  = B("🏷 属性パネル",     self.toggle_context_panel)
        S()
        self.rec_btn = B("⏺ 記録", self.toggle_record)

        dnd_txt = "  ← ドロップでファイルを開く" if HAS_DND else \
                  "  (pip install tkinterdnd2 でD&D対応)"
        tk.Label(tb, text=dnd_txt, fg="#999999", font=("", 9)).pack(side="left")

        self.autosave_lamp = tk.Label(
            tb, text="自動保存: ON", fg="#009900", font=("", 9), padx=8)
        self.autosave_lamp.pack(side="right")

    def _create_statusbar(self):
        bar = tk.Frame(self.outer, bd=1, relief="sunken")
        bar.pack(fill="x", side="bottom")

        self._st_cursor = tk.Label(bar, text="行: 1  列: 1", anchor="w", padx=8)
        self._st_cursor.pack(side="left")
        tk.Frame(bar, width=1, bg="#aaaaaa").pack(side="left", fill="y", pady=2)
        self._st_chars = tk.Label(bar, text="文字数: 0", anchor="w", padx=8)
        self._st_chars.pack(side="left")
        tk.Frame(bar, width=1, bg="#aaaaaa").pack(side="left", fill="y", pady=2)
        self._st_xml = tk.Label(bar, text="", fg="#0055aa", anchor="w", padx=8,
                                font=("", 8))
        self._st_xml.pack(side="left", fill="x", expand=True)
        tk.Frame(bar, width=1, bg="#aaaaaa").pack(side="left", fill="y", pady=2)
        self._st_autosave = tk.Label(bar, text="", fg="#888888", anchor="w", padx=8)
        self._st_autosave.pack(side="left")
        self._st_rec = tk.Label(bar, text="", fg="red", anchor="e", padx=8)
        self._st_rec.pack(side="right")

    # ── タブ管理 ───────────────────────────────────────────────
    def new_tab(self, filepath=None):
        pane  = EditorPane(self.notebook, self, filepath)
        title = os.path.basename(filepath) if filepath else "無題"
        self.notebook.add(pane, text=f"  {title}  ")
        self.notebook.select(pane)
        pane.text.focus_set()
        return pane

    def current_pane(self):
        try:
            w = self.notebook.nametowidget(self.notebook.select())
            return w if isinstance(w, EditorPane) else None
        except Exception:
            return None

    def _update_tab_title(self, pane):
        try:
            idx  = self.notebook.index(pane)
            base = os.path.basename(pane.filepath) if pane.filepath else "無題"
            mark = " ●" if pane.modified else ""
            self.notebook.tab(idx, text=f"  {base}{mark}  ")
        except Exception:
            pass

    def close_tab(self, pane=None):
        p = pane or self.current_pane()
        if not p:
            return
        if p.modified:
            name = os.path.basename(p.filepath) if p.filepath else "無題"
            ans  = messagebox.askyesnocancel(
                "保存確認", f"「{name}」は変更されています。保存しますか？",
                parent=self)
            if ans is None:
                return
            if ans:
                p.save()
        try:
            self.notebook.forget(p)
        except Exception:
            pass
        if self.notebook.index("end") == 0:
            self.new_tab()

    def _close_tab_middle(self, event):
        try:
            idx  = self.notebook.index(f"@{event.x},{event.y}")
            pane = self.notebook.nametowidget(self.notebook.tabs()[idx])
            self.close_tab(pane)
        except Exception:
            pass

    def _on_tab_changed(self, _=None):
        p = self.current_pane()
        if p:
            p.text.focus_set()
            self._update_statusbar(p)
            self._update_xml_status(p)
            if self._tree_visible:
                self._tree_panel.refresh(p)
            if self._ctx_panel_visible:
                # タブ切り替え時はパネルを即時リセットしてから更新
                self._ctx_panel._cur_tag   = None
                self._ctx_panel._cur_start = None
                self._ctx_panel.schedule_update(p)

    # ── ファイル操作 ───────────────────────────────────────────
    def open_file(self):
        paths = filedialog.askopenfilenames(
            filetypes=[
                ("XML/TEI", "*.xml *.tei *.odd *.rng *.xsl *.xslt"),
                ("テキスト", "*.txt"),
                ("Python",   "*.py"),
                ("すべて",   "*.*"),
            ])
        for path in paths:
            self._open_path(path)

    def _open_path(self, path):
        path = path.strip().strip("{}")
        if not os.path.isfile(path):
            return
        for tab in self.notebook.tabs():
            pane = self.notebook.nametowidget(tab)
            if isinstance(pane, EditorPane) and pane.filepath == path:
                self.notebook.select(pane)
                return
        p = self.current_pane()
        if p and not p.filepath and not p.modified and \
                p.text.get("1.0", "end-1c") == "":
            p._load_file(path)
            self._update_tab_title(p)
            self._update_xml_status(p)
        else:
            pane = self.new_tab(path)
            self._update_xml_status(pane)

    def save_file(self):
        p = self.current_pane()
        if p:
            p.save()

    def save_file_as(self):
        p = self.current_pane()
        if p:
            p.save_as()
            self._update_tab_title(p)

    def _on_drop(self, event):
        raw   = event.data
        paths = re.findall(r'\{([^}]+)\}|(\S+)', raw)
        for a, b in paths:
            path = (a or b).strip()
            if path:
                self.after(0, self._open_path, path)

    # ── ステータスバー ─────────────────────────────────────────
    def _update_statusbar(self, pane=None):
        p = pane or self.current_pane()
        if not p:
            return
        row, col = p.text.index("insert").split(".")
        self._st_cursor.config(text=f"行: {row}  列: {int(col) + 1}")
        self._st_chars.config(
            text=f"文字数: {len(p.text.get('1.0', 'end-1c'))}")
        # コンテキストパネルをデバウンス更新
        if self._ctx_panel_visible:
            self._ctx_panel.schedule_update(p)

    def _update_xml_status(self, pane=None):
        p = pane or self.current_pane()
        if not p:
            self._st_xml.config(text="")
            return
        if p.is_xml_mode:
            self._st_xml.config(text="XML/TEI ✓  │  ボタンにマウスを重ねるとテンプレートを表示",
                                fg="#0055aa")
        else:
            self._st_xml.config(text="")

    # ── 検索・置換 ─────────────────────────────────────────────
    def open_search(self):
        if self._search_dlg and self._search_dlg.winfo_exists():
            self._search_dlg.lift()
        else:
            self._search_dlg = SearchReplaceDialog(self, self)

    def goto_line(self):
        p = self.current_pane()
        if not p:
            return
        total = int(p.text.index("end-1c").split(".")[0])
        line  = simpledialog.askinteger(
            "行ジャンプ", f"行番号 (1〜{total}):",
            parent=self, minvalue=1, maxvalue=total)
        if line:
            p.text.mark_set("insert", f"{line}.0")
            p.text.see(f"{line}.0")
            self._update_statusbar(p)

    # ── スニペットバー表示/非表示 ────────────────────────────────
    def toggle_snippet_bar(self):
        """TEI スニペットバーを表示/非表示 (F10)"""
        if self._snip_bar_visible:
            self._snippet_bar.pack_forget()
            self._snip_bar_visible = False
            self.snip_btn.config(relief="flat")
        else:
            # ツールバーの次, _main_frame の前に挿入
            self._snippet_bar.pack(
                fill="x", before=self._main_frame)
            self._snip_bar_visible = True
            self.snip_btn.config(relief="sunken")

    # ── コンテキストパネル表示/非表示 ───────────────────────────
    def toggle_context_panel(self):
        """TEI 属性コンテキストパネルを表示/非表示 (F11)"""
        if self._ctx_panel_visible:
            self._ctx_panel.pack_forget()
            self._ctx_panel_visible = False
            self.ctx_btn.config(relief="flat")
        else:
            self._ctx_panel.pack(side="right", fill="y",
                                 after=self.notebook)
            self._ctx_panel.configure(width=TEIContextPanel.WIDTH)
            self._ctx_panel.pack_propagate(False)
            self._ctx_panel_visible = True
            self.ctx_btn.config(relief="sunken")
            # 直ちに更新
            p = self.current_pane()
            if p:
                self._ctx_panel.schedule_update(p)

    # ── スニペット適用 (コアロジック) ─────────────────────────────
    def _apply_snippet_action(self, template: str):
        """
        テキスト選択中: タグで選択テキストを折り返し挿入
        選択なし      : テンプレートをカーソル位置に挿入
        カーソル位置  : {} があればその位置、なければタグ末尾
        """
        p = self.current_pane()
        if not p:
            return

        # 選択テキストを取得
        try:
            sel_text  = p.text.get("sel.first", "sel.last")
            sel_start = p.text.index("sel.first")
            sel_end   = p.text.index("sel.last")
            has_sel   = True
        except tk.TclError:
            sel_text  = ""
            sel_start = p.text.index("insert")
            sel_end   = None
            has_sel   = False

        # {} を選択テキストで置換
        if '{}' in template:
            result       = template.replace('{}', sel_text, 1)
            cursor_chars = template.index('{}') + len(sel_text)
        else:
            # void 要素など {} なし: そのまま挿入
            result       = template
            cursor_chars = len(result)

        # 挿入
        if has_sel:
            p.text.delete(sel_start, sel_end)
        p.text.insert(sel_start, result)

        # カーソルを適切な位置へ
        # tkinter の "+Nc" は文字数ではなく index 演算なので変換
        try:
            new_idx = f"{sel_start} + {cursor_chars}c"
            p.text.mark_set("insert", new_idx)
            p.text.see("insert")
        except Exception:
            pass

        p.modified = True
        self._update_tab_title(p)
        if p.is_xml_mode and p._highlighter:
            p._highlighter.schedule()
        p.text.focus_set()
        self._update_statusbar(p)

    # ── TEI/XML 機能 ───────────────────────────────────────────
    def toggle_xml_mode(self):
        p = self.current_pane()
        if not p:
            return
        if p.is_xml_mode:
            p.disable_xml_mode()
            self._st_xml.config(text="")
        else:
            p.enable_xml_mode(self.dark_mode)
            self._update_xml_status(p)

    def validate_xml(self):
        p = self.current_pane()
        if not p:
            return
        content = p.text.get("1.0", "end-1c").strip()
        if not content:
            messagebox.showinfo("バリデーション", "内容が空です")
            return
        p.text.tag_remove("xml_error_line", "1.0", "end")
        try:
            ET.fromstring(content)
            messagebox.showinfo(
                "バリデーション ✔",
                "XML は整形式です (Well-formed)。\n\nエラーなし。\n\n"
                "※ TEI スキーマ検証 (ODD/RNG) には\n"
                "  oXygen XML Editor や Jing が必要です。")
        except ET.ParseError as e:
            msg = str(e)
            m = re.search(r'line (\d+)', msg)
            if m:
                line = int(m.group(1))
                p.text.tag_add("xml_error_line",
                                f"{line}.0", f"{line}.end")
                p.text.tag_config("xml_error_line",
                    background="#ffdddd", foreground="#cc0000")
                p.text.see(f"{line}.0")
            messagebox.showerror(
                "バリデーション エラー",
                f"XML エラー:\n{e}\n\n"
                "エラー行を赤色でハイライトしました。")

    def format_xml(self):
        p = self.current_pane()
        if not p:
            return
        content = p.text.get("1.0", "end-1c").strip()
        if not content:
            return
        try:
            decl = ""
            if content.startswith("<?xml"):
                end_pos = content.index("?>") + 2
                decl    = content[:end_pos] + "\n"
                content = content[end_pos:].lstrip()

            root = ET.fromstring(content)
            self._indent_element(root, level=0)
            formatted = ET.tostring(root, encoding="unicode",
                                    xml_declaration=False)
            result = decl + formatted

            p.text.delete("1.0", "end")
            p.text.insert("1.0", result)
            p.modified = True
            self._update_tab_title(p)
            if p.is_xml_mode and p._highlighter:
                p._highlighter.schedule()
            self._update_statusbar(p)
        except ET.ParseError as e:
            messagebox.showerror("整形エラー",
                f"XML のパースに失敗しました:\n{e}")

    def _indent_element(self, elem, level: int):
        indent = "\n" + "  " * level
        if len(elem):
            if not elem.text or not elem.text.strip():
                elem.text = indent + "  "
            if not elem.tail or not elem.tail.strip():
                elem.tail = indent
            last = None
            for child in elem:
                self._indent_element(child, level + 1)
                last = child
            if last is not None and (not last.tail or not last.tail.strip()):
                last.tail = indent
        else:
            if level and (not elem.tail or not elem.tail.strip()):
                elem.tail = indent

    def new_tei_document(self):
        pane = self.new_tab()
        pane.text.insert("1.0", TEI_SNIPPETS["TEI 新規ドキュメント"])
        pane.modified = True
        pane.enable_xml_mode(self.dark_mode)
        self._update_xml_status(pane)
        self._update_tab_title(pane)

    def insert_snippet(self, snippet: str):
        """メニューからのフルテンプレート挿入"""
        p = self.current_pane()
        if not p:
            return
        try:
            sel_start = p.text.index("sel.first")
            sel_end   = p.text.index("sel.last")
            p.text.delete(sel_start, sel_end)
            p.text.insert(sel_start, snippet)
        except tk.TclError:
            p.text.insert("insert", snippet)
        p.modified = True
        self._update_tab_title(p)
        if p.is_xml_mode and p._highlighter:
            p._highlighter.schedule()

    def toggle_tree_panel(self):
        if self._tree_visible:
            self._tree_panel.pack_forget()
            self._tree_visible = False
        else:
            self._tree_panel.pack(
                side="left", fill="y", before=self.notebook)
            self._tree_panel.configure(width=250)
            self._tree_panel.pack_propagate(False)
            self._tree_visible = True
            self._tree_panel.refresh()

    def open_xpath_dialog(self):
        if self._xpath_dlg and self._xpath_dlg.winfo_exists():
            self._xpath_dlg.lift()
        else:
            self._xpath_dlg = XPathDialog(self, self)

    # ── マクロ ─────────────────────────────────────────────────
    def open_macro_dialog(self):
        if self._macro_dlg and self._macro_dlg.winfo_exists():
            self._macro_dlg.lift()
        else:
            self._macro_dlg = MacroDialog(self, self)

    def _load_macros(self):
        macros = {}
        if os.path.exists(MACRO_FILE):
            try:
                with open(MACRO_FILE, encoding="utf-8") as f:
                    macros = json.load(f)
            except Exception:
                pass
        end = self.macro_menu.index("end")
        if end is not None:
            for i in range(int(end), 2, -1):
                try:
                    self.macro_menu.delete(i)
                except Exception:
                    pass
        for name, script in macros.items():
            self.macro_menu.add_command(
                label=name,
                command=lambda s=script: self._exec_macro(s))

    def _exec_macro(self, script):
        p = self.current_pane()
        if not p:
            return
        try:
            exec(script, {"editor": p.text,  # noqa: S102
                          "text": p.text.get("1.0", "end-1c"),
                          "re": re, "tk": tk})
        except Exception as e:
            messagebox.showerror("マクロエラー", str(e))

    def toggle_record(self):
        self._recording = not self._recording
        if self._recording:
            self._macro_buffer = []
            self.rec_btn.config(text="⏹ 記録中...")
            self._st_rec.config(text="● REC")
        else:
            self.rec_btn.config(text="⏺ 記録")
            self._st_rec.config(text="")
            if self._macro_buffer:
                self._save_recorded()

    def _save_recorded(self):
        name = simpledialog.askstring("マクロ保存", "マクロ名を入力:", parent=self)
        if not name:
            return
        chars  = "".join(self._macro_buffer
                         ).replace("\\", "\\\\").replace("'", "\\'")
        script = f"editor.insert('insert', '{chars}')"
        macros = {}
        if os.path.exists(MACRO_FILE):
            try:
                with open(MACRO_FILE, encoding="utf-8") as f:
                    macros = json.load(f)
            except Exception:
                pass
        macros[name] = script
        with open(MACRO_FILE, "w", encoding="utf-8") as f:
            json.dump(macros, f, ensure_ascii=False, indent=2)
        self._load_macros()
        messagebox.showinfo("保存完了", f"「{name}」として保存しました")

    # ── 自動保存 ───────────────────────────────────────────────
    def _autosave_settings(self):
        dlg = tk.Toplevel(self)
        dlg.title("自動保存設定")
        dlg.resizable(False, False)
        dlg.attributes("-topmost", True)
        tk.Checkbutton(dlg, text="自動保存を有効にする",
                       variable=self.autosave_enabled).grid(
            row=0, column=0, columnspan=2, padx=12, pady=8, sticky="w")
        tk.Label(dlg, text="保存間隔 (秒):").grid(
            row=1, column=0, padx=12, pady=4, sticky="e")
        tk.Spinbox(dlg, textvariable=self.autosave_interval,
                   from_=10, to=3600, width=8).grid(row=1, column=1, padx=12)
        tk.Button(dlg, text="OK", width=10,
                  command=lambda: (self._restart_autosave(), dlg.destroy())
                  ).grid(row=2, column=0, columnspan=2, pady=8)

    def _start_autosave(self):
        self._autosave_stop.clear()
        threading.Thread(target=self._autosave_loop, daemon=True).start()

    def _restart_autosave(self):
        self._autosave_stop.set()
        enabled = self.autosave_enabled.get()
        self.autosave_lamp.config(
            text=f"自動保存: {'ON' if enabled else 'OFF'}",
            fg="#009900" if enabled else "#888888")
        self._start_autosave()

    def _autosave_loop(self):
        while not self._autosave_stop.wait(self.autosave_interval.get()):
            if self.autosave_enabled.get():
                self.after(0, self._do_autosave)

    def _do_autosave(self):
        saved = 0
        for tab in self.notebook.tabs():
            pane = self.notebook.nametowidget(tab)
            if isinstance(pane, EditorPane):
                try:
                    if pane.autosave():
                        saved += 1
                except Exception:
                    pass
        if saved:
            self._st_autosave.config(
                text=f"自動保存: {time.strftime('%H:%M:%S')}  ({saved} タブ)")

    # ── テーマ ─────────────────────────────────────────────────
    def toggle_dark(self):
        self.dark_mode = not self.dark_mode
        th = DARK if self.dark_mode else LIGHT
        for tab in self.notebook.tabs():
            pane = self.notebook.nametowidget(tab)
            if isinstance(pane, EditorPane):
                pane.apply_theme(th, self.dark_mode)

    def toggle_frame(self):
        self._frame_idx = (self._frame_idx + 1) % len(self.FRAME_COLORS)
        self.outer.config(bg=self.FRAME_COLORS[self._frame_idx])

    def change_font(self):
        size = simpledialog.askinteger(
            "フォントサイズ", "サイズを入力:",
            parent=self, initialvalue=11, minvalue=6, maxvalue=72)
        if size:
            for tab in self.notebook.tabs():
                pane = self.notebook.nametowidget(tab)
                if isinstance(pane, EditorPane):
                    pane.text.config(font=("Courier New", size))
                    pane.line_numbers.config(font=("Courier New", size))

    # ── 終了 ───────────────────────────────────────────────────
    def _on_close(self):
        for tab in list(self.notebook.tabs()):
            pane = self.notebook.nametowidget(tab)
            if isinstance(pane, EditorPane) and pane.modified:
                self.notebook.select(pane)
                name = os.path.basename(pane.filepath) if pane.filepath else "無題"
                ans  = messagebox.askyesnocancel(
                    "保存確認", f"「{name}」は変更されています。保存しますか？",
                    parent=self)
                if ans is None:
                    return
                if ans:
                    pane.save()
        self._autosave_stop.set()
        self.destroy()


# ===========================================================
# 実行
# ===========================================================
if __name__ == "__main__":
    if not HAS_DND:
        print("ヒント: ドラッグ&ドロップを有効にするには  "
              "pip install tkinterdnd2  を実行してください")
    app = TentekeEditor()
    app.mainloop()
