# 実例

| ファイル | 中身 |
|---|---|
| `hokudai-ht3/build_ht3_draft.py` | 北大病院の情報公開用文書（HT3、オプトアウト）の下書きを作る。行の中で色が混ざるひな形の片付け方、①単独版・②多機関版の2ページの扱い、表の例示行、ヘッダーの先頭ページ別指定。`--template <HT3雛形> --variant single|multi --out 出力.docx` |

ひな形は機関の配布物なので同梱していない。機関のページから取って研究のフォルダの `様式/` に置き、`--template` で渡す。

## 新しい研究の計画書・説明文書（HT1/HT2）を作るとき

1. `scripts/dump_template.py HT1….docx --runs --out 構造.txt` でひな形を読む。
2. 研究のフォルダの `tools/build_<研究名>.py` を作り、冒頭に定数（題名、日付、研究者、機関、出力先）を置く。
3. `from docx_tools import *` で関数を読み、`SKILL.md` の「作成の手順」6の流れで書く。章ごとの本文は `Inserter.after(r"^N．…$", [...])` に入れる（書き方は `references/sections.md`）。
4. 背景の引用は `{{n}}`、文献リストは `references_meta.json` の `reference` を貼る。
5. まず作業フォルダに出力し、`audit_docs.py` で要修正0件にしてから提出フォルダへ。

出力先が提出フォルダのスクリプトは、テストやサブエージェントに実行させない（提出物を上書きしてしまう）。
