# CLAUDE.md

LIME FIT 売上ダッシュボード（GitHub Pages 公開）。

- `index.html` … ダッシュボード本体（GitHub Pages が `main` ブランチから配信）
- `data/data.json` … GAS が生成する実績・目標データ
- `gas/main.gs` … `data.json` を生成する Google Apps Script のソース

## プッシュ運用ルール

- **ダッシュボード（`index.html` など）に変更を加え、動作に問題がなければ、ユーザーに言われなくても commit して `origin/main` に push する。** 確認を待たずにプッシュしてよい。
- GitHub Pages は `main` ブランチから配信されるため、コミットは `main` に直接行う（ブランチを切らない）。
- コミットメッセージは日本語で `feat:` / `fix:` / `docs:` などの接頭辞を付ける。
- コミット時は対象ファイルを明示的に `git add` する（`.claude/settings.local.json` などローカル設定を巻き込まない）。
- `data/data.json` は GAS の `exportToGitHub` が自動コミット・プッシュするファイル。手動で編集・コミットしない。
- push 前に `git fetch` し、リモートが進んでいれば `git rebase origin/main` してから push する（GAS の自動コミットと衝突しやすいため）。

## 分析・実装ルール

課題分析・施策提案・UI実装を行う際は、`CLAUDE_RULES.md` と `knowledge/` 配下の markdown を先に読むこと。
