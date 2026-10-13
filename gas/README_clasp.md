# clasp の設定（main.gs を手で貼らなくて済むようにする）

**設定は完了しています。** これからは下の2コマンドだけで Apps Script に反映できます。

```
git pull          ← VS Code側の修正を取り込む
clasp push        ← Apps Script に反映
```

---

## 設定済みの内容（2026-10-13）

| 項目 | 内容 |
|---|---|
| clasp | グローバルに導入（3.4.1） |
| ログイン | `s.kimura@lime-fit.com`（認証情報はユーザーフォルダの `.clasprc.json`） |
| `.clasp.json` | スクリプトID ／ `rootDir: gas` ／ `fileExtension: gs` |
| `.claspignore` | `gas/` の中の `.md` だけ送らない（`index.html` や `data/` は元から対象外） |
| `gas/appsscript.json` | Apps Script側の本物の設定（JST・V8・Webアプリ設定）を取り込み済み |
| `.gitignore` | 認証情報（`.clasprc.json`）を除外 |

`fileExtension: gs` を入れてあるので、`clasp pull` でも `.js` ではなく `.gs` で降りてきます。

---

## 毎回の使い方

### リポジトリ → Apps Script

```
clasp push
```

送られるファイルを先に確認したいとき：

```
clasp status
```

### Apps Script → リポジトリ（画面で直接いじったとき）

```
clasp pull
git diff gas/
```

---

## 気をつけること

| 注意 | 理由 |
|---|---|
| **ファイル名は Apps Script 側と完全に揃える** | 名前が違うと**新しいファイルが増え、関数が二重定義になります**。`hpb.gs` → `HPB.gs` に直したのはこのため |
| **関数名を他のファイルと重複させない** | Apps Script は全ファイルが同じスコープです。同名だと後から読まれたほうが勝ちます（実例は下記） |
| `clasp push` は Apps Script 側を**上書き**する | 画面で直接いじった分があると消えます。心配なときは先に `clasp pull` で確認 |
| 認証情報は共有しない | Apps Script を操作できる鍵です |
| スクリプトプロパティは clasp では送られない | `SS_MASTER` などは Apps Script の画面で管理します |

### 実際に起きた関数名の衝突（2026-10-13）

`hpb.gs` と `main.gs` の両方に `getHpbSS_()` がありました。
`hpb.gs` 側は「`SS_HPB` が未設定なら例外を投げる」作りで、そちらが勝っていたため、
**ブランド別内訳の PV・CVR・ACR が168件すべて空**になっていました。
`main.gs` 側を `getHpbSSForBrand_()` に改名して解決しています。

衝突の確認方法：

```
cat gas/*.gs | grep -o '^function [A-Za-z0-9_]*' | sort | uniq -d
```

何も出なければ重複なしです。**`gas/` にファイルを足すときは必ずこれを確認してください。**

---

## うまくいかないとき

| 症状 | 対処 |
|---|---|
| `User has not enabled the Apps Script API` | https://script.google.com/home/usersettings でAPIをオンにする |
| `Could not read API credentials` | `clasp login` をやり直す |
| `Invalid script ID` | `.clasp.json` の `scriptId` を確認する |
| push したのに反映されない | Apps Script の画面を再読み込み。それでもダメなら `clasp status` で送信対象を確認 |
