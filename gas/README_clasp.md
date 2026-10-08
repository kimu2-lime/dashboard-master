# clasp の設定（main.gs を手で貼らなくて済むようにする）

clasp は Google 公式の Apps Script 用コマンドです。
これを設定すると、**`clasp push` の1コマンドで `gas/` の中身が Apps Script に反映**されます。

準備はもう済ませてあります。**キム兄さんにやってもらうのは「①ログイン」と「②IDの登録」の2つだけ**です。

| 済んでいること | 内容 |
|---|---|
| clasp のインストール | `npm install -g @google/clasp`（バージョン 3.4.1） |
| `.claspignore` | `gas/` の中だけを送る設定。`index.html` や `data/` は送られません |
| `gas/appsscript.json` | Apps Script の設定ファイル（タイムゾーン Asia/Tokyo・V8） |
| `.gitignore` | ログイン情報（`.clasprc.json`）をGitHubに上げない設定 |

---

## ① Googleにログインする（1回だけ）

ターミナルで、このフォルダに移動して実行します。

```
cd "C:\Users\skimu\OneDrive\デスクトップ\dashboard-master"
clasp login
```

ブラウザが開くので、**ダッシュボードのApps Scriptを持っているGoogleアカウント**でログインし、
アクセスを許可してください。

> 「Apps Script APIが無効です」と言われたら、
> https://script.google.com/home/usersettings を開いて
> **「Google Apps Script API」をオン**にしてから、もう一度 `clasp login` してください。

ログイン情報は `C:\Users\skimu\.clasprc.json` に保存されます。
**このファイルは他人に渡さないでください**（Apps Scriptを操作できる鍵です）。

---

## ② スクリプトIDを登録する（1回だけ）

**スクリプトIDの調べ方：**

1. https://script.google.com を開き、ダッシュボードのプロジェクトを開く
2. 左下の **⚙️（プロジェクトの設定）** をクリック
3. 「スクリプト ID」をコピー（`1A2B3C...` のような長い文字列）

コピーしたIDで、このフォルダに `.clasp.json` を作ります。

```
cd "C:\Users\skimu\OneDrive\デスクトップ\dashboard-master"
clasp clone <ここにスクリプトID> --rootDir gas
```

> `clasp clone` は **Apps Script側の中身をこちらに持ってきます**。
> 先に持ってくることで、「リポジトリに無いファイルが向こうにある」場合に
> それを消してしまう事故を防げます。

持ってきたあと、**何が変わったかを必ず確認してください**。

```
git status
git diff gas/
```

- `main.gs` だけが変わっている（＝私の修正が入っていない古い版が降ってきた）なら、
  `git checkout gas/main.gs` で戻してから次に進みます
- **見覚えのない `.gs` ファイルが増えていたら、それは Apps Script 側にしか無いコードです。**
  消さずにそのままコミットしてください（次回の `clasp push` で消えてしまうため）

---

## 毎回の使い方

### リポジトリ → Apps Script に反映する

```
clasp push
```

送られるファイルの一覧は、先に確認できます。

```
clasp status
```

### Apps Script → リポジトリに取り込む（画面で直接いじったとき）

```
clasp pull
git diff gas/
```

---

## これからの流れ

私（VS Code側のClaude）が `gas/` を直したら、こう案内します。

```
git pull          ← 私の修正を取り込む
clasp push        ← Apps Script に反映
```

**貼り付け作業はもう不要になります。**

---

## 気をつけること

| 注意 | 理由 |
|---|---|
| `clasp push` は Apps Script 側を**上書き**します | 画面で直接いじった分があると消えます。心配なときは先に `clasp pull` して確認 |
| `.clasprc.json` は共有しない | Apps Script を操作できる認証情報です。`.gitignore` 済み |
| `gas/appsscript.json` を消さない | これが無いと `clasp push` が失敗します |
| スクリプトプロパティは clasp では送られません | `SS_MASTER` などの設定は Apps Script の画面で管理します |

---

## うまくいかないとき

| 症状 | 対処 |
|---|---|
| `User has not enabled the Apps Script API` | https://script.google.com/home/usersettings でAPIをオンにする |
| `Could not read API credentials` | `clasp login` をやり直す |
| `Invalid script ID` | `.clasp.json` の `scriptId` を確認する（プロジェクトの設定からコピーし直す） |
| push したのに反映されない | Apps Script の画面を再読み込みする。それでもダメなら `clasp status` で送信対象を確認 |
