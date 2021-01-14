# ExcelFormulaEditor - エクセル数式らくらく入力アドイン

## 概要

Excelの数式を手軽に良い感じでインデントして、数式の編集を効率化するためのExcelアドインです。

* 数式を自動インデント
  * 1行表示
  * ネストレベルを指定したブロック表記
  * ツリー表記
* 不完全な数式を赤く表示
* 数式の実行結果を表示しログが残る
* 操作はキーボードオンリーにも、マウス操作にも対応
* アクティブセルに対する情報の取得する隠し機能（[value]、[text]、[formula]、[row]、[col]プロパティ）

![](https://www.dropbox.com/s/85zcc6qy31md56i/20200730_%E3%82%A8%E3%82%AF%E3%82%BB%E3%83%AB%E3%81%AE%E6%95%B0%E5%BC%8F%E5%85%A5%E5%8A%9B%E3%82%92%E6%A5%BD%E3%81%AB%E3%81%99%E3%82%8B%E3%82%A2%E3%83%89%E3%82%A4%E3%83%B3%E3%82%92%E4%BD%9C%E3%81%A3%E3%81%A6%E3%81%BF%E3%81%9F_01.png?raw=1)

<br>

## 使い方

### ダウンロード

[ExcelFormulaEditor/master/bin/エクセル数式らくらく入力アドイン.xlam](
https://raw.githubusercontent.com/KotorinChunChun/ExcelFormulaEditor/master/bin/%E3%82%A8%E3%82%AF%E3%82%BB%E3%83%AB%E6%95%B0%E5%BC%8F%E3%82%89%E3%81%8F%E3%82%89%E3%81%8F%E5%85%A5%E5%8A%9B%E3%82%A2%E3%83%89%E3%82%A4%E3%83%B3.xlam) からダウンロード

### 初めての立ち上げ方

単にダウンロードして開いても、全く反応が無いことがあります。

* binフォルダよりアドインファイルをダウンロード
* ファイルのプロパティを開いて「セキュリティ：～～～～～☑許可する」にチェックを入れてOK

  ![](https://www.dropbox.com/s/kmr79wzlu9xzg4j/20200730_%E3%82%A8%E3%82%AF%E3%82%BB%E3%83%AB%E3%81%AE%E6%95%B0%E5%BC%8F%E5%85%A5%E5%8A%9B%E3%82%92%E6%A5%BD%E3%81%AB%E3%81%99%E3%82%8B%E3%82%A2%E3%83%89%E3%82%A4%E3%83%B3%E3%82%92%E4%BD%9C%E3%81%A3%E3%81%A6%E3%81%BF%E3%81%9F_02.png?raw=1)
  
* ダブルクリックで開く
* マクロを有効化する
* Excelのリボンに機能が増える

![](https://www.dropbox.com/s/lh20i9s3qfaejn7/20200730_%E3%82%A8%E3%82%AF%E3%82%BB%E3%83%AB%E3%81%AE%E6%95%B0%E5%BC%8F%E5%85%A5%E5%8A%9B%E3%82%92%E6%A5%BD%E3%81%AB%E3%81%99%E3%82%8B%E3%82%A2%E3%83%89%E3%82%A4%E3%83%B3%E3%82%92%E4%BD%9C%E3%81%A3%E3%81%A6%E3%81%BF%E3%81%9F_03.png?raw=1)

### 常駐させたい場合

* ファイルをアドインフォルダにコピー
* 起動時に立ち上がるように設定

詳しくはGoogleで。

### フォームの開き方

三通りの方法があります。

1. `Ctrl+2`キー　（太字／解除 を上書きします。Ctrl+Bがあるから不要ですよね？）
2. 監視中に数式の入ったセルをダブルクリック
3. 「数式エディタ起動」コマンドを実行

※アドインを開いた時点で、自動的に「`Ctrl+2`の上書き」と「ダブルクリックの監視」が始まります。

### フォーマットの採用

三通りの方法で画面右側のフォーマットを切り替えられます。

1. `Alt+数字` キー
2. `Ctrl+TAB` / `Ctrl+Shift+TAB` キー
3. タブのマウスクリック

二通りの方法で選択中のフォーマットを数式入力ボックスに採用できます。

1. `Alt+数字` キーの連打
2. `F5` キー

### 入力の確定

* `Ctrl+Enter` で入力を確定（※配列数式チェックボックスの状態に従う）
* `Ctrl+Shift+Enter` で配列数式として入力を確定

### 入力の取り消し

* `Ctrl+Z` で編集を元に戻す（※フォーマット採用時点以降の変更のみ）
* `ESC` キーでダイアログを閉じる
* マウスで×をクリックする

<br>

## 利用風景

![](https://www.dropbox.com/s/jnq6612el3y8tq1/20200730_%E3%82%A8%E3%82%AF%E3%82%BB%E3%83%AB%E3%81%AE%E6%95%B0%E5%BC%8F%E5%85%A5%E5%8A%9B%E3%82%92%E6%A5%BD%E3%81%AB%E3%81%99%E3%82%8B%E3%82%A2%E3%83%89%E3%82%A4%E3%83%B3%E3%82%92%E4%BD%9C%E3%81%A3%E3%81%A6%E3%81%BF%E3%81%9F_01.gif?raw=1)

![](https://www.dropbox.com/s/b0j90th4vurxiqh/20200730_%E3%82%A8%E3%82%AF%E3%82%BB%E3%83%AB%E3%81%AE%E6%95%B0%E5%BC%8F%E5%85%A5%E5%8A%9B%E3%82%92%E6%A5%BD%E3%81%AB%E3%81%99%E3%82%8B%E3%82%A2%E3%83%89%E3%82%A4%E3%83%B3%E3%82%92%E4%BD%9C%E3%81%A3%E3%81%A6%E3%81%BF%E3%81%9F_04.gif?raw=1)

<br>

## お約束

* このプログラムは[MITライセンス](https://ja.wikipedia.org/wiki/MIT_License)で公開しています。
* 利用は自己責任でお願いします。
* バグ報告は大歓迎ですが、必ず修正するとは限りません。

### MITライセンスの要約

こんな感じです。

```
* コピーも、再配布も、変更も、業務利用も、販売品に含めてもいいよ
* でも、著作者表記とライセンス表記は消さないでね
* このプログラムを使って何が起きても作者は関知しませんよ
```

<br>

### その心は？

**世界中に私の生きた証を残すのじゃー！**

<br>

## 作者情報

@KotorinChunChun

https://twitter.com/KotorinChunChun

https://www.excel-chunchun.com/

<br>

## 更新履歴

### 2020/07/27 Ver.0.10

* 最低機能で初回リリース

### 2020/07/28 Ver.0.11

* 若干バグ修正

### 2021/01/15 Ver.0.12 配列数式対応と操作性向上

#### 全般
* 全コントロールフォーカス時のキー操作に対応
* Ctrl+Shift+Enterで配列数式自動設定に対応
* チェックボックスで配列数式切り替えに対応

#### 入力ボックス
* スクロールバー追加
* 文字サイズ変更ボタン追加
* エラー時の背景色を淡い赤色に変更

#### フォーマット候補
* アクティブタブに着色
* Ctrl+TAB / Ctrl+Shift+TABでタブ切り替えに対応
* Ctrl+数字の連続入力でフォーマット採用

<br>

## 残りの課題

* 式が255文字以上になるとエラーで確定できない
* 参照しているセルの値ベースの数式プレビュー
* 数式バーの高さ変更機能
* 複数セルへの一括適用機能
* 『元に戻す』を壊さないようクリップボードコピーを追加
* Excelに新たに実装されたLET関数に適したフォーマット開発

