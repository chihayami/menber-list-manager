# 名簿マネージャー（統合プロジェクト）

本プロジェクトは, LINE Botを介して名簿データを自動更新するマスターシステム（master）と, 更新された名簿データをコピーしやすい形式でプレビュー表示およびコピーできるビューアシステム（viewer）を統合して管理するリポジトリである.

---

## ディレクトリ構造
- master/ : LINEからのメッセージに応じて会員名簿スプレッドシートを更新するプロジェクト.
  - src/ : Google Apps Script（GAS）にアップロードされるプログラムコード（README.jsを含む）.
- viewer/ : 事務作業用に名簿データを取得し, コピペ用プレビューを生成するプロジェクト.
  - src/ : GASにアップロードされるプログラムコードおよびHTMLファイル（README.jsを含む）.
- docs/ : システムの各種説明ドキュメント.
  - [メンテナンスガイド (maintainer_guide.md)](file:///c:/Users/AI/workspace_AI/member-list-manager/docs/maintainer_guide.md) : メンテナーおよびシステム管理者向けのメンテナンスおよび初期設定ガイド.
  - [利用ガイド (user_guide.md)](file:///c:/Users/AI/workspace_AI/member-list-manager/docs/user_guide.md) : 実務担当者および役員向けのシステム利用ガイド.

---

## 開発とテスト
masterおよびviewerの各ディレクトリにおいて, Node.jsを利用したローカル環境でのテストとclaspによるデプロイが可能である.

### 1. 依存関係のインストール
プロジェクトの各ディレクトリで以下を実行する.
- masterの場合:
  cd master
  npm install
- viewerの場合:
  cd viewer
  npm install

### 2. ローカルテストの実行
各ディレクトリにて, モックを使用したテストを実行できる.
- コマンド:
  npm test

### 3. GASへのデプロイ
claspを利用して, ローカルのコードをGoogle Apps Scriptへアップロードできる.
- コマンド:
  npm run push
