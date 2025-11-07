# 🏌️‍♂️ Google Places API インドアゴルフ施設収集ツール

このプロジェクトは、**Google Places API (New)** と **Google Apps Script (GAS)** を利用して、  
関東エリア（茨城・栃木・群馬・埼玉・千葉・東京・神奈川）の  
インドアおよびアウトドアのゴルフ施設情報を自動収集・一覧化するシステムです。

---

## 🚀 主な機能

### 1. Googleスプレッドシート連携
- スプレッドシート上にカスタムメニュー「IndoorGolf」を追加  
- 「関東収集（実行）」ボタンからワンクリックでデータ収集を開始  
- 結果は自動的に `Results` シートへ表形式で出力

### 2. 取得項目
| 項目名 | 説明 |
|--------|------|
| 店舗ID | Google上の店舗ID |
| 店舗名 | ゴルフ施設の名称 |
| 住所 | 施設住所 |
| 種別 | アウトドア or インドア |
| 電話番号 | 店舗の電話番号 |
| WebサイトURL | 公式サイト |
| 評価 (rating) | Google上の平均評価 |
| 口コミ数 (userRatingCount) | レビュー件数 |
| 営業ステータス (businessStatus) | OPEN / CLOSED など |
| 検索地域名・キーワード | 検索条件 |

### 3. 検索条件
- キーワード例：「インドアゴルフ」「ゴルフ練習場」「ゴルフスクール」など  
- 関東主要都市の中心座標から半径15〜20kmで検索  
- 最大 500〜1000件程度のデータを想定

### 4. CSV出力機能
- 取得結果シート (`Results`) をワンクリックでCSV形式に変換  
- 自動でGoogleドライブに保存

### 5. エラーハンドリング
- APIエラーや実行中の例外発生時にアラートを表示  
- 原因を特定しやすいログを出力

---

## ⚙️ 使用技術

- **Google Apps Script (GAS)**
- **Google Spreadsheet**
- **Google Places API (New)**
- **Google Drive API**

---

## 🧩 アーキテクチャ概要

ユーザー操作
↓
Googleスプレッドシート
↓ (カスタムメニュー)
Apps Script 実行
↓
Google Places API へリクエスト
↓
施設情報をJSONで受信
↓
スプレッドシート「Results」へ出力
↓
（オプション）CSVとしてDriveに保存

---

## 🧰 実行方法

1. スプレッドシートを開き、メニューから「IndoorGolf」→「関東収集（実行）」をクリック  
2. データ取得が開始され、「Results」シートに自動で結果が出力されます  
3. CSV保存が必要な場合は、「CSV保存」をクリックしてGoogleドライブにエクスポート  

---

## 🔒 注意事項

- Google Places APIキーは、実行ユーザーが別途用意する必要があります  
- Apps Scriptの実行時間（6分）制限を考慮し、100件ごとのバッチ処理に対応  
- 将来的に全国対応へ拡張可能な構成で設計  

---

## 🧾 デモ・サンプル

- 📊 [デモスプレッドシート（閲覧用）](https://docs.google.com/spreadsheets/d/1ld_pRSNTG1ni-uqLju9SPJx3ccxb6b44EwvDca32oRE/edit?gid=644591532#gid=644591532)  
- 💾 [CSV出力例（Googleドライブ）](https://drive.google.com/file/d/1yHtIVS7p6wIMp8vsFOCb7ivRCEawItpv/view?usp=drive_link)

---

## 🧑‍💻 開発者コメント

本ツールは「API連携 × 自動化 × スプレッドシート操作」を組み合わせた  
GASプロジェクトとして開発しました。  
実行効率と拡張性を重視し、将来的に検索エリアやキーワードを追加することで  
全国規模のデータ収集にも対応可能です。

---

## 📄 ライセンス

MIT License

---
