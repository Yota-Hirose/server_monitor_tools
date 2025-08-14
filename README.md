# スポーツコーダー監視作業記録システム

Webブラウザから監視データを入力して、自動的にExcelファイルを更新するシステムです。

## 🌐 Web版の使用方法（推奨）

### 1. GitHub Pagesでアクセス
- **URL**: `https://[あなたのGitHubユーザー名].github.io/server_monitor_tools/docs/`
- ブラウザで直接監視データを入力できます

### 2. リポジトリ設定（初回のみ）
1. GitHubリポジトリの **Settings** → **Pages** を開く
2. **Source** を `Deploy from a branch` に設定
3. **Branch** を `main` / `docs` に設定
4. **Personal Access Token** を生成:
   - [GitHub Settings > Developer settings > Personal access tokens](https://github.com/settings/tokens)
   - **Scopes**: `repo` にチェック
   - 生成されたトークンをコピー

### 3. Web入力からExcel自動更新
1. Web上のフォームで監視データを入力
2. **「📤 GitHub送信」** タブを選択
3. GitHub設定（トークン、リポジトリオーナー、リポジトリ名）を入力
4. **「📤 GitHubに送信」** ボタンをクリック
5. 🎉 **GitHub Actionsが自動実行され、Excelファイルが更新されます**

## 💾 ローカル版の使用方法

### 1. データ入力
- `docs/index.html` をブラウザで開く
- 監視記録を入力

### 2. Excel更新
- **「💾 ローカル保存」** → **「📁 JSONで保存」**
- `scripts/excel_update.bat` をダブルクリック

## 📋 入力項目

### 基本情報
- 確認日（YYYY/M/D形式）
- 確認時刻
- 確認者
- 確認結果（デフォルト: 問題なし）
- サーバ時刻同期（デフォルト: 問題なし）

### FTサーバユーティリティ
- CPUモジュール、PCIモジュール
- SCSIエンクロージャ

### HDD残容量
- C:, D:, E:, Y:, Z:ドライブ

### メモリ・CPU使用状況
- SQLServerメモリ、全体メモリ使用量
- CPU使用率、CPU確認時刻
- メモリ使用量状況

### サーバーランプ
- 上段・下段サーバーランプ点灯数

### SC機
- **HDD残容量**: SC-1〜SC-12のC/D/Hドライブ
- **CPU[%]/メモリ[GB]**: SC-1〜SC-12のCPU使用率・メモリ使用量

### 備考
- 特記事項（Z列に配置）

## 🔧 技術仕様

### Excel配置
| データ | セル位置 | 備考 |
|--------|----------|------|
| 確認日 | B5 | 日付のみ |
| 確認者 | C5 | |
| 確認結果 | D5 | 手動入力優先 |
| サーバ時刻同期 | J9 | |
| SC機HDD C | H24-H35 | SC-1〜SC-12 |
| SC機HDD D | M24-M35 | SC-1〜SC-12 |
| SC機HDD H | R24-R35 | SC-1〜SC-12 |
| SC機CPU（奇数） | G列 | SC-1,3,5,7,9,11 |
| SC機CPU（偶数） | Q列 | SC-2,4,6,8,10,12 |
| SC機メモリ（奇数） | K列 | SC-1,3,5,7,9,11 |
| SC機メモリ（偶数） | U列 | SC-2,4,6,8,10,12 |
| 備考 | Z5 | |

### GitHub Actions連携
- **トリガー**: `repository_dispatch`
- **イベントタイプ**: `update-monitoring-data`
- **自動処理**: Excelファイル更新 → コミット → プッシュ

## 📦 必要な環境

### ローカル版
- Python 3.x
- openpyxlライブラリ

### Web版
- GitHubアカウント
- Personal Access Token (repo権限)
- GitHub Pagesが有効なリポジトリ

## 🚀 セットアップ

1. このリポジトリをフォークまたはクローン
2. GitHub Pages設定を有効化
3. Personal Access Tokenを生成
4. Webフォームからトークン設定
5. 監視データ入力 → 送信で自動更新開始！

## 📄 出力形式

- **Excel**: `data/スポーツコーダ監視作業履歴.xlsx`
- **CSV**: ブラウザダウンロード
- **JSON**: ブラウザダウンロード

---

**🎯 Web版なら、どこからでもアクセスして即座にExcelを更新できます！**