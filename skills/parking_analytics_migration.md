# Skill: 駐車場分析ダッシュボード構築とGoogle Drive連携

## 概要
このスキルは、ローカルCSVベースの駐車場分析ダッシュボードを、Google Drive API連携版に移行し、Streamlit Cloudで安定稼働させるための完全なガイドラインです。

## 技術スタック
- **Python**: 3.12 (Streamlit Cloudでの依存関係ビルド安定のため)
- **ライブラリ**: `streamlit`, `pandas`, `plotly`, `google-api-python-client`, `google-auth`
- **データソース**: Google Drive 共有ディレクトリ (Excel/CSV)

## 1. データ構造と加工ロジック
### 駐車場コードの変換
- **440**: 南１
- **441**: 南２
- **442**: 南３
- **443**: 南４
- **444**: 北１
- **445**: 北２
- **446**: 北３

### KPI計算
- **最大稼働率**: (在庫台数のピーク値 / 指定された収容台数) * 100
- **一般回転率**: ((一般入庫合計 + 一般出庫合計) / 2) / 収容台数

## 2. Google Drive API連携 (標準パターン)
サービスアカウントを使用してファイルを探索し、バイナリとしてダウンロードする。
```python
def download_file_from_drive(file_id):
    # build('drive', 'v3', credentials=credentials) を使用
    request = service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    # ... chunked download ...
    return fh
```

## 3. UI/UX デザイン要件 (極限可読モード)
Streamlit Cloud の暗い背景で KPI ラベルを確実に読ませるための CSS テンプレート。
- **KPIラベル**: `font-weight: 900 !important; color: #FFFFFF !important; text-shadow: 1px 1px 3px rgba(0,0,0,1.0);`
- **フォント**: `Inter` (Google Fonts)

## 4. デプロイ・トラブルシューティング
- **Oh no エラー**: 多くの場合は Python バージョンの不整合。必ず 3.12 ベースで構成する。
- **Secretsの形式**: `gcp_service_account` は1つの文字列として、サービスアカウントJSONを `'''` で囲んで貼り付ける。
- **403エラー**: Google Cloud Console で "Google Drive API" を「有効」にすることを忘れない。
- **404エラー**: フォルダIDの綴りミス（特に '7' と 'z'、'j' と 'Z' の取り違え）を最優先で確認する。

## このスキルの呼び出し方
「駐車場データ分析スキルの手順に従って、[新機能名] を追加して」と指示してください。
