# 【管理者向け】駐車場データダッシュボード 運用マニュアル

このファイルは、システムの管理者や開発者がダッシュボードを保守・運用するためのマニュアルです。

## 1. データの追加・更新手順
データソースは `data` フォルダ内に集約されています。

### ① 経営ダッシュボード（全社・一般利用集計）の更新
- `data/updated_integrated_data_FY2025.csv.gz` がソースファイルです。
- 新しい月や年度のデータを追加した場合は、同じファイル名で上書き保存してください。
- （圧縮せずに `.csv` のまま置きたい場合は、`unified_dashboard.py` の 221行目にある `file_path` を `.csv` に書き換えてください）

### ② 稼働分析プロ（時間帯別・タブ比較分析）の更新
- `data/excel2025_with_avg` 等のフォルダが存在します。
- 新年度（例: 2026年）のExcelファイルを受領した際は、ここに `excel2026_with_avg` フォルダを新設し、その中に各拠点のExcelファイルを入れてください。
- プログラムは自動的に「2023」「2024」「2025」までの年を読み込む設定になっているため、新年度を追加する場合は `unified_dashboard.py` の 398行目 `years = [2023, 2024, 2025]` に `2026` を足してください。

## 2. アプリの修正とGitHubへの反映
プログラム自体は `unified_dashboard.py` となります。

### GitHubへの保存・同期
現在のローカル変更を非公開リポジトリ（`sugichango/parking`）の `main` ブランチに同期するには、コマンドプロンプトやPowerShellで以下を実行します。

```bash
cd c:\Users\sugitamasahiko\Documents\parking_system
git add .
git commit -m "Update dashboard modifications"
git push origin main
```

### Streamlit Community Cloud (オンライン版) の更新
- GitHub上の `main` ブランチにプッシュされると、[share.streamlit.io](https://share.streamlit.io) にデプロイされているクラウド版も「数分以内に自動的に更新」されます。
- 初期設定のパスワードは `.streamlit/secrets.toml` に保存されていますが、クラウド版のパスワードはStreamlit Cloud側の「Advanced settings > Secrets」にて変更可能です。
