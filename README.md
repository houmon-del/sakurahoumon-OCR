# sakurahoumon-OCR

相談シート・居宅療養管理指導報告書のOCR → AI構造化 → DentNet CSV出力 Webアプリ

## 技術スタック
- **OCRエンジン**: yomitoku v0.12.0（日本語AI OCR、4モデル構成）
- **AI構造化**: Claude Sonnet（Vision API + OCRテキスト）
- **バックエンド**: Flask + gunicorn
- **デプロイ**: Google Cloud Run（gen2, 8Gi, 4 vCPU）
- **CI/CD**: Cloud Build（GitHub push トリガー）

## ローカル起動
```bash
pip install -r requirements.txt
python app.py  # http://localhost:5002
```

## Cloud Run デプロイ

### 重要: Cloud Build トリガー設定
現在のCloud Buildトリガー (`d57efa97`) はデフォルトの「Deploy to Cloud Run」設定を使っており、
**cloudbuild.yaml を使っていない**。そのためメモリ/CPU設定が反映されない。

**対策（どちらか）:**
1. GCPコンソールでトリガーを「Cloud Build構成ファイル（cloudbuild.yaml）」に変更
2. pushのたびに `gcloud run services update` で手動適用:
```bash
gcloud run services update sakurahoumon-ocr \
  --region=asia-northeast1 \
  --project=elite-campus-480105-h9 \
  --memory=8Gi --cpu=4 --timeout=300s \
  --min-instances=1 --max-instances=1 \
  --session-affinity --execution-environment=gen2 --cpu-boost
```

### Dockerイメージ構成
- ベース: `python:3.11`（フルイメージ、slimだとPyTorch/ONNX用ライブラリ不足でSIGABRT）
- PyTorch CPU-only + ONNX Runtime
- yomitokuの4モデルをビルド時にプリダウンロード (`RUN python -m yomitoku.cli.download_model`)
- スレッド制御: OMP/MKL/TORCH/ORT すべて4（vCPU数と一致）
- ONNX推論モード有効（PyTorch推論より2-5倍高速）

## 機能

### 通常OCRモード
PDF/画像アップロード → yomitoku OCR → 結果表示（段落+テーブル）→ エクスポート（JSON/MD/CSV/XLSX）

### 相談シートモード
複数PDF/ZIP一括アップロード → OCR → Claude AI構造化抽出 → インライン編集 → DentNet CSV出力（CP932, 22列）

## ファイル構成
- `app.py` — Flaskルート（通常OCR + 相談シートモード）
- `ocr_engine.py` — yomitoku OCRエンジン（lazy loading, ONNX推論, バッチ処理）
- `ai_corrector.py` — Claude AI（OCR校正 + 構造化抽出 + 相談シート専用抽出）
- `consultation_csv.py` — DentNet CSV変換（和暦→西暦、22列フォーマット）
- `Dockerfile` — Cloud Run用コンテナ（python:3.11 + PyTorch CPU + ONNX + モデルプリDL）
- `cloudbuild.yaml` — Cloud Build設定（※トリガーが使っていない場合あり、上記参照）

## 既知の課題
- CPU推論は1ページ1-3分（初回はモデルロード30秒追加）
- インメモリジョブ管理のためmax-instances=1必須（スケールアウト不可）
- Cloud Buildトリガーがcloudbuild.yamlを使わない問題（手動設定適用が必要）

## 環境変数
- `ANTHROPIC_API_KEY` — Claude AI用（.envファイルに設定）
- `PORT` — Cloud Runが自動設定（デフォルト8080）
