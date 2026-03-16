# sakurahoumon-OCR

相談シート・居宅療養管理指導報告書のOCR → AI構造化 → DentNet CSV出力 Webアプリ

## 技術スタック
- **OCRエンジン**: yomitoku（日本語AI OCR、4モデル構成、PyTorch CUDA推論）
- **AI構造化**: Claude Sonnet（Vision API + OCRテキスト）
- **バックエンド**: Flask + gunicorn
- **デプロイ**: Google Compute Engine GPU VM（T4 GPU）
- **コンテナ**: Docker（pytorch/pytorch:2.6.0-cuda12.6-cudnn9-runtime）

## デプロイ（GPU VM）

### 前提
- GCP プロジェクト: `elite-campus-480105-h9`
- Docker イメージ: Artifact Registry に `:gpu` タグでビルド済み
- ファイアウォール: `ocr-server` タグで port 8080 開放済み

### 1. Docker イメージをビルド（Cloud Shell）
```bash
cd ~/sakurahoumon-OCR
gcloud builds submit --tag asia-northeast1-docker.pkg.dev/elite-campus-480105-h9/cloud-run-source-deploy/sakurahoumon-ocr/sakurahoumon-ocr:gpu .
```

### 2. startup.sh を作成（Cloud Shell）
```bash
cat > startup.sh << 'SCRIPT'
#!/bin/bash
cos-extensions install gpu
mount --bind /var/lib/nvidia /var/lib/nvidia
mount -o remount,exec /var/lib/nvidia
sleep 10
docker-credential-gcr configure-docker --registries=asia-northeast1-docker.pkg.dev
docker rm -f ocr 2>/dev/null
docker run -d --name ocr --restart=always \
  --volume /var/lib/nvidia/lib64:/usr/local/nvidia/lib64 \
  --volume /var/lib/nvidia/bin:/usr/local/nvidia/bin \
  --device /dev/nvidia0:/dev/nvidia0 \
  --device /dev/nvidia-uvm:/dev/nvidia-uvm \
  --device /dev/nvidiactl:/dev/nvidiactl \
  -e LD_LIBRARY_PATH=/usr/local/nvidia/lib64 \
  -e ANTHROPIC_API_KEY="YOUR_KEY_HERE" \
  -p 8080:8080 \
  asia-northeast1-docker.pkg.dev/elite-campus-480105-h9/cloud-run-source-deploy/sakurahoumon-ocr/sakurahoumon-ocr:gpu
SCRIPT
```
※ `YOUR_KEY_HERE` を実際の Anthropic API キーに置き換えること

### 3. GPU VM を作成（Cloud Shell）
```bash
gcloud compute instances create sakurahoumon-ocr-gpu \
  --zone=us-central1-a \
  --machine-type=n1-standard-2 \
  --accelerator=type=nvidia-tesla-t4,count=1 \
  --maintenance-policy=TERMINATE \
  --image-family=cos-stable \
  --image-project=cos-cloud \
  --tags=ocr-server \
  --scopes=cloud-platform \
  --boot-disk-size=50GB \
  --metadata-from-file=startup-script=startup.sh
```
※ アジアリージョンはT4枯渇が多い。`us-central1-a` が確実。

### 4. 起動・停止
```bash
# 起動
gcloud compute instances start sakurahoumon-ocr-gpu --zone=ZONE

# 停止
gcloud compute instances stop sakurahoumon-ocr-gpu --zone=ZONE

# IP確認（起動後）
gcloud compute instances describe sakurahoumon-ocr-gpu --zone=ZONE \
  --format='get(networkInterfaces[0].accessConfigs[0].natIP)'
```
起動スクリプトが自動でGPUドライバ＋コンテナを立ち上げるので、2〜3分待てば `http://IP:8080` でアクセス可能。

## ローカル起動
```bash
pip install -r requirements.txt
python app.py  # http://localhost:5002
```

## 機能

### 通常OCRモード
PDF/画像アップロード → yomitoku OCR → 結果表示（段落+テーブル）→ エクスポート（JSON/MD/CSV/XLSX）

### 相談シートモード
複数PDF/ZIP一括アップロード → OCR → Claude AI構造化抽出 → インライン編集 → DentNet CSV出力（CP932, 22列）

## ファイル構成
- `app.py` — Flaskルート（通常OCR + 相談シートモード）
- `ocr_engine.py` — yomitoku OCRエンジン（lazy loading, PyTorch CUDA推論, バッチ処理）
- `ai_corrector.py` — Claude AI（OCR校正 + 構造化抽出 + 相談シート専用抽出）
- `consultation_csv.py` — DentNet CSV変換（和暦→西暦、22列フォーマット）
- `Dockerfile` — GPU対応コンテナ（PyTorch CUDA + モデルプリDL）
- `cloudbuild.yaml` — Cloud Build設定

## 技術メモ
- `infer_onnx: False` — ONNX Runtime の opset 17/18 変換エラーを回避
- GPU (CUDA) 推論で ONNX 不要（PyTorch 直接推論で十分高速）
- COS (Container-Optimized OS) では `cos-extensions install gpu` が必要
- インメモリジョブ管理のためインスタンスは1つのみ

## 環境変数
- `ANTHROPIC_API_KEY` — Claude AI用（VM起動スクリプトに設定）
- `PORT` — デフォルト8080
