#!/bin/bash
# GPU VMでコンテナを再起動するスクリプト
# 引数: $1 = PROJECT_ID
set -e
PROJECT_ID="$1"
DOCKERCFG=$(mktemp -d)

TOKEN=$(curl -sf http://metadata.google.internal/computeMetadata/v1/instance/service-accounts/default/token \
  -H Metadata-Flavor:Google | python3 -c "import sys,json; print(json.load(sys.stdin)['access_token'])")

echo "$TOKEN" | docker --config "$DOCKERCFG" login -u oauth2accesstoken --password-stdin \
  https://asia-northeast1-docker.pkg.dev

docker --config "$DOCKERCFG" pull \
  asia-northeast1-docker.pkg.dev/${PROJECT_ID}/cloud-run-source-deploy/sakurahoumon-ocr/sakurahoumon-ocr:gpu

docker rm -f ocr || true

ANTHROPIC_API_KEY=$(curl -sf \
  "https://secretmanager.googleapis.com/v1/projects/${PROJECT_ID}/secrets/anthropic-api-key/versions/latest:access" \
  -H "Authorization: Bearer $TOKEN" | \
  python3 -c "import sys,json,base64; print(base64.b64decode(json.load(sys.stdin)['payload']['data']).decode())")

docker --config "$DOCKERCFG" run -d --name ocr --restart=always \
  --volume /var/lib/nvidia/lib64:/usr/local/nvidia/lib64 \
  --volume /var/lib/nvidia/bin:/usr/local/nvidia/bin \
  --device /dev/nvidia0:/dev/nvidia0 \
  --device /dev/nvidia-uvm:/dev/nvidia-uvm \
  --device /dev/nvidiactl:/dev/nvidiactl \
  -e LD_LIBRARY_PATH=/usr/local/nvidia/lib64 \
  -e ANTHROPIC_API_KEY="$ANTHROPIC_API_KEY" \
  -p 8080:8080 \
  asia-northeast1-docker.pkg.dev/${PROJECT_ID}/cloud-run-source-deploy/sakurahoumon-ocr/sakurahoumon-ocr:gpu

rm -rf "$DOCKERCFG"
echo "Deploy complete!"
