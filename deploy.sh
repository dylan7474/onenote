#!/usr/bin/env bash

set -euo pipefail

PORT_ARG=${1:-3020}
PROJECT_NAME="OneNote Web"
IMAGE_NAME="onenote-web"
CONTAINER_NAME="onenote-web"
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

if ! [[ "${PORT_ARG}" =~ ^[0-9]+$ ]] || (( PORT_ARG < 1 || PORT_ARG > 65535 )); then
  echo "Error: PORT must be an integer between 1 and 65535."
  exit 1
fi

if ! command -v docker >/dev/null 2>&1; then
  echo "Error: docker is required but was not found in PATH."
  exit 1
fi

if [ ! -f "${SCRIPT_DIR}/index.html" ]; then
  echo "Error: index.html was not found in ${SCRIPT_DIR}."
  exit 1
fi

if [ ! -f "${SCRIPT_DIR}/server.js" ]; then
  echo "Error: server.js was not found in ${SCRIPT_DIR}."
  exit 1
fi

if [ ! -f "${SCRIPT_DIR}/vendor/dompurify.min.js" ]; then
  echo "Error: vendor/ assets were not found in ${SCRIPT_DIR}."
  exit 1
fi

echo "=== Deploying ${PROJECT_NAME} on http://localhost:${PORT_ARG} ==="

BUILD_DIR="$(mktemp -d)"
cleanup() {
  rm -rf "${BUILD_DIR}"
}
trap cleanup EXIT

cp "${SCRIPT_DIR}/index.html" "${SCRIPT_DIR}/server.js" "${SCRIPT_DIR}/graph.js" "${BUILD_DIR}/"
cp -r "${SCRIPT_DIR}/vendor" "${BUILD_DIR}/vendor"

cat > "${BUILD_DIR}/Dockerfile" <<'DOCKER_EOF'
FROM node:22-alpine
WORKDIR /app
COPY index.html server.js graph.js ./
COPY vendor ./vendor
ENV DATA_FILE=/data/state.json
VOLUME ["/data"]
CMD ["node", "server.js"]
DOCKER_EOF

echo "Building Docker image..."
docker build -t "${IMAGE_NAME}" "${BUILD_DIR}"

echo "Stopping existing container (if any)..."
docker rm -f "${CONTAINER_NAME}" >/dev/null 2>&1 || true

echo "Starting container..."
docker run -d \
  --name "${CONTAINER_NAME}" \
  -p "${PORT_ARG}:${PORT_ARG}" \
  -e "PORT=${PORT_ARG}" \
  -v "onenote-data:/data" \
  --restart unless-stopped \
  "${IMAGE_NAME}" >/dev/null

echo "========================================="
echo "Deployed ${PROJECT_NAME}."
echo "URL: http://localhost:${PORT_ARG}/"
echo "App file: http://localhost:${PORT_ARG}/index.html"
echo "========================================="
