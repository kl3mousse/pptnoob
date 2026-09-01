#!/bin/bash

set -euo pipefail

ROOT_DIR="$(cd "$(dirname "$0")/.." && pwd)"
VERSION="${1:-$(node -p "require('${ROOT_DIR}/package.json').version")}"
VERSION="${VERSION#v}"
BUILD_DIR="${ROOT_DIR}/release/build"
PAYLOAD_DIR="${BUILD_DIR}/payload"
OUTPUT_DIR="${ROOT_DIR}/release"
PACKAGE_PATH="${OUTPUT_DIR}/pptNoob-${VERSION}-macOS.pkg"

rm -rf "${BUILD_DIR}"
mkdir -p "${PAYLOAD_DIR}/Library/Application Support/pptNoob" "${OUTPUT_DIR}"
cp "${ROOT_DIR}/manifest.xml" "${PAYLOAD_DIR}/Library/Application Support/pptNoob/manifest.xml"

pkgbuild \
  --root "${PAYLOAD_DIR}" \
  --scripts "${ROOT_DIR}/installer/scripts" \
  --identifier "com.pptnoob.powerpoint-addin" \
  --version "${VERSION}" \
  --install-location "/" \
  "${PACKAGE_PATH}"

rm -rf "${BUILD_DIR}"
echo "Created ${PACKAGE_PATH}"