#!/usr/bin/env bash
set -euo pipefail

APPIMAGE_URL="https://ftp.gwdg.de/pub/tdf/libreoffice/stable/24.2.5/AppImage/LibreOffice_24.2.5_Linux_x86-64.AppImage"
APPIMAGE_DIR="libreoffice-appimage"
SOFFICE_PATH="${APPIMAGE_DIR}/usr/bin/soffice"

if [ ! -x "${SOFFICE_PATH}" ]; then
  echo "Fetching portable LibreOffice for XLSB conversion..."
  TMP_APPIMAGE="LibreOffice.AppImage"
  curl -L -o "${TMP_APPIMAGE}" "${APPIMAGE_URL}"
  chmod +x "${TMP_APPIMAGE}"
  "./${TMP_APPIMAGE}" --appimage-extract >/dev/null
  rm -rf "${APPIMAGE_DIR}"
  mv squashfs-root "${APPIMAGE_DIR}"
  rm -f "${TMP_APPIMAGE}"
fi

pip install -r requirements.txt

cd dashboard
npm install
npm run build

rm -rf ../dashboard_static
mkdir -p ../dashboard_static
cp -R .next/server/app/* ../dashboard_static/
mkdir -p ../dashboard_static/_next
cp -R .next/static ../dashboard_static/_next/
cp -R public/* ../dashboard_static/ 2>/dev/null || true
