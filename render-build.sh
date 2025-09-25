#!/usr/bin/env bash
set -euo pipefail

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
