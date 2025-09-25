#!/usr/bin/env bash
set -euo pipefail

pip install -r requirements.txt

cd dashboard
npm install
npm run build
npm run export -- --outdir ../dashboard_static
