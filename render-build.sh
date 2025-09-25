#!/usr/bin/env bash
set -euo pipefail

pip install -r requirements.txt

# Build the Next.js dashboard into a static directory the Flask app can serve.
cd dashboard
npm install
npm run build
npm run export -- --outdir ../dashboard_static
