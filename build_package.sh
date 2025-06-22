#!/usr/bin/env bash
set -euo pipefail
mkdir -p package
zip -r package/MagicWand_v0.3-dev.zip README.md CHANGELOG.md changelog.txt FAQ.md REFERENCE.md ROADMAP.md forms modules
