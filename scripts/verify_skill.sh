#!/usr/bin/env bash
set -euo pipefail

# Anchor the working directory to the skill package.
# 将工作目录固定在技能包根目录。
cd "$(dirname "${BASH_SOURCE[0]}")/.."
bash ./scripts/debug_skill.sh inspect --output json
bash ./scripts/debug_skill.sh list-tools --output content
"${PYTHON_COMMAND:-python3}" -m unittest discover -s tests -v
bash ./scripts/debug_skill.sh validate
bash ./scripts/debug_skill.sh package
