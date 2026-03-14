#!/usr/bin/env bash
set -euo pipefail

if [ "$#" -lt 1 ]; then
  echo "Nutzung: scripts/add_prompt.sh \"dein Prompt-Text\""
  exit 1
fi

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
PROMPTS_FILE="$REPO_ROOT/prompts.txt"

prompt_text="$*"

if [ ! -f "$PROMPTS_FILE" ]; then
  touch "$PROMPTS_FILE"
fi

last_num="$(grep -E '^[0-9]+\.' "$PROMPTS_FILE" | tail -n 1 | sed -E 's/^([0-9]+)\..*/\1/' || true)"
if [ -z "$last_num" ]; then
  next_num=1
else
  next_num=$((last_num + 1))
fi

printf "\n%s. %s\n" "$next_num" "$prompt_text" >> "$PROMPTS_FILE"

git -C "$REPO_ROOT" add "$PROMPTS_FILE"
echo "Prompt #$next_num zu $PROMPTS_FILE hinzugefügt und gestaged."
