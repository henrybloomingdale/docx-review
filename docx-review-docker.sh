#!/bin/bash
# docx-review — Shell wrapper that runs the Docker container
# Mounts current directory as /work so file paths work naturally.
#
# Usage:
#   docx-review input.docx edits.json -o reviewed.docx
#   cat edits.json | docx-review input.docx -o reviewed.docx
#
set -euo pipefail

IMAGE="docx-review"

# Check if stdin is piped
if [ -t 0 ]; then
    docker run --rm -v "$(pwd):/work" -w /work "$IMAGE" "$@"
else
    docker run --rm -i -v "$(pwd):/work" -w /work "$IMAGE" "$@"
fi
