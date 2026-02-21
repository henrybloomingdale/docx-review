#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"
BUILD_DIR="$PROJECT_DIR/build"
IMAGE="docx-review"
PASSED=0
FAILED=0

run_docx() {
    docker run --rm -v "$PROJECT_DIR:/work" -w /work "$IMAGE" "$@" 2>&1
}

run_docx_stdin() {
    docker run --rm -i -v "$PROJECT_DIR:/work" -w /work "$IMAGE" "$@" 2>&1
}

assert_pass() {
    local name="$1"
    echo "  [PASS] $name"
    PASSED=$((PASSED + 1))
}

assert_fail() {
    local name="$1"
    local detail="${2:-}"
    echo "  [FAIL] $name"
    [[ -n "$detail" ]] && echo "    $detail"
    FAILED=$((FAILED + 1))
}

if ! command -v jq >/dev/null 2>&1; then
    echo "Error: jq is required for this test"
    exit 1
fi

mkdir -p "$BUILD_DIR"

echo "=== Comment Update Tests ==="

if ! docker image inspect "$IMAGE" >/dev/null 2>&1; then
    echo "Building docker image..."
    (cd "$PROJECT_DIR" && make docker >/dev/null)
fi

# fresh base
run_docx --create -o /work/build/comment_update_base.docx --json >/dev/null

# Test 1: Backward-compatible add comment still works
RESULT=$(cat <<'JSON' | run_docx_stdin /work/build/comment_update_base.docx -o /work/build/comment_update_t1.docx --json
{
  "author": "Tester",
  "comments": [
    {"anchor": "Specific Aims", "text": "Original reviewer comment"}
  ]
}
JSON
)

SUCCESS=$(echo "$RESULT" | jq -r '.success')
CTYPE=$(echo "$RESULT" | jq -r '.results[0].type')
if [[ "$SUCCESS" == "true" && "$CTYPE" == "comment" ]]; then
    assert_pass "add comment remains backward-compatible"
else
    assert_fail "add comment failed" "success=$SUCCESS, type=$CTYPE"
fi

# Test 2: Update existing comment by ID
RESULT=$(cat <<'JSON' | run_docx_stdin /work/build/comment_update_t1.docx -o /work/build/comment_update_t2.docx --json
{
  "author": "DocRevise",
  "comments": [
    {
      "op": "update",
      "id": 0,
      "text": "Original reviewer comment\n\nDocRevise action: tightened efficacy claim language in Section 4.2."
    }
  ]
}
JSON
)

SUCCESS=$(echo "$RESULT" | jq -r '.success')
CTYPE=$(echo "$RESULT" | jq -r '.results[0].type')
if [[ "$SUCCESS" == "true" && "$CTYPE" == "comment_update" ]]; then
    READ=$(run_docx /work/build/comment_update_t2.docx --read --json)
    UPDATED_TEXT=$(echo "$READ" | jq -r '.comments[] | select(.id == "0") | .text')
    if [[ "$UPDATED_TEXT" == *"DocRevise action:"* ]]; then
        assert_pass "update comment by id persists in read-back"
    else
        assert_fail "update comment text not found in read-back" "$UPDATED_TEXT"
    fi
else
    assert_fail "update comment failed" "success=$SUCCESS, type=$CTYPE"
fi

# Test 3: Invalid update ID gives structured failure
RESULT=$(cat <<'JSON' | run_docx_stdin /work/build/comment_update_t2.docx -o /work/build/comment_update_t3.docx --json || true
{
  "comments": [
    {"op": "update", "id": 999, "text": "no-op"}
  ]
}
JSON
)

SUCCESS=$(echo "$RESULT" | jq -r '.success')
MSG=$(echo "$RESULT" | jq -r '.results[0].message')
if [[ "$SUCCESS" == "false" && "$MSG" == *"Comment id not found"* ]]; then
    assert_pass "invalid comment id returns clear failure"
else
    assert_fail "invalid id did not fail as expected" "success=$SUCCESS, msg=$MSG"
fi

# Test 4: Dry-run validates update target existence
RESULT=$(cat <<'JSON' | run_docx_stdin /work/build/comment_update_t2.docx --dry-run --json || true
{
  "comments": [
    {"op": "update", "id": 0, "text": "x"},
    {"op": "update", "id": 999, "text": "x"}
  ]
}
JSON
)

OK_COUNT=$(echo "$RESULT" | jq -r '.comments_succeeded')
if [[ "$OK_COUNT" == "1" ]]; then
    assert_pass "dry-run validates existing vs missing ids"
else
    assert_fail "dry-run validation mismatch" "comments_succeeded=$OK_COUNT"
fi

rm -f "$BUILD_DIR"/comment_update_base.docx "$BUILD_DIR"/comment_update_t1.docx "$BUILD_DIR"/comment_update_t2.docx "$BUILD_DIR"/comment_update_t3.docx

echo "=== Comment Update Tests: $PASSED passed, $FAILED failed ==="
if [[ $FAILED -gt 0 ]]; then
    exit 1
fi
