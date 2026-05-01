#!/usr/bin/env bash
# new_case.sh — scaffold a new case folder from templates
# Usage: bash scripts/new_case.sh case_01
set -e

CASE_ID="${1:?Usage: bash scripts/new_case.sh <case_id>}"
REPO_ROOT="$(cd "$(dirname "$0")/.." && pwd)"
CASE_DIR="$REPO_ROOT/cases/$CASE_ID"
SCHEMAS="$REPO_ROOT/schemas"

if [ -d "$CASE_DIR" ]; then
  echo "Error: $CASE_DIR already exists."
  exit 1
fi

mkdir -p "$CASE_DIR/transcripts" "$CASE_DIR/validation"

cp "$SCHEMAS/meta.yaml.template"           "$CASE_DIR/meta.yaml"
cp "$SCHEMAS/placeholder_map.yaml.template" "$CASE_DIR/placeholder_map.yaml"
cp "$SCHEMAS/human_grades.yaml.template"   "$CASE_DIR/human_grades.yaml"

# Patch case_id in meta.yaml
sed -i '' "s/^case_id: case_01/case_id: $CASE_ID/" "$CASE_DIR/meta.yaml"

echo "Created: $CASE_DIR"
echo ""
echo "Next steps:"
echo "  1. Drop transcript files into $CASE_DIR/transcripts/"
echo "  2. Fill in $CASE_DIR/meta.yaml  (criteria_in_scope, product_type, etc.)"
echo "  3. Fill in $CASE_DIR/placeholder_map.yaml"
echo "  4. Run: python scripts/adapt_case.py --case cases/$CASE_ID"
echo "  5. Run LLM on adapted_indicators.xlsx → save results to $CASE_DIR/validation/run_output.json"
echo "  6. Fill in $CASE_DIR/human_grades.yaml"
echo "  7. Run: python scripts/validate_case.py --case cases/$CASE_ID"
echo "  8. Annotate $CASE_DIR/validation/diff_report.yaml (error_types, fix_applied, etc.)"
