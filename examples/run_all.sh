#!/bin/bash
# Generate all example spreadsheets
# Usage: cd xlsx-writer && bash examples/run_all.sh

set -e

BINARY="${BINARY:-cargo run --release --}"
OUTDIR="${OUTDIR:-examples/output}"

mkdir -p "$OUTDIR"

echo "Building xlsx-writer..."
cargo build --release

BINARY="target/release/xlsx-writer"

for input in examples/*.txt; do
    name=$(basename "$input" .txt)
    output="$OUTDIR/${name}.xlsx"
    echo "  $input → $output"
    "$BINARY" "$output" < "$input"
done

echo ""
echo "Done! Generated files in $OUTDIR/"
ls -lh "$OUTDIR"/*.xlsx
