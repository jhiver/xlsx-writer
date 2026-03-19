#!/bin/bash
# Benchmark xlsx-writer with 100,000 rows under different scenarios
set -e

BIN=target/release/xlsx-writer
ROWS=100000
OUTDIR=/tmp/xlsx-bench
mkdir -p "$OUTDIR"

echo "=== xlsx-writer benchmark ($ROWS rows) ==="
echo ""

# --- Benchmark 1: Plain TEXT + NUM (no formatting) ---
echo "1) Plain TEXT + NUM (no formatting)"
{
  echo "ADD_WORKSHEET Plain"
  for i in $(seq 1 $ROWS); do
    echo "FAST $((i-1)) 0 _ $i"
    echo "FAST $((i-1)) 1 _ Item_$i"
    echo "FAST $((i-1)) 2 _ $((RANDOM % 10000)).$((RANDOM % 100))"
  done
} > "$OUTDIR/plain_input.txt"
echo "  Input: $(wc -l < "$OUTDIR/plain_input.txt") lines, $(du -h "$OUTDIR/plain_input.txt" | cut -f1) on disk"
time "$BIN" "$OUTDIR/plain.xlsx" < "$OUTDIR/plain_input.txt" > /dev/null
echo "  Output: $(du -h "$OUTDIR/plain.xlsx" | cut -f1)"
echo ""

# --- Benchmark 2: FAST with named styles ---
echo "2) FAST with named styles (bold header + num_format)"
{
  echo "ADD_WORKSHEET Styled"
  echo "STYLE hdr bold 1 bg_color navy color white"
  echo "STYLE num num_format #,##0.00"
  echo "STYLE txt font Courier"
  echo "FAST 0 0 hdr ID"
  echo "FAST 0 1 hdr Name"
  echo "FAST 0 2 hdr Amount"
  for i in $(seq 1 $ROWS); do
    echo "FAST $i 0 txt $i"
    echo "FAST $i 1 txt Item_$i"
    echo "FAST $i 2 num $((RANDOM % 10000)).$((RANDOM % 100))"
  done
} > "$OUTDIR/styled_input.txt"
echo "  Input: $(wc -l < "$OUTDIR/styled_input.txt") lines, $(du -h "$OUTDIR/styled_input.txt" | cut -f1) on disk"
time "$BIN" "$OUTDIR/styled.xlsx" < "$OUTDIR/styled_input.txt" > /dev/null
echo "  Output: $(du -h "$OUTDIR/styled.xlsx" | cut -f1)"
echo ""

# --- Benchmark 3: ROW command (space-separated) ---
echo "3) ROW command (space-separated, auto type detection)"
{
  echo "ADD_WORKSHEET ROW"
  echo "SETCOL A:A 10"
  echo "SETCOL B:B 20"
  echo "SETCOL C:C 12"
  echo "ROW .bold:1 A1 ID Name Amount"
  for i in $(seq 2 $((ROWS+1))); do
    echo "ROW A$i $((i-1)) Item_$((i-1)) $((RANDOM % 10000)).$((RANDOM % 100))"
  done
} > "$OUTDIR/row_input.txt"
echo "  Input: $(wc -l < "$OUTDIR/row_input.txt") lines, $(du -h "$OUTDIR/row_input.txt" | cut -f1) on disk"
time "$BIN" "$OUTDIR/row.xlsx" < "$OUTDIR/row_input.txt" > /dev/null
echo "  Output: $(du -h "$OUTDIR/row.xlsx" | cut -f1)"
echo ""

# --- Benchmark 4: Inline styles on every cell ---
echo "4) Inline styles on every cell (worst case formatting)"
{
  echo "ADD_WORKSHEET HeavyFormat"
  for i in $(seq 1 $ROWS); do
    echo "TEXT .bold:1 .color:navy .border:1 .font_size:10 A$i Row $i"
    echo "NUM .num_format:#,##0.00 .bg_color:#EEEEEE .border:1 B$i $((RANDOM % 10000)).$((RANDOM % 100))"
    echo "NUM .num_format:0.0% .italic:1 .border:1 C$i 0.$((RANDOM % 100))"
  done
} > "$OUTDIR/heavy_input.txt"
echo "  Input: $(wc -l < "$OUTDIR/heavy_input.txt") lines, $(du -h "$OUTDIR/heavy_input.txt" | cut -f1) on disk"
time "$BIN" "$OUTDIR/heavy.xlsx" < "$OUTDIR/heavy_input.txt" > /dev/null
echo "  Output: $(du -h "$OUTDIR/heavy.xlsx" | cut -f1)"
echo ""

# --- Benchmark 5: Mixed commands (realistic workload) ---
echo "5) Mixed commands (realistic: header + formulas + freeze + autofilter)"
{
  echo "ADD_WORKSHEET Report"
  echo "SET_PROPERTY title Benchmark Report"
  echo "SETCOL A:A 10"
  echo "SETCOL B:B 20"
  echo "SETCOL C:C 15"
  echo "SETCOL D:D 15"
  echo "FREEZE A2"
  echo "TEXT .bold:1 .bg_color:navy .color:white A1 ID"
  echo "TEXT .bold:1 .bg_color:navy .color:white B1 Product"
  echo "TEXT .bold:1 .bg_color:navy .color:white C1 Price"
  echo "TEXT .bold:1 .bg_color:navy .color:white D1 Tax"
  for i in $(seq 2 $((ROWS+1))); do
    echo "NUM A$i $((i-1))"
    echo "TEXT B$i Product_$((i-1))"
    echo "NUM .num_format:#,##0.00 C$i $((RANDOM % 10000)).$((RANDOM % 100))"
    echo "FORMULA .num_format:#,##0.00 D$i =C$i*0.2"
  done
  echo "AUTOFILTER A1:D$((ROWS+1))"
} > "$OUTDIR/mixed_input.txt"
echo "  Input: $(wc -l < "$OUTDIR/mixed_input.txt") lines, $(du -h "$OUTDIR/mixed_input.txt" | cut -f1) on disk"
time "$BIN" "$OUTDIR/mixed.xlsx" < "$OUTDIR/mixed_input.txt" > /dev/null
echo "  Output: $(du -h "$OUTDIR/mixed.xlsx" | cut -f1)"
echo ""

# --- Summary ---
echo "=== Output file sizes ==="
ls -lh "$OUTDIR"/*.xlsx | awk '{print "  " $NF ": " $5}'

# Cleanup temp input files
rm -f "$OUTDIR"/*_input.txt
