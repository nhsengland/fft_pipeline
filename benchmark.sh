#!/usr/bin/env bash
# FFT Pipeline Performance Benchmark
# Captures timing and success/fail, prints TL;DR summary

set -e

RESULTS=$(mktemp)
START_TIME=$(date +%s)

cleanup() {
    rm -f "$RESULTS"
}
trap cleanup EXIT

log_result() {
    local service="$1"
    local status="$2"
    local time="$3"
    echo "$service|$status|$time" >> "$RESULTS"
}

run_benchmark() {
    local flag="$1"
    local service="$2"
    echo "* Running $service..."
    
    OUTPUT=$( { time uv run -m fft "$flag" 2>&1; } 2>&1 )
    TIME=$( echo "$OUTPUT" | grep "^real" | awk '{print $2}' )
    
    if echo "$OUTPUT" | grep -q "Pipeline completed successfully"; then
        log_result "$service" "PASS" "$TIME"
        echo "  $service: PASS ($TIME)"
    else
        log_result "$service" "FAIL" "$TIME"
        echo "  $service: FAIL ($TIME)"
    fi
    echo ""
}

# Run benchmarks
echo "========================================"
echo "  FFT Pipeline Performance Benchmark"
echo "========================================"
echo ""

run_benchmark "--ip" "Inpatient"
run_benchmark "--ae" "A&E"
run_benchmark "--amb" "Ambulance"

# Calculate duration
END_TIME=$(date +%s)
TOTAL_SECS=$((END_TIME - START_TIME))

# Print TL;DR
echo "========================================"
echo "  TL;DR"
echo "========================================"
echo ""
echo "| Service   | Status | Time |"
echo "|-----------|--------|------|"

while IFS='|' read -r service status time; do
    printf "| %-11s| %-6s | %s |\n" "$service" "$status" "$time"
done < "$RESULTS"

echo ""
echo "Total elapsed: ${TOTAL_SECS}s"

# Summary
PASSED=$(grep "|PASS|" "$RESULTS" | wc -l | tr -d ' ')
FAILED=$(grep "|FAIL|" "$RESULTS" | wc -l | tr -d ' ')
TOTAL=3

if [ "$FAILED" -eq 0 ]; then
    echo ""
    echo "ALL TESTS PASSED ($PASSED/$TOTAL)"
    exit 0
else
    echo ""
    echo "SOME TESTS FAILED: $PASSED passed, $FAILED failed"
    exit 1
fi
