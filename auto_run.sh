#!/bin/bash
# auto_run.sh — daily NIBSS reconciliation guard.
#
# Fires the fetch + ETL for YESTERDAY's data, but only when:
#   1. the clock is at/after the agreed time (default 07:00)
#   2. NIBSS is reachable (i.e. the Mac has internet)
#   3. yesterday's report has not already been produced (marker file)
#
# Designed to be triggered repeatedly (launchd: StartCalendarInterval +
# RunAtLoad + StartInterval); it exits cheaply when there is nothing to do.
set -u

DIR="$(cd "$(dirname "$0")" && pwd)"
AGREED_TIME="${NIBSS_RUN_AFTER:-07:00}"
LOG="$DIR/logs/auto_run.log"
REPORT_DATE="$(date -v-1d '+%d_%m_%Y')"   # report covers the previous day
MARKER="$DIR/logs/.ran_$REPORT_DATE"

mkdir -p "$DIR/logs"
log() { printf '%s %s\n' "$(date '+%Y-%m-%d %H:%M:%S')" "$*" >> "$LOG"; }

# Already produced this report?
[ -f "$MARKER" ] && exit 0

# Before the agreed time?
NOW="$(date '+%H:%M')"
if [[ "$NOW" < "$AGREED_TIME" ]]; then
    exit 0
fi

# Internet / NIBSS reachable?  (retry briefly — the server is occasionally flaky)
up=0
for _ in 1 2 3; do
    if curl -sS --max-time 10 -o /dev/null \
        "https://nibsswebserver.nibss-plc.com.ng/ThinClient/WTM/public/"; then
        up=1
        break
    fi
    sleep 2
done
if [ "$up" -ne 1 ]; then
    log "NIBSS unreachable — skipping (will retry on next trigger)"
    exit 0
fi

# Run the pipeline for yesterday's data.
log "starting run for $REPORT_DATE"
if "$DIR/run.sh" "$REPORT_DATE" >> "$LOG" 2>&1; then
    touch "$MARKER"
    log "done — report for $REPORT_DATE produced"
else
    log "run failed (exit $?) — will retry on next trigger"
fi
