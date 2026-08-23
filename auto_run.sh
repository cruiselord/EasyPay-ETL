#!/bin/bash
# auto_run.sh — daily NIBSS reconciliation guard.
#
# Fires the fetch + ETL for YESTERDAY's data, but only when:
#   1. the clock is at/after the agreed time (default 07:00)
#   2. the Mac is in an active session (lid open AND display awake) — never
#      during clamshell sleep or a DarkWake/maintenance wake
#   3. NIBSS is reachable (i.e. the Mac has internet)
#   4. yesterday's report has not already been produced (marker file)
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
NOTIFIER="/opt/homebrew/bin/terminal-notifier"
[ -x "$NOTIFIER" ] || NOTIFIER="/usr/local/bin/terminal-notifier"
notify() {
    if [ -x "$NOTIFIER" ]; then
        "$NOTIFIER" -title "NIBSS ETL" -message "$1" >/dev/null 2>&1
    else
        osascript -e "display notification \"$1\" with title \"NIBSS ETL\"" 2>/dev/null
    fi
}

log "triggered (report date $REPORT_DATE)"

# Already produced this report?
if [ -f "$MARKER" ]; then
    log "already done — marker present, skipping"
    exit 0
fi

# Before the agreed time?
NOW="$(date '+%H:%M')"
if [[ "$NOW" < "$AGREED_TIME" ]]; then
    log "before agreed time $AGREED_TIME — skipping"
    exit 0
fi

# Active session?  Skip clamshell / sleep / DarkWake — the lid must be open
# (AppleClamshellState "No") and the display must be awake.
if ! ioreg -r -k AppleClamshellState -d 1 2>/dev/null | grep -q '"AppleClamshellState" = No'; then
    log "skipping — lid closed (clamshell); will retry when you log in"
    exit 0
fi
if ! pmset -g assertions 2>/dev/null | grep -q 'Prevent sleep while display is on'; then
    log "skipping — display asleep; will retry when you log in"
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
    notify "Report for $REPORT_DATE produced and emailed."
else
    log "run failed (exit $?) — will retry on next trigger"
    notify "NIBSS ETL failed for $REPORT_DATE — will retry."
fi
