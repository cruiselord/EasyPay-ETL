#!/bin/bash
# install_launchd.sh — schedule the NIBSS ETL to run daily on macOS.
#
# Creates and loads a LaunchAgent that fires auto_run.sh:
#   - daily at NIBSS_RUN_AFTER (default 07:00)
#   - on login/wake (RunAtLoad)
#   - every 15 min as a catch-up (idempotent — a marker in logs/ prevents re-runs)
#
# Usage:
#   ./install_launchd.sh                # default 07:00
#   NIBSS_RUN_AFTER=06:00 ./install_launchd.sh
#
# NOTE: macOS blocks launchd from reading ~/Documents. If this project lives
# under ~/Documents, grant Full Disk Access to /bin/bash first:
#   System Settings → Privacy & Security → Full Disk Access → add /bin/bash
set -e

DIR="$(cd "$(dirname "$0")" && pwd)"
# Reverse-DNS label for the LaunchAgent.  Replace `yourorg` with your own
# organisation short-name — this is the identifier that shows up under
# ~/Library/LaunchAgents/ and in `launchctl list`.  Override with NIBSS_LABEL.
LABEL="${NIBSS_LABEL:-com.yourorg.nibss-etl}"
PLIST="$HOME/Library/LaunchAgents/$LABEL.plist"
RUN_AFTER="${NIBSS_RUN_AFTER:-07:00}"
HOUR="${RUN_AFTER%%:*}"
MIN="${RUN_AFTER##*:}"

mkdir -p "$DIR/logs"

cat > "$PLIST" <<EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key><string>$LABEL</string>
    <key>ProgramArguments</key>
    <array><string>/bin/bash</string><string>$DIR/auto_run.sh</string></array>
    <key>WorkingDirectory</key><string>$DIR</string>
    <key>RunAtLoad</key><true/>
    <key>StartCalendarInterval</key>
    <dict><key>Hour</key><integer>$HOUR</integer><key>Minute</key><integer>$MIN</integer></dict>
    <key>StartInterval</key><integer>900</integer>
    <key>StandardOutPath</key><string>$DIR/logs/auto_run.log</string>
    <key>StandardErrorPath</key><string>$DIR/logs/auto_run.log</string>
</dict>
</plist>
EOF

launchctl unload "$PLIST" 2>/dev/null || true
launchctl load -w "$PLIST"
echo "Installed '$LABEL' — runs daily at $RUN_AFTER (and on login/wake)."
echo "To remove: launchctl unload $PLIST && rm $PLIST"

# Verify launchd can actually read the project (catches the missing
# Full-Disk-Access case, where /bin/bash is denied ~/Documents).
launchctl kickstart -k "gui/$UID/$LABEL" 2>/dev/null || true
sleep 2
if grep -q "triggered" "$DIR/logs/auto_run.log" 2>/dev/null; then
    echo "OK: launchd successfully executed the job."
else
    echo "WARNING: no 'triggered' entry in logs/auto_run.log."
    echo "  Grant Full Disk Access to /bin/bash, then rerun this script:"
    echo "  System Settings → Privacy & Security → Full Disk Access → add /bin/bash"
fi
