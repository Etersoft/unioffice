#!/bin/sh
# Run only in a disposable Wine prefix, inside a graphical session or xvfb-run.
set -eu
: "${WINEPREFIX:?Set WINEPREFIX to the dedicated OpenOffice test prefix}"
wine_loader=${WINE:-wine}
office_exe=${OFFICE_EXE:-'C:\Program Files (x86)\OpenOffice 4\program\soffice.exe'}
probe_script=${1:-openoffice-backend.vbs}
pass_marker=${2:-OpenOffice COM backend: PASS}
case "${CLIENT_BITS:-32}" in
    32) script_host='C:\windows\syswow64\cscript.exe' ;;
    64) script_host='C:\windows\system32\cscript.exe' ;;
    *) echo 'CLIENT_BITS must be 32 or 64' >&2; exit 2 ;;
esac
test_root=$(CDPATH= cd -- "$(dirname -- "$0")/.." && pwd)
"$wine_loader" "$office_exe" -headless -nofirststartwizard -norestore -nodefault &
office_launcher=$!
probe_log=$(mktemp "$HOME/tmp/claude/unioffice-backend-XXXXXX.log")
trap 'kill "$office_launcher" 2>/dev/null || true' EXIT HUP INT TERM
sleep 3
probe_status=0
timeout -k 3 60 "$wine_loader" "$script_host" //nologo "Z:$test_root/tests/openoffice-ready.vbs" > "$probe_log" 2>&1 || probe_status=$?
cat "$probe_log"
[ "$probe_status" -eq 0 ] || exit "$probe_status"
grep -Fq 'OpenOffice readiness: PASS' "$probe_log" || exit 1
case "$probe_script" in
    *.exe) set -- "$wine_loader" "Z:$test_root/$probe_script" ;;
    *) set -- "$wine_loader" "$script_host" //nologo "Z:$test_root/tests/$probe_script" ;;
esac
timeout -k 3 60 "$@" > "$probe_log" 2>&1 || probe_status=$?
cat "$probe_log"
[ "$probe_status" -eq 0 ] || exit "$probe_status"
# Wine cscript can return zero after a VBScript runtime error.
grep -Fq "$pass_marker" "$probe_log"
