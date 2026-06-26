#!/usr/bin/env bash
#
# Process files in $process_dir that are newer than the timestamp stored
# in timeprocessed.txt, then update timeprocessed.txt to "now".

set -u

TIMESTAMP_FILE="timeprocessed.txt"
process_dir="liudata/"   # set/override as needed
RUNNER="python3 process-call-data.py"

# --- 1. Sanity checks ---
if [[ ! -d "$process_dir" ]]; then
    echo "Error: process_dir does not exist: $process_dir" >&2
    exit 1
fi

# --- 2. Capture "now" BEFORE processing ---
# We snapshot the current time up front so any file created while this
# script is running will be picked up on the *next* run, not skipped.
now_ts="$(date '+%a %b %e %H:%M:%S %Z %Y')"

# --- 3. Read the previous timestamp (if any) ---
if [[ -f "$TIMESTAMP_FILE" ]]; then
    last_ts="$(< "$TIMESTAMP_FILE")"
    last_ts="${last_ts//$'\n'/}"   # strip stray newlines
    last_ts="${last_ts##[[:space:]]}"
    last_ts="${last_ts%%[[:space:]]}"
else
    last_ts=""
fi

# --- 4. Build the find command ---
# If there's no prior timestamp, process everything in the directory.
if [[ -n "$last_ts" ]]; then
    echo "Looking for files in '$process_dir' newer than: $last_ts"
    find_args=( "$process_dir" -maxdepth 1 -type f -name "*.xlsx" -newermt "$last_ts" -print0 )
else
    echo "No previous timestamp found; processing all files in '$process_dir'"
    find_args=( "$process_dir" -maxdepth 1 -type f -name "*.xlsx" -print0 )
fi

# --- 5. Process each matching file ---
count=0
failed=0
while IFS= read -r -d '' file; do
    echo ">>> Processing: $file"
    
    if $RUNNER "$file"; then
        count=$((count + 1))
    else
        echo "    !! $RUNNER failed (exit $?) for: $file" >&2
        failed=$((failed + 1))
    fi
done < <(find "${find_args[@]}")

echo "Processed $count file(s); $failed failure(s)."

# --- 6. Update the timestamp file ---
# Only write the snapshot we took at the start, so we don't lose any
# files that may have appeared during the run.
echo "$now_ts" > "$TIMESTAMP_FILE"
echo "Updated $TIMESTAMP_FILE -> $now_ts"