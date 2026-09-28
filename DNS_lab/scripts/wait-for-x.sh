#!/bin/sh
set -eu

until /usr/bin/xdpyinfo -display "${DISPLAY:-:0}" >/dev/null 2>&1; do
    sleep 0.2
done

exec "$@"
