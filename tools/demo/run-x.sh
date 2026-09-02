#!/bin/bash
# Xvfb + openbox を用意してから、指定のスクリプトを実行する
set -e
exec xvfb-run -a --server-args="-screen 0 1920x1080x24 -nolisten tcp" bash -c '
  openbox --sm-disable >/dev/null 2>&1 &
  sleep 1.5
  xsetroot -solid "#1f3a2e" 2>/dev/null || true
  exec "$@"
' _ "$@"
