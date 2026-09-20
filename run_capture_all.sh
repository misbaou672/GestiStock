#!/bin/bash
export DISPLAY=:99
Xvfb :99 -screen 0 1600x1000x24 &
XVFB_PID=$!
sleep 2

python3 capture_all.py

kill $XVFB_PID

