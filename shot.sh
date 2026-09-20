#!/bin/bash
export DISPLAY=:99
Xvfb :99 -screen 0 1024x768x24 &
XVFB_PID=$!
sleep 2
python3 gestion_stocks.py &
APP_PID=$!
sleep 4
import -window root /home/msb/.gemini/antigravity/scratch/portfolio-react/src/components/projects/media/gestistock.jpg
kill $APP_PID
kill $XVFB_PID
