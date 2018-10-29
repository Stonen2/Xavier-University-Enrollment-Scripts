#!/bin/bash
export DISPLAY=:0
#xdg-open https://www.xavier.edu
#gpicview /home/pi/Desktop/QRCode.jpg
pkill /home/pi/startup_vid.sh OMXplayer
eog -f /home/pi/Desktop/QRCode.jpg

