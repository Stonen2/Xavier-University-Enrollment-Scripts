#!/bin/bash


#Export DISPLAY


export DISPLAY =:0.0 

#Call Gnome EOG

/usr/bin/eog -f /home/pi/Desktop/QRCode.jpg &

#Time to display

sleep 100

killall eog


