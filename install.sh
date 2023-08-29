#!/bin/bash

rm -fr build/
rm -fr dist/
source venv/Scripts/activate
pyinstaller -wF -p venv/Scripts -p venv/Lib/site-packages --add-data "./fish.jpg:." --add-data "./kb.ini:." -i fish.ico -n 授渔 main.py
