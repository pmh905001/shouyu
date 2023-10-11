#!/bin/bash

rm -fr build/
rm -fr dist/
source venv310/Scripts/activate
pyinstaller -wF -p venv310/Scripts -p venv310/Lib/site-packages --add-data "./kb.ini;." --add-data "./resources/icons;./resources/icons" -i ./resources/icons/fish.ico -n 授渔 main.py
