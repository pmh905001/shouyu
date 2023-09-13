#!/bin/bash

rm -fr build/
rm -fr dist/
source venv310/Scripts/activate
pyinstaller -wF -p venv310/Scripts -p venv310/Lib/site-packages --add-data "./fish.jpg:." --add-data "./kb.ini:." --add-data "./failed.png:." --add-data "./success.png:." -i fish.ico -n 授渔 main.py
