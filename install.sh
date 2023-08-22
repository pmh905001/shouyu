#!/bin/bash

source venv/Scripts/activate
pyinstaller -wF -p venv/Scripts -p venv/Lib/site-packages --add-data "./fish.jpg:." -i fish.ico -n 授渔 main.py
