#!/bin/bash

rm -fr build/
rm -fr dist/
source venv310/Scripts/activate
# --collect-data rapidocr_onnxruntime: RapidOCR ships its default OCR models
# and config.yaml as package *data*, not code - a plain import scan (which is
# all --add-data-less PyInstaller runs do) never picks those up, so the OCR
# capture hotkey fails at RapidOCR() init with a FileNotFoundError pointing
# at .../_MEIxxxxx/rapidocr_onnxruntime/config.yaml. Don't drop this flag.
pyinstaller -wF -p venv310/Scripts -p venv310/Lib/site-packages --add-data "./kb.ini;." --add-data "./resources/icons;./resources/icons" --collect-data rapidocr_onnxruntime -i ./resources/icons/fish.png -n shouyu main.py
