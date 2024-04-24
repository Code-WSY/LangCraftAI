#!/bin/bash
echo "Welcome to the translation tool! This tool is developed by Suyun."
echo "Please select the function you want to use:"
echo "[1] Translate"
echo "[2] PDF to TXT"
echo

read -p "Enter your choice: " choice

case $choice in
    1)
        cd /Users/suyun/OneDrive/Project/translate/src
        /usr/local/bin/python GUI.py
        ;;
    2)
        cd /Users/suyun/OneDrive/Project/translate/test
        /usr/local/bin/python pdftotxt.py
        ;;
    *)
        echo "输入无效，请输入 1 或 2。"
        ;;
esac
