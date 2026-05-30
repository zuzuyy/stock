#!/bin/bash

set -e
cd "$(dirname "$0")"

echo "🔍 检查虚拟环境..."
if [ ! -f "myenv/bin/activate" ]; then
    echo "⚠️ 检查虚拟环境，正在重建 myenv ..."
    rm -rf myenv
    python3 -m venv myenv
fi

source myenv/bin/activate
echo "📦 安装/更新依赖..."
pip install --upgrade pip -q
pip install requests pandas numpy scikit-learn xgboost openpyxl img2table[tesseract] paddle easyocr

echo "🚀 开始分析..."
python3 gp.py
echo "✅ 完成！报告在 涨停分析/ 目录"
