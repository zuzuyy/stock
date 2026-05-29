#!/bin/bash
# 增强版启动脚本：自动修复损坏的虚拟环境

set -e
cd "$(dirname "$0")"

echo "🔍 检查虚拟环境..."
if [ ! -f "myenv/bin/activate" ]; then
    echo "⚠️ 虚拟环境损坏，正在重建 myenv ..."
    rm -rf myenv
    python3 -m venv myenv
fi

source myenv/bin/activate
echo "📦 安装/更新依赖..."
pip install --upgrade pip -q
pip install requests pandas numpy scikit-learn xgboost openpyxl

echo "🚀 开始分析..."
python3 gp.py
echo "✅ 完成！报告在 涨停分析/ 目录"
