#!/bin/bash
# macOS / Linux 用：chmod +x 一键转换.sh 后双击或在终端运行
cd "$(dirname "$0")" || exit 1
export PYTHONIOENCODING=utf-8

echo
echo "============================================"
echo "  docx 转 Dify 知识库 Markdown"
echo "============================================"
echo

if ! command -v python3 >/dev/null 2>&1; then
    echo "[错误] 没有检测到 python3，请先安装 Python 3。"
    exit 1
fi

if ! python3 -c "import docx" >/dev/null 2>&1; then
    echo "[提示] 首次运行需要安装依赖，正在安装 python-docx ..."
    python3 -m pip install -r "./requirements.txt" --quiet || {
        echo "[提示] 镜像源安装失败，尝试默认源 ..."
        python3 -m pip install -r "./requirements.txt" --quiet || {
            echo "[错误] 依赖安装失败，请手动执行：python3 -m pip install -r requirements.txt"
            exit 1
        }
    }
    echo
fi

echo "请输入存放 docx 的文件夹路径（直接回车 = 本文件夹）："
read -r SRCDIR

if [ -z "$SRCDIR" ]; then
    python3 ./docx2md.py .
else
    python3 ./docx2md.py "$SRCDIR"
fi

echo
echo "处理完成，请查看上方日志和输出目录中的「处理清单.csv」。"
echo "更多选项（--flat / -o / --log-file 等）见 README.md：python3 docx2md.py --help"
