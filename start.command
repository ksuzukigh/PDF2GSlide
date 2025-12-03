#!/bin/bash

# 1. このファイルがある場所に移動する
cd "$(dirname "$0")"

# 2. Anacondaの機能を読み込む（これが重要！）
source /opt/homebrew/anaconda3/etc/profile.d/conda.sh

# 3. pdf-slide 環境に切り替える
conda activate pdf-slide

# 4. Webアプリを起動する
python -m streamlit run app.py