# IELTS Vocabulary Trainer (IELTS 词汇训练营)

基于 [Streamlit](https://streamlit.io/) 的雅思词汇学习网页应用，提供 **28 天系统化学习计划**，覆盖 **3,484 个核心雅思词汇**。搭配短语和测试题目由大模型预生成，支持多模式测试、智能复习和进度追踪。

## 功能概览

| 功能 | 说明 |
|------|------|
| 28 天学习计划 | 3,484 词按原始顺序均匀分配，每天约 125 词 |
| 交互式词汇卡片 | 英文单词 + 音标 + 音频 + 英文释义，可展开查看中文释义 / 搭配 / 例句 / 同根词 / 相近词 |
| 4 种测试题型 | 汉译英、英译汉、短语搭配翻译、句子词汇填空 |
| 智能评分 | Levenshtein 模糊匹配，≥85% 相似度即算正确 |
| 进度持久化 | 学习进度自动保存至本地 JSON，关闭浏览器不丢失 |
| 词汇掌握度追踪 | 每词独立记录尝试次数和正确率，绿色 / 黄色 / 红色三色标识 |
| 5 种复习模式 | 艾宾浩斯间隔重复 / 薄弱词强化 / 随机抽查 / 按天复习 / 掌握度检查 |
| 学习报告 | 统计仪表板 + 每日正确率柱状图 + 测试历史记录 |

## 项目结构

```
collapp/
├── app.py                              # 主应用入口 (Streamlit)
├── config.example.py                   # API 密钥配置模板
├── config.py                           # API 密钥配置 (gitignore)
├── requirements.txt                    # Python 依赖
├── .gitignore
│
├── 雅思英文词汇表（完整版）.xlsx        # 源数据 (3,484 词 × 9 列)
│
├── generate_collocations.py            # 搭配短语生成器 (调用 LLM)
├── generate_test_bank.py               # 测试题库生成器 (调用 LLM)
├── generate_audio.py                   # 批量音频生成脚本
├── audio_generator.py                  # 音频引擎 (edge-tts / Volcengine TTS)
│
├── test_bank_llm/                      # 预生成测试题库
│   ├── day_01.json ... day_28.json
│   └── checkpoint.json
├── audio/                              # MP3 发音文件 (0.mp3 ~ 3483.mp3)
│   └── generation_progress.json
└── progress.json                       # 用户学习进度 (运行时生成)
```

## 环境要求

- Python 3.10+
- Windows / macOS / Linux

## 快速开始

### 1. 克隆仓库

```bash
git clone https://github.com/Fullkon/ielts-voc.git
cd ielts-voc
```

### 2. 创建虚拟环境并安装依赖

```bash
python -m venv venv
venv\Scripts\activate       # Windows
# source venv/bin/activate  # macOS / Linux

pip install -r requirements.txt
```

### 3. 配置 API 密钥

```bash
copy config.example.py config.py    # Windows
# cp config.example.py config.py    # macOS / Linux
```

编辑 `config.py`，填入火山方舟 API 密钥：

```python
ARK_KEY = "your-ark-api-key-here"
ARK_URL = "https://ark.cn-beijing.volces.com/api/v3/chat/completions"
MODEL   = "doubao-seed-2-1-pro-260628"
```

> API 密钥可在[火山方舟控制台](https://console.volcengine.com/ark/)获取。

### 4. （可选）生成搭配短语

为每个词汇生成典型的考试搭配短语：

```bash
python generate_collocations.py
```

- 约 70 批次，每批 50 词，预计耗时约 6 小时
- 支持断点续传，中断后重新运行即可从上次位置继续
- 结果写回 Excel 文件的「常见搭配」列

### 5. （可选）生成测试题库

为 28 天每天生成 4 种题型的测试题：

```bash
python generate_test_bank.py
```

- cn2en / en2cn 直接从 Excel 生成（无需 LLM，秒完成）
- collocation / sentence 调用 LLM 生成，预计约 24 小时
- 输出至 `test_bank_llm/` 目录，支持断点续传

### 6. 生成音频文件

```bash
python generate_audio.py
```

- 使用免费的 Microsoft Edge TTS（`en-US-JennyNeural` 美式女声）
- 3,484 个 MP3，并发度 5，约 15–20 分钟

### 7. 启动应用

```bash
streamlit run app.py
```

浏览器打开 `http://localhost:8501` 即可开始学习。

## 使用指南

### 词汇学习

1. 左侧边栏选择「学习日」（1–28）
2. 每个词汇卡片显示英文单词、音标、发音按钮和英文释义
3. 点击「点击查看」展开该词的完整信息——汉语释义、常见搭配、例句、同根词、相近词
4. 页面底部有翻页导航，可快速前后翻页
5. 页面底部「直接进入测试」区域，选择题型和数量后点击「进入测试」

### 进入测试

1. 选择测试题型——汉译英、英译汉、短语搭配翻译、句子词汇填空
2. 选择测试数量（10 / 20 / 30 / 50 / 全部）
3. 点击「进入测试」，页面底部出现「开始测试」按钮
4. 答题后自动评分并记录

### 复习强化

提供 5 种复习模式：
- **智能复习**：基于艾宾浩斯遗忘曲线推荐需要复习的天数
- **薄弱词强化**：自动筛选掌握度 <60% 的词汇
- **随机抽查**：从已学词汇中随机抽取 20 个
- **按天复习**：选定某一天重新测试
- **掌握度检查**：按掌握度区间（<40% / 40–60% / 60–80% / ≥80%）筛选

### 学习报告

查看整体统计数据——已完成天数、总体正确率、已掌握词数、每日正确率趋势图及测试历史记录。

## 数据流

```
雅思英文词汇表.xlsx
    │
    ├─→ generate_collocations.py ──→ 更新 xlsx「常见搭配」列
    ├─→ generate_test_bank.py   ──→ test_bank_llm/day_*.json
    └─→ generate_audio.py       ──→ audio/*.mp3
                │
                ▼
           app.py (Streamlit 运行)
           读取 xlsx + JSON + MP3
                │
                ▼
         用户浏览器 http://localhost:8501
```

## 依赖说明

| 包 | 用途 |
|----|------|
| `streamlit` | Web 应用框架 |
| `pandas` | Excel 和数据处理 |
| `openpyxl` | Excel 文件读写 |
| `numpy` | 数值计算 |
| `edge-tts` | 免费文字转语音 (Microsoft Edge) |
| `requests` | HTTP API 调用 |

## 常见问题

<details>
<summary><b>Q: 音频没有声音或发音不对应？</b></summary>

确保已运行 `python generate_audio.py` 生成所有 MP3 文件。不要删除或重命名 `audio/` 目录下的文件，文件名 (`0.mp3` ~ `3483.mp3`) 与 Excel 行号一一对应。
</details>

<details>
<summary><b>Q: 页面提示"测试题库尚未生成"？</b></summary>

运行 `python generate_test_bank.py` 生成至少一天（约 27 分钟/天）的题库后即可测试。
</details>

<details>
<summary><b>Q: 如何清空学习进度重新开始？</b></summary>

删除 `progress.json` 文件，或在应用侧边栏点击「重置所有进度」按钮。
</details>

## License

MIT
