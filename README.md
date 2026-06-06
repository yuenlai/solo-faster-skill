# solo-faster-skill

`solo-faster` 是一个面向 Codex 的本地 skill，用于在单个代码仓库内批量生成主提示词，并配合 Trae、Excel、git 完成串行投递、人工等待、验收、回填和逐轮提交推送。

## 安装

将本仓库 clone 到本地后，把整个目录放到你的 Codex skills 目录下，并命名为 `solo-faster`：

```bash
git clone https://github.com/yuenlai/solo-faster-skill.git
mkdir -p "$CODEX_HOME/skills"
rm -rf "$CODEX_HOME/skills/solo-faster"
cp -R ./solo-faster-skill "$CODEX_HOME/skills/solo-faster"
```

如果你的环境没有预先设置 `CODEX_HOME`，通常可以先这样设置：

```bash
export CODEX_HOME="$HOME/.codex"
mkdir -p "$CODEX_HOME/skills"
```

安装完成后，可在 Codex 中直接使用：

```text
使用 $solo-faster 处理当前项目
```

## 目录说明

- `SKILL.md`：主说明文档
- `scripts/`：批量生成、Excel 维护、Trae 辅助与状态流转脚本
- `references/`：不同任务类型的执行手册
- `agents/openai.yaml`：skill agent 元数据

## 使用前提

这个 skill 依赖本地环境能力，至少包括：

- Python 3
- git
- macOS
- Trae CN
- Codex 的 `@电脑` 能力
- Codex 内置浏览器 `@browser`

其中，Trae 相关发送、Session ID 复制和窗口切换逻辑是围绕本地 Trae CN 与 macOS 自动化设计的。

## 本地运行态数据

下列内容是本地运行态数据，不应提交到仓库：

- `data/prompt_history.jsonl`
- `data/trae_open_state.json`
- `__pycache__/`
- `.DS_Store`

仓库里的 `data/` 目录可以为空；脚本运行时会按需创建本地数据文件。

## 自检

```bash
python3 -m py_compile "$CODEX_HOME/skills/solo-faster/scripts/"*.py
python3 "$CODEX_HOME/skills/solo-faster/scripts/solo_faster_v2_dry_run.py"
```
