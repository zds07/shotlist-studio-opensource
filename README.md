# shotlist-studio · 短视频分镜脚本转换器

把甲方原始稿（docx / xlsx / 自由文本）一键转换为标准化的短视频分镜拍摄脚本 xlsx。

![Python](https://img.shields.io/badge/Python-3.9%2B-blue)
![License](https://img.shields.io/badge/License-MIT-green)
![Platform](https://img.shields.io/badge/Platform-macOS%20%7C%20Linux%20%7C%20Windows-lightgrey)

## 解决什么问题

影像制作者经常收到两类稿件：

- **甲方稿**：只有"主题 + 风格 + 时长 + 几段故事"，没有镜头概念，颗粒度粗
- **制作稿**：需要固定 10 列结构（镜号 / 拍摄地点 / 画面内容 / 镜头设计 / 台词·旁白 / 时长 / 备注 / 道具 / 音效 / 参考画面），能直接交给摄制组、演员、甲方三方对稿

本 skill 把前者自动转换成后者。

## 核心特性

- **旁白铁律**：原稿旁白一字不改，按语义分段，每段旁白匹配对应画面
- **固定视觉规范**：微软雅黑、#383838 深灰标题行、#D9EAD3 浅绿幕标题、白/浅灰交替镜头行、全表文字居中
- **打印即用**：横向 A3 纸张、适合一页、水平居中，Ctrl+P 直接打印对稿
- **拆镜规则**：识别幕结构、判断镜头切分粒度、补充镜头设计与道具备注

## 文件结构

```
shotlist-studio/
├── SKILL.md                     # skill 主入口：触发条件、制作流程、判断规则
├── references/
│   ├── format-spec.md           # 格式规范（列结构 / 视觉 / 打印设置）
│   ├── conversion-rules.md      # 甲方稿→制作稿的 11 条转换规则
│   └── examples.md              # 11 个真实脚本的差异对照
├── scripts/
│   ├── build_xlsx.py            # 生成 xlsx 的核心脚本
│   └── parse_brief.py           # 解析原始稿（docx/xlsx）
└── assets/
    └── template_structure.json  # 标准模板的 JSON Schema
```

## 安装

### 方式 1：作为 WorkBuddy / Claude Code / Codex 等 Agent skill 安装

```bash
# 任意 agent 平台通用安装路径：~/.{platform}/skills/shotlist-studio/
git clone https://github.com/zds07/shotlist-studio-opensource.git \
  ~/.workbuddy/skills/shotlist-studio
# 或
git clone https://github.com/zds07/shotlist-studio-opensource.git \
  ~/.claude/skills/shotlist-studio
# 或
git clone https://github.com/zds07/shotlist-studio-opensource.git \
  ~/.codex/skills/shotlist-studio
```

### 方式 2：作为 Python 包本地使用

```bash
git clone https://github.com/zds07/shotlist-studio-opensource.git
cd shotlist-studio-opensource
python3 -m venv .venv
source .venv/bin/activate
pip install openpyxl python-docx
python scripts/build_xlsx.py
```

## 使用示例

### 1. 准备结构化数据

按照 `assets/template_structure.json` 的 Schema 准备一个 dict：

```python
from scripts.build_xlsx import build

meta = {
    "video_name": "示例宣传片",
    "video_topic": "品牌 30s 短片",
    "shoot_date": "2026-11-01",
    "deadline": "2026-11-15",
    "shot_count": 6,
    "props": "演员服装、道具若干",
    "main_product": "示例产品 A",
}

acts = [
    {
        "name": "第一幕：开场",
        "shots": [
            {
                "shot_no": 1,
                "location": "外景·城市天际线",
                "action": "日出航拍，镜头从城市建筑群中升起",
                "camera": "航拍 缓慢上推",
                "dialogue": "清晨的城市正在醒来……",
                "duration": "5s",
                "notes": "光线柔和、转场用",
                "props": "—",
                "sfx": "城市环境音 + 轻音乐",
                "reference": "",
            },
            # ... 更多镜头
        ],
    },
    # ... 更多幕
]

build("output.xlsx", meta, acts)
```

### 2. 解析甲方原始稿

```python
from scripts.parse_brief import parse_brief

data = parse_brief("甲方稿.docx")
print(data["meta"]["video_topic"])
for act in data["acts"]:
    print(f"{act['name']}: {len(act['shots'])} 镜")
```

## 在 Agent 中使用

当你在支持的 Agent（WorkBuddy / Claude Code / Codex / Grok 等）中调用本 skill 时，agent 会自动识别以下触发词：

- 「把这个稿子做成制作版分镜」
- 「按客户经理示例的格式出分镜」
- 「上传甲方稿 → 生成 xlsx 拍摄脚本」

详细触发词清单见 [`SKILL.md`](./SKILL.md)。

## 制作流程（6 步）

1. **预处理原稿**：`.docx` / `.xlsx` 直接读；`.doc` 用 `textutil`（macOS）/ `antiword`（Linux）转文本
2. **提取不可动部分**：旁白原文（**一字不改**）、场景结构、总时长、拍摄点位
3. **按语义拆镜**：每段旁白对应一组画面，做到"旁白 ↔ 画面"一一匹配
4. **遵守三条铁律**：
   - 旁白原话不动
   - 没给参考图 → 参考画面列留空
   - 采访/口播类标"受访者自由发挥，文案仅供参考"
5. **生成统一格式 xlsx**：固定 10 列、视觉规范、打印设置
6. **交付 → 迭代**：用户改 → agent 重跑，秒出新版

## 字段映射表

| 字段名 | 中文列名 | 必填 | 说明 |
|---|---|---|---|
| `shot_no` | 镜号 | ✓ | 数字，从 1 开始递增 |
| `location` | 拍摄地点 | ✓ | 外景/内景+具体地点 |
| `action` | 画面内容 | ✓ | 写"到现场拍什么"，具体动作清单 |
| `camera` | 镜头设计 | ✓ | 景别+运动方式 |
| `dialogue` | 台词/旁白 | ✓ | 采访原话 / 旁白原文（不改） |
| `duration` | 时长 | ✓ | 单位秒，如 "5s" |
| `notes` | 备注 |  | 合规要点、提示、特殊说明 |
| `props` | 道具 |  | 服装、物品清单 |
| `sfx` | 音效 |  | 现场音 / 配乐描述 |
| `reference` | 参考画面 |  | 没给参考图时留空 |

## 安装

把仓库克隆到 `~/.workbuddy/skills/shotlist-studio/` 即可被 WorkBuddy / Claude Code / Codex / Grok 等 Agent 自动识别：

```bash
git clone https://github.com/zds07/shotlist-studio-opensource.git \
  ~/.workbuddy/skills/shotlist-studio
```

## 使用

把甲方原始稿（`.doc` / `.docx` / `.md` / 自由文本）丢给 Agent，让它按 SKILL.md 的流程做。下面是触发示例：

- "按 shotlist-studio 流程，把这份甲方稿做成客户经理示例的一天那种格式"
- "把这份银行拍摄脚本整理成制作版"
- "我的脚本是 .doc 格式，老格式，先转 txt 再走 skill"

## License

MIT — 自由使用、修改、商用。引用时附仓库地址即可。

## 致谢

本 skill 的格式规范参考了金融行业常见的拍摄脚本模板，迭代自 11 份真实银行 / 政府 / 文旅客户的脚本样本。
