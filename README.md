## v2 优化方向

这一版把 PPT 生成从“逐页套版式”改成“先建立视觉圣经和故事线，再生成候选素材、安排版式节奏和连续动效”：

- `references/视觉叙事系统.md`：视觉圣经、故事线、上下文账本。
- `references/版式变体与节奏.md`：不规则版式家族和整套节奏约束。
- `references/动画连续性规范.md`：Morph 对象命名、页面内动效和兼容性边界。
- `scripts/compile_prompts.py`：为每个图位编译 3 张有上下文的候选提示词。
- `scripts/deck_lint.py`：检查叙事字段、版式重复、候选图和运动对象。
- `scripts/motion.py`：写入 Morph/淡出转场并为跨页对象命名。
- `scripts/deck_builder.py`：新增 asym_split、diagonal、orbit、staircase、collage、quote 版式。

# challenge-cup-ppt

做挑战杯 / 互联网+ / 创赛路演 PPT 的 agent skill：从 100GB+ 模板素材库里按需取用，按九步工作流产出可上台的 PPT。

## 这是什么

**不是模板填充工具，是一套让 AI 自己做设计判断的思考框架。**
同一份内容，不同 AI 用它应该做出不同风格的 PPT——分叉发生在"定主张、排结构、写版式"这些判断上，而不是换配色。

**项目进展与决策记录见 [PROGRESS.md](PROGRESS.md)** —— 上下文被压缩或换人接手时，先看这份。

两种用法：

- **给 agent 用**：`SKILL.md`（带脚本工具链）
- **给任意聊天 AI 用**：把 `prompt.md` 里的提示词整段复制粘贴

包含：

| 文件 | 作用 |
| --- | --- |
| `SKILL.md` | 技能主体：像设计负责人一样思考（brief→主张→结构→版式→素材→自我批评→迭代） |
| `prompt.md` | 可直接粘贴给任意 AI 的通用版提示词（不依赖工具） |
| `references/` | 素材地图、配色与版式、字体与动画坑、完整经验文档 |
| `scripts/gen_assets.py` | **自己生成素材**：图表 / 流程 / 时间轴 / 背景（透明底、主题色） |
| `scripts/layout_check.py` | **排版自检**：文字溢出、元素重叠、字号过小、空页 |
| `scripts/template_map.py` | **打印现成模板的结构**（每页每元素的编号/位置/文字），用于改造模板 |
| `scripts/fill_template.py` | **把内容填进现成模板**，设计全部保留（首选路线） |
| `scripts/deck_builder.py` | 从零生成整套 PPT（9 种页面类型，兜底方案，设计不如现成模板） |
| `scripts/find_template.py` | 在模板索引里检索，避免加载整个素材库 |
| `scripts/pptx_inspect.py` | PPTX 体检：页数、版式、字体、字号、占位符残留 |
| `assets/模板索引.json` | 468 个模板包 / 2697 个 pptx 的索引（赛道分类、预览图数、字体数、体积） |
| `examples/demo.yaml` | 一套 8 页的示例规格，可直接改成自己的项目 |

## 核心思路：渐进式读取，不加载全量

素材库 100+ GB，绝对不要整体加载。三层读取：

1. **查索引**（1.2 MB，秒级）
   ```bash
   python3 scripts/find_template.py 蓝色 科技 --limit 5
   python3 scripts/find_template.py --cat 红旅非遗 --min-pptx 20
   ```
2. **只解压命中包的预览图**（几十 MB）挑模板
3. **只解压选中的那一个 pptx**

```bash
unzip "$PPT_LIBRARY_ZIP" "<包路径>/*预览*" -d /tmp/ppt
unzip "$PPT_LIBRARY_ZIP" "<包路径>/xxx.pptx" -d /tmp/ppt
```

## 安装

```bash
git clone https://github.com/Gui88KAI/challenge-cup-ppt.git
cp -r challenge-cup-ppt ~/.agents/skills/
```

其他 agent（Claude Code / OpenClaw / Cursor 等）也可以直接把 `SKILL.md` 作为指令读入，脚本按需调用。

## 素材库

素材库本体（约 103 GB，6590 个文件）**不在本仓库**：

- GitHub 单文件上限 100 MB（素材库里 174 个 pptx 超限），单仓库建议 < 5 GB，102 GB 无法托管
- 本仓库只放**索引 + 脚本 + 经验**，素材库用网盘托管，下载链接见 [`library/sources.json`](library/sources.json)
- agent 先查索引，再只下载/解压需要的那一个包

用环境变量指定素材库位置：

```bash
export PPT_LIBRARY_ZIP=/path/to/挑战杯PPT制作素材库.zip   # 默认 /sdcard/Download/DSH-Go/...
```

素材库结构：

```
挑战杯PPT素材合集/
├── 01-百度网盘-999套PPT模板/   74 GB  通用模板库 + 国赛创赛模板 + 字体库
├── 02-夸克-创赛挑战杯/         28 GB  按赛道配色分好的创赛模板
└── PPT制作经验.md
```

## 两条路线

> **素材库里的模板是设计师做的**（深色科技底、光效、图标、数据卡片）。
> 想做出那种效果，必须改造现成模板；从零生成只能保证结构正确。

### 路线 A：改造现成模板（默认，质量最高）

仓库自带 3 套精选模板（`templates/tech|red|rural/`，各带 preview.jpg 和 mapping 骨架），clone 下来就能直接用，不需要先下载 100GB 素材库：

```bash
# 0) 用自带模板（最简单）
python3 scripts/fill_template.py templates/tech/template.pptx templates/tech/mapping.example.yaml -o out.pptx

# 1) 或去素材库挑模板（先看预览图）
python3 scripts/find_template.py 科技 --limit 8

# 2) 看模板结构
python3 scripts/template_map.py 模板.pptx --slide 1

# 3) 写映射（格式见 examples/template_fill.yaml）

# 4) 填内容，设计不动
python3 scripts/fill_template.py 模板.pptx 映射.yaml -o 输出.pptx
```

### 路线 B：从零生成（兜底）

**方式一：从零生成（最快）**

```bash
pip install python-pptx pyyaml
python3 scripts/deck_builder.py examples/demo.yaml -o demo.pptx --check
```

改 `demo.yaml` 里的文案和配色，再生成一遍就是自己的 PPT。

**方式二：用素材库模板（要设计感）**

```
1. 需求访谈 → 01-需求.md（赛道、时长页数、必讲模块、现场设备）
2. 建项目目录
3. 检索素材 → 02-检索结果.md（find_template.py）
4. 选模板（3 选 1，先看预览图）→ 03-选型.md
5. 模板体检 → 04-模板结构.md（pptx_inspect.py）
6. 写文案大纲 → 05-大纲.md（一页一个结论，≤60 字，见 references/文案句式库.md）
7. 版式映射 → 06-映射表.md（2–3 种内容版式循环）
8. 生成 → deck_builder.py 从零拼 / python-pptx 批量替换模板
9. 检查交付 → 07-检查报告.md
```

## 页面类型

`deck_builder.py` 内置 9 种页面，覆盖一套路演 PPT 的全部需要：

| 类型 | 用途 |
| --- | --- |
| `cover` | 封面（标题 + 副标题 + 可选背景图） |
| `section` | 章节过渡页（大数字 + 章节名） |
| `cards` | 2×2 / 3 列卡片（讲几个要点） |
| `bullets` | 要点列表 |
| `kpi` | 数据大屏（3 个大数字） |
| `timeline` | 时间轴 / 规划 |
| `compare` | 左右对比（VS） |
| `image_text` | 左图右文（产品展示） |
| `closing` | 结尾页 |

## 依赖

```bash
pip install python-pptx
```

## License

MIT
