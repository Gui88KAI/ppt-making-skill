# ppt-making-skill

面向创赛、商业计划书、融资路演、答辩和课程的 PPT 制作 Skill，支持按项目自动生成素材与版面。

先读 [SKILL.md](SKILL.md)。入口会复用已知需求、询问缺项，并根据用途读取以下一个独立方法：

| 方向 | 特点 | 当前验证状态 |
| --- | --- | --- |
| [01 创赛](directions/01-创赛/GUIDE.md) | 比赛分流、项目贡献、证据与饱满视觉 | 有旧版实践；新方法与子类待独立成品验证 |
| [02 商业计划书](directions/02-商业计划书/GUIDE.md) | 支持独立阅读和经营判断 | 新方法成品待验证 |
| [03 融资路演](directions/03-融资路演/GUIDE.md) | 面向投资人的限时讲述 | 新方法成品待验证 |
| [04 答辩汇报](directions/04-答辩汇报/GUIDE.md) | 分学术、结题、个人评优子场景 | 子场景成品待验证 |
| [05 课程与主题班会](directions/05-课程与主题班会/GUIDE.md) | 帮助理解、讨论与行动 | 有班会实践；新方法与其他课程子场景待验证 |

五类不仅更换栏目或页数：各有问答、制作顺序、页型、视觉与评价标准。旧 `types/` 路径只保留兼容跳转，不维护第二套规则。

结构化需求可用 `python scripts/intake.py brief.json` 检查。脚本只核对已提取字段，不能理解自然语言、核实授权或证明视觉和事实。实际制作仍需读取所选GUIDE并使用宿主可用的生图、PPTX和渲染工具。

## 安装

```sh
git clone --branch codex/wps-compat-experiment https://github.com/Gui88KAI/ppt-making-skill.git
mkdir -p ~/.agents/skills
cp -r ppt-making-skill ~/.agents/skills/ppt-making-skill
```

`~/.agents/skills` 为示例位置，按宿主的 Skill 安装目录调整。

## 素材与模板

默认从零生成。使用 `gen_assets.py` 生成图表、流程图、时间轴和背景，需要题材图时调用可用的生图工具。用户提供的原图和证据仍按实际来源使用。

```sh
python -m pip install python-pptx pyyaml pillow matplotlib
python scripts/deck_builder.py examples/demo.yaml -o demo.pptx --check
```

此命令只演示规格转 PPTX。修改 YAML 中的文案、配色和页型可重新生成。配图需按项目实际生成，并将图片路径写入对应页的 `image` 字段。图表 PNG 需要 Matplotlib，流程图、时间轴及背景使用 Pillow。命令和输入格式可通过 `python scripts/gen_assets.py --help` 和 `python scripts/gen_assets.py chart --help` 查看。系统没有合适的中文字体时，可将 `PPT_ASSET_FONT` 设为本机字体文件路径。

`templates/` 中的 tech、red、rural 供可选参考。存在且风格贴合时，优先看已有预览并借鉴配色和构图，再从零生成。用户明确指定模板时，才按[模板指南](templates/README.md)填充副本。缺少模板或预览不影响默认流程。

WPS 目标读[兼容规范](references/WPS兼容规范.md)。分别记录静态检查、视觉渲染和实机播放。

明确指定 tech 模板且文件可用时，可将映射改成项目内容后生成副本：

```sh
python scripts/fill_template.py templates/tech/template.pptx templates/tech/mapping.example.yaml -o out.pptx
```

支持Skill发现的宿主根据 `SKILL.md` 的描述加载；不支持自动发现的宿主需要用户或宿主提供入口与相关文件。当前实验版不自动修改或合并主分支；请在独立环境试用，按项目已有授权维护。
