# WPS 兼容制作与验收

用户指定 WPS 演示、WPS Office 或教室电脑以 WPS 播放时读取本文件。它处理播放环境，不限定题材、风格、页数或受众。已明确的信息直接沿用，版本未知可继续保守制作。

## 配置

```yaml
compatibility:
  target: wps
  platform: windows
  version: null
  preferred_transition: fade
theme:
  font_title: 微软雅黑
  font_body: 微软雅黑
  font_num: Arial
```

上述字体是 Windows 示例，不保证所有 WPS 平台自带。已有指定字体先保留；缺少时选择目标设备确实安装且授权允许使用的字体，替换后重新检查换行。正文保持可编辑，必要的艺术字可转清晰 PNG，不得将整套 PPT 截成图片掩盖排版问题。

## 制作原则

- 交付标准 `.pptx`，保持用户指定的画布比例。图片优先使用本地嵌入 PNG/JPEG；复杂 SVG/EMF/WMF 转换后核对清晰度。图表和需编辑的正文保持原生对象。
- 显式声明 Latin、East Asian、Complex Script 字体，避免中文依赖隐含主题回退。填写字体名不等于嵌入字体；安装字体也不等于目标电脑已安装。
- 未完成目标 WPS 放映验证前默认 `fade` 或 `none`，不依赖 Morph、3D、OLE、宏和外链媒体表达关键信息。用户明确要求高级动画时记录需求和待验证项，并提供独立可读的静态版本。
- 封面与模板路线同样执行该配置，优先级高于旧参考中的开屏/Morph 要求；开屏可以保留，但无高级动画时各页仍须读得懂。
- `entry_groups` 是制作计划，生成器未实现元素级进入动画，不得写成已完成。

## 工具

```bash
python scripts/deck_builder.py deck.json -o deck.pptx --target wps --check
python scripts/wps_compat.py deck.pptx --report wps-report.json --strict
```

`--target wps` 优先于规格中的目标；否则读取 `compatibility.target`。WPS 模式仅写淡出/无转场，原有 Morph 请求记录回退。静态检查读取实际 PPTX 的转场、字体声明和资源关系，检查外链图片/媒体、图片格式、宏/ActiveX/OLE 和字体嵌入资源，不把人工“已验证”字段当成证据。

现成模板和其他脚本的输出可生成单独的兼容副本：

```bash
python scripts/wps_compat.py input.pptx --prepare output-wps.pptx --report prepare-report.json
```

它规范字体声明、将转场改为淡出，不改变文字、布局或图片位置，不覆盖原件。仅在明确需要静态副本时加 `--strip-animations`；时间线移除后可能出现叠层内容同时可见，须重新逐页检查。`--font 微软雅黑` 会有意替换全部原生文字字体，确认目标设备和用户需求后再用。

## 实机验收和交付

用目标 WPS 打开最终导出文件，确认无修复提示；逐页检查字体、换行、遮挡、图片裁切，放映检查翻页、过渡、内部动画和音视频。记录 WPS 版本、系统、文件 SHA-256、检查页码和问题处理。重新导出后原验证记录失效。

条件允许时从 WPS 导出 PDF 和逐页图片，并检查页数和内容。其他工具导出的 PDF 标明来源，Pillow 预览仅用于草图。

没有 WPS 时继续交付可编辑 PPTX 和静态报告，明确写“已按 WPS 保守配置生成，尚未在目标 WPS 实机验证”。不得伪造版本、截图、PDF 或通过状态。静态检查通过不证明字体替换、透明效果、换行和实际播放正确。

## 官方依据与边界

- [WPS 学堂：嵌入字体](https://www.wps.cn/learning/question/detail/id/1733)：嵌入字体属于保存选项，需要核对实际文件和目标设备结果。
- [Microsoft：Morph 扩展](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-pptx/68d26d78-f7f5-47ab-835d-4e6c82ff39f0)：Morph 不属于基础 `p` 命名空间，不能写成 `p:morph`。
- [Microsoft：转场扩展与回退](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-pptx/22ebe6b5-2ade-43d9-977a-98fa194725c2)：扩展使用 AlternateContent 和 Fallback。

淡出/无转场、嵌入图片和显式字体是本 Skill 的保守制作策略，不是 WPS 全版本兼容保证。
