# 共享工具：不决定五类版式

各方向独立决定需求、叙事、页型、视觉密度和动画节奏。这里仅索引共同操作；先读根[入口](../SKILL.md)并按所选方向确定计划，再按需使用。

- 生图和边缘处理：[配图生成规范](../references/配图生成规范.md)、[用户偏好](../references/用户偏好与反馈沉淀.md)。有原生工具时实际调用，独立主题图不重复；用户认可原图只修边时保留原图。
- 素材生成：`scripts/gen_assets.py` 生成图表、流程、时间轴和背景。题材配图按[配图生成规范](../references/配图生成规范.md)调用可用工具。用户明确指定模板时才读取[模板指南](../templates/README.md)。
- PPTX导出与检查：`scripts/deck_builder.py` 是兜底生成器，不负责自主问答、创作配图或审美判断；`scripts/delivery_check.py`、`layout_check.py`、`image_usage_check.py` 只检查相应对象和结构。`scripts/title_editability_check.py` 可核对命名的艺术标题是否逐字为独立原生文本框，不代替WPS实机检查。
- WPS目标：[WPS兼容规范](../references/WPS兼容规范.md)、`scripts/wps_compat.py`。静态检查、实际渲染和目标软件放映分别记录。
- 代表页及交付：[代表页与交付验收](../references/代表页与交付验收.md)。其中通用检查不覆盖各方向专用的用途与审美评价。

文件、动画对象及材料绑定要对应实际产物。工具能力不可用时如实记录，不能把占位或计划当成完成。
