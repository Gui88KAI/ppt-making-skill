# 生涯规划成长赛道 Demo

从已制作的本地《生涯发展展示·成长赛道》提炼了 5 种页面任务，文案用通用示例呈现。随附的 5 张图是原项目生成的场景示意，个人照片、身份、奖项和原始项目证据留在原稿。

| Demo 页 | 素材 | 原作对应场景 |
| --- | --- | --- |
| 1 封面 | `cover-photo.jpg` | 电子实验台 |
| 2 职业探索 | `origin-photo.jpg` | 笔记桌与开发板 |
| 3 学习实践 | `sensors.jpg` | 开发板与传感器 |
| 4 成长计划 | `horizon-photo.jpg` | 面向部署的机械臂工作台 |
| 5 职业目标 | `robot.jpg` | 机器人与视觉模组 |

在仓库根目录执行：

```sh
python -m pip install python-pptx pyyaml pillow
python scripts/deck_builder.py examples/demo.yaml -o demo.pptx --check
```

`photo_story` 的图片路径相对于规格文件。通过 `text_x`、`text_y`、`text_width`（厘米）将文字放在画面留白处，`ink` 和 `highlight` 分别指定正文与强调色，`lettered` 将标题拆成可移动的单字文本框。几页之间切换深浅底与文字区域。素材定位来自已有照片构图，新项目应按内容选择相应场景。

[成品 PPTX](demo.pptx) · [实际渲染预览](preview.jpg)。使用华文行楷与微软雅黑，字体由打开设备提供。样例已渲染并通过 WPS 静态检查。
