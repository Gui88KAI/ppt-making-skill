# templates/ — 随仓库分发的精选模板

素材库有 100+GB、468 个模板包，太大没法全放进仓库。这里放 **3 套精选模板**，保证任何人 clone 下来就能跑出和素材库同样水准的效果（而不是从零画的简约风）。

| 目录 | 风格 | 适用赛道 | 页数 | 体积 |
| --- | --- | --- | --- | --- |
| `tech/` | 深蓝科技 + 光效 + 机器人配图 | 科技 / AI / 芯片 / 医疗设备 | 8 | 31 MB |
| `red/` | 红金 + 非遗质感 | 红旅 / 非遗 / 文化传承 | 16 | 34 MB |
| `rural/` | 绿意 + 乡村实景 | 乡村振兴 / 农业 | 20 | 20 MB |

每套目录里：

```
template.pptx            模板本体（保留全部设计）
preview.jpg              原作者做的预览图（看风格用；rural 无预览）
mapping.example.yaml     映射骨架，元素编号已列好，填文案就行
```

## 怎么用

```bash
# 1. 看结构（可选，mapping 骨架里已经有编号）
python3 scripts/template_map.py templates/tech/template.pptx --slide 1

# 2. 改 mapping.example.yaml：把 "" 换成你的文案
#    注释里是原模板的文字，照着改最省事

# 3. 生成
python3 scripts/fill_template.py templates/tech/template.pptx templates/tech/mapping.example.yaml -o out.pptx
```

## 字体（已随仓库附带）

`templates/fonts/` 里放了这三套模板实际用到的 9 个字体（69MB），装完即可 100% 还原设计：

```bash
# Windows：全选 templates/fonts/ 里的字体 → 右键 安装 → 重启 Office
# Linux:
mkdir -p ~/.fonts && cp templates/fonts/* ~/.fonts/ && fc-cache -f
```

缺字体时用 `思源黑体 CN` 兜底，版式不变、字形会变。详见 `templates/fonts/README.md`。

## 三条硬提醒

1. **字体**：这些模板用了汉仪旗黑 / 汉仪雅酷黑 / 汉仪行楷等商业字体。不装字体，打开后会被替换成系统字体，设计感会掉。素材库的模板包里有字体文件（版权归原字体厂商，本仓库不打包）。
   - 替代方案：把字统一换成「思源黑体 / 阿里巴巴普惠体」，版式不变、气质略变。
2. **文字长度**：模板的文本框尺寸是设计好的。`mapping.example.yaml` 注释里写了原文，**新文案不要明显超过原文长度**，否则会溢出文本框。
3. **图片**：`images: {元素编号: "图片路径"}` 会按原框等比裁切铺满。

## 想要更多模板

素材库（100+GB）不在仓库里，用 `library/sources.json` 里的网盘链接下载，然后用 `scripts/find_template.py` 按赛道检索。
