# 可选模板

默认按 [SKILL.md](../SKILL.md) 从零生成。只有本地存在 `.pptx` 且风格适合时，才参考对应的配色和版式。先看 `preview.jpg`（如有），没有预览时可查看模板已有页面，或者直接跳过。整个目录缺失或为空时继续默认路线。

| 目录 | 风格 | 随仓版本预览 |
| --- | --- | --- |
| `tech/` | 深蓝科技 | 有 |
| `red/` | 红金、文化主题 | 有 |
| `rural/` | 绿色、乡村与农业主题 | 无 |

用户明确指定模板时，才使用以下流程。先查看元素编号，复制并修改映射文件，再输出 PPTX 副本。

```sh
python scripts/template_map.py templates/tech/template.pptx --slide 1
python scripts/fill_template.py templates/tech/template.pptx templates/tech/mapping.example.yaml -o out.pptx
```

映射中的 `index` 为从 1 开始的页码，元素编号从 0 开始，组内路径写成 `"4.2.0"`。示例映射的空字符串用于演示清空，制作时改成实际内容并检查未映射的旧文案和图片。指定模板文件确实不可用时说明情况，并交付从零生成的可做部分。

模板内字体和图片的授权状态应按实际用途核实。打开成品后检查字体替换、文本长度与放映效果。目标为 WPS 时可加 `--target wps --check` 执行静态检查。
