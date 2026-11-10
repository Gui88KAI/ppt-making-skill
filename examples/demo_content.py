"""deck2_content.py — 示范 PPT 内容（智慧农业无人机植保巡检）。

图片由 gen_image.py 生成于 assets/。
"""

DECK = {
    "pages": [
        dict(type="cover", bg="cover",
             kicker="第十届中国国际“互联网+”大学生创新创业大赛",
             title="慧飞植保\n让每一亩地都用上无人机",
             subtitle="飞防服务 + 数据驱动的精准施药",
             tags=["智慧农业", "无人机", "精准施药"],
             org="慧飞科技 · 大连民族大学 · 2026"),

        dict(type="section", number="01", title="痛点与机会",
             note="为什么现在，为什么是我们"),

        dict(type="image_cards", title="一亩地，人工打药要 40 分钟", tag="项目痛点",
             sub="三个痛点，每个都能算成钱",
             image="assets/pain.jpg",
             source="数据来源：2025 年东北地区植保作业调研，样本 186 户",
             cards=[
                 ("打药靠人扛", "背负式喷雾器日均作业 3–5 亩，50 岁以上农户占 62%", "人力瓶颈"),
                 ("雾滴粗、浪费大", "人工喷雾粒径 200μm 以上，农药流失率超 40%", "成本高企"),
                 ("不会看虫情", "凭经验施药，错过最佳窗口，平均减产 8–12%", "决策盲目"),
             ]),

        dict(type="image_hero", title="无人机打一遍，8 分钟一亩", tag="解决方案",
             sub="飞防服务 + 数据驱动的精准施药",
             image="assets/solution.jpg",
             points=[
                 "多旋翼植保机，载药 20L，单架次日作业 300–500 亩",
                 "雾滴 80–120μm，穿透性好，农药利用率提升到 70%",
                 "多光谱巡检定位虫害区域，只打有虫的地块",
             ],
             source="来源：2025 年田间对比试验，3 个试验点"),

        dict(type="kpi", title="三个数字，都能查到出处", tag="成果验证",
             sub="不是估算，是实测",
             bg="data",
             source="数据来源：2025 年 7–9 月现场作业记录",
             items=[
                 dict(value="8", unit="分钟/亩", label="对比人工 40 分钟"),
                 dict(value="70%", unit="农药利用率", label="行业平均 35%"),
                 dict(value="130", unit="元/亩 省", label="含农药与人工"),
             ]),

        dict(type="compare", title="同一块地，两条路", tag="价值对比",
             sub="按 500 亩规模测算",
             left=dict(title="人工打药", items=[
                 "40 分钟/亩，单季需 20 人天",
                 "农药流失 40% 以上",
                 "虫情靠经验，错过窗口",
                 "综合成本 190 元/亩",
             ]),
             right=dict(title="无人机服务", items=[
                 "8 分钟/亩，单季 4 人天",
                 "雾滴细，利用率到 70%",
                 "多光谱巡检先定位再打",
                 "综合成本 60 元/亩",
             ])),

        dict(type="end", title="谢谢",
             bg="team",
             note="慧飞科技 · 让每一亩地都用上无人机"),
    ]
}
