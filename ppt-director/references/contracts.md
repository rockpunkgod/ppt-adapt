# Director契约v1

本契约允许额外字段。所有文本必填项为非空字符串，所有数值须有限；ID在所属集合唯一。目录相对引用按plan文件目录解析。spec须先存在，`design_spec`必须解析到本次传入的spec，同deck_id和version=1。尺寸用cm、字号用pt、时间用ms；执行COM时才转换为秒。

## design-spec.json

必填：`version:1, deck_id, canvas_cm:{width,height}, audience,purpose,decision, tokens, variants, quality`。

- tokens.fonts.heading/body各有`ea,latin`字体名。
- tokens.font_sizes_pt包含`title,body,caption,footer`正字号；heading角色使用title字号。对象可选`font_size_pt`覆盖，body覆盖仍不得小于quality.min_body_pt。
- tokens.colors包含paper/ink/accent1/accent2/muted/night，值为六位`RRGGBB`或`#RRGGBB`；允许扩展on_dark等色。默认页用paper/ink；变体text_role是语义色角色，由执行层解析（建议把on_dark明确列入colors）。
- tokens.spacing_cm含`margin,gap`非负。页面尺寸不锁死16:9；根请求建议33.87×19.05。
- variants为对象，可以为空；default隐含不可重定义。每个具名变体含`pages:[页面id]`、`background_role`（colors现存key）、`text_role`、`reason`。范围与页面variant双向精确一致。
- quality含`min_body_pt`正数、`max_motion_groups_per_page`非负整数。建议18与4，不强行将用户明确要求的可读字号改为默认；对caption/footer允许小于正文下限。

## deck-plan.json

必填：`version:1,deck_id,design_spec,narrative,evidence,pages`。

叙事narrative含opening/throughline/closing_ask非空。evidence每项含id/kind/path：kind为user/source/data/synthetic等说明性字符串；本地path必须存在（文件或目录），https/http URL只检查格式，不声称已下载或核真；纯原创内容也用本地简短来源说明，记录synthetic，不能空引用。敏感凭据不要写进URL。

每页必填id、order（整套从1连续）、role、chapter、message、evidence_refs（至少一项已声明证据）、layout_family、required_mechanisms（数组，可空）、objects（数组，可空）、relations（数组，可空）、motion。可选page_job补充role，variant省略即default，rhythm_reason解释连续同角色或同layout；对象和页可添加制作需要的字段。

required_mechanisms每项可以是非空字符串，或`{query,recipe?,reuse_scope?,excluded_capabilities?}`：query非空字符串，recipe可选字符串用来明确真实索引入口；reuse_scope若提供须非空，写本次实际借用的原语/适用参数；excluded_capabilities若提供须为字符串数组，每项非空，列未复用或未验证的整原子能力。额外字段允许，不能用它们替代query。旧字符串继续兼容；显式只借稳定对象身份与Morph插值不等于通过整套遮罩/边缘无漏缝能力，计划和最终验收须保留边界。

objects每项含id/kind/bounds_cm:[x,y,w,h]，w/h必须正且对象在画布内。文本kind为text/title/heading/body/caption/footer或含text字段时，必须text非空、font_role明确（heading/title/body/caption/footer）；非文字对象不强制字体。可选font_size_pt是实际覆盖值。`data_ref`若给出须为本页evidence_refs中的证据id，或在其本地JSON证据中唯一可解析的字段路径（例如`roadmap.phases[0]`）；多源歧义时用`evidence-id#field.path`限定。不存在、格式错误或跨页未引用的数据不能通过；URL数据不在本地假解析。chart/table必须有data_ref。morph_key可选非空，同页唯一，跨页重复表示同一连续对象。ID在页内唯一，执行层使用page-id/object-id命名空间。

relations每项为对象，三种形式：from/to引用本页对象；type+objects数组引用至少两个本页对象；或`type:'morph-pair',target_page,keys`，要求另一个真实页面且keys在两页各存在一次。类型标签描述语义，不代证几何间距真的相等。motion含transition=none|morph、duration_ms非负、groups数组；none时duration_ms必须0。每组为对象含id；可选duration_ms/delay_ms非负，targets或object_ids为本页对象id数组，depends_on为本页组id数组且无环。组数不超过spec门限。

## 校验界限与使用

`--page s05`仍检查全稿身份、叙事证据、页序和变体范围，仅把对象/bounds/关系/时序深检限制到s05，不能用它绕开全局契约。未知页是错误。所有错误阻止执行，退出码1；成功0；JSON/文件/输入格式失败同样为1。`--report`可写新JSON报告，拒覆盖现有文件。stdout始终输出JSON，含errors/warnings/checked_pages。

节奏检查是建议而非套版规则：多页只有单layout或四页连续同role且无rhythm_reason会警告；不强制opening/closing角色名字或固定“呼吸页”数量。全稿真正的叙事质量、事实支持度、视觉层次仍由Director判断。

计划校验不是作品通过。执行层还须检查真实对象类型、文字溢出与阅读字号、实际素材/数据、原生PNG；动画须看完整周期、进入退出和最终保持，交互须真点击与负对照。每份报告要绑定实际输入与最终文件，不能用旧稿验收代替当前稿。

维护校验器后可运行`python tests/test_director.py`。测试只使用标准库与临时目录，覆盖契约正负例，不启动Office或生成PPT；它也不代替内容与画面验收。
