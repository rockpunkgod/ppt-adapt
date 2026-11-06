# 契约、冲突和整套验收

## v1输入

`design-spec.json`：`version:1`、`deck_id`、`canvas_cm:[width,height]`或`{width,height}`（厘米）、`tokens.fonts.heading/body.{ea,latin}`、`tokens.font_sizes_pt.{title,body,caption,footer}`（点）、`tokens.colors`、`tokens.spacing_cm`。可声明 `quality.min_body_pt/max_motion_groups_per_page`，以及全局 `master_id/layout_id`。可选 `variants.night` 包含明确 `pages:["s05","s06"]`、`background_role/text_role` 与局部 `tokens`；只作用于列明页面。

`deck-plan.json`：`version:1`、`deck_id`、`pages`。每页：

```json
{
  "id":"s01", "order":1, "role":"evidence", "chapter":"需求",
  "message":"证据支持的本页论点", "evidence_refs":["simulation"],
  "layout_family":"chart-plus-verdict", "required_mechanisms":["chart_build_*"],
  "variant":"default",
  "objects":[{"id":"chart","kind":"chart","bounds_cm":[1.5,4,24,11],"data_ref":"demand"}],
  "relations":[],
  "motion":{"transition":"none","duration_ms":0,"groups":[
    {"id":"reveal1","trigger":"click","sequence":"main","click_index":1,
     "start_ms":0,"duration_ms":500,
     "writes":[{"object":"chart","property":"category[1].visibility","value":true}]}
  ]}
}
```

证据注册表可在spec或deck-plan的 `evidence` 中，使用 `{id,...}` 数组或id字典。此处检查引用存在；原始数字的正确性、数据文件字段解析和标题是否被证据支持，仍需内容验收。`data_ref`是制作层输入引用，不会被执行为代码。

CLI另绑定真实本地证据：`source_files.evidence`是`[{id,declared_path,path,sha256,bytes,local:true}]`，其中path为绝对路径，相对输入按**deck-plan文件的目录**解析。缺失本地文件直接报错；远程URL只记录uri/local:false，不联网或伪造SHA。制作入口在执行任何写入前调用 `verify_source_files(plan)`，同时检查spec、deck-plan、全部本地证据和配方索引现文件；数据改值后旧计划不得继续制作。`compile_plan(spec,deck,index)`保持纯数据API，CLI的真实文件绑定不可省略来代称数据已冻结。

文本对象带 `font_role`：heading/title用标题字号与heading字体，body用正文，caption/footer用相应字号和body字体；保留 `text`。显式 `fonts` 冲突可按角色token修正。字号的规范字段为 **font_size_pt**，优先于旧size_pt；缺省继承token，两个字段输出同步。显式字号必须正且有限，并遵守逐对象可读性底线；heading/title/body受min_body_pt约束，caption/footer仅受明确min_caption_pt/min_footer_pt约束。低于底线则阻断，不能改大后声称输入没有问题，也不能缩字解决溢出。

命名字形Morph使用同一 `morph_key`，输出原生命名 `!!key`，同页不可重复；其他对象规范名字是 **deck_id+'.'+obj.id**（obj.id已包括s01.title等页面前缀）。`native_name`与`name`镜像一致，制作层必须消费effective字段，不另造命名。

支持空间关系：`left_of`（from/to或objects，gap_cm）、`equal-width`、`align-left`、`align-top`（objects，tolerance_cm），以及跨页 `morph-pair`（target_page、keys）。不支持的关系阻断，不默默当通过。bounds是逻辑几何约束；不代证字形/旋转视觉范围与文本实际溢出。

机制可用字符串或 `{query,recipe?,reuse_scope,excluded_capabilities}`。query必须匹配真实索引机制/入口；recipe用于明确现有入口/id，不是上传一个新引擎。`reuse_scope`明确本页实际复用哪部分；`excluded_capabilities`列未复用的完整配方能力，例如只复用!!身份/700ms几何Morph，排除真实四矩形孔罩。相同页重复选择同一索引行会去重。每项选中的 `scope` 保留索引原文，不被应用范围覆盖；backend、行号与索引字节SHA都保留。未写reuse_scope可以暂时planned/unverified，但application_review标unspecified，不能升级组合verified。

## 冲突决策

|类别|检测与处理|
|---|---|
|字体|缺EA/Latin、未知角色、显式字号/最低正文要求不符阻断；角色字体不符可改回token，保文字|
|主题|语义色按token修；variant仅允许明确列页；未知variant/颜色角色阻断|
|母版|存在显式全局master_id时恢复正确作用域；允许variant另定；未声明则不虚构母版对象|
|布局|同上处理layout_id；坏bounds、未知关系、等宽/对齐/间距失败阻断，要求真正重排|
|命名|可生成稳定原生命名；对象id重复使引用不明确时阻断，不自作主张改引用|
|Morph|同页多对象同key、缺跨页配对、对象种类变化、自动局部几何动画抢过渡属性阻断|
|时间轴|ms为唯一规范，seconds=ms/1000；错误显式秒值有修复日志；重叠/不可比时钟同属性双写阻断|
|接口|未知/多义配方、错误deck_id、未知backend、COM缺CLI关闭依赖阻断|

同一main sequence中的**不同click_index**是不同事件时钟，两个0ms不等于同时；相同/未知click或独立交互sequence写同属性时，必须声明明确互斥/顺序再实施，当前保守阻断。不能随意挪动画到后面把依赖语义改掉。`writes`须完整描述本次组合实际会改的属性；索引文字不足以推导全部脚本内部副作用，制作适配层还要核对，不能因未声明就声称没有冲突。

操作DAG默认串行页面作者、关闭节点、必要COM精修；`operations`扩展节点可有 `id/backend/depends_on`，环或未知依赖阻断。它是交接计划，不运行命令，不自行取得共享PowerPoint使用权。具体配方参数/资产适配读取原参考；没有把多个完整示例脚本都运行到同一PPTX上当拼装器。

`conflicts.json`分清真正冲突、自动修复、缺省派生值。所有修复记录before/after/reason/priority。固定优先级是内容完整正确、可读性、全局一致、动画、装饰；涉及内容删改、字号缩小或无法判断的重排保持阻断。

## verified的严格含义

规划与渲染分离；初次只能planned/unverified。每项配方必须先明确reuse_scope，才允许绑定整套验收；不能因为字符串匹配索引一行就继承其整原子通过。绑定报告要求：

```json
{
  "scope":"whole-deck", "kind":"native-powerpoint", "passed":true,
  "acceptance_id":"本次独立原生验收运行ID",
  "plan_sha256":"composition-plan.json的plan_sha256",
  "deck_sha256":"本次final.pptx的文件SHA256",
  "index_sha256":"plan.index.sha256",
  "spec_sha256":"plan.input_hashes.spec",
  "deck_plan_sha256":"plan.input_hashes.deck_plan",
  "pages":["s01","s02"],
  "evidence":[
    {"kind":"png","page_id":"s01","path":"绝对路径/s01.png","sha256":"真实文件SHA256"},
    {"kind":"png","page_id":"s02","path":"绝对路径/s02.png","sha256":"真实文件SHA256"},
    {"kind":"native-recording","path":"绝对路径/actual-show.mp4","sha256":"真实文件SHA256"}
  ]
}
```

`input_hashes`是输入解析后的规范JSON SHA；`source_files`另存输入原始字节SHA及路径。计划SHA计算规范JSON并排除plan_sha256/status/verification这三个可变字段，其余完整内容（包括输入路径与索引路径）都绑定。报告必须覆盖精确完整页面集合与每页真实PNG；有非cut/none转场或动画组时必须附真实录屏。旧配方的PNG/影片不能改scope冒充本次8页整体验收。此绑定检查文件完整性和范围，不是影像真实性鉴定；制作方需保存真实播放/点击日志与人工/自动核验项。

## 来源与非来源规则

- [storytelling with data：Plan your PPT presentation in under 10 minutes](https://www.youtube.com/watch?v=XSDRpIdzze8)，1:14–3:08先故事板、3:09–6:10开场/证据/结尾角色、6:18–7:12另评讲义扫描性。
- [Firm Learning：HOW CONSULTANTS MAKE SLIDES](https://www.youtube.com/watch?v=Qv3mrOtFMzM)，2:02–5:03一格一页的论点标题/证据草图、5:06–6:35先整套storyline；其MickeyMouseMaster是草稿，不是PowerPoint母版。
- [Presentation Process：All about PowerPoint Slide Master](https://www.youtube.com/watch?v=qxvIUhFOdrs)，6:12–7:44封面隐藏背景例外、15:02–16:26布局Fade/dim、16:33–17:59章节第二母版。HideBackground会连带隐藏其他母版图形；全局转场不得覆盖局部Morph/Zoom。
- [Kevin Stratvert：How to use PowerPoint Slide Master](https://www.youtube.com/watch?v=kHRc-3rMCg4)，1:56–3:06全局字体、3:08–3:48重复标识与本地图像遮挡风险、4:14–5:10仅标题布局的人像。

这四片已完整读取可用字幕1201段、查看26个定点帧；不是连续全片视觉观看。已有SCQ→A→理由/证据复用 `officecli-ppt-design/references/layout-story-geometry.md` 的原有边界。视频未规定动画密度数字/固定8页；DAG、hash绑定、write sets与当前schema是本工程组合规范。作者关于默认字体的喜好不当禁令，关于重复图片线性增加文件体积的说法不当通用规律。
