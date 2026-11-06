# 字体、网格与真实轮廓配方

## 使用与边界

四片源教程完整自动转写与108分布操作帧已核查，并非逐秒连续观看。作品是机制复现，不携带作者模板、照片或字体。配方为静态；本机PowerPoint静态图足够验收，不额外添加动画。

# 来源与已查看范围

2026-11-06；Ego Lite独立space9。四片完整自动转写均已读；字幕可能误识别字体名，以下结合操作画面。共108帧均已查看：TeZ36、fCH24、ZJB36、Xz12。不是声称逐秒观看每个中间画面。

| ID / 实际标题 | 作者 / 时长 | 已还原的操作 |
|---|---|---|
| [TeZ-y45Ei1c — 3 ESSENTIAL Tips for Better Slide Design](https://www.youtube.com/watch?v=TeZ-y45Ei1c) | One Skill PPT / 19:00 | 网格对齐、字距行距字重、配色与主题色 |
| [fCHriW2sUJ4 — How to Select Two Fonts That Go Together: Font Pairing Explained](https://www.youtube.com/watch?v=fCHriW2sUJ4) | Strong Slides / 12:16元数据（播放器约12:15） | form model / contrast matrix筛选字体搭配，局限与10组示例 |
| [ZJB3trPn9_w — How to Create Awesome Typography Designs in PowerPoint](https://www.youtube.com/watch?v=ZJB3trPn9_w) | One Skill PPT / 8:59 | 文字变形、转轮廓、白边分层、相减高光、条纹相交阴影 |
| [XzIR8ep4vw0 — What Are the Best Fonts to Use in PowerPoint?](https://www.youtube.com/watch?v=XzIR8ep4vw0) | Envato Tuts+ / 4:14 | Georgia标题+Garamond正文的阅读层级；不是列出五款字体排名 |


# 原子技巧与后端

后端表示实际实现途径：officecli-native=只需OfficeCLI；powerpoint-com=原生PowerPoint自动化；hybrid=OfficeCLI基础+COM精修；external-render=外部渲染素材。本组没有必须external-render的已验证能力。

| 原子能力 | 源时间 | backend | 操作映射与验收 |
|---|---|---|---|
| F01 用网格决定坐标而非只靠视觉拖动 | TeZ02:28–05:33 | officecli-native | 按slideWidth/12、slideHeight/8算位置；标题/正文/按钮共享x；辅助网格不留在成品。基础页1已验证 |
| F02 时间轴先组内对齐再组间分布 | TeZ06:14–08:57 | officecli-native | 年号共底、说明共顶，每组共中轴；中心位置用等差序列。基础页2已验证 |
| F03 字体类别、字重与语气匹配 | TeZ09:19–13:19 | officecli-native | font/size/bold/italic；先查本机字体；可替代但记下字体差异。基础页3、字体配对页已验证 |
| F04 字距和行距必须独立控制 | TeZ11:11–12:51 | officecli-native | spacing是pt，lineSpacing支持.9x等倍数；不要把字距当倍数，.7行距不是通用标准。原生检查不碰字、不掉字 |
| F05 把palette写入主题槽 | TeZ17:19–18:12 | officecli-native | set /theme accent1..6与headingFont/bodyFont；具体形状可直接HEX或方案色。基础页4与theme XML检查 |
| F06 形态矩阵是候选生成法 | fCH01:10–09:58 | officecli-native | 先定标题/正文角色，再试a/e/O、粗细对比、serif；不确定可记混合形态；禁止把矩阵当确定审美评分器。字体3页实际试排 |
| F07 标题与正文需要角色差 | fCH10:19–10:54；Xz01:30–03:11 | officecli-native | Georgia+Calibri/Arial、Georgia+Garamond；标题明显更大、正文仍足够可读。font-pairing与font-hierarchy已原生验证 |
| F08 斜基线文字变形 | ZJB01:40–02:42 | officecli-native | set textWarp=textSlantUp；这是字形变形，不是整体rotation。1.0.160需修正文档元素顺序，详下 |
| F09 文字转轮廓后分层 | ZJB02:46–04:28 | hybrid | CLI建字和底；COM字与覆盖矩形Intersect。转后是freeform custom geometry，不再可改文字；保留before文件作为可编辑源。已验证5个原生路径对象 |
| F10 相减生成真实边缘高光 | ZJB04:31–05:27 | powerpoint-com | 棕色字形－右移少量字形；MergeShapes(4,primary)。不是模糊阴影或细线手描。934节点相减轮廓已验证 |
| F11 条纹先Union再字形Intersect | ZJB06:12–07:56 | powerpoint-com | 细矩形等距→MergeShapes(1)→与阴影字形MergeShapes(3)；最后棕色外轮廓叠回。1369节点条纹、原生真孔已验证 |

## 从作品修回的规则

1. OfficeCLI1.0.160将prstTxWarp追加在noAutofit后面会validate失败，PowerPoint可能忽略变形。局部raw-set移除noAutofit再追加到bodyPr末尾；正确顺序为prstTxWarp→noAutofit。修后validate与原生斜基线都通过；不能靠“原生打开成功”略过schema。
2. MergeShapes实测需要Open(...WithWindow=-1)使演示文稿有窗口。无窗口时返回UnexpectedHRESULT。仅关闭自己打开的presentation，不退出共享PowerPoint。
3. PowerPoint Duplicate自带偏移，复制后立即将Left/Top复位到源形状，再施加设计所需偏移；否则多次复制会出现错位描边。COM setter给Left/Top传single，避免Double→Single封送失败。
4. MergeShapes返回void；记录操作前IDs，从操作后新增ID取新形状，不能假设返回值就是结果，也不能假设它一定排在最后。
5. Enum MsoMergeCmd本机PIA核实：Union1、Combine2、Intersect3、Subtract4、Fragment5；不根据界面顺序猜数字。Range用明确名称数组，PrimaryShape用实际COM对象。
6. 初版封面正文换行挤近按钮：缩短文字后重验。网格对齐不等于文本没有溢出；实测内容改变后再检查实际行数。

## 当前代表作品

- foundations：4页，网格手机封面、等距时间轴、字体语气、主题色板；v2正文已修。
- font-pairing：3页，Georgia/Calibri、Georgia/Arial、Times New Roman/Arial。
- font-hierarchy：1页，Georgia/Garamond。
- typography-polished-v2：1页，Script MT Bold替代Milkshake，原生路径、白边、相减高光、条纹阴影。

这些是机制复现，非原教程完整美术资产复刻；本组是静态技巧，以本机PowerPoint原生静态图验收，不硬加动画。

## 执行入口

所有路径由调用方传入；输出使用新路径，拒绝覆盖。构建只用Python标准库编排OfficeCLI，不用python-pptx。需Windows PowerPoint才能执行COM精修和原生导图。

```powershell
python scripts/type_build.py --officecli "E:/ppt skills/tools/officecli.exe" --outdir "E:/work/type-new" --mode all
powershell -NoProfile -File scripts/type_polish.ps1 -InputDeck "E:/work/type-new/typography-before.pptx" -OutputDeck "E:/work/type-new/typography-polished.pptx" -Evidence "E:/work/type-evidence-new"
python scripts/type_verify_geometry.py "E:/work/type-new/typography-polished.pptx" --report "E:/work/type-geometry-new.json"
powershell -NoProfile -File scripts/type_native_verify.ps1 -DeckPath "E:/work/type-new/foundations.pptx" -OutDir "E:/work/type-native-new"
```

`--mode`可选all/foundations/pairing/hierarchy/typography。实际命令中的脚本路径应相对本Skill目录解析，输入输出使用绝对路径。OfficeCLI与PowerPoint所需字体必须本机可用；先查字体，缺失不自动安装。原教程Milkshake缺失，验证作品明确用Script MT Bold 150pt替代原160pt；Montserrat等例子也用本机字体进行机制验证。原生输出中可见字体差异，不宣称逐像素一致。

COM执行需与其他正在进行的PowerPoint自动化排队；只关闭自身打开的演示文稿。MergeShapes需可见窗口，并选择目标slide，尤其跨页合并。

## 最低验收

1. OfficeCLI validate零错误；文本不能裁切、碰撞或缺字。
2. 原生PNG复查网格、时间轴、字体层级；字体配对矩阵是启发式，不输出客观审美分数。
3. Typography必须真原生轮廓：black-front具有正负绕向闭合路径形成孔，零图片替代；原生PNG确认孔与高光条纹。
4. 修改文字、字体、字重或变形后重新精修并重验；before保留可编辑文字，polished转路径后不能再直接改字。

研究和首轮原生证据：`E:/ppt skills/distill/youtube-type-foundations/test-report.md`。此绝对路径仅为本次工作证据，不是运行依赖。封装后的复验见同研究目录`packaged-test-report.md`。
