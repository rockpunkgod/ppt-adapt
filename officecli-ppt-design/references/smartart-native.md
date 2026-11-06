# 原生SmartArt节点、层级与图片配对

backend为`hybrid`：OfficeCLI负责基础页面、标题、说明与notes；PowerPoint COM添加真正SmartArt并保留数据模型。不能把普通组合形状、Mermaid生成shape或截图叫SmartArt。

来自[BV1VTDGYsEmd P13](https://www.bilibili.com/video/BV1VTDGYsEmd/?p=13)。986秒操作帧与字幕确认文本窗格Tab降为子级；1670秒确认选中图片后使用图片版式。来源按3秒内嵌字幕OCR和关键帧核查，不声称全课逐秒听完。此配方只通过hierarchy1、hierarchy2、pList1三个布局及下述节点操作，不覆盖整课全部SmartArt效果。

## 先判断内容关系

先分标题与条目，再明确并列/父子关系，选可承载实际字量的布局。图片应与节点标题明确配对；页面旁边独立摆着的文字不保证自动匹配。转为普通形状后失去SmartArt自动布局，转为文字可能丢图片；如需转换，保留原生版。

## 运行固定验收样例

```text
python scripts/bili_layout_smartart_build.py --officecli officecli --out ./new/smartart-base.pptx
powershell -NoProfile -File scripts/bili_layout_smartart_polish.ps1 -Deck ./new/smartart-base.pptx -Output ./new/smartart.pptx -RenderDirectory ./new/native -Landscape /path/wide.jpg -Portrait /path/portrait.jpg
python scripts/bili_layout_smartart_verify.py ./new/smartart.pptx --prepared-dir ./new/native --report ./new/structure.json
```

全部使用新路径，脚本拒覆盖。照片由调用方提供有权使用的JPEG，Skill不捆照片、字体、课件或截图。样例节点文案为产品研究/需求洞察/原型测试/反馈归纳，图片题注为日落/步道；替换主题时同步修改节点输入、图片标签映射和独立校验预期，不把它当自动识图。

基础生成器批处理后显式`officecli close`，再直接读取磁盘ZIP断言三页，之后才交COM。resident的CLI内存validate成功不等于外部程序立即能读到最新文件；漏flush曾使COM读到0页。

## 已验证的操作与修正

1. 本机隐藏文稿下SmartArtLayouts返回null，PowerShell的null.Count为0，不能据此判断机器没有布局。App.Visible=-1且**自己的文稿Open WithWindow=-1**后可读取159布局；脚本从真实目录选择已观察的三个ID，缺失就失败，不猜GUID。脚本仅关闭自己的文稿，不Quit PowerPoint。
2. hierarchy1创建真实四节点，Level=1/2/2/3；插入临时第五节点并删除后恢复四节点。本机插删后保留旧COM引用时Promote断言失败；改为按稳定文本/业务键重新取得当前节点，再Promote/Demote，样例正确恢复了反馈归纳的原型测试父节点。真实内容可能有重名，应使用唯一业务键或另保映射。
3. hierarchy1切换hierarchy2后逐节点比较文字和Level；保存、关闭、重开后再比较，两个布局保持同一信息结构。
4. `Font.Name`只显式写Latin时中文可能仍用其他字体；同时设`NameFarEast`。两个层级图和图片题注的中文字体、22pt字号均已在原生PNG与XML检查。
5. pList1每个节点要求唯一空图片形状候选，按该窗口实际宽高先等比cover裁剪JPEG，再UserPicture填图。图像与标签沿`presAssocID`关联，独立校验追踪到嵌入关系和SHA256，不只看两张图摆在附近。
6. diagram连接的`type`省略时为parOf；独立验证按该缺省构建父子图，不能只统计显式type而误算所有Level=0。保存结果包含真正graphicFrame、diagram data/layout/drawing关系，不是普通group。

## 验收范围

三页原生静态已检查；schema通过，两个层级图的四节点关系与布局ID正确，第三页两个图文节点的嵌入照片hash和窗口比例独立通过，预裁后比例误差小于0.2%。样例没有动画，不要求用影片代替对象与静态检查。

封装后另在`validation/forward-smartart`完全用包内三脚本重建，未包外补修；独立结构校验与3张原生PNG再度通过。主agent亦复看三页，父子结构、字体和图片配对清晰。

换布局/字量/字体后重看原生PNG，不能由三个布局外推全部159种都可自动套用。图片配对通过的是显式节点映射，不保证任意外来PPT中的散落文字自动归属正确照片。
