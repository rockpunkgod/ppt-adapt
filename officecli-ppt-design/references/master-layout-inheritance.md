# 真正的母版与CustomLayout继承

用于一处更新共同标识、按页面角色维护分支样式。backend=`hybrid`：OfficeCLI制作每页内容，COM建立共享母版与真正的CustomLayout绑定。共享标识放master，角色标识放layout，页内内容留在slide；不要把三份独立图形当作继承。

来源[BV1VTDGYsEmd P23](https://www.bilibili.com/video/BV1VTDGYsEmd/?p=23) 453–618秒。510秒整帧可见主母版Logo影响子layout；603秒原画总结“母版影响子模板、子模板元素独立”。这是采样字幕和操作帧核查，未声称逐秒完整口述转写。

## 复现样例

```text
python scripts/bili_layout_master_build.py --officecli officecli --out ./new/master-base.pptx
powershell -NoProfile -File scripts/bili_layout_master_polish.ps1 -Deck ./new/master-base.pptx -Output ./new/inheritance.pptx -RenderDirectory ./new/native
python scripts/bili_layout_master_verify.py ./new/inheritance.pptx --report ./new/structure.json
```

用全新输出路径。基础PPT、最终PPT、两个中间PPT和PNG目录都拒绝覆盖；COM仅关闭自己的文稿，不Quit应用。基础生成器批处理后close/flush并检查磁盘三页，再交外部PowerPoint读取。

脚本保留三种状态：`inheritance.before.pptx`为共同master+A/A/B绑定；`inheritance.master-update.pptx`只更新共同master标识；`inheritance.pptx`再只更新A layout。native目录对应before/master/after各三张PNG。

## 关键操作

1. 在同一`SlideMaster.CustomLayouts`里建立A、B两个命名layout。清除新增layout自动带入的占位符再赋给页；本例只需要固定角色条，避免新增占位对象遮挡原页内容。
2. 页1/2绑定A，页3绑定B；layout和slide都显示master shapes。共享标识仅存在master一次，A/B条各自仅存在对应layout；slide.Shapes仍只有原来的三个本地文字对象。
3. 改master的同一个标识对象，再原生导出全部三页确认更新。随后只改A layout条，确认A两页同步、B仍保持原样。
4. 保存后关闭并重开最终文稿，核对每页原文、对象数、CustomLayout名和最终原生PNG。母版和版式分别由真实关系链连接，不仅凭最终外观判断。

独立verify读取三份PPT，追踪每页slide→slideLayout→slideMaster，验证A/A/B两种绑定、两layout共享同一master及母版反向列出两layout；检查共享标识和角色条所在层、三阶段文本/填色、完整本地正文不变。还需实际看九张原生PNG。

## 已通过范围

三阶段共九张原生PNG已检查。额外像素差分证实：共同master区域在三页均改变、改master时分支条不变、改A时只有A两页分支区改变，B第三页master-update与after整页完全一致；所有本地正文区域逐像素不变。schema和真实继承关系检查通过。

包内三脚本在`validation/forward-master-inheritance`全新前向重建后，上述真实关系、九张原生PNG及独立像素差分再次通过；无包外补修。主agent另查看三状态总览与差分报告。

本原子只验证共享普通母版图形、A/B CustomLayout分支和本地文字保留。主题字体继承、母版动画、参考线锁定及所有placeholder类型属于另外的能力，需要分别测试。样例的页内文字是普通可编辑文本；不能由它推断已验证标题placeholder的主题联动。
