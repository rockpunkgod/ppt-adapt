# 真标题／正文占位符：层级继承与本地覆盖

backend：hybrid。OfficeCLI准备原创内容底稿；PowerPoint COM建立真Title/Body placeholder及master/layout样式，执行真实SlideReset，再按关系、实际字体和原生像素验证。不是在每页复制相同字框。

来源BV1VTDGYsEmd P23 228–435s、747–1182s，已读对应OCR并核270/390/798/936/1062/1170s整帧。源重点是布局影响普通页、真正占位符、旧内容以无格式纯文本迁入；本机Reset/direct/换布局实验是新增边界核验，不冒称源逐步都演示过。主题字体/主题色引用不属于本配方。

## 操作及真实边界

1. 从CLI源标题和两段正文取纯文本，应用A/A/A/B两个真实CustomLayout，取得已实例化的Title1/Body2；仅在真placeholder内赋文本，正文IndentLevel=1/2，移除旧普通标题/正文框，保留本地脚注。每页应仅有2个真placeholder和1个普通脚注。
2. **两类枚举不能混用**：PpPlaceholderType的Title=1、Body=2；PpTextStyleType的TitleStyle=2、BodyStyle=3。OfficeCLI空master先建立真实Master Body2，再设置TextStyles.Item(3).Levels(1/2)。本机研究仅改正确枚举仍未落盘，新增真实Master Body后style及原生继承才成功；不能只看setter未报错。
3. A布局标题30pt蓝色→36pt青色/新位置，正常A页1/3同步，B页4不变；页2局部标题40pt橙色与自定义位置保持覆盖。母版正文级别1/2由24/20pt改26/18pt，两布局对应段落同步，正文层级仍为1/2。
4. 覆盖按**具体属性**判断：页2标题40pt和橙色是真rPr覆盖。正文第一段仅橙色持续覆盖；设置的24pt等于初始继承值，本机没有保存冗余sz，因此母版更新后仍跟随变26pt。不要把一次Font.Size赋值调用等同永远有显式字号覆盖。
5. 页1/2换到B布局后，正常标题采用B的28pt绿色；页2direct字号/颜色仍保持，但其几何位置按B重排。换布局会重排哪些属性须按实际对象检查，不能承诺所有本地位置始终保留。
6. 选中自有页2，执行官方Office命令表的SlideReset（6264）；这次标题和正文direct格式被清除，回B标题及当前母版正文样式；本地标题/两段正文/脚注和正文级别全部保留。不是把换CustomLayout冒充原生Reset，也不删除文本重建假装重置。

## 复现入口

构建/结构核验仅Python标准库，像素核验用已有Pillow环境；原生步骤需Windows PowerPoint，串行操作自有文稿。全部输出显式指定，新路径拒覆盖；COM仅关闭自己打开的文稿，保留共享应用和其他文稿。

```powershell
& $Py "$Skill/scripts/placeholder_build.py" --officecli $OfficeCLI --out "$Out/base.pptx"
& "$Skill/scripts/placeholder_polish.ps1" -Deck "$Out/base.pptx" -Evidence "$Out/native"
& $Py "$Skill/scripts/placeholder_verify.py" --base "$Out/base.pptx" --evidence "$Out/native" --report "$Out/structure.json"
& $PillowPy "$Skill/scripts/placeholder_pixels.py" --evidence "$Out/native" --report "$Out/pixels.json"
& $OfficeCLI validate "$Out/native/placeholder-final.pptx" --json
```

输出五个4页PPTX：before、layout-update、body-role-update、layout-swap、placeholder-final；另有六阶段共24张原生PNG（末阶段为保存重开），states.json记录实际文本/级别/字体/坐标，role-readback.json记录Style即时和保存后读数。最终作品是native/placeholder-final.pptx。

## 验收

独立解析真实slide→layout→master关系、title/body的type+idx匹配，确认真实Master Body2。按实际master/bodyStyle→layout lstStyle→slide/paragraph/run覆盖计算字号/颜色，逐条与COM resolved字体核对；保原始CLI三类本地文本、两段真实层级，验证A/B共享范围、direct属性、换布局和Reset前后。像素核验A改动不污染B、标题更新不改正文、正文更新不改标题、Reset只改页2、脚注各阶段不变、重开全图相同，并核真实继承颜色与橙色覆盖消除。

研究v3：以上结构/cascade、24原生PNG和像素隔离均通过，已目视24状态联系表及受影响页。先前错误枚举v1与无Master Body的v2保留distill/placeholder-inheritance，未当通过；其余未执行的MasterBody假设脚本不作为证据。当前仅四页英文字体/两正文级别、局部字色/字号/位置范围，不外推任意placeholder类型、全部9级、复杂混合文本、中文主题字体或导入任意已有PPT。

包内全新前向已完成：validation/forward-placeholder-inheritance/。四入口从新CLI底稿生成，独立cascade/真实placeholder关系/本地文本与级别全部通过；六阶段24张原生PNG已目视，36项像素隔离/保留及24状态颜色检查通过，最终schema零错、输入底稿SHA不变、拒覆盖通过。原生SlideReset与保存重开均真实执行；完整报告见该目录README.md。
