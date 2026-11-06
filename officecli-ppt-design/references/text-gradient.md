# 真正可编辑文字的多stop渐变与透明端

backend：officecli-native。OfficeCLI直接给a:rPr写gradFill；COM只读取、导图、保存重开，不参与渐变authoring。文本框本身noFill，未用整框渐变、字形布尔转换或位图代替。

## 制作与复验

1. `text_gradient_build.py --officecli <exe> --outdir <fresh-dir>`。生成text-gradient.pptx：深/浅两背景，每页单一OPEN文本框，Arial Black170pt；并生成相同字形的独立白字mask对照text-mask.pptx。
2. `text_gradient_native.ps1 -Deck <gradient> -EvidenceDirectory <fresh-native-dir>`。读取真实Font.Fill和stop，导原生PNG，另存roundtrip.pptx，再打开导同页图。对白字mask也运行到另一个新目录。
3. `text_gradient_edit.py --officecli <exe> --input <native-dir/roundtrip.pptx> --output <fresh-edited.pptx>`。CLI直接把真实OPEN改成OPAL，保留原渐变。再对编辑稿运行同一native脚本到新目录，目证新A及字孔。
4. `text_gradient_verify.py --deck <gradient> --native-dir <native> --mask-dir <mask-native> --edited-dir <edited-native> --output <fresh-json>`。需要Pillow/NumPy；builder/editor只用Python标准库。

全部输入输出用绝对路径，拒绝覆盖已有输出。这里是静态设计技巧；验证原生可读字形、透明端、保存重开和真改字，不硬加动画。

## 关键属性

`textFill=F26F4E@0-9E81EA@30-67CFD2@55-67CFD2+alpha0@85-67CFD2+alpha0@100-0`

5stop分别0/30/55/85/100%，后两stop透明青；最后0是横向0°。使用明确+alpha0。实际CLI探针发现8位hex输入按ARGB解释，不要把67CFD200臆测为RGBA透明青；Get返回的8位显示规范值也不能盲目当输入语法重放。

不是给shape.gradient赋值：应在XML看到p:txBody/a:p/a:r/a:rPr/a:gradFill，形状p:spPr为noFill。编辑后仍是a:t真实文字，字孔露底并非画一块背景色遮住。

## 证据与限制

源Masterclass9Ewb0Rd3P3E，1:00:43–1:01:46：文字渐变、多色标及透明端；源1:01:31–38白色单词/白影为可选层，本配方未声称该效果已过。Ego精确源3678s明确Text Options→Text Fill→Gradient Fill，Linear/90°；本原创横向OPEN使用0°，不是逐值复制源。3694s右侧已切为Format Background，不能当文字菜单证据。

原型48项已过：原生重开5stop、可编辑真文字；两背景RGB差反解字形透明度中位数1.000→0.625→0.000，独立白字mask找出8592像素真实封闭字孔，两背景误差0。首开/重开逐像素一致；保存后CLI改OPAL再次原生打开，A真实出现且渐变仍在。包内4脚本从头构建的exact前向 validation/forward-text-gradient 同样48项全部通过，基础/原生重开/改字后原生重开三份PPT均OfficeCLI validate无误，12张原生PNG已生成并复看关键深浅/改字页。未复用研究PPT。

字体依赖Arial Black；替换字体会改变宽度、字孔和采样区域，需重新导图验证。不同渐变角度/stop布局也需复测；RGB差法是本例两纯色背景下的局部透明度对照，不是从一张PNG假定提取alpha通道。
