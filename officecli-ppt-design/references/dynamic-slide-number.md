# 母版样式与真实动态页码字段

backend：**hybrid**。OfficeCLI建立三页基础内容；COM建立母版/版式页码占位符和真实slidenum字段。来源：[PPT完整大师课](https://youtu.be/9Ewb0Rd3P3E)，25:38–27:24，母版设置页码样式、打开Slide Number、复制后自动2→3。源文字已读，视觉只按相关帧对照，不称整片连续观看。

Ego定点1558/1612/1633秒已看，1612秒明确为母版圆形页码字段；当前复现范围是字段、母版样式继承与自动编号，圆框装饰不是本样例外观复刻目标。

`slide_number_field_build.py --officecli <exe> --out <new-base.pptx>`生成三页；`slide_number_field_native.ps1 -Deck <base> -Evidence <new-directory>`生成母版字号42 Arial深蓝样式，通过layout承接、逐页启用页码，导出初始3页、复制4页、重排4页、另存重开4页。只关闭自己的文稿，保留原CLI文件。`slide_number_field_verify.py --evidence <directory> --report <new-report>`独立检查字段/页面身份与真实字形。

制作依赖OfficeCLI、Windows PowerShell与本机PowerPoint；图像验收另需Python NumPy/Pillow。所有输出需新路径，精修入口只接受本例三页基础稿，不是任意已有文稿的页码修复器。

关键限制：从CLI空白版式开始，母版加入sldNum并打开母版/页级HeadersFooters仍不会显示；实际CustomLayout必须有页码placeholder。添加版式placeholder后，页级启用会物化原生字段，复制和重排由PowerPoint更新。不能用手写数字或普通文本框充当页码；本样例额外普通数字只用于独立图像参考。

AddPlaceholder(13)会自带字段，盲目再InsertSlideNumber会变成两个。当前母版先清空Text再插一次、版式保留默认一次；包内验收检查母版、版式及4页各恰有一个slidenum。母版样式由layout继承，不对实际页码逐页硬编码。

原生字形分别与独立1/2/3/4模板比对：正确模板IoU≥.95，并比所有错误数字高>.1；实测正确.989–.998，错误最高约.456。不同纵坐标的抗锯齿最大RGB差14–49，旧“每像素差≤1”检查未通过；保留此失败，改为语义字形判据，不声称两处零像素差。相同位置的保存重开四页则逐像素相同。当前仅验证3→4页、复制第二页及尾页移到首页，不覆盖任意编号起始/隐藏页/分节重编号或页码动画。
