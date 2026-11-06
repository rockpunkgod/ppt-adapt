# 主题色槽、EA字体引用与直接格式隔离

backend：`officecli-native`。主题、字体引用及页面全部由OfficeCLI高层命令制作；COM仅用于本机导图、保存重开及读回，不参与authoring。不是再次复制母版Logo。

## 来源和设计规则

[27P P26](https://www.bilibili.com/video/BV1VTDGYsEmd/?p=26)660–711秒明确：采用主题色的对象可统一换色，标准/特定色不能；[P23](https://www.bilibili.com/video/BV1VTDGYsEmd/?p=23)1602–1638秒展示设计菜单中的整体颜色/字体设置。[One Skill大师课](https://www.youtube.com/watch?v=9Ewb0Rd3P3E)14:16–16:05建立heading/body字体对及accent1–3。新增精确源帧878秒显示双字体槽，925秒显示accent1–6；输入中的字体/色值不拼为最终定稿。P18 Logo取色/画布外调色板只是配色流程，不能当主题引用语义证据。

希望品牌修改时统一联动的颜色/字体使用主题引用；必须保留的局部强调使用直接RGB/明确字体。标题major、正文minor，以及中文EA与Latin脚本槽各自有范围；不要把一个裸font属性误当所有脚本和角色的全局开关。业务状态红绿若必须固定，不应仅因换品牌主题就被任意替换。

## 已测OfficeCLI映射

在presentation根路径`/`设置：

```text
theme.color.accent1=1C6561
theme.font.major.eastAsia=Microsoft YaHei
theme.font.minor.eastAsia=Microsoft YaHei
theme.font.major.latin=Arial
theme.font.minor.latin=Arial
```

对象填充/文字颜色采用`fill=accent1`或`color=accent1`，实际存为`a:schemeClr`；直接`fill=1C6561`存为`a:srgbClr`。中文标题分别写`font.ea=+mj-ea`、`font.latin=+mj-lt`，正文写`+mn-ea/+mn-lt`，中文run语言`lang=zh-CN`。直接格式使用明确的微软雅黑/Arial字体名。

两页三状态：base→只换accent1为AE4E32→保留换色且只改major EA为SimSun。三份页面对象解析树不变，仅theme参数变更；无需逐页重写RGB或逐个替换字体。未测试的theme色transform、缺失字体fallback、多master/themeOverride、chart/SmartArt主题响应不能由本例代证。

## 复现与原生验收

全部输出显式给定，已有目录/报告拒覆盖。构建使用Python标准库及OfficeCLI。`PY_PIL`指已有Pillow的验证Python，Pillow仅分析原生PNG。

```powershell
python scripts/theme_semantics_build.py --officecli OFFICECLI --outdir NEW
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/theme_semantics_native.ps1 -InputDirectory NEW -EvidenceDirectory NEW/native
PY_PIL scripts/theme_semantics_verify.py --input-directory NEW --evidence-directory NEW/native --output NEW/checks.json
```

原生脚本打开三稿各两页，导图、SaveAs到新文件、关闭重开再导图，共12张1600×900PNG；只关闭自己文稿，不退出共享PowerPoint。此为静态主题语义测试，没有动画或录屏通过的扩展声明。

最终exact前向：`E:/ppt skills/validation/forward-theme-semantics`，全新包内构建、原生验证、146项独立结构/像素/读回、三份重保存PPTX validate均通过，12图已复看。

- accent1引用绿→赭红，跨两页的引用对象联动，直接绿卡/绿线保持；原生ForeColor.ObjectThemeColor为5，对照为0。
- major EA微软雅黑→宋体，中文标题ROI变化6867像素；minor正文、Latin以及direct控制ROI均0变化。原生NameFarEast解析SimSun，minor/direct仍微软雅黑，NameAscii仍Arial。不能仅凭中文全角文本总宽度判断字体是否改变。
- 6对首次/保存重开PNG逐像素相同，重保存XML保留`schemeClr`及`+mj-ea/+mn-ea/+mj-lt/+mn-lt`，没有被转成直接色/固定字体。
- XML命名空间声明可被提升到根，比较解析后的对象树，不能把xmlns位置变化误判为局部对象被重写。

已测字体为本机安装的Microsoft YaHei、SimSun、Arial，主要中文zh-CN和Latin en-US。更换字体、脚本、语言或复杂主题继承后重新确认引用、实际解析字体及原生结果；不宣称包含字体文件或保证缺失字体时同样外观。
