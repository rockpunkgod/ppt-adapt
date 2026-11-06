# 原生可编辑多色软色场

来源：50P BV1oPHgeiEgb P24 1875.5–2067.5，指定段完整字幕与1948/1982/2013/2031操作帧核验。源在Photoshop白底/透明层上先画暖黄与冷青两大面，再加绿/青绿/蓝/紫/紫红局部软刷过渡。此配方还原**多个可定位颜色贡献的层叠设计机制**，不称PowerPoint径向渐变与Photoshop刷头算法相同。

backend：**officecli-native**。全部形状、颜色/alpha停点与排版由OfficeCLI制作；PowerPoint只读重开、读真实对象和导PNG。没有照片/位图背景，没有外部渲染。

## 使用规则

1. 冷暖两个大面先定位，再放中间局部色。用独立圆形渐变shape控制中心/直径/颜色/强度，避免把多色二维背景压成一个双色线性渐变。大圆可在画布外，画布边缘原生裁切。
2. `radial`高层可创建原生circle path fill；per-stop alpha用OfficeCLI raw-set补充。同一个stamp所有stop保持同色，只改变alpha，避免色彩插值与alpha插值混在一起。
3. 本机原生探针确认**位置0是中心**。圆形shape的圆边对应约70.711%的包围盒对角半径：采用0/.17678/.35355/.53033/.70711/1停点、相对alpha1/.72/.35/.10/0/0，使圆边先达到透明。初始逆方向探针留下硬圈，不能复用。当前仅圆形标定，不外推任意非等比椭圆/偏心focus。
4. 混色为普通normal source-over：`Cout=Cb*ab+(Ca*aa+Cbg*(1-aa))*(1-ab)`。AB与BA结果通常不同；不能以透明度叠加冒充Multiply。利用single A、single B独立控制估计实际alpha，再验AB/BA，两种顺序都测。
5. 局部改色要改目标shape每个stop的RGB并保留alpha/位置/几何/层序。直接设置solidFill会丢失软场；重建渐变时也不要只改首stop导致彩色环。本工程第3页仅把mid-lime从绿改淡紫，其他色场保留。
6. 内容区可放半透明浅底保证正文可读，但其颜色不应参与底层混色公式对照。当前验收对照页是纯白底，不混入卡片。

## 从包内独立制作与核验

依赖：OfficeCLI1.0.160、本机PowerPoint、基础Python；独立像素检查需Pillow/NumPy。无Shapely/联网素材依赖。设定`$skillRoot`为技能目录，`$taskPython`为基础Python，`$taskPixelPython`为含Pillow/NumPy的Python，`$taskCli`为OfficeCLI，`$taskOut`为不存在的新输出目录。

```powershell
& $taskPython "$skillRoot/scripts/soft_field_build.py" --officecli $taskCli --out $taskOut --seed 18
& "$skillRoot/scripts/soft_field_native.ps1" -Deck "$taskOut/soft-color-field.pptx" -Evidence "$taskOut/native"
& $taskPixelPython "$skillRoot/scripts/soft_field_verify.py" --deck "$taskOut/soft-color-field.pptx" --model "$taskOut/model.json" --native "$taskOut/native" --out "$taskOut/pixel-check.json"
& $taskCli validate "$taskOut/soft-color-field.pptx" --json
```

查看三张原生PNG：原创多色场、single A/B与AB/BA顺序、单层改色。shape Type1、Fill.Type3、6个GradientStops，两个native只读重开均保留。脚本拒覆盖，不退出共享PowerPoint；静态配方不宣称做了动态实播。

范围：默认seed18及本配色/圆形停点/正常叠加/单层改色；其他种子、尺寸、透明曲线、椭圆、焦点、混合模式须重新通过像素与原生检查。此机制应并入颜色/原生渐变用途，不作为重复的大Skill。

状态：包内fresh通过，详见validation/forward-soft-color-field/README.md与distill/soft-color-field/source-notes.md。研究阶段最大normal公式RGB差1.873（容差3），正反序最大差64；局部色修改支持圆外0变。

包内fresh从零生成三页，240项native stop值/6径向alpha/9位置正反叠加/局部改色均通过；三PNG逐页复看、3拒覆盖、脚本与PPT SHA不变、schema通过。
