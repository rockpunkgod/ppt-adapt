# 柔边半椭圆与折页分界

backend：`hybrid`。OfficeCLI制作柔边椭圆与最终页面，本机PowerPoint将自己的椭圆原生导成透明PNG，OfficeCLI裁半、邻接同底色矩形后组合旋转。不是普通均匀矩形阴影，也不需要外部渲染器。

来源：[BV1qj411Y7Up](https://www.bilibili.com/video/BV1qj411Y7Up/)，25–64秒。字幕和六张分布整帧已核对：灰色椭圆去轮廓→柔化边缘→粘贴为图片→裁半→邻接背景同色矩形→组合旋转→数字置底。下列使用自拟文案、尺寸、色值，不复制作者海报。

## 执行

```text
python scripts/shadow_fold_build.py --stage base --out NEW/base.pptx --officecli OFFICECLI_PATH
powershell -NoProfile -File scripts/shadow_fold_raster.ps1 -Deck NEW/base.pptx -Png NEW/ellipse.png
python scripts/shadow_fold_build.py --stage final --shadow NEW/ellipse.png --out NEW/fold.pptx --officecli OFFICECLI_PATH
python scripts/shadow_fold_verify.py NEW/fold.pptx --report NEW/audit.json
```

输出必须为新路径。基础文件保留4×14cm原生椭圆、18pt柔边；PNG仅将阴影素材栅格化，标题、数字和衬底仍是可编辑对象。半图宽2cm高14cm、右裁50%，与同底色矩形紧贴后整体旋转32°。这些是本样参数，原视频未给清晰稳定的全部数值，不当作逐项抄录。

背景颜色改变时连同衬底一起改。灰阴影在深红底会形成亮灰条，因此深色示例用原生图片亮度-100转黑，再设50%不透明度；白底仍保留灰色。只换底色而不检查合成效果会破坏阴影。

## 原生导出的比例陷阱

PowerPoint Shape.Export的ScaleWidth/Height是相对整页的导出尺度，不能直接当目标shape像素。最初填400/1400得到63×1377的细长图，实际比例失真，已排除。改为整页宽高使用同一两倍倍率，得到302×1059 RGBA，4:14比例误差约0.19%。再裁半后不改变图像比例。改倍率后仍检查原图和裁片比例。

独立验证读取真实嵌入PNG的尺寸/透明通道、srcRect裁切、图片显示长宽比、组旋转及红底亮度/alpha，不接受只看命令成功。该检查不证明视觉结果，仍须本机PowerPoint查看最终静态页。

## 通过范围

白底和红底两页原生画面已看，schema与issues均无误；真实灰柔边、同色前景和数字遮挡共同形成斜向分界。红底经历灰条→过重黑条→50%合成修正。示例是静态设计，不声明源演示中额外动画已复现，也不为了静态页增加无意义动画测试。

包内三脚本在`validation/forward-shadow-fold`另生成新椭圆底稿、新透明PNG与新成品，未复用研究成品。比例/裁半/旋转/亮度/alpha独立审计通过，最终两张本机PowerPoint原生PNG再次逐张复看通过。
