# 跨底色局部字形与真实剪角

backend：`hybrid`。OfficeCLI生成基础文字和双色版式；本机PowerPoint COM求交/剪除，形成真实可编辑矢量轮廓。静态设计以原生PNG验证，不硬加动画。

来源：[BV1VTDGYsEmd P09](https://www.bilibili.com/video/BV1VTDGYsEmd/?p=9)759–939秒：编号复制、下半区域矩形求交后原位重叠，跨不同底色局部换色；[P08](https://www.bilibili.com/video/BV1VTDGYsEmd/?p=8)1392–1674秒：03右下角真实剪除，再加独立细斜线。本项目另核1392/1450/1550/1620/1674/1700秒完整帧。当前示例用Arial和双色背景重建机制，未复制源字体、照片或模板。

## 已通过操作

1. `01`保留原文本。Duplicate后明确恢复副本Left/Top，避免PowerPoint默认复制偏移。副本与从y=12.1cm开始的矩形执行Intersect3，原文字作为PrimaryShape；结果填A57119、无线，叠回原位。下部为真实字形曲线，顶部仍为原文本。
2. `03`先与全页矩形Intersect3，文字转为真实轮廓；以右下角RightTriangle8并水平翻转执行Subtract4。再添加独立1.2pt斜线。最终没有遮盖背景色的三角形，切口可透出背景。
3. 使用可见文稿窗口，Activate、Normal9、选中目标页，取真实ShapeRange并Select，再以实际形状对象作为PrimaryShape。不得用有文字的普通矩形冒充字形转换，也不将光栅图充当可编辑轮廓。

## 使用

需要Windows本机PowerPoint。基础构建只用Python标准库调用OfficeCLI，并依赖同包`build_components.py`。所有输出明确给定且拒覆盖。`PY_PIL`表示已有Pillow的验证Python，Pillow只读原生PNG，不生成PPT。

```powershell
python scripts/glyph_slices_build.py --officecli OFFICECLI --out NEW/base.pptx
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/glyph_slices_polish.ps1 -InputDeck NEW/base.pptx -OutputDeck NEW/final.pptx -Evidence NEW/native
PY_PIL scripts/glyph_slices_verify.py NEW/final.pptx --native-evidence NEW/native --report NEW/checks.json
```

## 原生前向与失败边界

最终新前向：`E:/ppt skills/validation/forward-glyph-slices-v2`，全新包内builder生成两页→同包精修→4张before/after原生1600×900PNG→独立结构和像素检查→最终OfficeCLI validate均通过。两张最终页已人工复看。

- 第1页上部原字形区域逐像素完全未变，分界以上金色像素0；下部金色10701像素，原下部白字10815像素，轮廓面积差约1.05%，在抗锯齿容差内。
- 第2页剪除减少3722金色像素，仍保留41619金色像素；结构确认为customGeom、无原文本/三角形遮盖、斜线独立，全部页面无光栅图片。
- 像素阈值限定当前固定字体/字号/版式及1600×900导出。换数字、字体、边界或布局后重新生成并检查，不直接复用这些像素坐标。

封装前向曾在旧PowerPoint PID11080连续E_FAIL，视图1/9、名称/索引Range均失败，官方ShapesIntersect也disabled。随后同进程出现正式“可能导致PowerPoint不稳定，请保存并重启”的NUIDialog。主agent确认0文稿/0放映后正常退出旧server并在正确仅子进程环境启动新server；同一基底、未修改的精修脚本随即完成全部Boolean，新构建v2也独立成功。这支持本次会话状态是关键条件，不声称所有E_FAIL都同因。

精修脚本遇错立即停止并关闭自己的文稿，不自动Quit/杀死共享PowerPoint、不无限重试。若出现正式不稳定提示，应先保护现有文稿并按当前授权协调维护。诊断和原型成功稿均不替代最终fresh前向证据。

轮廓可编辑路径/颜色，但不能继续把它当普通文本改数字。可编辑原文字保留在`base.pptx`，改数字后从基础稿重建。第1页原白色文本仍保留；第2页原数字已转轮廓。此配方不宣称文字动态继承背景色、自动随背景换色或任何动画通过。
