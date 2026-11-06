# 原生3D入场、持续转盘与点击跳跃

backend: **hybrid**。沿用OfficeCLI真`model3d`和PowerPoint COM时间轴；静态GLB也可使用这些原生3D预设，不要求模型自带骨骼动画。已有双面卡片模型制作见[双面GLB卡片](glb-double-sided-card.md)，不重复实现模型引擎。

## 制作顺序与空间

先为模型的完整运动留空间，再排文字。模型跳跃的包络大于静态外框；本例33.87×19.05cm页面，模型位于20/5.5cm，外框6×10cm。早期8×13cm模型位置过高，跳跃越出页面，缩小并下移后重新实播通过。不要只检查静态截图或照搬这组尺寸到不同模型。

同一个`interactive-model`承载三种效果：

1. 主序列Arrive，单击开始，1秒；原生淡入、3D旋转和缩放。
2. 主序列Turntable，与上一动画同时，10秒一圈；关闭加减速，使用原生“连续”。
3. 独立InteractiveSequence的Jump & Turn，点击模型自身，2秒；不放到普通下一次单击组。

本机原生菜单实际创建后的COM类型分别为151、152、154。保存后预设分别为entr/60、emph/37、emph/39；转盘动画属性是`3d.view.rotation.y`，跳跃涉及`3d.object.translation.y`与`3d.object.rotation.y`，不能用二维Spin或图片路径替代。旧PIA枚举未列出这些类型时，按实际原生效果读回判断，不猜编号。

## 连续与指定触发

本机“连续”选项落盘为转盘预设父`cTn`的`repeatCount=indefinite`，同时取消加减速。COM读取`RepeatCount`为INT_MIN，却拒绝把该负值写回；示例先由COM创建原生Turntable，再用OfficeCLI只修改该预设的重复属性，保存重开核实。不要据INT_MIN区分“直到下次单击”和“直到幻灯片末尾”，该区别仍沿用[点击相位保持](until-next-click.md)的标准。

指定触发沿用既有交互序列：`InteractiveSequences.Add().AddTriggerEffect(model,154,4,model,'',0)`。主序列保持Arrive和Turntable，交互组保持Jump & Turn；共享同一模型，但视图旋转与对象旋转/位移承担不同运动。新模型或其他效果组合仍需实播确认是否相互覆盖。

## 运行

```powershell
python scripts/native_3d_effects_build.py --officecli OFFICECLI --model LOCAL.glb --out NEW/base.pptx
powershell -NoProfile -File scripts/native_3d_effects_polish.ps1 -InputDeck NEW/base.pptx -OutputDeck NEW/final.pptx -Officecli OFFICECLI
```

需要已有OfficeCLI、Python、Windows PowerShell及支持3D动画的桌面PowerPoint；字体Noto Serif SC/Noto Sans SC。只读取本地GLB，不下载素材。精修脚本只接受自身的一页新底稿，不清理业务稿已有时间轴。

## 实播范围

`examples/native-3d-effects/`保留可编辑PPTX、原创GLB、40秒原生播放与预览。发布源已完整重建并保存重开：主序列151/152、交互组154、10秒转盘、模型自身触发、无限重复读回一致，OfficeCLI验证零错。

最终录制只有四次计划鼠标点击，无额外输入：约2.28秒开始，8.64/16.26/28.28秒点击模型。实际输入时间包含录制和截图开销，不把计划7秒误报为实际7秒。27个分布画面复看，包含入场、三次跳跃高点、侧面、正背面及末段持续旋转。10fps均匀采样400个画面，三次跳跃模型色区中心由约215–223px上升到约139–141px后返回；所有可见模型样本在640×400画面内的上下界约39–306px，幻灯片有效纵向20–380px。该测量来自重采样，不声称录屏原始帧率恒定。

已测单页、三次间隔点击与约38秒运行，不宣称跨页返回、快速连击、骨骼动画或所有模型都通过。前台被打断的两轮录制作废；早期跳跃越界稿也未入库。参考操作：[原生3D动画章节](https://www.youtube.com/watch?v=9Ewb0Rd3P3E&t=16165s)，4:29:25–4:31:42；已核对Arrive、With Previous/10秒、模型指定触发和Continuous操作画面。
