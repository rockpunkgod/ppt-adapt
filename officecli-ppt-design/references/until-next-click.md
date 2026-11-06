# 下一次单击时保持当前相位

backend: **hybrid**。OfficeCLI制作原创可编辑转盘，PowerPoint COM添加Spin与后续OnClick Fade，再用OfficeCLI保留原生的UntilNextClick停止条件。不用预定角度、固定重复次数或到时截断代替用户点击。

第一点击启动1秒/360°Spin；第二点击结束重复、保持此刻角度，同时让下一个Fade显示状态。两次物理Left回退后，点击可重新开始。第二页使用UntilEnd作对照：同样点击显示状态，转盘继续旋转。

## 原生停止条件

已通过本机真正Timing/Repeat下拉分别制作UntilNextClick和UntilEnd并保存重开。两者COM的RepeatCount都回读−2147483648，RepeatDuration也同为0，因此不能靠这个值区分或猜特殊枚举。

两者Spin父节点都有`repeatCount=indefinite`、`fill=hold`。UntilNextClick另有：

```xml
<p:endCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:endCondLst>
```

该条件位于Spin父节点的`stCondLst`之后、`childTnLst`之前。只修改原生Spin的父条件，保留旋转行为、目标、保持状态以及后续独立点击组；没有此条件的UntilEnd仍会继续旋转。不要把它放到整条时间轴，或把原生`onNext`误写成任意对象的指定触发。

## 制作

```powershell
python scripts/until_next_build.py --officecli OFFICECLI --out NEW/base.pptx
powershell -NoProfile -File scripts/until_next_polish.ps1 -InputDeck NEW/base.pptx -OutputDeck NEW/until-next.pptx -Native NEW/native -Officecli OFFICECLI
```

使用新输出。脚本限两个无既有动画的原创基础页，建立各自唯一的Spin与Fade；不覆盖已有业务稿的完整时间轴。Windows PowerShell向CLI传递XML时，片段内属性用单引号，避免双引号丢失；写入失败应停止，不能继续录屏并按成功文件处理。需要Windows桌面PowerPoint、Windows PowerShell、OfficeCLI、Python及Noto Serif SC/Noto Sans SC与Arial。

## 已实播范围

`examples/until-next-click/`包含可编辑PPTX及两个原生片段拼接的播放文件：第一页32秒，六次真实输入覆盖启动、停止、Left两次、重新启动和再次停止；第二页15秒，两次点击检验UntilEnd对照。24个分布帧与逐帧相位均已检查。

第一次停止后的2.3秒采样角度变化不足0.021°，停止角距初始相位约156°；第二次停止后的5.2秒采样保持不变，角度亦远离周期端点。两次Left的原生ClickIndex为1、0，回到起始角后重新点击恢复运动。对照页显示状态后，5秒内仍实际转动约1800°。不是用画面消失、播放器停止或刚好一整圈结束冒充保持当前相位。文案为预置状态，未做转盘分区识别。

OfficeCLI验证零错。主片段完成后立即关闭文稿曾报错，已通过同一任务实例单独关闭测试稿，确认仅任务宿主剩余、无放映；对照片段加入退出后的短等待后正常收尾。早期前台变化、停止条件写入失败的试验均未计通过。

只验证当前Spin/Fade、普通Next点击及两次Left路径。不外推任意动画、指定按钮触发、跨页跳返或其他播放器；新的组合仍须完整实播。

制作依据：[完整大师课](https://www.youtube.com/watch?v=9Ewb0Rd3P3E)4:36:04–4:38:19。复读完整归档对应转写及核看归档转盘/Result操作帧；0.5秒Wave菜单帧不作为此处已选UntilNextClick的证据。本轮以本机真正选项制作和原生播放验证，未计新增在线播放覆盖。
