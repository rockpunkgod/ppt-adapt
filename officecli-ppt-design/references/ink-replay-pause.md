# 原生 Ink Replay 与末尾停顿循环

backend: **hybrid**。OfficeCLI建立页面和自制静音，标准库按官方InkML格式加入原创笔迹及本项目自产的原生Replay时序；PowerPoint COM建立媒体书签、交互序列并移动原Replay行为。最终对象是真Ink，不是普通线条的Wipe。

来源：[完整大师课](https://www.youtube.com/watch?v=9Ewb0Rd3P3E)4:57:21–5:01:59。复读对应完整转写，核看4:57:35、4:58:08、4:58:35、4:59:07、5:00:59操作帧；分布样本不等同连续观看。作者先画笔迹，选择动画Replay并设5秒，再以7秒静音的起点书签驱动它，留下2秒完整画面的停顿。

## 操作与映射

1. `ink_pause_build.py`用OfficeCLI建立四个点及嵌入式7秒PCM静音。`close`明确落盘后再扩展InkML，避免COM读到resident尚未写完的空稿。
2. 原创三条trace使用设备坐标，T为整数毫秒，范围0–5000；ContentPart引用真正的InkML部件。InkML为L3文件扩展，OfficeCLI当前没有原生`add ink`命令。不要把一张笔迹图片换名当Ink。
3. 自产的2.2KB时序种子含原生preset63入口和`drawProgress`从0到1。只保留本项目自制对象的行为和占位ID，构建时重新绑定；不捆绑作者课件或笔迹。
4. COM为静音起点加InkCycle书签、循环播放，在交互序列加极短且位于页外的anchor；将原Replay `MoveAfter(anchor)`，再明确`Effect.Shape=原Ink`。旧COM接口把Replay读成EffectType0，按0重建会丢失drawProgress。
5. 移动行为后，重新按索引获取交互序列的Ink效果，设5秒、WithPrevious/零延迟，删除主序列中同一Ink的残留入口。静音7秒与动画5秒分别核验，再以实际放映确认重画和停顿。

## 运行

```powershell
python scripts/ink_pause_build.py --officecli "E:/ppt skills/tools/officecli.exe" --outdir "E:/work/ink-new/base"
powershell -NoProfile -File scripts/ink_pause_polish.ps1 -InputDeck "E:/work/ink-new/base/ink-before.pptx" -OutputDeck "E:/work/ink-new/ink.pptx" -Evidence "E:/work/ink-new/native"
python scripts/ink_pause_verify.py "E:/work/ink-new/ink.pptx" --before "E:/work/ink-new/base/ink-before.pptx" --report "E:/work/ink-new/structure.json"
powershell -NoProfile -File scripts/ink_pause_record.ps1 -Deck "E:/work/ink-new/ink.pptx" -Evidence "E:/work/ink-new/play" -Ffmpeg "C:/path/to/ffmpeg.exe" -Seconds 29
python scripts/ink_pause_analyze.py "E:/work/ink-new/play/native.mp4" --ffmpeg "C:/path/to/ffmpeg.exe" --report "E:/work/ink-new/analysis.json"
```

路径必须显式指定、拒绝覆盖。原生窗口串行使用；脚本只关闭自己的文稿，不退出共享PowerPoint。生成路径不依赖Draw配置、系统剪贴板或软件笔输入。录屏以自己的全屏放映窗口做前景守卫和桌面裁切，硬件加速放映不能用静态PrintWindow代证。

像素分析只针对本原创三trace示范，并结合接触表目视复核。它按影片尺寸计算16:9页面居中后的笔迹区域；不是任意PPT的通用动画验收器。改变笔迹、颜色或时长后应调整探针并重录。

## 证据与当前状态

原型`distill/audio-bookmark/ink-replay-v6.pptx`在29秒原生桌面录屏中通过。`ink-pause-recovered-v1/analysis.json`检测四次重画6.467、13.467、20.533、27.500秒，间隔7.000、7.066、6.967秒；后续完整周期约4.83–4.90秒到近完成，保持约2.07–2.20秒。近完成阈值涵盖末尾4%绘制，略早于完全画完；准确5秒/7秒另由原生时序和嵌入PCM审计。58幅分布帧已目视确认绘制→保持→清空重画。影片0秒晚于放映Run，首轮被截断，验收采用后续完整周期。

维护前同一单次Replay稿曾全程静态；清理空自动化实例并修正其子进程环境后，同稿6秒录屏明确逐笔绘制。两项条件同时改变，不能推断唯一原因。随后29秒循环是在另一次自动恢复的PowerPoint会话中通过。遇到原生错误提示或RPC断连应保留失败证据，停止该轮；不要仅由XML存在推断通过，也不要重启含有用户文稿的共享实例。

**包内独立前向已通过**：`verification/ink-pause-forward/`直接运行包内脚本从头生成OfficeCLI底稿、新InkML、新成品、原生PNG及29秒真实桌面影片。schema零错，独立12项审计通过；原生保存后的一阶/二阶差分InkML已按W3C解码，183个XYT采样逐点等于构建前数据。影片四次重画6.400、13.467、20.567、27.533秒，间隔7.067、7.100、6.966秒；三个完整周期均4.9秒到近完成，后续保持2.13、2.17、2.03秒。58幅分布帧与静态原生图均已目视；`analysis.json`通过。

官方格式：[MS-ODRAWXML Ink内容子集](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-odrawxml/096dacae-0d2c-4861-bc4d-c8e4c6405ad3)、[W3C InkML trace及差分编码](https://www.w3.org/TR/InkML/#trace)。原生接口：[AddTriggerEffect](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.sequence.addtriggereffect)、[MediaBookmarks.Add](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.mediabookmarks.add)。
