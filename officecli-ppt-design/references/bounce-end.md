# Fly末段 Bounce End：对照、GUI与真实质心验证

backend：`hybrid`。OfficeCLI做可编辑基础页；PowerPoint COM精修原生Fly的末段Bounce End；实际放映录屏量测。它不是独立Bounce入口预设，不用手工路径或视频贴图替代。

## 来源和已测参数

[One Skill PPT · TOP 10 PowerPoint Animations](https://www.youtube.com/watch?v=o12PIUTskS0)。3:32–3:51设置Fly的Bounce End0.8秒；5:11–5:29设置From Top、总时长2秒、Bounce End1.5秒，随后按0.1秒递增Delay。原UI字幕已复读，328–330秒精确操作帧已复看。

配方只有一组已测参数：原生Fly2、Top10、Duration2秒、Delay1秒、平滑起止/加减速0。对照关闭Bounce；变体在Effect.Timing及两个Property行为的Timing上设置BounceEnd=-1、BounceEndIntensity=0.75。赋值使用`[single]0.75`，禁止`[single].75`：后者在本机PowerShell5.1实际为null，曾造成读回0和假阴性。

通过OfficeDev官方controls表里的`EffectOptionsDialog`打开本机“飞入”选项，直接读得AutomationId643为“1.5 秒”，同稿Duration2秒/Intensity0.75，平滑起止0。证据在研究`bounce-end/test-v1/gui-official-command/values.json`及`native-options.png`。只确认这个参数组合；改总时长/方向/版本后仍需复测，不把单点读数当通用秒数换算。

原生保存为mc:Choice现代timing加mc:Fallback线性兼容timing。现代分支含父`p14:presetBounceEnd=75000`和X/Y anim的`p14:bounceEnd=75000`；旧分支保留无回弹Fly。交付保留两分支，审计本机支持的Choice，不能全树双计为4个Fly。旧阅读器若走Fallback，不承诺同样回弹。

## 复现

全部输出显式给定且拒覆盖。精修仅接受无既有动画的全新一页CLI基底。Windows本机PowerPoint和OfficeCLI；FFmpeg/FFprobe用于桌面录制与实际时间戳。基础构建只用Python标准库；`PY_NUMPY`表示已有NumPy的分析Python，NumPy只参与像素量测，未用于生成PPT。

```powershell
python scripts/bounce_end_build.py --officecli OFFICECLI --outdir NEW
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/bounce_end_polish.ps1 -InputDeck NEW/bounce-base.pptx -OutputDeck NEW/bounce-native.pptx -EvidenceDirectory NEW/native
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/bounce_end_record.ps1 -Deck NEW/bounce-native.pptx -Report NEW/desktop.json -OutputDirectory NEW/desktop -Ffmpeg FFMPEG
PY_NUMPY scripts/bounce_end_measure.py --movie NEW/desktop/bounce-end-live.mp4 --ffmpeg FFMPEG --ffprobe FFPROBE --output NEW/centroids.json
python scripts/bounce_end_verify.py --deck NEW/bounce-native.pptx --measure-report NEW/centroids.json --output NEW/checks.json
```

录屏用真实desktop+DWM物理放映范围，显式Arrow1并回读，单页重复开始一次以捕捉完整自动动画；不取title窗口缓存。脚本仅关闭自己的文稿，不Quit共享PowerPoint，和其它原生UI工作串行。

## 量测定义与前向结果

最终exact包内前向：`E:/ppt skills/validation/forward-bounce-end-v2/bounce-native.pptx`。全新CLI基础→包内COM→原生60fps导片→真实桌面→独立13项→最终OfficeCLI validate均通过。`desktop/contact.jpg`9时刻与`centroid-trajectory.svg/png`已人工复看，文字无乱码。

桌面录制请求60fps，实际238个独立VFR帧、平均39.67fps；按FFprobe真实PTS并以passthrough解码，未把补帧当独立样本。固定1280宽实验，自动剔除上下黑边，按两种纯色mask求质心；限定全圆可见、1.0≤PTS<3.8秒为运动窗口，4–5秒中位数为终点，排除录屏开头复位前瞬态。原始轨迹在`centroids.json`。

- 变体终点y=453.12px，1.90秒已过终点线，2.00秒y=508.42px，末段偏离终点55.30px，2.20秒返回终点另一侧，随后减幅收敛。向下为像素正方向；核心判据是实际非单调往返。
- 无回弹对照单调到位，正向偏差约0.004px，低于1px测量容差；两个对象最终均稳定。
- 13项同时检查原生Fly2而非Bounce预设、父/行为Bounce属性、2000ms与1000ms、真实PTS样本、对照单调性、变体非单调性与清晰偏离、终态稳定。参数存在本身不构成视觉通过。

分析脚本仅适用于此固定纯色圆对照版式，不冒充照片/任意复杂对象跟踪器。更换版式、颜色、对象、方向或时长后应调整量测区域并重新放映。GUI辅助探查脚本有一次取消provider异常，已保留研究记录，不随Skill发布；最终生成/录屏/量测不依赖该辅助脚本。
