# 双幅网格的连续滚动循环

backend：`officecli-native`。OfficeCLI创建所有可编辑几何、组合、Fade与原生motion path；COM只读导图/放映验证。没有外部影片背景或逐帧位图。

## 来源与空间条件

[One Skill完整大师课](https://www.youtube.com/watch?v=9Ewb0Rd3P3E&t=20140s)5:35:40–5:40:01。已读此段完整转写并看5:37:48整帧：15列/8行的单幅网格复制到右侧，组合成两页宽；向左移动一页，关闭平滑起止，17秒、重复至页末，另加1秒Fade同起。源交界重叠线稍厚，口述也明确指出。

当前原创样稿用纯色底与单独细条构成15列/8行，双幅共30条竖边/9条横边，避免重复边框叠深。细条厚度等效0.5pt、真实白色alpha20%；不是预混背景色冒充透明度。固定标题在上层。未复制源彩色背景、字体或模板。

连续性依赖**位移恰等于一个图案周期**。本例一幅宽等于SlideWidth；右幅几何等于左幅加SlideWidth。两幅作为同一组平移−1页，17秒后复位时，右幅与左幅显示相同图案。不是任意距离的repeat都能无缝，也不能用AutoReverse实现持续向左。

## 制作入口

```powershell
python scripts/seamless_grid_build.py --officecli OFFICECLI --out NEW/grid.pptx
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/seamless_grid_export.ps1 -Deck NEW/grid.pptx -Evidence NEW/native
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/native_record_states.ps1 -Deck NEW/grid.pptx -Plan SKILL/assets/grid-loop-record.json -Directory NEW/live -Ffmpeg FFMPEG
PY_PIXELS scripts/seamless_grid_measure.py --movie NEW/live/native.mp4 --ffmpeg FFMPEG --ffprobe FFPROBE --report NEW/measure.json
python scripts/seamless_grid_verify.py --deck NEW/grid.pptx --measure NEW/measure.json --report NEW/checks.json
```

构建和结构核验只用Python标准库；构建依赖同包`build_components.py`。原生导图/放映需Windows PowerPoint，录像与时间戳需FFmpeg/FFprobe，像素相位分析需已有NumPy。所有输出新路径、拒覆盖，不安装依赖，不退出共享PowerPoint。

CLI先在透明周期边界形状上建立两条动画，组合40个原生形状后，使用OfficeCLI raw-set将实际spTgt统一改成真实group ID，并只对path效果设repeatCount=indefinite。不能让透明边界自己运动而网格不动。路径`M 0 0 L -1 0 E`以页面为单位，dur17000、无accel/decel/autoRev；Fade仅一次1000ms、同起。大组和右侧一幅本来在画布外；仅这些已定义坐标和半条线越界有意保留，不对普通文字溢出放行。

## 原生前向与证据

`validation/forward-seamless-grid/`：包内全新CLI→原生PNG→54秒真实屏幕→17项独立结构/运动检查→OfficeCLI validate通过。40个可编辑成员、39处真实alpha20000、两个空间周期、全部动画目标group均核对；原生画面与跨周期采样已查看。

实际1440个独立VFR帧，按FFprobe PTS、passthrough解码，无补帧。15列图案的空间傅里叶相位在2–17、17–34、34–51秒窗口持续变化；实测向左速度分别约75.298、75.285、75.288px/s，理论1280/17=75.294。三窗误差均<0.013%，R²>0.99999，整体相位RMS约0.0636rad（约0.86px）。后两窗仍保持相同速度，覆盖三次周期边界，没有在首轮后停住。

傅里叶相位按一格周期存在歧义；本例验证均匀周期网格的可见连续运动，不能推广为任意照片纹理/对象身份跟踪。所有阈值在研究量测前已固定，包内alpha版本重新54秒验证，没有降低标准。早期分析器NumPy bool序列化问题只影响报告写出，修正后按相同影片重算，未冒称失败报告已通过。

录屏0秒晚于放映Run；54秒超过三倍17秒时长，但不是从原生启动首帧开始，不能据此声称精确首秒Fade同步。Fade时长/一次性有结构证据；动态通过范围是持续运动与循环连续性。改变周期、网格间距、组边界、路径或时长后重算位移并重新录制，不能只延长repeat。
