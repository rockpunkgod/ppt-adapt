# 连续路线：双向时间轴

`push-return.pptx`为三页可编辑的路线演示，以观察、测试、改进和决策连接三个章节。每页保留原生开放路径和文字；共享底色、同色3pt路线及同x的跨页端点，让原生Push在前进与返回时都能连接画面。

本机PowerPoint以所有幻灯片模式播放，Right、Right、Left、Left重复两轮，实际路线1→2→3→2→1→2→3→2→1。`playback.mp4`保留44秒连续录制；8次物理按键、无额外输入。`preview.jpg`展示每次运动的早、中、晚状态。

复用既有[Push制作方法](../../officecli-ppt-design/references/push-timeline-seam.md)，没有改写原有制作脚本或结构验收标准。重建只需运行`push_seam_build.py`并传入本地OfficeCLI路径和新的输出文件；不需要下载素材。中文改稿、斜线接点、纹理底色和快速打断需重新预览；本例不代表任意Push或其他导航入口均已通过。
