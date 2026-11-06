# C20：完整2.5D笔记本铰链与打开页视频层

backend：hybrid。原创机身、48键、触控板、屏框均为OfficeCLI原生形状；COM精修两组透视相机和Morph/媒体动画。视频是自产6秒程序测试片，未用作者模板、Freepik照片或库存视频；FFmpeg只产生测试资产，PPT场景不是外部渲染影片。

## 必须先理解的范围

本配方与源最后演示相同：开合期间旋转的是屏幕内静帧poster；完全打开的目标页再用一个独立、平面的原生视频层淡入播放。它不是“实时视频贴在旋转3D表面上”。只有打开页屏幕RotationY=0、正对观众时，本例矩形视频才与屏内画面对齐。改变打开角度、屏幕弯曲、任意透视或要求持续贴面视频，须另做原生能力验证，不能直接把二维视频框当作随屏幕透视的纹理，也不因此默认宣称整场景必须external-render。

内屏与原创16:9视频尺寸预先一致，本例无需源中对任意比例素材做Intersect裁切；没有声称“圆角视频Intersect保留媒体”已经本例专项测试。闭合/打开两页poster来自同一视频首帧，修正源示范中上一页仍为旧绿照片的已承认不一致。

## 轴与真实对象

每个body外框18.2×10.6cm。屏幕可见面在铰链上方，等尺寸无填充/无线条helper放下方；键盘可见面在铰链下方，等尺寸helper放上方。各自成组后bbox高度21.2cm，两个组中心都在y11.9cm铰链。helper要匹配外框，不能错误拿更小的内屏视频矩形作轴补偿。

相机preset为本机PIA确认的PerspectiveRelaxedModerately=62，FieldOfView45°，X/Z归零。Closed屏幕Y90、键盘Y270；Open屏幕Y0、键盘Y290。COM可规范读回270→-90、290→-70，按同角度理解；本机实际生成camera.latitude=角度×60000。不能凭字段叫Y就猜笛卡尔旋转方向，已看原生实图确认横向铰链开合。其余预设须重测。

机身壳与屏框深度3pt，底部bevel5×5pt，原创使用circle预设；两页Morph2秒。视频为真正Type16/MediaTypeMovie3内嵌H264，打开页MainSequence为Fade1秒WithPrevious和MediaPlay WithPrevious；Loop/PlayOnEntry/Rewind开启，StopAfterSlides1。默认原生视频点击暂停的交互节点不等于多了主序列点击步骤。

## 包内可复现流程

每个输出新建且拒绝覆盖。builder/资产helper只用Python标准库调用OfficeCLI/FFmpeg；像素验证需Pillow、NumPy。使用Windows本机PowerPoint进行原生精修和真实播放。

1. `laptop_hinge_assets.py --ffmpeg <exe> --outdir <fresh-assets>`：生成original-screen.mp4与匹配first-frame.png。
2. `laptop_hinge_build.py --officecli <exe> --poster <first-frame.png> --video <original-screen.mp4> --outdir <fresh-base>`：制作完整两页含真媒体；视频参数必需，避免把无视频轴探针误作完整作品。
3. `laptop_hinge_polish.ps1 -InputDeck <base/laptop-base.pptx> -OutputDeck <fresh-final.pptx> -EvidenceDirectory <fresh-native>`：原生相机/动画，另存后重开读回与导图。
4. `laptop_hinge_record.ps1 -Deck <final> -EvidenceDirectory <fresh-live> -Ffmpeg <exe>`：22秒真实桌面录屏；Right/Home/Right物理按键开、合、再开，8状态图及只读Player样本，不用程序seek/start假造播放进度。只有自己的放映窗口获得前台才发送键，结束只关自己的文稿。
5. `laptop_hinge_verify.py --deck <final> --native-dir <native> --assets-dir <assets> --live-dir <live> --ffprobe <exe> --output <fresh-report>`。可用`--structure-only`先静态审核；可加`--input-plan <json>`检查测试前记录的inputSha256是否仍同。

## 研究闭环与独立证据

full-v1/verify-v2.json：59项通过。22秒真放映影片20帧已复看：屏幕与透视键盘开合，中间poster随屏转；打开后出现1秒淡入混合帧并接活视频；闭合归位后再开正常。

不能只拿Player进度代证画面：实际打开页视频两个时刻屏内平均差9.03RGB级；屏幕矩形外机身/边框像素最大差0；闭合前与播放后闭合整帧逐像素一致。45条有效数值Player样本，5703→117ms及5996→289ms真实循环。过渡时某些Player位置为null，未当作播放证明。另存重开保持真实媒体与2个自动主序列效果。

包内fresh前向 validation/forward-laptop-hinge-scene 已独立完成：48项静态结构/嵌入/轴/相机/视频层核验；全新22秒真实播放与60项最终检查全部通过，含输入PPT前后SHA一致。8状态全部片内，20帧fresh联系图已复看。fresh有效Player数值样本44条，屏内帧差9.47RGB级，稳定打开页屏外机身/边框差0，闭合前后场景差0；OfficeCLI validate通过。修改尺寸、相机、开合角度或视频比例后必须重新导图/实播，稳定打开时屏外0的量测不自动推广到所有过渡与任意角度。

## 来源范围

大师课9Ewb0Rd3P3E，3:15:53–3:36:50转写全文与章节6分布帧已读；Ego自有space29补8精确帧后finish。12035s等尺寸复制体，12057s沿边翻转组；12185s仍可见下方补偿屏，不能冒称已隐藏；12276s键盘274.8°是调270°的中间态；12695s屏幕X/Y/Z0与Perspective45°清楚；12921s素材裁切，12939s仍是旧绿照片；12976s Fade WithPrevious1.00清楚。无填充helper、270°/290°、视频自动播放等结合完整字幕定位，不把中间UI帧冒充所有最终参数。
