# 悬停菜单、孔洞遮罩与真正 Slide Zoom

backend：hybrid。OfficeCLI建原创5页基础、真正Slide Zoom、双孔custom geometry及海报关联；PowerPoint COM补MouseOver动作与Morph时间、导原生目标海报。保留既有hover-state-navigation.md“移出保持”的独立配方；本配方补的是移出回home＋点击深入及Zoom原生返回。

## 结构与关键行为

home1、hoverA2、hoverB3、detailA4、detailB5。卡片移入指定hover页；active卡不链接自身，避免不断重启。图标缩小上移，Zoom/label从遮罩外上移进孔。标题等在遮罩前，Zoom和label在遮罩后。

遮罩是一块真实自定义几何：整页外轮廓＋两个反向闭合子路径，孔处透明允许卡片与Zoom接受输入；不是alpha0矩形。hover页遮罩原生MouseOver链接home1。卡片自身、其上图标/文字的MouseOver一致，避免可见内容造成输入死区。

Zoom必须是Microsoft slidezoom namespace的Choice/graphicFrame/sldZmObj，目标为真实内容页，returnToParent=true、transitionDur500ms；不是普通页面链接，也不换成Section Zoom。PowerPoint正常另存后可省略returnToParent：官方schema默认true，验证器接受这一合法默认，同时仍要求实际返回事件。旧阅读器Fallback只是普通海报/链接，不能承诺同样的缩放返回。

## 可复现流程

所有输出新建且拒绝覆盖。基础builder/poster修改器只用Python标准库；独立像素核验需要Pillow。Windows PowerPoint精修/实播与FFmpeg/ffprobe必需。

1. `hover_zoom_build.py --officecli <exe> --outdir <fresh-base-dir>`。
2. `hover_zoom_polish.ps1 -InputDeck <base.pptx> -OutputDeck <fresh-native.pptx> -EvidenceDirectory <fresh-native-dir>`。
3. `hover_zoom_posters.py --officecli <exe> --input <native.pptx> --output <fresh-final.pptx> --native-dir <native-dir>`。关联真实原生detail4/5 PNG，隐藏海报缓存；保持真正Zoom Choice。
4. `hover_zoom_test.ps1 -Deck <final> -Directory <fresh-live> -Ffmpeg <exe>`。默认120秒真实桌面录屏；两分支各两轮、21个状态。每轮移入→孔内保持→物理点击Zoom深入→物理点击空白Zoom返回→穿出背景回home。无GotoSlide/Next驱动。全过程只操作自己的可见放映窗口，前台/WindowFromPoint/逐步cursor守卫，Arrow指针，结束只关自己文稿。
5. `hover_zoom_verify.py --deck <final> --live-dir <live> --ffprobe <exe> --output <fresh-json>`。可加`--input-plan <json>`，其中inputSha256为测试前最终PPT哈希，以验证测试前后原稿不变。

## 输入与坐标校准

真实MOVE不是定位后推断成功。初始只在中性点2,6cm用一次SetCursorPos；受测动作均用连续小步相对mouse MOVE，并逐步回读实际cursor。每段8秒上限、测试120秒输入截止；终点允许1物理px量化误差。按真实client/DWM矩形与纵横比推导slide的letterbox，不能将PowerPoint COM窗口宽高数字直接当物理像素。

研究校准真实client2560×1600，slide实显2560×1439.858，顶部/底部约80.071px；COM窗口读值1536×960不是同一数值。原32,18cm点实际在slide内，普通实填矩形正控制两次Hover返回都通过，且鼠标处RGB与框色完全相同。早期校准汇总存在PowerShell逗号算术错误，保留原失败记录，独立从原始输入事件重建通过；不是把汇总失败掩成无错误运行。

## 已测证据与因果边界

研究prototype/live-continuous：129项全过，21状态、2904逐步cursor样本、26段MOVE，最后状态65.74秒在完整120秒影片内。输入PPT测试前后SHA256相同。逐状态卡片与遮罩像素核验；18个状态帧与12个中间帧已复看，包含图标缩小、海报进孔、真实Zoom放大与返回、四次背景复位。每状态六次稳定抽样，通过所列路线；不宣称任意极速横扫、全部显示器/输入设备、网页版均无抖动。

早期直接SetCursorPos＋同点微动在外侧背景与重复进入上存在阴性。延长等待、无Morph、仅改为原生Subtract生成的同一mask几何均未修复；彩色mask静态导图显示真实覆盖，但不代证show命中。最终同一原稿采用连续真实MOVE路线实播通过。保留全部对照，不能把它写成已唯一确认的PPT内部根因，也不能说原稿几何曾被修好。

包内fresh exact validation/forward-hover-zoom-menu已完成：42项结构与海报字节检查、5页原生布局逐像素一致；新PPT独立120秒实播与135项最终核验全部通过，21状态最后65.965秒均在影片内，2904轨迹样本。原稿测试前后SHA256一致，16帧新影片联系图已复看，含Morph/Zoom中间态与每轮背景复位。最终PPT OfficeCLI validate通过。

## 来源与复刻范围

Masterclass9Ewb0Rd3P3E，字幕2:30:06–2:38:39：2:31:00 Morph0.35秒；2:31:31 MouseOver卡→hover；2:32:10背景MouseOver回home；2:35:21 Return to Zoom及随后0.5秒；2:36:37–2:38:16照片Subtract孔洞、遮罩层级与返回动作。源四卡照片玻璃视觉转为原创两卡几何机制还原，非逐像素复制。

六精确源帧见distill/hover-zoom-menu/source-detail.jpg。9095/9133s仍是Action on click初始页，不能当MouseOver已选；9330s Return已勾但Duration仍1.00，属于调整前；9445s清楚展示mask置顶只露孔内卡，9494s标题/标识移到mask前。完整字幕与帧样本联合定位具体操作，边界保留。
