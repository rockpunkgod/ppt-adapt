# Kiosk 答题门控与改题依赖迁移

backend：hybrid。OfficeCLI负责两页原创题目、可编辑答案形状、Next形状与指向结果页的内部动作；PowerPoint COM负责原生指定对象触发动画、Kiosk放映设置和已有Next动画迁组。

## 适用行为

三个答案是同页指定触发器。错误答案变红；正确答案变绿并在同组动画后显示Next，只有点击出现的Next才跳结果。默认题2+3=5，可用配套改题流程改为2+4=6。Kiosk阻止本配方的空白点击、方向键、PageDown、Space绕过；Esc正常退出。不是安全或权限隔离，未声称阻止Tab、开发接口等其他路径。

## 可复现步骤

以下均为包内 scripts 文件名，输入/输出用绝对路径。每个输出必须新建，不覆盖原稿。基础构建只需标准库Python和OfficeCLI；像素验证额外需要Pillow；原生精修/实播需要Windows PowerPoint与FFmpeg。

1. `kiosk_quiz_build.py --officecli <exe> --outdir <fresh-dir>`：生成quiz-base.pptx。Next保留形状动作，删除OfficeCLI附加的文本Run动作，防止白字被超链接蓝色下划线覆盖。
2. `kiosk_quiz_polish.ps1 -InputDeck <base> -OutputDeck <fresh-native.pptx> -EvidenceDirectory <fresh-polish> -CorrectAnswer 5`：三个原生InteractiveSequence，答案自身触发FillColor+FontColor，0.25秒；正确组Next Fly From Bottom，After Previous，0.5秒。保存ShowType3Kiosk，AdvanceMode手动；页点击保留启用，以供Speaker负对照。
3. `kiosk_quiz_test.ps1 -Deck <native> -EvidenceDirectory <fresh-live> -CorrectAnswer 5 -Ffmpeg <exe> -SpeakerControl`：仅对自己的可见放映窗口物理点击与按键，Arrow指针，前台/命中/实际坐标护栏。自动保存原生桌面录屏、状态PNG与日志；Speaker只改内存，不保存覆盖Kiosk文件。
4. `kiosk_quiz_verify.py --deck <native> --evidence-directory <live> --correct-answer 5 --output <fresh-json>`：结构、动画实际依赖、页面动作、真实输入结果、反馈白字像素共同检查。不得只凭XML关系存在宣称通过。

## 改题不是只换文字

1. `kiosk_quiz_prepare_variant.py --officecli <exe> --input <native-A> --output <fresh-variant.pptx> --addend 4`。仅改变真实算式和结果，刻意保留旧正确组，便于检查迁移。
2. `kiosk_quiz_rebind.ps1 -InputDeck <variant> -OutputDeck <fresh-B.pptx> -EvidenceDirectory <fresh-migration> -CorrectAnswer 6`。旧5绿色改红、新6红色改绿，把同一个Next Effect移到6的FontColor后，保留真实页面动作。
3. 对B运行同一个test，CorrectAnswer6。verify添加 `--migration <migration-dir/migration.json>`。旧5必须显示红色且Next仍隐藏，新6绿并显Next，物理点击才到正确结果页。

COM的旧Effect对象Parent在MoveAfter后可仍回报旧5组；不能据此判断实际归属。脚本重新枚举三个InteractiveSequences查唯一Next成员，再交叉保存XML和实际点击验证。

## 证据与范围

源：Masterclass9Ewb0Rd3P3E，7:20:03–7:26:55为Kiosk/指定触发/正确Next，7:27:02–7:28:07为正确答案变更后的真实动画迁移。精读扩展字幕7:19:30–7:29:30；另经Ego补8精确源帧（Kiosk选项、Object/Font Color、0.25/0.50时长与触发面板）；26878s恰是Next拖移前旧B组，不把该帧当迁移后证明。实际源画面分布样本及补帧覆盖范围见distill/kiosk-quiz/source-notes.md。原创算式与三个大按钮还原机制，不宣称复制源的四菱形视觉。

研究A首次门控通过；研究B 2+4迁组后68项通过，有8秒实际桌面录屏与10状态PNG。改题后三答案反馈白字像素各自完全保留，Next白字2374像素。包内exact全新前向 validation/forward-kiosk-quiz 已完成：A 67项、B 68项全部通过；两份OfficeCLI validate无误，两段各8秒真实桌面录屏和每题10状态PNG。B影片8帧联系图已逐帧复看，包含初态、两个错误、正确显Next和2+4=6结果。代码不复用研究PPT，A从包内builder新建，B由A通过包内CLI改题与原生迁组生成。
