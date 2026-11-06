# 夜读试点：整套组合验收

原创8页管理层决策演示。全部数字为模拟规划数据，非真实调查。打开 `library-night-pilot.pptx` 编辑文字、原生表格、原生图表和Morph对象；图表采用原生数值，不承诺PowerPoint的Excel数据工作簿编辑功能。

`brief.md`、`director-rationale.md`、`design-spec.json`、`deck-plan.json`分别记录任务、设计理由、全局规范和逐页规划。`composition-release`记录制作前的配方选择、依赖图、空间/时序约束及冲突检查。最终验证状态和证据见原生验收报告；单项配方的既有通过不代替本次组合结果。

`native-verification`保留本机PowerPoint的8页原生图、完整实屏录制、实际方向键事件和对象读回。`theme-regression`是单独的主题强调色变更对照：全局色槽联动，夜间章节背景保持。对照稿是验收证据，不是第二套设计交付。

重新制作时，先在仓库根目录执行：

```powershell
python ppt-director/scripts/validate_deck_plan.py --spec examples/library-night-pilot/design-spec.json --plan examples/library-night-pilot/deck-plan.json
python ppt-composition-planner/scripts/plan_composition.py --spec examples/library-night-pilot/design-spec.json --deck-plan examples/library-night-pilot/deck-plan.json --index officecli-ppt-design/references/recipe-index.md --out NEW-COMPOSITION
python examples/library-night-pilot/build_demo.py --officecli PATH-TO-OFFICECLI --out NEW-BASE.pptx --composition NEW-COMPOSITION/composition-plan.json
powershell -NoProfile -File officecli-ppt-design/scripts/polish_timing.ps1 -InputPath NEW-BASE.pptx -OutputPath NEW-FINAL.pptx -PlanPath NEW-BASE.timing-plan.json
```

基础页和动效来自现有配方原语，示例适配器不构成另一套PPT引擎。COM脚本原样调用；本例向它提供一个JSON对象精修首个图表效果，四个效果的500ms时长均需在原生验收中读回。

原生录制用 `validate_native.ps1 -Deck ... -Plan ... -Spec ... -Evidence NEW-DIR -Ffmpeg ...`，需要本机PowerPoint、FFmpeg和可独占的放映界面；脚本只关闭自己的稿件。独立验收用 `verify_integration.py --deck ... --plan ... --spec ... --data ... --composition ... --evidence ... --regression-deck ... --regression-evidence ... --ffmpeg ... --out NEW-REPORT.json`。Pillow为图像核查依赖。只有实际检查过原生画面和记录的审核者才填写 `--visual-reviewer`。

验收绑定本次文件SHA及原始绝对路径，是这次作品的记录。换机器、改输入、换素材或重新生成文稿后，应生成新组合计划并重新原生验收；不能将归档记录当作另一份稿件的通过证明。
