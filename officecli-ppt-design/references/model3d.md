# 原生3D：模型内部动作与跨页场景

来源：[One Skill PPT 3D Morph](https://www.youtube.com/watch?v=3Qo_Hx21KUc)，重点01:51–04:01、05:12–11:10。

## 分清两套动画

Floating、running、waving是GLB中的内部动画；位置、大小、三维朝向的跨页变化是Morph。二维PNG缩放、shape.spin或一个没有动画轨道的GLB不能等价替代。

本轮通过的是**OfficeCLI改造作者公开原生模板，保留3D动画，再用本机PowerPoint确认**，不是从零调用CLI生成库存模型。

## 复现方法

1. 由调用方提供合法取得的含动画GLB的PPTX模板。技能不捆绑完整模板和模型。
2. 使用`build_morph.py --mode 3d-template --source-3d <file> --outdir <new-folder>`改写可编辑文案、备注与字体。
3. 保留model3d、关系、embedAnim、animPr、posterFrame与spTgt；比较模型字节哈希和各页动画结构，不只验证文件中“有.glb”。
4. 同一主体相邻页保留唯一`!!name`，不同内部动作场景可以是不同主体。不要把所有宇航员都改成同一个名字。
5. 屏外模型不做无意义循环；固定页码/导航不随长场景一起滚动。来源在四屏纵向场景中用2秒Morph。
6. 用本机PowerPoint播放，检查内部动作、三维姿态和跨页位移同时成立，并留意卡顿。

## 本轮对照修正

源字体fallback造成文案挤压，改为实际可用Arial并清除旧kern后重新原生导图和录屏。不能因为保留原模板就假设排版不会坏。

4页模型与原生动画节点和源保持一致，原生视频确认不是静态替代图。不同素材、中文文案仍需重新排版。导航点击是另一项测试，模型/转场通过不表示菜单也通过。

## 扩展条件

用户希望从零插入新模型或新增专属3D动画时，检查本机PowerPoint COM/VBA是否暴露对应能力，保留精修脚本与回读证据。若只能通过GUI库存选择或缺少动画资产，报告具体缺口；不编造`scene=`、`animationScene=`等CLI参数。
