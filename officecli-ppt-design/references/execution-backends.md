# 四类实现后端

这是用户指定的工程分类，和“字体/版式/Morph/交互”等设计类别正交。每个原子配方标注实际采用的后端，不按教程标题或视觉风格预先决定。

|后端|适用范围|交付与验证|
|---|---|---|
|`officecli-native`|OfficeCLI可完成的页面、排版、素材、形状、基础动画、Morph及经验证的局部OOXML实现|CLI基础脚本、可编辑PPT、原生PowerPoint图像/动态测试|
|`powerpoint-com`|在既有文稿/对象上，需要精细控制的原生PowerPoint动画与行为|可重放COM/VBA精修脚本、参数计划、精修后PPT、原生行为与节奏验证|
|`hybrid`|OfficeCLI建立基础页面，再由PowerPoint COM完成几何或动效；确无受测API的Summary容器明确采用原生UI创建|基础PPT+基础脚本、精修/原生UI作者脚本+最终PPT，各层贡献和UI语言/依赖分开记录|
|`external-render`|核心视觉采用尚无已验证原生配方的外部像素计算或渲染，如限定通道分离、alpha提取或外部动画|原件与参数、外部工程/生成脚本、PNG或视频、嵌入PPT及适用的原生视觉/播放验收；明确丢失的可编辑性，不将接口未实现推断为原生绝不可能|

## 路由规则

先判断目标是静态设计、原生对象状态变化、原生时间轴/事件，还是外部渲染视觉。优先选择能保留编辑性且实际满足目标的最简单后端。

- CLI高层属性失效，不等于PowerPoint做不到；先检查COM/VBA和实际原生结构。
- 某个COM接口不暴露，也不等于视觉本身必须AE；可能是GUI原生能力或需要最小XML补全。
- 普通图片处理、模糊素材准备不自动将整份PPT归为external-render；分类依据核心效果在哪里产生。
- 已有动画GLB模板通过CLI改文案，可标officecli-native的“模板保留工作流”，不能扩大为“从零创建模型动画”。
- `external-render`没有经过作品验证时，仅保留路由和待办，不编造已完成的AE能力、不擅自安装外部软件。

当前已测示例：简化卡片/花瓣/固定窗口/现代Morph为officecli-native；通用Timing计划为powerpoint-com；真实卡片数字Union、Fill.Background、三轮延迟拖尾和精修液面为hybrid。完整分项见[配方索引](recipe-index.md)。外部渲染保留明确边界，未用近似原生效果填充通过名单。
