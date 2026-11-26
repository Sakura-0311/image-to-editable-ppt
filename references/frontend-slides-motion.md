# 动效、原生 `.pptx` Morph XML 与图片取景规范
 
本规范整合自 `zarazhangrui/frontend-slides` 与 `HY-hue113/ppt-effects`（离线 `zipfile` + `lxml` 实现，不依赖外部仓库脚本），覆盖四项能力：
1. `grill-me` 动态效果必问模板与交付后提醒；
2. 原生 `.pptx` **Morph 平滑切换**与**页内自动衔接动画（`<p:cTn>`）** XML 注入；
3. **原生形状填充取景（`blipFill` + `custGeom` 三项校验）**与固定 Z 序；
4. **素材明暗度量化门禁**及单文件 HTML 网页幻灯片选项。
 
## 1. `grill-me` 必问：动态效果决策题模板
 
在每一轮 `grill-me` 提问中，必须包含如下关于动态效果的决策题（若尚未明确）：
 
```markdown
Q - 动态效果与范围：这份演示文稿是否需要加入动态效果？如果需要，更倾向原生 `.pptx` 的 **Morph 平滑切换** / **页内自动衔接动画**，还是独立网页版动效？范围是只加在关键页（封面、章节过渡页、核心数据/大图页），还是全篇每页都加？
- 推荐答案：仅在封面、章节页与核心数据/大图页加入原生 `.pptx` Morph 平滑切换（旧版自动降级淡入淡出）或克制的 200-300ms 自动衔接入场；信息密集页保持静态，避免干扰阅读。
- 理由：只在重点页加动效能突出叙事节奏且不影响信息密度；全篇都加容易拖慢汇报节奏或造成阅读疲劳。
```
 
### 交付后的固定提示（当用户选择“暂不需要动态效果”时）
 
如果用户在 `grill-me` 中选择**当前不需要动态效果**，在最终交付 `.pptx` 文件时，必须在回复中明确附上以下说明：
 
> 当前版本未添加动态效果；如果你后续希望为其中某几页（例如封面、章节过渡页或重点数据页）或全部页面增加动态效果，直接告诉我具体页码和想要的节奏风格，我可以随时为你补充原生 `.pptx` Morph 平滑切换 / 页内自动衔接动画（也支持导出带完整 CSS 动效的独立 HTML 网页版）。
 
## 2. 原生 `.pptx` XML 动效注入规范（纯离线 `zipfile` + `lxml`）
 
### A. Morph 平滑切换（PowerPoint 2019+ 原生生效，旧版自动降级淡入淡出）
 
通过 `zipfile` + `lxml.etree` 修改目标页 `ppt/slides/slideX.xml`，在 `</p:cSld>` 之后（`</p:sld>` 之前）注入标准 `<mc:AlternateContent>` 兼容性节点：
- 声明命名空间：`xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"`、`xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"`、`xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main"`。
- `<mc:Choice Requires="p159">` 内写入：
  `<p:transition spd="slow" p14:dur="1200"><p159:morph option="byObject"/></p:transition>`（商务节奏可调 `dur="800"`–`"1200"`，电影感可调 `"1800"`）。
- `<mc:Fallback>` 内写入降级切换：
  `<p:transition spd="slow" p14:dur="800"><p:fade/></p:transition>`。
- **Morph 生效前提**：前后两页需要平滑过渡的同一元素必须保持同类型对象，并建议以 `!!` 开头命名形状（如 `!!hero_card`）强制按对象匹配。
 
### B. 页内时间轴自动衔接（`<p:cTn>` 补丁）
 
当同一页存在多个连贯动画元素、希望避免演说者反复点击鼠标时：
- 在 `slideX.xml` 的 `<p:timing>` 节点中定位后续动画的 `<p:cTn>`，将 `nodeType="clickEffect"` 改为 `nodeType="afterEffect"`（或并行 `withEffect`），并将 `<p:stCondLst><p:cond delay="0"/></p:stCondLst>` 的延迟设为 `0`（或错落 `150`–`250`ms）。
- **落地铁律**：只在字少图大的重点页加动效，去掉动画后页面核心信息不得受损。
 
## 3. 原生形状填充取景（`blipFill`）、`custGeom` 校验与素材明暗门禁
 
1. **图片作为形状填充**：优先用原生形状（`p:sp` + `prstGeom`/`custGeom` + `blipFill`）承接图片，而非 Pillow 预裁静态圆角 PNG，保留 PowerPoint 内二次调整裁切能力。
   - 宽画幅仅用不变形形状（如 `capsule`、`cut`）；圆形用 `ellipse` 不拉伸；同款异形框单页或连续页最多出现 `<= 2` 次。
2. **`custGeom` 自定义轮廓三项硬校验**（缺一不可）：
   - `<a:path>` 必须显式带 `fill="norm"`；
   - `<p:spPr>` 子元素严格按 `xfrm -> geom -> 填充(blipFill/solidFill) -> ln` 顺序排列；
   - 路径必须闭合（`<a:close/>`）且坐标严格落在 `[0, 100000]` 范围内。
3. **固定 Z 序**：卡片内填图必须加局部蒙版以保证文字可读，Z 序固定为 `底层卡片形状 -> 中层局部蒙版 -> 上层文字 -> 旁侧独立取景图`。
4. **素材明暗量化门禁**（执行前用 Pillow 实测图片平均亮度）：
   - 浅色底 / 黑字页压图要求亮图（平均亮度 `< 140` 触发预警，需换图或调浅遮罩）；
   - 深色底 / 白字页要求暗图或加深 scrim 蒙版；
   - 透明/白底抠图槽仅接受干净白底或透明 PNG，杂乱背景直接拦截或先做抠图清理。
 
## 4. 情绪-动效匹配表（Effect-to-Feeling Guide）
 
| 氛围目标 | 推荐动效节奏 | 适用视觉线索 |
| :--- | :--- | :--- |
| **商务 / 专业（Professional）** | 克制快速（200-300ms）、干脆平滑入场 / 800ms Morph | 深蓝/石板灰、精确间距、数据图表主导 |
| **杂志 / 编辑感（Editorial）** | 错落文字揭示（staggered `150-250ms`）、图文交错浮现 | 强字号对比、引文块、衬线标题 + 无衬线正文 |
| **电影感 / 强冲击（Dramatic）** | 慢速淡入（1-1.5s）、1800ms Morph 缩放过渡 | 深色底、聚光氛围、全幅大图 |
| **极简 / 克制（Calm）** | 极缓微弱渐显、轻柔淡出 | 高留白、低饱和配色、大内边距 |
| **科技 / 未来感（Techy）** | 网格揭示、微光高亮、数据节点依次显现 | 网格底纹、等宽数字、高对比强调色 |
 
## 5. 单文件 HTML 网页幻灯片选项（仅当用户明确要求网页版/高交互动效时）
 
如果用户要求生成可在浏览器直接演示的高动效 HTML 网页版（而非 `.pptx`）：
1. **零依赖单文件**：全部 CSS/JS 内联，无需 `npm` 或构建工具。
2. **固定 16:9 舞台（`1920x1080`）**：整体等比缩放适配视口，禁止在手机端重排内部版式；通过 `.active` / `.visible` 的 `visibility`、`opacity` 和 `pointer-events` 控制切页（禁止用 `display: none/block` 覆盖 flex 布局）。
3. **必备细节**：支持 `prefers-reduced-motion` 无障碍降级、键盘左右键/空格翻页、以及浏览器内联文本编辑快捷键（如 `E` 键）。
