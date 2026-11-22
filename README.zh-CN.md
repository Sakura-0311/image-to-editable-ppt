# image-to-editable-ppt

[English](README.md) | **简体中文**

把图片 / PDF / 扁平幻灯片重建成**可编辑的现代 `.pptx`**，同时给已有演示文稿做「去 AI 味」改造的离线技能包。

整合三部分规范：

- **de-ai-pptx v3.1**：合规闸门 → 取证打分（五层指纹）→ `grill-me` 拷问对齐 → 五层改造 → 验收复扫；另含从零生成模式、三套离线视觉配方与提示词模板库
- **frontend-slides 动效规范**：动效询问模板、情绪-动效匹配表、单文件 HTML 网页幻灯片选项
- **视觉逆向重建**：页面归一化 → 分层解析 → 资产剥离清理 → `python-pptx` 原生重建 → 渲染质检

## 触发词

> **为我使用ppt-xy**

说出这句即自动触发本技能。触发后首轮先执行 `grill-me` 提问对齐（包含视觉风格、是否需要动态效果及页码范围），对齐后才动手；交付强制为最新现代 `.pptx`（OOXML），不降级为老旧 `.ppt`，不把整页压成图片。

## 文件结构

```
image-to-editable-ppt/
├── SKILL.md                                  # 主入口：输出格式、逐页定制设计、字体配对、硬禁令、任务路由
└── references/
    ├── de-ai-pptx-v3.1.md                    # 去 AI 味完整规范（15 章）
    ├── frontend-slides-motion.md             # 动效与 HTML 网页幻灯片扩展
    └── visual-reconstruction.md              # 视觉逆向重建（Route 4）
```

## 四条路线

1. **改造**：已有 `.pptx` 去 AI 味，五阶段流水线走完交付改造版
2. **体检**：只取证打分，输出双评分报告，不改文件
3. **生成**：从主题/文档从零做一份不像 AI 的 PPT
4. **视觉重建**：截图、图片、PDF、图片版 PPT 转可编辑 `.pptx`

## 使用

把本仓库放进你的 AI 助手技能目录（如 `skills/image-to-editable-ppt/`），或把 `SKILL.md` 与 `references/` 一起交给助手读取，然后说触发词「为我使用ppt-xy」即可。

依赖：`python-pptx`（必需）、`Pillow`（可选），以及系统工具 `soffice`、`pdftoppm` 用于渲染质检。全程离线，无需联网下载模板或脚本。

## 鸣谢

特别鸣谢 **HXY** 对本技能的贡献与支持。

同时感谢开源项目 [zarazhangrui/frontend-slides](https://github.com/zarazhangrui/frontend-slides)（动效规范来源）与 `grill-me` 拷问协议。
