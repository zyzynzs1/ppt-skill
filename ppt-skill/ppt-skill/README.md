# ppt-skill

专业级中文商务/技术汇报 PPT 生成技能（Agent Skill）。

## 内容

- `ppt-designer/SKILL.md` — 技能主文件：工作流程（需求访谈 → 大纲策划 → 生成交付）、设计系统、组件规范、写作规范、质量检查清单
- `ppt-designer/templates/design-tokens.md` — 设计令牌：画布/栅格/配色/字体/组件全部参数
- `ppt-designer/templates/page-archetypes.md` — 页面原型：封面、目录、内容页（5 种变体）、章节过渡、数据页、结尾页
- `ppt-designer/scripts/pptx_helpers.py` — python-pptx 辅助函数，一键应用全部设计令牌

## 风格定位

结构化信息图风格：高密度卡片化布局、蓝金双色系统、微软雅黑全稿统一、
结论式标题、整页淡蓝底 + 统一页脚，适用于方案汇报、产品介绍、投标述标、培训课件、咨询报告。

## 使用

将 `ppt-designer/` 目录复制到你的 Agent skills 目录即可。
