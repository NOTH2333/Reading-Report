# 股票大作手回忆录 / GU PIAO DA ZUO SHOU

本项目已从“外层仓库 + 内层实际项目”的双层目录结构收平成单项目根，当前目录就是实际工作目录。

## 主要入口

- `output/`：最终交付物，包括 PPT、Markdown 和共用图片素材。
- `source/`：源 PDF。
- `data/`：中间文本与结构化数据。
- `scripts/`：生成与检查脚本。
- `README_目录说明.md`：更完整的目录约定、重建顺序和维护原则。

## 当前保留策略

- `node_modules/` 仍保留在项目根，用于保证现有 Node 构建链路不被打断。
- `tmp/` 仍保留临时渲染结果和日志，不与最终交付物混放。

## 常用位置

- 阅读报告：`output/md/阅读报告.md`
- 关键知识总结：`output/md/关键知识总结.md`
- PPT：`output/ppt/`

## 进入项目

```powershell
cd "01_股票大作手回忆录__GU_PIAO_DA_ZUO_SHOU"
```
