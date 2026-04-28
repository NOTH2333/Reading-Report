# 《交易心理分析》学习包

## 目录

- 总阅读报告：`reports/阅读报告.md`
- 章节索引：`summaries/00-章节索引.md`
- 章节 PPT：
- 第 1 章：`ppt/01-成功之路：基本面、技术面或思想分析？.pptx` 与 `ppt/01-成功之路：基本面、技术面或思想分析？.js`
- 第 2 章：`ppt/02-交易的诱惑（和危险）.pptx` 与 `ppt/02-交易的诱惑（和危险）.js`
- 第 3 章：`ppt/03-自己承担责任.pptx` 与 `ppt/03-自己承担责任.js`
- 第 4 章：`ppt/04-持续一致性：一种思想状态.pptx` 与 `ppt/04-持续一致性：一种思想状态.js`
- 第 5 章：`ppt/05-认知的动力.pptx` 与 `ppt/05-认知的动力.js`
- 第 6 章：`ppt/06-市场的角度.pptx` 与 `ppt/06-市场的角度.js`
- 第 7 章：`ppt/07-交易者的优势：考虑概率.pptx` 与 `ppt/07-交易者的优势：考虑概率.js`
- 第 8 章：`ppt/08-和信念一起工作.pptx` 与 `ppt/08-和信念一起工作.js`
- 第 9 章：`ppt/09-信念的天性.pptx` 与 `ppt/09-信念的天性.js`
- 第 10 章：`ppt/10-信念对交易的影响.pptx` 与 `ppt/10-信念对交易的影响.js`
- 第 11 章：`ppt/11-像交易者一样思考.pptx` 与 `ppt/11-像交易者一样思考.js`
- 章节知识总结：
- `summaries/01-成功之路：基本面、技术面或思想分析？.md`
- `summaries/02-交易的诱惑（和危险）.md`
- `summaries/03-自己承担责任.md`
- `summaries/04-持续一致性：一种思想状态.md`
- `summaries/05-认知的动力.md`
- `summaries/06-市场的角度.md`
- `summaries/07-交易者的优势：考虑概率.md`
- `summaries/08-和信念一起工作.md`
- `summaries/09-信念的天性.md`
- `summaries/10-信念对交易的影响.md`
- `summaries/11-像交易者一样思考.md`
- 配图资产：`assets/ch01` 至 `assets/ch11`，以及 `assets/common`

## 设计说明

- 语言：中文为主，关键英文术语只在需要时补充。
- 风格：图解启蒙风，适合零基础读者阅读，也适合二次讲解。
- 配图：全部为本地原创 SVG 图解与场景图，不依赖外部版权图片。

## 生成链

- 文本抽取：`python scripts/extract_book.py`
- 统一生成：`node scripts/generate_learning_pack.js`
- 抽样校验：`powershell -ExecutionPolicy Bypass -File scripts/validate_decks.ps1`
