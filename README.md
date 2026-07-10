# Talent Pool 人力画像自动化分析

这是一个静态网页工具，用于上传 `survey pool.xlsx` 并自动生成可复制、可下载的人力画像分析结果。

## 在线共享给同事

这个项目可以直接部署到 GitHub Pages。同事不需要安装任何环境，只需要打开 Pages 链接并上传 Excel。

部署步骤：

1. 把本仓库推送到 GitHub 的 `main` 分支。
2. 进入 GitHub 仓库页面，打开 `Settings` -> `Pages`。
3. 在 `Build and deployment` 里选择 `Source: GitHub Actions`。
4. 回到 `Actions`，等待 `Deploy GitHub Pages` 工作流完成。
5. 部署完成后，GitHub 会给出一个 Pages 网址，例如 `https://你的用户名.github.io/仓库名/`。

隐私说明：Excel 文件在同事自己的浏览器里解析，当前版本不会把上传的 Excel 发到 GitHub 或任何后端服务器。

## 使用方式

1. 打开 `index.html`。
2. 上传表头不变的 Excel 文件。
3. 页面会按 `Contact Email` 去重统计人数，并生成 KPI、分类图表、汇总表。
4. 可下载图表 PNG、汇总 CSV、Markdown 报告。

## 当前支持的分析维度

- 去重人数、项目经验覆盖、人均项目记录、国家覆盖、高英语能力
- 工作时长分布
- 项目数量分布
- 提交时间趋势
- 项目类型
- 英语水平
- 学历
- 专业/领域
- 语种-国家
- 国家/地区
- PM、国家、学历、英语水平、关键词筛选

## 数据口径

- 统计人数按 `Contact Email` 去重。
- 如果 `Contact Email` 为空，页面会把该行当作单独记录保留，避免误合并。
- 项目经验会合并 `Project No. 1-3` 下的 `Project Types`、`Duration`、`Role` 等字段。
- 国家/地区统计优先使用 P 列 `Resident Country/Region`。
- 专业统计优先使用 V 列 `Major/Specialty (学科专业)`。
- Business Line 统计优先使用 AN、AZ、BL 三列，对应三个项目块的 `Business Line`。
- Major 和 Business Line 按每次有效出现计数；同一人员可在同一统计中贡献多次。
- Major 图表和汇总仅显示 V 列路径的第一级英文学科名称，并移除四位学科分类代码。

## 依赖

页面通过 CDN 加载 SheetJS 和 Chart.js。如果需要完全离线版本，可以把这两个库下载到本仓库并修改 `index.html` 的 script 地址。
