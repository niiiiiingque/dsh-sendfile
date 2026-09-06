# 兼容与维护记录

版本：dsh-sendfile 0.1.4。验证日期：2026-09-06。0.1.4 为纯 UI 变更（移除拖拽全屏遮罩，拖放行为不变），宿主连接点与 0.1.2 完全相同，兼容性结论沿用 0.1.2 核验；22 项自动测试全过。

## 已核验环境

| 项目 | 核验范围 |
| --- | --- |
| DSH Desktop | 本机应用版本 2.0.4，读取其安装包和内置源码 |
| DSH 内核 | 应用内置 `@deepseek-ai/dsh@0.1.2-alpha.1`，另起隔离 Web 实例运行 |
| Node | 本机 26.8.1；package engines 沿用宿主最低要求，其他 Node 版本未实测 |
| 文件解析 | 本机 macOS；anydoc 0.1.9 预编译本地模块，旧 DOC 可由系统 textutil 兜底 |
| 正式桌面安装 | 0.1.1 已完成真实模型读取、修改、新建与下载验证；0.1.2 收拢输入区并迁移设置入口 |
| Mac mini / 其他系统 | 尚未验收；不能仅因共享工作区就视为已安装 |

## 与宿主的连接点

- npm 包 `dsh.bundle.patch` 声明 host 挂载；`dsh.client` 声明 web 客户端与注入依赖。
- 客户端产物使用宿主 `window.__ModuleLoader__.load`，React 与 DSH 模块由宿主提供。
- 单个回形针按钮注入 `conversation.input.left`；附件视图经 `conversation.input.dock` 挂载，用 React portal 定位到 `[data-composer-card] [data-slot="conversation.input.attachments"]`。
- 参数界面注入 `settings.general.item`，显示在 DSH 的常规设置页；会话输入区不再显示文件面板和设置按钮。
- 发送包装器调用原型上的原始 `sendSession(session,text,imageIds,mode,signal)`。成功条件沿用 `{kind:'success'}`，其余结果恢复本批附件。
- 历史卡片适配当前版本 `[data-time-hover-root] [class*="_bubble"]`。原始消息和完整附件文本仍保留。
- host 注入 `webServer`、`workspaceRegistry`、`tools`、`settings`。工具参数使用宿主 ToolDefinition 的 JSON Schema；不依赖 npm 未发布的 alpha builder 包。
- 文件只归属于工作区注册表中的目标会话；HTTP 只支持本机回环连接。

DSH 升级后需重新核对上述接口。DOM 适配失效时，可能出现卡片位置变化或消息显示完整提取文本；正式升级前先在隔离 profile 重做输入、发送、图片混合和重启恢复验收。

## 依赖与安装

顶层运行依赖在 package.json 固定版本，开发目录附 package-lock.json。anydoc 使用随 npm 包分发的本机架构二进制，不在安装时编译；本机 `--ignore-scripts` 安装已成功。其他机器仍需能取得对应平台的 optional dependency。

安装前需要联网下载依赖，文档解析时不调用云端解析服务。文档发送给所选模型遵循 DSH 自身的模型配置。

安装器默认仅显示计划，正式执行需 `--apply`。它检查桌面端/内核版本，完整备份插件环境、会话、索引、附件和设置，用官方 CLI 安装 tgz，通过覆盖层禁用 DragView 并保留 bundle 与依赖，最后核对有效配置。完整回退在模拟环境实际运行通过；日常环境未为演练而回退。

## 文件边界

DOCX/XLSX/PPTX 的创建与结构读取已有自动测试；这不代表在 Word/WPS/PowerPoint 中完成了逐页视觉验收。Office 预览显示提取文字和数据，不复制原版式。XLSX 修改不计算公式，清除缓存并设置打开重算；复杂图表、修订痕迹、外部链接等需专项样例验证。

PDF/HTML 提供页面预览入口，实际页面渲染尚待桌面验收。HTML 被限制在沙箱中，只支持内嵌脚本、样式和 data 图片，不加载远程资源。
