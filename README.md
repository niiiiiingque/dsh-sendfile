# 发文件 · dsh-sendfile 0.1.2

为 DSH Desktop 定制的文档工作流：把文件放入对话，AI 读到提取内容，按要求修改副本或生成新文件，再预览、下载结果。

本版对应 DSH Desktop 2.0.4 内置内核 `0.1.2-alpha.1`。已完成隔离验证，并在本机 desktop profile 验证真实模型读取附件、修改 Word 和生成 Excel。0.1.2 将输入区收拢为单个回形针图标，移除会话“文件”面板和齿轮按钮，把参数移到 DSH 的常规设置页。最终桌面验收结果以随附实机验收记录为准。

## 使用方式

1. 在 DSH 选择带工作区的会话。点击输入框内的回形针图标，或把文档拖入窗口。
2. 卡片出现在**输入框内部上沿**。等待「上传中 → 解析中 → 已就绪」，输入要求并发送。仅有附件也能发送。
3. 文档文字自动随消息进入模型上下文。对话流保留文件卡片，点击可看内容和下载。
4. 让 AI 修改或生成文件，例如「把实收净重 198.6 改成 199.2，另存修订稿」或「把以上数据做成 Excel」。
5. AI 生成或修改的文件会在回复中提供预览、下载链接，无需进入单独的文件面板。

文档字符数和单文件大小上限位于 DSH「设置 → 常规 → 发文件插件」。

图片继续使用原生图片入口。文档与图片的混合拖拽处理已实现，桌面端实际混合拖拽仍待安装后验收。

## 支持范围

| 操作 | 本版支持 | 边界 |
| --- | --- | --- |
| 读取 | DOCX、DOC、XLSX、XLS、PPTX、文字 PDF、MD、TXT、CSV | 图片、图表、扫描页的视觉内容不进入文字提取结果 |
| 新建 | DOCX、XLSX、PPTX、MD、TXT、CSV、HTML | Word 为简单 A4 正文；PPTX 为简单标题正文版式 |
| 修改 | DOCX 精确文字替换，XLSX 指定单元格改值，文本整体修订 | 只写新副本；DOCX 不跨段替换；不提供通用 Office 页面编辑器 |
| 预览 | Office 提取内容；PDF、HTML 页面预览入口 | Office 内容预览不等于原版式；PDF/HTML 页面入口的实际渲染兼容性待桌面验收 |
| 登记输出 | 工作区内的支持格式文件 | 其他工具生成的 PDF 等可在模型回复中返回预览和下载链接 |

不支持旧版 `.ppt` 上传，会明确提示另存为 `.pptx` 或 PDF。不做扫描 PDF OCR、语音转写、原地覆盖、旧版 DOC/XLS 回写、原稿 PDF 编辑。

DOCX 修改默认清空全部分节的普通/首页/偶数页页眉页脚及相关引用，正文域保持。跨字体 run 的精确替换可用，替换文字采用起始 run 的格式。复杂修订痕迹、域内部文字、文本框与特殊结构仍需人工核对。

XLSX 修改保留已支持的单元格样式与公式表达式，清除公式缓存结果并设置打开时重算。插件不计算公式；用 Excel/WPS 打开后核对。复杂图表、数据透视、外部连接等高级功能不作保真承诺。

## 本地文件与数据

- 读取前复制到 `<工作区>/.dsh-sendfile/<会话摘要>/<随机ID>/original/<原文件名>`。
- 原件、提取文字、元数据分别保存；同名文件使用不同 ID，不覆盖。
- 原文件保持原位；输出另建 ID，可追溯 `sourceId`。
- 没有数据库、云端解析、OCR API 或宏执行。旧 DOC 在 macOS 上可使用系统 `textutil` 本地兜底。
- **点击发送后，提取文字会交给 DSH 当前配置的模型。**「本地解析」不表示模型推理离线。插件不会改变 DSH 本身的模型、遥测或联网设置。
- 插件文件接口仅接受本机回环连接，拒绝跨站请求；写接口检查随机令牌。客户端不能指定任意服务器路径。
- 默认每条消息文档总计最多 200,000 个 JavaScript 字符，单文件 32 MB；设置可调为 1,000–200,000 字符、1–64 MB。
- 超长内容按段落/行边界截断，单段过长时按字符截断。消息包含截断标记，AI 可用 `sendfile_read` 分段读取后文。500 万字符以上需拆分。
- 所有附件按发送时的目标会话归属。没有读到的内容不能视作模型已经读取。

## 提供给 AI 的工具

| 工具 | 用途 |
| --- | --- |
| `sendfile_read` | 读取已登记文档的指定字符段 |
| `sendfile_create` | 新建文档并返回预览、下载链接 |
| `sendfile_edit` | 核对旧文字/旧值后修改，另存副本 |
| `sendfile_publish` | 登记同工作区其他工具的输出并返回预览、下载链接 |
| `sendfile_list` | 列出本会话附件和输出 |

示例参数：

```json
{"name":"报告.docx","spec":{"title":"交付情况","text":"本月计划交付200吨。\n实收198.6吨。"}}
```

```json
{"sourceId":"附件ID","name":"修订稿.docx","spec":{"replacements":[{"find":"198.6","replace":"199.2","expectedCount":1}]}}
```

```json
{"sourceId":"附件ID","name":"数据-修订.xlsx","spec":{"cells":[{"sheet":"交付","cell":"C2","expected":198.6,"value":199.2}]}}
```

## 安装与回退

**需要用户确认后才执行正式安装。**安装会更改 `~/.dsh/profiles/desktop`。为避免输入事件冲突，通过 profile 覆盖层 `- id: drag-file`、`disabled: true` 停用 DragView；保留其 bundle 登记、安装包和原配置，便于回退。必须先退出 DSH Desktop，结束正在运行的任务。

随源码交付的 `scripts/install.mjs` 默认只打印计划；加 `--apply` 才安装。运行入口是 macOS 终端或本机 Agent。脚本检查桌面端与内核版本，调用 `scripts/recovery.py` 完整备份并校验插件环境、会话、索引、附件和设置，禁用 DragView 后通过官方 CLI 安装同目录 tgz。

```sh
node scripts/install.mjs
# 获得确认且退出 DSH 后：
node scripts/install.mjs --apply
```

安装器不会自动重启应用，不改全局 AGENTS、Skills、API 密钥或 iCloud 记忆。Mac mini 应单独安装；共享文件不等于插件已经安装。

若需回退，先退出 DSH，运行 `node scripts/rollback.mjs <安装器打印的备份目录> --apply`。脚本校验并恢复完整旧插件环境（含依赖），将当前环境改名保留；不覆盖会话、项目索引和附件。之后重启。恢复工具经过模拟故障与真实文件恢复测试。

## 开发与验证

```sh
npm ci --ignore-scripts
npm run build
npm test
npm run check
```

构建无需访问用户 DSH 配置。测试用虚构样例，HTTP 测试仅监听 `127.0.0.1`，旧 DOC 测试需要 macOS 的 `textutil`。测试夹具 `legacy.xls` 为生成的 BIFF8 文件，不含真实业务数据。

当前内核的 `@deepseek-ai/dsh-tools@0.1.2-alpha.1` 未在 npm 发布，所以直接采用宿主公开的 `ToolDefinition` JSON Schema 注册契约，不安装不同版本的工具内核。所有模型参数在执行前校验。

待发送卡片通过 React portal 放入原生输入框附件区域，图片组件继续由 DSH 管理。发送适配器调用原有 `sendSession`，保留图片、queue/steer、取消与乐观回显流程。卸载时还原。

消息卡片采用当前版本 DOM 展示适配，原始消息节点和完整记录保持；卸载插件后会显示完整附文。DSH 升级后必须检查输入框选择器、`sendSession`、消息气泡结构与工具注册接口。

## 依据

- 用户提供的《DSH定制插件-发文件-来龙去脉与规格.md》，作为需求背景审阅。
- 本机 DSH Desktop 2.0.4 的内置内核源码、slot 声明、`sendSession`、工具和设置注册接口。
- [DeepSeek Harness 官方仓库](https://github.com/deepseek-ai/deepseek-harness)与[官方 Cordis 开发教程](https://deepseek-harness.github.io/deepseek-harness/en/develop/cordis-tutorial/)。
- [anydoc 原始项目](https://github.com/firecrawl/anydoc)，本插件锁定本机验证过的 `0.1.9` 本地解析入口，不启用 hosted OCR。
- 参考 DragView 的 npm bundle 结构与 DSH 模块加载包装方式；本插件的解析、保存、发送状态和 UI 为新实现。
