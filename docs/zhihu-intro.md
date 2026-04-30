# onenote-cli: 让你的 OneNote 笔记在 AI 时代存活下来

你的 OneNote 笔记本里藏着多少年的记忆？工作笔记、学习资料、日记、灵感碎片……这些内容，AI 能帮你搜索吗？

答案是：现在可以了。

onenote-cli 是一个用 Bun + TypeScript 构建的命令行工具，通过 Microsoft Graph API 直接操作你的 OneNote 笔记本。最关键的功能：**全文搜索，精确到页面级别，点击 URL 直接跳转到匹配的那一页。**

## 为什么需要这个？

OneNote 的搜索功能只能在桌面端或网页端使用，无法被 AI 工具调用。当你想让 AI 帮你找一条几年前的笔记时，它做不到——因为 OneNote 没有暴露搜索 API 给第三方。

更糟的是，如果你的 OneDrive 里有超过 5000 个 OneNote 项目（笔记本+分区+分区组），微软的 Graph API 会直接返回 403 错误，连列出分区都做不到。

onenote-cli 解决了这两个问题。

## 它能做什么？

### 全文搜索（精确到页面）

```bash
$ onenote search "项目计划"

# (20240315) Q2 项目计划
  Section: Work Notes | Notebook: My Notebook
  ...下季度**项目计划**：1. 完成用户认证模块重构 2. 上线新的推荐算法...
  https://onenote.com/...?wd=target(...)

3 page-level results found.
```

搜索结果直接给出 OneNote Online 的页面级 URL，点击即可跳转到对应页面。不是打开整个分区，而是精确到那一页。

### 笔记本管理

```bash
$ onenote notebooks list
$ onenote sections list -n <notebook-id>
$ onenote pages create -s <section-id> -t "Meeting Notes" -b "<p>内容</p>"
```

### 5000 项目限制的解决方案

当 Graph API 因为 SharePoint 文档库超过 5000 项而返回 403 时，onenote-cli 会：

1. 通过 OneDrive API 直接下载 `.one` 二进制文件
2. 从二进制中提取页面内容（支持 UTF-8 和 UTF-16LE，中日韩文字都能搜到）
3. 提取页面 GUID，通过 OneNote API 获取官方页面 URL
4. 构建本地缓存索引

### AI 集成

作为 Claude Code Skill 一键安装：

```bash
$ npx skills add snomiao/onenote-cli
```

安装后，AI 可以直接搜索你的 OneNote 笔记，输出 Markdown 格式的结果：

```markdown
[Q2 项目计划](https://onenote.com/...?wd=target(...))
  Work Notes | My Notebook
  ...下季度**项目计划**：1. 完成用户认证模块重构...
```

## 技术亮点

- **设备代码流认证**：支持 SSH / 无头环境
- **.one 二进制解析**：从 MS-ONESTORE 格式中提取页面文本和 GUID
- **官方页面 URL**：通过 `/me/onenote/sections/0-{guid}/pages` 端点获取，绕过 OneNote Online 的会话缓存
- **跨目录运行**：`.env.local` 和缓存从包目录自动加载

## 开始使用

```bash
git clone https://github.com/snomiao/onenote-cli.git
cd onenote-cli
bun install
cp .env.example .env.local  # 填入你的 Azure AD Client ID
bun run src/index.ts auth login
bun run src/index.ts sync
bun run src/index.ts search "你想找的内容"
```

详细的 Azure AD 配置步骤见 GitHub 仓库的 [docs/setup.md](https://github.com/snomiao/onenote-cli/blob/main/docs/setup.md)。

---

**GitHub**: https://github.com/snomiao/onenote-cli

MIT License | by snomiao
