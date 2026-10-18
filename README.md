# WordTool

基于 Chrome / Edge 油猴（Tampermonkey）的**选词发音 & 翻译工具**。

双击单词即可真人发音 + 释义，支持技术标识符智能拆读（`deepseek-api`、`userMessageDelta` 等），ALT 拖选翻译，整页翻译，输入框三空格一键译成英文。

---

### 📱 兼容性

在使用脚本之前，你需要为浏览器安装脚本管理器插件。

#### 🌐 浏览器支持（任选其一即可）

- **Chrome** 或基于 Chromium 内核的浏览器（Edge、360 浏览器、CentBrowser 等）*推荐*
- Firefox（功能兼容）

#### 🙈 需要安装的脚本管理器（任选其一即可）

- [Tampermonkey（新油猴）](https://www.tampermonkey.net/) *推荐*
- [Violentmonkey（暴力猴）](https://violentmonkey.github.io/get-it/)

> 注：Greasemonkey（旧油猴）因过于古老，不保证完全兼容。

#### 🧪 支持的操作系统

Windows、macOS、Linux 等支持 Chrome / Edge / Firefox 的操作系统。  
（移动端如 iOS、Android 暂不支持）

---

### 📄 安装脚本

1. 先安装上面的脚本管理器（Tampermonkey 推荐）
2. 点击下方安装链接安装脚本：

**[点击安装 WordTool](https://github.com/你的用户名/WordTool/raw/main/WordTool.user.js)**

（把链接换成你实际上传后的 raw 地址）

或者手动创建：
1. 打开 Tampermonkey 管理面板 → 添加新脚本
2. 把本仓库的 `WordTool.user.js` 全部内容复制粘贴进去
3. 保存即可

---

### ✨ 主要功能

| 功能 | 说明 |
|------|------|
| **双击发音** | 双击任意英文单词，自动播放有道真人发音（美音/英音可切换），并显示音标和释义 |
| **技术标识符智能拆读** | 支持 `deepseek-api`、`user_message_delta`、`XMLHttpRequest`、camelCase 等，自动拆成子词顺序朗读，并显示逐词释义 |
| **技术词读法优化** | 内置常见开发术语读法（pytest → pie test、github → git hub、typescript → type script 等） |
| **ALT 拖选翻译** | 按住 `Alt` 键拖动选中文字，使用浏览器本地 Translator API 翻译成中文 |
| **整页翻译** | 点击左下角浮动按钮，开启/关闭整页双语对照翻译 |
| **输入框三空格译英** | 在任意输入框/文本域中连续按 3 次空格，自动把当前内容翻译成英文并替换 |
| **有道音频响度归一化** | 自动调整不同单词的音量，听感更统一 |
| **CSP 兼容** | 通过 Web Audio 内存播放，兼容 GitHub 等严格 CSP 网站 |

---

### 🎮 使用方法

1. **发音**：双击页面上的英文单词或技术标识符即可
2. **选词翻译**：按住 `Alt` 键，拖动鼠标选中一段文字
3. **整页翻译**：点击页面左下角蓝色浮动按钮（可拖动位置）
4. **输入框译英**：在输入框里输入中文后，连续按三下空格
5. **设置**：油猴菜单 →「Word Tool 设置」，可切换触发方式（双击 / Ctrl+单击 / Ctrl+悬浮）和美音/英音

按 `Esc` 可关闭当前弹窗并停止发音。

---

### ⚙️ 设置说明

通过油猴扩展图标 → Word Tool 设置：

- **触发方式**
  - 双击选词自动发音（默认）
  - CTRL + 单击发音
  - CTRL + 悬浮发音
- **有道语音类型**：美音 / 英音
- **检测 Translator API**：检查当前浏览器是否支持本地翻译（Edge 148+ 推荐）

---

### 🛠️ 技术说明

- 发音优先使用有道真人音频，失败时回退到浏览器 Speech Synthesis
- 翻译使用浏览器内置 Translator API（本地处理，无需 API Key）
- 支持 camelCase、snake_case、kebab-case 等技术命名自动拆分
- 有道音频自动进行响度归一化，减少不同单词音量差异

---

### 🚀 贡献

欢迎提交 Issue 反馈问题，或直接提 Pull Request。

修改版本时请同步更新：
- 脚本头部的 `@version`
- README 中的版本说明（如有）

---

### 📜 License

GPL 3

---

如果这个脚本对你有帮助，欢迎给个 Star ⭐