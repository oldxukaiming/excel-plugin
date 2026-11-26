# Excel 双引擎翻译插件（Bing / Google）

一个基于 **Office.js** 的现代 Excel 侧边栏加载项（兼支持自定义公式），支持在 **Bing 翻译** 与 **Google 翻译** 之间一键切换，将 `A1` 等单元格（或批量区域）文本翻译后自动填充至 `A2`、右侧单元格或指定目标区域。

## 核心特性

1. **双引擎切换**：支持 `Bing 翻译` 与 `Google 翻译` 随时切换。
2. **免 Key 优先 + 官方 Key 可选**：
   - 默认无需 API Key 即可使用内置免费接口（Bing Edge/Web 接口 & Google 公共翻译接口）。
   - 在侧边栏“高级设置”中（或通过环境变量）支持填入 Google Cloud Translation Key / 自定义镜像端点，以及 Azure Translator Key / Region。
3. **直观单元格映射（如 A1 → A2）**：
   - 默认模式为 **紧邻下方单元格**（选中 `A1` 时自动计算目标为 `A2`，多行区域自动整体下移）。
   - 支持 **紧邻右侧单元格**（`A1 → B1`）、**自定义目标单元格**（如 `C1`、`Sheet2!A1`）以及 **原地覆盖**。
   - 支持实时跟随 Excel 当前选区，并支持多单元格批量翻译。
4. **自定义公式（可选）**：
   - `=XLTRANS.TRANSLATE(A1, "zh-CN", "auto", "bing")`（第 4 参数可改为 `"google"`）。

## 快速开始

### 1. 安装依赖与生成图标/证书

```powershell
npm install
npm run icons
npm run cert:trust
```

> 注：`npm run cert:trust` 会生成本地开发用的 `localhost` HTTPS 证书并提示安装受信任根证书（Office 加载项强制要求 HTTPS）。如果只生成证书文件而不自动安装系统信任，可执行 `npm run cert:gen`，然后手动将 `.certs/localhost.crt` 导入“受信任的根证书颁发机构”。

### 2. 启动本地翻译服务

```powershell
npm start
```

服务默认同时监听：
- `https://localhost:3443`（Office.js 侧边栏加载项使用）
- `http://127.0.0.1:3444`（伴随 HTTP 接口，供 VBA 宏免证书调用）

你可以在浏览器打开 `https://localhost:3443/taskpane.html` 预览界面。

### 3. 在 Excel 中加载插件（无需文件夹共享 / 无需管理员权限）

由于受限机器无法设置网络共享文件夹，请按以下任一方式加载：

#### 方式一：当前用户注册表本地旁加载（推荐，保留完整 Office.js 侧边栏）

直接将 `manifest.xml` 注册到当前用户（`HKCU`）的 Office 开发人员加载项目录，并把证书导入当前用户根证书存储（均不需要管理员权限与文件夹共享）：

```powershell
npm run addin:register
```

完成后重启 Excel，进入 **插入 / 开始** → **加载项** → **我的加载项** → **开发人员加载项**（或直接在“开始”选项卡右侧看到 **翻译面板**）。如需取消注册可运行 `npm run addin:unregister`。

#### 方式二：Excel 网页版直接上传 `manifest.xml`

如果您的企业有 Microsoft 365 网页版 Excel，可直接打开网页版工作簿 → **加载项** → **更多加载项** → **上传我的加载项**，选中项目根目录的 `manifest.xml` 即可。

#### 方式三：独立网页翻译工作台（免 VBA、免共享、免加载项，受限电脑 100% 可用）

如果本机既无法设定共享，又提示“初始化 VBA 库出错 1004”（企业策略禁用或精简了 VBA 组件），可直接打开内置工作台：

1. 启动服务：`npm start`（或 `npm run start:http`）
2. 浏览器访问：`http://127.0.0.1:3444/workbench.html`（或 `https://localhost:3443/workbench.html`）
3. 工作台提供三种无权限要求的使用方式（均支持 Bing / Google 引擎切换）：
   - **`.xlsx` 文件直译**：直接上传本地 `.xlsx`，把 `A1` 或 `A1:A20` 翻译填充到 `A2`（或右侧/自定义区域）后下载新文件。
   - **Excel 窗口联动**：尝试直接读取您已手动打开的 Excel 窗口选中单元格并写回目标格。
   - **复制粘贴快翻**：从 Excel `Ctrl+C` 复制单元格或整列文本，翻译后保持行列结构一键复制，切回 Excel `A2` 直接 `Ctrl+V`。

#### 方式四：纯本地 VBA 加载宏（仅适用于未禁用 VBA 的环境）

如果系统组策略连 Office Web 加载项也禁用，可使用本项目附带的 VBA 模块 [src/vba/ExcelTranslator.bas](src/vba/ExcelTranslator.bas)：

1. 保持服务运行：`npm start`（或 `npm run start:http`）。
2. 自动构建 `.xlam` 加载宏并安装到个人 `%APPDATA%\Microsoft\AddIns` 目录：
   ```powershell
   npm run vba:build
   ```
   > 若 Excel 安全设置禁止脚本访问 VBA 对象模型，可在 Excel 按 `Alt + F11` 打开 VBA 编辑器，右键左侧项目 → **导入文件** → 选择 `src/vba/ExcelTranslator.bas`（可另存为 `.xlam` 加载宏）。
3. 可用的宏与函数：
   - 宏 `TranslateSelectionBelow`：把选中单元格（如 `A1`）翻译填充到下方单元格（如 `A2`），可绑定到快速访问工具栏或快捷键。
   - 宏 `TranslateSelectionRight`：翻译填充到右侧单元格（如 `A1 → B1`）。
   - 宏 `SwitchTranslateEngine` / `SetEngineBing` / `SetEngineGoogle`：在 **Bing** 与 **Google** 间切换引擎。
   - 宏 `ConfigureLanguages`：配置源语言与目标语言。
   - 公式 `=XLTRANS(A1, "zh-CN", "auto", "bing")`：直接在单元格通过公式翻译（第 4 参数可切换 `"bing"` / `"google"`）。

## 目录结构

- `manifest.xml`：Office 加载项清单文件
- `cloudflare/worker.js` & `wrangler.toml`：Cloudflare Workers 一体化部署配置（API + 静态前端）
- `src/server/`：Node.js + Express 本地 HTTPS/HTTP 服务与 Bing / Google 翻译引擎
- `src/client/`：Office.js 侧边栏（`taskpane.*`）、独立工作台（`workbench.*`）、浏览器端 `.xlsx` 处理（`xlsx-browser.js`）与自定义函数（`functions.*`）
- `scripts/`：图标生成、HTTPS 证书生成、静态站点构建与清单校验脚本
- `tests/`：单元测试与接口集成测试

## 部署到 Cloudflare（推荐，绑定自定义域名 + 免费 HTTPS）

### 方式一：Cloudflare Workers 一体化部署（前后端同域，最简单）

1. 在 Cloudflare Dashboard → **Workers & Pages** → **Create** → 选择连接 GitHub 仓库 `oldxukaiming/excel-plugin`（或在本地执行 `npx wrangler login && npx wrangler deploy`）。
   - 构建命令：`npm run build:static`（已在 [wrangler.toml](file:///d:/Users/xukai/Desktop/excel-plugin/wrangler.toml) 中配置）
   - 若要让下载的 `manifest.xml` 自动指向您的域名，可在构建环境变量设置 `PUBLIC_URL=https://你的域名`。
2. 在 Worker 的 **Settings → Domains & Routes** 添加您的自定义域名（例如 `excel.yourdomain.com`），Cloudflare 会自动签发免费 HTTPS 证书。
3. 浏览器打开 `https://excel.yourdomain.com` 即可使用（`.xlsx` 文件全程在浏览器本地解析，翻译请求自动走同域 `/api/translate/batch`）。

### 方式二：GitHub Pages 托管前端 + Cloudflare Worker 托管 API

1. 在 GitHub 仓库 **Settings → Pages** 中将 Source 设为 **GitHub Actions**（工作流已内置在 [.github/workflows/deploy-pages.yml](file:///d:/Users/xukai/Desktop/excel-plugin/.github/workflows/deploy-pages.yml)）。
2. 在 GitHub 仓库 **Settings → Secrets and variables → Actions → Variables** 添加：
   - `API_BASE_URL`：您的 Cloudflare Worker 地址（例如 `https://api-excel.yourdomain.com`）
   - `PUBLIC_URL`：前端访问地址（例如 `https://excel.yourdomain.com`）
   - `PAGES_CNAME`（可选）：前端自定义域名（例如 `excel.yourdomain.com`）
