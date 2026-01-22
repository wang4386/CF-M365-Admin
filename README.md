# CF-M365-Admin (二开版)

基于 Cloudflare Workers 的 Microsoft 365 E5 开发者订阅自助管理系统。

更新说明： 本版本对核心逻辑进行了重构，引入了 **“强制邀请码注册”** 机制，配合 Cloudflare KV 存储实现邀请码的生命周期管理。同时重绘了前端 UI，提供更现代化的用户体验，并增强了后台安全性。

## ✨ 主要特性

* **自助注册 (前台)**
  * **强制邀请码** ：注册必须使用管理员分发的邀请码，杜绝恶意注册。
  * **安全验证** ：集成 Cloudflare Turnstile 人机验证。
  * **多订阅支持** ：支持前台选择不同订阅（通过 SKU 映射隐藏真实 ID）。
* **管理后台 (后台)**
  * **邀请码管理** ：批量生成邀请码、一键复制未使用码、查看邀请码使用状态（绑定用户/使用时间）。
  * **安全登录** ：独立的后台登录页面，支持 Turnstile 验证和 HttpOnly Cookie 鉴权。
  * **智能联动** ： **删除用户的同时，系统会自动识别并清理该用户绑定的邀请码记录** ，保持数据整洁。
  * **用户管理** ：支持按时间排序、查看订阅状态、批量删除用户。
  * **库存查询** ：实时查询 Azure 租户内的许可证总数与剩余数。
* **系统机制**
  * **KV 存储** ：利用 Cloudflare KV 存储邀请码数据，读写速度快。
  * **灵活配置** ：支持自定义后台路径、隐藏管理员账号。

## 📸 界面预览

| 注册页面 | 管理登录 |
| --- | --- |
| ![https://assets.qninq.cn/qning/7F1zp2sk.webp](https://assets.qninq.cn/qning/7F1zp2sk.webp) | ![https://assets.qninq.cn/qning/aMxfqka5.webp](https://assets.qninq.cn/qning/aMxfqka5.webp) |

| 用户管理 | 邀请码管理 |
| --- | --- |
| ![https://assets.qninq.cn/qning/Dx8cw9Vq.webp](https://assets.qninq.cn/qning/Dx8cw9Vq.webp) | ![https://assets.qninq.cn/qning/wrIRpG7y.webp](https://assets.qninq.cn/qning/wrIRpG7y.webp) |

## 🛠️ 部署前提

1. **Cloudflare 账号** ：用于部署 Workers 和配置 KV。
2. **Microsoft 365 全局管理员权限** ：用于注册 Azure AD 应用。
3. **Azure AD 应用 (App Registration)** ：
   * 注册一个新应用。
   * 获取 `Client ID` (客户端 ID) 和 `Tenant ID` (租户 ID)。
   * 生成 `Client Secret` (客户端密钥)。
   * **API 权限 (必须)** ：
     * 类型：**Application (应用程序权限)**
     * 权限：`User.ReadWrite.All`, `Directory.ReadWrite.All`
     * **重要** ：添加后必须点击“代表 xxx 授予管理员同意 (Grant admin consent)”。
4. **Cloudflare Turnstile** ：
   * 在 Cloudflare 后台申请 Turnstile，获取 `Site Key` 和 `Secret Key`。

## 🚀 部署指南

### 1. 创建 Cloudflare Worker

1. 在 Cloudflare Dashboard 创建一个新的 Worker。
2. 复制本项目 `worker.js` 的完整代码并覆盖。

### 2. 配置 KV 存储 (关键步骤)

**注意：本版本强依赖 KV 存储邀请码，不配置将无法运行。**

1. 在 Cloudflare Workers 首页左侧菜单找到  **KV** 。
2. 点击  **Create a Namespace** ，命名为 `m365_invites` (名字可自定义)。
3. 进入你的 Worker -> **Settings** ->  **Variables** 。
4. 在 **KV Namespace Bindings** 处点击  **Add binding** 。
5. **Variable name** 必须填：`INVITE_CODES` (注意全大写)。
6. Namespace 选择刚才创建的 `m365_invites`。
7. 点击 **Deploy** 保存。

### 3. 配置环境变量

在 Worker -> **Settings** -> **Variables** 中添加以下环境变量：

| 变量名                    | 说明                 | 示例 / 备注                                    | 是否必须  |
| --------------------------- | ---------------------- | ------------------------------------------------ | ----------- |
| `AZURE_TENANT_ID`     | Azure 租户 ID        | `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`     | ✅        |
| `AZURE_CLIENT_ID`     | Azure 应用 ID        | `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`     | ✅        |
| `AZURE_CLIENT_SECRET` | Azure 应用密钥       | `Value`(注意不是 Secret ID)                | ✅        |
| `CF_TURNSTILE_SECRET` | CF Turnstile 密钥    | 后端验证用 Secret Key                          | ✅ (强制) |
| `TURNSTILE_SITE_KEY`  | CF Turnstile 站点 ID | 前端显示用 Site Key                            | ✅ (强制) |
| `DEFAULT_DOMAIN`      | 默认域名后缀         | `example.onmicrosoft.com`(不带 @)          | ✅        |
| `ADMIN_TOKEN`         | 后台访问密码         | `MySuperSecretToken123`                    | ✅        |
| `SKU_MAP`             | 订阅映射表           | 见下方获取教程                                 | ✅        |
| `ADMIN_PATH`          | 后台访问路径         | `/admin`或`/console`(默认`/admin`) | ❌        |
| `HIDDEN_USER`         | 隐藏管理员账号       | `admin@example.onmicrosoft.com`            | ❌        |
| `ENABLE_DEBUG`        | 调试模式             | `true`                                     | ❌        |

### **请务必注意，所有变量的变量类型都必须使用 Text 或 Secret！包括 SKU_MAP 也是！不要用 JSON 变量类型！**

## ⚙️ 关于 SKU_MAP 的获取

两种获取方式（都需要参考下方的json进行自行填写）

json格式：`{"E5开发版":"你的SKU_ID_1", "A1学生版":"你的SKU_ID_2"}`

**第一种方式（推荐）** ：
部署前，你可以先随便填一个 JSON，访问 `/admin` 登录后台，点击“ **订阅用量** ”按钮，弹窗中会直接显示你租户下所有真实的 SKU ID 和名称。

**第二种方式** ：
部署并设置好变量后，先访问 `/admin` 进行登录。登录成功后，手动在浏览器访问： `https://你的worker域名.workers.dev/你的后台路径/api/licenses`
页面会返回一段 JSON。找到 `skuPartNumber` 是 `DEVELOPERPACK_E5` 或者 `ENTERPRISEPACK` (即 E3) 的那一项。

## 📖 使用指南

### 用户端

* 访问 `https://你的worker域名.workers.dev`
* 用户需输入有效的邀请码、选择订阅、输入用户名密码即可创建。

### 管理端

* 访问 `https://你的worker域名.workers.dev/admin` (或你自定义的 `ADMIN_PATH`)
* 输入 `ADMIN_TOKEN` 登录。
* **邀请码管理** ：
  * 点击“生成邀请码”批量生成。
  * 点击“复制所有未使用”可快速导出。
* **用户管理** ：
  * 支持批量勾选删除用户（删除时会自动释放/清理对应的邀请码记录）。

## ⚠️ 免责声明

本项目仅供学习和内部管理使用。

* 请勿用于非法分发 Microsoft 账号。
* 作者不对因配置错误导致的数据丢失或账号被封禁承担责任。
* 涉及删除操作，请务必谨慎。

**License** : MIT
