/**
 * 核心配置变量 (需要在 Cloudflare 后台配置):
 * AZURE_TENANT_ID: 租户 ID
 * AZURE_CLIENT_ID: 客户端 ID
 * AZURE_CLIENT_SECRET: 客户端密钥
 * CF_TURNSTILE_SECRET: Turnstile Secret Key
 * TURNSTILE_SITE_KEY: Turnstile Site Key
 * DEFAULT_DOMAIN: 你的邮箱后缀 (不带@)
 * SKU_MAP: JSON字符串，映射前台名称到SKU ID。例: {"E5开发版":"你的SKU_ID_1", "A1教育版":"你的SKU_ID_2"}
 * ADMIN_TOKEN: 管理员访问密码
 * HIDDEN_USER: (可选) 隐藏的特权账户完整邮箱
 * ENABLE_DEBUG: (可选) 设置为 'true' 开启调试日志
 */

const debugLog = (env, ...args) => {
    if (env.ENABLE_DEBUG === 'true') console.log('[DEBUG]', ...args);
};

// --- 辅助函数：密码强度校验 (四选三) ---
function checkPasswordComplexity(pwd) {
    if (!pwd || pwd.length < 8) return false;
    let score = 0;
    if (/[a-z]/.test(pwd)) score++; // 小写
    if (/[A-Z]/.test(pwd)) score++; // 大写
    if (/\d/.test(pwd)) score++;    // 数字
    if (/[^a-zA-Z0-9]/.test(pwd)) score++; // 符号
    return score >= 3;
}

// --- GitHub 图标 SVG (硬编码复用) ---
const GITHUB_ICON = `<svg viewBox="0 0 16 16" version="1.1" width="20" height="20" aria-hidden="true" fill="currentColor"><path d="M8 0C3.58 0 0 3.58 0 8c0 3.54 2.29 6.53 5.47 7.59.4.07.55-.17.55-.38 0-.19-.01-.82-.01-1.49-2.01.37-2.53-.49-2.69-.94-.09-.23-.48-.94-.82-1.13-.28-.15-.68-.52-.01-.53.63-.01 1.08.58 1.23.82.72 1.21 1.87.87 2.33.66.07-.52.28-.87.51-1.07-1.78-.2-3.64-.89-3.64-3.95 0-.87.31-1.59.82-2.15-.08-.2-.36-1.02.08-2.12 0 0 .67-.21 2.2.82.64-.18 1.32-.27 2-.27.68 0 1.36.09 2 .27 1.53-1.04 2.2-.82 2.2-.82.44 1.1.16 1.92.08 2.12.51.56.82 1.27.82 2.15 0 3.07-1.87 3.75-3.65 3.95.29.25.54.73.54 1.48 0 1.07-.01 1.93-.01 2.2 0 .21.15.46.55.38A8.013 8.013 0 0016 8c0-4.42-3.58-8-8-8z"></path></svg>`;

// --- 1. 前台注册页面 HTML (大幅美化版) ---
const HTML_REGISTER_PAGE = (siteKey, skuOptions) => `
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Office 365 自助开通</title>
    <script src="https://challenges.cloudflare.com/turnstile/v0/api.js" async defer></script>
    <style>
        :root {
            --primary: #4f46e5;
            --primary-hover: #4338ca;
            --bg-gradient: linear-gradient(-45deg, #ee7752, #e73c7e, #23a6d5, #23d5ab);
            --glass-bg: rgba(255, 255, 255, 0.85);
            --glass-border: rgba(255, 255, 255, 0.4);
            --text-main: #1f2937;
            --text-sub: #6b7280;
        }
        @keyframes gradient {
            0% { background-position: 0% 50%; }
            50% { background-position: 100% 50%; }
            100% { background-position: 0% 50%; }
        }
        @keyframes fadeInUp {
            from { opacity: 0; transform: translateY(20px); }
            to { opacity: 1; transform: translateY(0); }
        }
        body {
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
            background: var(--bg-gradient);
            background-size: 400% 400%;
            animation: gradient 15s ease infinite;
            display: flex;
            justify-content: center;
            align-items: center;
            min-height: 100vh;
            margin: 0;
            color: var(--text-main);
        }
        .card {
            background: var(--glass-bg);
            backdrop-filter: blur(12px);
            -webkit-backdrop-filter: blur(12px);
            padding: 40px;
            border-radius: 20px;
            box-shadow: 0 15px 35px rgba(0,0,0,0.1), 0 5px 15px rgba(0,0,0,0.05);
            width: 100%;
            max-width: 420px;
            border: 1px solid var(--glass-border);
            animation: fadeInUp 0.8s cubic-bezier(0.2, 0.8, 0.2, 1);
            position: relative;
            overflow: hidden;
            box-sizing: border-box;
            margin: 20px;
        }
        .header-row {
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 12px;
            margin-bottom: 30px;
        }
        h2 { margin: 0; font-weight: 700; color: #111827; letter-spacing: -0.5px; }
        .github-link {
            color: var(--text-main);
            transition: transform 0.3s ease, color 0.3s ease;
            display: flex;
            align-items: center;
        }
        .github-link:hover { transform: scale(1.1) rotate(5deg); color: var(--primary); }

        /* 自定义 Input 样式 */
        .input-group { margin-bottom: 20px; text-align: left; position: relative; }
        .label { font-size: 13px; font-weight: 600; color: var(--text-sub); margin-bottom: 8px; display: block; }
        input {
            width: 100%;
            padding: 14px 16px;
            border: 2px solid #e5e7eb;
            border-radius: 12px;
            box-sizing: border-box;
            font-size: 15px;
            transition: all 0.3s ease;
            background: rgba(255,255,255,0.6);
        }
        input:focus {
            border-color: var(--primary);
            box-shadow: 0 0 0 4px rgba(79, 70, 229, 0.1);
            outline: none;
            background: white;
        }

        /* 自定义 Select 样式 */
        .custom-select {
            position: relative;
            cursor: pointer;
            user-select: none;
        }
        .select-trigger {
            background: rgba(255,255,255,0.6);
            border: 2px solid #e5e7eb;
            border-radius: 12px;
            padding: 14px 16px;
            font-size: 15px;
            color: var(--text-main);
            display: flex;
            justify-content: space-between;
            align-items: center;
            transition: all 0.3s;
        }
        .select-trigger:hover { border-color: #d1d5db; }
        .select-trigger.active { border-color: var(--primary); background: white; }
        .select-arrow { transition: transform 0.3s; width: 10px; height: 10px; border-right: 2px solid #6b7280; border-bottom: 2px solid #6b7280; transform: rotate(45deg) translateY(-2px); }
        .select-trigger.active .select-arrow { transform: rotate(225deg) translateY(-2px); }
        
        .options-container {
            position: absolute;
            top: 110%;
            left: 0;
            right: 0;
            background: white;
            border-radius: 12px;
            box-shadow: 0 10px 25px rgba(0,0,0,0.1);
            opacity: 0;
            visibility: hidden;
            transform: translateY(-10px);
            transition: all 0.3s cubic-bezier(0.2, 0.8, 0.2, 1);
            z-index: 100;
            overflow: hidden;
            border: 1px solid #f3f4f6;
        }
        .options-container.open { opacity: 1; visibility: visible; transform: translateY(0); }
        .option {
            padding: 12px 16px;
            transition: background 0.2s;
            font-size: 14px;
        }
        .option:hover { background: #f3f4f6; color: var(--primary); }
        .option.selected { background: #e0e7ff; color: var(--primary); font-weight: 600; }

        /* 按钮与消息 */
        button {
            width: 100%;
            padding: 14px;
            background-color: var(--primary);
            color: white;
            border: none;
            border-radius: 12px;
            font-size: 16px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.3s ease;
            margin-top: 10px;
            box-shadow: 0 4px 6px rgba(79, 70, 229, 0.2);
        }
        button:hover { background-color: var(--primary-hover); transform: translateY(-2px); box-shadow: 0 6px 12px rgba(79, 70, 229, 0.3); }
        button:active { transform: translateY(0); }
        button:disabled { background-color: #9ca3af; cursor: not-allowed; transform: none; box-shadow: none; }
        
        .message { margin-top: 20px; font-size: 14px; padding: 12px; border-radius: 8px; display: none; line-height: 1.5; text-align: left; animation: fadeInUp 0.3s; }
        .error { background-color: #fee2e2; color: #991b1b; border: 1px solid #fecaca; }
        .success { background-color: #dcfce7; color: #166534; border: 1px solid #bbf7d0; }

        .cf-turnstile { display: flex; justify-content: center; margin: 20px 0; }
        
        .footer {
            margin-top: 30px;
            padding-top: 20px;
            border-top: 1px solid #e5e7eb;
            font-size: 12px;
            color: var(--text-sub);
            display: flex;
            flex-direction: column;
            align-items: center;
            gap: 5px;
        }
        .footer a { color: var(--text-sub); text-decoration: none; display: flex; align-items: center; gap: 6px; transition: color 0.2s; }
        .footer a:hover { color: var(--primary); }
    </style>
</head>
<body>
    <div class="card">
        <div class="header-row">
            <h2>Office 365 自助开通</h2>
            <a href="https://github.com/zixiwangluo/CF-M365-Admin" target="_blank" class="github-link" title="View Source on GitHub">
                ${GITHUB_ICON}
            </a>
        </div>
        
        <form id="regForm">
            <!-- 隐藏的真实 Input，用于存储选择的值 -->
            <input type="hidden" id="skuName" name="skuName" value="">
            
            <div class="input-group">
                <span class="label">选择订阅类型</span>
                <div class="custom-select">
                    <div class="select-trigger" id="selectTrigger">
                        <span>请选择...</span>
                        <div class="select-arrow"></div>
                    </div>
                    <div class="options-container" id="optionsContainer">
                        ${skuOptions.map((opt, index) => `<div class="option" data-value="${opt}">${opt}</div>`).join('')}
                    </div>
                </div>
            </div>

            <div class="input-group">
                <span class="label">用户名 (仅字母和数字)</span>
                <input type="text" id="username" placeholder="例如: admin" required pattern="[a-zA-Z0-9]+">
            </div>
            
            <div class="input-group">
                <span class="label">密码 (8位+, 含大/小写/数字/符号 3种)</span>
                <input type="password" id="password" placeholder="设置您的强密码" required>
            </div>
            
            <div class="cf-turnstile" data-sitekey="${siteKey}"></div>
            
            <button type="submit" id="btn">立即创建账号</button>
        </form>
        
        <div id="msg" class="message"></div>
        
        <div class="footer">
            <div>Powered By CloudFlare Workers</div>
            <a href="https://github.com/zixiwangluo/CF-M365-Admin" target="_blank">
                ${GITHUB_ICON}
                <span>CF-M365-Admin</span>
            </a>
        </div>
    </div>

    <script>
        // --- 自定义下拉框逻辑 ---
        const trigger = document.getElementById('selectTrigger');
        const container = document.getElementById('optionsContainer');
        const options = document.querySelectorAll('.option');
        const hiddenInput = document.getElementById('skuName');
        const triggerText = trigger.querySelector('span');

        // 默认选择第一项
        if (options.length > 0) {
            selectOption(options[0]);
        }

        function toggleSelect(e) {
            e.stopPropagation();
            trigger.classList.toggle('active');
            container.classList.toggle('open');
        }

        function closeSelect() {
            trigger.classList.remove('active');
            container.classList.remove('open');
        }

        function selectOption(option) {
            const val = option.getAttribute('data-value');
            hiddenInput.value = val;
            triggerText.innerText = val;
            options.forEach(o => o.classList.remove('selected'));
            option.classList.add('selected');
            closeSelect();
        }

        trigger.addEventListener('click', toggleSelect);
        
        options.forEach(opt => {
            opt.addEventListener('click', (e) => {
                e.stopPropagation();
                selectOption(opt);
            });
        });

        document.addEventListener('click', closeSelect);

        // --- 密码与表单逻辑 ---
        function checkComplexity(pwd) {
            let score = 0;
            if (/[a-z]/.test(pwd)) score++;
            if (/[A-Z]/.test(pwd)) score++;
            if (/\\d/.test(pwd)) score++;
            if (/[^a-zA-Z0-9]/.test(pwd)) score++;
            return score >= 3;
        }

        document.getElementById('regForm').addEventListener('submit', async (e) => {
            e.preventDefault();
            const btn = document.getElementById('btn');
            const msg = document.getElementById('msg');
            const username = document.getElementById('username').value;
            const password = document.getElementById('password').value;
            const skuVal = hiddenInput.value;

            if (!skuVal) {
                msg.style.display = 'block'; msg.className = 'message error';
                msg.innerText = '⚠️ 请先选择订阅类型';
                return;
            }

            if (password.toLowerCase().includes(username.toLowerCase())) {
                msg.style.display = 'block'; msg.className = 'message error';
                msg.innerText = '❌ 安全警告：密码不能包含用户名，请重设';
                return;
            }

            if (password.length < 8 || !checkComplexity(password)) {
                msg.style.display = 'block'; msg.className = 'message error';
                msg.innerText = '❌ 密码太简单：需8位以上，且包含大写、小写、数字、符号中的至少3种';
                return;
            }

            btn.disabled = true; btn.innerHTML = '<span style="display:inline-block;animation:spin 1s linear infinite">↻</span> 正在部署资源...';
            msg.style.display = 'none';

            const formData = new FormData();
            formData.append('skuName', skuVal);
            formData.append('username', username);
            formData.append('password', password);
            formData.append('cf-turnstile-response', document.querySelector('[name="cf-turnstile-response"]').value);
            
            try {
                const res = await fetch('/', { method: 'POST', body: formData });
                const data = await res.json();
                msg.style.display = 'block';
                if (data.success) {
                    msg.className = 'message success';
                    msg.innerHTML = '🎉 <b>开通成功！</b><br>账号: ' + data.email + '<br>密码: (您刚才设置的)<br><a href="https://portal.office.com" target="_blank" style="color:#15803d;font-weight:bold;margin-top:5px;display:inline-block">👉 前往 Office.com 登录</a>';
                    document.getElementById('regForm').reset();
                    // 重置下拉框到默认
                    if (options.length > 0) selectOption(options[0]);
                    if(typeof turnstile !== 'undefined') turnstile.reset();
                } else {
                    msg.className = 'message error';
                    msg.innerText = '❌ ' + data.message;
                    if(typeof turnstile !== 'undefined') turnstile.reset();
                }
            } catch (err) {
                msg.style.display = 'block'; msg.className = 'message error';
                msg.innerText = '网络连接似乎断了，请稍后重试';
            } finally {
                btn.disabled = false; btn.innerText = '立即创建账号';
            }
        });
        
        // 添加简单的旋转动画样式
        const styleSheet = document.createElement("style");
        styleSheet.innerText = "@keyframes spin { 100% { transform: rotate(360deg); } }";
        document.head.appendChild(styleSheet);
    </script>
</body>
</html>
`;

// --- 2. 后台管理页面 HTML (美化版) ---
const HTML_ADMIN_PAGE = (skuMapJson) => `
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Office 365 用户管理控制台</title>
    <style>
        :root { --admin-primary: #0078d4; --admin-bg: #f3f2f1; }
        body { font-family: "Segoe UI", -apple-system, sans-serif; padding: 0; margin: 0; background: var(--admin-bg); color: #201f1e; }
        .nav { background: white; padding: 15px 30px; box-shadow: 0 1px 4px rgba(0,0,0,0.1); display: flex; justify-content: space-between; align-items: center;}
        .nav h1 { margin: 0; font-size: 20px; font-weight: 600; display: flex; align-items: center; gap: 10px; }
        .container { max-width: 1400px; margin: 30px auto; padding: 0 20px; }
        
        .card { background: white; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); padding: 25px; margin-bottom: 20px; animation: fadeIn 0.5s; }
        @keyframes fadeIn { from { opacity: 0; transform: translateY(10px); } to { opacity: 1; transform: translateY(0); } }

        .toolbar { display: flex; gap: 12px; margin-bottom: 20px; flex-wrap: wrap; }
        button { padding: 10px 20px; border: none; border-radius: 6px; cursor: pointer; font-weight: 600; font-size: 14px; transition: all 0.2s; display: flex; align-items: center; gap: 6px;}
        button:hover { opacity: 0.9; transform: translateY(-1px); }
        button:active { transform: translateY(0); }
        
        .btn-refresh { background: white; color: var(--admin-primary); border: 1px solid var(--admin-primary); }
        .btn-refresh:hover { background: #f0f8ff; }
        .btn-del { background: #d13438; color: white; box-shadow: 0 2px 5px rgba(209, 52, 56, 0.3); }
        .btn-pwd { background: #ffaa44; color: white; box-shadow: 0 2px 5px rgba(255, 170, 68, 0.3); }
        .btn-lic { background: #107c10; color: white; box-shadow: 0 2px 5px rgba(16, 124, 16, 0.3); }

        table { width: 100%; border-collapse: separate; border-spacing: 0; margin-top: 10px; font-size: 14px; }
        th { background: #f8f9fa; padding: 15px; text-align: left; border-bottom: 2px solid #eee; color: #605e5c; font-weight: 600; cursor: pointer; user-select: none; }
        td { padding: 15px; border-bottom: 1px solid #f3f2f1; vertical-align: middle; transition: background 0.2s; }
        tr:hover td { background: #faf9f8; }
        tr:last-child td { border-bottom: none; }
        
        .tag { padding: 4px 10px; border-radius: 12px; font-size: 12px; font-weight: 500; display:inline-block; margin:2px;}
        .tag-blue { background: #e0efff; color: #005a9e; }
        
        .modal { display: none; position: fixed; top:0; left:0; width:100%; height:100%; background: rgba(0,0,0,0.4); backdrop-filter: blur(4px); justify-content: center; align-items: center; z-index: 1000; animation: fadeIn 0.2s;}
        .modal-content { background: white; padding: 30px; border-radius: 12px; width: 600px; max-width: 90%; max-height: 85vh; overflow-y: auto; box-shadow: 0 10px 30px rgba(0,0,0,0.2); transform: scale(0.95); animation: popIn 0.3s forwards;}
        @keyframes popIn { to { transform: scale(1); } }
        
        .close { float: right; cursor: pointer; font-size: 24px; color: #605e5c; transition: color 0.2s; }
        .close:hover { color: #000; }
        
        .loading { text-align: center; padding: 40px; color: #605e5c; }
        input[type="checkbox"] { width: 18px; height: 18px; cursor: pointer; }
        
        /* 许可证条 */
        .progress-bar { background: #edebe9; border-radius: 4px; height: 10px; width: 120px; display: inline-block; overflow: hidden; vertical-align: middle; margin-left: 10px;}
        .progress-fill { height: 100%; background: var(--admin-primary); transition: width 0.5s ease; }

        .footer { text-align: center; padding: 20px; color: #888; font-size: 12px; display: flex; justify-content: center; align-items: center; gap: 8px;}
        .footer a { color: #888; text-decoration: none; display: flex; align-items: center; gap: 5px; }
        .footer a:hover { color: var(--admin-primary); }
    </style>
</head>
<body>
    <div class="nav">
        <h1><span>⚡</span> Office 365 Admin</h1>
        <div style="font-size:12px; color:#666;">安全模式: Enabled</div>
    </div>
    
    <div class="container">
        <div class="card">
            <div class="toolbar">
                <button class="btn-refresh" onclick="loadUsers()">🔄 刷新列表</button>
                <button class="btn-lic" onclick="openLicModal()">📊 查看订阅余量</button> 
                <button class="btn-pwd" onclick="openPwdModal()">🔑 重置密码</button>
                <button class="btn-del" onclick="bulkDelete()">🗑️ 批量删除</button>
            </div>
            <div id="status" style="margin-bottom:10px; height:20px; color:#107c10; font-weight:600;"></div>
            
            <div style="overflow-x: auto;">
                <table id="mainTable">
                    <thead>
                        <tr>
                            <th width="40"><input type="checkbox" id="selectAll" onclick="toggleAll(this)"></th>
                            <th onclick="sortTable('displayName')">用户名 <span id="sort-displayName"></span></th>
                            <th onclick="sortTable('userPrincipalName')">账号(邮箱) <span id="sort-userPrincipalName"></span></th>
                            <th>当前订阅</th>
                            <th onclick="sortTable('createdDateTime')">创建时间 <span id="sort-createdDateTime"></span></th>
                            <th>UUID</th>
                        </tr>
                    </thead>
                    <tbody id="userTableBody"></tbody>
                </table>
            </div>
            <div class="loading" id="loading">正在从 Microsoft Graph 加载数据...</div>
        </div>
        
        <div class="footer">
             Powered by CloudFlare Workers
             <span>|</span>
             <a href="https://github.com/zixiwangluo/CF-M365-Admin" target="_blank">
                ${GITHUB_ICON} CF-M365-Admin
             </a>
        </div>
    </div>

    <!-- 密码模态框 -->
    <div id="pwdModal" class="modal">
        <div class="modal-content" style="width: 400px;">
            <span class="close" onclick="closeModal('pwdModal')">&times;</span>
            <h3 style="margin-top:0">重置密码</h3>
            <div style="margin: 20px 0;">
                <label style="display:block; margin-bottom:10px; cursor:pointer;">
                    <input type="radio" name="pwdType" value="auto" checked onclick="togglePwdInput(false)"> 
                    🎲 自动生成高强度密码
                </label>
                <label style="display:block; margin-bottom:10px; cursor:pointer;">
                    <input type="radio" name="pwdType" value="custom" onclick="togglePwdInput(true)"> 
                    ✏️ 自定义密码
                </label>
            </div>
            <input type="text" id="customPwd" placeholder="输入新密码" style="width:100%; padding:10px; border:1px solid #ccc; border-radius:4px; display:none; box-sizing:border-box;">
            <div style="margin-top:25px; text-align:right;">
                <button class="btn-pwd" onclick="submitPwdReset()" style="width:100%; justify-content:center;">确认重置</button>
            </div>
            <div id="pwdResult" style="margin-top:15px; font-size:13px; color:#005a9e; word-break:break-all; background:#f0f8ff; padding:10px; border-radius:4px; display:none;"></div>
        </div>
    </div>

    <!-- 许可证模态框 -->
    <div id="licModal" class="modal">
        <div class="modal-content">
            <span class="close" onclick="closeModal('licModal')">&times;</span>
            <h3 style="margin-top:0">全局订阅使用情况</h3>
            <div id="licLoading" style="display:none; text-align:center; padding:20px;">正在查询...</div>
            <table id="licTable">
                <thead>
                    <tr>
                        <th>订阅名称 / SKU ID</th>
                        <th>总量</th>
                        <th>已分配</th>
                        <th>剩余可用</th>
                    </tr>
                </thead>
                <tbody id="licBody"></tbody>
            </table>
        </div>
    </div>

    <script>
        const RAW_MAP = ${skuMapJson || '{}'};
        const ID_TO_NAME = {};
        for(let key in RAW_MAP) ID_TO_NAME[RAW_MAP[key]] = key;

        let allUsers = [];
        const API_BASE = '/admin/api';

        // --- 用户列表逻辑 ---
        async function loadUsers() {
            document.getElementById('loading').style.display = 'block';
            document.getElementById('userTableBody').innerHTML = '';
            resetSortIcons();
            try {
                const res = await fetch(API_BASE + '/users?token=' + getToken());
                const data = await res.json();
                if (!res.ok || data.error) throw new Error(data.error ? data.error.message : 'API Error');
                if (!Array.isArray(data.value)) throw new Error('API 响应格式错误');
                allUsers = data.value;
                renderTable(allUsers);
            } catch (e) { alert('加载失败: ' + e.message); } 
            finally { document.getElementById('loading').style.display = 'none'; }
        }

        function renderTable(users) {
            const tbody = document.getElementById('userTableBody');
            if(users.length === 0) {
                tbody.innerHTML = '<tr><td colspan="6" style="text-align:center; padding:20px;">暂无用户</td></tr>';
                return;
            }
            tbody.innerHTML = users.map(u => {
                let licenses = '<span style="color:#999">无订阅</span>';
                if (u.assignedLicenses && u.assignedLicenses.length > 0) {
                    licenses = u.assignedLicenses.map(l => {
                        const name = ID_TO_NAME[l.skuId] || l.skuId;
                        return \`<span class="tag tag-blue">\${name}</span>\`;
                    }).join('');
                }
                return \`
                <tr>
                    <td><input type="checkbox" class="u-check" value="\${u.id}" data-name="\${u.userPrincipalName}"></td>
                    <td><strong>\${u.displayName}</strong></td>
                    <td>\${u.userPrincipalName}</td>
                    <td>\${licenses}</td>
                    <td>\${new Date(u.createdDateTime).toLocaleString()}</td>
                    <td style="font-size:11px; color:#aaa; font-family:monospace;">\${u.id}</td>
                </tr>\`;
            }).join('');
        }

        let sortConfig = { key: null, dir: 1 };
        function sortTable(key) {
            if (sortConfig.key === key) sortConfig.dir *= -1;
            else sortConfig = { key: key, dir: 1 };

            resetSortIcons();
            document.getElementById('sort-' + key).innerText = sortConfig.dir === 1 ? '↓' : '↑';

            allUsers.sort((a, b) => {
                let valA = a[key] || '';
                let valB = b[key] || '';
                if (typeof valA === 'string') return sortConfig.dir * valA.localeCompare(valB, 'zh-CN'); 
                return valA > valB ? sortConfig.dir : -sortConfig.dir;
            });
            renderTable(allUsers);
        }

        function resetSortIcons() {
            ['displayName', 'userPrincipalName', 'createdDateTime'].forEach(k => {
                const el = document.getElementById('sort-' + k);
                if(el) el.innerText = ''; 
            });
        }

        function getToken() { return new URLSearchParams(window.location.search).get('token') || prompt('请输入管理员 Token:'); }
        function toggleAll(source) { document.querySelectorAll('.u-check').forEach(c => c.checked = source.checked); }
        function getSelected() { return Array.from(document.querySelectorAll('.u-check:checked')).map(c => ({id: c.value, name: c.getAttribute('data-name')})); }

        async function bulkDelete() {
            const selected = getSelected();
            if (selected.length === 0) return alert('请先选择用户');
            if (!confirm(\`⚠️ 高危操作确认\n\n确定要永久删除选中的 \${selected.length} 个用户吗？\n此操作不可恢复！\`)) return;
            
            const statusDiv = document.getElementById('status');
            statusDiv.innerText = '正在执行批量删除...';
            
            for (const u of selected) {
                try {
                    const res = await fetch(API_BASE + '/users/' + u.id + '?token=' + getToken(), { method: 'DELETE' });
                    if(res.status === 403) console.error(u.name + ' 删除失败: 受保护的账户');
                } catch(e) {}
            }
            statusDiv.innerText = '✅ 删除操作完成';
            setTimeout(() => statusDiv.innerText = '', 3000);
            loadUsers();
        }

        // --- 模态框通用逻辑 ---
        function closeModal(id) { document.getElementById(id).style.display = 'none'; }
        
        // --- 密码相关 ---
        function openPwdModal() { 
            if (getSelected().length === 0) return alert('请先选择用户'); 
            document.getElementById('pwdModal').style.display = 'flex'; 
            document.getElementById('pwdResult').style.display = 'none';
            document.getElementById('pwdResult').innerText = ''; 
        }
        function togglePwdInput(show) { document.getElementById('customPwd').style.display = show ? 'block' : 'none'; }
        
        function generatePass() {
            const chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*";
            let pass = ""; 
            for (let i = 0; i < 12; i++) pass += chars.charAt(Math.floor(Math.random() * chars.length));
            return pass + "Aa1!"; 
        }

        async function submitPwdReset() {
            const selected = getSelected();
            const type = document.querySelector('input[name="pwdType"]:checked').value;
            let password = (type === 'custom') ? document.getElementById('customPwd').value : '';
            if (type === 'custom' && !password) return alert('请输入密码');
            
            const resultDiv = document.getElementById('pwdResult');
            resultDiv.style.display = 'block';
            resultDiv.innerText = '正在请求 Graph API...';
            
            let successList = [];
            for (const u of selected) {
                const finalPwd = (type === 'auto') ? generatePass() : password;
                try {
                    await fetch(API_BASE + '/users/' + u.id + '/password?token=' + getToken(), {
                        method: 'PATCH',
                        headers: {'Content-Type': 'application/json'},
                        body: JSON.stringify({ password: finalPwd })
                    });
                    successList.push(\`\${u.name} -> \${finalPwd}\`);
                } catch(e) {}
            }
            resultDiv.innerHTML = '<b>✅ 操作完成，请复制保存:</b><br><br>' + successList.join('<br>');
        }

        // --- 许可证查询相关 ---
        async function openLicModal() {
            document.getElementById('licModal').style.display = 'flex';
            document.getElementById('licBody').innerHTML = '';
            document.getElementById('licLoading').style.display = 'block';

            try {
                const res = await fetch(API_BASE + '/licenses?token=' + getToken());
                const data = await res.json();
                
                if(!res.ok) throw new Error(data.error || 'Fetch Error');
                
                document.getElementById('licBody').innerHTML = data.map(lic => {
                    const friendlyName = ID_TO_NAME[lic.skuId] ? 
                        \`<span style="font-weight:bold; color:#0078d4">\${ID_TO_NAME[lic.skuId]}</span>\` : 
                        lic.skuPartNumber;
                        
                    const usagePercent = lic.total > 0 ? Math.round((lic.used / lic.total) * 100) : 0;
                    const remaining = lic.total - lic.used;
                    
                    return \`
                    <tr>
                        <td>
                            \${friendlyName}
                            <div style="font-size:11px; color:#999; margin-top:2px;">\${lic.skuId}</div>
                        </td>
                        <td>\${lic.total}</td>
                        <td>
                            \${lic.used}
                            <div class="progress-bar" title="\${usagePercent}%">
                                <div class="progress-fill" style="width: \${usagePercent}%"></div>
                            </div>
                        </td>
                        <td style="color: \${remaining < 5 ? '#d13438' : '#107c10'}; font-weight:bold;">
                            \${remaining}
                        </td>
                    </tr>
                    \`;
                }).join('');

            } catch(e) {
                document.getElementById('licBody').innerHTML = \`<tr><td colspan="4" style="color:red; text-align:center;">查询失败: \${e.message}</td></tr>\`;
            } finally {
                document.getElementById('licLoading').style.display = 'none';
            }
        }

        if(window.location.search.includes('token=')) loadUsers();
    </script>
</body>
</html>
`;

export default {
    async fetch(request, env) {
        const url = new URL(request.url);

        // --- A. 后台管理路由 (/admin) ---
        if (url.pathname.startsWith('/admin')) {
            const token = url.searchParams.get('token');
            if (token !== env.ADMIN_TOKEN) {
                return new Response('401 Unauthorized', { status: 401 });
            }

            if (url.pathname === '/admin' || url.pathname === '/admin/') {
                return new Response(HTML_ADMIN_PAGE(env.SKU_MAP), { headers: { 'Content-Type': 'text/html;charset=UTF-8' } });
            }

            // API: 获取许可证列表 (新增)
            if (url.pathname === '/admin/api/licenses' && request.method === 'GET') {
                const accessToken = await getAccessToken(env);
                const resp = await fetch('https://graph.microsoft.com/v1.0/subscribedSkus', {
                    headers: { 'Authorization': `Bearer ${accessToken}` }
                });
                const data = await resp.json();
                
                // 格式化数据返回给前端
                const result = data.value ? data.value.map(s => ({
                    skuPartNumber: s.skuPartNumber, 
                    skuId: s.skuId,                 
                    total: s.prepaidUnits.enabled,  
                    used: s.consumedUnits           
                })) : [];
                
                return new Response(JSON.stringify(result), { status: resp.status, headers: { 'Content-Type': 'application/json' } });
            }

            // API: 获取用户列表
            if (url.pathname === '/admin/api/users' && request.method === 'GET') {
                const accessToken = await getAccessToken(env);
                const graphUrl = 'https://graph.microsoft.com/v1.0/users?$select=id,displayName,userPrincipalName,createdDateTime,assignedLicenses&$top=100&$orderby=createdDateTime desc&$count=true';
                debugLog(env, 'Fetching users from:', graphUrl);

                const resp = await fetch(graphUrl, { 
                    headers: { 'Authorization': `Bearer ${accessToken}`, 'ConsistencyLevel': 'eventual' } 
                });
                
                const data = await resp.json();
                
                // 隐藏账户过滤
                if (data.value && env.HIDDEN_USER) {
                    data.value = data.value.filter(u => u.userPrincipalName.toLowerCase() !== env.HIDDEN_USER.toLowerCase());
                }
                
                return new Response(JSON.stringify(data), { status: resp.status, headers: { 'Content-Type': 'application/json' } });
            }

            // API: 删除用户
            if (url.pathname.match(/\/admin\/api\/users\/[^/]+$/) && request.method === 'DELETE') {
                const userId = url.pathname.split('/').pop();
                const accessToken = await getAccessToken(env);

                if (env.HIDDEN_USER) {
                    const checkResp = await fetch(`https://graph.microsoft.com/v1.0/users/${userId}?$select=userPrincipalName`, {
                        headers: { 'Authorization': `Bearer ${accessToken}` }
                    });
                    if (checkResp.ok) {
                        const checkUser = await checkResp.json();
                        if (checkUser.userPrincipalName.toLowerCase() === env.HIDDEN_USER.toLowerCase()) {
                            return new Response(JSON.stringify({error: 'Forbidden'}), { status: 403 });
                        }
                    }
                }
                const resp = await fetch(`https://graph.microsoft.com/v1.0/users/${userId}`, {
                    method: 'DELETE',
                    headers: { 'Authorization': `Bearer ${accessToken}` }
                });
                return new Response(null, { status: resp.status });
            }

            // API: 重置密码
            if (url.pathname.endsWith('/password') && request.method === 'PATCH') {
                const userId = url.pathname.split('/')[4]; 
                const body = await request.json();
                const accessToken = await getAccessToken(env);
                
                if (env.HIDDEN_USER) {
                     const checkResp = await fetch(`https://graph.microsoft.com/v1.0/users/${userId}?$select=userPrincipalName`, {
                        headers: { 'Authorization': `Bearer ${accessToken}` }
                    });
                    if (checkResp.ok && (await checkResp.json()).userPrincipalName.toLowerCase() === env.HIDDEN_USER.toLowerCase()) {
                        return new Response(JSON.stringify({error: 'Forbidden'}), { status: 403 });
                    }
                }

                const payload = { passwordProfile: { forceChangePasswordNextSignIn: false, password: body.password } };
                const resp = await fetch(`https://graph.microsoft.com/v1.0/users/${userId}`, {
                    method: 'PATCH',
                    headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
                    body: JSON.stringify(payload)
                });
                return new Response(null, { status: resp.status });
            }
        }

        // --- B. 前台注册路由 ---
        if (request.method === 'GET') {
            let skuOptions = [];
            try {
                const map = JSON.parse(env.SKU_MAP || '{}');
                skuOptions = Object.keys(map);
            } catch (e) { return new Response('Config Error: SKU_MAP is invalid', {status:500}); }
            
            return new Response(HTML_REGISTER_PAGE(env.TURNSTILE_SITE_KEY, skuOptions), {
                headers: { 'Content-Type': 'text/html;charset=UTF-8' },
            });
        }

        if (request.method === 'POST') {
            try {
                const formData = await request.formData();
                const username = formData.get('username');
                const password = formData.get('password');
                const skuName = formData.get('skuName'); 
                const turnstileToken = formData.get('cf-turnstile-response');
                const ip = request.headers.get('CF-Connecting-IP');

                debugLog(env, 'Register attempt:', username, ip);

                const verifyRes = await fetch('https://challenges.cloudflare.com/turnstile/v0/siteverify', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ secret: env.CF_TURNSTILE_SECRET, response: turnstileToken, remoteip: ip })
                });
                if (!(await verifyRes.json()).success) return Response.json({ success: false, message: '人机验证失败，请刷新页面重试' });

                let skuId = null;
                try { skuId = JSON.parse(env.SKU_MAP || '{}')[skuName]; } catch(e){}
                if (!skuId) return Response.json({ success: false, message: '请选择有效的订阅类型' });

                if (!/^[a-zA-Z0-9]+$/.test(username)) return Response.json({ success: false, message: '用户名格式错误，仅允许字母和数字' });
                
                // 后端二次校验
                if (password.toLowerCase().includes(username.toLowerCase())) {
                    return Response.json({ success: false, message: '密码不能包含用户名（或用户名的部分）' });
                }
                if (!checkPasswordComplexity(password)) {
                    return Response.json({ success: false, message: '密码需包含大小写/数字/符号中的至少3种' });
                }

                const accessToken = await getAccessToken(env);
                const userEmail = `${username}@${env.DEFAULT_DOMAIN}`;

                if (env.HIDDEN_USER && userEmail.toLowerCase() === env.HIDDEN_USER.toLowerCase()) {
                     return Response.json({ success: false, message: '该用户名已被占用' });
                }

                const userPayload = {
                    accountEnabled: true,
                    displayName: username,
                    mailNickname: username,
                    userPrincipalName: userEmail,
                    passwordProfile: { forceChangePasswordNextSignIn: false, password: password },
                    usageLocation: "CN" 
                };
                
                const createReq = await fetch('https://graph.microsoft.com/v1.0/users', {
                    method: 'POST',
                    headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
                    body: JSON.stringify(userPayload)
                });

                if (!createReq.ok) {
                    const err = await createReq.json();
                    debugLog(env, 'Create User Error:', err);
                    
                    const errMsg = err.error?.message || '';
                    if (errMsg.includes('another object')) return Response.json({ success: false, message: '该用户名已被占用，请换一个试试' });
                    if (errMsg.includes('Password cannot contain username')) return Response.json({ success: false, message: '创建失败：密码不能包含用户名' });
                    if (errMsg.includes('PasswordProfile') || errMsg.includes('weak')) return Response.json({ success: false, message: '创建失败：密码过于简单或不符合策略' });

                    throw new Error(errMsg);
                }

                const newUser = await createReq.json();

                const licenseReq = await fetch(`https://graph.microsoft.com/v1.0/users/${newUser.id}/assignLicense`, {
                    method: 'POST',
                    headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        addLicenses: [{ disabledPlans: [], skuId: skuId }],
                        removeLicenses: []
                    })
                });

                if (!licenseReq.ok) {
                    const licErr = await licenseReq.json();
                    return Response.json({ success: false, message: '账号创建成功但订阅分配失败: ' + licErr.error.message });
                }

                return Response.json({ success: true, email: userEmail });
            } catch (e) {
                return Response.json({ success: false, message: '系统繁忙或错误: ' + e.message });
            }
        }

        return new Response('Method Not Allowed', { status: 405 });
    }
};

async function getAccessToken(env) {
    const params = new URLSearchParams();
    params.append('client_id', env.AZURE_CLIENT_ID);
    params.append('scope', 'https://graph.microsoft.com/.default');
    params.append('client_secret', env.AZURE_CLIENT_SECRET);
    params.append('grant_type', 'client_credentials');

    const res = await fetch(`https://login.microsoftonline.com/${env.AZURE_TENANT_ID}/oauth2/v2.0/token`, {
        method: 'POST',
        body: params
    });
    const data = await res.json();
    return data.access_token;
}
