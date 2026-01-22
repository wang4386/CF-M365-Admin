/**
 * CF-M365-Admin v4.0 (Zen-iOS Design & Secure Login)
 * * [环境变量配置]
 * AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET: 微软 Graph API 配置
 * CF_TURNSTILE_SECRET: Turnstile 后端密钥 (必须)
 * TURNSTILE_SITE_KEY: Turnstile 前端 Site Key (必须)
 * DEFAULT_DOMAIN: 邮箱后缀 (不带@)
 * SKU_MAP: JSON SKU 映射
 * ADMIN_TOKEN: 管理员密码
 * HIDDEN_USER: (可选) 隐藏管理员账号
 * ADMIN_PATH: (可选) 后台路径，默认 "/admin"
 * ENABLE_DEBUG: (可选) "true"
 * * [KV 绑定]
 * Variable Name: INVITE_CODES
 */

const debugLog = (env, ...args) => {
    if (env.ENABLE_DEBUG === 'true') console.log('[DEBUG]', ...args);
};

// --- 辅助函数 ---
const getEnv = (val, defaultVal = '') => val ? val.trim() : defaultVal;

function checkPasswordComplexity(pwd) {
    if (!pwd || pwd.length < 8) return false;
    let score = 0;
    if (/[a-z]/.test(pwd)) score++;
    if (/[A-Z]/.test(pwd)) score++;
    if (/\d/.test(pwd)) score++;
    if (/[^a-zA-Z0-9]/.test(pwd)) score++;
    return score >= 3;
}

// --- 通用 HTML 头部 (Tailwind + Fonts) ---
const HTML_HEAD = (title) => `
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
    <title>${title}</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <script src="https://challenges.cloudflare.com/turnstile/v0/api.js" async defer></script>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;800&display=swap" rel="stylesheet">
    <script src="https://unpkg.com/lucide@latest"></script>
    <script>
        tailwind.config = {
            theme: {
                extend: {
                    fontFamily: { sans: ['Inter', 'sans-serif'] },
                    colors: {
                        ios: {
                            bg: '#F2F2F7',
                            card: 'rgba(255, 255, 255, 0.65)',
                            primary: '#1C1C1E',
                            secondary: '#FFFFFF',
                            danger: '#FF3B30',
                            success: '#34C759',
                            border: 'rgba(255, 255, 255, 0.6)',
                            borderOut: 'rgba(0, 0, 0, 0.05)'
                        }
                    },
                    boxShadow: {
                        'zen': '0 24px 48px -12px rgba(0, 0, 0, 0.08), 0 4px 12px -4px rgba(0,0,0,0.02)',
                        'inner-light': 'inset 0 2px 4px 0 rgba(0, 0, 0, 0.03)'
                    },
                    backdropBlur: {
                        'xs': '2px',
                    }
                }
            }
        }
    </script>
    <style>
        body { background-color: #F2F2F7; -webkit-font-smoothing: antialiased; }
        .glass-panel {
            background-color: rgba(255, 255, 255, 0.65);
            backdrop-filter: blur(40px);
            -webkit-backdrop-filter: blur(40px);
            border: 1px solid rgba(255, 255, 255, 0.6);
            box-shadow: 0 0 0 1px rgba(0,0,0,0.03), 0 24px 48px -12px rgba(0, 0, 0, 0.08);
        }
        .input-zen {
            background-color: rgba(243, 244, 246, 0.6);
            box-shadow: inset 0 2px 4px 0 rgba(0, 0, 0, 0.03);
            border: 1px solid transparent;
            transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1);
        }
        .input-zen:focus {
            background-color: #FFFFFF;
            border-color: rgba(0,0,0,0.1);
            box-shadow: 0 4px 12px rgba(0,0,0,0.05);
            outline: none;
        }
        .btn-primary {
            background-color: #1C1C1E;
            color: white;
            transition: transform 0.1s;
        }
        .btn-primary:active { transform: scale(0.98); }
        .btn-secondary {
            background-color: #FFFFFF;
            color: #1C1C1E;
            border: 1px solid rgba(0,0,0,0.05);
            box-shadow: 0 2px 8px rgba(0,0,0,0.04);
            transition: transform 0.1s, background-color 0.2s;
        }
        .btn-secondary:active { transform: scale(0.98); }
        .btn-secondary:hover { background-color: #F9FAFB; }
        
        /* 隐藏滚动条但保留功能 */
        .no-scrollbar::-webkit-scrollbar { display: none; }
        .no-scrollbar { -ms-overflow-style: none; scrollbar-width: none; }
        
        .loading-spinner {
            border: 3px solid rgba(0,0,0,0.1);
            border-left-color: #1C1C1E;
            border-radius: 50%;
            width: 24px; height: 24px;
            animation: spin 1s linear infinite;
        }
        @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
    </style>
</head>
`;

// --- 1. 登录页面 (Login) ---
const HTML_LOGIN_PAGE = (siteKey, adminPath) => `
<!DOCTYPE html>
<html lang="zh-CN">
${HTML_HEAD('Admin Login')}
<body class="flex items-center justify-center min-h-screen p-6">
    <div class="glass-panel w-full max-w-[380px] rounded-[40px] p-10 flex flex-col items-center">
        <div class="mb-8 p-4 bg-white/50 rounded-[24px] shadow-sm">
            <i data-lucide="shield-check" class="w-8 h-8 text-gray-800"></i>
        </div>
        
        <h2 class="text-2xl font-extrabold text-gray-900 mb-2 tracking-tight">管理员登录</h2>
        <p class="text-gray-500 text-sm mb-8 font-medium">请输入访问令牌以继续</p>

        <form id="loginForm" class="w-full space-y-5">
            <div class="space-y-2">
                <label class="text-[10px] font-bold text-gray-400 uppercase tracking-widest pl-1">Access Token</label>
                <input type="password" id="token" 
                    class="input-zen w-full px-4 py-3.5 rounded-2xl text-gray-800 text-sm placeholder-gray-400" 
                    placeholder="••••••••••••" required>
            </div>

            <div class="flex justify-center pt-2">
                <div class="cf-turnstile" data-sitekey="${siteKey}" data-theme="light"></div>
            </div>

            <button type="submit" id="loginBtn" class="btn-primary w-full py-3.5 rounded-2xl font-semibold text-sm shadow-lg mt-4 flex justify-center items-center gap-2">
                <span>进入控制台</span>
                <i data-lucide="arrow-right" class="w-4 h-4"></i>
            </button>
        </form>
        <div id="msg" class="mt-4 text-center text-xs font-medium text-red-500 hidden"></div>
    </div>
    <script>
        lucide.createIcons();
        document.getElementById('loginForm').addEventListener('submit', async (e) => {
            e.preventDefault();
            const btn = document.getElementById('loginBtn');
            const msg = document.getElementById('msg');
            const token = document.getElementById('token').value;
            const turnstileEl = document.querySelector('[name="cf-turnstile-response"]');
            
            if (!turnstileEl || !turnstileEl.value) {
                msg.textContent = '请完成人机验证';
                msg.classList.remove('hidden');
                return;
            }

            btn.disabled = true;
            btn.innerHTML = '<div class="loading-spinner w-4 h-4 border-2"></div>';
            msg.classList.add('hidden');

            try {
                const res = await fetch('${adminPath}/login', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ token, cf_turnstile: turnstileEl.value })
                });
                
                if (res.ok) {
                    window.location.reload(); // 登录成功，刷新页面进入后台
                } else {
                    const data = await res.json();
                    throw new Error(data.message || '验证失败');
                }
            } catch (err) {
                msg.textContent = err.message;
                msg.classList.remove('hidden');
                btn.disabled = false;
                btn.innerHTML = '<span>重试</span><i data-lucide="refresh-cw" class="w-4 h-4"></i>';
                turnstile.reset();
            }
        });
    </script>
</body>
</html>
`;

// --- 2. 注册页面 (Register) ---
const HTML_REGISTER_PAGE = (siteKey, skuOptions) => `
<!DOCTYPE html>
<html lang="zh-CN">
${HTML_HEAD('Office 365 邀请注册')}
<body class="flex items-center justify-center min-h-screen p-6">
    <div class="glass-panel w-full max-w-[420px] rounded-[48px] p-8 md:p-12 relative overflow-hidden">
        <!-- 装饰背景 -->
        <div class="absolute -top-20 -right-20 w-64 h-64 bg-blue-400/10 rounded-full blur-3xl pointer-events-none"></div>
        <div class="absolute -bottom-20 -left-20 w-64 h-64 bg-purple-400/10 rounded-full blur-3xl pointer-events-none"></div>

        <div class="relative z-10">
            <div class="text-center mb-10">
                <div class="inline-flex items-center justify-center w-16 h-16 bg-white/80 rounded-[24px] shadow-sm mb-6">
                    <i data-lucide="user-plus" class="w-8 h-8 text-gray-800"></i>
                </div>
                <h2 class="text-3xl font-extrabold text-gray-900 tracking-tight">Create Account</h2>
                <p class="text-gray-500 text-sm mt-2 font-medium">使用管理员分发的邀请码激活</p>
            </div>

            <form id="regForm" class="space-y-6">
                <div class="space-y-2">
                    <label class="text-[10px] font-bold text-gray-400 uppercase tracking-widest pl-1">Invite Code</label>
                    <div class="relative">
                        <input type="text" id="inviteCode" class="input-zen w-full pl-11 pr-4 py-3.5 rounded-2xl text-gray-800 text-sm font-mono tracking-wide" placeholder="XXXX-XXXX" required>
                        <i data-lucide="key" class="w-4 h-4 text-gray-400 absolute left-4 top-1/2 -translate-y-1/2"></i>
                    </div>
                </div>

                <div class="space-y-2">
                    <label class="text-[10px] font-bold text-gray-400 uppercase tracking-widest pl-1">Subscription</label>
                    <div class="relative">
                        <select id="skuSelect" class="input-zen w-full pl-11 pr-10 py-3.5 rounded-2xl text-gray-800 text-sm appearance-none bg-transparent">
                            ${skuOptions.map(opt => `<option value="${opt}">${opt}</option>`).join('')}
                        </select>
                        <i data-lucide="package" class="w-4 h-4 text-gray-400 absolute left-4 top-1/2 -translate-y-1/2"></i>
                        <i data-lucide="chevron-down" class="w-4 h-4 text-gray-400 absolute right-4 top-1/2 -translate-y-1/2"></i>
                    </div>
                </div>

                <div class="grid grid-cols-2 gap-4">
                    <div class="space-y-2">
                        <label class="text-[10px] font-bold text-gray-400 uppercase tracking-widest pl-1">Username</label>
                        <input type="text" id="username" class="input-zen w-full px-4 py-3.5 rounded-2xl text-gray-800 text-sm" placeholder="Alice" pattern="[a-zA-Z0-9]+" required>
                    </div>
                    <div class="space-y-2">
                         <label class="text-[10px] font-bold text-gray-400 uppercase tracking-widest pl-1">Password</label>
                        <input type="password" id="password" class="input-zen w-full px-4 py-3.5 rounded-2xl text-gray-800 text-sm" placeholder="********" required>
                    </div>
                </div>

                <div class="flex justify-center pt-2">
                    <div class="cf-turnstile" data-sitekey="${siteKey}"></div>
                </div>

                <button type="submit" id="btn" class="btn-primary w-full py-4 rounded-2xl font-bold text-sm shadow-xl hover:shadow-2xl transition-all">
                    立即激活账号
                </button>
            </form>

            <div id="msg" class="mt-6 p-4 rounded-2xl text-sm font-medium hidden"></div>
        </div>
    </div>

    <script>
        lucide.createIcons();
        document.getElementById('regForm').addEventListener('submit', async (e) => {
            e.preventDefault();
            const btn = document.getElementById('btn');
            const msg = document.getElementById('msg');
            
            // UI Reset
            msg.classList.add('hidden');
            msg.className = 'mt-6 p-4 rounded-2xl text-sm font-medium hidden';
            btn.disabled = true; 
            const originalBtnText = btn.innerHTML;
            btn.innerHTML = '<div class="flex items-center justify-center gap-2"><span>Processing</span><div class="loading-spinner w-4 h-4 border-white/30 border-l-white"></div></div>';

            const formData = new FormData();
            formData.append('inviteCode', document.getElementById('inviteCode').value.trim());
            formData.append('skuName', document.getElementById('skuSelect').value);
            formData.append('username', document.getElementById('username').value);
            formData.append('password', document.getElementById('password').value);
            
            const turnstileEl = document.querySelector('[name="cf-turnstile-response"]');
            if(turnstileEl) formData.append('cf-turnstile-response', turnstileEl.value);

            try {
                const res = await fetch('/', { method: 'POST', body: formData });
                const data = await res.json();
                
                msg.classList.remove('hidden');
                if (data.success) {
                    msg.classList.add('bg-green-100/80', 'text-green-800', 'border', 'border-green-200');
                    msg.innerHTML = \`<div class="flex flex-col gap-1"><span class="font-bold flex items-center gap-2"><i data-lucide="check-circle" class="w-4 h-4"></i> Success</span><span>账号: \${data.email}</span><a href="https://portal.office.com" target="_blank" class="underline mt-1 opacity-70 hover:opacity-100">前往登录 &rarr;</a></div>\`;
                    lucide.createIcons();
                    document.getElementById('regForm').reset();
                } else {
                    throw new Error(data.message);
                }
            } catch (err) {
                msg.classList.add('bg-red-50/80', 'text-red-600', 'border', 'border-red-100');
                msg.innerHTML = \`<div class="flex items-center gap-2"><i data-lucide="alert-circle" class="w-4 h-4"></i> <span>\${err.message || '网络错误'}</span></div>\`;
                lucide.createIcons();
            } finally {
                btn.disabled = false; 
                btn.innerHTML = originalBtnText;
                if(typeof turnstile !== 'undefined') turnstile.reset();
            }
        });
    </script>
</body>
</html>
`;

// --- 3. 后台管理页面 HTML ---
const HTML_ADMIN_PAGE = (skuMapJson, hiddenUser, adminPath) => `
<!DOCTYPE html>
<html lang="zh-CN">
${HTML_HEAD('M365 Console')}
<body class="p-6 md:p-10 min-h-screen">
    <div class="max-w-7xl mx-auto space-y-8">
        <!-- Header -->
        <div class="flex flex-col md:flex-row md:items-center justify-between gap-4">
            <div>
                <h1 class="text-3xl font-extrabold text-gray-900 tracking-tight">Dashboard</h1>
                <p class="text-gray-500 text-sm font-medium mt-1">M365 Subscription & User Management</p>
            </div>
            <div class="flex items-center gap-3">
                <div class="px-4 py-2 bg-white/60 rounded-full text-xs font-bold text-gray-500 border border-white/60 shadow-sm backdrop-blur-md">
                    v4.0.0
                </div>
            </div>
        </div>

        <!-- Navigation Tabs -->
        <div class="flex p-1.5 bg-gray-200/50 rounded-[20px] w-fit backdrop-blur-md">
            <button onclick="switchTab('users')" id="tab-btn-users" class="px-6 py-2.5 rounded-2xl text-sm font-semibold transition-all shadow-sm bg-white text-gray-900">
                用户列表
            </button>
            <button onclick="switchTab('invites')" id="tab-btn-invites" class="px-6 py-2.5 rounded-2xl text-sm font-semibold text-gray-500 hover:text-gray-700 transition-all ml-1">
                邀请码管理
            </button>
        </div>

        <!-- Content Area -->
        <div class="glass-panel rounded-[40px] min-h-[600px] relative overflow-hidden">
            
            <!-- Users Tab -->
            <div id="view-users" class="p-8 h-full flex flex-col">
                <div class="flex flex-wrap items-center justify-between gap-4 mb-8">
                    <div class="flex items-center gap-3">
                        <button onclick="loadUsers()" class="btn-secondary w-10 h-10 rounded-xl flex items-center justify-center">
                            <i data-lucide="refresh-ccw" class="w-4 h-4"></i>
                        </button>
                        <span class="text-sm font-bold text-gray-400 uppercase tracking-widest pl-2" id="user-count-label">Loading...</span>
                    </div>
                    <div class="flex items-center gap-3">
                        <button onclick="toggleDeleteMode()" id="btn-del-mode" class="btn-secondary px-5 py-2.5 rounded-xl text-sm font-semibold text-gray-600">
                            批量管理
                        </button>
                        <button onclick="confirmBatchDelete()" id="btn-del-confirm" class="bg-red-500 text-white px-5 py-2.5 rounded-xl text-sm font-semibold shadow-lg hover:bg-red-600 active:scale-95 transition-all hidden">
                            确认删除 (<span id="del-count">0</span>)
                        </button>
                    </div>
                </div>

                <div class="flex-1 overflow-auto -mx-8 px-8 no-scrollbar">
                    <table class="w-full text-left border-collapse">
                        <thead>
                            <tr class="border-b border-gray-200/50">
                                <th class="py-4 pl-2 w-12 del-col hidden">
                                    <input type="checkbox" onchange="toggleSelectAll(this)" class="rounded-md border-gray-300 text-gray-900 focus:ring-0 w-5 h-5 bg-white/50">
                                </th>
                                <th class="py-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest">User</th>
                                <th class="py-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest">Email (UPN)</th>
                                <th class="py-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest">Status</th>
                                <th class="py-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest">Created</th>
                            </tr>
                        </thead>
                        <tbody id="user-list" class="text-sm text-gray-700">
                            <!-- JS Render -->
                        </tbody>
                    </table>
                </div>
            </div>

            <!-- Invites Tab -->
            <div id="view-invites" class="p-8 h-full flex flex-col hidden">
                <div class="flex flex-wrap items-center justify-between gap-4 mb-8">
                    <div class="flex items-center gap-3">
                        <button onclick="promptGenerate()" class="btn-primary px-5 py-2.5 rounded-xl text-sm font-semibold flex items-center gap-2 shadow-lg">
                            <i data-lucide="plus" class="w-4 h-4"></i> Generate
                        </button>
                        <button onclick="loadInvites()" class="btn-secondary w-10 h-10 rounded-xl flex items-center justify-center">
                            <i data-lucide="refresh-ccw" class="w-4 h-4"></i>
                        </button>
                        <button onclick="copyUnused()" class="btn-secondary px-5 py-2.5 rounded-xl text-sm font-semibold text-gray-600">
                            复制未使用
                        </button>
                    </div>
                    <div class="flex items-center gap-4">
                        <span class="text-sm font-bold text-gray-400 uppercase tracking-widest" id="invite-stats">...</span>
                        <button onclick="clearAllInvites()" class="text-red-500 hover:bg-red-50 px-4 py-2 rounded-xl text-xs font-bold transition-colors">
                            清空所有
                        </button>
                    </div>
                </div>

                <div class="flex-1 overflow-auto -mx-8 px-8 no-scrollbar">
                    <table class="w-full text-left border-collapse">
                        <thead>
                            <tr class="border-b border-gray-200/50">
                                <th class="py-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest">Code</th>
                                <th class="py-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest">Status</th>
                                <th class="py-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest">Bound User</th>
                                <th class="py-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest">Created</th>
                                <th class="py-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest w-20">Action</th>
                            </tr>
                        </thead>
                        <tbody id="invite-list" class="text-sm text-gray-700">
                            <!-- JS Render -->
                        </tbody>
                    </table>
                </div>
            </div>

            <!-- Loading Overlay -->
            <div id="loading-overlay" class="absolute inset-0 bg-white/40 backdrop-blur-sm flex items-center justify-center z-50 hidden">
                <div class="loading-spinner w-8 h-8 border-gray-400/30 border-l-gray-800"></div>
            </div>
        </div>
    </div>

    <script>
        const API_BASE = '${adminPath}/api';
        const HIDDEN_USER = ${JSON.stringify(hiddenUser || '')}.toLowerCase();
        let usersCache = [];
        let invitesCache = [];
        let deleteMode = false;

        // Init
        document.addEventListener('DOMContentLoaded', () => {
            loadUsers();
            loadInvites(true); // silent load
            lucide.createIcons();
        });

        // Tab Switcher
        function switchTab(tab) {
            ['users', 'invites'].forEach(t => {
                const btn = document.getElementById('tab-btn-' + t);
                const view = document.getElementById('view-' + t);
                
                if (t === tab) {
                    btn.className = 'px-6 py-2.5 rounded-2xl text-sm font-semibold transition-all shadow-sm bg-white text-gray-900';
                    view.classList.remove('hidden');
                } else {
                    btn.className = 'px-6 py-2.5 rounded-2xl text-sm font-semibold text-gray-500 hover:text-gray-700 transition-all ml-1';
                    view.classList.add('hidden');
                }
            });
            if (tab === 'users') loadUsers();
            if (tab === 'invites') loadInvites();
        }

        // --- Users Logic ---
        async function loadUsers() {
            setLoading(true);
            try {
                const res = await fetch(API_BASE + '/users');
                if (res.status === 401) return window.location.reload(); // Cookie expired
                const data = await res.json();
                usersCache = data.value || [];
                renderUsers();
            } catch(e) { console.error(e); alert('Failed to load users'); }
            setLoading(false);
        }

        function renderUsers() {
            const tbody = document.getElementById('user-list');
            document.getElementById('user-count-label').innerText = \`\${usersCache.length} USERS\`;
            
            tbody.innerHTML = usersCache.map(u => {
                const isHidden = u.userPrincipalName.toLowerCase() === HIDDEN_USER;
                const checkbox = isHidden ? '' : \`<input type="checkbox" class="user-check rounded border-gray-300 text-gray-900 focus:ring-0 w-5 h-5 bg-white/50" value="\${u.id}" onchange="updateDelCount()">\`;
                const badge = isHidden ? '<span class="px-2 py-0.5 bg-gray-200 rounded text-[10px] font-bold text-gray-500 ml-2">ADMIN</span>' : '';
                const license = u.assignedLicenses.length 
                    ? '<span class="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium bg-green-100 text-green-800">Active</span>'
                    : '<span class="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium bg-gray-100 text-gray-800">None</span>';
                
                return \`
                <tr class="group hover:bg-white/40 transition-colors border-b border-gray-100/50 last:border-0">
                    <td class="py-4 pl-2 del-col \${deleteMode ? '' : 'hidden'}">\${checkbox}</td>
                    <td class="py-4 font-medium text-gray-900">\${u.displayName} \${badge}</td>
                    <td class="py-4 font-mono text-xs text-gray-500">\${u.userPrincipalName}</td>
                    <td class="py-4">\${license}</td>
                    <td class="py-4 text-gray-500 text-xs">\${new Date(u.createdDateTime).toLocaleDateString()}</td>
                </tr>
                \`;
            }).join('');
        }

        function toggleDeleteMode() {
            deleteMode = !deleteMode;
            document.querySelectorAll('.del-col').forEach(el => el.classList.toggle('hidden', !deleteMode));
            
            const btnMode = document.getElementById('btn-del-mode');
            const btnConfirm = document.getElementById('btn-del-confirm');
            
            if (deleteMode) {
                btnMode.innerText = '取消';
                btnMode.classList.add('bg-gray-200');
                btnConfirm.classList.remove('hidden');
            } else {
                btnMode.innerText = '批量管理';
                btnMode.classList.remove('bg-gray-200');
                btnConfirm.classList.add('hidden');
                document.querySelectorAll('.user-check').forEach(c => c.checked = false);
                updateDelCount();
            }
        }

        function toggleSelectAll(src) {
            document.querySelectorAll('.user-check').forEach(c => c.checked = src.checked);
            updateDelCount();
        }

        function updateDelCount() {
            document.getElementById('del-count').innerText = document.querySelectorAll('.user-check:checked').length;
        }

        async function confirmBatchDelete() {
            const checks = document.querySelectorAll('.user-check:checked');
            if(checks.length === 0) return;
            if(!confirm(\`确定要永久删除这 \${checks.length} 个用户吗？关联的邀请码也将被清理。\n此操作不可撤销。\n\nAre you sure?\`)) return;

            setLoading(true);
            
            // Sync invites first to ensure cleanup works
            if(invitesCache.length === 0) await loadInvites(true);

            for(const cb of checks) {
                const uid = cb.value;
                try {
                    // 1. Delete User
                    const res = await fetch(API_BASE + '/users/' + uid, { method: 'DELETE' });
                    if(res.ok || res.status === 404) {
                        // 2. Cleanup Invite
                        const user = usersCache.find(u => u.id === uid);
                        if(user) {
                            const invite = invitesCache.find(i => i.user && i.user.toLowerCase() === user.userPrincipalName.toLowerCase());
                            if(invite) {
                                await fetch(API_BASE + '/invites?code=' + invite.code, { method: 'DELETE' });
                            }
                        }
                    }
                } catch(e) { console.error(e); }
            }
            
            toggleDeleteMode();
            loadUsers();
            setTimeout(() => loadInvites(true), 1000);
            setLoading(false);
        }

        // --- Invites Logic ---
        async function loadInvites(silent = false) {
            if(!silent) setLoading(true);
            try {
                const res = await fetch(API_BASE + '/invites');
                const data = await res.json();
                invitesCache = data;
                if(!silent) renderInvites();
            } catch(e) {}
            if(!silent) setLoading(false);
        }

        function renderInvites() {
            const tbody = document.getElementById('invite-list');
            document.getElementById('invite-stats').innerText = \`\${invitesCache.length} CODES\`;
            
            if(invitesCache.length === 0) {
                tbody.innerHTML = '<tr><td colspan="5" class="py-10 text-center text-gray-400">No data</td></tr>';
                return;
            }

            tbody.innerHTML = invitesCache.map(i => {
                const isUsed = i.status === 'used';
                const status = isUsed 
                    ? '<span class="inline-flex items-center px-2 py-0.5 rounded text-xs font-bold bg-gray-100 text-gray-500">USED</span>'
                    : '<span class="inline-flex items-center px-2 py-0.5 rounded text-xs font-bold bg-green-100 text-green-700">UNUSED</span>';
                
                return \`
                <tr class="border-b border-gray-100/50 last:border-0 hover:bg-white/40">
                    <td class="py-4 font-mono font-bold text-gray-800 cursor-pointer hover:text-blue-600 transition-colors" onclick="copyText('\${i.code}')">\${i.code}</td>
                    <td class="py-4">\${status}</td>
                    <td class="py-4 text-gray-500 font-mono text-xs">\${i.user || '-'}</td>
                    <td class="py-4 text-gray-400 text-xs">\${new Date(i.createdAt).toLocaleDateString()}</td>
                    <td class="py-4">
                        <button onclick="deleteInvite('\${i.code}')" class="text-red-400 hover:text-red-600 p-2 rounded-lg hover:bg-red-50 transition-all">
                            <i data-lucide="trash-2" class="w-4 h-4"></i>
                        </button>
                    </td>
                </tr>
                \`;
            }).join('');
            lucide.createIcons();
        }

        async function promptGenerate() {
            const count = prompt('生成数量 (1-50):', '1');
            if(!count) return;
            setLoading(true);
            try {
                await fetch(API_BASE + '/invites', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({count: parseInt(count)})
                });
                loadInvites();
            } catch(e) { alert('Error'); }
            setLoading(false);
        }

        async function deleteInvite(code) {
            if(!confirm('Delete this invite code?')) return;
            await fetch(API_BASE + '/invites?code=' + code, { method: 'DELETE' });
            loadInvites();
        }

        async function clearAllInvites() {
            if(!confirm('DANGER: Clear ALL invite codes?')) return;
            setLoading(true);
            await fetch(API_BASE + '/invites?action=clear_all', { method: 'DELETE' });
            loadInvites();
            setLoading(false);
        }

        function copyUnused() {
            const codes = invitesCache.filter(i => i.status !== 'used').map(i => i.code).join('\\n');
            if(!codes) return alert('No unused codes');
            navigator.clipboard.writeText(codes).then(() => alert('Copied!'));
        }

        function copyText(t) {
            navigator.clipboard.writeText(t).then(() => alert('Code copied: ' + t));
        }

        function setLoading(state) {
            const el = document.getElementById('loading-overlay');
            if(state) el.classList.remove('hidden');
            else el.classList.add('hidden');
        }
    </script>
</body>
</html>
`;

export default {
    async fetch(request, env) {
        const url = new URL(request.url);
        const adminPath = getEnv(env.ADMIN_PATH, '/admin').replace(/\/$/, ''); // 移除尾部斜杠
        
        // --- 0. 路由分发 ---
        
        // 静态资源 (注册页)
        if (url.pathname === '/' && request.method === 'GET') {
            let skuOptions = [];
            try { skuOptions = Object.keys(JSON.parse(env.SKU_MAP || '{}')); } catch (e) {}
            return new Response(HTML_REGISTER_PAGE(env.TURNSTILE_SITE_KEY, skuOptions), {
                headers: { 'Content-Type': 'text/html;charset=UTF-8' },
            });
        }

        // 注册 API
        if (url.pathname === '/' && request.method === 'POST') {
            return handleRegister(request, env);
        }

        // --- 后台相关路由 ---
        
        if (url.pathname.startsWith(adminPath)) {
            
            // 登录页面 & API
            if (url.pathname === adminPath || url.pathname === `${adminPath}/`) {
                // 检查 Cookie
                if (await checkAuth(request, env)) {
                    return new Response(HTML_ADMIN_PAGE(env.SKU_MAP, getEnv(env.HIDDEN_USER), adminPath), {
                        headers: { 'Content-Type': 'text/html;charset=UTF-8' }
                    });
                }
                return new Response(HTML_LOGIN_PAGE(env.TURNSTILE_SITE_KEY, adminPath), {
                    headers: { 'Content-Type': 'text/html;charset=UTF-8' }
                });
            }

            // 登录提交 API
            if (url.pathname === `${adminPath}/login` && request.method === 'POST') {
                const body = await request.json();
                
                // 1. Turnstile 验证
                const turnstileRes = await fetch('https://challenges.cloudflare.com/turnstile/v0/siteverify', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        secret: env.CF_TURNSTILE_SECRET,
                        response: body.cf_turnstile,
                        remoteip: request.headers.get('CF-Connecting-IP')
                    })
                });
                const turnstileData = await turnstileRes.json();
                if (!turnstileData.success) {
                    return Response.json({ success: false, message: 'Turnstile verification failed' }, { status: 403 });
                }

                // 2. 密码验证
                if (body.token === env.ADMIN_TOKEN) {
                    // 设置 HttpOnly Cookie, 有效期 1 天
                    const headers = new Headers();
                    headers.append('Set-Cookie', `m365_admin_token=${env.ADMIN_TOKEN}; Path=/; Max-Age=86400; HttpOnly; SameSite=Strict; Secure`);
                    return new Response(JSON.stringify({ success: true }), { status: 200, headers });
                }
                return Response.json({ success: false, message: 'Invalid Token' }, { status: 401 });
            }

            // --- API 路由 (需要鉴权) ---
            if (url.pathname.startsWith(`${adminPath}/api`)) {
                if (!(await checkAuth(request, env))) return Response.json({ error: 'Unauthorized' }, { status: 401 });

                // Users API
                if (url.pathname === `${adminPath}/api/users`) {
                    if (request.method === 'DELETE') {
                         // Batch delete is not supported on this path, handled by single ID route
                    }
                    if (request.method === 'GET') return handleGetUsers(env);
                }
                
                // Single User Delete
                if (url.pathname.startsWith(`${adminPath}/api/users/`) && request.method === 'DELETE') {
                    const uid = url.pathname.split('/').pop();
                    return handleDeleteUser(uid, env);
                }

                // Invites API
                if (url.pathname === `${adminPath}/api/invites`) {
                    if (request.method === 'GET') return handleGetInvites(env);
                    if (request.method === 'POST') return handleCreateInvites(request, env);
                    if (request.method === 'DELETE') return handleDeleteInvites(request, env);
                }
            }
        }

        return new Response('404 Not Found', { status: 404 });
    }
};

// --- 业务逻辑处理器 ---

async function checkAuth(req, env) {
    // 优先检查 Cookie
    const cookie = req.headers.get('Cookie');
    if (cookie && cookie.includes(`m365_admin_token=${env.ADMIN_TOKEN}`)) return true;
    
    // 兼容旧版 URL Token (可选，如果不需要可删除)
    const url = new URL(req.url);
    if (url.searchParams.get('token') === env.ADMIN_TOKEN) return true;
    
    return false;
}

// 注册逻辑 (未改动核心逻辑，仅适配新结构)
async function handleRegister(req, env) {
    try {
        const formData = await req.formData();
        const inviteCode = formData.get('inviteCode');
        
        // 1. Verify Turnstile
        if (env.CF_TURNSTILE_SECRET) {
            const tRes = await fetch('https://challenges.cloudflare.com/turnstile/v0/siteverify', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ secret: env.CF_TURNSTILE_SECRET, response: formData.get('cf-turnstile-response'), remoteip: req.headers.get('CF-Connecting-IP') })
            });
            if (!(await tRes.json()).success) return Response.json({ success: false, message: '人机验证失败' });
        }

        // 2. Verify Invite Code
        if (!env.INVITE_CODES) return Response.json({ success: false, message: '系统配置错误(KV)' });
        const kvKey = `invite:${inviteCode}`;
        const inviteData = await env.INVITE_CODES.get(kvKey, { type: 'json' });
        if (!inviteData) return Response.json({ success: false, message: '邀请码无效' });
        if (inviteData.status === 'used') return Response.json({ success: false, message: '邀请码已被使用' });

        // 3. Create User
        const username = formData.get('username');
        const password = formData.get('password');
        const skuName = formData.get('skuName');
        
        let skuId = null;
        try { skuId = JSON.parse(env.SKU_MAP || '{}')[skuName]; } catch(e){}
        if (!skuId) return Response.json({ success: false, message: '订阅类型无效' });
        if (!checkPasswordComplexity(password)) return Response.json({ success: false, message: '密码需包含大小写字母和数字，且长度>8' });

        const accessToken = await getAccessToken(env);
        const userEmail = `${username}@${getEnv(env.DEFAULT_DOMAIN)}`;
        const hiddenUser = getEnv(env.HIDDEN_USER);
        if (hiddenUser && userEmail.toLowerCase() === hiddenUser.toLowerCase()) return Response.json({ success: false, message: '用户名不可用' });

        const createReq = await fetch('https://graph.microsoft.com/v1.0/users', {
            method: 'POST',
            headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
            body: JSON.stringify({
                accountEnabled: true,
                displayName: username,
                mailNickname: username,
                userPrincipalName: userEmail,
                passwordProfile: { forceChangePasswordNextSignIn: false, password: password },
                usageLocation: "CN"
            })
        });

        if (!createReq.ok) {
            const err = await createReq.json();
            return Response.json({ success: false, message: err.error?.message || '创建用户失败' });
        }
        const newUser = await createReq.json();

        // 4. Assign License
        await fetch(`https://graph.microsoft.com/v1.0/users/${newUser.id}/assignLicense`, {
            method: 'POST',
            headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
            body: JSON.stringify({ addLicenses: [{ disabledPlans: [], skuId: skuId }], removeLicenses: [] })
        });

        // 5. Update Invite Status
        inviteData.status = 'used';
        inviteData.user = userEmail;
        inviteData.usedAt = Date.now();
        await env.INVITE_CODES.put(kvKey, JSON.stringify(inviteData));

        return Response.json({ success: true, email: userEmail });
    } catch (e) {
        return Response.json({ success: false, message: e.message });
    }
}

async function handleGetUsers(env) {
    try {
        const token = await getAccessToken(env);
        // Fallback sort
        const res = await fetch('https://graph.microsoft.com/v1.0/users?$select=id,displayName,userPrincipalName,createdDateTime,assignedLicenses&$top=100', {
            headers: { 'Authorization': `Bearer ${token}` }
        });
        const data = await res.json();
        if (data.value) data.value.sort((a, b) => new Date(b.createdDateTime) - new Date(a.createdDateTime));
        return Response.json(data);
    } catch (e) { return Response.json({ error: e.message }, { status: 500 }); }
}

async function handleDeleteUser(uid, env) {
    try {
        const token = await getAccessToken(env);
        // Check Hidden
        const hiddenUser = getEnv(env.HIDDEN_USER);
        if (hiddenUser) {
            const check = await fetch(`https://graph.microsoft.com/v1.0/users/${uid}?$select=userPrincipalName`, { headers: { Authorization: `Bearer ${token}` } });
            if (check.ok) {
                const u = await check.json();
                if (u.userPrincipalName.toLowerCase() === hiddenUser.toLowerCase()) return Response.json({ error: 'Protected' }, { status: 403 });
            }
        }
        await fetch(`https://graph.microsoft.com/v1.0/users/${uid}`, { method: 'DELETE', headers: { Authorization: `Bearer ${token}` } });
        return Response.json({ success: true });
    } catch (e) { return Response.json({ error: e.message }, { status: 500 }); }
}

async function handleGetInvites(env) {
    const list = await env.INVITE_CODES.list({ prefix: 'invite:', limit: 1000 });
    const invites = [];
    for (const key of list.keys) {
        const v = await env.INVITE_CODES.get(key.name, { type: 'json' });
        if (v) invites.push(v);
    }
    invites.sort((a, b) => b.createdAt - a.createdAt);
    return Response.json(invites);
}

async function handleCreateInvites(req, env) {
    const body = await req.json();
    const count = Math.min(Math.max(parseInt(body.count) || 1, 1), 50);
    const created = [];
    for (let i = 0; i < count; i++) {
        const code = generateRandomString(10);
        const data = { code, status: 'unused', createdAt: Date.now(), user: null, usedAt: null };
        await env.INVITE_CODES.put(`invite:${code}`, JSON.stringify(data));
        created.push(data);
    }
    return Response.json(created);
}

async function handleDeleteInvites(req, env) {
    const url = new URL(req.url);
    const action = url.searchParams.get('action');
    const code = url.searchParams.get('code');

    if (action === 'clear_all') {
        const list = await env.INVITE_CODES.list({ prefix: 'invite:', limit: 1000 });
        for (const k of list.keys) await env.INVITE_CODES.delete(k.name);
        return Response.json({ success: true });
    }
    if (code) {
        await env.INVITE_CODES.delete(`invite:${code}`);
        return Response.json({ success: true });
    }
    return Response.json({ error: 'Invalid parameters' }, { status: 400 });
}

// --- OAuth ---
async function getAccessToken(env) {
    const params = new URLSearchParams();
    params.append('client_id', env.AZURE_CLIENT_ID);
    params.append('scope', 'https://graph.microsoft.com/.default');
    params.append('client_secret', env.AZURE_CLIENT_SECRET);
    params.append('grant_type', 'client_credentials');
    const res = await fetch(`https://login.microsoftonline.com/${env.AZURE_TENANT_ID}/oauth2/v2.0/token`, { method: 'POST', body: params });
    const data = await res.json();
    if (data.error) throw new Error(data.error_description || 'Auth Failed');
    return data.access_token;
}

function generateRandomString(len) {
    const c = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
    let r = '';
    for (let i = 0; i < len; i++) r += c.charAt(Math.floor(Math.random() * c.length));
    return r;
}
