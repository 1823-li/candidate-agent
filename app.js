// ========== 全局状态 ==========
const state = {
    files: { resume: null, profile: null, gpi: null, assessment: null },
    texts: { resume: '', profile: '', gpi: '', assessment: '' },
    reportText: '',
    interviewText: '',
    score: 0,
    profileLib: {},        // { 岗位名: 画像文本 }
    profileLibMeta: null,  // { source, time, count }
    selectedPosition: ''
};

const LIB_STORAGE_KEY = 'hm_profile_library_v1';
const DEFAULT_PROFILE_URL = './data/人才画像.xlsx';
const HISTORY_STORAGE_KEY = 'hm_analysis_history_v1';
const MAX_HISTORY_ITEMS = 50; // 最多保留50条历史记录

// ========== 历史记录管理 ==========
function loadHistory() {
    try {
        const data = localStorage.getItem(HISTORY_STORAGE_KEY);
        return data ? JSON.parse(data) : [];
    } catch (e) {
        console.warn('历史记录加载失败', e);
        return [];
    }
}

function saveHistoryList(list) {
    try {
        // 限制条数，超出时移除最旧的
        if (list.length > MAX_HISTORY_ITEMS) {
            list = list.slice(0, MAX_HISTORY_ITEMS);
        }
        localStorage.setItem(HISTORY_STORAGE_KEY, JSON.stringify(list));
    } catch (e) {
        console.warn('历史记录保存失败', e);
        // localStorage 可能已满，尝试清理旧记录
        if (e.name === 'QuotaExceededError' || e.code === 22) {
            try {
                const trimmed = list.slice(0, Math.floor(list.length / 2));
                localStorage.setItem(HISTORY_STORAGE_KEY, JSON.stringify(trimmed));
                showToast('⚠ 存储空间不足，已自动清理较早的记录', 'error');
            } catch (e2) {
                console.error('清理后仍保存失败', e2);
            }
        }
    }
}

function addToHistory(record) {
    const list = loadHistory();
    list.unshift(record); // 最新的排最前
    saveHistoryList(list);
    renderHistory();
}

function renderHistory() {
    const list = loadHistory();
    const section = document.getElementById('historySection');
    const listEl = document.getElementById('historyList');

    if (list.length === 0) {
        section.classList.remove('active');
        return;
    }

    section.classList.add('active');
    listEl.innerHTML = list.map((item, idx) => {
        const scoreClass = item.score >= 80 ? 'score-high' : item.score >= 60 ? 'score-mid' : 'score-low';
        const dateStr = item.time || '-';
        const position = item.position || '未选择';
        return `<li class="history-item">
            <div class="history-score ${scoreClass}">${item.score}</div>
            <div class="history-info">
                <div class="h-name">${escapeHtml(item.name)}</div>
                <div class="h-meta">🏢 ${escapeHtml(position)} 📅 ${escapeHtml(dateStr)}</div>
            </div>
            <div class="history-actions">
                <button class="btn-hist view" onclick="viewHistoryItem(${idx})" title="查看详情">👁 查看</button>
                <button class="btn-hist dl-docx" onclick="downloadHistoryItem(${idx}, 'match', 'docx')" title="下载报告Word">📥 报告</button>
                <button class="btn-hist dl-docx" onclick="downloadHistoryItem(${idx}, 'interview', 'docx')" title="下载面试提纲Word">📥 面试提纲</button>
                <button class="btn-hist del" onclick="deleteHistoryItem(${idx})" title="删除此记录">🗑</button>
            </div>
        </li>`;
    }).join('');
}

function viewHistoryItem(idx) {
    const list = loadHistory();
    const item = list[idx];
    if (!item) return;

    const scoreClass = item.score >= 80 ? 'score-high' : item.score >= 60 ? 'score-mid' : 'score-low';
    const body = document.getElementById('historyModalBody');

    let scoreLabel = '';
    if (item.score >= 80) scoreLabel = '匹配度高 - 推荐进入下一轮';
    else if (item.score >= 60) scoreLabel = '匹配度中等 - 建议进一步评估';
    else scoreLabel = '匹配度偏低 - 谨慎考虑';

    body.innerHTML = `
        <h2>📋 ${escapeHtml(item.name)} - 分析记录详情</h2>
        <div style="font-size:0.82rem;color:var(--text-secondary);margin-bottom:1rem;">
            🏢 应聘岗位：${escapeHtml(item.position || '未选择')}  📅 分析时间：${escapeHtml(item.time || '-')}
        </div>
        <div class="modal-score">
            <div class="score-circle ${scoreClass}">${item.score}</div>
            <div style="font-size:0.85rem;color:var(--text-secondary);margin-top:0.4rem;">${scoreLabel}</div>
        </div>
        <div class="modal-section">
            <h3>📝 综合岗位匹配度报告</h3>
            <div class="modal-content">${escapeHtml(cleanMarkdown(item.reportText) || '（无内容）')}</div>
        </div>
        <div class="modal-section">
            <h3>🎤 面试提纲</h3>
            <div class="modal-content">${escapeHtml(cleanMarkdown(item.interviewText) || '（无内容）')}</div>
        </div>
        <div style="text-align:center;margin-top:1rem;display:flex;gap:0.5rem;justify-content:center;flex-wrap:wrap;">
            <button class="btn-download docx" onclick="downloadHistoryItem(${idx}, 'match', 'docx')">📥 下载报告 (Word)</button>
            <button class="btn-download docx" onclick="downloadHistoryItem(${idx}, 'interview', 'docx')">📥 下载面试提纲 (Word)</button>
        </div>
    `;

    document.getElementById('historyModalOverlay').classList.add('active');
}

function closeHistoryModal(event) {
    // 点击遮罩层关闭，或显式关闭
    if (!event || event.target === document.getElementById('historyModalOverlay')) {
        document.getElementById('historyModalOverlay').classList.remove('active');
    }
}

function deleteHistoryItem(idx) {
    if (!confirm('确定删除此条分析记录吗？')) return;
    const list = loadHistory();
    list.splice(idx, 1);
    saveHistoryList(list);
    renderHistory();
    showToast('🗑 已删除该记录', 'success');
}

function clearAllHistory() {
    if (!confirm('确定清空所有分析记录吗？此操作不可恢复！')) return;
    localStorage.removeItem(HISTORY_STORAGE_KEY);
    renderHistory();
    showToast('🗑 已清空所有记录', 'success');
}

function downloadHistoryItem(idx, type, format) {
    const list = loadHistory();
    const item = list[idx];
    if (!item) return;

    let content, filename;
    if (type === 'match') {
        content = item.reportText || '';
        filename = `${item.name}_综合岗位匹配度报告`;
    } else {
        content = item.interviewText || '';
        filename = `${item.name}_面试提纲`;
    }

    if (format === 'docx') {
        if (typeof docx === 'undefined') {
            showToast('⚠ Word 导出库未加载，请检查网络后重试', 'error');
            return;
        }
        generateDocx(content, filename).then(blob => {
            saveFile(blob, filename + '.docx');
            showToast(`📦 已下载: ${filename}.docx`, 'success');
        }).catch(err => {
            showToast(`⚠ Word 导出失败: ${err.message}`, 'error');
        });
        return;
    }

    const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
    saveFile(blob, filename + '.txt');
    showToast(`📥 已下载: ${filename}.txt`, 'success');
}

function escapeHtml(text) {
    if (!text) return '';
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

// 页面初始化时加载历史记录
loadHistory();
renderHistory();

// ========== 岗位人才画像库 ==========
async function initProfileLibrary() {
    // 优先使用本地缓存（导入过的更新版本），否则加载默认文件
    try {
        const cached = localStorage.getItem(LIB_STORAGE_KEY);
        if (cached) {
            const data = JSON.parse(cached);
            if (data.lib && Object.keys(data.lib).length > 0) {
                state.profileLib = data.lib;
                state.profileLibMeta = data.meta;
                renderProfileLibrary();
                return;
            }
        }
    } catch (e) { console.warn('缓存读取失败', e); }

    await loadDefaultProfiles();
}

async function loadDefaultProfiles() {
    try {
        const resp = await fetch(DEFAULT_PROFILE_URL + '?t=' + Date.now());
        if (!resp.ok) throw new Error('HTTP ' + resp.status);
        const buf = await resp.arrayBuffer();
        const lib = parseProfileWorkbook(buf);
        state.profileLib = lib;
        state.profileLibMeta = { source: '默认文件 (data/人才画像.xlsx)', time: new Date().toLocaleString('zh-CN') };
        renderProfileLibrary();
    } catch (e) {
        document.getElementById('libStatus').innerHTML =
            '⚠️ 默认画像文件加载失败（' + e.message + '），请点击"导入/更新画像文件"手动上传 .xlsx';
    }
}

function parseProfileWorkbook(buf) {
    if (typeof XLSX === 'undefined') throw new Error('Excel 解析库未加载，请检查网络');
    const wb = XLSX.read(buf, { type: 'array' });
    const lib = {};
    for (const name of wb.SheetNames) {
        // 跳过空表/临时表
        if (/^sheet\d+$/i.test(name.trim())) continue;
        const text = XLSX.utils.sheet_to_csv(wb.Sheets[name], { FS: '\t', blankrows: false }).trim();
        // 只保留包含岗位 JD 特征的有效画像表（排除如操作说明类 Sheet）
        if (text && text.includes('岗位名称') && text.replace(/[\t\n\r ]/g, '').length > 20) {
            lib[name.trim()] = text;
        }
    }
    if (Object.keys(lib).length === 0) throw new Error('未解析到有效的岗位画像（Sheet）');
    return lib;
}

function renderProfileLibrary() {
    const select = document.getElementById('jobPosition');
    const names = Object.keys(state.profileLib);
    select.innerHTML = '<option value="">—— 请选择岗位 ——</option>' +
        names.map(n => `<option value="${n.replace(/"/g, '&quot;')}">${n}</option>`).join('');

    const meta = state.profileLibMeta || {};
    document.getElementById('libStatus').innerHTML =
        `当前画像库：<b>${names.length}</b> 个岗位 | 来源：${meta.source || '默认文件'} | 更新时间：${meta.time || '-'}`;

    // 恢复之前的选中项
    if (state.selectedPosition && state.profileLib[state.selectedPosition]) {
        select.value = state.selectedPosition;
        onPositionSelect();
    } else {
        document.getElementById('profilePreview').classList.remove('active');
        state.texts.profile = '';
    }
}

function onPositionSelect() {
    const select = document.getElementById('jobPosition');
    const name = select.value;
    state.selectedPosition = name;
    const preview = document.getElementById('profilePreview');
    if (!name) {
        state.texts.profile = '';
        preview.classList.remove('active');
        return;
    }
    state.texts.profile = state.profileLib[name];
    preview.textContent = state.texts.profile;
    preview.classList.add('active');
}

async function importProfileFile(input) {
    const file = input.files[0];
    if (!file) return;
    try {
        const buf = await file.arrayBuffer();
        const lib = parseProfileWorkbook(buf);
        state.profileLib = lib;
        state.profileLibMeta = { source: `导入文件 (${file.name})`, time: new Date().toLocaleString('zh-CN') };
        // 持久化到浏览器，下次打开自动使用新版
        localStorage.setItem(LIB_STORAGE_KEY, JSON.stringify({ lib: state.profileLib, meta: state.profileLibMeta }));
        state.selectedPosition = '';
        renderProfileLibrary();
        showToast(`✅ 画像库已更新：${Object.keys(lib).length} 个岗位`, 'success');
    } catch (err) {
        showToast(`❌ 画像文件解析失败: ${err.message}`, 'error');
    } finally {
        input.value = '';
    }
}

function resetToDefaultProfiles() {
    localStorage.removeItem(LIB_STORAGE_KEY);
    state.selectedPosition = '';
    loadDefaultProfiles();
    showToast('✅ 已恢复为默认画像文件', 'success');
}

// 页面加载时初始化画像库
initProfileLibrary();

// ========== 文件上传处理 ==========
document.querySelectorAll('input[type="file"]').forEach(input => {
    if (!input.dataset.target) return; // 跳过画像库导入等非候选人材料输入
    input.addEventListener('change', async (e) => {
        const target = e.target.dataset.target;
        const file = e.target.files[0];
        if (!file) return;
        state.files[target] = file;
        document.getElementById(`fname-${target}`).textContent = file.name;
        document.getElementById(`upload-${target}`).classList.add('uploaded');
        try {
            state.texts[target] = await extractText(file);
            showToast(`✅ ${file.name} 解析成功`, 'success');
        } catch (err) {
            // 解析失败：回滚上传状态，避免误认为已就绪
            state.files[target] = null;
            state.texts[target] = '';
            document.getElementById(`fname-${target}`).textContent = '';
            document.getElementById(`upload-${target}`).classList.remove('uploaded');
            e.target.value = '';
            showToast(`❌ ${err.message}`, 'error');
        }
    });
});

// ========== 文件文本提取 ==========
async function extractText(file) {
    const ext = file.name.split('.').pop().toLowerCase();
    let text;
    if (['txt', 'csv', 'md', 'markdown'].includes(ext)) text = await file.text();
    else if (ext === 'html' || ext === 'htm') text = await extractHtmlText(file);
    else if (ext === 'pdf') return await extractPdfText(file);
    else if (ext === 'docx') return await extractDocxText(file);
    else if (ext === 'doc') throw new Error('不支持旧版 .doc 格式，请用 Word 另存为 .docx 后重新上传');
    else throw new Error('不支持的文件格式: .' + ext);

    // 文本类文件：检测是否为二进制/损坏文件
    if (isBinaryText(text)) {
        throw new Error(`文件 "${file.name}" 内容为二进制乱码，可能已损坏，请重新下载/导出后再上传`);
    }
    if (!text.trim()) {
        throw new Error(`文件 "${file.name}" 未提取到任何文本内容，请确认文件是否完好`);
    }
    return text;
}

function isBinaryText(text) {
    const sample = text.slice(0, 800);
    if (!sample.trim()) return text.length > 0; // 有内容但全是空白/控制符
    const controls = (sample.match(/[\x00-\x08\x0E-\x1F]/g) || []).length;
    return controls / sample.length > 0.1;
}

async function extractHtmlText(file) {
    const html = await file.text();
    const doc = new DOMParser().parseFromString(html, 'text/html');
    // 移除脚本和样式
    doc.querySelectorAll('script, style, noscript').forEach(el => el.remove());
    // 将块级元素边界转为换行，保留表格结构
    doc.querySelectorAll('tr').forEach(tr => tr.insertAdjacentText('beforeend', '\n'));
    doc.querySelectorAll('td, th').forEach(cell => cell.insertAdjacentText('beforeend', ' | '));
    doc.querySelectorAll('p, div, li, h1, h2, h3, h4, h5, h6, br').forEach(el => el.insertAdjacentText('beforeend', '\n'));
    return (doc.body.textContent || '').replace(/[ \t]+/g, ' ').replace(/\n{3,}/g, '\n\n').trim();
}

async function extractPdfText(file) {
    const buf = await file.arrayBuffer();
    pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
    const pdf = await pdfjsLib.getDocument({ data: buf }).promise;
    let text = '';
    for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i);
        const content = await page.getTextContent();
        text += content.items.map(item => item.str).join(' ') + '\n';
    }
    return text.trim();
}

async function extractDocxText(file) {
    const buf = await file.arrayBuffer();
    const result = await mammoth.extractRawText({ arrayBuffer: buf });
    return result.value;
}

// ========== AI 平台预设与连接测试 ==========
const AI_PRESETS = {
    deepseek: {
        base: 'https://api.deepseek.com',
        model: 'deepseek-chat',
        hint: 'DeepSeek 官方接口，按量计费（需充值）。注册地址：platform.deepseek.com'
    },
    siliconflow: {
        base: 'https://api.siliconflow.cn/v1',
        model: 'deepseek-ai/DeepSeek-V3',
        hint: '🆓 免费方案：访问 cloud.siliconflow.cn 注册（送免费额度）→ 账号管理 → API密钥 → 创建并粘贴到上方 Key 框。DeepSeek-V3 在免费档，不花钱即可分析'
    },
    ollama: {
        base: 'http://localhost:11434/v1',
        model: 'qwen2.5:7b',
        hint: '🏠 完全免费、无需注册、无需 Key：① 安装 Ollama（ollama.com/download）；② 命令行执行 ollama pull qwen2.5:7b 下载模型（约4.7GB，电脑内存需≥8GB）；③ 回到本页面直接测试使用。模型越大效果越好，可选 deepseek-r1:14b 等'
    },
    zhipu: {
        base: 'https://open.bigmodel.cn/api/paas/v4',
        model: 'glm-4-flash',
        hint: '智谱开放平台：访问 open.bigmodel.cn 注册，GLM Flash 系列为免费档。注册后在"API密钥"页创建 Key'
    },
    openrouter: {
        base: 'https://openrouter.ai/api/v1',
        model: 'deepseek/deepseek-chat-v3.1:free',
        hint: 'OpenRouter：访问 openrouter.ai 注册，选择带 :free 后缀的免费模型（国内访问可能较慢）'
    },
    custom: {
        base: '',
        model: '',
        hint: '自定义：填写任意 OpenAI 兼容接口地址与模型名'
    }
};

function applyProvider() {
    const key = document.getElementById('aiProvider').value;
    const p = AI_PRESETS[key];
    if (!p) return;
    if (p.base) document.getElementById('apiBase').value = p.base;
    if (p.model) document.getElementById('modelName').value = p.model;
    document.getElementById('providerHint').textContent = p.hint;
    document.getElementById('connStatus').textContent = '';
}

async function testConnection() {
    const statusEl = document.getElementById('connStatus');
    const apiKey = document.getElementById('apiKey').value.trim();
    const apiBase = document.getElementById('apiBase').value.trim();
    const isLocal = /localhost|127\.0\.0\.1/.test(apiBase);
    if (!apiKey && !isLocal) {
        statusEl.textContent = '⚠ 请先填写 API Key（或选择本地 Ollama 方案）';
        statusEl.style.color = 'var(--danger)';
        return;
    }
    statusEl.textContent = '🔍 正在测试连接...';
    statusEl.style.color = 'var(--text-secondary)';
    try {
        const reply = await callDeepSeek('你是一个助手，请用一个字回答。', '你好', 16);
        statusEl.textContent = `✅ 连接成功！模型已响应（${document.getElementById('modelName').value.trim()}），可以开始分析`;
        statusEl.style.color = 'var(--success)';
    } catch (err) {
        let msg = err.message;
        if (/402|Insufficient|余额/.test(msg)) msg = '余额不足，请换一个免费平台（如硅基流动）';
        else if (/401|Unauthorized|Authentication/i.test(msg)) msg = 'API Key 无效，请检查是否复制完整';
        else if (/404|model.*not.*found/i.test(msg)) msg = '模型名不存在，请检查模型名称填写';
        statusEl.textContent = `❌ 连接失败：${msg}`;
        statusEl.style.color = 'var(--danger)';
    }
}

// 初始化：默认选中硅基流动预设并显示提示
applyProvider();

// ========== DeepSeek API 调用 ==========
async function callDeepSeek(systemPrompt, userPrompt, maxTokens = 4096) {
    const apiKey = document.getElementById('apiKey').value.trim();
    const apiBase = document.getElementById('apiBase').value.trim();
    const modelName = document.getElementById('modelName').value.trim() || 'deepseek-chat';
    const isLocal = /localhost|127\.0\.0\.1/.test(apiBase);
    if (!apiKey && !isLocal) throw new Error('请先填写 API Key（或选择本地 Ollama 方案）');
    const authKey = apiKey || 'ollama'; // 本地服务无需真实密钥

    // 兼容带/不带 /v1 后缀的 API 地址
    const base = apiBase.replace(/\/+$/, '');
    const endpoint = /\/v\d+$/.test(base) ? `${base}/chat/completions` : `${base}/v1/chat/completions`;

    const resp = await fetch(endpoint, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${authKey}`
        },
        body: JSON.stringify({
            model: modelName,
            messages: [
                { role: 'system', content: systemPrompt },
                { role: 'user', content: userPrompt }
            ],
            max_tokens: maxTokens,
            temperature: 0.3,
            stream: false
        })
    });

    if (!resp.ok) {
        const errText = await resp.text();
        throw new Error(`API 请求失败 (${resp.status}): ${errText}`);
    }

    const data = await resp.json();
    return data.choices[0].message.content;
}

// ========== 截断文本 ==========
function truncate(text, maxLen) {
    if (!text) return '(无内容)';
    return text.length > maxLen ? text.substring(0, maxLen) + '...(已截断)' : text;
}

// ========== 更新进度 ==========
function setStep(stepNum, status) {
    const el = document.getElementById(`step-${stepNum}`);
    el.className = status;
    const icon = el.querySelector('.step-icon');
    if (status === 'active') icon.textContent = '⏳';
    else if (status === 'done') icon.textContent = '✅';
    else icon.textContent = '○';
    const progress = Math.min(Math.round((stepNum / 6) * 100), 100);
    document.getElementById('progressBar').style.width = progress + '%';
}

// ========== 主分析流程 ==========
async function startAnalysis() {
    const candidateName = document.getElementById('candidateName').value.trim() || '候选人';

    // 验证岗位选择与文件
    const missing = [];
    if (!state.texts.profile) missing.push('应聘岗位（请在画像库中选择）');
    const labels = { resume: '简历', gpi: 'GPI个性测验报告', assessment: '职能支持岗综合测评报告' };
    for (const [key, label] of Object.entries(labels)) {
        if (!state.texts[key]) missing.push(label);
    }
    if (missing.length > 0) {
        showToast(`请先上传: ${missing.join('、')}`, 'error');
        return;
    }

    // 禁用按钮，显示进度
    const btn = document.getElementById('btnGenerate');
    btn.disabled = true;
    document.getElementById('progressSection').classList.add('active');
    document.getElementById('resultsSection').classList.remove('active');
    for (let i = 1; i <= 6; i++) setStep(i, '');

    try {
        // Step 1: 解析文件
        setStep(1, 'active');
        await sleep(500);
        setStep(1, 'done');

        // Step 2: 简历与岗位匹配度分析
        setStep(2, 'active');
        const resumeMatchResult = await analyzeResumeMatch(candidateName);
        setStep(2, 'done');

        // Step 3: 测评报告三维度分析
        setStep(3, 'active');
        const assessmentResult = await analyzeAssessment(candidateName);
        setStep(3, 'done');

        // Step 4: 生成综合报告
        setStep(4, 'active');
        const fullReport = await generateFullReport(candidateName, resumeMatchResult, assessmentResult);
        state.reportText = fullReport;
        // 提取分数
        const scoreMatch = fullReport.match(/综合匹配度评分[：:\s]*(\d{1,3})\s*分/) || fullReport.match(/综合评分[：:\s]*(\d{1,3})\s*分/);
        if (scoreMatch) {
            state.score = parseInt(scoreMatch[1]);
        } else {
            const scoreMatch2 = fullReport.match(/(\d{1,3})\s*\/\s*100/);
            state.score = scoreMatch2 ? parseInt(scoreMatch2[1]) : 75;
        }
        setStep(4, 'done');

        // Step 5: 设计面试题
        setStep(5, 'active');
        const interviewQuestions = await designInterviewQuestions(candidateName, fullReport);
        state.interviewText = interviewQuestions;
        setStep(5, 'done');

        // Step 6: 完成
        setStep(6, 'done');
        document.getElementById('progressBar').style.width = '100%';

        // 显示结果
        displayResults();

        // 💾 保存到历史记录
        const historyRecord = {
            id: Date.now(),
            name: candidateName,
            position: state.selectedPosition || '未选择',
            score: state.score,
            reportText: state.reportText,
            interviewText: state.interviewText,
            time: new Date().toLocaleString('zh-CN')
        };
        addToHistory(historyRecord);

        showToast('🎉 分析完成！已自动保存到历史记录', 'success');

    } catch (err) {
        showToast(`分析失败: ${err.message}`, 'error');
        console.error(err);
    } finally {
        btn.disabled = false;
    }
}

// ========== 分析函数 ==========
async function analyzeResumeMatch(name) {
    const position = state.selectedPosition || '目标岗位';
    const systemPrompt = `你是一位资深的人力资源专家，擅长通过简历分析候选人与岗位的匹配度。请从简历角度，结合岗位人才画像，详细分析候选人的优势与劣势。`;

    const userPrompt = `请分析以下候选人与目标岗位的匹配度。

【岗位人才画像 - ${position}】
${truncate(state.texts.profile, 5000)}

【候选人简历 - ${name}】
${truncate(state.texts.resume, 5000)}

请从以下方面进行分析：
1. 教育背景匹配度
2. 工作经验匹配度（行业、岗位年限、职责相关性）
3. 专业技能匹配度
4. 项目经历/业绩亮点
5. 简历中体现的优势（至少3点）
6. 简历中体现的劣势/风险点（至少3点）

请详细、客观地分析，输出结构化内容。`;

    return await callDeepSeek(systemPrompt, userPrompt, 3000);
}

async function analyzeAssessment(name) {
    const systemPrompt = `你是一位专业的心理测评分析师，擅长解读北森GPI个性测验和职能支持岗综合测评报告。请按"成长的基础"、"成长的潜力"、"成长的渴望"三个维度综合分析候选人与鸿萌橡塑的匹配度。`;

    const userPrompt = `请根据以下两份测评报告，按三个维度分析候选人 ${name} 与鸿萌橡塑的匹配度。

【北森GPI个性测验测评报告】
${truncate(state.texts.gpi, 5000)}

【职能支持岗综合测评报告】
${truncate(state.texts.assessment, 5000)}

请按以下三个维度进行分析：

一、成长的基础
- 候选人的基础知识储备、基本能力素质
- 个性特征中与岗位匹配的部分
- 测评数据支撑

二、成长的潜力
- 学习能力、适应能力、发展潜能
- 测评中体现的高潜力指标
- 可能的成长方向

三、成长的渴望
- 内在驱动力、成就动机
- 职业兴趣与岗位契合度
- 价值观与企业文化匹配度

每个维度请给出具体分析和匹配度评价（高/中/低）。`;

    return await callDeepSeek(systemPrompt, userPrompt, 3000);
}

async function generateFullReport(name, resumeResult, assessmentResult) {
    const position = state.selectedPosition || '目标岗位';
    const systemPrompt = `你是一位资深高级人力资源顾问，负责为鸿萌橡塑撰写专业的候选人综合岗位匹配度分析报告。
核心要求：报告必须严格遵循以下章节结构和格式。大量使用markdown表格呈现信息，使内容清晰、美观、一目了然。客观全面，有数据支撑，最终给出100分制综合评分。
绝对禁止：不得使用XX、xxx、占位符等任何敷衍文字，所有内容必须用真实数据填写。`;

    const userPrompt = `请为候选人 ${name}（应聘岗位：${position}）生成一份完整的综合岗位匹配度报告。

【岗位人才画像】
${truncate(state.texts.profile, 3000)}

【简历匹配度分析】
${resumeResult}

【测评三维度分析】
${assessmentResult}

请严格按照以下格式生成报告。极其重要：
1. 所有表格必须使用markdown表格格式（用 | 作为列分隔符，第二行用 |---|---| 作为分隔行）
2. 禁止使用XX、xxx、占位符等任何敷衍文字，所有内容必须用简历和测评中的真实数据填写
3. 禁止使用 # 或 ## 或 ### 等markdown标题符号

综合匹配度评分：72分/100分

一、评估结论总览

| 评估维度 | 得分(满分100) | 权重 | 加权得分 |
|---------|-------------|------|----------|
| 专业硬技能匹配度 | 78 | 40% | 31.2 |
| 软素质与岗位特质匹配度 | 65 | 30% | 19.5 |
| 成长潜力与发展适配度 | 72 | 20% | 14.4 |
| 文化价值观与动机匹配度 | 70 | 10% | 7.0 |
| 综合得分 | | | 72分 |

二、与JD各项要求的逐条匹配分析

☑ 完全匹配项

| JD要求 | 简历对应 | 匹配情况 |
|--------|---------|----------|
| 本科及以上学历，会计学/财务管理专业 | 南京航空航天大学金城学院，会计学本科 | 完全符合 |
| 8年以上制造业财务经验 | 11年财务经验，其中近7年在制造业上市公司 | 超出要求 |

部分匹配/待验证项

| JD要求 | 简历对应 | 分析 |
|--------|---------|------|
| 熟练使用用友NC | 简历写"熟练操作金蝶、用友、浪潮" | 用友NC是集团版产品，需确认具体使用版本 |

明显缺失项

| JD要求 | 简历对应 | 分析 |
|--------|---------|------|
| 编制集团合并财务报表及附注 | 简历未提及合并报表经验 | 最大短板，下辖8家子公司合并报表是核心职责 |

三、核心优势

1. 上市公司IPO全流程经验
王云芳在菲沃泰任职期间深度参与公司科创板IPO项目，涵盖核算规范化、研发合规专项、税务风险防控等核心环节。贵司规划三年内上市，这一经验极具价值，能直接迁移到上市筹备工作中。

2. 研发费用精细化管理能力
主导研发费用归集、核算及追踪评估，建立全流程资料归档机制，完成研发费用加计扣除合规备案。贵司作为高新技术企业，研发费用管理是税务合规的核心。

四、主要劣势与风险

1. 无合并报表经验
贵司下辖8家子公司，JD明确要求编制集团合并财务报表及附注。王云芳的简历中完全没有合并报表相关经历，这是最需要关注的风险点。

2. 无团队管理经验
JD要求指导总账、税务、应收应付会计工作，王云芳过往均为独立岗位，未体现带团队或指导他人的经历。

四、综合人岗匹配诊断

| 成长维度 | 匹配度评分(满分10) | 核心结论 |
|---------|------------------|----------|
| 成长的基础 | 7.0/10 | 专业硬实力扎实，但系统思维和影响力缺失 |
| 成长的潜力 | 4.5/10 | 学习具体知识尚可，但极度悲观、缺乏自信 |
| 成长的渴望 | 4.0/10 | 安于现状，回避竞争，缺乏挑战高目标的欲望 |

五、与同岗位其他候选人的差异化竞争力

| 对比维度 | 该候选人 | 一般同业候选人 | 优势/短板 |
|---------|---------|--------------|----------|
| IPO经验 | 完整经历（菲沃泰科创板） | 多数没有 | 稀缺性高 |
| 上市公司财务规范 | 7年（上市前后全程） | 多数没有 | 经验质量高 |
| 合并报表 | 空白 | 部分有 | 短板明显 |
| 管理经验 | 空白 | 部分有 | 短板明显 |

六、面试阶段需进一步验证的问题

| 验证项目 | 重要性 | 面试问题(建议STAR法则追问) |
|---------|-------|---------------------------|
| 合并报表知识 | 高 | 假设集团内母子公司有未实现内部交易损益，合并抵消时如何处理？ |
| 用友NC经验 | 高 | 请具体描述你在用友NC中完成过哪些模块的配置和操作？ |
| 团队管理潜力 | 中 | 假设你的一名下属连续3个月做错同一项费用分摊，你会如何沟通和纠正？ |
| 抗压能力 | 中 | 请分享一次你在工作中面临巨大压力或几乎想放弃的经历 | |

七、录用建议

三种可行方案：

| 方案 | 岗位定位 | 匹配度评估 |
|------|---------|------------|
| 方案一：直接录用为会计主管（培养型） | 会计主管，前3-6个月由上级带教合并报表 | 72分 |
| 方案二：录用为高级总账会计 | 高级总账会计/税务专家，不设直接下属 | 85分 |
| 方案三：继续寻访 | 继续寻找更匹配的候选人 | - |

综合判断：王云芳的专业硬实力（IPO经验、税务合规、ERP上线）与岗位要求高度契合，但软素质（合并报表空白、无管理经验）与岗位要求存在结构性差距。整体呈"偏科型"特征——长处突出、短板明显。建议以"高级总账会计"岗录用，匹配度约85分；如以"会计主管"岗直接录用，匹配度约72分，需接受3-6个月的培养过渡期。`;

    return await callDeepSeek(systemPrompt, userPrompt, 4096);
}

async function designInterviewQuestions(name, fullReport) {
    const systemPrompt = `你是一位资深面试官，擅长设计结构化面试提纲。题目针对性强、有深度，能有效验证候选人真实能力和潜在风险。输出必须以表格为主要呈现形式，整洁清晰、美观专业。
绝对禁止：不得使用XX、xxx、占位符等任何敷衍文字，所有内容必须针对该候选人的具体情况填写。`;

    const userPrompt = `请根据以下综合报告，为候选人 ${name} 设计一份结构化面试提纲，共约10道题。

【综合报告】
${fullReport}

设计要求：
1. 必须设计8-10道面试题，不得少于8道
2. 重点针对报告中的低分项、风险项和待验证项设计题目
3. 题目类型覆盖：专业能力验证（3-4题）、行为面试STAR（3-4题）、情景模拟（1-2题）、动机与价值观（1-2题）
4. 每题标注重要性（高/中/低）
5. 每题包含面试主问题 + 2-3个递进式追问要点

请严格按以下格式输出，核心内容必须以表格呈现：

${name} 面试提纲

一、面试问题总览

| 序号 | 验证项目 | 重要性 | 面试问题 |
|------|---------|--------|----------|
| 1 | 合并报表知识 | 高 | 假设集团内母子公司有未实现内部交易损益，合并抵消时如何处理？请讲出你的逻辑框架。 |
| 2 | 用友NC经验 | 高 | 请具体描述你在用友NC中完成过哪些模块的配置和操作？与用友U8的主要区别是什么？ |
| 3 | 团队管理潜力 | 中 | 假设你的一名下属连续3个月做错同一项费用分摊，你会如何沟通和纠正？请现场模拟。 |
| 4 | 抗压能力 | 中 | 请分享一次你在工作中面临巨大压力或几乎想放弃的经历，当时发生了什么？你做了什么？结局如何？ |
| 5 | 特殊会计处理 | 中 | 你是否处理过资产减值测试、递延所得税计算或政府补助确认？请举例说明。 |
| 6 | 成就动机 | 低 | 如果公司上市因外部环境延期2年，你个人的职业规划会如何调整？ |
| 7 | 跨部门协作 | 中 | 请描述一次你与业务部门就财务数据产生分歧的经历，你是如何处理的？ |
| 8 | 流程优化能力 | 中 | 你在过往工作中主导过哪些财务流程的优化？具体做了什么？效果如何？ |
| 9 | 税务风险防控 | 高 | 请分享一次你发现并处理税务风险的经历，你是如何识别和应对的？ |
| 10 | 职业规划 | 低 | 未来3-5年你的职业发展目标是什么？你希望在这个岗位上获得什么？ |

二、各题详细追问要点

第1题  合并报表知识  重要性：高
追问要点：
· 请描述你处理过的最复杂的合并报表场景
· 内部往来对账出现差异时，你的处理流程是什么？
· 少数股东权益的计算和列报你是否有实操经验？
考察重点：是否具备集团合并报表的实操能力

第2题  用友NC经验  重要性：高
追问要点：
· 你在NC中主要使用哪些模块？（总账、应收应付、固定资产等）
· NC与U8在集团化管理方面有哪些关键区别？
· 你是否参与过NC系统的初始化或流程配置？
考察重点：是否真正使用过用友NC集团版系统

第3题  团队管理潜力  重要性：中
追问要点：
· 你是否有过指导新人的经历？具体是怎么做的？
· 如果团队成员不配合你的工作安排，你会怎么处理？
· 你如何评估一个下属的工作质量？
考察重点：是否具备带团队和管理他人的潜力

第4题  抗压能力  重要性：中
追问要点：
· 那次压力经历中，最困难的时刻是什么？你是怎么度过的？
· 你平时如何调节工作压力？有什么具体的方法？
· 如果同时面对多个紧急任务，你会如何安排优先级？
考察重点：面对高压环境时的心理韧性和应对策略

第5题  特殊会计处理  重要性：中
追问要点：
· 资产减值测试的具体步骤是什么？你做过哪些资产的减值测试？
· 递延所得税资产和负债的确认条件是什么？请举例说明
· 政府补助的会计处理方法有哪些？你处理过哪种类型？
考察重点：对复杂会计准则的理解和实操能力

第6题  成就动机  重要性：低
追问要点：
· 如果上市延期，你会如何调整自己的工作重点？
· 你过去是否遇到过计划被打乱的情况？你是如何应对的？
· 你个人的职业成就感主要来自哪里？
考察重点：内在驱动力和职业稳定性

第7题  跨部门协作  重要性：中
追问要点：
· 那次分歧的核心问题是什么？你采取了什么沟通策略？
· 最终结果如何？你从中学到了什么？
· 你如何平衡财务合规要求和业务部门的效率需求？
考察重点：沟通协调能力和跨部门合作经验

第8题  流程优化能力  重要性：中
追问要点：
· 你发现流程问题的切入点是什么？
· 优化过程中遇到了哪些阻力？你是如何克服的？
· 优化后的效果如何量化？
考察重点：系统思维和持续改进能力

第9题  税务风险防控  重要性：高
追问要点：
· 你是如何发现那个税务风险的？
· 你采取了哪些具体措施来化解风险？
· 事后你是否建立了预防机制？具体是什么？
考察重点：税务合规意识和风险防控能力

第10题  职业规划  重要性：低
追问要点：
· 你为什么选择应聘这个岗位？
· 你希望在这个岗位上实现什么目标？
· 你如何看待这个岗位与你长期职业目标的关系？
考察重点：职业目标清晰度和岗位匹配意愿

三、面试评分参考

| 评估维度 | 优秀表现 | 警示信号 |
|---------|---------|----------|
| 专业能力 | 能清晰描述处理流程和逻辑框架，有具体案例支撑 | 回答模糊、缺乏具体案例、概念混淆 |
| 管理潜力 | 能举例说明如何指导和培养下属，有管理思维 | 回避管理话题、缺乏团队意识、只关注个人工作 |
| 抗压能力 | 能客观分析压力来源并给出解决方案，情绪稳定 | 情绪化、归因于外部因素、无法给出具体应对方法 |
| 动机匹配 | 职业目标清晰，与岗位发展路径一致 | 目标模糊、对岗位了解不足、缺乏长期规划 |

注意：禁止使用 # 或 ## 或 ### 等markdown标题符号，用中文数字编号即可。`;

    return await callDeepSeek(systemPrompt, userPrompt, 4096);
}

// ========== 清除markdown格式符号 ==========
function cleanMarkdown(text) {
    if (!text) return '';
    return text
        .replace(/^#{1,6}\s*/gm, '')       // 移除 # ## ### 等标题符号
        .replace(/\*\*(.+?)\*\*/g, '$1')    // 移除 **加粗** 保留文字
        .replace(/\*(.+?)\*/g, '$1')        // 移除 *斜体* 保留文字
        .replace(/={2,}/g, '')              // 移除 == 标记
        .replace(/^\s*[-]{3,}\s*$/gm, '')   // 移除 --- 分隔线
        .replace(/^\|[\s\-:|]+\|$/gm, '')   // 移除表格分隔行 |---|---|
        .replace(/\n{3,}/g, '\n\n');         // 压缩多余空行
}

// ========== 显示结果 ==========
function displayResults() {
    const section = document.getElementById('resultsSection');
    section.classList.add('active');

    // 评分
    const circle = document.getElementById('scoreCircle');
    circle.textContent = state.score;
    circle.className = 'score-circle ' + (state.score >= 80 ? 'score-high' : state.score >= 60 ? 'score-mid' : 'score-low');

    const label = document.getElementById('scoreLabel');
    if (state.score >= 80) label.textContent = '匹配度高 - 推荐进入下一轮';
    else if (state.score >= 60) label.textContent = '匹配度中等 - 建议进一步评估';
    else label.textContent = '匹配度偏低 - 谨慎考虑';

    // 报告内容（清除markdown符号后显示）
    document.getElementById('reportContent').textContent = cleanMarkdown(state.reportText);
    document.getElementById('interviewContent').textContent = cleanMarkdown(state.interviewText);

    // 滚动到结果
    section.scrollIntoView({ behavior: 'smooth' });
}

// ========== 下载报告 ==========
async function downloadReport(type, format = 'docx') {
    const candidateName = document.getElementById('candidateName').value.trim() || '候选人';
    let content, filename;

    if (type === 'match') {
        content = state.reportText;
        filename = `${candidateName}_综合岗位匹配度报告`;
    } else {
        content = state.interviewText;
        filename = `${candidateName}_面试提纲`;
    }

    if (format === 'docx') {
        if (typeof docx === 'undefined') {
            showToast('⚠ Word 导出库未加载，请检查网络后重试', 'error');
            return;
        }
        try {
            const blob = await generateDocx(content, filename);
            saveFile(blob, filename + '.docx');
            showToast(`📦 已下载: ${filename}.docx`, 'success');
        } catch (err) {
            showToast(`⚠ Word 导出失败: ${err.message}`, 'error');
            console.error(err);
        }
        return;
    }

    const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
    saveFile(blob, filename + '.txt');
    showToast(`📥 已下载: ${filename}.txt`, 'success');
}

function saveFile(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
}

// ========== Word (DOCX) 导出 ==========
async function generateDocx(text, title) {
    const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
            Footer, PageNumber, Table, TableRow, TableCell, WidthType, BorderStyle,
            VerticalAlign, ShadingType } = docx;
    const FONT = '微软雅黑';
    const lines = text.split(/\r?\n/);
    const children = [];
    let i = 0;
    let isTitleHandled = false;

    // ---- 辅助：解析 markdown 表格块 ----
    function parseTableBlock(tableLines) {
        const rows = [];
        for (const ln of tableLines) {
            const t = ln.trim();
            if (!t.startsWith('|')) continue;
            if (/^\|[\s\-:|]+\|$/.test(t)) continue;
            const cells = t.split('|').slice(1, -1).map(c => c.trim());
            if (cells.length > 0 && cells.some(c => c)) rows.push(cells);
        }
        return rows;
    }

    // ---- 辅助：创建表格行 ----
    function mkRow(cells, maxCols, isHeader, rowIdx) {
        const isEven = rowIdx % 2 === 0;
        return new TableRow({
            tableHeader: isHeader,
            children: cells.map(text => {
                const cleanText = (text || '').replace(/\*\*/g, '');
                return new TableCell({
                    verticalAlign: VerticalAlign.CENTER,
                    shading: isHeader
                        ? { type: ShadingType.SOLID, color: '2563EB', fill: '2563EB' }
                        : (!isHeader && isEven ? { type: ShadingType.SOLID, color: 'F2F7FB', fill: 'F2F7FB' } : undefined),
                    margins: { top: 50, bottom: 50, left: 100, right: 100 },
                    children: [new Paragraph({
                        spacing: { before: 40, after: 40 },
                        children: [new TextRun({
                            text: cleanText,
                            bold: isHeader,
                            size: 20, font: FONT,
                            color: isHeader ? 'FFFFFF' : '333333'
                        })]
                    })]
                });
            })
        });
    }

    // ---- 辅助：创建表格 ----
    function mkTable(dataRows) {
        if (!dataRows.length) return null;
        const maxCols = Math.max(...dataRows.map(r => r.length));
        const borderDef = { style: BorderStyle.SINGLE, size: 1, color: 'BFBFBF' };
        return new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            borders: { top: borderDef, bottom: borderDef, left: borderDef, right: borderDef, insideHorizontal: borderDef, insideVertical: borderDef },
            rows: dataRows.map((cells, idx) => mkRow(cells, maxCols, idx === 0, idx))
        });
    }

    // ---- 辅助：处理行内加粗 **text** ----
    function parseInlineRuns(text) {
        const runs = [];
        const re = /\*\*(.+?)\*\*/g;
        let last = 0, m;
        while ((m = re.exec(text)) !== null) {
            if (m.index > last) runs.push(new TextRun({ text: text.slice(last, m.index), size: 22, font: FONT }));
            runs.push(new TextRun({ text: m[1], bold: true, size: 22, font: FONT, color: '1d4ed8' }));
            last = re.lastIndex;
        }
        if (last < text.length) runs.push(new TextRun({ text: text.slice(last), size: 22, font: FONT }));
        return runs.length ? runs : [new TextRun({ text: text, size: 22, font: FONT })];
    }

    // ---- 主循环 ----
    while (i < lines.length) {
        const raw = lines[i];
        const line = raw.trimEnd();
        const trimmed = line.trim();

        // 空行跳过
        if (!trimmed) { i++; continue; }

        // 检测 markdown 表格块（连续 | 开头行）
        if (trimmed.startsWith('|')) {
            const tableLines = [];
            while (i < lines.length && lines[i].trim().startsWith('|')) {
                tableLines.push(lines[i]);
                i++;
            }
            const dataRows = parseTableBlock(tableLines);
            const tbl = mkTable(dataRows);
            if (tbl) {
                children.push(tbl);
                children.push(new Paragraph({ spacing: { before: 60, after: 60 }, children: [] }));
            }
            continue;
        }

        // 文档主标题
        if (!isTitleHandled) {
            isTitleHandled = true;
            children.push(new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 200, after: 100 },
                children: [new TextRun({ text: trimmed, bold: true, size: 36, font: FONT, color: '1d4ed8' })]
            }));
            i++; continue;
        }

        // 综合评分行 → 大号居中显示
        if (/综合匹配度评分/.test(trimmed) || /综合评分[：:]/.test(trimmed)) {
            const sm = trimmed.match(/(\d{1,3})\s*分/);
            if (sm) {
                const sc = parseInt(sm[1]);
                const clr = sc >= 80 ? '059669' : sc >= 60 ? 'd97706' : 'dc2626';
                children.push(new Paragraph({ spacing: { before: 200, after: 0 }, alignment: AlignmentType.CENTER, children: [] }));
                children.push(new Paragraph({
                    alignment: AlignmentType.CENTER,
                    spacing: { before: 0, after: 40 },
                    children: [new TextRun({ text: String(sc), bold: true, size: 64, font: FONT, color: clr })]
                }));
                children.push(new Paragraph({
                    alignment: AlignmentType.CENTER,
                    spacing: { before: 0, after: 200 },
                    children: [new TextRun({ text: '综合匹配度评分（满分100分）', size: 20, font: FONT, color: '666666' })]
                }));
            } else {
                children.push(new Paragraph({
                    alignment: AlignmentType.CENTER,
                    spacing: { before: 120, after: 120 },
                    children: [new TextRun({ text: trimmed, bold: true, size: 28, font: FONT, color: '1d4ed8' })]
                }));
            }
            i++; continue;
        }

        // 分隔线
        if (/^[=\-—_*]{3,}$/.test(trimmed)) {
            children.push(new Paragraph({
                border: { bottom: { color: 'CCCCCC', size: 4, style: 'single' } },
                spacing: { before: 100, after: 100 }
            }));
            i++; continue;
        }

        // 章节标题：一、二、三...
        if (/^[一二三四五六七八九十]+、/.test(trimmed)) {
            children.push(new Paragraph({
                heading: HeadingLevel.HEADING_1,
                spacing: { before: 300, after: 120 },
                border: { bottom: { color: '2563EB', size: 6, style: 'single' } },
                children: [new TextRun({ text: trimmed, bold: true, size: 28, font: FONT, color: '1d4ed8' })]
            }));
            i++; continue;
        }

        // 子标题：第N题 / ☑/✗ 小标题 / 重要性标注行
        if (/^第\d+题/.test(trimmed) || /^[☑✓✗✘×]/.test(trimmed) || /^【.+】$/.test(trimmed) && trimmed.length <= 30) {
            children.push(new Paragraph({
                heading: HeadingLevel.HEADING_2,
                spacing: { before: 200, after: 80 },
                children: [new TextRun({ text: trimmed, bold: true, size: 26, font: FONT, color: '2563eb' })]
            }));
            i++; continue;
        }

        // 重要性行 ★
        if (/重要性[：:]/.test(trimmed)) {
            children.push(new Paragraph({
                spacing: { before: 80, after: 40 },
                children: [new TextRun({ text: trimmed, bold: true, size: 22, font: FONT, color: 'd97706' })]
            }));
            i++; continue;
        }

        // ☑/✗ 开头行 → 彩色显示
        if (/^[☑✓]/.test(trimmed)) {
            children.push(new Paragraph({
                spacing: { before: 120, after: 60 },
                children: [new TextRun({ text: trimmed, bold: true, size: 24, font: FONT, color: '059669' })]
            }));
            i++; continue;
        }
        if (/^[✗✘×]/.test(trimmed)) {
            children.push(new Paragraph({
                spacing: { before: 120, after: 60 },
                children: [new TextRun({ text: trimmed, bold: true, size: 24, font: FONT, color: 'dc2626' })]
            }));
            i++; continue;
        }

        // 数字编号短行 → 加粗小标题
        if (/^\d+[.、]\s*/.test(trimmed) && trimmed.length <= 50) {
            children.push(new Paragraph({
                spacing: { before: 120, after: 60 },
                children: [new TextRun({ text: trimmed, bold: true, size: 24, font: FONT })]
            }));
            i++; continue;
        }

        // ■ 开头
        if (trimmed.startsWith('■')) {
            children.push(new Paragraph({
                spacing: { before: 80, after: 40 },
                children: [new TextRun({ text: trimmed.replace(/^■\s*/, ''), bold: true, size: 24, font: FONT, color: '2563eb' })]
            }));
            i++; continue;
        }

        // · / • / - 列表项
        if (/^[·•\-\*]\s*/.test(trimmed) || /^\s*\d+[.)、]/.test(line)) {
            children.push(new Paragraph({
                indent: { left: 480 },
                spacing: { before: 40, after: 40 },
                children: parseInlineRuns(trimmed)
            }));
            i++; continue;
        }

        // 普通段落（支持 **加粗** 解析）
        children.push(new Paragraph({
            spacing: { before: 60, after: 60 },
            children: parseInlineRuns(trimmed)
        }));
        i++;
    }

    // ---- 构建文档 ----
    const now = new Date();
    const dateStr = `${now.getFullYear()}年${now.getMonth() + 1}月${now.getDate()}日`;

    const doc = new Document({
        creator: '鸿萌橡塑人才评估系统',
        title: title,
        description: '候选人岗位匹配度分析报告',
        styles: {
            default: {
                document: { run: { font: FONT, size: 22 } }
            }
        },
        sections: [{
            properties: {
                page: { margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } }
            },
            footers: {
                default: new Footer({
                    children: [new Paragraph({
                        alignment: AlignmentType.CENTER,
                        children: [
                            new TextRun({ text: `鸿萌橡塑 · 生成于 ${dateStr} · 第 `, size: 18, font: FONT, color: '888888' }),
                            new TextRun({ children: [PageNumber.CURRENT], size: 18, font: FONT, color: '888888' }),
                            new TextRun({ text: ' 页', size: 18, font: FONT, color: '888888' })
                        ]
                    })]
                })
            },
            children: children
        }]
    });

    return await Packer.toBlob(doc);
}

// ========== 工具函数 ==========
function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

function showToast(msg, type) {
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    toast.textContent = msg;
    document.body.appendChild(toast);
    setTimeout(() => toast.remove(), 3000);
}
