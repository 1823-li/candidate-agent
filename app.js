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
            '⚠️ 默认画像文件加载失败（' + e.message + '），请点击“导入/更新画像文件”手动上传 .xlsx';
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
        `当前画像库：<b>${names.length}</b> 个岗位　|　来源：${meta.source || '默认文件'}　|　更新时间：${meta.time || '-'}`;

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
    showToast('↩ 已恢复为默认画像文件', 'success');
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
        hint: '智谱开放平台：访问 open.bigmodel.cn 注册，GLM Flash 系列为免费档。注册后在“API密钥”页创建 Key'
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
        statusEl.textContent = '❌ 请先填写 API Key（或选择本地 Ollama 方案）';
        statusEl.style.color = 'var(--danger)';
        return;
    }
    statusEl.textContent = '⏳ 正在测试连接...';
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
    if (status === 'active') icon.textContent = '◉';
    else if (status === 'done') icon.textContent = '✓';
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
        const scoreMatch = fullReport.match(/综合评分[：:\s]*(\d{1,3})\s*分/);
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
        showToast('🎉 分析完成！', 'success');

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
    const systemPrompt = `你是一位高级人力资源顾问，负责为鸿萌橡塑撰写专业的候选人综合岗位匹配度报告。报告需要客观、全面、有数据支撑，最终给出100分制的综合评分。`;

    const userPrompt = `请为候选人 ${name} 生成一份完整的综合岗位匹配度报告。

【简历匹配度分析】
${resumeResult}

【测评三维度分析】
${assessmentResult}

请生成一份完整的综合报告，格式如下：

========================================
${name} 综合岗位匹配度报告
鸿萌橡塑
========================================

一、基本信息概述
（简要概括候选人背景）

二、简历维度分析
（基于简历的匹配度分析，含优势和劣势）

三、测评维度分析
1. 成长的基础
2. 成长的潜力  
3. 成长的渴望
（每个维度的详细分析）

四、综合优势（至少5点，每点具体说明）

五、综合劣势/风险项（至少5点，每点具体说明）

六、综合评分
综合评分：XX分（满分100分）
评分依据：
- 简历匹配度：XX分（权重30%）
- 成长的基础：XX分（权重25%）
- 成长的潜力：XX分（权重25%）
- 成长的渴望：XX分（权重20%）

七、总体建议
（是否推荐进入下一轮，及理由）

请确保评分客观合理，优势和劣势分析具体详实。`;

    return await callDeepSeek(systemPrompt, userPrompt, 4096);
}

async function designInterviewQuestions(name, fullReport) {
    const systemPrompt = `你是一位资深面试官，擅长设计结构化面试题目来验证候选人的真实能力和潜在风险。题目要针对性强、有深度，能有效区分候选人水平。`;

    const userPrompt = `请根据以下综合报告，为候选人 ${name} 设计8-10道结构化面试题。

【综合报告】
${fullReport}

设计要求：
1. 重点针对报告中的低分项和风险项设计题目
2. 每道题从以下三个方面展开：
   - 题目：具体的面试问题
   - 追问：2-3个递进式追问
   - 考察点：该题目要验证的核心能力/素质

3. 题目类型应覆盖：
   - 专业能力验证题（2-3题）
   - 行为面试题（STAR法）（2-3题）
   - 情景模拟题（1-2题）
   - 动机与价值观题（1-2题）

请按以下格式输出：

========================================
${name} 结构化面试题
========================================

第1题：[题目类型]
■ 题目：...
■ 追问：
  1. ...
  2. ...
  3. ...
■ 考察点：...

（依次列出所有题目）

========================================
面试评分参考标准
========================================
（简要说明各题目的评分要点）`;

    return await callDeepSeek(systemPrompt, userPrompt, 4096);
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

    // 报告内容
    document.getElementById('reportContent').textContent = state.reportText;
    document.getElementById('interviewContent').textContent = state.interviewText;

    // 滚动到结果
    section.scrollIntoView({ behavior: 'smooth' });
}

// ========== 下载报告 ==========
async function downloadReport(type, format = 'txt') {
    const candidateName = document.getElementById('candidateName').value.trim() || '候选人';
    let content, filename;

    if (type === 'match') {
        content = state.reportText;
        filename = `${candidateName}_综合岗位匹配度报告`;
    } else {
        content = state.interviewText;
        filename = `${candidateName}_结构化面试题`;
    }

    if (format === 'docx') {
        if (typeof docx === 'undefined') {
            showToast('❌ Word 导出库未加载，请检查网络后重试', 'error');
            return;
        }
        try {
            const blob = await generateDocx(content, filename);
            saveFile(blob, filename + '.docx');
            showToast(`📥 已下载: ${filename}.docx`, 'success');
        } catch (err) {
            showToast(`❌ Word 导出失败: ${err.message}`, 'error');
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
    const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Footer, PageNumber } = docx;
    const FONT = '微软雅黑';
    const lines = text.split(/\r?\n/);
    const children = [];
    let isTitleHandled = false;

    for (const raw of lines) {
        const line = raw.trimEnd();
        const trimmed = line.trim();
        if (!trimmed) continue;

        // 分隔线（==== 或 ----）跳过
        if (/^[=\-—_*]{3,}$/.test(trimmed)) {
            children.push(new Paragraph({
                border: { bottom: { color: '999999', size: 6, style: 'single' } },
                spacing: { before: 100, after: 100 }
            }));
            continue;
        }

        // 文档主标题（第一行或包含“报告”“面试题”且较短的行）
        if (!isTitleHandled) {
            isTitleHandled = true;
            children.push(new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 200, after: 200 },
                children: [new TextRun({ text: trimmed, bold: true, size: 36, font: FONT })]
            }));
            continue;
        }

        // 章节标题：一、二、三... 或 第N题 或 【...】
        if (/^[一二三四五六七八九十]+、/.test(trimmed) || /^第\d+题/.test(trimmed) || (/^【.+】$/.test(trimmed) && trimmed.length <= 30)) {
            children.push(new Paragraph({
                heading: HeadingLevel.HEADING_2,
                spacing: { before: 240, after: 120 },
                children: [new TextRun({ text: trimmed, bold: true, size: 28, font: FONT, color: '1d4ed8' })]
            }));
            continue;
        }

        // 小标题：数字编号的短行（如 “1. 教育背景匹配度”）或 ■ 开头的行
        if (/^\d+[.、]\s*/.test(trimmed) && trimmed.length <= 40) {
            children.push(new Paragraph({
                spacing: { before: 120, after: 60 },
                children: [new TextRun({ text: trimmed, bold: true, size: 24, font: FONT })]
            }));
            continue;
        }

        if (trimmed.startsWith('■')) {
            children.push(new Paragraph({
                spacing: { before: 80, after: 40 },
                children: [new TextRun({ text: trimmed.replace(/^■\s*/, ''), bold: true, size: 24, font: FONT, color: '2563eb' })]
            }));
            continue;
        }

        // 列表项：- / • / 缩进数字
        if (/^[-•·]\s*/.test(trimmed) || /^\s*\d+[.)、]/.test(line)) {
            children.push(new Paragraph({
                indent: { left: 480 },
                spacing: { before: 40, after: 40 },
                children: [new TextRun({ text: trimmed, size: 22, font: FONT })]
            }));
            continue;
        }

        // 普通段落
        children.push(new Paragraph({
            spacing: { before: 60, after: 60 },
            children: [new TextRun({ text: trimmed, size: 22, font: FONT })]
        }));
    }

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
