// ==UserScript==
// @name         Word Tool - 选词发音 & 翻译（Edge/Chrome）
// @namespace    simpread/word-tool
// @version      4.6.0
// @description  双击单词/技术标识符发音（_/-/camelCase 自动拆词顺序朗读 + 子词逐行释义 + 有道真人音频自动响度归一化 + TTS兜底，CSP兼容）+ ALT拖动选词翻译 + 整页翻译 + 输入框连续3次空格翻译成英文
// @match        *://*/*
// @run-at       document-idle
// @sandbox      DOM
// @inject-into  content
// @grant        GM_xmlhttpRequest
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_registerMenuCommand
// @connect      dict.youdao.com
// @connect      *.youdao.com
// @supportURL   https://github.com/Hulkhao/Word-Tool/issues/new
// @downloadURL  https://raw.githubusercontent.com/Hulkhao/Word-Tool/main/Word-Tool.user.js
// @updateURL    https://raw.githubusercontent.com/Hulkhao/Word-Tool/main/Word-Tool.user.js
// ==/UserScript==

(function () {
  'use strict';

  // 防止 iframe 重复注入
  if (window !== window.top) return;

  const CONFIG = {
    triggerMode: 'dblclick',
    youdaoType: 2,
    debounceMs: 120,
    targetLanguage: 'zh-CN'
  };

  function loadConfig() {
    CONFIG.triggerMode = GM_getValue('wt-trigger-mode', CONFIG.triggerMode);
    CONFIG.youdaoType = Number(GM_getValue('wt-youdao-type', CONFIG.youdaoType));
  }

  function saveConfig() {
    GM_setValue('wt-trigger-mode', CONFIG.triggerMode);
    GM_setValue('wt-youdao-type', CONFIG.youdaoType);
  }

  // ==================== 浏览器 / Translator API ====================

  function getPageWindow() {
    try {
      if (typeof unsafeWindow !== 'undefined' && unsafeWindow) return unsafeWindow;
    } catch {}
    return window;
  }

  function getTranslatorAPI() {
    const candidates = [];
    try { candidates.push(globalThis.Translator); } catch {}
    try { candidates.push(window.Translator); } catch {}
    try { candidates.push(getPageWindow().Translator); } catch {}

    return candidates.find(api =>
      api && typeof api.availability === 'function' && typeof api.create === 'function'
    ) || null;
  }

  function browserName() {
    const ua = navigator.userAgent;
    if (/Edg\//.test(ua)) return 'Edge';
    if (/Chrome\//.test(ua)) return 'Chrome';
    return '当前浏览器';
  }

  const translators = new Map();
  const translatorPromises = new Map();

  async function getTranslator(sourceLang, targetLang = CONFIG.targetLanguage) {
    const cacheKey = `${sourceLang}->${targetLang}`;
    if (translators.has(cacheKey)) return translators.get(cacheKey);
    if (translatorPromises.has(cacheKey)) return translatorPromises.get(cacheKey);

    const promise = (async () => {
      const TranslatorAPI = getTranslatorAPI();
      if (!TranslatorAPI) {
        throw new Error(
          `${browserName()} 未向当前 Userscript 暴露 Translator API。Edge 需要 148+，并建议使用最新版 Tampermonkey。`
        );
      }

      const opts = {
        sourceLanguage: sourceLang,
        targetLanguage: targetLang
      };

      const availability = await TranslatorAPI.availability(opts);
      console.log('[Word Tool] Translator availability:', sourceLang, '->', targetLang, availability);

      if (availability === 'unavailable') {
        throw new Error(`不支持 ${sourceLang} → ${targetLang} 的本地翻译模型`);
      }

      const translator = await TranslatorAPI.create({
        ...opts,
        monitor(monitor) {
          monitor.addEventListener('downloadprogress', event => {
            const total = Number(event.total) || 1;
            const loaded = Number(event.loaded) || 0;
            const percent = Math.max(0, Math.min(100, Math.round((loaded / total) * 100)));
            console.log(`[Word Tool] 下载翻译模型 ${sourceLang}→${targetLang}: ${percent}%`);
          });
        }
      });

      translators.set(cacheKey, translator);
      return translator;
    })();

    translatorPromises.set(cacheKey, promise);
    try {
      return await promise;
    } finally {
      translatorPromises.delete(cacheKey);
    }
  }

  function detectLanguage(text) {
    const t = text.trim();
    if (/[\u3040-\u309F\u30A0-\u30FF]/.test(t)) return 'ja';
    if (/[\uAC00-\uD7AF\u1100-\u11FF]/.test(t)) return 'ko';
    if (/[\u4E00-\u9FFF\u3400-\u4DBF]/.test(t)) return 'zh-CN';
    return 'en';
  }

  async function translateWithBrowser(text, targetLang = CONFIG.targetLanguage) {
    const sourceLang = detectLanguage(text);
    if (sourceLang === targetLang) return text;
    const translator = await getTranslator(sourceLang, targetLang);
    return await translator.translate(text);
  }

  async function checkTranslatorStatus() {
    const api = getTranslatorAPI();
    if (!api) {
      alert(`Word Tool：未检测到 Translator API。\n\n浏览器：${browserName()}\n如果你使用 Edge，请确认版本为 148 或更高，并更新 Tampermonkey。`);
      return;
    }

    try {
      const status = await api.availability({
        sourceLanguage: 'en',
        targetLanguage: CONFIG.targetLanguage
      });
      alert(`Word Tool：Translator API 已检测到。\n\n浏览器：${browserName()}\n英文 → 简体中文：${status}`);
    } catch (err) {
      alert(`Word Tool：检测 Translator API 时出错：\n${err?.message || String(err)}`);
    }
  }

  // ==================== 发音 ====================

  // 使用 Tampermonkey 跨域请求下载音频，再通过 Web Audio API 从内存播放。
  // 这样不会让当前网页直接加载 dict.youdao.com，可绕过 GitHub 等网站严格的 media-src CSP。
  const audioBufferCache = new Map();
  const audioBufferPromises = new Map();
  const audioNormalizationGainCache = new WeakMap();
  let audioContext = null;
  let currentAudioSource = null;
  let pronunciationRunId = 0;
  const IDENTIFIER_PART_PAUSE_MS = 180;

  // 有道真人语音响度归一化：
  // - 每个 AudioBuffer 只分析一次并缓存；
  // - 使用“有效语音区 RMS”避免首尾静音把增益错误放大；
  // - 限制最大放大 / 衰减，并用峰值上限避免削波。
  const YOUDAO_TARGET_RMS = 0.115;
  const YOUDAO_ACTIVE_SAMPLE_THRESHOLD = 0.01;
  const YOUDAO_MIN_GAIN = Math.pow(10, -8 / 20); // 最多衰减约 8 dB
  const YOUDAO_MAX_GAIN = Math.pow(10,  6 / 20); // 最多放大约 6 dB
  const YOUDAO_PEAK_CEILING = 0.95;
  const YOUDAO_MASTER_GAIN = 0.90;

  function getAudioContext() {
    if (audioContext && audioContext.state !== 'closed') return audioContext;

    const AudioContextCtor = globalThis.AudioContext || globalThis.webkitAudioContext;
    if (!AudioContextCtor) return null;

    audioContext = new AudioContextCtor();
    return audioContext;
  }

  function prepareAudioContext() {
    const ctx = getAudioContext();
    if (!ctx) return Promise.reject(new Error('当前浏览器不支持 Web Audio API'));

    // 这个函数会从双击/点击等用户操作事件中立即调用，尽量在 user activation 仍有效时 resume。
    if (ctx.state === 'suspended') {
      return ctx.resume().then(() => ctx);
    }
    return Promise.resolve(ctx);
  }

  function getSpeechSynthesis() {
    try {
      return getPageWindow().speechSynthesis || globalThis.speechSynthesis || window.speechSynthesis || null;
    } catch {
      return null;
    }
  }

  // 常见开发者术语：页面仍显示原词，只把更适合朗读的文本交给浏览器 TTS。
  // key 必须与 cleanWord() 处理后的结果一致。
  const PRONUNCIATION_OVERRIDES = new Map(Object.entries({
    pytest: 'pie test',
    pypi: 'pie pie',
    numpy: 'num pie',
    scipy: 'sigh pie',
    pandas: 'pan das',
    pytorch: 'pie torch',
    github: 'git hub',
    gitlab: 'git lab',
    npm: 'N P M',
    npx: 'N P X',
    pnpm: 'P N P M',
    yarn: 'yarn',
    vite: 'veet',
    vitest: 'vee test',
    typescript: 'type script',
    javascript: 'java script',
    nodejs: 'node J S',
    nextjs: 'next J S',
    vuejs: 'view J S',
    threejs: 'three J S',
    webgl: 'web G L',
    webgpu: 'web G P U',
    pytest: 'pie test',
    pypi: 'pie pie',
    numpy: 'num pie',
    scipy: 'sigh pie',
    pandas: 'pan das',
    pytorch: 'pie torch',
    github: 'git hub',
    gitlab: 'git lab',
    npm: 'N P M',
    npx: 'N P X',
    pnpm: 'P N P M',
    yarn: 'yarn',
    vite: 'veet',
    vitest: 'vee test',
    typescript: 'type script',
    javascript: 'java script',
    nodejs: 'node J S',
    nextjs: 'next J S',
    vuejs: 'view J S',
    threejs: 'three J S',
    webgl: 'web G L',
    webgpu: 'web G P U',
    claude: 'claud', // 很多人错读成 clau-de，实际类似 cloud
    mistral: 'mis tral',
    pytorch: 'py torch',
    huggingface: 'hugging face',
    langchain: 'lang chain',
    llama: 'la ma', // 西班牙语源，双l发l或y音，技术界普遍读 la-ma
    qwen: 'chiu wen', // 千问，外国人常读错，TTS也需要校准
    midjourney: 'mid journey',
    'stable-diffusion': 'stable diffusion',
    rag: 'rag', // Retrieval-Augmented Generation，通常作为完整单词读 [ræɡ]，而不是 R-A-G
    rlhf: 'R L H F',
    vue: 'view', // 极多新手误读为 V-U-E
    vite: 'veet', // 官方读音类似 veet（微特），常被错读为 bite 改发音
    svelte: 'svelt', // 尾音e不发音
    typescript: 'type script',
    nextjs: 'next J S',
    nuxtjs: 'nuxt J S',
    tailwindcss: 'tailwind C S S',
    pnpm: 'P N P M', // 或读 p-N-P-M，TTS常卡壳
    bun: 'bun',
    esbuild: 'E S build',
    webpack: 'web pack',
    kubernetes: 'koo ber net es', // 缩写 k8s 读 'K eights'
    nginx: 'engine X', // 经典易错词，绝非 n-g-i-n-x
    sqlite: 'S Q L lite', // 官方推荐读法，而非 S-Q-L-i-t-e
    postgresql: 'post gres Q L', // 通常简称为 postgres 'post gres'
    redis: 'reh dis', // 容易被错读为 ree-dis
    mongodb: 'mongo D B',
    prisma: 'priz ma',
    docker: 'dock er',
    kafka: 'cough ka', // 类似 kafka
    grpc: 'G R P C',
    graphql: 'graph Q L',
    wasm: 'wah zum', // WebAssembly，普遍读成一个单词 [wæzəm]，而非 W-A-S-M
    yaml: 'yam el', // 读音类似 yam-el，而非 Y-A-M-L
    json: 'jay sun', // 读音类似 jay-son
    cron: 'kron', // 定时任务，读如 ch-ron
    'ci-cd': 'C I C D',
    devops: 'dev ops',
    oauth: 'O auth', // 开放授权，读 O-auth
    jwt: 'jot', // JSON Web Token，业界常口语读作 'jot'，TTS常傻读 J-W-T
    sudo: 'soo doo', // superuser do，常被错读为 su-doh
    epoch: 'eh puck', // 机器学习的 epoch，常被误读为 ee-poch
    asus: 'ay soos', // 华硕官方读音 ay-soos，常被读错
    xiaomi: 'shiaw mee', // 小米，海外TTS和老外常读成 zia-omi
    huawei: 'wah way', // 华为
    linux: 'lin nux', // 经典易错，非 line-nux
    adobe: 'ah doh bee', // 尾音e发音，常被漏掉
    gif: 'jif', // 开发者常吵架的读音，多数倾向 jif 或 gif，此处可转为直观发音
    moore: 'moor', // 摩尔定律的 moore
    cache: 'cash', // 缓存，极多非母语者误读为 ca-shay
    archive: 'ar kive'
  }));

  function stopCurrentAudio(cancelSequence = true) {
    if (cancelSequence) pronunciationRunId += 1;

    if (currentAudioSource) {
      try { currentAudioSource.stop(0); } catch {}
      try { currentAudioSource.disconnect(); } catch {}
      currentAudioSource = null;
    }

    const synth = getSpeechSynthesis();
    if (synth?.speaking || synth?.pending) {
      try { synth.cancel(); } catch {}
    }
  }

  function delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  function getPreferredEnglishVoice(synth) {
    const voices = synth?.getVoices?.() || [];
    if (!voices.length) return null;

    const preferredPrefix = CONFIG.youdaoType === 0 ? 'en-GB' : 'en-US';
    return voices.find(v => v.lang?.toLowerCase() === preferredPrefix.toLowerCase())
      || voices.find(v => v.lang?.toLowerCase().startsWith(preferredPrefix.slice(0, 2).toLowerCase()))
      || null;
  }

  function speakWithBrowserTTS(text) {
    return new Promise((resolve, reject) => {
      const synth = getSpeechSynthesis();
      const UtteranceCtor = getPageWindow().SpeechSynthesisUtterance || globalThis.SpeechSynthesisUtterance;
      if (!synth || !UtteranceCtor) {
        reject(new Error('当前浏览器不支持 Speech Synthesis TTS'));
        return;
      }

      try { synth.cancel(); } catch {}

      const utterance = new UtteranceCtor(text);
      utterance.lang = CONFIG.youdaoType === 0 ? 'en-GB' : 'en-US';
      utterance.rate = 0.92;
      utterance.pitch = 1;

      const voice = getPreferredEnglishVoice(synth);
      if (voice) utterance.voice = voice;

      let settled = false;
      const finish = (fn, value) => {
        if (settled) return;
        settled = true;
        fn(value);
      };

      utterance.onend = () => finish(resolve);
      utterance.onerror = event => finish(reject, new Error(`浏览器 TTS 播放失败：${event.error || 'unknown'}`));

      try {
        synth.speak(utterance);
      } catch (err) {
        finish(reject, err);
      }
    });
  }

  // 普通词清洗：用于词典查询 / 单词音频 URL。
  function cleanWord(word) {
    const cleaned = String(word || '')
      .trim()
      .toLowerCase()
      .replace(/[^a-z'-]/gi, '');
    if (!cleaned || cleaned.length < 1 || cleaned.length > 50) return null;
    return cleaned;
  }

  // 双击技术标识符时保留 "_" / "-"，避免 user_message_delta 被压成 usermessagedelta。
  function cleanPronounceToken(token) {
    const cleaned = String(token || '')
      .trim()
      .replace(/[‐‑‒–—−]/g, '-')
      .replace(/^[^a-z0-9]+|[^a-z0-9_'’-]+$/gi, '')
      .replace(/[^a-z0-9_'’-]/gi, '');
    if (!cleaned || cleaned.length < 1 || cleaned.length > 120) return null;
    if (!/[a-z]/i.test(cleaned)) return null;
    return cleaned;
  }

  // 支持：
  //   STILL_GENERATING      -> still | generating
  //   user_message_delta    -> user | message | delta
  //   reply-timeout         -> reply | timeout
  //   userMessageDelta      -> user | message | delta
  //   XMLHttpRequest        -> xml | http | request
  function splitPronounceToken(token) {
    const cleanedToken = cleanPronounceToken(token);
    if (!cleanedToken) return [];

    const rawParts = cleanedToken
      .split(/[_-]+/)
      .flatMap(part => part
        .replace(/([a-z0-9])([A-Z])/g, '$1 $2')
        .replace(/([A-Z]+)([A-Z][a-z])/g, '$1 $2')
        .split(/\s+/)
      )
      .map(part => cleanWord(part))
      .filter(Boolean);

    return rawParts;
  }

  function isCompoundPronounceToken(token, parts = splitPronounceToken(token)) {
    return parts.length > 1 || /[_-]/.test(String(token || ''));
  }

  async function fetchCompoundWordInfo(parts) {
    const infos = await Promise.all(parts.map(part => fetchWordInfo(part)));

    // 保留每个子词自己的结果。即使某个词没有查到，也保留该行，
    // 避免 PROFILE_NAME / user_message_delta 视觉上像“只翻译了第一个词”。
    return {
      usphone: null,
      meanings: [],
      compoundParts: parts.map((part, index) => ({
        word: part,
        usphone: infos[index]?.usphone || null,
        meanings: Array.isArray(infos[index]?.meanings)
          ? infos[index].meanings.slice(0, 2)
          : []
      }))
    };
  }

  function getYoudaoUrl(word) {
    return `https://dict.youdao.com/dictvoice?audio=${encodeURIComponent(word)}&type=${CONFIG.youdaoType}`;
  }

  function fetchWordInfo(word) {
    return new Promise(resolve => {
      GM_xmlhttpRequest({
        method: 'GET',
        url: `https://dict.youdao.com/jsonapi?q=${encodeURIComponent(word)}`,
        timeout: 5000,
        onload: res => {
          try {
            const data = JSON.parse(res.responseText);
            const wordInfo = data.ec?.word?.[0];
            if (!wordInfo) return resolve(null);

            const usphone = wordInfo.usphone ? `/${wordInfo.usphone}/` : null;
            const meanings = (wordInfo.trs || [])
              .slice(0, 3)
              .map(item => item?.tr?.[0]?.l?.i?.[0])
              .filter(Boolean);

            resolve({ usphone, meanings });
          } catch {
            resolve(null);
          }
        },
        onerror: () => resolve(null),
        ontimeout: () => resolve(null)
      });
    });
  }

  function fetchAudioArrayBuffer(url) {
    return new Promise((resolve, reject) => {
      GM_xmlhttpRequest({
        method: 'GET',
        url,
        responseType: 'arraybuffer',
        timeout: 8000,
        headers: {
          Accept: 'audio/mpeg,audio/*;q=0.9,*/*;q=0.8'
        },
        onload: res => {
          if (res.status < 200 || res.status >= 300) {
            reject(new Error(`有道语音请求失败：HTTP ${res.status}`));
            return;
          }
          if (!(res.response instanceof ArrayBuffer) || res.response.byteLength === 0) {
            reject(new Error('有道语音返回了空音频'));
            return;
          }
          resolve(res.response);
        },
        onerror: () => reject(new Error('有道语音网络请求失败')),
        ontimeout: () => reject(new Error('有道语音请求超时'))
      });
    });
  }

  async function getAudioBuffer(url, ctx) {
    if (audioBufferCache.has(url)) return audioBufferCache.get(url);
    if (audioBufferPromises.has(url)) return audioBufferPromises.get(url);

    const promise = (async () => {
      const arrayBuffer = await fetchAudioArrayBuffer(url);
      // slice(0) 给 decodeAudioData 一份独立缓冲，兼容部分 Chromium 实现。
      const decoded = await ctx.decodeAudioData(arrayBuffer.slice(0));
      audioBufferCache.set(url, decoded);
      return decoded;
    })();

    audioBufferPromises.set(url, promise);
    try {
      return await promise;
    } finally {
      audioBufferPromises.delete(url);
    }
  }


  function getNormalizedGainForBuffer(buffer) {
    if (audioNormalizationGainCache.has(buffer)) {
      return audioNormalizationGainCache.get(buffer);
    }

    let sumSquares = 0;
    let activeCount = 0;
    let peak = 0;

    for (let channel = 0; channel < buffer.numberOfChannels; channel += 1) {
      const data = buffer.getChannelData(channel);
      for (let i = 0; i < data.length; i += 1) {
        const sample = data[i];
        const abs = Math.abs(sample);
        if (abs > peak) peak = abs;

        // 忽略近似静音样本，避免不同录音的首尾静音长度影响 RMS。
        if (abs >= YOUDAO_ACTIVE_SAMPLE_THRESHOLD) {
          sumSquares += sample * sample;
          activeCount += 1;
        }
      }
    }

    let gain = 1;
    if (activeCount > 0 && sumSquares > 0) {
      const rms = Math.sqrt(sumSquares / activeCount);
      gain = YOUDAO_TARGET_RMS / rms;
    }

    gain = Math.max(YOUDAO_MIN_GAIN, Math.min(YOUDAO_MAX_GAIN, gain));

    // 在归一化之后仍为峰值留出安全余量，降低瞬态削波风险。
    if (peak > 0) {
      gain = Math.min(gain, YOUDAO_PEAK_CEILING / peak);
    }

    gain *= YOUDAO_MASTER_GAIN;
    audioNormalizationGainCache.set(buffer, gain);
    return gain;
  }

  async function playFromUrl(url, preparedContextPromise = null) {
    stopCurrentAudio(false);

    const ctx = await (preparedContextPromise || prepareAudioContext());
    if (ctx.state === 'suspended') {
      await ctx.resume();
    }

    const buffer = await getAudioBuffer(url, ctx);
    stopCurrentAudio(false);

    return await new Promise((resolve, reject) => {
      const source = ctx.createBufferSource();
      source.buffer = buffer;

      const gainNode = ctx.createGain();
      gainNode.gain.value = getNormalizedGainForBuffer(buffer);

      // 原生动态压缩器只负责兜住偶发峰值；主体音量统一由 RMS 归一化完成。
      const compressor = ctx.createDynamicsCompressor();
      compressor.threshold.value = -14;
      compressor.knee.value = 8;
      compressor.ratio.value = 3;
      compressor.attack.value = 0.003;
      compressor.release.value = 0.18;

      source.connect(gainNode);
      gainNode.connect(compressor);
      compressor.connect(ctx.destination);

      let settled = false;
      const finish = (err = null) => {
        if (settled) return;
        settled = true;
        if (currentAudioSource === source) currentAudioSource = null;
        try { source.disconnect(); } catch {}
        try { gainNode.disconnect(); } catch {}
        try { compressor.disconnect(); } catch {}
        if (err) reject(err);
        else resolve();
      };

      source.onended = () => finish();
      currentAudioSource = source;

      try {
        source.start(0);
      } catch (err) {
        finish(err);
      }
    });
  }

  let pronounceTooltipRemoveHandler = null;

  function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = String(text ?? '');
    return div.innerHTML;
  }

  function createPronounceTooltip(word, icon, wordInfo, isLoading, pronunciationText = '', pronunciationMeta = '') {
    document.getElementById('wt-pronounce-tooltip')?.remove();

    if (pronounceTooltipRemoveHandler) {
      document.removeEventListener('mousemove', pronounceTooltipRemoveHandler, true);
      pronounceTooltipRemoveHandler = null;
    }

    const tip = document.createElement('div');
    tip.id = 'wt-pronounce-tooltip';

    const safeWord = escapeHtml(word);
    const phone = wordInfo?.usphone ? escapeHtml(wordInfo.usphone) : '';
    const safePronunciation = pronunciationText ? escapeHtml(pronunciationText) : '';
    const meanings = Array.isArray(wordInfo?.meanings)
      ? wordInfo.meanings.map(escapeHtml).join('；')
      : '';
    const compoundParts = Array.isArray(wordInfo?.compoundParts)
      ? wordInfo.compoundParts
      : [];

    // pronunciationText 可用于两种场景：
    // 1) pytest -> pie test 这类固定技术词读法；
    // 2) user_message_delta -> user → message → delta 这类标识符拆读预览。
    if (safePronunciation) {
      const safeMeta = escapeHtml(pronunciationMeta || '技术词 · 浏览器 TTS');
      tip.innerHTML = `
        <div style="font-size:17px;font-weight:700;color:#111;">${icon} ${safeWord}</div>
        <div style="font-size:16px;color:#111;margin-top:7px;">
          <span style="font-size:12px;color:#777;margin-right:6px;">读法</span>
          <span style="font-weight:700;letter-spacing:.15px;">${safePronunciation}</span>
        </div>
        <div style="font-size:11px;color:#5f6368;margin-top:5px;">${safeMeta}</div>`;
      if (phone) {
        tip.innerHTML += `<div style="font-size:13px;color:#666;margin-top:5px;">词典音标：${phone}</div>`;
      }
    } else {
      tip.innerHTML = phone
        ? `<div style="font-size:17px;font-weight:700;color:#111;">${icon} ${phone}</div>
           <div style="font-size:14px;color:#666;margin-top:3px;">${safeWord}</div>`
        : `<div style="font-size:17px;font-weight:700;color:#111;">${icon} ${safeWord}</div>`;
    }

    if (compoundParts.length) {
      const rows = compoundParts.map(part => {
        const safePart = escapeHtml(part.word);
        const safePhone = part.usphone ? escapeHtml(part.usphone) : '';
        const partMeanings = Array.isArray(part.meanings) && part.meanings.length
          ? part.meanings.map(escapeHtml).join('；')
          : '<span style="color:#999;">未获取到词典释义</span>';

        return `
          <div style="padding:9px 0;border-top:1px solid #e7e7e7;">
            <div style="display:flex;align-items:baseline;gap:7px;margin-bottom:3px;">
              <span style="font-size:15px;font-weight:700;color:#111;">${safePart}</span>
              ${safePhone ? `<span style="font-size:12px;color:#777;">${safePhone}</span>` : ''}
            </div>
            <div style="font-size:14px;color:#333;line-height:1.6;word-wrap:break-word;">${partMeanings}</div>
          </div>`;
      }).join('');

      tip.innerHTML += `
        <div style="font-size:12px;color:#666;margin-top:9px;margin-bottom:3px;font-weight:600;">逐词释义</div>
        <div style="max-width:340px;">${rows}</div>`;
    } else if (meanings) {
      tip.innerHTML += `<div style="font-size:14px;color:#333;margin-top:6px;max-width:320px;line-height:1.6;word-wrap:break-word;">${meanings}</div>`;
    }
    if (isLoading) {
      tip.innerHTML += '<div style="font-size:12px;color:#777;margin-top:5px;">加载词义...</div>';
    }

    tip.style.cssText = [
      'position:fixed',
      'z-index:2147483647',
      'background:#fff',
      'color:#111',
      'padding:14px 16px',
      'border-radius:12px',
      'font-size:13px',
      'pointer-events:none',
      'opacity:0',
      'transition:opacity .15s',
      'box-shadow:0 10px 30px rgba(0,0,0,.22),0 2px 8px rgba(0,0,0,.12)',
      "font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif",
      'max-width:420px'
    ].join(';');

    document.body.appendChild(tip);

    const sel = window.getSelection();
    if (sel?.rangeCount) {
      const rect = sel.getRangeAt(0).getBoundingClientRect();
      let left = rect.left + rect.width / 2 - tip.offsetWidth / 2;
      let top = rect.top - tip.offsetHeight - 8;
      left = Math.max(8, Math.min(window.innerWidth - tip.offsetWidth - 8, left));
      if (top < 8) top = Math.min(window.innerHeight - tip.offsetHeight - 8, rect.bottom + 8);
      tip.style.left = `${left}px`;
      tip.style.top = `${top}px`;
    } else {
      tip.style.left = '16px';
      tip.style.top = '16px';
    }

    requestAnimationFrame(() => { tip.style.opacity = '1'; });

    let lastMouseX = null;
    let lastMouseY = null;
    pronounceTooltipRemoveHandler = e => {
      if (lastMouseX === null) {
        lastMouseX = e.clientX;
        lastMouseY = e.clientY;
        return;
      }
      if (Math.abs(e.clientX - lastMouseX) > 5 || Math.abs(e.clientY - lastMouseY) > 5) {
        tip.style.opacity = '0';
        setTimeout(() => tip.remove(), 150);
        document.removeEventListener('mousemove', pronounceTooltipRemoveHandler, true);
        pronounceTooltipRemoveHandler = null;
      }
    };
    document.addEventListener('mousemove', pronounceTooltipRemoveHandler, true);
  }

  async function pronounce(token) {
    const cleanedToken = cleanPronounceToken(token);
    if (!cleanedToken) return;

    const parts = splitPronounceToken(cleanedToken);
    if (!parts.length) return;

    const runId = ++pronunciationRunId;
    const compound = isCompoundPronounceToken(cleanedToken, parts);
    const singleWord = parts[0];
    const singleOverride = parts.length === 1
      ? (PRONUNCIATION_OVERRIDES.get(singleWord) || '')
      : '';

    // 新一轮发音开始时，立即停止上一轮音频，但不要把刚创建的 runId 自己取消掉。
    stopCurrentAudio(false);

    // 技术标识符展示拆词结果；普通固定技术词展示 override。
    const pronunciationText = compound
      ? parts.join(' → ')
      : singleOverride;
    const pronunciationMeta = compound
      ? `标识符拆读 · 间隔 ${IDENTIFIER_PART_PAUSE_MS}ms`
      : (singleOverride ? '技术词 · 浏览器 TTS' : '');

    createPronounceTooltip(cleanedToken, '🔊', null, true, pronunciationText, pronunciationMeta);

    const wordInfoPromise = compound
      ? fetchCompoundWordInfo(parts)
      : fetchWordInfo(singleWord);

    // 在双击的用户激活链里尽早创建 / resume AudioContext。
    // 后面即使需要等网络音频，也不容易被 Chromium 自动播放策略拦截。
    const preparedContextPromise = prepareAudioContext().catch(err => {
      console.warn('[Word Tool] Web Audio unavailable, will use browser TTS when needed:', err);
      return null;
    });

    const speakPart = async part => {
      if (runId !== pronunciationRunId) return false;

      const overrideText = PRONUNCIATION_OVERRIDES.get(part) || '';
      if (overrideText) {
        try {
          await speakWithBrowserTTS(overrideText);
        } catch (err) {
          console.warn('[Word Tool] Technical-term TTS failed:', part, err);
        }
        return runId === pronunciationRunId;
      }

      try {
        const prepared = await preparedContextPromise;
        if (!prepared) throw new Error('Web Audio context unavailable');
        await playFromUrl(getYoudaoUrl(part), Promise.resolve(prepared));
      } catch (err) {
        console.warn('[Word Tool] Youdao audio failed, fallback to browser TTS:', part, err);
        try {
          await speakWithBrowserTTS(part);
        } catch (ttsErr) {
          console.warn('[Word Tool] Browser TTS fallback failed:', part, ttsErr);
        }
      }

      return runId === pronunciationRunId;
    };

    const audioPromise = (async () => {
      for (let i = 0; i < parts.length; i += 1) {
        if (runId !== pronunciationRunId) return;

        const shouldContinue = await speakPart(parts[i]);
        if (!shouldContinue || runId !== pronunciationRunId) return;

        if (i < parts.length - 1) {
          await delay(IDENTIFIER_PART_PAUSE_MS);
        }
      }
    })();

    const wordInfo = await wordInfoPromise;
    if (runId === pronunciationRunId) {
      createPronounceTooltip(
        cleanedToken,
        '🔊',
        wordInfo,
        false,
        pronunciationText,
        pronunciationMeta
      );
    }

    await audioPromise;
  }

  // ==================== 选区 ====================

  function getSelectedWord() {
    const sel = window.getSelection();
    if (!sel || sel.isCollapsed || !sel.rangeCount) return null;

    const text = sel.toString().trim();
    const firstToken = text.split(/\s+/)[0];
    const cleaned = cleanPronounceToken(firstToken);
    if (!cleaned) return null;

    const range = sel.getRangeAt(0);
    const container = range.commonAncestorContainer;
    const parentEl = container.nodeType === Node.TEXT_NODE ? container.parentElement : container;
    if (parentEl?.closest?.('input, textarea, [contenteditable="true"], .notranslate, script, style')) return null;

    return cleaned;
  }

  function getSelectedText() {
    const sel = window.getSelection();
    if (!sel || sel.isCollapsed || !sel.rangeCount) return null;

    const text = sel.toString().trim();
    if (!text || text.length < 2 || text.length > 1500) return null;

    const range = sel.getRangeAt(0);
    const container = range.commonAncestorContainer;
    const parentEl = container.nodeType === Node.TEXT_NODE ? container.parentElement : container;
    if (parentEl?.closest?.('input, textarea, [contenteditable="true"], .notranslate, script, style, code, pre')) return null;

    return { text, range: range.getBoundingClientRect() };
  }

  function getWordAtPosition(x, y) {
    let range = null;
    let offsetNode = null;
    let offset = null;

    if (document.caretPositionFromPoint) {
      const pos = document.caretPositionFromPoint(x, y);
      if (pos?.offsetNode) {
        offsetNode = pos.offsetNode;
        offset = pos.offset;
        range = document.createRange();
        range.setStart(pos.offsetNode, pos.offset);
        range.collapse(true);
      }
    } else if (document.caretRangeFromPoint) {
      range = document.caretRangeFromPoint(x, y);
      if (range) {
        offsetNode = range.startContainer;
        offset = range.startOffset;
      }
    }

    if (!range) return null;

    // 对技术 token 优先自己向左右扫描。
    // 浏览器的 Range.expand('word') 对 "_" / "-" 的边界判定在不同页面可能不一致。
    if (offsetNode?.nodeType === Node.TEXT_NODE && Number.isInteger(offset)) {
      const raw = offsetNode.textContent || '';
      const isTokenChar = ch => /[a-z0-9_'’‐‑‒–—−-]/i.test(ch || '');

      let start = Math.min(offset, raw.length);
      let end = Math.min(offset, raw.length);

      // 鼠标可能落在字符右缘；先允许向左吃到当前 token。
      while (start > 0 && isTokenChar(raw[start - 1])) start -= 1;
      while (end < raw.length && isTokenChar(raw[end])) end += 1;

      const scanned = cleanPronounceToken(raw.slice(start, end));
      if (scanned) return scanned;
    }

    // 普通单词回退到 Chromium / Edge 的 word expansion。
    try {
      if (typeof range.expand === 'function') range.expand('word');
    } catch {}

    return cleanPronounceToken(range.toString().trim());
  }

  // ==================== ALT 拖动翻译 ====================

  let translatePopup = null;

  function removeTranslatePopup() {
    const existing = document.getElementById('wt-translate-popup');
    if (existing) {
      existing.style.opacity = '0';
      existing.style.transform = 'translateY(8px)';
      setTimeout(() => existing.remove(), 160);
    }
    translatePopup = null;
  }

  function ensureCommonStyles() {
    if (document.getElementById('wt-styles')) return;
    const style = document.createElement('style');
    style.id = 'wt-styles';
    style.textContent = `
      .wt-header{display:flex;align-items:center;gap:8px;margin-bottom:8px;}
      .wt-engine{font-size:11px;color:#999;flex:1;}
      .wt-actions{display:flex;gap:4px;}
      .wt-btn{width:28px;height:28px;border:none;background:transparent;border-radius:6px;cursor:pointer;font-size:14px;display:flex;align-items:center;justify-content:center;}
      .wt-btn:hover{background:#f0f0f0;}
      .wt-translated{color:#333;font-size:15px;font-weight:500;word-break:break-word;padding:4px 0;}
    `;
    (document.head || document.documentElement).appendChild(style);
  }

  function createTranslatePopup(translatedText, rect, engineLabel = `${browserName()} 本地翻译`) {
    removeTranslatePopup();
    ensureCommonStyles();

    const popup = document.createElement('div');
    popup.id = 'wt-translate-popup';
    popup.innerHTML = `
      <div class="wt-header">
        <span class="wt-engine">${escapeHtml(engineLabel)}</span>
        <div class="wt-actions">
          <button class="wt-btn wt-close-btn" title="关闭">✕</button>
        </div>
      </div>
      <div class="wt-translated">${escapeHtml(translatedText)}</div>
    `;

    popup.style.cssText = [
      'position:fixed',
      'z-index:2147483647',
      'background:#fff',
      'color:#333',
      'border-radius:12px',
      'padding:12px 16px',
      "font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif",
      'box-shadow:0 4px 24px rgba(0,0,0,.15),0 0 0 1px rgba(0,0,0,.05)',
      'max-width:420px',
      'min-width:200px',
      'opacity:0',
      'transform:translateY(8px)',
      'transition:opacity .2s ease,transform .2s ease',
      'cursor:default',
      'user-select:text',
      'font-size:14px',
      'line-height:1.5'
    ].join(';');

    document.body.appendChild(popup);

    const popupRect = popup.getBoundingClientRect();
    let left = rect.left + rect.width / 2 - popupRect.width / 2;
    let top = rect.top - popupRect.height - 12;

    if (left < 8) left = 8;
    if (left + popupRect.width > window.innerWidth - 8) left = window.innerWidth - popupRect.width - 8;
    if (top < 8) top = rect.bottom + 12;
    if (top + popupRect.height > window.innerHeight - 8) top = Math.max(8, window.innerHeight - popupRect.height - 8);

    popup.style.left = `${left}px`;
    popup.style.top = `${top}px`;

    requestAnimationFrame(() => {
      popup.style.opacity = '1';
      popup.style.transform = 'translateY(0)';
    });

    popup.querySelector('.wt-close-btn')?.addEventListener('click', e => {
      e.stopPropagation();
      removeTranslatePopup();
    });

    const clickOutside = e => {
      if (!popup.contains(e.target)) {
        removeTranslatePopup();
        document.removeEventListener('click', clickOutside, true);
      }
    };
    setTimeout(() => document.addEventListener('click', clickOutside, true), 100);

    translatePopup = popup;
  }

  function createStatusPopup(message, rect, isError = false) {
    removeTranslatePopup();
    const popup = document.createElement('div');
    popup.id = 'wt-translate-popup';
    popup.textContent = message;
    popup.style.cssText = [
      'position:fixed',
      'z-index:2147483647',
      'background:#fff',
      `color:${isError ? '#d93025' : '#666'}`,
      'font-size:13px',
      'padding:9px 12px',
      'border-radius:8px',
      'box-shadow:0 4px 16px rgba(0,0,0,.14)',
      'max-width:360px',
      'line-height:1.5',
      `left:${Math.max(8, rect.left)}px`,
      `top:${Math.max(8, rect.top - 44)}px`
    ].join(';');
    document.body.appendChild(popup);
    translatePopup = popup;
    return popup;
  }

  // ==================== 整页翻译 ====================

  const PAGE_PREFIX = 'wt-page';
  const GENERATED_CLASS = `${PAGE_PREFIX}-translation`;
  const CONTAINER_CLASS = `${PAGE_PREFIX}-translated-container`;

  const BLOCK_TAGS = new Set([
    'div', 'p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6',
    'article', 'section', 'header', 'footer', 'main', 'aside', 'nav',
    'blockquote', 'li', 'td', 'th'
  ]);

  const EXCLUDE_TAGS = new Set([
    'script', 'style', 'noscript', 'iframe', 'object', 'embed',
    'canvas', 'svg', 'math', 'code', 'pre', 'kbd', 'samp', 'var',
    'textarea', 'input', 'select', 'option', 'template', 'meta', 'link', 'head', 'title', 'button'
  ]);

  const EXCLUDE_SELECTORS = [
    '[data-translate="no"]', '.notranslate', '[translate="no"]',
    '[role="tooltip"]', '.tooltip', '.alt-text',
    `.${GENERATED_CLASS}`
  ];

  function isExcluded(el) {
    if (!el || !el.tagName) return true;
    const tag = el.tagName.toLowerCase();
    if (EXCLUDE_TAGS.has(tag)) return true;
    if (el.isContentEditable) return true;
    if (el.closest?.(`#${PAGE_PREFIX}-btn, #wt-translate-popup, #wt-pronounce-tooltip, #wt-settings-modal`)) return true;
    return EXCLUDE_SELECTORS.some(sel => {
      try { return Boolean(el.closest?.(sel)); } catch { return false; }
    });
  }

  function isVisible(el) {
    if (!el) return false;
    const style = window.getComputedStyle(el);
    if (style.display === 'none' || style.visibility === 'hidden' || style.opacity === '0') return false;
    const rect = el.getBoundingClientRect();
    return rect.width > 0 || rect.height > 0;
  }

  function findBlockContainer(el) {
    let current = el;
    while (current && current !== document.body) {
      if (BLOCK_TAGS.has(current.tagName?.toLowerCase())) return current;
      current = current.parentElement;
    }
    return el;
  }

  function extractTextNodes(root) {
    const nodes = [];
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, {
      acceptNode(node) {
        const parent = node.parentElement;
        if (!parent) return NodeFilter.FILTER_REJECT;
        if (isExcluded(parent)) return NodeFilter.FILTER_REJECT;
        if (!isVisible(parent)) return NodeFilter.FILTER_REJECT;

        const text = node.textContent.trim();
        if (!text || text.length < 2) return NodeFilter.FILTER_REJECT;
        if (/^[\d\W]+$/u.test(text)) return NodeFilter.FILTER_REJECT;
        return NodeFilter.FILTER_ACCEPT;
      }
    });

    let node;
    while ((node = walker.nextNode())) nodes.push(node);
    return nodes;
  }

  function groupByParagraph(textNodes) {
    const groups = new Map();
    for (const node of textNodes) {
      const container = findBlockContainer(node.parentElement);
      if (!container || container.classList?.contains(CONTAINER_CLASS)) continue;
      if (!groups.has(container)) groups.set(container, { container, nodes: [] });
      groups.get(container).nodes.push(node);
    }
    return Array.from(groups.values());
  }

  const translatedElements = new Set();
  const translatedContainers = new Set();

  async function translateGroup(group) {
    const { container, nodes } = group;
    if (!container?.isConnected || container.classList.contains(CONTAINER_CLASS)) return;

    const combinedText = nodes.map(n => n.textContent.trim()).filter(Boolean).join(' ');
    if (!combinedText || combinedText.length < 3) return;
    if (detectLanguage(combinedText) === 'zh-CN') return;

    try {
      const translated = await translateWithBrowser(combinedText);
      if (!translated || translated === combinedText || !isPageTranslating) return;

      const bilingualEl = document.createElement('div');
      bilingualEl.className = GENERATED_CLASS;
      bilingualEl.style.cssText = 'margin:6px 0;padding:6px 10px;background:rgba(0,200,100,.04);font-size:.9em;color:#444;line-height:1.6;opacity:0;transition:opacity .3s;border-radius:4px;';
      bilingualEl.textContent = translated;

      container.classList.add(CONTAINER_CLASS);
      container.appendChild(bilingualEl);
      translatedContainers.add(container);
      translatedElements.add(bilingualEl);
      requestAnimationFrame(() => { bilingualEl.style.opacity = '1'; });
    } catch (err) {
      console.warn('[Word Tool] Page translate failed:', err);
    }
  }

  let isPageTranslating = false;
  let pageTranslationQueue = [];
  let pageProcessing = false;
  let pageObserver = null;

  async function processPageQueue() {
    if (pageProcessing) return;
    pageProcessing = true;
    try {
      while (pageTranslationQueue.length > 0 && isPageTranslating) {
        const group = pageTranslationQueue.shift();
        await translateGroup(group);
      }
    } finally {
      pageProcessing = false;
    }
  }

  async function startPageTranslation() {
    if (isPageTranslating) return;

    if (!getTranslatorAPI()) {
      const rect = floatingBtn?.getBoundingClientRect() || { left: 16, top: 60 };
      createStatusPopup(`未检测到 ${browserName()} Translator API。Edge 需 148+；也请更新 Tampermonkey。`, rect, true);
      return;
    }

    isPageTranslating = true;
    updateFloatingButton(true);

    const groups = groupByParagraph(extractTextNodes(document.body));

    pageObserver?.disconnect();
    pageObserver = new IntersectionObserver(entries => {
      for (const entry of entries) {
        if (!entry.isIntersecting || !isPageTranslating) continue;
        const group = entry.target.__wtTranslationGroup;
        if (group && !entry.target.classList.contains(CONTAINER_CLASS)) {
          pageTranslationQueue.push(group);
          processPageQueue();
        }
        pageObserver.unobserve(entry.target);
      }
    }, { rootMargin: '250px 0px' });

    for (const group of groups) {
      group.container.__wtTranslationGroup = group;
      pageObserver.observe(group.container);
    }
  }

  function stopPageTranslation() {
    isPageTranslating = false;
    pageTranslationQueue = [];
    pageObserver?.disconnect();
    pageObserver = null;
    updateFloatingButton(false);

    for (const el of translatedElements) {
      el.style.opacity = '0';
      setTimeout(() => el.remove(), 180);
    }
    translatedElements.clear();

    for (const container of translatedContainers) {
      container.classList.remove(CONTAINER_CLASS);
      try { delete container.__wtTranslationGroup; } catch {}
    }
    translatedContainers.clear();
  }

  function togglePageTranslation() {
    if (isPageTranslating) stopPageTranslation();
    else startPageTranslation();
  }

  // ==================== 浮动按钮 ====================

  let floatingBtn = null;

  function createFloatingButton() {
    if (floatingBtn || !document.body) return;

    floatingBtn = document.createElement('div');
    floatingBtn.id = `${PAGE_PREFIX}-btn`;
    floatingBtn.innerHTML = '<svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor"><path d="M12.87 15.07l-2.54-2.51.03-.03c1.74-1.94 2.98-4.17 3.71-6.53H17V4h-7V2H8v2H1v1.99h11.17C11.5 7.92 10.44 9.75 9 11.35 8.07 10.32 7.3 9.19 6.69 8h-2c.73 1.63 1.73 3.17 2.98 4.56l-5.09 5.02L4 19l5-5 3.11 3.11.76-2.04zM18.5 10h-2L12 22h2l1.12-3h4.75L21 22h2l-4.5-12zm-2.62 7l1.62-4.33L19.12 17h-3.24z"/></svg>';
    floatingBtn.style.cssText = 'position:fixed;bottom:92px;left:24px;z-index:2147483647;width:32px;height:32px;border-radius:50%;background:#1a73e8;color:#fff;display:flex;align-items:center;justify-content:center;cursor:pointer;box-shadow:0 2px 8px rgba(0,0,0,.25);transition:background .2s ease,transform .15s ease;user-select:none;';
      // 改图标位置 ⬆️
    let isDragging = false;
    let hasMoved = false;
    let startX = 0;
    let startY = 0;
    let startLeft = 0;
    let startTop = 0;

    floatingBtn.addEventListener('mouseenter', () => {
      if (!isDragging) floatingBtn.style.transform = 'scale(1.15)';
    });
    floatingBtn.addEventListener('mouseleave', () => {
      if (!isDragging) floatingBtn.style.transform = 'scale(1)';
    });

    floatingBtn.addEventListener('mousedown', e => {
      if (e.button !== 0) return;
      isDragging = true;
      hasMoved = false;
      startX = e.clientX;
      startY = e.clientY;
      const rect = floatingBtn.getBoundingClientRect();
      startLeft = rect.left;
      startTop = rect.top;
      floatingBtn.style.transition = 'none';
      floatingBtn.style.transform = 'none';
      e.preventDefault();
      e.stopPropagation();
    });

    document.addEventListener('mousemove', e => {
      if (!isDragging) return;
      const dx = e.clientX - startX;
      const dy = e.clientY - startY;
      if (Math.abs(dx) > 3 || Math.abs(dy) > 3) hasMoved = true;
      if (!hasMoved) return;

      const newLeft = Math.max(0, Math.min(window.innerWidth - 32, startLeft + dx));
      const newTop = Math.max(0, Math.min(window.innerHeight - 32, startTop + dy));
      floatingBtn.style.left = `${newLeft}px`;
      floatingBtn.style.top = `${newTop}px`;
      floatingBtn.style.right = 'auto';
      floatingBtn.style.bottom = 'auto';
    }, true);

    document.addEventListener('mouseup', e => {
      if (!isDragging || e.button !== 0) return;
      isDragging = false;
      floatingBtn.style.transition = 'background .2s ease,transform .15s ease';
      if (!hasMoved) togglePageTranslation();
    }, true);

    document.body.appendChild(floatingBtn);
    updateFloatingButton(false);
  }

  function updateFloatingButton(translating) {
    if (!floatingBtn) return;
    floatingBtn.style.background = translating ? '#34a853' : '#1a73e8';
    floatingBtn.title = translating ? '点击停止翻译' : `点击翻译此页（${browserName()} 本地 Translator API）`;
  }

  // ==================== 输入框三空格 → 英文 ====================

  const tripleSpaceState = new WeakMap();
  const TRIPLE_SPACE_MAX_GAP_MS = 1000;

  function getEditableTarget(target) {
    const el = target instanceof Element ? target : target?.parentElement;
    if (!el) return null;

    if (el instanceof HTMLTextAreaElement) {
      if (el.disabled || el.readOnly) return null;
      return el;
    }

    if (el instanceof HTMLInputElement) {
      const allowedTypes = new Set(['text', 'search', 'url', 'email', 'tel', '']);
      if (!allowedTypes.has((el.type || '').toLowerCase()) || el.disabled || el.readOnly) return null;
      return el;
    }

    const editable = el.closest?.('[contenteditable="true"], [contenteditable="plaintext-only"]');
    return editable?.isContentEditable ? editable : null;
  }

  function getEditableText(el) {
    if (el instanceof HTMLInputElement || el instanceof HTMLTextAreaElement) return el.value;
    return el.innerText || el.textContent || '';
  }

  function setNativeTextControlValue(el, value) {
    const proto = el instanceof HTMLTextAreaElement ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
    const setter = Object.getOwnPropertyDescriptor(proto, 'value')?.set;
    if (setter) setter.call(el, value);
    else el.value = value;
    el.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertReplacementText', data: value }));
    el.dispatchEvent(new Event('change', { bubbles: true }));
  }

  function setEditableText(el, value) {
    if (el instanceof HTMLInputElement || el instanceof HTMLTextAreaElement) {
      setNativeTextControlValue(el, value);
      try {
        const end = value.length;
        el.setSelectionRange(end, end);
      } catch {}
      return;
    }

    el.focus();
    const selection = window.getSelection();
    const range = document.createRange();
    range.selectNodeContents(el);
    selection?.removeAllRanges();
    selection?.addRange(range);

    let replaced = false;
    try { replaced = document.execCommand('insertText', false, value); } catch {}
    if (!replaced) {
      el.textContent = value;
      el.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertReplacementText', data: value }));
    }
  }

  function removeTwoTriggerSpaces(el) {
    if (el instanceof HTMLInputElement || el instanceof HTMLTextAreaElement) {
      const start = el.selectionStart;
      const end = el.selectionEnd;
      if (start == null || end == null || start !== end) return;
      const before = el.value.slice(0, start);
      if (!/[ \u00A0]{2}$/.test(before)) return;
      const next = before.slice(0, -2) + el.value.slice(end);
      setNativeTextControlValue(el, next);
      try { el.setSelectionRange(start - 2, start - 2); } catch {}
      return;
    }

    // contenteditable：第三次空格触发前，删除前两次已经输入的空格。
    try {
      document.execCommand('delete', false);
      document.execCommand('delete', false);
    } catch {}
  }

  async function translateEditableToEnglish(el) {
    const original = getEditableText(el).trim();
    if (!original) return;

    const rect = el.getBoundingClientRect();
    const popup = createStatusPopup('正在翻译成英文…', rect, false);

    try {
      const translated = await translateWithBrowser(original, 'en');
      if (!el.isConnected) return;
      setEditableText(el, translated);
      popup.textContent = '已翻译成英文';
      popup.style.color = '#188038';
      setTimeout(() => popup.remove(), 1400);
    } catch (err) {
      console.error('[Word Tool] Input translate to English failed:', err);
      popup.textContent = `翻译失败：${err?.message || String(err)}`;
      popup.style.color = '#d93025';
      setTimeout(() => popup.remove(), 4000);
    }
  }

  function handleTripleSpaceToEnglish(e) {
    const editable = getEditableTarget(e.target);
    if (!editable) return;

    const isPlainSpace = e.key === ' ' || e.code === 'Space';
    const state = tripleSpaceState.get(editable) || { count: 0, lastAt: 0 };

    if (!isPlainSpace || e.ctrlKey || e.altKey || e.metaKey) {
      state.count = 0;
      state.lastAt = 0;
      tripleSpaceState.set(editable, state);
      return;
    }

    const now = Date.now();
    if (now - state.lastAt > TRIPLE_SPACE_MAX_GAP_MS) state.count = 0;
    state.count += 1;
    state.lastAt = now;
    tripleSpaceState.set(editable, state);

    if (state.count < 3) return;

    e.preventDefault();
    e.stopPropagation();
    state.count = 0;
    state.lastAt = 0;
    removeTwoTriggerSpaces(editable);
    translateEditableToEnglish(editable);
  }

  // ==================== 事件 ====================

  let pronounceDebounceTimer = null;
  let isAltDragging = false;

  function debouncedPronounce(word) {
    clearTimeout(pronounceDebounceTimer);
    pronounceDebounceTimer = setTimeout(() => pronounce(word), CONFIG.debounceMs);
  }

  function initEventListeners() {
    // 在输入框 / textarea / contenteditable 中，连续按 3 次空格：整段翻译成英文并原位替换。
    document.addEventListener('keydown', handleTripleSpaceToEnglish, true);

    // ALT + 左键用于框选时，避免链接触发“另存为/下载”等默认行为。
    const blockAltLinkAction = e => {
      if (!e.altKey || e.button !== 0) return;
      const target = e.target instanceof Element ? e.target : e.target?.parentElement;
      const link = target?.closest?.('a[href]');
      if (link) {
        e.preventDefault();
        e.stopPropagation();
      }
    };

    document.addEventListener('mousedown', blockAltLinkAction, true);
    document.addEventListener('click', blockAltLinkAction, true);

    // 防止单独按 Alt 激活浏览器菜单；只在页面拥有焦点时处理。
    document.addEventListener('keydown', e => {
      if (e.key === 'Alt' && e.target === document.body) e.preventDefault();
    }, true);

    document.addEventListener('mousedown', e => {
      if (e.altKey && e.button === 0) {
        isAltDragging = true;
        removeTranslatePopup();
        stopCurrentAudio();
      }
    }, true);

    document.addEventListener('mouseup', e => {
      if (!isAltDragging || e.button !== 0) return;
      isAltDragging = false;

      const selection = getSelectedText();
      if (!selection) return;

      const loadingPopup = createStatusPopup('翻译中…', selection.range, false);
      translateWithBrowser(selection.text)
        .then(translated => createTranslatePopup(translated, selection.range))
        .catch(err => {
          console.error('[Word Tool] Translation failed:', err);
          loadingPopup.textContent = `翻译失败：${err?.message || String(err)}`;
          loadingPopup.style.color = '#d93025';
          setTimeout(() => loadingPopup.remove(), 4000);
        });
    }, true);

    if (CONFIG.triggerMode === 'dblclick') {
      document.addEventListener('dblclick', e => {
        const target = e.target instanceof Element ? e.target : e.target?.parentElement;
        const link = target?.closest?.('a[href]');
        if (link) {
          if (!e.altKey) return;
          e.preventDefault();
          e.stopPropagation();
          e.stopImmediatePropagation();
        }

        let word = getSelectedWord();
        if (!word) word = getWordAtPosition(e.clientX, e.clientY);
        if (word) {
          // 双击模式不延迟，尽量保留 Edge 的 user activation，避免音频被自动播放策略拦截。
          pronounce(word);
        }
      }, true);
    }

    if (CONFIG.triggerMode === 'ctrl+click') {
      document.addEventListener('click', e => {
        if (!e.ctrlKey) return;
        const word = getSelectedWord() || getWordAtPosition(e.clientX, e.clientY);
        if (word) pronounce(word);
      }, true);
    }

    if (CONFIG.triggerMode === 'ctrl+hover') {
      let lastWord = '';
      document.addEventListener('mousemove', e => {
        if (!e.ctrlKey) {
          lastWord = '';
          return;
        }
        const word = getWordAtPosition(e.clientX, e.clientY) || getSelectedWord();
        if (word && word !== lastWord) {
          lastWord = word;
          debouncedPronounce(word);
        }
      }, true);
      document.addEventListener('keyup', e => {
        if (e.key === 'Control') lastWord = '';
      }, true);
    }

    document.addEventListener('keydown', e => {
      if (e.key === 'Escape') {
        removeTranslatePopup();
        stopCurrentAudio();
      }
    }, true);
  }

  // ==================== 设置 ====================

  function showSettings() {
    document.getElementById('wt-settings-modal')?.remove();

    const modal = document.createElement('div');
    modal.id = 'wt-settings-modal';
    modal.innerHTML = `
      <div class="wt-settings-content">
        <h3>Word Tool 设置</h3>

        <div class="wt-section-title">发音设置</div>
        <div class="wt-setting-group">
          <label>触发方式</label>
          <select id="wt-setting-trigger">
            <option value="dblclick" ${CONFIG.triggerMode === 'dblclick' ? 'selected' : ''}>双击选词自动发音</option>
            <option value="ctrl+click" ${CONFIG.triggerMode === 'ctrl+click' ? 'selected' : ''}>CTRL + 单击发音</option>
            <option value="ctrl+hover" ${CONFIG.triggerMode === 'ctrl+hover' ? 'selected' : ''}>CTRL + 悬浮发音</option>
          </select>
        </div>

        <div class="wt-setting-group">
          <label>有道语音类型</label>
          <select id="wt-setting-youdao">
            <option value="2" ${CONFIG.youdaoType === 2 ? 'selected' : ''}>美音</option>
            <option value="0" ${CONFIG.youdaoType === 0 ? 'selected' : ''}>英音</option>
          </select>
        </div>

        <div class="wt-section-title">翻译设置</div>
        <div class="wt-setting-group">
          <label>翻译引擎</label>
          <div style="padding:10px;background:#f5f5f5;border-radius:6px;font-size:13px;color:#666;">
            ${browserName()} 内置 Translator API（本地处理，自动识别英/日/韩 → 简体中文）
          </div>
        </div>

        <div class="wt-setting-group">
          <button id="wt-check-translator" class="wt-settings-btn" style="width:100%;">检测 Translator API</button>
        </div>

        <div class="wt-setting-group">
          <label style="color:#999;font-size:12px;line-height:1.6;">
            整页翻译：点击左下角浮动按钮<br>
            选词翻译：按住 ALT 拖动选中文字
          </label>
        </div>

        <div class="wt-setting-actions">
          <button id="wt-settings-cancel" class="wt-settings-btn">取消</button>
          <button id="wt-settings-save" class="wt-settings-btn wt-primary">保存</button>
        </div>
      </div>
      <div id="wt-settings-overlay"></div>
    `;

    if (!document.getElementById('wt-settings-styles')) {
      const style = document.createElement('style');
      style.id = 'wt-settings-styles';
      style.textContent = `
        #wt-settings-modal{position:fixed;inset:0;z-index:2147483647;display:flex;align-items:center;justify-content:center;}
        #wt-settings-overlay{position:absolute;inset:0;background:rgba(0,0,0,.5);}
        .wt-settings-content{position:relative;z-index:1;background:#fff;padding:24px;border-radius:12px;min-width:340px;max-width:400px;max-height:80vh;overflow-y:auto;box-shadow:0 8px 32px rgba(0,0,0,.2);font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;}
        .wt-settings-content h3{margin:0 0 20px;font-size:18px;color:#333;}
        .wt-section-title{font-size:12px;font-weight:600;color:#1890ff;text-transform:uppercase;letter-spacing:1px;margin:16px 0 8px;padding-bottom:4px;border-bottom:1px solid #f0f0f0;}
        .wt-setting-group{margin-bottom:14px;}
        .wt-setting-group label{display:block;margin-bottom:6px;font-size:13px;color:#666;}
        .wt-setting-group select{width:100%;padding:10px;border:1px solid #ddd;border-radius:6px;font-size:14px;background:#fff;}
        .wt-setting-actions{display:flex;gap:10px;justify-content:flex-end;margin-top:24px;padding-top:16px;border-top:1px solid #eee;}
        .wt-settings-btn{padding:10px 20px;border-radius:6px;font-size:14px;cursor:pointer;border:1px solid #ddd;background:#fff;}
        .wt-settings-btn:hover{background:#f5f5f5;}
        .wt-settings-btn.wt-primary{background:#1890ff;color:#fff;border-color:#1890ff;}
        .wt-settings-btn.wt-primary:hover{background:#40a9ff;}
      `;
      (document.head || document.documentElement).appendChild(style);
    }

    document.body.appendChild(modal);

    document.getElementById('wt-check-translator')?.addEventListener('click', checkTranslatorStatus);
    document.getElementById('wt-settings-cancel')?.addEventListener('click', () => modal.remove());
    document.getElementById('wt-settings-overlay')?.addEventListener('click', () => modal.remove());
    document.getElementById('wt-settings-save')?.addEventListener('click', () => {
      CONFIG.triggerMode = document.getElementById('wt-setting-trigger').value;
      CONFIG.youdaoType = Number(document.getElementById('wt-setting-youdao').value);
      saveConfig();
      modal.remove();
      alert('设置已保存。触发方式修改后请刷新页面生效。');
    });
  }

  // ==================== 初始化 ====================

  function init() {
    loadConfig();
    initEventListeners();
    createFloatingButton();

    GM_registerMenuCommand('Word Tool 设置', showSettings);
    GM_registerMenuCommand('检测 Translator API', checkTranslatorStatus);

    let sandboxMode = 'unknown';
    try { sandboxMode = GM_info?.sandboxMode || 'unknown'; } catch {}

    console.log(
      `[Word Tool] v4.6.0 initialized | browser=${browserName()} | sandbox=${sandboxMode} | Translator=${Boolean(getTranslatorAPI())}`
    );
  }

  function safeInit() {
    if (window.__WORD_TOOL_INITIALIZED__) return;
    window.__WORD_TOOL_INITIALIZED__ = true;
    init();
  }

  if (document.body) safeInit();
  else window.addEventListener('DOMContentLoaded', safeInit, { once: true });
})();
