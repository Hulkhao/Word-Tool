// ==UserScript==
// @name         Word Tool - 选词发音 & 翻译
// @namespace    simpread/word-tool
// @version      4.0.0
// @description  双击单词发音（有道真人美音）+ ALT拖动选词翻译 + 整页翻译（Chrome内置翻译API）
// @match        *://*/*
// @grant        GM_xmlhttpRequest
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_registerMenuCommand
// @connect      dict.youdao.com
// ==/UserScript==

/*
 * ========================================
 * Word Tool v4.0 - 发音 + 选词翻译 + 整页翻译
 * ========================================
 *
 * 功能一：双击发音
 * - 双击单词自动发音（有道真人美音）
 * - 超链接内需按住Alt+双击
 * - 显示美式音标和中文释义
 *
 * 功能二：ALT拖动翻译
 * - 按住ALT键拖动选中文字，自动翻译
 * - 悬浮窗显示翻译结果
 *
 * 功能三：整页翻译
 * - 右下角浮动按钮，一键翻译整个页面
 * - 再次点击取消翻译
 *
 * 翻译引擎：Chrome内置翻译API（本地处理，自动识别语言→中文）
 */

(function() {
  'use strict';

  // ============= 防止 iframe 重复注入 =============
  if (window !== window.top) return;

  // ============= 配置管理 =============
  const CONFIG = {
    triggerMode: 'dblclick',
    youdaoType: 2,
    debounceMs: 200
  };

  function loadConfig() {
    CONFIG.triggerMode = GM_getValue('wt-trigger-mode', CONFIG.triggerMode);
    CONFIG.youdaoType = GM_getValue('wt-youdao-type', CONFIG.youdaoType);
  }

  function saveConfig() {
    GM_setValue('wt-trigger-mode', CONFIG.triggerMode);
    GM_setValue('wt-youdao-type', CONFIG.youdaoType);
  }

  // ============= 音频缓存 =============
  const audioCache = new Map();
  let currentAudio = null;

  function stopCurrentAudio() {
    if (currentAudio) {
      currentAudio.pause();
      currentAudio.currentTime = 0;
      currentAudio = null;
    }
  }

  // ============= 语言检测 =============
  function detectLanguage(text) {
    const t = text.trim();
    if (/[\u3040-\u309F\u30A0-\u30FF]/.test(t)) return 'ja';
    if (/[\uAC00-\uD7AF\u1100-\u11FF]/.test(t)) return 'ko';
    if (/[\u4e00-\u9fff\u3400-\u4DBF]/.test(t)) return 'zh';
    return 'en';
  }

  // ============= Chrome 内置翻译 API =============
  const translators = new Map();
  const translatorLoading = new Set();

  async function getTranslator(sourceLang) {
    if (translators.has(sourceLang)) return translators.get(sourceLang);
    if (translatorLoading.has(sourceLang)) return null;
    if (!window.Translator) return null;

    translatorLoading.add(sourceLang);
    try {
      const availability = await window.Translator.availability({ sourceLanguage: sourceLang, targetLanguage: 'zh' });
      if (availability === 'unavailable') return null;

      const translator = await window.Translator.create({
        sourceLanguage: sourceLang,
        targetLanguage: 'zh',
        ...(availability !== 'available' ? {
          monitor: (monitor) => {
            monitor.addEventListener('downloadprogress', (e) => {
              console.log(`[Word Tool] Downloading ${sourceLang}→zh: ${Math.floor(e.loaded * 100)}%`);
            });
          }
        } : {})
      });
      translators.set(sourceLang, translator);
      return translator;
    } catch (err) {
      console.error(`[Word Tool] Translator init failed (${sourceLang}):`, err);
      return null;
    } finally {
      translatorLoading.delete(sourceLang);
    }
  }

  async function translateWithChrome(text) {
    const sourceLang = detectLanguage(text);
    if (sourceLang === 'zh') return text;
    const translator = await getTranslator(sourceLang);
    if (!translator) throw new Error('Translator not available');
    return await translator.translate(text);
  }

  // ============= 选区工具 =============
  function cleanWord(word) {
    const c = word.trim().toLowerCase().replace(/[^a-z'-]/gi, '');
    if (!c || c.length < 2 || c.length > 50) return null;
    return c;
  }

  function getSelectedWord() {
    const sel = window.getSelection();
    if (!sel || sel.isCollapsed || !sel.rangeCount) return null;
    const text = sel.toString().trim();
    const firstWord = text.split(/\s+/)[0];
    const cleaned = cleanWord(firstWord);
    if (!cleaned) return null;
    const range = sel.getRangeAt(0);
    const container = range.commonAncestorContainer;
    const parentEl = container.nodeType === Node.TEXT_NODE ? container.parentElement : container;
    if (parentEl?.closest('input, textarea, [contenteditable="true"], .notranslate, script, style')) return null;
    return cleaned;
  }

  function getSelectedText() {
    const sel = window.getSelection();
    if (!sel || sel.isCollapsed || !sel.rangeCount) return null;
    const text = sel.toString().trim();
    if (!text || text.length < 2 || text.length > 500) return null;
    const range = sel.getRangeAt(0);
    const container = range.commonAncestorContainer;
    const parentEl = container.nodeType === Node.TEXT_NODE ? container.parentElement : container;
    if (parentEl?.closest('input, textarea, [contenteditable="true"], .notranslate, script, style, code, pre')) return null;
    return { text, range: range.getBoundingClientRect() };
  }

  function getWordAtPosition(x, y) {
    let range = null;
    if (document.caretPositionFromPoint) {
      const pos = document.caretPositionFromPoint(x, y);
      if (pos) { range = document.createRange(); range.setStart(pos.offsetNode, pos.offset); range.setEnd(pos.offsetNode, pos.offset); }
    } else if (document.caretRangeFromPoint) {
      range = document.caretRangeFromPoint(x, y);
    }
    if (!range) return null;
    range.expand('word');
    return cleanWord(range.toString().trim());
  }

  // =============================================
  //  功能一：双击发音
  // =============================================

  function getYoudaoUrl(word) {
    return `https://dict.youdao.com/dictvoice?audio=${encodeURIComponent(word)}&type=${CONFIG.youdaoType}`;
  }

  function checkYoudaoAudio(word) {
    return new Promise((resolve) => {
      GM_xmlhttpRequest({
        method: 'HEAD', url: getYoudaoUrl(word), timeout: 3000,
        onload: (res) => resolve(res.status === 200),
        onerror: () => resolve(false),
        ontimeout: () => resolve(false)
      });
    });
  }

  function fetchWordInfo(word) {
    return new Promise((resolve) => {
      GM_xmlhttpRequest({
        method: 'GET', url: `https://dict.youdao.com/jsonapi?q=${encodeURIComponent(word)}`, timeout: 3000,
        onload: (res) => {
          try {
            const data = JSON.parse(res.responseText);
            const wordInfo = data.ec?.word?.[0];
            if (!wordInfo) { resolve(null); return; }
            const usphone = wordInfo.usphone ? `/${wordInfo.usphone}/` : null;
            const meanings = wordInfo.trs?.slice(0, 3).map(t => t.tr[0].l.i[0]).filter(Boolean) || [];
            resolve({ usphone, meanings });
          } catch (e) { resolve(null); }
        },
        onerror: () => resolve(null),
        ontimeout: () => resolve(null)
      });
    });
  }

  function playFromUrl(url) {
    stopCurrentAudio();
    const audio = new Audio(url);
    currentAudio = audio;
    audio.play().catch(() => { currentAudio = null; });
    audio.onended = () => { currentAudio = null; };
    audio.onerror = () => { currentAudio = null; };
  }

  let pronounceTooltipRemoveHandler = null;

  function createPronounceTooltip(word, icon, wordInfo, isLoading) {
    const existing = document.getElementById('wt-pronounce-tooltip');
    if (existing) existing.remove();
    if (pronounceTooltipRemoveHandler) {
      document.removeEventListener('mousemove', pronounceTooltipRemoveHandler);
      pronounceTooltipRemoveHandler = null;
    }

    const tip = document.createElement('div');
    tip.id = 'wt-pronounce-tooltip';
    let content = '';
    if (wordInfo?.usphone) {
      content += `<div style="font-size:16px;font-weight:600;color:#fff;">${icon} ${wordInfo.usphone}</div>`;
      content += `<div style="font-size:13px;color:#888;margin-top:2px;">${word}</div>`;
    } else {
      content += `<div style="font-size:14px;font-weight:500;">${icon} ${word}</div>`;
    }
    if (wordInfo?.meanings?.length > 0) {
      content += `<div style="font-size:13px;color:#ccc;margin-top:4px;max-width:280px;word-wrap:break-word;">${wordInfo.meanings.join('；')}</div>`;
    }
    if (isLoading) {
      content += `<div style="font-size:11px;color:#888;margin-top:2px;">加载中...</div>`;
    }
    tip.innerHTML = content;
    tip.style.cssText = `position:fixed;z-index:99999;background:#1e1e1e;color:#fff;padding:8px 12px;border-radius:6px;font-size:13px;pointer-events:none;opacity:0;transition:opacity 0.15s;box-shadow:0 2px 8px rgba(0,0,0,0.3);font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;max-width:320px;`;
    document.body.appendChild(tip);

    const sel = window.getSelection();
    if (sel.rangeCount) {
      const rect = sel.getRangeAt(0).getBoundingClientRect();
      tip.style.left = `${rect.left + rect.width / 2 - tip.offsetWidth / 2}px`;
      tip.style.top = `${rect.top - tip.offsetHeight - 8}px`;
    }
    requestAnimationFrame(() => { tip.style.opacity = '1'; });

    let lastMouseX = 0, lastMouseY = 0, initialized = false;
    pronounceTooltipRemoveHandler = (e) => {
      if (!initialized) { lastMouseX = e.clientX; lastMouseY = e.clientY; initialized = true; return; }
      if (Math.abs(e.clientX - lastMouseX) > 5 || Math.abs(e.clientY - lastMouseY) > 5) {
        tip.style.opacity = '0';
        setTimeout(() => tip.remove(), 150);
        document.removeEventListener('mousemove', pronounceTooltipRemoveHandler);
        pronounceTooltipRemoveHandler = null;
      }
    };
    document.addEventListener('mousemove', pronounceTooltipRemoveHandler);
  }

  async function pronounce(word) {
    const cleaned = cleanWord(word);
    if (!cleaned) return;
    createPronounceTooltip(cleaned, '🔊', null, true);
    const wordInfoPromise = fetchWordInfo(cleaned);

    const cached = audioCache.get(cleaned);
    if (cached?.url) {
      playFromUrl(cached.url);
      const wordInfo = await wordInfoPromise;
      if (wordInfo) { cached.usphone = wordInfo.usphone; cached.meanings = wordInfo.meanings; }
      createPronounceTooltip(cleaned, '🔊', cached, false);
      return;
    }

    try {
      const hasYoudao = await checkYoudaoAudio(cleaned);
      if (hasYoudao) {
        const url = getYoudaoUrl(cleaned);
        const wordInfo = await wordInfoPromise;
        const cacheEntry = { url, usphone: wordInfo?.usphone, meanings: wordInfo?.meanings };
        audioCache.set(cleaned, cacheEntry);
        playFromUrl(url);
        createPronounceTooltip(cleaned, '🔊', cacheEntry, false);
        return;
      }
    } catch (err) { console.warn('[Word Tool] Youdao check failed:', err); }

    const wordInfo = await wordInfoPromise;
    const cacheEntry = { url: null, usphone: wordInfo?.usphone, meanings: wordInfo?.meanings };
    createPronounceTooltip(cleaned, '❌', cacheEntry, false);
  }

  // =============================================
  //  功能二：ALT拖动翻译（悬浮窗）
  // =============================================

  function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  }

  let translatePopup = null;

  function removeTranslatePopup() {
    const existing = document.getElementById('wt-translate-popup');
    if (existing) {
      existing.style.opacity = '0';
      existing.style.transform = 'translateY(8px)';
      setTimeout(() => existing.remove(), 200);
    }
    translatePopup = null;
  }

  function createTranslatePopup(translatedText, rect) {
    removeTranslatePopup();
    const popup = document.createElement('div');
    popup.id = 'wt-translate-popup';

    popup.innerHTML = `
      <div class="wt-header">
        <span class="wt-engine">Chrome AI</span>
        <div class="wt-actions">
          <button class="wt-btn wt-close-btn" title="关闭">✕</button>
        </div>
      </div>
      <div class="wt-translated">${escapeHtml(translatedText)}</div>
    `;

    popup.style.cssText = `position:fixed;z-index:2147483647;background:#fff;color:#333;border-radius:12px;padding:12px 16px;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;box-shadow:0 4px 24px rgba(0,0,0,0.15),0 0 0 1px rgba(0,0,0,0.05);max-width:380px;min-width:200px;opacity:0;transform:translateY(8px);transition:opacity 0.2s ease,transform 0.2s ease;cursor:default;user-select:text;font-size:14px;line-height:1.5;`;

    if (!document.getElementById('wt-styles')) {
      const style = document.createElement('style');
      style.id = 'wt-styles';
      style.textContent = `
        .wt-header{display:flex;align-items:center;gap:8px;margin-bottom:8px;}
        .wt-engine{font-size:11px;color:#999;flex:1;}
        .wt-actions{display:flex;gap:4px;}
        .wt-btn{width:28px;height:28px;border:none;background:transparent;border-radius:6px;cursor:pointer;font-size:14px;display:flex;align-items:center;justify-content:center;transition:background 0.15s;}
        .wt-btn:hover{background:#f0f0f0;}
        .wt-translated{color:#333;font-size:15px;font-weight:500;word-break:break-word;padding:4px 0;}
      `;
      document.head.appendChild(style);
    }

    document.body.appendChild(popup);

    const popupRect = popup.getBoundingClientRect();
    let left = rect.left + rect.width / 2 - popupRect.width / 2;
    let top = rect.top - popupRect.height - 12;
    if (left < 8) left = 8;
    if (left + popupRect.width > window.innerWidth - 8) left = window.innerWidth - popupRect.width - 8;
    if (top < 8) top = rect.bottom + 12;
    popup.style.left = `${left}px`;
    popup.style.top = `${top}px`;

    requestAnimationFrame(() => { popup.style.opacity = '1'; popup.style.transform = 'translateY(0)'; });

    popup.querySelector('.wt-close-btn').addEventListener('click', (e) => { e.stopPropagation(); removeTranslatePopup(); });

    const clickOutside = (e) => {
      if (!popup.contains(e.target)) { removeTranslatePopup(); document.removeEventListener('click', clickOutside); }
    };
    setTimeout(() => document.addEventListener('click', clickOutside), 100);

    translatePopup = popup;
  }

  // =============================================
  //  功能三：整页翻译
  // =============================================

  const PAGE_PREFIX = 'wt-page';
  const BILINGUAL_CLASS = `${PAGE_PREFIX}-bilingual`;

  const BLOCK_TAGS = new Set([
    'div', 'p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6',
    'article', 'section', 'header', 'footer', 'main', 'aside', 'nav',
    'blockquote', 'pre', 'li', 'td', 'th', 'tr', 'ul', 'ol'
  ]);

  const EXCLUDE_TAGS = new Set([
    'script', 'style', 'noscript', 'iframe', 'object', 'embed',
    'canvas', 'svg', 'math', 'code', 'pre', 'kbd', 'samp', 'var',
    'textarea', 'input', 'select', 'option', 'template', 'meta', 'link', 'head', 'title'
  ]);

  const EXCLUDE_SELECTORS = [
    '[data-translate="no"]', '.notranslate', '[translate="no"]',
    '[role="tooltip"]', '.tooltip', '.alt-text'
  ];

  function isExcluded(el) {
    if (!el || !el.tagName) return true;
    const tag = el.tagName.toLowerCase();
    if (EXCLUDE_TAGS.has(tag)) return true;
    if (el.contentEditable === 'true') return true;
    if (['input', 'textarea', 'button', 'select'].includes(tag)) return true;
    if (el.closest(`.${BILINGUAL_CLASS}`)) return true;
    return EXCLUDE_SELECTORS.some(sel => {
      try { return el.matches(sel); } catch { return false; }
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
        if (!node.parentElement) return NodeFilter.FILTER_REJECT;
        if (isExcluded(node.parentElement)) return NodeFilter.FILTER_REJECT;
        if (!isVisible(node.parentElement)) return NodeFilter.FILTER_REJECT;
        const text = node.textContent.trim();
        if (!text || text.length < 2) return NodeFilter.FILTER_REJECT;
        if (/^[\d\W]*$/.test(text)) return NodeFilter.FILTER_REJECT;
        return NodeFilter.FILTER_ACCEPT;
      }
    });
    let node;
    while (node = walker.nextNode()) { nodes.push(node); }
    return nodes;
  }

  function groupByParagraph(textNodes) {
    const groups = new Map();
    textNodes.forEach(node => {
      const container = findBlockContainer(node.parentElement);
      const key = getElementPath(container);
      if (!groups.has(key)) groups.set(key, { container, nodes: [] });
      groups.get(key).nodes.push(node);
    });
    return Array.from(groups.values());
  }

  function getElementPath(el) {
    const path = [];
    let current = el;
    while (current && current !== document.body) {
      const tag = current.tagName?.toLowerCase() || 'unknown';
      const siblings = Array.from(current.parentElement?.children || []).filter(c => c.tagName?.toLowerCase() === tag);
      const idx = siblings.indexOf(current);
      path.unshift(`${tag}${siblings.length > 1 ? `[${idx}]` : ''}`);
      current = current.parentElement;
    }
    return path.join('>');
  }

  const translatedElements = new Set();

  async function translateGroup(group) {
    const { container, nodes } = group;
    const combinedText = nodes.map(n => n.textContent.trim()).join(' ');
    if (!combinedText || combinedText.length < 3) return;
    const lang = detectLanguage(combinedText);
    if (lang === 'zh') return;

    try {
      const translated = await translateWithChrome(combinedText);
      if (!translated || translated === combinedText) return;

      const bilingualEl = document.createElement('div');
      bilingualEl.className = BILINGUAL_CLASS;
      bilingualEl.style.cssText = 'margin:6px 0;padding:6px 10px;background:rgba(0,200,100,0.04);font-size:0.9em;color:#444;line-height:1.6;opacity:0;transition:opacity 0.3s;border-radius:4px;';
      bilingualEl.textContent = translated;

      if (container.classList) container.classList.add(BILINGUAL_CLASS);
      container.appendChild(bilingualEl);
      requestAnimationFrame(() => { bilingualEl.style.opacity = '1'; });
      translatedElements.add(bilingualEl);
    } catch (err) {
      console.warn('[Word Tool] Page translate failed:', err);
    }
  }

  let isPageTranslating = false;
  let pageTranslationQueue = [];
  let pageProcessing = false;

  async function processPageQueue() {
    if (pageProcessing) return;
    pageProcessing = true;
    while (pageTranslationQueue.length > 0 && isPageTranslating) {
      const group = pageTranslationQueue.shift();
      await translateGroup(group);
    }
    pageProcessing = false;
  }

  async function startPageTranslation() {
    if (isPageTranslating) return;
    isPageTranslating = true;
    updateFloatingButton(true);

    const textNodes = extractTextNodes(document.body);
    const groups = groupByParagraph(textNodes);

    const observer = new IntersectionObserver((entries) => {
      entries.forEach(entry => {
        if (entry.isIntersecting && isPageTranslating) {
          const group = groups.find(g => g.container === entry.target);
          if (group && !translatedElements.has(group.container)) {
            pageTranslationQueue.push(group);
            processPageQueue();
          }
          observer.unobserve(entry.target);
        }
      });
    }, { rootMargin: '200px' });

    groups.forEach(g => {
      if (!translatedElements.has(g.container)) observer.observe(g.container);
    });
  }

  function stopPageTranslation() {
    isPageTranslating = false;
    pageTranslationQueue = [];
    pageProcessing = false;
    updateFloatingButton(false);

    translatedElements.forEach(el => {
      el.style.opacity = '0';
      setTimeout(() => el.remove(), 300);
    });
    translatedElements.clear();

    document.querySelectorAll(`.${BILINGUAL_CLASS}`).forEach(el => {
      if (el.classList.contains(BILINGUAL_CLASS) && !el.className.includes(PAGE_PREFIX)) {
        el.classList.remove(BILINGUAL_CLASS);
      }
    });
  }

  function togglePageTranslation() {
    if (isPageTranslating) stopPageTranslation();
    else startPageTranslation();
  }

  // ============= 浮动翻译按钮 =============
  let floatingBtn = null;

  function createFloatingButton() {
    if (floatingBtn) return;
    floatingBtn = document.createElement('div');
    floatingBtn.id = `${PAGE_PREFIX}-btn`;
    floatingBtn.innerHTML = `<svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor"><path d="M12.87 15.07l-2.54-2.51.03-.03c1.74-1.94 2.98-4.17 3.71-6.53H17V4h-7V2H8v2H1v1.99h11.17C11.5 7.92 10.44 9.75 9 11.35 8.07 10.32 7.3 9.19 6.69 8h-2c.73 1.63 1.73 3.17 2.98 4.56l-5.09 5.02L4 19l5-5 3.11 3.11.76-2.04zM18.5 10h-2L12 22h2l1.12-3h4.75L21 22h2l-4.5-12zm-2.62 7l1.62-4.33L19.12 17h-3.24z"/></svg>`;
    floatingBtn.style.cssText = `position:fixed;bottom:24px;right:24px;z-index:2147483647;width:32px;height:32px;border-radius:50%;background:#1a73e8;color:#fff;display:flex;align-items:center;justify-content:center;cursor:pointer;box-shadow:0 2px 8px rgba(0,0,0,0.25);transition:background 0.2s ease;`;
    floatingBtn.addEventListener('mouseenter', () => { if (!isDragging) floatingBtn.style.transform = 'scale(1.15)'; });
    floatingBtn.addEventListener('mouseleave', () => { if (!isDragging) floatingBtn.style.transform = 'scale(1)'; });

    // 拖动支持
    let isDragging = false;
    let hasMoved = false;
    let startX, startY, startLeft, startTop;

    floatingBtn.addEventListener('mousedown', (e) => {
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
    });

    document.addEventListener('mousemove', (e) => {
      if (!isDragging) return;
      const dx = e.clientX - startX;
      const dy = e.clientY - startY;
      if (Math.abs(dx) > 3 || Math.abs(dy) > 3) hasMoved = true;
      if (!hasMoved) return;
      const newLeft = Math.max(0, Math.min(window.innerWidth - 32, startLeft + dx));
      const newTop = Math.max(0, Math.min(window.innerHeight - 32, startTop + dy));
      floatingBtn.style.left = newLeft + 'px';
      floatingBtn.style.top = newTop + 'px';
      floatingBtn.style.right = 'auto';
      floatingBtn.style.bottom = 'auto';
    });

    document.addEventListener('mouseup', () => {
      if (!isDragging) return;
      isDragging = false;
      floatingBtn.style.transition = 'background 0.2s ease, transform 0.15s ease';
      if (!hasMoved) togglePageTranslation();
    });

    document.body.appendChild(floatingBtn);
  }

  function updateFloatingButton(translating) {
    if (!floatingBtn) return;
    floatingBtn.style.background = translating ? '#34a853' : '#1a73e8';
    floatingBtn.title = translating ? '点击停止翻译' : '点击翻译此页';
  }

  // =============================================
  //  事件绑定
  // =============================================

  let pronounceDebounceTimer = null;

  function debouncedPronounce(word) {
    clearTimeout(pronounceDebounceTimer);
    pronounceDebounceTimer = setTimeout(() => pronounce(word), CONFIG.debounceMs);
  }

  let isAltDragging = false;

  function initEventListeners() {
    // 阻止Alt+单击下载
    const blockAltDownload = (e) => {
      if (e.altKey && e.button === 0) {
        const linkParent = e.target.closest('a[href]');
        if (linkParent) { e.preventDefault(); e.stopPropagation(); }
      }
    };
    document.addEventListener('mousedown', blockAltDownload, true);
    document.addEventListener('click', blockAltDownload, true);

    document.addEventListener('keydown', (e) => { if (e.key === 'Alt') e.preventDefault(); }, true);

    // ALT拖动翻译
    document.addEventListener('mousedown', (e) => {
      if (e.altKey && e.button === 0) {
        isAltDragging = true;
        removeTranslatePopup();
        stopCurrentAudio();
      }
    }, true);

    document.addEventListener('mouseup', (e) => {
      if (isAltDragging && e.button === 0) {
        isAltDragging = false;
        const selection = getSelectedText();
        if (selection) {
          const loadingPopup = document.createElement('div');
          loadingPopup.id = 'wt-translate-popup';
          loadingPopup.style.cssText = `position:fixed;z-index:2147483647;background:#fff;color:#999;font-size:13px;padding:8px 12px;border-radius:8px;box-shadow:0 4px 16px rgba(0,0,0,0.12);left:${selection.range.left}px;top:${selection.range.top - 40}px;`;
          loadingPopup.textContent = '翻译中...';
          document.body.appendChild(loadingPopup);

          translateWithChrome(selection.text).then(translated => {
            createTranslatePopup(translated, selection.range);
          }).catch(err => {
            console.error('[Word Tool] Translation failed:', err);
            loadingPopup.textContent = '翻译失败';
            loadingPopup.style.color = '#ff4d4f';
            setTimeout(() => loadingPopup.remove(), 2000);
          });
        }
      }
    }, true);

    // 双击发音
    if (CONFIG.triggerMode === 'dblclick') {
      const handler = (e) => {
        const linkParent = e.target.closest('a[href]');
        if (linkParent) {
          if (!e.altKey) return;
          e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation();
        }
        let word = getSelectedWord();
        if (!word) word = getWordAtPosition(e.clientX, e.clientY);
        if (word) debouncedPronounce(word);
      };
      document.addEventListener('dblclick', handler);
    }

    if (CONFIG.triggerMode === 'ctrl+click') {
      document.addEventListener('click', (e) => {
        if (e.ctrlKey) { const word = getSelectedWord(); if (word) debouncedPronounce(word); }
      });
    }

    if (CONFIG.triggerMode === 'ctrl+hover') {
      let lastWord = '';
      document.addEventListener('mousemove', (e) => {
        if (!e.ctrlKey) { lastWord = ''; return; }
        const word = getSelectedWord();
        if (word && word !== lastWord) { lastWord = word; debouncedPronounce(word); }
      });
      document.addEventListener('keyup', (e) => { if (e.key === 'Control') lastWord = ''; });
    }

    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') { removeTranslatePopup(); stopCurrentAudio(); }
    });
  }

  // =============================================
  //  设置界面
  // =============================================

  function showSettings() {
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
            <option value="ctrl+click" ${CONFIG.triggerMode === 'ctrl+click' ? 'selected' : ''}>CTRL+单击发音</option>
            <option value="ctrl+hover" ${CONFIG.triggerMode === 'ctrl+hover' ? 'selected' : ''}>CTRL+悬浮发音</option>
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
            Chrome 内置翻译 API（本地处理，自动识别语言→中文）
          </div>
        </div>

        <div class="wt-setting-group">
          <label style="color:#999;font-size:12px;">
            整页翻译：点击右下角浮动按钮<br>
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
        #wt-settings-modal{position:fixed;top:0;left:0;right:0;bottom:0;z-index:2147483647;display:flex;align-items:center;justify-content:center;}
        #wt-settings-overlay{position:absolute;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.5);}
        .wt-settings-content{position:relative;z-index:1;background:#fff;padding:24px;border-radius:12px;min-width:340px;max-width:400px;max-height:80vh;overflow-y:auto;box-shadow:0 8px 32px rgba(0,0,0,0.2);font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;}
        .wt-settings-content h3{margin:0 0 20px 0;font-size:18px;color:#333;}
        .wt-section-title{font-size:12px;font-weight:600;color:#1890ff;text-transform:uppercase;letter-spacing:1px;margin:16px 0 8px 0;padding-bottom:4px;border-bottom:1px solid #f0f0f0;}
        .wt-setting-group{margin-bottom:14px;}
        .wt-setting-group label{display:block;margin-bottom:6px;font-size:13px;color:#666;}
        .wt-setting-group select{width:100%;padding:10px;border:1px solid #ddd;border-radius:6px;font-size:14px;background:#fff;}
        .wt-setting-actions{display:flex;gap:10px;justify-content:flex-end;margin-top:24px;padding-top:16px;border-top:1px solid #eee;}
        .wt-settings-btn{padding:10px 20px;border-radius:6px;font-size:14px;cursor:pointer;border:1px solid #ddd;background:#fff;transition:all 0.15s;}
        .wt-settings-btn:hover{background:#f5f5f5;}
        .wt-settings-btn.wt-primary{background:#1890ff;color:#fff;border-color:#1890ff;}
        .wt-settings-btn.wt-primary:hover{background:#40a9ff;}
      `;
      document.head.appendChild(style);
    }

    document.body.appendChild(modal);

    document.getElementById('wt-settings-cancel').addEventListener('click', () => modal.remove());
    document.getElementById('wt-settings-overlay').addEventListener('click', () => modal.remove());
    document.getElementById('wt-settings-save').addEventListener('click', () => {
      CONFIG.triggerMode = document.getElementById('wt-setting-trigger').value;
      CONFIG.youdaoType = parseInt(document.getElementById('wt-setting-youdao').value);
      saveConfig();
      modal.remove();
    });
  }

  // ============= 初始化 =============
  function init() {
    loadConfig();
    initEventListeners();
    createFloatingButton();
    GM_registerMenuCommand('Word Tool 设置', showSettings);
    console.log('[Word Tool] v4.0 initialized');
  }

  init();
})();
