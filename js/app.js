// ==================== كود الواجهة والتفاعل ====================
// دوال الحساب الصافية في js/calc.js — هذا الملف مسؤول عن DOM فقط

// ==================== UTILITIES ====================
// تهريب HTML لأي بيانات تُحقن في innerHTML (حماية من XSS)
function escapeHtml(str) {
    return String(str)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

// قراءة آمنة من localStorage — قيمة فاسدة لا تعطّل التطبيق
function safeParse(key, fallback) {
    try {
        const value = JSON.parse(localStorage.getItem(key) || JSON.stringify(fallback));
        return Array.isArray(fallback) && !Array.isArray(value) ? fallback : value;
    } catch (e) {
        return fallback;
    }
}

function showToast(msg, type = '') {
    const t = document.getElementById('toast');
    t.textContent = msg;
    t.className = 'toast show' + (type ? ' ' + type : '');
    setTimeout(() => t.classList.remove('show'), 3000);
}

function debounce(fn, wait) {
    let t;
    return function (...args) {
        clearTimeout(t);
        t = setTimeout(() => fn.apply(this, args), wait);
    };
}

// ==================== STATE ====================
let currentCategory = null;
let currentTemplate = null;
let currentTemplateIdx = null;
let favorites = safeParse('courtFavorites', []);
let viewHistory = safeParse('courtHistory', []);

// ==================== INITIALIZATION ====================
document.addEventListener('DOMContentLoaded', function () {
    migrateStorageKeys();
    loadCategories();
    setupSearch();
    setupQuickTafqit();
    setupDelegatedListeners();
    loadFavorites();
    loadHistory();

    if (localStorage.getItem('courtTheme') === 'dark') {
        document.body.setAttribute('data-theme', 'dark');
    }

    // Count total templates
    let total = 0;
    for (let cat in templatesData) {
        total += templatesData[cat].length;
    }
    document.getElementById('totalCount').textContent = total;
});

// ==================== TEMPLATE KEYS (مفاتيح ثابتة للمفضلة والسجل) ====================
// المفتاح يعتمد رقم النموذج (num) لا موقعه في المصفوفة،
// حتى لا تفسد المفضلة عند إعادة ترتيب النماذج في ملف البيانات
function getTemplateNum(t, idx) {
    return t && t.num != null && t.num !== '' ? String(t.num) : String(idx + 1);
}

function getFavKey(cat, num) {
    return `${cat}::${num}`;
}

function findTemplateIndexByNum(cat, num) {
    const arr = templatesData[cat];
    if (!arr) return -1;
    return arr.findIndex((t, i) => getTemplateNum(t, i) === String(num));
}

// ترحيل المفاتيح القديمة (تصنيف::فهرس) إلى الصيغة الجديدة (تصنيف::رقم) لمرة واحدة
function migrateStorageKeys() {
    if (localStorage.getItem('courtKeysVersion') === '2') return;
    const migrate = key => {
        const [cat, idxStr] = String(key).split('::');
        const arr = templatesData[cat];
        const idx = parseInt(idxStr, 10);
        const t = arr && arr[idx];
        if (!t) return null;
        return getFavKey(cat, getTemplateNum(t, idx));
    };
    favorites = favorites.map(migrate).filter(Boolean);
    viewHistory = viewHistory.map(migrate).filter(Boolean);
    localStorage.setItem('courtFavorites', JSON.stringify(favorites));
    localStorage.setItem('courtHistory', JSON.stringify(viewHistory));
    localStorage.setItem('courtKeysVersion', '2');
}

// ==================== DELEGATED LISTENERS ====================
// معالجة النقر عبر data-* بدل حقن onclick داخل HTML مبني من بيانات (حماية من XSS)
function setupDelegatedListeners() {
    document.getElementById('categoryList').addEventListener('click', function (e) {
        const btn = e.target.closest('.category-header[data-cat]');
        if (btn) selectCategory(btn.dataset.cat, btn);
    });

    document.getElementById('templatesList').addEventListener('click', function (e) {
        const el = e.target.closest('[data-action]');
        if (!el) return;
        const cat = el.dataset.cat || currentCategory;
        const idx = parseInt(el.dataset.idx, 10);
        const templates = templatesData[cat];
        if (!templates || !templates[idx]) return;

        if (el.dataset.action === 'copy') {
            copyToClipboard(templates[idx].content);
        } else if (el.dataset.action === 'view') {
            currentCategory = cat;
            viewTemplate(idx);
        }
    });

    const detailContent = document.getElementById('detailContent');
    detailContent.addEventListener('click', function (e) {
        const ph = e.target.closest('.placeholder');
        if (ph) editPlaceholder(ph);
    });
    detailContent.addEventListener('keydown', function (e) {
        if ((e.key === 'Enter' || e.key === ' ') && e.target.classList.contains('placeholder')) {
            e.preventDefault();
            editPlaceholder(e.target);
        }
    });

    document.getElementById('favoritesList').addEventListener('click', function (e) {
        const rm = e.target.closest('.remove-btn[data-key]');
        if (rm) {
            removeFavorite(rm.dataset.key);
            return;
        }
        const open = e.target.closest('.fav-open[data-key]');
        if (open) openFromKey(open.dataset.key);
    });

    document.getElementById('historyList').addEventListener('click', function (e) {
        const open = e.target.closest('.history-item[data-key]');
        if (open) openFromKey(open.dataset.key);
    });
}

// ==================== CATEGORIES ====================
function loadCategories() {
    const list = document.getElementById('categoryList');
    list.innerHTML = '';

    for (let cat in templatesData) {
        if (templatesData[cat].length === 0) continue;

        const li = document.createElement('li');
        li.className = 'category-item';

        li.innerHTML = `
            <button type="button" class="category-header" data-cat="${escapeHtml(cat)}">
                <span>${escapeHtml(cat)}</span>
                <span class="category-count">${templatesData[cat].length}</span>
            </button>
        `;
        list.appendChild(li);
    }
}

function selectCategory(cat, btnEl) {
    currentCategory = cat;

    // Update active state
    document.querySelectorAll('.category-header').forEach(h => h.classList.remove('active'));
    if (btnEl) {
        btnEl.classList.add('active');
    }

    document.getElementById('templatesList').style.display = 'block';
    // Show templates list
    showTemplatesList(templatesData[cat], cat);

    // Hide detail view
    document.getElementById('templateDetail').classList.remove('show');

    // Close mobile sidebar
    document.getElementById('sidebar').classList.remove('show');
}

// ==================== TEMPLATES LIST ====================
function showTemplatesList(templates, categoryName) {
    const container = document.getElementById('templatesList');

    if (!templates || templates.length === 0) {
        container.innerHTML = `
            <div class="empty-state">
                <div class="empty-state-icon">📭</div>
                <h3>لا توجد قوالب في هذا التصنيف</h3>
            </div>
        `;
        return;
    }

    let html = `
        <div class="content-header">
            <h2>${escapeHtml(categoryName)}</h2>
            <span>${templates.length} نموذج</span>
        </div>
        <div style="overflow-x: auto;">
            <table class="templates-table">
                <thead>
                    <tr>
                        <th style="width: 50px;">#</th>
                        <th style="width: 120px;">الاختصار</th>
                        <th>المحتوى</th>
                        <th style="width: 80px;">نسخ</th>
                    </tr>
                </thead>
                <tbody>
    `;

    templates.forEach((t, idx) => {
        const num = t.num || (idx + 1);
        const keyword = t.keyword || '-';
        const preview = t.content.substring(0, 80) + (t.content.length > 80 ? '...' : '');

        html += `
            <tr>
                <td class="template-num">${escapeHtml(num)}</td>
                <td class="template-keyword"><button type="button" class="template-keyword-btn" data-action="view" data-idx="${idx}">${escapeHtml(keyword)}</button></td>
                <td class="template-preview" data-action="view" data-idx="${idx}" style="cursor: pointer;">${escapeHtml(preview)}</td>
                <td class="template-actions-cell">
                    <button type="button" class="copy-btn" data-action="copy" data-idx="${idx}">نسخ</button>
                </td>
            </tr>
        `;
    });

    html += '</tbody></table></div>';
    container.innerHTML = html;
}

// ==================== TEMPLATE DETAIL ====================
function viewTemplate(idx) {
    const templates = templatesData[currentCategory];
    if (!templates || !templates[idx]) return;

    currentTemplate = templates[idx];
    currentTemplateIdx = idx;

    const keyword = currentTemplate.keyword || currentCategory;
    document.getElementById('detailTitle').textContent = keyword;

    let content = formatContent(currentTemplate.content);
    document.getElementById('detailContent').innerHTML = content;

    // Update favorite button
    updateFavButton();

    // Add to history
    addToHistory(currentCategory, getTemplateNum(currentTemplate, idx));

    document.getElementById('templatesList').style.display = 'none';
    document.getElementById('templateDetail').classList.add('show');
}

function backToList() {
    document.getElementById('templateDetail').classList.remove('show');
    document.getElementById('templatesList').style.display = 'block';
}

function formatContent(content) {
    // التهريب أولاً ثم إبراز الحقول القابلة للتعبئة — لمنع تنفيذ أي HTML وارد في البيانات
    return escapeHtml(content)
        .replace(/(\d{2}\s*\/\s*\d{2}\s*\/\s*\d{4})/g, '<span class="placeholder" role="button" tabindex="0">$1</span>')
        .replace(/(0{5,})/g, '<span class="placeholder" role="button" tabindex="0">$1</span>')
        .replace(/(\.{5,})/g, '<span class="placeholder" role="button" tabindex="0">$1</span>')
        .replace(/\(\s*اسم القاضي\s*\)/g, '<span class="placeholder" role="button" tabindex="0">(اسم القاضي)</span>');
}

function editPlaceholder(el) {
    const newValue = prompt('أدخل القيمة الجديدة:', el.textContent);
    if (newValue !== null) {
        el.textContent = newValue;
        el.classList.remove('placeholder');
        el.removeAttribute('role');
        el.removeAttribute('tabindex');
    }
}

// ==================== COPY ====================
function copyCurrentTemplate() {
    if (!currentTemplate) return;
    const content = document.getElementById('detailContent').innerText;
    copyToClipboard(content);
}

function copyToClipboard(text) {
    navigator.clipboard.writeText(text).then(() => {
        showToast('تم نسخ النموذج بنجاح ✓', 'success');
    }).catch(() => {
        const ta = document.createElement('textarea');
        ta.value = text;
        document.body.appendChild(ta);
        ta.select();
        document.execCommand('copy');
        document.body.removeChild(ta);
        showToast('تم نسخ النموذج بنجاح ✓', 'success');
    });
}

function printTemplate() {
    window.print();
}

// ==================== FAVORITES ====================
function toggleFavorite() {
    if (!currentCategory || currentTemplateIdx === null) return;

    const num = getTemplateNum(currentTemplate, currentTemplateIdx);
    const key = getFavKey(currentCategory, num);
    const index = favorites.indexOf(key);

    if (index > -1) {
        favorites.splice(index, 1);
        showToast('تم إزالة النموذج من المفضلة', '');
    } else {
        favorites.push(key);
        showToast('تم إضافة النموذج للمفضلة ⭐', 'success');
    }

    localStorage.setItem('courtFavorites', JSON.stringify(favorites));
    updateFavButton();
    loadFavorites();
}

function updateFavButton() {
    const btn = document.getElementById('favBtn');
    if (!btn || !currentCategory || currentTemplateIdx === null) return;

    const key = getFavKey(currentCategory, getTemplateNum(currentTemplate, currentTemplateIdx));
    const isFav = favorites.includes(key);
    btn.textContent = isFav ? '⭐' : '☆';
    btn.setAttribute('aria-pressed', String(isFav));
}

function loadFavorites() {
    const list = document.getElementById('favoritesList');
    if (!list) return;

    if (favorites.length === 0) {
        list.innerHTML = `<div class="empty-state" style="padding: 1rem; font-size: 0.9rem;"><p>لا توجد قوالب مفضلة</p></div>`;
        return;
    }

    list.innerHTML = favorites.map(key => {
        const [cat, num] = String(key).split('::');
        const idx = findTemplateIndexByNum(cat, num);
        const template = idx > -1 ? templatesData[cat][idx] : null;
        if (!template) return '';

        const title = template.keyword || cat;
        return `
            <div class="favorite-item">
                <button type="button" class="fav-open" data-key="${escapeHtml(key)}">${escapeHtml(title)}</button>
                <button type="button" class="remove-btn" data-key="${escapeHtml(key)}" aria-label="إزالة ${escapeHtml(title)} من المفضلة">✕</button>
            </div>
        `;
    }).filter(Boolean).join('');
}

function removeFavorite(key) {
    const index = favorites.indexOf(key);
    if (index > -1) {
        favorites.splice(index, 1);
        localStorage.setItem('courtFavorites', JSON.stringify(favorites));
        loadFavorites();
        updateFavButton();
    }
}

function openFromKey(key) {
    const [cat, num] = String(key).split('::');
    const idx = findTemplateIndexByNum(cat, num);
    if (idx === -1) return;
    currentCategory = cat;
    viewTemplate(idx);
}

// ==================== HISTORY ====================
function addToHistory(cat, num) {
    const key = getFavKey(cat, num);

    // Remove if exists
    viewHistory = viewHistory.filter(k => k !== key);

    // Add to beginning
    viewHistory.unshift(key);

    // Keep only last 15
    viewHistory = viewHistory.slice(0, 15);

    localStorage.setItem('courtHistory', JSON.stringify(viewHistory));
    loadHistory();
}

function loadHistory() {
    const list = document.getElementById('historyList');
    if (!list) return;

    if (viewHistory.length === 0) {
        list.innerHTML = `<div class="empty-state" style="padding: 1rem; font-size: 0.9rem;"><p>لا يوجد سجل</p></div>`;
        return;
    }

    list.innerHTML = viewHistory.map(key => {
        const [cat, num] = String(key).split('::');
        const idx = findTemplateIndexByNum(cat, num);
        const template = idx > -1 ? templatesData[cat][idx] : null;
        if (!template) return '';

        const title = template.keyword || cat;
        return `<button type="button" class="history-item" data-key="${escapeHtml(key)}"><span>${escapeHtml(title)}</span></button>`;
    }).filter(Boolean).join('');
}

function clearHistory() {
    viewHistory = [];
    localStorage.setItem('courtHistory', JSON.stringify(viewHistory));
    loadHistory();
}

// ==================== SEARCH ====================
function setupSearch() {
    document.getElementById('searchInput').addEventListener('input', debounce(function () {
        const query = this.value.trim().toLowerCase();
        if (query.length < 2) {
            if (currentCategory) {
                showTemplatesList(templatesData[currentCategory], currentCategory);
            }
            return;
        }

        // Search all categories (مقارنة غير حساسة لحالة الأحرف من الجهتين)
        const results = [];
        for (let cat in templatesData) {
            templatesData[cat].forEach((t, idx) => {
                if (t.content.toLowerCase().includes(query) ||
                    (t.keyword && t.keyword.toLowerCase().includes(query)) ||
                    (t.num && String(t.num).toLowerCase().includes(query))) {
                    results.push({ ...t, category: cat, originalIdx: idx });
                }
            });
        }

        showSearchResults(results, query);
    }, 300));
}

function showSearchResults(results, query) {
    const container = document.getElementById('templatesList');

    if (results.length === 0) {
        container.innerHTML = `
            <div class="empty-state">
                <div class="empty-state-icon">🔍</div>
                <h3>لا توجد نتائج لـ "${escapeHtml(query)}"</h3>
            </div>
        `;
        return;
    }
    let html = `
        <div class="content-header">
            <h2>نتائج البحث</h2>
            <span>${results.length} نتيجة</span>
        </div>
        <div style="overflow-x: auto;">
            <table class="templates-table">
                <thead>
                    <tr>
                        <th>#</th>
                        <th>الرقم</th>
                        <th>الاختصار</th>
                        <th>المحتوى</th>
                        <th>نسخ</th>
                    </tr>
                </thead>
                <tbody>
    `;
    results.forEach((t, idx) => {
        const preview = t.content.substring(0, 70) + '...';
        const templateIdx = t.originalIdx;
        html += `
            <tr data-action="view" data-cat="${escapeHtml(t.category)}" data-idx="${templateIdx}" style="cursor: pointer;">
                <td class="template-num">${idx + 1}</td>
                <td class="template-num">${escapeHtml(t.num || '-')}</td>
                <td class="template-keyword"><button type="button" class="template-keyword-btn" data-action="view" data-cat="${escapeHtml(t.category)}" data-idx="${templateIdx}">${escapeHtml(t.keyword || '-')}</button></td>
                <td class="template-preview">${escapeHtml(preview)}</td>
                <td class="template-actions-cell">
                    <button type="button" class="copy-btn" data-action="copy" data-cat="${escapeHtml(t.category)}" data-idx="${templateIdx}">نسخ</button>
                </td>
            </tr>
        `;
    });
    html += '</tbody></table></div>';
    container.innerHTML = html;
}

// ==================== THEME ====================
function toggleTheme() {
    const isDark = document.body.getAttribute('data-theme') === 'dark';
    document.body.setAttribute('data-theme', isDark ? '' : 'dark');
    localStorage.setItem('courtTheme', isDark ? '' : 'dark');
}

// ==================== SIDEBAR ====================
function toggleSidebar() {
    document.getElementById('sidebar').classList.toggle('show');
}

// ==================== MODALS (مع إدارة التركيز للوصولية) ====================
let lastFocusedElement = null;

function getFocusable(container) {
    return Array.from(container.querySelectorAll(
        'button, [href], input, select, textarea, [tabindex]:not([tabindex="-1"])'
    )).filter(el => !el.disabled && el.offsetParent !== null);
}

function openModal(id) {
    const overlay = document.getElementById(id);
    if (!overlay) return;
    lastFocusedElement = document.activeElement;
    overlay.classList.add('show');
    const focusable = getFocusable(overlay);
    if (focusable.length) focusable[0].focus();
}

function closeModal(id) {
    const overlay = document.getElementById(id);
    if (!overlay) return;
    overlay.classList.remove('show');
    if (lastFocusedElement && document.body.contains(lastFocusedElement)) {
        lastFocusedElement.focus();
    }
}

// Escape يغلق المودال/القائمة/الدرج المفتوح، وTab محصور داخل المودال المفتوح
document.addEventListener('keydown', function (e) {
    const openOverlay = document.querySelector('.modal-overlay.show');

    if (e.key === 'Escape') {
        if (openOverlay) {
            closeModal(openOverlay.id);
            return;
        }
        const dropdown = document.getElementById('calcDropdown');
        if (dropdown && dropdown.classList.contains('show')) {
            dropdown.classList.remove('show');
        }
        return;
    }

    if (e.key === 'Tab' && openOverlay) {
        const focusable = getFocusable(openOverlay);
        if (!focusable.length) return;
        const first = focusable[0];
        const last = focusable[focusable.length - 1];
        if (e.shiftKey && document.activeElement === first) {
            e.preventDefault();
            last.focus();
        } else if (!e.shiftKey && document.activeElement === last) {
            e.preventDefault();
            first.focus();
        }
    }
});

// ==================== DROPDOWN ====================
function toggleDropdown(e) {
    e.stopPropagation();
    document.getElementById('calcDropdown').classList.toggle('show');
}

function selectCalc(modalId) {
    document.getElementById('calcDropdown').classList.remove('show');
    openModal(modalId);
}

function openMobileCalcs() {
    openModal('mobileCalcsModal');
}

function selectMobileCalc(modalId) {
    closeModal('mobileCalcsModal');
    setTimeout(() => openModal(modalId), 200);
}

// إغلاق القائمة المنسدلة عند الضغط خارجها
document.addEventListener('click', function (e) {
    const dropdown = document.getElementById('calcDropdown');
    if (dropdown && !dropdown.contains(e.target)) {
        dropdown.classList.remove('show');
    }
});

document.querySelectorAll('.modal-overlay').forEach(overlay => {
    overlay.addEventListener('click', function (e) {
        if (e.target === this) closeModal(this.id);
    });
});

function switchTab(tabEl, tabId) {
    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
    tabEl.classList.add('active');
    document.getElementById(tabId).classList.add('active');
}

// ==================== TOOLS ====================
function toggleTool(header) {
    const body = header.nextElementSibling;
    body.classList.toggle('collapsed');
    if (header.tagName === 'BUTTON') {
        header.setAttribute('aria-expanded', String(!body.classList.contains('collapsed')));
    }
}

// ==================== TAFQIT (Number to Text) ====================
function convertToText() {
    const num = parseFloat(document.getElementById('tafqitNumber').value);
    const currency = document.getElementById('tafqitCurrency').value;
    const style = document.getElementById('tafqitStyle').value;
    const resultDiv = document.getElementById('tafqitResult');

    if (isNaN(num) || num < 0) {
        resultDiv.textContent = 'أدخل رقماً صحيحاً';
        return;
    }

    const intPart = Math.floor(num);
    const decPart = Math.round((num - intPart) * 100);

    let result = numberToArabicWords(intPart, style);

    const curr = currencies[currency];
    if (currency !== 'NONE') {
        if (intPart === 1) result += ' ' + curr.singular;
        else if (intPart === 2) result = result.replace('اثنان', '') + curr.dual;
        else if (intPart >= 3 && intPart <= 10) result += ' ' + curr.plural;
        else result += ' ' + curr.singular;

        if (decPart > 0) {
            result += ' و' + numberToArabicWords(decPart, style);
            if (decPart === 1) result += ' ' + curr.subunit;
            else if (decPart === 2) result += ' ' + curr.subunit + 'ان';
            else if (decPart >= 3 && decPart <= 10) result += ' ' + curr.subunitPlural;
            else result += ' ' + curr.subunit;
        }

        result += ' فقط لا غير';
    }

    resultDiv.textContent = result;
}

function copyTafqitResult() {
    const result = document.getElementById('tafqitResult').textContent;
    if (result && result !== 'أدخل الرقم لتحويله إلى كتابة') {
        navigator.clipboard.writeText(result);
        showToast('تم نسخ النتيجة', 'success');
    }
}

function setupQuickTafqit() {
    document.getElementById('quickTafqitInput').addEventListener('input', function () {
        const num = parseFloat(this.value);
        const resultDiv = document.getElementById('quickTafqitResult');

        if (isNaN(num) || num < 0) {
            resultDiv.textContent = 'أدخل المبلغ لتحويله';
            return;
        }

        const intPart = Math.floor(num);
        let result = numberToArabicWords(intPart);

        if (intPart === 1) result += ' ريال سعودي';
        else if (intPart === 2) result = result.replace('اثنان', '') + 'ريالان سعوديان';
        else if (intPart >= 3 && intPart <= 10) result += ' ريالات سعودية';
        else result += ' ريال سعودي';

        result += ' فقط لا غير';
        resultDiv.textContent = result;
    });
}

function copyQuickTafqit() {
    const result = document.getElementById('quickTafqitResult').textContent;
    if (result !== 'أدخل المبلغ لتحويله') {
        navigator.clipboard.writeText(result);
        showToast('تم نسخ النتيجة', 'success');
    }
}

// ==================== INHERITANCE CALCULATOR ====================
function calculateInheritance() {
    // Get heir counts
    const heirs = {
        sons: parseInt(document.getElementById('sons').value) || 0,
        daughters: parseInt(document.getElementById('daughters').value) || 0,
        grandsonsSon: parseInt(document.getElementById('grandsonsSon').value) || 0,
        granddaughtersSon: parseInt(document.getElementById('granddaughtersSon').value) || 0,
        father: parseInt(document.getElementById('father').value) || 0,
        mother: parseInt(document.getElementById('mother').value) || 0,
        grandfather: parseInt(document.getElementById('grandfather').value) || 0,
        grandmotherMother: parseInt(document.getElementById('grandmotherMother').value) || 0,
        grandmotherFather: parseInt(document.getElementById('grandmotherFather').value) || 0,
        husband: parseInt(document.getElementById('husband').value) || 0,
        wives: parseInt(document.getElementById('wives').value) || 0,
        brothersFull: parseInt(document.getElementById('brothersFull').value) || 0,
        sistersFull: parseInt(document.getElementById('sistersFull').value) || 0,
        brothersFather: parseInt(document.getElementById('brothersFather').value) || 0,
        sistersFather: parseInt(document.getElementById('sistersFather').value) || 0,
        brothersMother: parseInt(document.getElementById('brothersMother').value) || 0,
        sistersMother: parseInt(document.getElementById('sistersMother').value) || 0,
        nephewsFull: parseInt(document.getElementById('nephewsFull').value) || 0,
        nephewsFather: parseInt(document.getElementById('nephewsFather').value) || 0,
        unclesFull: parseInt(document.getElementById('unclesFull').value) || 0,
        unclesFather: parseInt(document.getElementById('unclesFather').value) || 0,
        cousinsFull: parseInt(document.getElementById('cousinsFull').value) || 0,
        cousinsFather: parseInt(document.getElementById('cousinsFather').value) || 0
    };

    const estateValue = parseFloat(document.getElementById('estateValue').value) || 0;

    // Calculate shares
    const result = computeInheritance(heirs, estateValue);

    // Display results
    displayInheritanceResult(result, estateValue);

    // Switch to result tab
    switchTab(document.querySelectorAll('.tab')[1], 'resultTab');
}

function displayInheritanceResult(result, estateValue) {
    const container = document.getElementById('inheritanceResult');

    if (result.shares.length === 0) {
        container.innerHTML = `
            <div class="empty-state">
                <div class="empty-state-icon">⚠️</div>
                <p>لم يتم إدخال أي ورثة</p>
            </div>
        `;
        return;
    }

    let html = '<div class="inheritance-result">';

    // العول
    if (result.awl) {
        html += `
            <div class="evidence-box" style="background: #fef3c7; border-color: #f59e0b; color: #92400e; margin-bottom: 1rem;">
                <strong>⚠️ العول:</strong> ${result.awl.explanation}
            </div>
        `;
    }

    // الرد
    if (result.radd) {
        html += `
            <div class="evidence-box" style="background: #dbeafe; border-color: #3b82f6; color: #1e40af; margin-bottom: 1rem;">
                <strong>📊 الرد:</strong> ${result.radd.explanation}
            </div>
        `;
    }

    // حساب المبالغ التفصيلية
    const fardTotal = result.shares.filter(s => s.type === 'فرض').reduce((sum, s) => sum + s.shareValue, 0);
    const remainingForAsaba = estateValue > 0 ? estateValue * (1 - Math.min(fardTotal, 1)) : 0;

    // جدول الورثة
    html += `
        <div class="result-header">📋 توزيع التركة على الورثة</div>
        <table class="result-table">
            <thead>
                <tr>
                    <th>الوارث</th>
                    <th>النصيب</th>
                    <th>النوع</th>
                    ${estateValue > 0 ? '<th>المبلغ (ريال)</th>' : ''}
                </tr>
            </thead>
            <tbody>
    `;

    result.shares.forEach(share => {
        // حساب المبلغ الإجمالي للمجموعة
        let totalAmount = 0;
        if (estateValue > 0) {
            if (share.shareValue > 0) {
                let adjustedShare = share.shareValue;
                if (result.awl) {
                    adjustedShare = share.shareValue / result.awl.adjusted;
                }
                totalAmount = estateValue * adjustedShare;
            } else if (share.type.includes('تعصيب')) {
                totalAmount = remainingForAsaba;
            }
        }

        // تحديد عدد الذكور والإناث للتعصيب
        let maleCount = 0, femaleCount = 0;
        if (share.heir.includes('الأبناء') || share.heir.includes('أبناء')) {
            // للأبناء والبنات - للذكر مثل حظ الأنثيين
            if (share.heir.includes('البنات') || share.heir.includes('بنات')) {
                // استخراج العدد من البيانات الأصلية
                const heirsData = result.originalHeirs || {};
                maleCount = heirsData.sons || 0;
                femaleCount = heirsData.daughters || 0;
            } else {
                maleCount = share.count;
            }
        } else if (share.heir.includes('الإخوة') && share.heir.includes('الأخوات')) {
            const heirsData = result.originalHeirs || {};
            if (share.heir.includes('الأشقاء') || share.heir.includes('الشقيقات')) {
                maleCount = heirsData.brothersFull || 0;
                femaleCount = heirsData.sistersFull || 0;
            } else if (share.heir.includes('لأب')) {
                maleCount = heirsData.brothersFather || 0;
                femaleCount = heirsData.sistersFather || 0;
            }
        }

        // إذا كان هناك تعصيب بالغير (ذكور وإناث)
        if (share.type === 'تعصيب' && maleCount > 0 && femaleCount > 0) {
            const totalParts = maleCount * 2 + femaleCount;
            const partValue = totalAmount / totalParts;
            const maleShare = partValue * 2;
            const femaleShare = partValue;

            // عرض كل ابن في صف
            for (let i = 0; i < maleCount; i++) {
                html += `
                    <tr>
                        <td><strong>الابن ${maleCount > 1 ? (i + 1) : ''}</strong></td>
                        <td>للذكر مثل حظ الأنثيين</td>
                        <td>تعصيب</td>
                        ${estateValue > 0 ? `<td>${maleShare.toLocaleString('ar-SA', { maximumFractionDigits: 2 })}</td>` : ''}
                    </tr>
                `;
            }
            // عرض كل بنت في صف
            for (let i = 0; i < femaleCount; i++) {
                html += `
                    <tr>
                        <td><strong>البنت ${femaleCount > 1 ? (i + 1) : ''}</strong></td>
                        <td>للذكر مثل حظ الأنثيين</td>
                        <td>تعصيب</td>
                        ${estateValue > 0 ? `<td>${femaleShare.toLocaleString('ar-SA', { maximumFractionDigits: 2 })}</td>` : ''}
                    </tr>
                `;
            }
            // عرض الدليل مرة واحدة
            html += `
                <tr>
                    <td colspan="${estateValue > 0 ? 4 : 3}">
                        <div class="evidence-box">${share.evidence}</div>
                    </td>
                </tr>
            `;
        }
        // إذا كان هناك عدة ورثة يتقاسمون (مثل الزوجات أو البنات بالفرض)
        else if (share.count > 1 && estateValue > 0) {
            const individualAmount = totalAmount / share.count;
            const singularName = getSingularName(share.heir);

            for (let i = 0; i < share.count; i++) {
                html += `
                    <tr>
                        <td><strong>${singularName} ${share.count > 1 ? (i + 1) : ''}</strong></td>
                        <td>${share.share}</td>
                        <td>${share.type}</td>
                        <td>${individualAmount.toLocaleString('ar-SA', { maximumFractionDigits: 2 })}</td>
                    </tr>
                `;
            }
            html += `
                <tr>
                    <td colspan="4">
                        <div class="evidence-box">${share.evidence}</div>
                    </td>
                </tr>
            `;
        }
        // وارث واحد أو بدون قيمة تركة
        else {
            let amount = totalAmount > 0 ? totalAmount.toLocaleString('ar-SA', { maximumFractionDigits: 2 }) : '-';

            html += `
                <tr>
                    <td><strong>${share.heir}</strong></td>
                    <td>${share.share}</td>
                    <td>${share.type}</td>
                    ${estateValue > 0 ? `<td>${amount}</td>` : ''}
                </tr>
                <tr>
                    <td colspan="${estateValue > 0 ? 4 : 3}">
                        <div class="evidence-box">${share.evidence}</div>
                    </td>
                </tr>
            `;
        }
    });

    html += '</tbody></table>';

    // المحجوبون
    if (result.blocked.length > 0) {
        html += `
            <div class="result-header" style="margin-top: 1.5rem;">🚫 المحجوبون</div>
            <table class="result-table">
                <thead>
                    <tr>
                        <th>المحجوب</th>
                        <th>حاجبه</th>
                        <th>السبب</th>
                    </tr>
                </thead>
                <tbody>
        `;

        result.blocked.forEach(b => {
            html += `
                <tr>
                    <td>${b.heir}</td>
                    <td>${b.blockedBy}</td>
                    <td>${b.reason}</td>
                </tr>
            `;
        });

        html += '</tbody></table>';
    }

    html += '</div>';
    container.innerHTML = html;
}

// دالة مساعدة للحصول على الاسم المفرد
function getSingularName(pluralName) {
    const mapping = {
        'الزوجات': 'الزوجة',
        'البنات': 'البنت',
        'الأبناء': 'الابن',
        'الأخوات الشقيقات': 'الأخت الشقيقة',
        'الأخوات لأب': 'الأخت لأب',
        'الأخوات لأم': 'الأخت لأم',
        'الإخوة الأشقاء': 'الأخ الشقيق',
        'الإخوة لأب': 'الأخ لأب',
        'الإخوة لأم': 'الأخ لأم',
        'الجدات': 'الجدة',
        'بنات الابن': 'بنت الابن',
        'أبناء الابن': 'ابن الابن'
    };
    return mapping[pluralName] || pluralName;
}

// ==================== DIYAT CALCULATOR ====================
let diyatState = {
    gender: 'male',
    crime: 'Amd',
    unit: 'riyal',
    selectedItem: null
};

function setDiyatOption(btn) {
    const type = btn.dataset.type;
    const value = btn.dataset.value;

    // تحديث الأزرار النشطة في نفس المجموعة
    btn.parentElement.querySelectorAll('.option-btn').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');

    diyatState[type] = value;

    // تحديث النتيجة إذا كان هناك عنصر مختار
    if (diyatState.selectedItem) {
        showDiyatResult(diyatState.selectedItem);
    }
}

function loadDiyatList() {
    if (typeof diyatData === 'undefined') {
        document.getElementById('diyatList').innerHTML =
            '<div style="padding: 2rem; text-align: center; color: var(--text-light);">⚠️ لم يتم تحميل ملف البيانات (diyat-data.js)</div>';
        return;
    }

    const list = document.getElementById('diyatList');
    list.innerHTML = diyatData.map(item => `
        <button type="button" class="diyat-item" data-num="${item.num}" data-name="${escapeHtml(item.name)}" onclick="selectDiyatItem(${item.num})">
            <div style="display: flex; align-items: center; flex: 1;">
                <div class="diyat-item-num">${item.num}</div>
                <div class="diyat-item-name">${escapeHtml(item.name)}</div>
            </div>
            <span style="color: var(--text-light); font-size: 0.85rem;">←</span>
        </button>
    `).join('');
}

function filterDiyatList() {
    const query = document.getElementById('diyatSearch').value.trim();
    const items = document.querySelectorAll('.diyat-item');

    items.forEach(item => {
        const name = item.dataset.name || '';
        const num = item.dataset.num || '';
        const match = !query || name.includes(query) || num.includes(query);
        item.style.display = match ? 'flex' : 'none';
    });
}

function selectDiyatItem(num) {
    const item = diyatData.find(d => d.num === num);
    if (!item) return;

    diyatState.selectedItem = item;

    // إبراز العنصر المختار
    document.querySelectorAll('.diyat-item').forEach(i => i.classList.remove('selected'));
    document.querySelector(`.diyat-item[data-num="${num}"]`).classList.add('selected');

    showDiyatResult(item);
}

function showDiyatResult(item) {
    const key = diyatState.gender === 'male'
        ? (diyatState.crime === 'Amd' ? 'maleAmd' : 'maleKhata')
        : (diyatState.crime === 'Amd' ? 'femaleAmd' : 'femaleKhata');

    const fieldKey = key + '_' + diyatState.unit;
    const value = item[fieldKey];

    const genderLabel = diyatState.gender === 'male' ? 'الرجل' : 'المرأة';
    const crimeLabel = diyatState.crime === 'Amd' ? 'العمد وشبهه' : 'الخطأ';
    const unitLabel = diyatState.unit === 'riyal' ? 'ريال سعودي' : 'من الإبل';

    let displayValue = value;
    let tafqit = '';

    // التفقيط (للقيم الرقمية فقط بالريال)
    if (diyatState.unit === 'riyal' && /^\d+(\.\d+)?$/.test(value)) {
        const num = Math.floor(parseFloat(value));
        displayValue = num.toLocaleString('ar-SA') + ' ريال';
        let words = numberToArabicWords(num);
        if (num === 1) words += ' ريال سعودي';
        else if (num === 2) words = words.replace('اثنان', '') + 'ريالان سعوديان';
        else if (num >= 3 && num <= 10) words += ' ريالات سعودية';
        else words += ' ريال سعودي';
        words += ' فقط لا غير';
        tafqit = words;
    } else if (diyatState.unit === 'camel' && /^\d+(\.\d+)?$/.test(value)) {
        displayValue = value + ' من الإبل';
    } else if (value === 'حكومة') {
        displayValue = '⚖️ حكومة (تقدير القاضي)';
    } else if (value === 'مهر المثل') {
        displayValue = '💍 مهر المثل';
    } else if (value === 'قيمته') {
        displayValue = '💰 قيمته (السوقية)';
    } else if (/^[=\-]+$/.test(value)) {
        displayValue = '➖ لا تقدير له في هذه الحالة';
    } else {
        displayValue = value;
    }

    const resultDiv = document.getElementById('diyatResult');
    resultDiv.innerHTML = `
        <div class="diyat-result-card">
            <div class="label">
                <span class="diyat-info-tag">${genderLabel}</span>
                <span class="diyat-info-tag">${crimeLabel}</span>
                <span class="diyat-info-tag">${unitLabel}</span>
            </div>
            <div style="font-size: 1.05rem; opacity: 0.95; margin-bottom: 0.5rem;">
                ${escapeHtml(item.name)}
            </div>
            <div class="value">${escapeHtml(displayValue)}</div>
            ${tafqit ? `<div class="tafqit">${escapeHtml(tafqit)}</div>` : ''}
            <button type="button" class="btn" style="background: white; color: var(--primary); margin-top: 0.75rem;" onclick="copyDiyatResult()">
                📋 نسخ النتيجة
            </button>
        </div>
    `;
}

function copyDiyatResult() {
    if (!diyatState.selectedItem) return;

    const item = diyatState.selectedItem;
    const key = diyatState.gender === 'male'
        ? (diyatState.crime === 'Amd' ? 'maleAmd' : 'maleKhata')
        : (diyatState.crime === 'Amd' ? 'femaleAmd' : 'femaleKhata');
    const fieldKey = key + '_' + diyatState.unit;
    const value = item[fieldKey];

    const genderLabel = diyatState.gender === 'male' ? 'الرجل' : 'المرأة';
    const crimeLabel = diyatState.crime === 'Amd' ? 'العمد وشبهه' : 'الخطأ';
    const unitLabel = diyatState.unit === 'riyal' ? 'ريال سعودي' : 'من الإبل';

    const text = `حاسبة الديات والشجاج
العضو/المنفعة: ${item.name} (رقم ${item.num})
الجنس: ${genderLabel}
نوع الجناية: ${crimeLabel}
الدية: ${value} ${unitLabel}`;

    copyToClipboard(text);
}

// تحميل القائمة عند فتح المودال
document.addEventListener('DOMContentLoaded', function () {
    setTimeout(loadDiyatList, 500);
});

// ==================== EXPERT FEES CALCULATOR ====================
function calculateExpertFees() {
    const claim = parseFloat(document.getElementById('claimAmount').value) || 0;
    const awarded = parseFloat(document.getElementById('awardedAmount').value) || 0;
    const fees = parseFloat(document.getElementById('expertFees').value) || 0;

    const container = document.getElementById('expertResult');

    if (claim <= 0 || fees <= 0) {
        container.innerHTML = '<div class="result-box">أدخل البيانات لحساب التوزيع</div>';
        return;
    }

    if (awarded > claim) {
        container.innerHTML = '<div class="result-box" style="background: #fef3c7; border-color: #f59e0b;">المبلغ المحكوم به لا يمكن أن يتجاوز أصل المطالبة</div>';
        return;
    }

    // الحسابات
    const lossRatio = (claim - awarded) / claim;
    const plaintiffShare = fees * lossRatio;
    const defendantShare = fees - plaintiffShare;

    const formatNum = (n) => n.toLocaleString('ar-SA', { maximumFractionDigits: 2 });
    const formatPercent = (n) => (n * 100).toLocaleString('ar-SA', { maximumFractionDigits: 2 }) + '%';

    container.innerHTML = `
        <div class="result-header">📊 نتيجة التوزيع</div>
        <table class="result-table">
            <tr>
                <td><strong>نسبة ما خسره المدعي من دعواه</strong></td>
                <td>${formatPercent(lossRatio)}</td>
            </tr>
            <tr>
                <td><strong>ما يتحمله المدعي من أتعاب الخبرة</strong></td>
                <td>${formatNum(plaintiffShare)} ريال</td>
            </tr>
            <tr>
                <td><strong>ما يتحمله المدعى عليه من أتعاب الخبرة</strong></td>
                <td>${formatNum(defendantShare)} ريال</td>
            </tr>
        </table>
    `;
}

function copyExpertResult() {
    const claim = parseFloat(document.getElementById('claimAmount').value) || 0;
    const awarded = parseFloat(document.getElementById('awardedAmount').value) || 0;
    const fees = parseFloat(document.getElementById('expertFees').value) || 0;

    if (claim <= 0 || fees <= 0) {
        showToast('أدخل البيانات أولاً', '');
        return;
    }

    const lossRatio = (claim - awarded) / claim;
    const plaintiffShare = fees * lossRatio;
    const defendantShare = fees - plaintiffShare;

    const formatNum = (n) => n.toLocaleString('ar-SA', { maximumFractionDigits: 2 });
    const formatPercent = (n) => (n * 100).toLocaleString('ar-SA', { maximumFractionDigits: 2 }) + '%';

    const text = `حاسبة أتعاب الخبرة
أصل المطالبة: ${formatNum(claim)} ريال
المبلغ المحكوم به: ${formatNum(awarded)} ريال
أتعاب الخبير: ${formatNum(fees)} ريال

النتيجة:
- نسبة خسارة المدعي: ${formatPercent(lossRatio)}
- يتحمله المدعي: ${formatNum(plaintiffShare)} ريال
- يتحمله المدعى عليه: ${formatNum(defendantShare)} ريال`;

    copyToClipboard(text);
}

// ==================== DAMAGES CALCULATOR ====================
function calculateDamages() {
    const value = parseFloat(document.getElementById('damageValue').value) || 0;
    const ratio = parseFloat(document.getElementById('responsibilityRatio').value) || 0;
    const container = document.getElementById('damagesResult');

    if (value <= 0 || ratio < 0) {
        container.innerHTML = '<div class="result-box">أدخل البيانات لحساب التعويض</div>';
        return;
    }

    if (ratio > 100) {
        container.innerHTML = '<div class="result-box" style="background: #fef3c7; border-color: #f59e0b;">نسبة المسؤولية يجب ألا تتجاوز 100%</div>';
        return;
    }

    const compensation = value * (ratio / 100);
    const remaining = value - compensation;
    const formatNum = (n) => n.toLocaleString('ar-SA', { maximumFractionDigits: 2 });

    // التفقيط
    const intCompensation = Math.floor(compensation);
    let words = numberToArabicWords(intCompensation);
    if (intCompensation === 1) words += ' ريال سعودي';
    else if (intCompensation === 2) words = words.replace('اثنان', '') + 'ريالان سعوديان';
    else if (intCompensation >= 3 && intCompensation <= 10) words += ' ريالات سعودية';
    else words += ' ريال سعودي';
    words += ' فقط لا غير';

    container.innerHTML = `
        <div class="result-header">📊 نتيجة التعويض</div>
        <table class="result-table">
            <tr>
                <td><strong>قيمة الضرر الكلية</strong></td>
                <td>${formatNum(value)} ريال</td>
            </tr>
            <tr>
                <td><strong>نسبة المسؤولية</strong></td>
                <td>${ratio}%</td>
            </tr>
            <tr style="background: rgba(26, 95, 74, 0.1);">
                <td><strong>قيمة التعويض المستحق</strong></td>
                <td><strong>${formatNum(compensation)} ريال</strong></td>
            </tr>
            <tr>
                <td><strong>المتبقي على الطرف الآخر</strong></td>
                <td>${formatNum(remaining)} ريال</td>
            </tr>
        </table>
        <div class="evidence-box" style="margin-top: 1rem;">
            <strong>تفقيط التعويض:</strong> ${words}
        </div>
    `;
}

function copyDamagesResult() {
    const value = parseFloat(document.getElementById('damageValue').value) || 0;
    const ratio = parseFloat(document.getElementById('responsibilityRatio').value) || 0;

    if (value <= 0 || ratio < 0 || ratio > 100) {
        showToast('أدخل البيانات أولاً', '');
        return;
    }

    const compensation = value * (ratio / 100);
    const formatNum = (n) => n.toLocaleString('ar-SA', { maximumFractionDigits: 2 });

    const text = `حاسبة تقدير الأضرار
قيمة الضرر: ${formatNum(value)} ريال
نسبة المسؤولية: ${ratio}%
قيمة التعويض المستحق: ${formatNum(compensation)} ريال`;

    copyToClipboard(text);
}

// ==================== TRAFFIC DAMAGES CALCULATOR ====================
function calculateTraffic() {
    const damages = parseFloat(document.getElementById('trafficDamages').value) || 0;
    const fees = parseFloat(document.getElementById('trafficFees').value) || 0;
    const shipping = parseFloat(document.getElementById('trafficShipping').value) || 0;
    const other = parseFloat(document.getElementById('trafficOther').value) || 0;
    const container = document.getElementById('trafficResult');

    const total = damages + fees + shipping + other;

    if (total <= 0) {
        container.innerHTML = '<div class="result-box">أدخل البيانات لحساب المجموع</div>';
        return;
    }

    const formatNum = (n) => n.toLocaleString('ar-SA', { maximumFractionDigits: 2 });

    // التفقيط
    const intTotal = Math.floor(total);
    let words = numberToArabicWords(intTotal);
    if (intTotal === 1) words += ' ريال سعودي';
    else if (intTotal === 2) words = words.replace('اثنان', '') + 'ريالان سعوديان';
    else if (intTotal >= 3 && intTotal <= 10) words += ' ريالات سعودية';
    else words += ' ريال سعودي';
    words += ' فقط لا غير';

    container.innerHTML = `
        <div class="result-header">📊 إجمالي الأضرار المرورية</div>
        <table class="result-table">
            <tr>
                <td><strong>تقدير التلفيات</strong></td>
                <td>${formatNum(damages)} ريال</td>
            </tr>
            <tr>
                <td><strong>رسوم التقدير</strong></td>
                <td>${formatNum(fees)} ريال</td>
            </tr>
            <tr>
                <td><strong>رسوم شحن المركبة</strong></td>
                <td>${formatNum(shipping)} ريال</td>
            </tr>
            <tr>
                <td><strong>تلفيات أو أضرار أخرى</strong></td>
                <td>${formatNum(other)} ريال</td>
            </tr>
            <tr style="background: rgba(26, 95, 74, 0.1);">
                <td><strong>المجموع الكلي</strong></td>
                <td><strong>${formatNum(total)} ريال</strong></td>
            </tr>
        </table>
        <div class="evidence-box" style="margin-top: 1rem;">
            <strong>تفقيط المجموع:</strong> ${words}
        </div>
    `;
}

function copyTrafficResult() {
    const damages = parseFloat(document.getElementById('trafficDamages').value) || 0;
    const fees = parseFloat(document.getElementById('trafficFees').value) || 0;
    const shipping = parseFloat(document.getElementById('trafficShipping').value) || 0;
    const other = parseFloat(document.getElementById('trafficOther').value) || 0;
    const total = damages + fees + shipping + other;

    if (total <= 0) {
        showToast('أدخل البيانات أولاً', '');
        return;
    }

    const formatNum = (n) => n.toLocaleString('ar-SA', { maximumFractionDigits: 2 });

    const text = `حاسبة الأضرار المرورية
تقدير التلفيات: ${formatNum(damages)} ريال
رسوم التقدير: ${formatNum(fees)} ريال
رسوم شحن المركبة: ${formatNum(shipping)} ريال
تلفيات أخرى: ${formatNum(other)} ريال
المجموع الكلي: ${formatNum(total)} ريال`;

    copyToClipboard(text);
}

// ==================== END OF SERVICE BENEFIT CALCULATOR ====================
function setEosDurationMode(btn) {
    btn.parentElement.querySelectorAll('.option-btn').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    const isDates = btn.dataset.value === 'dates';
    document.getElementById('eosDatesInputs').style.display = isDates ? '' : 'none';
    document.getElementById('eosManualInputs').style.display = isDates ? 'none' : '';
    calculateEOS();
}

function getEosServiceDuration() {
    const datesMode = document.getElementById('eosDatesInputs').style.display !== 'none';
    if (datesMode) {
        const startVal = document.getElementById('eosStartDate').value;
        const endVal = document.getElementById('eosEndDate').value;
        if (!startVal || !endVal) return null;
        return dateDiffYMD(startVal, endVal);
    }
    const y = parseInt(document.getElementById('eosYears').value) || 0;
    const m = parseInt(document.getElementById('eosMonths').value) || 0;
    const d = parseInt(document.getElementById('eosDays').value) || 0;
    if (y === 0 && m === 0 && d === 0) return null;
    return { years: y + Math.floor(m / 12), months: m % 12, days: d };
}

function computeEOS() {
    const wage = parseFloat(document.getElementById('eosWage').value) || 0;
    const reason = document.getElementById('eosReason').value;
    const duration = getEosServiceDuration();

    const result = computeEOSCore(wage, reason, duration);
    if (!result || result.error) return result;

    return {
        ...result,
        reasonText: document.getElementById('eosReason').selectedOptions[0].textContent
    };
}

function formatEosDuration(d) {
    const parts = [];
    if (d.years > 0) parts.push(d.years + (d.years === 1 ? ' سنة' : d.years === 2 ? ' سنتان' : d.years <= 10 ? ' سنوات' : ' سنة'));
    if (d.months > 0) parts.push(d.months + (d.months === 1 ? ' شهر' : d.months === 2 ? ' شهران' : ' أشهر'));
    if (d.days > 0) parts.push(d.days + (d.days === 1 ? ' يوم' : d.days === 2 ? ' يومان' : d.days <= 10 ? ' أيام' : ' يوماً'));
    return parts.length ? parts.join(' و ') : 'صفر';
}

function calculateEOS() {
    const container = document.getElementById('eosResult');
    const result = computeEOS();

    if (!result) {
        container.innerHTML = '<div class="result-box">أدخل البيانات لحساب المكافأة</div>';
        return;
    }

    if (result.error) {
        container.innerHTML = `<div class="result-box" style="background: #fef3c7; border-color: #f59e0b;">${result.error}</div>`;
        return;
    }

    const formatNum = (n) => n.toLocaleString('ar-SA', { maximumFractionDigits: 2 });

    // التفقيط
    const intDue = Math.floor(result.dueAward);
    let words = numberToArabicWords(intDue);
    if (intDue === 1) words += ' ريال سعودي';
    else if (intDue === 2) words = words.replace('اثنان', '') + 'ريالان سعوديان';
    else if (intDue >= 3 && intDue <= 10) words += ' ريالات سعودية';
    else words += ' ريال سعودي';
    words += ' فقط لا غير';

    container.innerHTML = `
        <div class="result-header">📊 نتيجة الاحتساب</div>
        <table class="result-table">
            <tr>
                <td><strong>مدة الخدمة</strong></td>
                <td>${formatEosDuration(result.duration)}</td>
            </tr>
            <tr>
                <td><strong>الأجر الشهري الأخير</strong></td>
                <td>${formatNum(result.wage)} ريال</td>
            </tr>
            <tr>
                <td><strong>مكافأة السنوات الخمس الأولى (نصف شهر عن كل سنة)</strong></td>
                <td>${formatNum(result.firstFive)} ريال</td>
            </tr>
            <tr>
                <td><strong>مكافأة السنوات التالية (شهر عن كل سنة)</strong></td>
                <td>${formatNum(result.afterFive)} ريال</td>
            </tr>
            <tr>
                <td><strong>إجمالي المكافأة الكاملة</strong></td>
                <td>${formatNum(result.fullAward)} ريال</td>
            </tr>
            <tr>
                <td><strong>النسبة المستحقة</strong></td>
                <td>${result.ratioLabel}</td>
            </tr>
            <tr style="background: rgba(26, 95, 74, 0.1);">
                <td><strong>المكافأة المستحقة</strong></td>
                <td><strong>${formatNum(result.dueAward)} ريال</strong></td>
            </tr>
        </table>
        ${result.dueAward > 0 ? `
        <div class="evidence-box" style="margin-top: 1rem;">
            <strong>تفقيط المكافأة:</strong> ${words}
        </div>` : ''}
        <div class="result-box" style="margin-top: 1rem;">
            ⚖️ ${result.ruleNote}
        </div>
    `;
}

function copyEOSResult() {
    const result = computeEOS();

    if (!result || result.error) {
        showToast('أدخل البيانات أولاً', '');
        return;
    }

    const formatNum = (n) => n.toLocaleString('ar-SA', { maximumFractionDigits: 2 });

    const text = `حاسبة مكافأة نهاية الخدمة
سبب انتهاء العلاقة العمالية: ${result.reasonText}
مدة الخدمة: ${formatEosDuration(result.duration)}
الأجر الشهري الأخير: ${formatNum(result.wage)} ريال

النتيجة:
- مكافأة السنوات الخمس الأولى: ${formatNum(result.firstFive)} ريال
- مكافأة السنوات التالية: ${formatNum(result.afterFive)} ريال
- إجمالي المكافأة الكاملة: ${formatNum(result.fullAward)} ريال
- النسبة المستحقة: ${result.ratioLabel}
- المكافأة المستحقة: ${formatNum(result.dueAward)} ريال

${result.ruleNote}
(وفق المواد 84، 85، 87 من نظام العمل السعودي)`;

    copyToClipboard(text);
}
