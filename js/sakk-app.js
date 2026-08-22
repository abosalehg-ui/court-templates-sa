// ==================== واجهة معالج الصكوك ====================
// كود الواجهة والتفاعل لصفحة sakk.html
// منطق المعالجة الصافي في js/sakk.js (مغطى بالاختبارات)

// ==================== أدوات مساعدة ====================
function $(id) { return document.getElementById(id); }

function escapeHtml(str) {
    return String(str || '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function showToast(msg, type = '') {
    const toast = $('toast');
    toast.textContent = msg;
    toast.className = 'toast show ' + type;
    setTimeout(() => { toast.className = 'toast'; }, 2500);
}

// تهريب المحارف الخاصة بالتعابير النمطية داخل كلمات القاموس
function escapeRegExp(str) {
    return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

const arabicNum = (n) => Number(n).toLocaleString('ar-SA');

// ==================== مفاتيح التخزين المحلي ====================
const DRAFT_KEY = 'courtSakkDraft';
const THEME_KEY = 'courtTheme'; // نفس مفتاح الصفحة الرئيسية ليتزامن الوضع الليلي

// قراءة آمنة من localStorage — بيئة تمنع التخزين لا تعطّل الصفحة
function safeGet(key) {
    try { return localStorage.getItem(key); } catch (e) { return null; }
}
function safeSet(key, value) {
    try { localStorage.setItem(key, value); } catch (e) { /* التخزين غير متاح */ }
}
function safeRemove(key) {
    try { localStorage.removeItem(key); } catch (e) { /* التخزين غير متاح */ }
}

// ==================== الحالة ====================
// نتيجة آخر معالجة — تستهلكها أزرار النسخ والحفظ والطباعة
let lastResult = null;

// ==================== الوضع الليلي (متزامن مع الصفحة الرئيسية) ====================
function applyTheme(dark) {
    document.body.setAttribute('data-theme', dark ? 'dark' : '');
    $('darkModeToggle').textContent = dark ? '☀️' : '🌙';
}
applyTheme(safeGet(THEME_KEY) === 'dark');
$('darkModeToggle').addEventListener('click', () => {
    const nowDark = document.body.getAttribute('data-theme') !== 'dark';
    applyTheme(nowDark);
    safeSet(THEME_KEY, nowDark ? 'dark' : '');
});

// ==================== عدّاد الأحرف وحفظ المسودة ====================
let draftTimer;

function updateInputCount() {
    $('inputCharCount').textContent = 'عدد الأحرف: ' + arabicNum(countChars($('inputText').value));
}

$('inputText').addEventListener('input', function () {
    updateInputCount();
    clearTimeout(draftTimer);
    draftTimer = setTimeout(() => {
        const text = $('inputText').value;
        if (text.trim()) safeSet(DRAFT_KEY, text);
        else safeRemove(DRAFT_KEY);
    }, 500);
});

// استرجاع مسودة سابقة — تُعرض كخيار لا كاستبدال تلقائي
(function offerDraft() {
    const draft = safeGet(DRAFT_KEY);
    if (!draft || !draft.trim()) return;
    $('draftBanner').style.display = '';
    $('restoreDraftBtn').addEventListener('click', () => {
        $('inputText').value = draft;
        updateInputCount();
        $('draftBanner').style.display = 'none';
        showToast('تم استرجاع النص المحفوظ ✓', 'success');
    });
    $('discardDraftBtn').addEventListener('click', () => {
        safeRemove(DRAFT_KEY);
        $('draftBanner').style.display = 'none';
    });
})();

// ==================== فتح ملف نصي ====================
$('openFileBtn').addEventListener('click', () => $('fileInput').click());

$('fileInput').addEventListener('change', function (e) {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
        $('inputText').value = ev.target.result;
        updateInputCount();
        safeSet(DRAFT_KEY, ev.target.result);
        $('draftBanner').style.display = 'none';
        showToast('تم تحميل الملف ✓', 'success');
    };
    reader.onerror = () => showToast('تعذر قراءة الملف');
    reader.readAsText(file, 'utf-8');
    // تصفير القيمة حتى يعمل اختيار الملف نفسه مرة أخرى
    this.value = '';
});

// ==================== بناء عرض المقارنة ====================
// الترتيب مقصود: تهريب HTML أولاً، ثم تظليل الكلمات المصححة، ثم تحويل علامات التتبع
// إلى وسوم. أي تبديل في هذا الترتيب يفتح ثغرة حقن أو يكسر التظليل.
function buildComparison(inputText, result) {
    let originalHighlighted = escapeHtml(inputText);
    let processedHighlighted = escapeHtml(result.annotated);

    result.correctedWords.forEach(({ wrong, correct }) => {
        const wrongRegex = new RegExp('(^|[^\\u0600-\\u06FF])' + escapeRegExp(wrong) + '(?![\\u0600-\\u06FF])', 'g');
        const correctRegex = new RegExp('(^|[^\\u0600-\\u06FF])' + escapeRegExp(correct) + '(?![\\u0600-\\u06FF])', 'g');
        originalHighlighted = originalHighlighted.replace(wrongRegex, '$1<span class="corrected-word">' + wrong + '</span>');
        processedHighlighted = processedHighlighted.replace(correctRegex, '$1<span class="corrected-word">' + correct + '</span>');
    });

    processedHighlighted = processedHighlighted
        .replace(new RegExp(SAKK_MARKS.MERGE_START, 'g'), '<span class="merged-session" title="موضع دمج جلسات">')
        .replace(new RegExp(SAKK_MARKS.MERGE_END, 'g'), '</span>')
        .replace(new RegExp(SAKK_MARKS.DELETE_MARK, 'g'), '<span class="deleted-session" title="حُذفت هنا جلسة شطب">⟦ حُذفت جلسة شطب ⟧</span>');

    $('originalText').innerHTML = originalHighlighted;
    $('processedText').innerHTML = processedHighlighted;
    $('originalCharCount').textContent = 'عدد الأحرف: ' + arabicNum(countChars(inputText));
    $('processedCharCount').textContent = 'عدد الأحرف: ' + arabicNum(countChars(result.text));
}

function renderStats(stats) {
    $('sessionsCount').textContent = arabicNum(stats.sessionsMerged);
    $('deletedCount').textContent = arabicNum(stats.shatbDeleted);
    $('correctionsCount').textContent = arabicNum(stats.corrections);
    $('spacesCount').textContent = arabicNum(stats.spaces);
    $('colonsCount').textContent = arabicNum(stats.colons);
    $('emptyLinesCount').textContent = arabicNum(stats.emptyLines);
    $('statsCard').style.display = '';
}

// ==================== المعالجة ====================
function runProcessing() {
    const inputText = $('inputText').value;

    if (!inputText.trim()) {
        showToast('⚠️ الرجاء إدخال نص للمعالجة');
        return;
    }

    const btn = $('processBtn');
    const originalLabel = btn.textContent;
    btn.disabled = true;
    btn.textContent = '⏳ جارٍ المعالجة...';

    // تأخير قصير حتى ترسم الواجهة حالة الانتظار قبل المعالجة المتزامنة
    setTimeout(() => {
        try {
            const result = processSakk(inputText, { enableCorrection: $('enableCorrection').checked });
            lastResult = result;

            renderStats(result.stats);
            buildComparison(inputText, result);
            $('comparison').style.display = '';

            $('copyBtn').disabled = false;
            $('downloadBtn').disabled = false;
            $('printBtn').disabled = false;

            if (result.warnings.length) {
                $('skippedWarning').textContent = '⚠️ ' + result.warnings.join('؛ ');
                $('skippedWarning').style.display = '';
                showToast('تمت المعالجة مع تنبيهات');
            } else {
                $('skippedWarning').style.display = 'none';
                showToast('تمت المعالجة بنجاح ✓', 'success');
            }
        } catch (err) {
            console.error('خطأ أثناء معالجة نص الضبط:', err);
            showToast('حدث خطأ غير متوقع أثناء المعالجة');
        } finally {
            btn.disabled = false;
            btn.textContent = originalLabel;
        }
    }, 50);
}

$('processBtn').addEventListener('click', runProcessing);

// Ctrl+Enter من داخل حقل الإدخال يشغّل المعالجة
$('inputText').addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
        e.preventDefault();
        runProcessing();
    }
});

// ==================== محو النص ====================
$('clearBtn').addEventListener('click', () => {
    $('inputText').value = '';
    updateInputCount();
    safeRemove(DRAFT_KEY);
    lastResult = null;

    $('comparison').style.display = 'none';
    $('statsCard').style.display = 'none';
    $('skippedWarning').style.display = 'none';
    $('draftBanner').style.display = 'none';
    $('originalText').innerHTML = '';
    $('processedText').innerHTML = '';

    $('copyBtn').disabled = true;
    $('downloadBtn').disabled = true;
    $('printBtn').disabled = true;

    $('inputText').focus();
    showToast('تم محو النص');
});

// ==================== النسخ ====================
function fallbackCopy(text) {
    const ta = document.createElement('textarea');
    ta.value = text;
    ta.style.position = 'fixed';
    ta.style.opacity = '0';
    document.body.appendChild(ta);
    ta.select();
    try { document.execCommand('copy'); showToast('تم نسخ النتيجة ✓', 'success'); } catch (e) { showToast('تعذر النسخ'); }
    document.body.removeChild(ta);
}

$('copyBtn').addEventListener('click', () => {
    if (!lastResult) return;
    const text = lastResult.text;
    if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(text)
            .then(() => showToast('تم نسخ النتيجة ✓', 'success'))
            .catch(() => fallbackCopy(text));
    } else {
        fallbackCopy(text);
    }
});

// ==================== الحفظ كملف نصي ====================
$('downloadBtn').addEventListener('click', () => {
    if (!lastResult) return;
    const today = new Date();
    const stamp = today.getFullYear() + '-' +
        String(today.getMonth() + 1).padStart(2, '0') + '-' +
        String(today.getDate()).padStart(2, '0');

    const blob = new Blob([lastResult.text], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'صك_معالج_' + stamp + '.txt';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    // يُؤجَّل التحرير: إبطال الرابط فور النقر يُسقط اسم الملف المطلوب في بعض المتصفحات
    setTimeout(() => URL.revokeObjectURL(url), 1000);
    showToast('تم حفظ الملف ✓', 'success');
});

// ==================== الطباعة ====================
$('printBtn').addEventListener('click', () => window.print());

// ==================== التهيئة ====================
updateInputCount();
