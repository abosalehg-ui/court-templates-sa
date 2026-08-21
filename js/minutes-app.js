// ==================== واجهة مولّد الجلسة ====================
// كود الواجهة والتفاعل لصفحة minutes.html
// منطق الصياغة الصافي في js/minutes.js (مغطى بالاختبارات)

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

// ==================== مفاتيح التخزين المحلي ====================
const DRAFT_KEY = 'courtMinutesDraft';
const SETTINGS_KEY = 'courtMinutesSettings';
const THEME_KEY = 'courtTheme'; // نفس مفتاح الصفحة الرئيسية ليتزامن الوضع الليلي

// ==================== الحالة ====================
// state: حالة المحضر (تُمرَّر لمحرك الصياغة) — uiState: حالة الواجهة فقط
const state = freshMinutesState();
const uiState = {
    manualText: null,
    manualEditActive: false,
    currentStep: 0,
    clerkNoticeShown: false
};

// ==================== الوضع الليلي (متزامن مع الصفحة الرئيسية) ====================
function applyTheme(dark) {
    document.body.setAttribute('data-theme', dark ? 'dark' : '');
    $('darkModeToggle').textContent = dark ? '☀️' : '🌙';
}
applyTheme(localStorage.getItem(THEME_KEY) === 'dark');
$('darkModeToggle').addEventListener('click', () => {
    const nowDark = document.body.getAttribute('data-theme') !== 'dark';
    applyTheme(nowDark);
    localStorage.setItem(THEME_KEY, nowDark ? 'dark' : '');
});

// ==================== التاريخ الهجري في الترويسة ====================
(function setHijriDate() {
    try {
        const today = new Date();
        const weekday = new Intl.DateTimeFormat('ar-SA', { weekday: 'long' }).format(today);
        const hijri = new Intl.DateTimeFormat('ar-SA-u-ca-islamic-umalqura-nu-latn', { day: '2-digit', month: '2-digit', year: 'numeric' });
        const parts = hijri.formatToParts(today);
        const get = t => parts.find(p => p.type === t).value;
        $('hijriDate').textContent = `يوم ${weekday} الموافق ${get('day')} / ${get('month')} / ${get('year')} هـ`;
    } catch (e) {
        $('hijriDate').textContent = '';
    }
})();

// ==================== النوافذ المنبثقة ====================
function openMinutesModal(id) {
    $(id).classList.add('show');
}
function closeMinutesModal(id) {
    $(id).classList.remove('show');
}
[['confirmModal', ['confirmModalClose', 'confirmModalNo']],
 ['wakalaModal', ['wakalaModalClose', 'wakalaModalOk']],
 ['clerkModal', ['clerkModalClose', 'clerkModalOk']]].forEach(([modalId, btns]) => {
    btns.forEach(btnId => $(btnId).addEventListener('click', () => closeMinutesModal(modalId)));
    $(modalId).addEventListener('click', e => { if (e.target === $(modalId)) closeMinutesModal(modalId); });
});

// ==================== بيانات القاضي والمحكمة (محفوظة محليًا) ====================
(function initSettings() {
    let saved = {};
    try { saved = JSON.parse(localStorage.getItem(SETTINGS_KEY) || '{}'); } catch (e) { }
    $('judgeName').value = saved.judge || '';
    $('courtName').value = saved.court || '';
    state.opening.judge = $('judgeName').value;
    state.opening.court = $('courtName').value;
})();
function saveSettings() {
    try {
        localStorage.setItem(SETTINGS_KEY, JSON.stringify({ judge: state.opening.judge, court: state.opening.court }));
    } catch (e) { }
}
$('judgeName').addEventListener('input', e => { state.opening.judge = e.target.value; saveSettings(); render(); });
$('courtName').addEventListener('input', e => { state.opening.court = e.target.value; saveSettings(); render(); });

// ==================== إدراج بيانات الطرفين في نص الضبط ====================
// مُغلق افتراضًا لأن جدول نظام تقاضي أعلى صفحة الضبط والصك يحمل اسم الطرف ورقم هويته،
// فيُخفى حقلا الاسم والهوية ويُوصف الطرف بلقبه وحده. ويبقى تحديد النوع (فرد/شركة،
// مذكر/مؤنث) لأن عليه تدور المطابقة النحوية. وبيانات الوكيل لا يشملها المفتاح.
$('includePartyDataInText').addEventListener('change', e => {
    state.includePartyDataInText = e.target.checked;
    updateVisibility('plaintiff');
    updateVisibility('defendant');
    renderExtraCards('plaintiff');
    renderExtraCards('defendant');
    render();
});

// ==================== طريقة انعقاد الجلسة ====================
const SESSION_MODE_HINTS = {
    videoFull: 'تُدرج صيغة الانعقاد كاملة مع الاستناد لقرار رئيس المجلس الأعلى للقضاء رقم (17388) وتاريخ 5/10/1441هـ.',
    video: 'تُدرج عبارة «عبر الاتصال المرئي» وحدها في الافتتاح دون نص القرار، وتبقى بقية فقرات المحضر كما هي.',
    inPerson: 'يُفتتح المحضر دون ذكر الاتصال المرئي، وتُحذف الإشارة لرابط غرفة الاتصال المرئي من فقرات الغياب والشطب وتُستبدل بصيغة الحضور المباشر.'
};

function updateSessionModeUI() {
    const mode = state.opening.mode;
    $('sessionModeHint').textContent = SESSION_MODE_HINTS[mode] || '';
    const noVideoBtn = document.querySelector('.choice-group[data-group="plaintiff-specialCase"] .choice-btn[data-value="systemNoVideo"]');
    if (noVideoBtn) {
        noVideoBtn.textContent = mode === 'inPerson'
            ? '⚠ تبيّن حضوره بالنظام الإلكتروني، لكنه لم يحضر بمقر المحكمة'
            : '⚠ تبيّن حضوره بالنظام الإلكتروني، لكنه لم يحضر عبر الاتصال المرئي';
    }
}
updateSessionModeUI();

// ==================== وقت إغلاق الجلسة ====================
// وقت الافتتاح مثبت في ناجز أعلى صفحة القضية، فلا يُطلب هنا ولا يُكرَّر في متن الضبط،
// وبطاقة وقت الإغلاق تُذيَّل بها خطوة (الأسباب والحكم) — انظر showStep
function populateHourOptions() {
    const hourSel = $('closeHour');
    const prevValue = hourSel.value;
    hourSel.innerHTML = '';
    const hours = validHoursForPeriod($('closePeriod').value || 'ص');
    hours.forEach(h => {
        const opt = document.createElement('option');
        opt.value = h;
        opt.textContent = String(h).padStart(2, '0');
        hourSel.appendChild(opt);
    });
    hourSel.value = hours.map(String).includes(prevValue) ? prevValue : hours[0];
}

(function populateTimeSelects() {
    const minSel = $('closeMinute');
    for (let m = 0; m < 60; m += 5) {
        const opt = document.createElement('option');
        opt.value = m;
        opt.textContent = String(m).padStart(2, '0');
        minSel.appendChild(opt);
    }
    minSel.value = state.closing.minute;
    populateHourOptions();
    $('closeHour').value = state.closing.hour;
})();

function syncClosingTime() {
    state.closing.hour = Number($('closeHour').value);
    state.closing.minute = Number($('closeMinute').value);
    state.closing.period = $('closePeriod').value;
    $('closePeriodHint').textContent = `يُختم الضبط بـ: وأغلقت الجلسة الساعة ${buildTimeArabic(state.closing.hour, state.closing.minute, state.closing.period)}`;
}
$('closePeriod').addEventListener('change', () => { populateHourOptions(); syncClosingTime(); render(); });
['closeHour', 'closeMinute'].forEach(id => {
    $(id).addEventListener('change', () => { syncClosingTime(); render(); });
});
syncClosingTime();

// ==================== تحويل الأرقام المشرقية وحقول الأرقام فقط ====================
document.addEventListener('input', function (e) {
    const el = e.target;
    if (el.tagName !== 'INPUT' && el.tagName !== 'TEXTAREA') return;
    let val = convertArabicDigits(el.value);
    if (el.hasAttribute('data-numeric-only')) {
        val = val.replace(/[^0-9]/g, '');
    }
    if (val !== el.value) {
        const pos = el.selectionStart;
        el.value = val;
        try { el.selectionStart = el.selectionEnd = pos; } catch (err) { }
    }
}, true);

// ==================== قوائم صلة القرابة ====================
function generateNasabOptionsHTML(role) {
    const roleLabel = role === 'plaintiff' ? 'بالمدعي' : 'بالمدعى عليه';
    const byDegree = { 1: [], 2: [], 3: [], 4: [] };
    Object.keys(kinshipDegree).forEach(k => byDegree[kinshipDegree[k]].push(k));
    const groups = [1, 2, 3, 4].map(d =>
        `<optgroup label="الدرجة ${degreeOrdinal[d]}">${byDegree[d].map(k => `<option value="${k}">${k}</option>`).join('')}</optgroup>`
    ).join('');
    return `<option value="">اختر صلة القرابة ${roleLabel}...</option>${groups}<option value="أخرى">ليس قريبًا من درجات القرابة الأربعة</option>`;
}

function generateAsharOptionsHTML(direction) {
    const groups = asharTemplate.map(group => {
        const opts = group.items.map(itemTpl => {
            const text = itemTpl.replace('{S}', direction);
            return `<option value="${text}">${text}</option>`;
        }).join('');
        return `<optgroup label="الدرجة ${degreeOrdinal[group.degree]}">${opts}</optgroup>`;
    }).join('');
    return `<option value="">اختر صلة القرابة...</option>${groups}<option value="أخرى">ليس قريبًا من درجات القرابة الأربعة</option>`;
}

function populateAsharSelect(party) {
    $(`${party}-kinship-ashar`).innerHTML = generateAsharOptionsHTML(state[party].asharDirection);
}

function updateKinshipHint(party) {
    const s = state[party];
    const hintEl = $(`${party}-kinship-hint`);
    const isAshar = s.kinshipType === 'أصهار';
    const value = isAshar ? s.kinshipAshar : s.kinship;
    const degreeTable = isAshar ? asharDegree : kinshipDegree;

    if (!value) { hintEl.innerHTML = ''; return; }
    if (value === 'أخرى' || !degreeTable[value]) {
        hintEl.innerHTML = '<div class="kinship-degree-hint bad">⚠️ خارج الدرجات الأربع — سيُضاف تلقائيًا نص يفيد عدم أحقية الوكيل بالاستمرار في الإجابة، استنادًا للقرار الوزاري رقم (2044) وتاريخ 1443/8/4هـ.</div>';
    } else {
        const deg = degreeTable[value];
        const typeLabel = isAshar ? ' (أصهار)' : '';
        hintEl.innerHTML = `<div class="kinship-degree-hint ok">✓ صلة قرابة من الدرجة ${degreeOrdinal[deg]}${typeLabel} — مقبولة لوكيل غير محامٍ.</div>`;
    }
}

// ==================== بحث الجنسيات ====================
// يُستعمل للطرف الأول ولكل طرف إضافي، فبطاقة الهوية واحدة للجميع
function attachNationalitySearch(input, list, onValue) {
    if (!input || !list) return;
    let highlightedIdx = -1;
    let filtered = [];

    function renderList() {
        // تطبيع الهمزات والتشكيل حتى يجد «عماني» جنسية «عُماني»
        const q = normalizeArabicSearch(input.value);
        filtered = q ? NATIONALITY_OPTIONS.filter(n => normalizeArabicSearch(n).includes(q)) : NATIONALITY_OPTIONS;
        list.innerHTML = filtered.length === 0
            ? '<div class="nat-dropdown-empty">لا توجد نتائج مطابقة</div>'
            : filtered.map((n, i) => `<div class="nat-dropdown-item" data-idx="${i}">${n}</div>`).join('');
        highlightedIdx = -1;
        list.style.display = 'block';
    }

    function choose(value) {
        input.value = value;
        list.style.display = 'none';
        onValue(value);
    }

    input.addEventListener('focus', renderList);
    input.addEventListener('input', () => {
        renderList();
        onValue(input.value);
    });
    input.addEventListener('blur', () => {
        // مهلة تسمح بالنقر على عنصر القائمة، ولا تُخفيها إن عاد التركيز للحقل سريعًا
        setTimeout(() => { if (document.activeElement !== input) list.style.display = 'none'; }, 150);
    });
    input.addEventListener('keydown', e => {
        const items = list.querySelectorAll('.nat-dropdown-item');
        if (e.key === 'ArrowDown') {
            e.preventDefault();
            highlightedIdx = Math.min(highlightedIdx + 1, items.length - 1);
        } else if (e.key === 'ArrowUp') {
            e.preventDefault();
            highlightedIdx = Math.max(highlightedIdx - 1, 0);
        } else if (e.key === 'Enter') {
            if (highlightedIdx >= 0 && filtered[highlightedIdx]) {
                e.preventDefault();
                choose(filtered[highlightedIdx]);
            }
            return;
        } else if (e.key === 'Escape') {
            list.style.display = 'none';
            return;
        } else {
            return;
        }
        items.forEach((el, i) => el.classList.toggle('highlighted', i === highlightedIdx));
    });
    list.addEventListener('mousedown', e => {
        const item = e.target.closest('.nat-dropdown-item');
        if (item && item.dataset.idx !== undefined) choose(filtered[Number(item.dataset.idx)]);
    });
}

function setupNationalitySearch(party) {
    attachNationalitySearch($(`${party}-foreignNationality`), $(`${party}-foreignNationality-list`), value => {
        state[party].foreignNationality = value;
        render();
    });
}

// ==================== إظهار/إخفاء الحقول حسب الاختيارات ====================
function updateGenderedLabels(party) {
    const gender = state[party].gender;
    const attGroup = document.querySelector(`.choice-group[data-group="${party}-attendance"]`);
    if (attGroup) {
        const asalaBtn = attGroup.querySelector('[data-value="أصالة"]');
        const absentBtn = attGroup.querySelector('[data-value="لم يحضر"]');
        if (asalaBtn && !isCorporateEntity(state[party])) asalaBtn.textContent = gender === 'م' ? 'حضر أصالة' : 'حضرت أصالة';
        if (absentBtn) absentBtn.textContent = gender === 'م' ? 'لم يحضر' : 'لم تحضر';
    }
    const notifyGroup = document.querySelector(`.choice-group[data-group="${party}-notifyStatus"]`);
    if (notifyGroup) {
        const yesBtn = notifyGroup.querySelector('[data-value="تبلغ"]');
        const noBtn = notifyGroup.querySelector('[data-value="لم يتبلغ"]');
        if (yesBtn) yesBtn.textContent = gender === 'م' ? 'تبلّغ' : 'تبلّغت';
        if (noBtn) noBtn.textContent = gender === 'م' ? 'لم يتبلّغ' : 'لم تتبلّغ';
    }
}

function updateAgentGenderedLabels(party) {
    const isFem = state[party].agentGender === 'ف';
    $(`${party}-lawyer-q-label`).textContent = isFem ? 'هل الوكيلة محامية؟' : 'هل الوكيل محامٍ؟';
    $(`${party}-agent-name-label`).innerHTML = (isFem ? 'اسم الوكيلة' : 'اسم الوكيل') + ' <span class="req">*</span>';
    $(`${party}-agent-name`).placeholder = isFem ? 'اسم الوكيلة' : 'اسم الوكيل';
}

function updateCompanyAttendanceOptions(party) {
    const attGroup = document.querySelector(`.choice-group[data-group="${party}-attendance"]`);
    if (!attGroup) return;
    const asalaBtn = attGroup.querySelector('[data-value="أصالة"]');
    const representedBtn = attGroup.querySelector('[data-value="تمثيل"]');
    const isCompany = isCorporateEntity(state[party]);
    if (asalaBtn) asalaBtn.style.display = isCompany ? 'none' : '';
    if (representedBtn) representedBtn.textContent = 'حضر وكيل / ممثل';
    if (isCompany && state[party].attendance === 'أصالة') {
        state[party].attendance = 'تمثيل';
        state[party].hasAccompanyingAgent = false;
        $(`${party}-hasAccompanyingAgent`).checked = false;
        attGroup.querySelectorAll('.choice-btn').forEach(b => b.classList.remove('active'));
        if (representedBtn) representedBtn.classList.add('active');
    }
}

function updateRepresentationTypeOptions(party) {
    const group = document.querySelector(`.choice-group[data-group="${party}-repType"]`);
    if (!group) return;
    const isCompany = isCorporateEntity(state[party]);
    const isWaqf = state[party].entityType === ENTITY_TYPES.WAQF;
    const individualAgentBtn = group.querySelector('[data-value="وكيل"]');
    const companyAgentBtn = group.querySelector('[data-value="وكيل شركة"]');
    if (companyAgentBtn) companyAgentBtn.textContent = isWaqf ? 'وكيل عن الوقف' : 'وكيل عن الشركة';
    const companyRepresentativeBtn = group.querySelector('[data-value="ممثل"]');

    if (individualAgentBtn) individualAgentBtn.style.display = isCompany ? 'none' : '';
    if (companyAgentBtn) companyAgentBtn.style.display = isCompany ? '' : 'none';
    if (companyRepresentativeBtn) companyRepresentativeBtn.style.display = isCompany ? '' : 'none';

    const allowedValues = isCompany ? ['وكيل شركة', 'ممثل'] : ['وكيل'];
    if (!allowedValues.includes(state[party].repType)) {
        state[party].repType = allowedValues[0];
    }
    group.querySelectorAll('.choice-btn').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.value === state[party].repType);
    });

    const isAgent = state[party].repType === 'وكيل' || state[party].repType === 'وكيل شركة';
    $(`${party}-wakeel-fields`).style.display = isAgent ? 'block' : 'none';
    $(`${party}-repName-field`).style.display = state[party].repType === 'ممثل' ? 'block' : 'none';
    $(`${party}-rep-num-field`).style.display = state[party].repType === 'ممثل' ? 'block' : 'none';
    $(`${party}-kinship-field`).style.display =
        state[party].repType === 'وكيل' && state[party].repIsLawyer === 'لا' ? 'block' : 'none';
}

// وثيقة الشخصية الاعتبارية: السجل التجاري للشركة وصك الوقفية للوقف
function updateCorporateDocField(party) {
    const field = $(`${party}-crNum-field`);
    if (!field) return;
    const docLabel = `رقم ${corporateDocLabel(state[party])}`;
    const label = field.querySelector('label');
    if (label) label.innerHTML = `${docLabel} <span class="req">*</span>`;
    const input = $(`${party}-crNum`);
    if (input) input.placeholder = docLabel;
}

function updateEntityTypeHint(party) {
    const hint = $(`${party}-entityType-hint`);
    if (!hint) return;
    hint.textContent = state[party].entityType === ENTITY_TYPES.WAQF
        ? 'الوقف شخصية اعتبارية كالشركة: لا يحضر أصالةً وإنما بوكيل أو ممثل نظامي، والحكم عليه واجب التدقيق لزومًا.'
        : '';
}

// نوع الإفهام مع المدعى عليه الوقف: واجب التدقيق لزومًا لا اختيارًا
let noticeKindBeforeLock = null;
function updateNoticeKindLock() {
    const group = document.querySelector('.choice-group[data-group="ruling-noticeKind"]');
    if (!group) return;
    const forced = mandatoryReviewNotice(state);
    if (forced) {
        if (noticeKindBeforeLock === null) noticeKindBeforeLock = state.ruling.noticeKind;
        state.ruling.noticeKind = 'review';
    } else if (noticeKindBeforeLock !== null) {
        state.ruling.noticeKind = noticeKindBeforeLock;
        noticeKindBeforeLock = null;
    }
    group.querySelectorAll('.choice-btn').forEach(btn => {
        const isReview = btn.dataset.value === 'review';
        btn.classList.toggle('locked', forced && !isReview);
        btn.classList.toggle('active', btn.dataset.value === state.ruling.noticeKind);
    });
    const notice = $('noticeKindLockNotice');
    if (notice) notice.style.display = forced ? 'block' : 'none';
    updateClaimValueHint();
}

function updateSpecialCaseVisibility() {
    const s = state.plaintiff;
    const show = s.attendance !== 'لم يحضر';
    $('plaintiff-specialCase-section').style.display = show ? 'block' : 'none';
    if (!show && s.specialCase !== 'none') {
        s.specialCase = 'none';
        const group = document.querySelector('.choice-group[data-group="plaintiff-specialCase"]');
        group.querySelectorAll('.choice-btn').forEach(b => b.classList.remove('active'));
        group.querySelector('[data-value="none"]').classList.add('active');
    }
}

function updatePrimaryNameFieldVisibility(party) {
    const show = state.includePartyDataInText && state[party].attendance !== 'لم يحضر';
    $(`${party}-name-field`).style.display = show ? 'block' : 'none';
}

function updateOathAbsenceFieldVisibility() {
    const show = state.sessionType === 'previous' && state.defendant.attendance === 'لم يحضر';
    $('defendant-oathAbsence-field').style.display = show ? 'block' : 'none';
}

function updateOathDefendantStatusHint() {
    const el = $('oathDefendantStatusHint');
    el.textContent = getOathDefendantStatus(state) === 'غائب'
        ? 'المدعى عليه غائب حاليًا (حسب حالة حضوره في بيانات الأطراف) — ستُضاف تلقائيًا فقرة تحديد جلسة قادمة لإبلاغه بالحضور لأداء اليمين.'
        : 'المدعى عليه حاضر حاليًا (حسب حالة حضوره في بيانات الأطراف).';
}

// ==================== الخطوة 2: التسلسل الإجرائي ====================
// عرض الدعوى على المدعى عليه يسبق سؤال المدعي عن بينته، وتفريع الإقرار/الإنكار/الدفع

const STANCE_HINTS = {
    'إنكار': 'الإنكار يوجب تكليف المدعي بالبينة، فتظهر بطاقة بينة المدعي.',
    'إقرار': 'الإقرار يُغني عن البينة، فلا يُسأل المدعي عن بينته ولا عن اليمين، ويُقفل باب المرافعة مباشرة.',
    'دفع شكلي': 'يُثبت الدفع ويُعرض على المدعي للرد عليه، ثم يُنظر قبل الخوض في الموضوع.'
};

// هل بُلغ الضبط مرحلة الموضوع (فتُسأل البينات)؟
function meritsReached() {
    const c = state.claim;
    if (state.defendant.attendance === 'لم يحضر') return true;
    if (c.defendantStance === 'إقرار') return false;
    if (c.defendantStance === 'دفع شكلي' && c.answeredOnMerits !== 'نعم') return false;
    return true;
}

function updateStanceVisibility() {
    if (!state.sessionType) return;
    const isNew = state.sessionType === 'new';
    const defendantPresent = state.defendant.attendance !== 'لم يحضر';

    $('claimTextField').style.display = isNew ? 'block' : 'none';
    $('defendantResponseSection').style.display = (isNew && defendantPresent) ? 'block' : 'none';
    $('followUpSection').style.display = isNew ? 'none' : 'block';
    $('defendantStanceHint').textContent = STANCE_HINTS[state.claim.defendantStance] || '';
    $('formalPleaBlock').style.display = (isNew && defendantPresent && state.claim.defendantStance === 'دفع شكلي') ? 'block' : 'none';

    const showPlaintiffEvidence = isNew ? meritsReached() : !!state.followUp.plaintiffEvidence;
    $('plaintiffEvidenceCard').style.display = showPlaintiffEvidence ? 'block' : 'none';
    $('evidenceOrderNotice').style.display = defendantPresent ? 'block' : 'none';

    const showDefendantEvidence = isNew
        ? (defendantPresent && meritsReached())
        : !!state.followUp.defendantEvidence;
    $('defendantEvidenceCard').style.display = showDefendantEvidence ? 'block' : 'none';
    // في الجلسة التالية يكون السؤال مُختارًا من قائمة الإجراءات، فلا يُكرَّر
    $('askDefendantEvidenceBlock').style.display = isNew ? 'block' : 'none';
    $('defendantEvidenceFields').style.display =
        (isNew ? state.claim.askDefendantEvidence === 'نعم' : true) ? 'block' : 'none';

    // مطعن الخصم لا يُسأل عنه إن كان غائبًا
    $('plaintiffWitnessObjectionBlock').style.display = defendantPresent ? 'block' : 'none';
    $('defendantWitnessObjectionBlock').style.display = state.plaintiff.attendance !== 'لم يحضر' ? 'block' : 'none';
}

function updateClaimValueHint() {
    const r = state.ruling;
    const hint = $('claimValueHint');
    if (!hint) return;
    if (mandatoryReviewNotice(state)) {
        hint.textContent = 'الحكم على وقف واجب التدقيق لزومًا، فلا أثر لقيمة المطالبة في نوع الإفهام.';
        return;
    }
    if (r.noticeKind !== 'auto') {
        hint.textContent = 'نوع الإفهام محدَّد يدويًا، فلا أثر لقيمة المطالبة.';
        return;
    }
    const value = Number(String(r.claimValue).replace(/[^0-9.]/g, ''));
    if (!value) {
        hint.textContent = `لم تُدخل قيمة المطالبة — سيُعتمد "قابل للاستئناف". حدّ الدعاوى اليسيرة: ${YASEERA_CLAIM_LIMIT.toLocaleString('en-US')} ريال.`;
        return;
    }
    hint.textContent = value <= YASEERA_CLAIM_LIMIT
        ? `دعوى يسيرة (${value.toLocaleString('en-US')} ≤ ${YASEERA_CLAIM_LIMIT.toLocaleString('en-US')}) — الحكم نهائي مكتسب للقطعية.`
        : `تتجاوز حدّ الدعاوى اليسيرة (${YASEERA_CLAIM_LIMIT.toLocaleString('en-US')}) — الحكم قابل للاستئناف.`;
}

function updateVisibility(party) {
    const s = state[party];
    $(`${party}-accompany-toggle`).style.display = s.attendance === 'أصالة' ? 'block' : 'none';
    const showIdentity = state.includePartyDataInText && (s.attendance === 'أصالة' || isCorporateEntity(s));
    $(`${party}-selfIdentity-section`).style.display = showIdentity ? 'block' : 'none';
    updatePrimaryNameFieldVisibility(party);
    const showRepFields = s.attendance === 'تمثيل' || (s.attendance === 'أصالة' && s.hasAccompanyingAgent);
    $(`${party}-rep-fields`).style.display = showRepFields ? 'block' : 'none';
    $(`${party}-repType-chooser`).style.display = s.attendance === 'تمثيل' ? 'block' : 'none';
    if (s.attendance === 'أصالة' && s.hasAccompanyingAgent) {
        $(`${party}-wakeel-fields`).style.display = 'block';
        $(`${party}-rep-num-field`).style.display = 'none';
    }
    if (party === 'plaintiff') {
        updateSpecialCaseVisibility();
        $('plaintiff-absence-fields').style.display = s.attendance === 'لم يحضر' ? 'block' : 'none';
        $('defendantBlock').style.display = s.attendance === 'لم يحضر' ? 'none' : 'block';
    } else {
        $('defendant-absence-fields').style.display = s.attendance === 'لم يحضر' ? 'block' : 'none';
        $('defendant-oathPerformance-section').style.display = (state.sessionType === 'previous' && s.attendance !== 'لم يحضر') ? 'block' : 'none';
        updateOathAbsenceFieldVisibility();
        updateOathDefendantStatusHint();
    }
    updateStanceVisibility();
    if ($('mainApp').style.display !== 'none') showStep();
}

// ==================== معالجة أزرار الاختيار ====================
function handleChoice(key, value) {
    if (key === 'evidence-choice') {
        state.claim.evidenceChoice = value;
        $('evidenceTextField').style.display = value === 'has' ? 'block' : 'none';
        $('moreEvidenceSection').style.display = value === 'has' ? 'block' : 'none';
        $('oathSection').style.display = (value === 'none' || state.claim.hasMoreEvidence === 'لا') ? 'block' : 'none';
        return;
    }
    if (key === 'more-evidence') {
        state.claim.hasMoreEvidence = value;
        $('oathSection').style.display = value === 'لا' ? 'block' : 'none';
        $('moreEvidenceTextField').style.display = value === 'نعم' ? 'block' : 'none';
        return;
    }
    if (key === 'session-sameJudge') { state.sameJudge = value; return; }
    if (key === 'session-mode') { state.opening.mode = value; updateSessionModeUI(); return; }

    // ===== تكييف جواب المدعى عليه =====
    if (key === 'defendant-stance') {
        state.claim.defendantStance = value;
        $('formalPleaBlock').style.display = value === 'دفع شكلي' ? 'block' : 'none';
        updateStanceVisibility();
        return;
    }
    if (key === 'answered-on-merits') {
        state.claim.answeredOnMerits = value;
        updateStanceVisibility();
        return;
    }

    // ===== تزكية الشهود والمطعن فيهم =====
    if (key === 'plaintiff-tazkiya') {
        state.claim.tazkiya = value;
        $('plaintiffTazkiyaNamesField').style.display = value === 'presented' ? 'block' : 'none';
        return;
    }
    if (key === 'plaintiff-witnessObjection') {
        state.claim.witnessObjection = value;
        $('plaintiffWitnessObjectionField').style.display = value === 'نعم' ? 'block' : 'none';
        return;
    }
    if (key === 'defendant-tazkiya') {
        state.claim.defendantEvidence.tazkiya = value;
        $('defendantTazkiyaNamesField').style.display = value === 'presented' ? 'block' : 'none';
        return;
    }
    if (key === 'defendant-witnessObjection') {
        state.claim.defendantEvidence.objection = value;
        $('defendantWitnessObjectionField').style.display = value === 'نعم' ? 'block' : 'none';
        return;
    }

    // ===== دفوع المدعى عليه وبينته =====
    if (key === 'ask-defendant-evidence') {
        state.claim.askDefendantEvidence = value;
        $('defendantEvidenceFields').style.display = value === 'نعم' ? 'block' : 'none';
        return;
    }
    if (key === 'defendant-evidence-choice') {
        state.claim.defendantEvidence.choice = value;
        $('defendantEvidenceTextField').style.display = value === 'has' ? 'block' : 'none';
        return;
    }

    // ===== الأسباب والحكم =====
    if (key === 'ruling-pronounce') {
        state.ruling.pronounce = value;
        $('rulingFields').style.display = value === 'نعم' ? 'block' : 'none';
        return;
    }
    if (key === 'ruling-presence') { state.ruling.presence = value; return; }
    if (key === 'ruling-noticeKind') {
        state.ruling.noticeKind = value;
        updateClaimValueHint();
        return;
    }
    if (key === 'ruling-mandatoryReview') { state.ruling.mandatoryReview = value; return; }

    const dashIdx = key.indexOf('-');
    const party = key.slice(0, dashIdx);
    const field = key.slice(dashIdx + 1);

    if (field === 'gender') { state[party].gender = value; updateGenderedLabels(party); }
    else if (field === 'attendance') {
        state[party].attendance = value;
        updateVisibility(party);
    } else if (field === 'repType') {
        state[party].repType = value;
        $(`${party}-wakeel-fields`).style.display = (value === 'وكيل' || value === 'وكيل شركة') ? 'block' : 'none';
        $(`${party}-repName-field`).style.display = value === 'ممثل' ? 'block' : 'none';
        $(`${party}-rep-num-field`).style.display = value === 'ممثل' ? 'block' : 'none';
        if (value === 'وكيل شركة' || value === 'ممثل') {
            $(`${party}-kinship-field`).style.display = 'none';
        } else if (value === 'وكيل') {
            $(`${party}-kinship-field`).style.display = state[party].repIsLawyer === 'لا' ? 'block' : 'none';
        }
    } else if (field === 'repIsLawyer') {
        state[party].repIsLawyer = value;
        $(`${party}-license-field`).style.display = value === 'نعم' ? 'block' : 'none';
        $(`${party}-kinship-field`).style.display = (value === 'لا' && state[party].repType === 'وكيل') ? 'block' : 'none';
    } else if (field === 'agentGender') {
        state[party].agentGender = value;
        updateAgentGenderedLabels(party);
    } else if (field === 'kinshipType') {
        state[party].kinshipType = value;
        $(`${party}-nasab-block`).style.display = value === 'نسب' ? 'block' : 'none';
        $(`${party}-ashar-block`).style.display = value === 'أصهار' ? 'block' : 'none';
        updateKinshipHint(party);
    } else if (field === 'asharDirection') {
        state[party].asharDirection = value;
        state[party].kinshipAshar = '';
        populateAsharSelect(party);
        updateKinshipHint(party);
    } else if (field === 'occurrence') { state[party].occurrence = Number(value); }
    else if (field === 'notifyStatus') {
        state[party].notifyStatus = value;
        $(`${party}-notified-fields`).style.display = value === 'تبلغ' ? 'block' : 'none';
        $(`${party}-not-notified-notice`).style.display = value === 'لم يتبلغ' ? 'block' : 'none';
        showStep();
    } else if (field === 'specialCase') {
        state[party].specialCase = value;
        showStep();
    } else if (field === 'nationalityType') {
        state[party].nationalityType = value;
        $(`${party}-saudiId-field`).style.display = value === 'سعودي' ? 'block' : 'none';
        $(`${party}-foreignId-block`).style.display = value === 'غير ذلك' ? 'block' : 'none';
    } else if (field === 'entityType') {
        state[party].entityType = value;
        const isCompany = isCorporateEntity(state[party]);
        const isWaqf = value === ENTITY_TYPES.WAQF;
        $(`${party}-individualId-block`).style.display = isCompany ? 'none' : 'block';
        $(`${party}-crNum-field`).style.display = isCompany ? 'block' : 'none';
        $(`${party}-gender-label`).style.display = isCompany ? 'none' : 'block';
        const genderGroup = document.querySelector(`.choice-group[data-group="${party}-gender"]`);
        if (genderGroup) genderGroup.style.display = isCompany ? 'none' : 'flex';
        if (isCompany) {
            // الشركة مؤنثة (حضرت الشركة / المدعية) والوقف مذكر (حضر الوقف / المدعى عليه)
            const entityGender = isWaqf ? 'م' : 'ف';
            state[party].gender = entityGender;
            if (genderGroup) {
                genderGroup.querySelectorAll('.choice-btn').forEach(b => b.classList.remove('active'));
                genderGroup.querySelector(`[data-value="${entityGender}"]`).classList.add('active');
            }
            updateGenderedLabels(party);
        }
        updateCorporateDocField(party);
        updateEntityTypeHint(party);
        updateCompanyAttendanceOptions(party);
        updateRepresentationTypeOptions(party);
        updateVisibility(party);
        if (party === 'defendant') updateNoticeKindLock();
    } else if (field === 'oathAbsence') {
        state.defendant.oathAbsence = value;
        $('defendant-oathTabligh-field').style.display = value === 'نعم' ? 'block' : 'none';
        $('defendant-oathAbsence-notice').style.display = value === 'نعم' ? 'block' : 'none';
        $('defendant-normal-absence-block').style.display = value === 'نعم' ? 'none' : 'block';
        showStep();
    } else if (field === 'oathPerformanceSession') {
        state.defendant.oathPerformanceSession = value;
        $('defendant-oathAnswer-field').style.display = value === 'نعم' ? 'block' : 'none';
    }
}

// ربط أزرار الاختيار (عدا شاشة البداية)
document.querySelectorAll('.choice-group').forEach(group => {
    const key = group.dataset.group;
    if (key === 'session-landing') return;
    group.querySelectorAll('.choice-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            if (btn.classList.contains('locked')) return;
            group.querySelectorAll('.choice-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            handleChoice(key, btn.dataset.value);
            render();
        });
    });
});

// ==================== شاشة البداية ====================
document.querySelectorAll('.choice-group[data-group="session-landing"] .choice-btn').forEach(btn => {
    btn.addEventListener('click', () => {
        state.sessionType = btn.dataset.value;
        $('landingScreen').style.display = 'none';
        $('mainApp').style.display = 'block';
        $('stepNavBar').style.display = 'flex';
        $('sameJudgeCard').style.display = btn.dataset.value === 'previous' ? 'block' : 'none';
        updateOathAbsenceFieldVisibility();
        updateVisibility('defendant');
        updateStanceVisibility();
        uiState.currentStep = 0;
        showStep();
        render();
        if (!uiState.clerkNoticeShown) {
            uiState.clerkNoticeShown = true;
            openMinutesModal('clerkModal');
        }
    });
});
$('backLink').addEventListener('click', () => {
    $('mainApp').style.display = 'none';
    $('stepNavBar').style.display = 'none';
    $('landingScreen').style.display = 'block';
});

// ==================== ربط حقول الأطراف ====================
['plaintiff', 'defendant'].forEach(party => {
    $(`${party}-name`).addEventListener('input', e => { state[party].name = e.target.value; render(); });
    $(`${party}-saudiId`).addEventListener('input', e => { state[party].saudiId = e.target.value; render(); });
    $(`${party}-iqamaNum`).addEventListener('input', e => { state[party].iqamaNum = e.target.value; render(); });
    $(`${party}-crNum`).addEventListener('input', e => { state[party].crNum = e.target.value; render(); });
    $(`${party}-agent-name`).addEventListener('input', e => { state[party].agentName = e.target.value; render(); });
    setupNationalitySearch(party);

    $(`${party}-agent-id`).addEventListener('input', e => {
        const digitsOnly = e.target.value.replace(/[^0-9]/g, '').slice(0, 10);
        e.target.value = digitsOnly;
        state[party].agentId = digitsOnly;
        const showError = digitsOnly.length > 0 && digitsOnly.length !== 10;
        $(`${party}-agent-id-error`).style.display = showError ? 'block' : 'none';
        e.target.classList.toggle('invalid', showError);
        render();
    });

    $(`${party}-wakala-num`).addEventListener('input', e => { state[party].wakalaNum = e.target.value; render(); });
    $(`${party}-wakala-num`).addEventListener('change', e => {
        if (e.target.value.trim()) openMinutesModal('wakalaModal');
    });
    $(`${party}-license-num`).addEventListener('input', e => { state[party].licenseNum = e.target.value; render(); });

    $(`${party}-kinship`).innerHTML = generateNasabOptionsHTML(party);
    $(`${party}-kinship`).addEventListener('change', e => {
        state[party].kinship = e.target.value;
        updateKinshipHint(party);
        render();
    });
    populateAsharSelect(party);
    $(`${party}-kinship-ashar`).addEventListener('change', e => {
        state[party].kinshipAshar = e.target.value;
        updateKinshipHint(party);
        render();
    });

    $(`${party}-rep-num`).addEventListener('input', e => { state[party].repNum = e.target.value; render(); });
    $(`${party}-repName`).addEventListener('input', e => { state[party].repName = e.target.value; render(); });
    $(`${party}-tabligh`).addEventListener('input', e => { state[party].tabligh = e.target.value; render(); });
    $(`${party}-hasAccompanyingAgent`).addEventListener('change', e => {
        state[party].hasAccompanyingAgent = e.target.checked;
        updateVisibility(party);
        render();
    });

    updateVisibility(party);
    updateCompanyAttendanceOptions(party);
    updateRepresentationTypeOptions(party);
    $(`${party}-license-field`).style.display = state[party].repIsLawyer === 'نعم' ? 'block' : 'none';
});
$('defendant-oathTablighNum').addEventListener('input', e => { state.defendant.oathTablighNum = e.target.value; render(); });
$('defendant-oathAnswerText').addEventListener('input', e => { state.defendant.oathAnswerText = e.target.value; render(); });

// ==================== الأطراف الإضافية (تعدد المدعين/المدعى عليهم) ====================
// بطاقة الهوية للطرف الإضافي — نفس بطاقة الطرف الأول (تظهر عند الحضور أصالة)
function extraIdentityHTML(role, idx, p) {
    const isSaudi = p.nationalityType !== 'غير ذلك';
    const attrs = `data-role="${role}" data-idx="${idx}"`;
    return `
        <div class="sub-fields">
            <label>الجنسية</label>
            <div class="choice-group">
                <div class="choice-btn extra-choice ${isSaudi ? 'active' : ''}" ${attrs} data-field="nationalityType" data-value="سعودي">سعودي</div>
                <div class="choice-btn extra-choice ${isSaudi ? '' : 'active'}" ${attrs} data-field="nationalityType" data-value="غير ذلك">غير ذلك</div>
            </div>
            ${isSaudi ? `
            <div class="field">
                <label>رقم الهوية الوطنية <span class="req">*</span></label>
                <input type="text" class="form-control" ${attrs} data-field="saudiId" value="${escapeHtml(p.saudiId || '')}" placeholder="أدخل 10 أرقام فقط" maxlength="10" inputmode="numeric" data-numeric-only="true" required>
            </div>` : `
            <div class="field" style="position:relative;">
                <label>الجنسية <span class="req">*</span></label>
                <input type="text" class="form-control extra-nat-input" ${attrs} data-field="foreignNationality" value="${escapeHtml(p.foreignNationality || '')}" placeholder="اكتب للبحث عن الجنسية" autocomplete="off" required>
                <div class="nat-dropdown extra-nat-list" style="display:none;"></div>
            </div>
            <div class="field">
                <label>رقم الإقامة <span class="req">*</span></label>
                <input type="text" class="form-control" ${attrs} data-field="iqamaNum" value="${escapeHtml(p.iqamaNum || '')}" placeholder="أدخل 10 أرقام فقط" maxlength="10" inputmode="numeric" data-numeric-only="true" required>
            </div>`}
        </div>`;
}

function extraCardHTML(role, idx, p) {
    const ordinal = ordinalWord(idx + 2, p.gender);
    const roleLabel = role === 'plaintiff' ? 'مدعي' : 'مدعى عليه';
    return `
    <div class="extra-party-block">
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:0.5rem;">
            <div class="party-title" style="margin-bottom:0;">${roleLabel} ${ordinal}</div>
            <button type="button" class="remove-line-btn remove-extra-btn" data-role="${role}" data-idx="${idx}">✕ إزالة</button>
        </div>
        ${state.includePartyDataInText ? `
        <div class="field">
            <label>الاسم <span class="req">*</span></label>
            <input type="text" class="form-control" data-role="${role}" data-idx="${idx}" data-field="name" value="${escapeHtml(p.name)}" placeholder="اسم الطرف" required>
        </div>` : ''}
        <label>الجنس</label>
        <div class="choice-group">
            <div class="choice-btn extra-choice ${p.gender === 'م' ? 'active' : ''}" data-role="${role}" data-idx="${idx}" data-field="gender" data-value="م">مذكر</div>
            <div class="choice-btn extra-choice ${p.gender === 'ف' ? 'active' : ''}" data-role="${role}" data-idx="${idx}" data-field="gender" data-value="ف">مؤنث</div>
        </div>
        <label>حالة الحضور</label>
        <div class="choice-group">
            <div class="choice-btn extra-choice ${p.attendance === 'أصالة' ? 'active' : ''}" data-role="${role}" data-idx="${idx}" data-field="attendance" data-value="أصالة">${p.gender === 'م' ? 'حضر أصالة' : 'حضرت أصالة'}</div>
            <div class="choice-btn extra-choice ${p.attendance === 'تمثيل' ? 'active' : ''}" data-role="${role}" data-idx="${idx}" data-field="attendance" data-value="تمثيل">حضر وكيل</div>
            <div class="choice-btn extra-choice ${p.attendance === 'لم يحضر' ? 'active' : ''}" data-role="${role}" data-idx="${idx}" data-field="attendance" data-value="لم يحضر">${p.gender === 'م' ? 'لم يحضر' : 'لم تحضر'}</div>
        </div>
        ${p.attendance === 'أصالة' && state.includePartyDataInText ? extraIdentityHTML(role, idx, p) : ''}
        ${p.attendance === 'تمثيل' ? `
        <div class="sub-fields">
            <div class="field">
                <label>اسم الوكيل <span class="req">*</span></label>
                <input type="text" class="form-control" data-role="${role}" data-idx="${idx}" data-field="repName" value="${escapeHtml(p.repName)}" placeholder="اسم الوكيل" required>
            </div>
            <div class="field">
                <label>رقم الوكالة الشرعية <span class="req">*</span></label>
                <input type="text" class="form-control" data-role="${role}" data-idx="${idx}" data-field="repNum" value="${escapeHtml(p.repNum)}" placeholder="رقم الوكالة" inputmode="numeric" data-numeric-only="true" required>
            </div>
        </div>` : ''}
    </div>`;
}

function extraListOf(role) {
    return role === 'plaintiff' ? state.extraPlaintiffs : state.extraDefendants;
}

function renderExtraCards(role) {
    const container = $(role === 'plaintiff' ? 'plaintiffExtraContainer' : 'defendantExtraContainer');
    container.innerHTML = extraListOf(role).map((p, i) => extraCardHTML(role, i, p)).join('');
    // إعادة ربط بحث الجنسيات بعد كل إعادة بناء للبطاقات
    container.querySelectorAll('.extra-nat-input').forEach(input => {
        const idx = Number(input.dataset.idx);
        attachNationalitySearch(input, input.parentElement.querySelector('.extra-nat-list'), value => {
            const party = extraListOf(role)[idx];
            if (party) party.foreignNationality = value;
            render();
        });
    });
}

['plaintiffExtraContainer', 'defendantExtraContainer'].forEach(cid => {
    const container = $(cid);
    container.addEventListener('click', e => {
        const removeBtn = e.target.closest('.remove-extra-btn');
        if (removeBtn) {
            const role = removeBtn.dataset.role;
            extraListOf(role).splice(Number(removeBtn.dataset.idx), 1);
            renderExtraCards(role);
            render();
            return;
        }
        const choiceBtn = e.target.closest('.extra-choice');
        if (choiceBtn) {
            const { role, idx, field, value } = choiceBtn.dataset;
            extraListOf(role)[Number(idx)][field] = value;
            renderExtraCards(role);
            render();
        }
    });
    container.addEventListener('input', e => {
        const el = e.target;
        if (el.dataset && el.dataset.field && !el.classList.contains('extra-choice')) {
            const arr = extraListOf(el.dataset.role);
            if (arr[Number(el.dataset.idx)]) arr[Number(el.dataset.idx)][el.dataset.field] = el.value;
            render();
        }
    });
});

$('addExtraPlaintiffBtn').addEventListener('click', () => {
    state.extraPlaintiffs.push(freshExtraParty());
    renderExtraCards('plaintiff');
    render();
});
$('addExtraDefendantBtn').addEventListener('click', () => {
    state.extraDefendants.push(freshExtraParty());
    renderExtraCards('defendant');
    render();
});

// ==================== الدعوى والبينة ====================
$('claimText').addEventListener('input', e => {
    // نص الدعوى سطر واحد متصل في الضبط
    const cleaned = e.target.value.replace(/\s*\n+\s*/g, ' ').replace(/[ \t]{2,}/g, ' ');
    if (cleaned !== e.target.value) {
        const cursorAtEnd = e.target.selectionStart === e.target.value.length;
        e.target.value = cleaned;
        try { if (cursorAtEnd) e.target.selectionStart = e.target.selectionEnd = cleaned.length; } catch (err) { }
    }
    state.claim.text = cleaned;
    render();
});

// قائمة البينة — مشتركة بين المدعي والمدعى عليه
function setupEvidenceChecklist(cfg) {
    const box = $(cfg.checklistId);
    box.innerHTML = EVIDENCE_OPTIONS.map(opt =>
        `<label class="checkbox-row"><input type="checkbox" data-evidence-item="${opt}"> ${opt}</label>`
    ).join('');

    box.querySelectorAll('[data-evidence-item]').forEach(cb => {
        cb.addEventListener('change', e => {
            const item = e.target.dataset.evidenceItem;
            const block = cfg.block();
            if (e.target.checked) {
                if (!block.items.includes(item)) block.items.push(item);
                if (item === 'شهادة شهود') {
                    $(cfg.witnessSectionId).style.display = 'block';
                    if (block.witnesses.length === 0) {
                        block.witnesses.push(freshWitness());
                        cfg.renderWitnesses();
                    }
                }
                if (item === 'مستند رسمي آخر') $(cfg.otherDocFieldId).style.display = 'block';
            } else {
                block.items = block.items.filter(x => x !== item);
                if (item === 'شهادة شهود') $(cfg.witnessSectionId).style.display = 'none';
                if (item === 'مستند رسمي آخر') $(cfg.otherDocFieldId).style.display = 'none';
            }
            render();
        });
    });
}

// ==================== الشهود ====================
function witnessCardHTML(prefix, idx, w) {
    const ordinal = ordinalWord(idx + 1, 'م');
    const fields = [
        ['name', 'الاسم الكامل'], ['age', 'تاريخ الميلاد / العمر'], ['job', 'المهنة'],
        ['residence', 'مكان الإقامة'], ['relation', 'وجه الاتصال بأطراف الدعوى'], ['interest', 'المصلحة في هذه الدعوى'],
        ['testimony', 'نص الشهادة (أشهد بالله العظيم أن...)']
    ];
    return `
    <div class="extra-party-block">
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:0.5rem;">
            <div class="party-title" style="margin-bottom:0;">الشاهد ${ordinal}</div>
            ${idx > 0 ? `<button type="button" class="remove-line-btn remove-witness-btn" data-prefix="${prefix}" data-idx="${idx}">✕ إزالة</button>` : ''}
        </div>
        ${fields.map(([key, label]) => `
        <div class="field">
            <label>${label} <span class="req">*</span></label>
            <input type="text" class="form-control" data-witness-prefix="${prefix}" data-witness-idx="${idx}" data-witness-field="${key}" value="${escapeHtml(w[key])}" placeholder="${label}" required>
        </div>`).join('')}
    </div>`;
}

// كتلتا البينة: المدعي (حقول مسطّحة في claim) والمدعى عليه (كتلة مستقلة)
const plaintiffEvidenceBlock = () => ({
    get items() { return state.claim.evidenceItems; },
    set items(v) { state.claim.evidenceItems = v; },
    get witnesses() { return state.claim.witnesses; }
});
const defendantEvidenceBlock = () => state.claim.defendantEvidence;

function renderWitnesses() {
    $('witnessContainer').innerHTML = state.claim.witnesses.map((w, i) => witnessCardHTML('plaintiff', i, w)).join('');
}
function renderDefendantWitnesses() {
    $('defendantWitnessContainer').innerHTML = state.claim.defendantEvidence.witnesses.map((w, i) => witnessCardHTML('defendant', i, w)).join('');
}

function witnessListOf(prefix) {
    return prefix === 'plaintiff' ? state.claim.witnesses : state.claim.defendantEvidence.witnesses;
}

[['witnessContainer', 'plaintiff', renderWitnesses], ['defendantWitnessContainer', 'defendant', renderDefendantWitnesses]].forEach(([containerId, prefix, rerender]) => {
    $(containerId).addEventListener('input', e => {
        const el = e.target;
        if (el.dataset && el.dataset.witnessField) {
            const wit = witnessListOf(el.dataset.witnessPrefix)[Number(el.dataset.witnessIdx)];
            if (wit) wit[el.dataset.witnessField] = el.value;
            render();
        }
    });
    $(containerId).addEventListener('click', e => {
        const btn = e.target.closest('.remove-witness-btn');
        if (btn) {
            witnessListOf(btn.dataset.prefix).splice(Number(btn.dataset.idx), 1);
            rerender();
            render();
        }
    });
});

setupEvidenceChecklist({
    checklistId: 'evidenceChecklist', witnessSectionId: 'witnessSection', otherDocFieldId: 'otherDocumentField',
    block: plaintiffEvidenceBlock, renderWitnesses
});
setupEvidenceChecklist({
    checklistId: 'defendantEvidenceChecklist', witnessSectionId: 'defendantWitnessSection', otherDocFieldId: 'defendantOtherDocumentField',
    block: defendantEvidenceBlock, renderWitnesses: renderDefendantWitnesses
});

$('addWitnessBtn').addEventListener('click', () => {
    state.claim.witnesses.push(freshWitness());
    renderWitnesses();
    render();
});
$('addDefendantWitnessBtn').addEventListener('click', () => {
    state.claim.defendantEvidence.witnesses.push(freshWitness());
    renderDefendantWitnesses();
    render();
});

// ربط الحقول النصية بالحالة
[
    ['otherDocumentText', c => v => { c.otherDocumentText = v; }],
    ['moreEvidenceText', c => v => { c.moreEvidenceText = v; }],
    ['defendantResponseText', c => v => { c.defendantResponseText = v; }],
    ['formalPleaText', c => v => { c.formalPleaText = v; }],
    ['plaintiffReplyText', c => v => { c.plaintiffReplyText = v; }],
    ['plaintiffTazkiyaNames', c => v => { c.tazkiyaNames = v; }],
    ['plaintiffWitnessObjectionText', c => v => { c.witnessObjectionText = v; }],
    ['defendantPleasText', c => v => { c.defendantPleasText = v; }],
    ['defendantOtherDocumentText', c => v => { c.defendantEvidence.otherDocumentText = v; }],
    ['defendantTazkiyaNames', c => v => { c.defendantEvidence.tazkiyaNames = v; }],
    ['defendantWitnessObjectionText', c => v => { c.defendantEvidence.objectionText = v; }]
].forEach(([id, setterFor]) => {
    $(id).addEventListener('input', e => { setterFor(state.claim)(e.target.value); render(); });
});

// إجراءات الجلسة التالية
[['followUpPlaintiffEvidence', 'plaintiffEvidence'], ['followUpDefendantEvidence', 'defendantEvidence']].forEach(([id, key]) => {
    $(id).addEventListener('change', e => {
        state.followUp[key] = e.target.checked;
        updateStanceVisibility();
        showStep();
        render();
    });
});

// ==================== طلب اليمين ====================
$('requestOathCheckbox').addEventListener('change', e => {
    state.claim.requestOath = e.target.checked;
    if (e.target.checked) {
        state.claim.declineOath = false;
        $('declineOathCheckbox').checked = false;
    }
    $('oathDefendantStatus').style.display = e.target.checked ? 'block' : 'none';
    updateOathDefendantStatusHint();
    render();
});
$('declineOathCheckbox').addEventListener('change', e => {
    state.claim.declineOath = e.target.checked;
    if (e.target.checked) {
        state.claim.requestOath = false;
        $('requestOathCheckbox').checked = false;
        $('oathDefendantStatus').style.display = 'none';
    }
    render();
});

// ==================== مرحلة الأسباب والحكم ====================
// النماذج تُستهلك من مكتبة النماذج في data/templates.js فلا تُكرَّر هنا

// حدّ ما يُعرض من النتائج في القائمة المنسدلة
const TEMPLATE_RESULTS_LIMIT = 80;
// طول المقتطف المعروض تحت الكلمة المفتاحية للتفريق بين النماذج المتشابهة العنوان
const TEMPLATE_PREVIEW_LEN = 100;

function clipPreview(text) {
    return text.length > TEMPLATE_PREVIEW_LEN ? text.slice(0, TEMPLATE_PREVIEW_LEN).trim() + '…' : text;
}

// مقتطف من متن النموذج: الجملة التي وقع فيها البحث إن وُجدت، وإلا مطلع النموذج.
// نماذج التسبيب تتشابه مطالعها، فعرض موضع المطابقة هو ما يميّزها فعلًا.
function templatePreview(content, query) {
    const flat = String(content || '').replace(/\s+/g, ' ').trim();
    const terms = normalizeArabicSearch(query).split(' ').filter(Boolean);
    if (terms.length === 0) return clipPreview(flat);
    const parts = flat.split(/[.،؛:]\s+/);
    for (let i = 0; i < parts.length; i++) {
        const seg = normalizeArabicSearch(parts[i]);
        if (terms.some(t => seg.includes(t))) return (i > 0 ? '…' : '') + clipPreview(parts[i].trim());
    }
    return clipPreview(flat);
}

function setupTemplatePicker(cfg) {
    const input = $(cfg.searchId);
    const list = $(cfg.listId);
    const target = $(cfg.targetId);
    let filtered = [];
    let highlightedIdx = -1;

    function pool() {
        const all = templatesOfCategory(cfg.category);
        return cfg.curatedOnly && cfg.curatedOnly() ? all.slice(0, curatedReasonsCount(all)) : all;
    }

    function renderList() {
        const source = pool();
        // بحث مطبَّع بكلمات متعددة، فلا يفوته تشكيل النماذج المحدَّثة ولا اختلاف الهمزات
        filtered = searchTemplates(source, input.value);
        const shown = filtered.slice(0, TEMPLATE_RESULTS_LIMIT);
        const moreCount = filtered.length - shown.length;
        list.innerHTML = filtered.length === 0
            ? '<div class="nat-dropdown-empty">لا توجد نماذج مطابقة</div>'
            : shown.map((t, i) =>
                `<div class="nat-dropdown-item" data-idx="${i}">
                    <div class="tpl-item-title">${escapeHtml(String(t.keyword).replace(/\s+/g, ' '))}</div>
                    <div class="tpl-item-preview">${escapeHtml(templatePreview(t.content, input.value))}</div>
                </div>`
            ).join('') + (moreCount > 0
                ? `<div class="nat-dropdown-empty">وثمة نماذج أخرى مطابقة (${moreCount}) — تابع الكتابة لتضييق البحث</div>`
                : '');
        highlightedIdx = -1;
        list.style.display = 'block';
    }

    function choose(tpl) {
        if (!tpl) return;
        target.value = tpl.content;
        target.dispatchEvent(new Event('input', { bubbles: true }));
        list.style.display = 'none';
        input.value = '';
        showToast('أُدرج النموذج، ويمكن تعديله', 'success');
    }

    input.addEventListener('focus', renderList);
    input.addEventListener('input', renderList);
    input.addEventListener('blur', () => {
        setTimeout(() => { if (document.activeElement !== input) list.style.display = 'none'; }, 150);
    });
    input.addEventListener('keydown', e => {
        const items = list.querySelectorAll('.nat-dropdown-item');
        if (e.key === 'ArrowDown') { e.preventDefault(); highlightedIdx = Math.min(highlightedIdx + 1, items.length - 1); }
        else if (e.key === 'ArrowUp') { e.preventDefault(); highlightedIdx = Math.max(highlightedIdx - 1, 0); }
        else if (e.key === 'Enter') {
            if (highlightedIdx >= 0) { e.preventDefault(); choose(filtered[highlightedIdx]); }
            return;
        } else if (e.key === 'Escape') { list.style.display = 'none'; return; }
        else { return; }
        items.forEach((el, i) => el.classList.toggle('highlighted', i === highlightedIdx));
    });
    list.addEventListener('mousedown', e => {
        const item = e.target.closest('.nat-dropdown-item');
        if (item && item.dataset.idx !== undefined) choose(filtered[Number(item.dataset.idx)]);
    });
}

setupTemplatePicker({
    searchId: 'reasonsTemplateSearch', listId: 'reasonsTemplateList', targetId: 'reasonsText',
    category: 'أسباب الحكم', curatedOnly: () => $('curatedReasonsOnly').checked
});

setupTemplatePicker({
    searchId: 'verdictTemplateSearch', listId: 'verdictTemplateList', targetId: 'rulingText',
    category: 'اختصارات صندوق الحكم'
});

// عدّادا النماذج يُقرآن من مكتبة النماذج، فيتحدثان تلقائيًا مع كل تحديث لقوالب التسبيب
function updateReasonsTemplateCounts() {
    const all = templatesOfCategory('أسباب الحكم');
    const curated = curatedReasonsCount(all);
    $('curatedReasonsCountLabel').textContent = `الاقتصار على نماذج التسبيب المنسّقة (${curated} نموذجًا)`;
    $('reasonsTemplateHint').textContent =
        `تصنيف (أسباب الحكم) يضم ${all.length} نموذجًا. اختيار نموذج يُدرج نصه في الحقل أدناه، ويبقى قابلًا للتعديل الكامل.`;
}
updateReasonsTemplateCounts();

$('reasonsText').addEventListener('input', e => { state.ruling.reasonsText = e.target.value; render(); });
$('rulingText').addEventListener('input', e => { state.ruling.rulingText = e.target.value; render(); });
$('claimValue').addEventListener('input', e => {
    state.ruling.claimValue = e.target.value;
    updateClaimValueHint();
    render();
});
updateNoticeKindLock();

// ==================== المعاينة والتحذيرات ====================
function updateWarningsBox() {
    const warnings = collectWarnings(state);
    const box = $('missingFieldsWarning');
    if (warnings.length === 0) {
        box.innerHTML = '<div class="ok-box">جميع الحقول مكتملة ✓</div>';
    } else {
        box.innerHTML = `<div class="warning-box">تنبيه: هناك حقول لم تُعبّأ (سيظهر مكانها نقاط ....):<ul>${warnings.map(x => `<li>${escapeHtml(x)}</li>`).join('')}</ul></div>`;
    }
}

function render() {
    if (uiState.manualEditActive) return; // لا يُلمس النص أثناء التحرير اليدوي
    if (!state.sessionType) return;

    if (uiState.manualText !== null) {
        $('outputText').textContent = uiState.manualText;
    } else {
        $('outputText').textContent = composeMinutes(state);
    }
    updateWarningsBox();
    saveDraftToStorage();
}

// ==================== خطوات النموذج ====================
const stepTitles = { 0: 'بيانات الافتتاح', 1: 'بيانات الأطراف', 2: 'الدعوى والبينات', 3: 'الأسباب والحكم' };

function getStepList() {
    const steps = [0, 1];
    const plaintiffBlocks = state.plaintiff.attendance === 'لم يحضر' || state.plaintiff.specialCase !== 'none';
    const defendantBlocks = state.defendant.attendance === 'لم يحضر'
        && (state.defendant.notifyStatus === 'لم يتبلغ' || (state.sessionType === 'previous' && state.defendant.oathAbsence === 'نعم'));
    if (plaintiffBlocks || defendantBlocks) return steps;

    // الخطوة 2 متاحة في الجلستين: سماع البينات والشهود يقع في التالية غالبًا
    steps.push(2);

    // لا يُلحق الحكم إذا تأجلت الجلسة لتوجيه اليمين لمدعى عليه غائب
    const oathAdjournment = state.sessionType === 'new' && state.claim.requestOath
        && !state.claim.declineOath && state.defendant.attendance === 'لم يحضر';
    if (!oathAdjournment) steps.push(3);
    return steps;
}

function stepElement(n) {
    if (n === 0) return $('stepWrap0');
    if (n === 1) return $('stepWrap1');
    if (n === 2) return $('claimEvidenceCard');
    return $('rulingCard');
}

function showStep() {
    const list = getStepList();
    if (!list.includes(uiState.currentStep)) {
        const candidates = list.filter(n => n <= uiState.currentStep);
        uiState.currentStep = candidates.length ? candidates[candidates.length - 1] : list[0];
    }
    [0, 1, 2, 3].forEach(n => { stepElement(n).style.display = 'none'; });
    // بطاقة وقت الإغلاق تُذيَّل بها خطوة (الأسباب والحكم)، وفي المحاضر المختصرة التي تسقط
    // فيها هذه الخطوة (الشطب وعدم التبلّغ ونحوهما) تنتقل لآخر خطوة متاحة فتبقى في المتناول
    stepElement(list[list.length - 1]).appendChild($('closingTimeCard'));
    stepElement(uiState.currentStep).style.display = 'block';
    const idx = list.indexOf(uiState.currentStep);
    $('prevStepBtn').disabled = idx === 0;
    $('nextStepBtn').textContent = idx === list.length - 1 ? 'عرض النتيجة النهائية ✓' : 'التالي ←';
    $('stepLabel').textContent = `خطوة ${idx + 1} من ${list.length}: ${stepTitles[uiState.currentStep]}`;
    $('stepDots').innerHTML = list.map((n, i) => {
        const cls = i < idx ? 'done' : (i === idx ? 'current' : '');
        return `<div class="step-item"><div class="step-dot ${cls}">${i + 1}</div><div class="step-label">${stepTitles[n]}</div></div>`;
    }).join('');
    $('stepValidationMsg').style.display = 'none';
    saveDraftToStorage();
}

function isFieldHidden(el) {
    let node = el;
    while (node && node !== document.body) {
        if (node.style && node.style.display === 'none') return true;
        node = node.parentElement;
    }
    return false;
}

function validateCurrentStep() {
    const container = stepElement(uiState.currentStep);
    let firstInvalid = null;
    container.querySelectorAll('[required]').forEach(el => {
        if (isFieldHidden(el)) return;
        const val = (el.value || '').trim();
        let fieldValid = !!val;
        if (fieldValid && el.id && el.id.endsWith('-agent-id') && val.length !== 10) fieldValid = false;
        el.classList.toggle('invalid', !fieldValid);
        if (!fieldValid && !firstInvalid) firstInvalid = el;
    });
    if (firstInvalid) {
        try { firstInvalid.scrollIntoView({ behavior: 'smooth', block: 'center' }); } catch (e) { }
        try { firstInvalid.focus({ preventScroll: true }); } catch (e) { }
    }
    return !firstInvalid;
}

$('nextStepBtn').addEventListener('click', () => {
    if (!validateCurrentStep()) {
        $('stepValidationMsg').style.display = 'block';
        return;
    }
    const list = getStepList();
    const idx = list.indexOf(uiState.currentStep);
    if (idx < list.length - 1) {
        uiState.currentStep = list[idx + 1];
        showStep();
        window.scrollTo({ top: 0, behavior: 'smooth' });
    } else {
        render();
        enterReviewMode();
    }
});
$('prevStepBtn').addEventListener('click', () => {
    const list = getStepList();
    const idx = list.indexOf(uiState.currentStep);
    if (idx > 0) {
        uiState.currentStep = list[idx - 1];
        showStep();
        window.scrollTo({ top: 0, behavior: 'smooth' });
    }
});

// ==================== وضع المراجعة ====================
function enterReviewMode() {
    $('minutesLayout').classList.add('review-mode');
    $('backToEditBtn').style.display = 'inline-block';
    window.scrollTo({ top: 0, behavior: 'smooth' });
}
function exitReviewMode() {
    $('minutesLayout').classList.remove('review-mode');
    $('backToEditBtn').style.display = 'none';
}
$('backToEditBtn').addEventListener('click', exitReviewMode);

// ==================== محضر جديد ====================
['newSessionBtn', 'newSessionBtnStep'].forEach(id => {
    $(id).addEventListener('click', () => openMinutesModal('confirmModal'));
});
$('confirmModalYes').addEventListener('click', () => {
    try { localStorage.removeItem(DRAFT_KEY); } catch (e) { }
    state.sessionType = null; // يمنع إعادة حفظ المسودة وتحذير المغادرة
    location.reload();
});

// ==================== النسخ والطباعة وحجم الخط ====================
function copyOutputText() {
    const text = $('outputText').textContent;
    if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(text)
            .then(() => showToast('تم نسخ المحضر ✓', 'success'))
            .catch(() => fallbackCopy(text));
    } else {
        fallbackCopy(text);
    }
}
function fallbackCopy(text) {
    const ta = document.createElement('textarea');
    ta.value = text;
    ta.style.position = 'fixed';
    ta.style.opacity = '0';
    document.body.appendChild(ta);
    ta.select();
    try { document.execCommand('copy'); showToast('تم نسخ المحضر ✓', 'success'); } catch (e) { showToast('تعذر النسخ'); }
    document.body.removeChild(ta);
}
$('copyBtn').addEventListener('click', copyOutputText);
$('printBtn').addEventListener('click', () => window.print());

let outputFontSize = 17.5;
function applyFontSize() { $('outputText').style.fontSize = outputFontSize + 'px'; }
$('fontIncreaseBtn').addEventListener('click', () => { outputFontSize = Math.min(outputFontSize + 2, 28); applyFontSize(); });
$('fontDecreaseBtn').addEventListener('click', () => { outputFontSize = Math.max(outputFontSize - 2, 13); applyFontSize(); });

// ==================== التحرير اليدوي للنص ====================
$('editOutputBtn').addEventListener('click', () => {
    const outputEl = $('outputText');
    if (!uiState.manualEditActive) {
        uiState.manualEditActive = true;
        outputEl.contentEditable = 'true';
        outputEl.classList.add('editing');
        outputEl.focus();
        $('editOutputBtn').textContent = '✔️ تم';
    } else {
        uiState.manualText = outputEl.textContent;
        uiState.manualEditActive = false;
        outputEl.contentEditable = 'false';
        outputEl.classList.remove('editing');
        $('editOutputBtn').textContent = '✏️ تعديل';
        $('resetManualEditLink').style.display = 'inline';
        saveDraftToStorage();
    }
});
$('resetManualEditLink').addEventListener('click', () => {
    uiState.manualText = null;
    $('resetManualEditLink').style.display = 'none';
    render();
});

// ==================== حفظ المسودة واسترجاعها ====================
function collectFormSnapshot() {
    const snap = {
        texts: {}, checks: {}, groups: {}, selects: {},
        currentStep: uiState.currentStep,
        sessionType: state.sessionType,
        extraPlaintiffs: state.extraPlaintiffs,
        extraDefendants: state.extraDefendants,
        manualText: uiState.manualText,
        evidenceItems: state.claim.evidenceItems,
        witnesses: state.claim.witnesses,
        defendantEvidenceItems: state.claim.defendantEvidence.items,
        defendantWitnesses: state.claim.defendantEvidence.witnesses,
        followUp: state.followUp
    };
    document.querySelectorAll('input[type=text][id], textarea[id]').forEach(el => {
        if (el.disabled) return;
        snap.texts[el.id] = el.value;
    });
    document.querySelectorAll('input[type=checkbox][id]').forEach(el => {
        snap.checks[el.id] = el.checked;
    });
    document.querySelectorAll('select[id]').forEach(el => {
        snap.selects[el.id] = el.value;
    });
    document.querySelectorAll('.choice-group[data-group]').forEach(g => {
        const active = g.querySelector('.choice-btn.active');
        if (active) snap.groups[g.dataset.group] = active.dataset.value;
    });
    return snap;
}

function saveDraftToStorage() {
    if (!state.sessionType) return;
    try {
        localStorage.setItem(DRAFT_KEY, JSON.stringify(collectFormSnapshot()));
    } catch (e) { }
}

function applySnapshot(snap) {
    // دمج مع القيم الافتراضية حتى تعمل المسودّات المحفوظة قبل إضافة بطاقة الهوية
    const withDefaults = list => (list || []).map(p => Object.assign(freshExtraParty(), p));
    state.extraPlaintiffs = withDefaults(snap.extraPlaintiffs);
    state.extraDefendants = withDefaults(snap.extraDefendants);
    uiState.manualText = snap.manualText || null;
    if (uiState.manualText) $('resetManualEditLink').style.display = 'inline';
    renderExtraCards('plaintiff');
    renderExtraCards('defendant');

    state.claim.evidenceItems = snap.evidenceItems || [];
    state.claim.witnesses = snap.witnesses || [];
    state.claim.defendantEvidence.items = snap.defendantEvidenceItems || [];
    state.claim.defendantEvidence.witnesses = snap.defendantWitnesses || [];
    state.followUp = Object.assign(freshFollowUpState(), snap.followUp || {});
    $('followUpPlaintiffEvidence').checked = !!state.followUp.plaintiffEvidence;
    $('followUpDefendantEvidence').checked = !!state.followUp.defendantEvidence;

    [
        ['evidenceChecklist', state.claim.evidenceItems, 'witnessSection', 'otherDocumentField', renderWitnesses],
        ['defendantEvidenceChecklist', state.claim.defendantEvidence.items, 'defendantWitnessSection', 'defendantOtherDocumentField', renderDefendantWitnesses]
    ].forEach(([checklistId, items, witnessSectionId, otherDocFieldId, rerender]) => {
        items.forEach(item => {
            const cb = $(checklistId).querySelector(`[data-evidence-item="${item}"]`);
            if (cb) cb.checked = true;
            if (item === 'شهادة شهود') {
                $(witnessSectionId).style.display = 'block';
                rerender();
            }
            if (item === 'مستند رسمي آخر') $(otherDocFieldId).style.display = 'block';
        });
    });

    state.sessionType = snap.sessionType;
    $('sameJudgeCard').style.display = snap.sessionType === 'previous' ? 'block' : 'none';

    Object.entries(snap.groups || {}).forEach(([groupKey, val]) => {
        const btn = document.querySelector(`.choice-group[data-group="${groupKey}"] .choice-btn[data-value="${val}"]`);
        if (btn && groupKey !== 'session-landing') btn.click();
    });
    Object.entries(snap.checks || {}).forEach(([id, val]) => {
        const el = document.getElementById(id);
        if (el && el.checked !== val) { el.checked = val; el.dispatchEvent(new Event('change', { bubbles: true })); }
    });
    Object.entries(snap.texts || {}).forEach(([id, val]) => {
        const el = document.getElementById(id);
        if (el) { el.value = val; el.dispatchEvent(new Event('input', { bubbles: true })); }
    });
    Object.entries(snap.selects || {}).forEach(([id, val]) => {
        const el = document.getElementById(id);
        if (el) { el.value = val; el.dispatchEvent(new Event('change', { bubbles: true })); }
    });

    uiState.currentStep = snap.currentStep || 0;
    $('landingScreen').style.display = 'none';
    $('mainApp').style.display = 'block';
    $('stepNavBar').style.display = 'flex';
    updateOathAbsenceFieldVisibility();
    updateVisibility('defendant');
    updateStanceVisibility();
    updateNoticeKindLock();
    showStep();
    render();
}

(function checkForDraft() {
    try {
        const raw = localStorage.getItem(DRAFT_KEY);
        if (!raw) return;
        const snap = JSON.parse(raw);
        if (!snap || !snap.sessionType) return;
        $('draftBanner').style.display = 'block';
        $('restoreDraftBtn').addEventListener('click', () => {
            $('draftBanner').style.display = 'none';
            applySnapshot(snap);
        });
        $('discardDraftBtn').addEventListener('click', () => {
            try { localStorage.removeItem(DRAFT_KEY); } catch (e) { }
            $('draftBanner').style.display = 'none';
        });
    } catch (e) { }
})();

// ==================== سلوكيات عامة ====================
// زر الإدخال ينقل للحقل التالي بدل إرسال النموذج
document.addEventListener('keydown', e => {
    if (e.key !== 'Enter') return;
    const el = e.target;
    if (el.tagName === 'TEXTAREA') return;
    if (el.tagName !== 'INPUT' || el.type !== 'text') return;
    e.preventDefault();
    const focusable = Array.from(document.querySelectorAll('input[type=text]:not([disabled]), textarea'))
        .filter(node => node.offsetParent !== null);
    const idx = focusable.indexOf(el);
    if (idx > -1 && idx < focusable.length - 1) focusable[idx + 1].focus();
});

// تحذير قبل مغادرة الصفحة أثناء محضر غير مكتمل
window.addEventListener('beforeunload', e => {
    if (state.sessionType) {
        e.preventDefault();
        e.returnValue = '';
    }
});

// إغلاق النوافذ بزر Escape
document.addEventListener('keydown', e => {
    if (e.key === 'Escape') {
        document.querySelectorAll('.modal-overlay.show').forEach(m => m.classList.remove('show'));
    }
});
