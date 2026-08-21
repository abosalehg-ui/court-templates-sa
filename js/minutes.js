// ==================== محرك مولّد محضر الجلسة ====================
// هذا الملف يحتوي دوال الصياغة الصافية فقط (بلا أي تعامل مع DOM)
// حتى يمكن اختباره آلياً عبر `npm test` (انظر tests/minutes.test.js)
// واجهة المستخدم في js/minutes-app.js وصفحة minutes.html
//
// فكرة المولّد وإعداد الصياغ القضائية: القاضي/ زياد بن محمد البابطين

// ==================== ثوابت عامة ====================

const MINUTES_PLACEHOLDER = '...........';

// الصيغة المختصرة لانعقاد الجلسة مرئياً (تُدرج وحدها في الافتتاح المختصر)
const VIDEO_CALL_BRIEF = 'عبر الاتصال المرئي';

// المستند النظامي المرافق لصيغة الاتصال المرئي — يُدرج مع خيار «مع المستند» فقط
const VIDEO_CALL_BASIS = ' بأطراف الدَّعوى -من خلال الأنظمة الالكترونية المعتمدة لوزارة العدل، بناء على قرار رئيس المجلس الأعلى للقضاء برقم 07996819 تاريخ 17 / 06 / 1447هـ، وبناء على قرار فضيلة رئيس المجلس الأعلى للقضاء رقم 17388 وتاريخ 5/10/1441هـ والمبلغ بالتعميم رقم 1505 وتاريخ 5/5/1441هـ المتضمن استئناف عقد الجلسات في المحاكم بطريق التقاضي عن بعد. واستنادا إلى قرار وزير العدل رقم 8056 بتاريخ 5/10/1441هـ المبلغ بالتعميم رقم 13 /ت/8135 بتاريخ 5/10/1441هـ المتضمن إطلاق خدمة التقاضي عن بعد والموافقة على الدليل الاجرائي لخدمة التقاضي عن بعد (التقاضي الالكتروني). " -';

// صيغة انعقاد الجلسة عبر الاتصال المرئي مع مستندها النظامي (النص الكامل في الافتتاح)
const VIDEO_CALL_PHRASE = VIDEO_CALL_BRIEF + VIDEO_CALL_BASIS;

// الصيغة المختصرة لطريقة الانعقاد (تُستخدم في نص الشطب)
const VIDEO_CALL_SHORT = 'بالاتصال المرئي';

// صيغة الانعقاد الحضوري (تُستخدم في فقرات الغياب والشطب عند الجلسة الحضورية)
const IN_PERSON_SHORT = 'حضورياً بمقر المحكمة';

// طرق انعقاد الجلسة الثلاث المتاحة في بيانات الافتتاح
const SESSION_MODES = {
    VIDEO_FULL: 'videoFull',  // عبر الاتصال المرئي مع المستند النظامي
    VIDEO: 'video',           // عبر الاتصال المرئي مختصراً (الافتراضي)
    IN_PERSON: 'inPerson'     // جلسة حضورية
};

// طريقة الانعقاد المختارة، مع اعتماد المختصر افتراضياً لأي حالة قديمة بلا قيمة
function sessionMode(opening) {
    const mode = opening && opening.mode;
    return (mode === SESSION_MODES.VIDEO_FULL || mode === SESSION_MODES.IN_PERSON) ? mode : SESSION_MODES.VIDEO;
}

// هل الجلسة حضورية؟ (تُسقط كل الإشارات للاتصال المرئي من بقية الفقرات)
function isInPersonSession(state) {
    return sessionMode(state && state.opening) === SESSION_MODES.IN_PERSON;
}

const NATIONALITY_OPTIONS = ['آيسلندي', 'أذربيجاني', 'أرجنتيني', 'أردني', 'أرمني', 'أسترالي', 'أفريقي وسطي', 'أفغاني', 'ألباني', 'ألماني', 'أمريكي', 'أنغولي', 'أوروغوياني', 'أوزبكستاني', 'أوغندي', 'أوكراني', 'أيرلندي', 'إثيوبي', 'إريتري', 'إسباني', 'إستوني', 'إسواتيني', 'إكوادوري', 'إماراتي', 'إندونيسي', 'إيراني', 'إيطالي', 'الرأس الأخضر', 'بابوا غينيا الجديدة', 'باراغوياني', 'باكستاني', 'بحريني', 'برازيلي', 'برتغالي', 'بروناوي', 'بريطاني', 'بلجيكي', 'بلغاري', 'بنغلاديشي', 'بنمي', 'بنيني', 'بوتاني', 'بوتسواني', 'بوركينابي', 'بوروندي', 'بوسني', 'بولندي', 'بوليفي', 'بيروفي', 'بيلاروسي', 'تايلاندي', 'تايواني', 'تركمانستاني', 'تركي', 'ترينيدادي', 'تشادي', 'تشيكي', 'تشيلي', 'تنزاني', 'توغولي', 'تونسي', 'تيموري', 'جامايكي', 'جزائري', 'جزر القمر', 'جنوب أفريقي', 'جنوب سوداني', 'جورجي', 'جيبوتي', 'دنماركي', 'دومينيكاني', 'رواندي', 'روسي', 'روماني', 'زامبي', 'زيمبابوي', 'ساحل عاجي', 'ساو تومي وبرينسيبي', 'سريلانكي', 'سلفادوري', 'سلوفاكي', 'سلوفيني', 'سنغافوري', 'سنغالي', 'سوداني', 'سوري', 'سورينامي', 'سويدي', 'سويسري', 'سيراليوني', 'سيشيلي', 'صربي', 'صومالي', 'صيني', 'طاجيكستاني', 'عديم الجنسية', 'عراقي', 'عُماني', 'غابوني', 'غامبي', 'غاني', 'غواتيمالي', 'غياني', 'غيني', 'غيني استوائي', 'غيني بيساوي', 'فرنسي', 'فلبيني', 'فلسطيني', 'فنزويلي', 'فنلندي', 'فيتنامي', 'فيجي', 'قبرصي', 'قطري', 'قيرغيزستاني', 'كازاخستاني', 'كاميروني', 'كرواتي', 'كمبودي', 'كندي', 'كوبي', 'كوري جنوبي', 'كوري شمالي', 'كوستاريكي', 'كولومبي', 'كونغولي', 'كونغولي ديمقراطي', 'كويتي', 'كيني', 'لاتفي', 'لاوسي', 'لبناني', 'لوكسمبورغي', 'ليبي', 'ليبيري', 'ليتواني', 'ليسوتوي', 'مالطي', 'مالي', 'ماليزي', 'مدغشقري', 'مصري', 'مغربي', 'مغولي', 'مقدوني', 'مكسيكي', 'ملاوي', 'موريتاني', 'موريشيوسي', 'موزمبيقي', 'مولدوفي', 'مونتينيغري', 'ميانماري', 'ناميبي', 'نرويجي', 'نمساوي', 'نيبالي', 'نيجري', 'نيجيري', 'نيكاراغوي', 'نيوزيلندي', 'هايتي', 'هندوراسي', 'هندي', 'هنغاري', 'هولندي', 'هونغ كونغي', 'ياباني', 'يمني', 'يوناني', 'أخرى'];

// جنسيات الشاهد: قائمة الطرفين نفسها مع «سعودي» في صدرها،
// إذ ليس للشاهد مفتاح (سعودي / غير ذلك) وإنما حقل جنسية واحد
const WITNESS_NATIONALITY_OPTIONS = ['سعودي'].concat(NATIONALITY_OPTIONS);

const EVIDENCE_OPTIONS = ['العقد', 'شهادة شهود', 'الفاتورة', 'ورقة إقرار', 'سند لأمر', 'إيصال استلام', 'محضر تسليم', 'الحوالة', 'رسالة نصية / واتساب', 'تسجيل صوتي', 'تقرير خبرة', 'تقدير الأضرار', 'تقرير نجم', 'تقرير المرور', 'عرض سعر', 'صور فوتوغرافية', 'مقاطع فيديو', 'مستند رسمي آخر'];

// حدّ الدعاوى اليسيرة التي تكون أحكامها نهائية غير قابلة للاستئناف.
// ⚠️ الحد يُحدَّد بقرار من المجلس الأعلى للقضاء وقابل للتعديل — يجب التحقق منه قبل الاعتماد عليه.
const YASEERA_CLAIM_LIMIT = 50000;

// ==================== الوصول لمكتبة النماذج (data/templates.js) ====================
// نصوص الإفهام والتدقيق الوجوبي مصدرها الوحيد مكتبة النماذج، فلا تُكرَّر هنا.

function templatesSource() {
    try {
        if (typeof templatesData !== 'undefined' && templatesData) return templatesData;
    } catch (e) { /* لم تُحمَّل بعد */ }
    if (typeof require !== 'undefined') {
        try { return require('../data/templates.js'); } catch (e) { /* غير متاحة */ }
    }
    return null;
}

// جلب نص نموذج من المكتبة بتصنيفه وكلمته المفتاحية
function findTemplate(category, keyword) {
    const src = templatesSource();
    if (!src || !Array.isArray(src[category])) return null;
    const hit = src[category].find(t => t.keyword === keyword);
    return hit ? hit.content : null;
}

// قائمة نماذج تصنيف كامل (تستهلكها واجهة مرحلة الأسباب والحكم)
function templatesOfCategory(category) {
    const src = templatesSource();
    return (src && Array.isArray(src[category])) ? src[category] : [];
}

// ==================== بحث النماذج (بطاقة البحث في نماذج التسبيب) ====================
// نصوص القوالب تُحدَّث من ملفات Word وتُطبَّع بـ tools/normalize_arabic.py، فتحمل تشكيلًا
// وهمزات وعلامات اقتباس تُفشِل المطابقة الحرفية. لذلك يُطبَّع الطرفان قبل المقارنة.
const TASHKEEL_RE = /[\u064B-\u065F\u0670\u0640]/g;

function normalizeArabicSearch(text) {
    return convertArabicDigits(String(text == null ? '' : text))
        .replace(TASHKEEL_RE, '')
        .replace(/[\u0623\u0625\u0622\u0671]/g, 'ا')
        .replace(/[\u0649]/g, 'ي')
        .replace(/[\u0624]/g, 'و')
        .replace(/[\u0626]/g, 'ي')
        .replace(/[\u0629]/g, 'ه')
        .replace(/[^\u0621-\u064a0-9a-zA-Z]+/g, ' ')
        .replace(/\s+/g, ' ')
        .trim()
        .toLowerCase();
}

// مطلع نماذج التسبيب المنسّقة في مقدمة تصنيف (أسباب الحكم) — علامة تعرُّف على المجموعة
// (يُقارَن مطبَّعًا، فلا يضره تشكيل النص ولا اختلاف الألف المقصورة)
const CURATED_REASON_OPENER = 'فبيان الدعوى';
// احتياطي إن تغيّر مطلع النماذج المنسّقة مستقبلًا
const CURATED_REASONS_FALLBACK = 31;

// عدد النماذج المنسّقة = الطول المتصل في مقدمة التصنيف الذي يبدأ بمطلعها الموحّد
function curatedReasonsCount(list) {
    const arr = Array.isArray(list) ? list : templatesOfCategory('أسباب الحكم');
    const opener = normalizeArabicSearch(CURATED_REASON_OPENER);
    let count = 0;
    for (let i = 0; i < arr.length; i++) {
        // يكفي مطلع النص، فلا حاجة لتطبيع النموذج كاملاً
        const head = String((arr[i] && arr[i].content) || '').slice(0, opener.length * 3);
        if (!normalizeArabicSearch(head).startsWith(opener)) break;
        count++;
    }
    return count || Math.min(CURATED_REASONS_FALLBACK, arr.length);
}

// ذاكرة تطبيع النماذج — يُستدعى البحث مع كل حرف يُكتب، والتصنيف يتجاوز 140 نموذجاً
const normalizedTemplateCache = typeof WeakMap !== 'undefined' ? new WeakMap() : null;

function normalizedTemplate(tpl) {
    if (!tpl || typeof tpl !== 'object') {
        return { keyword: '', all: normalizeArabicSearch(tpl) };
    }
    const cached = normalizedTemplateCache && normalizedTemplateCache.get(tpl);
    if (cached) return cached;
    const keyword = normalizeArabicSearch(tpl.keyword);
    const entry = { keyword, all: keyword + ' ' + normalizeArabicSearch(tpl.content) };
    if (normalizedTemplateCache) normalizedTemplateCache.set(tpl, entry);
    return entry;
}

// بحث بكلمات متعددة (وليس مطابقة حرفية): تُطابَق كل كلمة على حدة في الكلمة المفتاحية والنص،
// وتُقدَّم النماذج التي طابقت كلمتُها المفتاحية على ما طابق متنُها فقط.
function searchTemplates(list, query) {
    const arr = Array.isArray(list) ? list : [];
    const terms = normalizeArabicSearch(query).split(' ').filter(Boolean);
    if (terms.length === 0) return arr.slice();
    const hits = [];
    arr.forEach(t => {
        const norm = normalizedTemplate(t);
        if (!terms.every(term => norm.all.includes(term))) return;
        hits.push({ tpl: t, rank: terms.every(term => norm.keyword.includes(term)) ? 0 : 1 });
    });
    return hits.sort((a, b) => a.rank - b.rank).map(h => h.tpl);
}

// ==================== تحويل الوقت إلى كتابة عربية ====================

const hourWords = { 1: 'الواحدة', 2: 'الثانية', 3: 'الثالثة', 4: 'الرابعة', 5: 'الخامسة', 6: 'السادسة', 7: 'السابعة', 8: 'الثامنة', 9: 'التاسعة', 10: 'العاشرة', 11: 'الحادية عشرة', 12: 'الثانية عشرة' };
const minuteWords1to10 = { 1: 'دقيقة واحدة', 2: 'دقيقتين', 3: 'ثلاث دقائق', 4: 'أربع دقائق', 5: 'خمس دقائق', 6: 'ست دقائق', 7: 'سبع دقائق', 8: 'ثماني دقائق', 9: 'تسع دقائق', 10: 'عشر دقائق' };
const minuteWordsTeens = { 11: 'إحدى عشرة', 12: 'اثنتي عشرة', 13: 'ثلاث عشرة', 14: 'أربع عشرة', 16: 'ست عشرة', 17: 'سبع عشرة', 18: 'ثماني عشرة', 19: 'تسع عشرة' };
const minuteWordsTens = { 20: 'عشرين', 30: 'ثلاثين', 40: 'أربعين', 50: 'خمسين' };
const unitsFeminine = { 1: 'إحدى', 2: 'اثنتين', 3: 'ثلاث', 4: 'أربع', 5: 'خمس', 6: 'ست', 7: 'سبع', 8: 'ثماني', 9: 'تسع' };

function minutesPhrase(min) {
    if (min >= 1 && min <= 10) return 'و' + minuteWords1to10[min];
    if (min >= 11 && min <= 19) return 'و' + minuteWordsTeens[min] + ' دقيقة';
    if (min % 10 === 0) return 'و' + minuteWordsTens[min] + ' دقيقة';
    const tensPart = Math.floor(min / 10) * 10;
    const unitPart = min % 10;
    return 'و' + unitsFeminine[unitPart] + ' و' + minuteWordsTens[tensPart] + ' دقيقة';
}

function buildTimeArabic(hour, minute, period) {
    const periodWord = period === 'ص' ? 'صباحًا' : 'مساءً';
    if (minute === 45) {
        const nextHour = hour === 12 ? 1 : hour + 1;
        return `${hourWords[nextHour]} إلا ربعًا ${periodWord}`;
    }
    let phrase = hourWords[hour];
    if (minute === 15) phrase += ' والربع';
    else if (minute === 30) phrase += ' والنصف';
    else if (minute > 0) phrase += ' ' + minutesPhrase(minute);
    return phrase + ' ' + periodWord;
}

// إضافة دقائق لوقت (بنظام 12 ساعة) مع قلب الفترة ص/م عند العبور
function addMinutesToTime(hour, minute, period, addMin) {
    let total = (hour % 12) * 60 + minute + addMin;
    const flips = Math.floor(total / (12 * 60));
    total = ((total % (12 * 60)) + (12 * 60)) % (12 * 60);
    let newHour = Math.floor(total / 60);
    const newMinute = total % 60;
    if (newHour === 0) newHour = 12;
    let newPeriod = period;
    if (flips % 2 !== 0) newPeriod = period === 'ص' ? 'م' : 'ص';
    return { hour: newHour, minute: newMinute, period: newPeriod };
}

// ساعات انعقاد الجلسات المتاحة لكل فترة
function validHoursForPeriod(period) {
    return period === 'ص' ? [8, 9, 10, 11] : [12, 1, 2, 3];
}

// تحويل الأرقام العربية المشرقية إلى أرقام لاتينية
function convertArabicDigits(str) {
    const arabicDigits = '٠١٢٣٤٥٦٧٨٩';
    return String(str).replace(/[٠-٩]/g, d => String(arabicDigits.indexOf(d)));
}

// ==================== المطابقة النحوية والألقاب ====================

const ordinalWordsMasc = { 1: 'الأول', 2: 'الثاني', 3: 'الثالث', 4: 'الرابع', 5: 'الخامس', 6: 'السادس', 7: 'السابع', 8: 'الثامن', 9: 'التاسع', 10: 'العاشر' };
const ordinalWordsFem = { 1: 'الأولى', 2: 'الثانية', 3: 'الثالثة', 4: 'الرابعة', 5: 'الخامسة', 6: 'السادسة', 7: 'السابعة', 8: 'الثامنة', 9: 'التاسعة', 10: 'العاشرة' };

function ordinalWord(n, gender) {
    const table = gender === 'ف' ? ordinalWordsFem : ordinalWordsMasc;
    return table[n] || `رقم ${n}`;
}

function partyLabel(role, gender) {
    if (role === 'plaintiff') return gender === 'م' ? 'المدعي' : 'المدعية';
    return gender === 'م' ? 'المدعى عليه' : 'المدعى عليها';
}

// لقب الطرف عند تعدد الأطراف: «المدعي الأول فلان»
function multiPartyLabel(role, gender, ordinalNum, name) {
    const base = partyLabel(role, gender);
    const ordinal = ordinalWord(ordinalNum, gender);
    const namePart = (name && name.trim()) ? ` ${name.trim()}` : '';
    return `${base} ${ordinal}${namePart}`;
}

// «وكيله / وكيلها / وكيلته / وكيلتها» حسب جنس الوكيل وجنس الأصيل
function agentPossessive(agentGender, partySuffix) {
    if (agentGender === 'ف') {
        return partySuffix === 'ه' ? 'وكيلته' : 'وكيلتها';
    }
    return partySuffix === 'ه' ? 'وكيله' : 'وكيلها';
}

// ==================== جداول درجات القرابة (لوكيل غير المحامي) ====================
// المادة (1/7) من اللائحة التنفيذية لنظام المرافعات الشرعية
// المعدَّلة بالقرار الوزاري رقم (2044) وتاريخ 1443/8/4هـ

const kinshipDegree = {
    'الأب': 1, 'الأم': 1, 'الابن': 1, 'البنت': 1,
    'الأخ': 2, 'الأخت': 2, 'الجد': 2, 'الجدة': 2, 'ابن الابن': 2, 'بنت الابن': 2, 'ابن البنت': 2, 'بنت البنت': 2,
    'العم': 3, 'الخال': 3, 'العمة': 3, 'الخالة': 3, 'ابن الأخ': 3, 'بنت الأخ': 3, 'ابن الأخت': 3, 'بنت الأخت': 3,
    'ابن العم': 4, 'بنت العم': 4, 'ابن العمة': 4, 'بنت العمة': 4, 'ابن الخال': 4, 'بنت الخال': 4, 'ابن الخالة': 4, 'بنت الخالة': 4
};

const degreeOrdinal = { 1: 'الأولى', 2: 'الثانية', 3: 'الثالثة', 4: 'الرابعة' };

// قرابة الأصهار: نفس الدرجات من جهة الزوج أو الزوجة
const asharTemplate = [
    { degree: 1, items: ['أب {S}', 'أم {S}', 'ابن {S} (الربيب)', 'بنت {S} (الربيبة)'] },
    { degree: 2, items: ['أخو {S}', 'أخت {S}', 'جد {S}', 'جدة {S}', 'ابن ابن {S}', 'بنت ابن {S}', 'ابن بنت {S}', 'بنت بنت {S}'] },
    { degree: 3, items: ['عم {S}', 'عمة {S}', 'خال {S}', 'خالة {S}', 'ابن أخ {S}', 'بنت أخ {S}', 'ابن أخت {S}', 'بنت أخت {S}'] },
    { degree: 4, items: ['ابن عم {S}', 'بنت عم {S}', 'ابن عمة {S}', 'بنت عمة {S}', 'ابن خال {S}', 'بنت خال {S}', 'ابن خالة {S}', 'بنت خالة {S}'] }
];

const asharDegree = {};
['الزوجة', 'الزوج'].forEach(spouse => {
    asharTemplate.forEach(group => {
        group.items.forEach(itemTpl => {
            asharDegree[itemTpl.replace('{S}', spouse)] = group.degree;
        });
    });
});

// صفات الأطراف المتاحة — الوقف يحمل خصائص الشركة نفسها ويفترق عنها في نوع الإفهام
const ENTITY_TYPES = { INDIVIDUAL: 'فرد', COMPANY: 'شركة', WAQF: 'وقف' };

// ==================== حالات ابتدائية ====================

function freshPartyState() {
    return {
        gender: 'م', name: '', attendance: 'أصالة', repType: 'وكيل', repIsLawyer: 'نعم', agentGender: 'م',
        agentName: '', agentId: '', wakalaNum: '', licenseNum: '', repNum: '', repName: '', occurrence: 1,
        tabligh: '', notifyStatus: 'تبلغ',
        kinshipType: 'نسب', kinship: '', asharDirection: 'الزوجة', kinshipAshar: '',
        hasAccompanyingAgent: false, specialCase: 'none',
        oathAbsence: 'لا', oathTablighNum: '',
        nationalityType: 'سعودي', saudiId: '', foreignNationality: '', iqamaNum: '',
        oathPerformanceSession: 'لا', oathAnswerText: '',
        entityType: 'فرد', crNum: ''
    };
}

function freshExtraParty() {
    return {
        name: '', gender: 'م', attendance: 'أصالة', repName: '', repNum: '',
        // بطاقة الهوية — كالطرف الأول تمامًا (تظهر عند الحضور أصالة)
        nationalityType: 'سعودي', saudiId: '', foreignNationality: '', iqamaNum: ''
    };
}

function freshWitness() {
    return {
        name: '', nationality: '', idNum: '', age: '', job: '', residence: '',
        // صلة الشاهد بكل خصم على حدة — المادة (71) من نظام الإثبات
        relationPlaintiff: '', relationDefendant: '', interest: '', testimony: '',
        // رقم الجوال لا يُدرج في نص الضبط، وإنما ليتمكن القاضي من الوصول للشاهد عند الحاجة
        phone: ''
    };
}

// بينة أحد الخصمين (يُستعمل للمدعى عليه، وللمدعي عبر plaintiffEvidenceBlock)
function freshEvidenceBlock() {
    return {
        choice: 'none', items: [], otherDocumentText: '',
        witnesses: [],
        tazkiya: 'none',        // 'none' لا تزكية | 'delay' طلب الإمهال | 'presented' حضر المعدِّلون
        tazkiyaNames: '',
        objection: 'لا',        // مطعن الخصم في الشهود أو في الشهادة
        objectionText: ''
    };
}

function freshClaimState() {
    return {
        text: '',
        // نوع الدعوى: الطلب المالي وحده تُشتق من قيمته صفة الحكم (يسير أو قابل للاستئناف)
        claimType: 'مالي',          // 'مالي' | 'غير مالي'
        claimValue: '',             // قيمة المطالبة — يُشتق منها نوع الإفهام
        // جواب المدعى عليه — يُعرض عليه قبل تكليف المدعي بالبينة
        defendantStance: 'إنكار',   // 'إقرار' | 'إنكار' | 'دفع شكلي'
        defendantResponseText: '',
        formalPleaText: '',
        plaintiffReplyText: '',
        answeredOnMerits: 'نعم',    // مع الدفع الشكلي: هل أجاب في الموضوع أيضًا؟
        // بينة المدعي
        evidenceChoice: 'none', evidenceItems: [], otherDocumentText: '',
        hasMoreEvidence: 'نعم', moreEvidenceText: '',
        witnesses: [],
        tazkiya: 'none', tazkiyaNames: '',
        witnessObjection: 'لا', witnessObjectionText: '',
        requestOath: false, declineOath: false,
        // دفوع المدعى عليه وبينته
        askDefendantEvidence: 'لا',
        defendantPleasText: '',
        defendantEvidence: freshEvidenceBlock()
    };
}

// الأسباب والحكم — مرحلة مستقلة تُستهلك فيها نماذج التسبيب من مكتبة النماذج
function freshRulingState() {
    return {
        pronounce: 'لا',            // هل يُنطق بالحكم في هذه الجلسة؟
        reasonsText: '',
        rulingText: '',
        presence: 'حضوري',          // 'حضوري' | 'غيابي'
        noticeKind: 'auto',         // 'auto' | 'final' | 'appealable' | 'sulh'
        // إجراء التدقيق الوجوبي — لا يُسأل عنه إلا مع إفهام (واجب التدقيق)، وهو صيغته لا فقرة زائدة عليه
        mandatoryReview: 'نعم'      // 'نعم' رفع الملف بعد مدة الاعتراض | 'إعادة' إعادة القضية للاستئناف
    };
}

// إجراءات الجلسة التالية (سماع البينات فيها لا في الجلسة التحضيرية وحدها)
function freshFollowUpState() {
    return { plaintiffEvidence: false, defendantEvidence: false };
}

// الحالة الكاملة للمحضر (تُمرَّر لكل دوال التوليد)
function freshMinutesState() {
    return {
        sessionType: null,           // 'new' جلسة تحضيرية | 'previous' منظورة سابقًا
        sameJudge: 'نعم',
        // إدراج بيانات الطرفين — الاسم والهوية (أو الإقامة أو السجل التجاري) — في نص الضبط.
        // مُغلق افتراضًا: جدول نظام تقاضي أعلى صفحة الضبط والصك يحمل اسم الطرف ورقم هويته،
        // فإثباتها في المتن تكرار، ويكفي وصفه بـ«المدعي/المدعية/المدعى عليه» مع ترقيمه عند التعدد.
        // ولا يشمل هذا المفتاح بيانات الوكيل (رقم الوكالة والتحقق منها بالمادة (51/3) ورخصة
        // المحاماة ودرجة القرابة) فتلك لا يكتبها الجدول وإثباتها في الضبط مطلوب نظامًا.
        includePartyDataInText: false,
        opening: { judge: '', court: '', mode: SESSION_MODES.VIDEO },
        // وقت الإغلاق يُدخله القاضي في آخر الضبط (لا يُشتق من وقت الافتتاح)
        closing: { hour: 8, minute: 30, period: 'ص' },
        plaintiff: freshPartyState(),
        defendant: freshPartyState(),
        extraPlaintiffs: [],
        extraDefendants: [],
        claim: freshClaimState(),
        followUp: freshFollowUpState(),
        ruling: freshRulingState()
    };
}

// ==================== بناة الجُمل ====================

function orDots(value) {
    const v = (value == null ? '' : String(value)).trim();
    return v || MINUTES_PLACEHOLDER;
}

// فقرة افتتاح الجلسة
// وقت الافتتاح مثبت أصلاً في صفحة القضية بناجز، فإعادته في متن الضبط ازدواج،
// ولذا اقتُصر على وقت الإغلاق في آخر الضبط.
function buildOpening(opening) {
    const judge = orDots(opening.judge);
    const court = orDots(opening.court);
    const mode = sessionMode(opening);
    let modePart = ` ${VIDEO_CALL_BRIEF}`;
    if (mode === SESSION_MODES.VIDEO_FULL) modePart = ` ${VIDEO_CALL_PHRASE}`;
    else if (mode === SESSION_MODES.IN_PERSON) modePart = '';
    return `فلديّ أنا ${judge} القاضي في المحكمة ${court} افتتحتُ الجلسة${modePart}،`;
}

// عبارة التحقق من الوكالة — المادة (51/3) من نظام المرافعات الشرعية
function buildAgencyVerificationClause(agentGender) {
    const agentSuffix = agentGender === 'ف' ? 'ها' : 'ه';
    return `، وقد جرى من الدَّائرة التحقق منها، وأنها تخوّل${agentSuffix} الإجراء المطلوب، استناداً للمادَّة (51/3) من نظام المرافعات الشرعية`;
}

// عبارة صلة قرابة الوكيل غير المحامي (أو عدم أحقيته إن كان خارج الدرجات الأربع)
function buildKinshipPhrase(s, name) {
    const isAshar = s.kinshipType === 'أصهار';
    const value = isAshar ? s.kinshipAshar : s.kinship;
    if (!value) return '';
    const agentSuffix = s.agentGender === 'ف' ? 'ها' : 'ه';

    if (value === 'أخرى') {
        const isFem = s.agentGender === 'ف';
        const agentNoun = isFem ? 'وكيلة' : 'وكيل';
        const heShe = isFem ? 'أنها' : 'أنه';
        const notRelated = isFem ? 'ليست قريبةً' : 'ليس قريبًا';
        const lahu = isFem ? 'لها' : 'له';
        return `، وعليه بعد التحقق من وكالة ${agentNoun} ${name} تبيّن ${heShe} ${notRelated} بأحد الدرجات الأربع المقررة بموجب المادة (1/7) من اللائحة التنفيذية لنظام المرافعات الشرعية المعدَّلة بالقرار الوزاري رقم (2044) وتاريخ 1443/8/4هـ، وعليه فلا يحق ${lahu} الاستمرار في الإجابة`;
    }

    const typeNote = isAshar ? ' (أصهار)' : '';
    return `، وتربط${agentSuffix} ب${name} صلة قرابة: ${value}${typeNote}`;
}

// عبارة الوكيل المرافق للطرف الحاضر أصالة
function buildAccompanyingAgentClause(s, name, suffix) {
    const wakalaNum = orDots(s.wakalaNum);
    const agentName = s.agentName.trim();
    const agentId = orDots(s.agentId);
    const namePart = agentName ? ` ${agentName}` : '';
    const agentVerb = s.agentGender === 'ف' ? 'حضرت' : 'حضر';
    const agentPoss = agentPossessive(s.agentGender, suffix);
    const lawyerWord = s.agentGender === 'ف' ? 'المحامية' : 'المحامي';
    const withHim = suffix === 'ه' ? 'معه' : 'معها';

    if (s.repIsLawyer === 'نعم') {
        const lic = orDots(s.licenseNum);
        return `، و${agentVerb} ${withHim} ${agentPoss} ${lawyerWord}${namePart}، بموجب الهوية رقم (${agentId})، بموجب الوكالة الشرعية رقم (${wakalaNum})${buildAgencyVerificationClause(s.agentGender)} ورخصة مزاولة المحاماة رقم (${lic})`;
    }
    return `، و${agentVerb} ${withHim} ${agentPoss}${namePart}، بموجب الهوية رقم (${agentId})، بموجب الوكالة الشرعية رقم (${wakalaNum})${buildAgencyVerificationClause(s.agentGender)}${buildKinshipPhrase(s, name)}`;
}

// صفة الطرف: الشركة والوقف كلاهما شخصية اعتبارية لا تحضر أصالةً وإنما بوكيل أو ممثل نظامي
function isCorporateEntity(s) {
    return !!s && (s.entityType === ENTITY_TYPES.COMPANY || s.entityType === ENTITY_TYPES.WAQF);
}

// وثيقة الشخصية الاعتبارية: السجل التجاري للشركة وصك الوقفية للوقف
function corporateDocLabel(s) {
    return (s && s.entityType === ENTITY_TYPES.WAQF) ? 'صك الوقفية' : 'السجل التجاري';
}

// إثبات هوية الطرف الحاضر (هوية وطنية / إقامة / سجل تجاري أو صك وقفية)
// includePartyData=false يُسقط العبارة من المخرج (بيانات الطرف مثبتة في جدول تقاضي)
function buildIdentityClause(s, includePartyData = true) {
    if (!includePartyData) return '';
    if (isCorporateEntity(s)) {
        return `، بموجب ${corporateDocLabel(s)} رقم (${orDots(s.crNum)})`;
    }
    if (s.nationalityType === 'سعودي') {
        return `، بموجب الهوية الوطنية رقم (${orDots(s.saudiId)})`;
    }
    return `، ${orDots(s.foreignNationality)} الجنسية، بموجب الإقامة النظامية رقم (${orDots(s.iqamaNum)})`;
}

// فقرة حضور/غياب الطرف الأساسي (المدعي أو المدعى عليه)
function buildPartyClause(state, role) {
    const s = state[role];
    const includePartyData = state.includePartyDataInText === true;
    const extraList = role === 'plaintiff' ? state.extraPlaintiffs : state.extraDefendants;
    const hasMultiple = extraList.length > 0;
    // عند إغلاق المفتاح يُوصف الطرف بلقبه وحده: «المدعي» أو «المدعي الأول» عند التعدد
    const givenName = includePartyData ? s.name : '';
    const name = hasMultiple
        ? multiPartyLabel(role, s.gender, 1, givenName)
        : (givenName.trim() ? `${partyLabel(role, s.gender)} ${givenName.trim()}` : partyLabel(role, s.gender));
    const suffix = s.gender === 'م' ? 'ه' : 'ها';

    if (s.attendance === 'أصالة') {
        const verb = s.gender === 'م' ? 'حضر' : 'حضرت';
        let clause = `${verb} ${name} أصالة${buildIdentityClause(s, includePartyData)}`;
        if (s.hasAccompanyingAgent) {
            clause += buildAccompanyingAgentClause(s, name, suffix);
        }
        return clause;
    }

    if (s.attendance === 'تمثيل') {
        if (s.repType === 'وكيل' || s.repType === 'وكيل شركة') {
            const wakalaNum = orDots(s.wakalaNum);
            const agentName = s.agentName.trim();
            const agentId = orDots(s.agentId);
            const namePart = agentName ? ` ${agentName}` : '';
            const agentVerb = s.agentGender === 'ف' ? 'حضرت' : 'حضر';
            const agentPoss = agentPossessive(s.agentGender, suffix);
            const lawyerWord = s.agentGender === 'ف' ? 'المحامية' : 'المحامي';
            if (s.repIsLawyer === 'نعم') {
                const lic = orDots(s.licenseNum);
                return `${agentVerb} عن ${name} ${agentPoss} ${lawyerWord}${namePart}، بموجب الهوية رقم (${agentId})، بموجب الوكالة الشرعية رقم (${wakalaNum})${buildAgencyVerificationClause(s.agentGender)} ورخصة مزاولة المحاماة رقم (${lic})`;
            }
            const kinshipPart = s.repType === 'وكيل شركة' ? '' : buildKinshipPhrase(s, name);
            return `${agentVerb} عن ${name} ${agentPoss}${namePart}، بموجب الهوية رقم (${agentId})، بموجب الوكالة الشرعية رقم (${wakalaNum})${buildAgencyVerificationClause(s.agentGender)}${kinshipPart}`;
        }
        // ممثل نظامي
        const repCapacity = orDots(s.repNum);
        const repName = s.repName.trim();
        const repNamePart = repName ? ` ${repName}` : '';
        return `حضر عن ${name} ممثل${suffix} النظامي${repNamePart}، بصفة: ${repCapacity}`;
    }

    const verb = s.gender === 'م' ? 'لم يحضر' : 'لم تحضر';
    return `${verb} ${name} رغم تبليغ${suffix} بالجلسة وفق مهمّة التبليغ رقم (${orDots(s.tabligh)})`;
}

// فقرة طرف إضافي (عند تعدد المدعين أو المدعى عليهم)
function buildExtraClause(role, idx, p, includePartyData = true) {
    const name = multiPartyLabel(role, p.gender, idx + 2, includePartyData ? p.name : '');
    const suffix = p.gender === 'م' ? 'ه' : 'ها';

    if (p.attendance === 'أصالة') {
        const verb = p.gender === 'م' ? 'حضر' : 'حضرت';
        return `${verb} ${name} أصالة${buildIdentityClause(p, includePartyData)}`;
    }
    if (p.attendance === 'تمثيل') {
        const repName = p.repName.trim();
        const namePart = repName ? ` ${repName}` : '';
        return `حضر عن ${name} وكيل${suffix}${namePart} بموجب الوكالة الشرعية رقم (${orDots(p.repNum)})`;
    }
    const verb = p.gender === 'م' ? 'لم يحضر' : 'لم تحضر';
    return `${verb} ${name}`;
}

// غياب طرف لم يتبلّغ: رفع الجلسة لإعادة التبليغ
function buildNotNotifiedClause(role, s) {
    const name = partyLabel(role, s.gender);
    const suffix = s.gender === 'م' ? 'ه' : 'ها';
    const verbAttend = s.gender === 'م' ? 'لم يحضر' : 'لم تحضر';
    const verbNotify = s.gender === 'م' ? 'لم يتبلّغ' : 'لم تتبلّغ';
    return `${verbAttend} ${name}، و${verbNotify} بالجلسة، وعليه رُفعت الجلسة لإعادة تبليغ${suffix} بحسب حال${suffix}`;
}

// الحالات الاستثنائية للمدعي (حضور بالنظام دون المرئي / المغادرة أثناء الجلسة)
function buildSpecialCaseText(state, role) {
    const s = state[role];
    const pName = partyLabel(role, s.gender);
    const isRepresented = s.attendance === 'تمثيل';
    const subjectGender = isRepresented ? s.agentGender : s.gender;
    const isFem = subjectGender === 'ف';
    const subjectName = isRepresented ? `وكيل${s.agentGender === 'ف' ? 'ة' : ''} ${pName}` : pName;

    if (s.specialCase === 'systemNoVideo') {
        const verbHadara = isFem ? 'تحضر' : 'يحضر';
        const heShe = isFem ? 'أنها' : 'أنه';
        const possessive = isFem ? 'حضورها' : 'حضوره';
        const howAttend = isInPersonSession(state) ? 'الجلسة بمقر المحكمة' : `الجلسة ${VIDEO_CALL_BRIEF}`;
        return `حيث تبيّن حضور ${subjectName} في النظام الإلكتروني، إلا ${heShe} لم ${verbHadara} ${howAttend}، ولم تتمكن الدائرة من شطب الجلسة بسبب ثبوت ${possessive} في النظام، وعليه جرى إثبات ذلك في محضر الجلسة`;
    }
    if (s.specialCase === 'leftDuringSession') {
        const verbGhadara = isFem ? 'غادرت' : 'غادر';
        const verbAada = isFem ? 'تعد' : 'يعد';
        const heShe = isFem ? 'أنها' : 'أنه';
        const possessiveIntizar = isFem ? 'انتظارها' : 'انتظاره';
        return `في أثناء انعقاد الجلسة ${verbGhadara} ${subjectName}، فجرى ${possessiveIntizar} حتى انتهاء الوقت المحدد للجلسة، إلا ${heShe} لم ${verbAada} للحضور، وعليه جرى إثبات ذلك في محضر الجلسة`;
    }
    return '';
}

// محضر شطب الدعوى عند غياب المدعي المتبلّغ — المادة (55) من نظام المرافعات الشرعية
function buildShatbText(state) {
    const opening = buildOpening(state.opening);
    const s = state.plaintiff;
    const name = partyLabel('plaintiff', s.gender);
    const suffix = s.gender === 'م' ? 'ه' : 'ها';
    const verbHadara = s.gender === 'م' ? 'يحضر' : 'تحضر';
    const verbTaqaddama = s.gender === 'م' ? 'يتقدم' : 'تتقدم';
    const lahuOrLaha = s.gender === 'م' ? 'له' : 'لها';
    const tabligh = orDots(s.tabligh);
    const llName = 'ل' + name.slice(1);

    // في الجلسة الحضورية تُحذف الإشارة لرابط غرفة الاتصال المرئي، وتُعتمد صيغة التبليغ
    // عبر الوسائل الإلكترونية كما في مكتبة قوالب ضبط الجلسات (تصنيف «الغياب» و«الشطب»)
    const notifyClause = isInPersonSession(state)
        ? `وقد جرى تبليغ ${name} بعقد هذه الجلسة - ${IN_PERSON_SHORT} - عبر الوسائل الإلكترونية وفق مهمّة التبليغ رقم ( ${tabligh} )، ولم يكن ${lahuOrLaha} وجود حتَّى انتهاءِ وقت الجلسة المخصَّص لنظر القضية في هذا اليوم،`
        : `وقد جرى تبليغ ${name} بعقد هذه الجلسة - ${VIDEO_CALL_SHORT} - وفق مهمّة التبليغ رقم ( ${tabligh} ) كما جرى بعث وتبليغ ${name} برابط الدخول لغرفة الاتصال المرئي، ولم يكن ${lahuOrLaha} وجود حتَّى انتهاءِ وقت الجلسة المُحدَّد،`;

    const bodyCore = `ولم ${verbHadara} فيها ${name}، ولا مَنْ يُمثِّل${suffix}، ولم ${verbTaqaddama} بعذر تقبله الدَّائرة، ${notifyClause}`;

    if (s.occurrence === 1) {
        const bodyMiddle = `${bodyCore} واستناداً لما نصت عليه الفقرة (2) من البند ( رابعاً ) من الدليل الإجرائي لخدمة التقاضي الالكتروني ونصّها: "يكون النطق بالحكم وشطب الدَّعوى، والحكم بالنكول من خلال ( الجلسة المرئية ) أو ( الجلسة الحضورية ) وتسلم الأحكام للأطراف الكترونياً" `;
        const closing = `لذا قرَّرت الدَّائرة شطب الدَّعوى للمرَّة الأُولى و${llName} المطالبة بإعادة استمرار السير في نظر الدعوى خلال (60) يوماً، عملًا بالمادَّة (55) من نظام المرافعات الشَّرعيَّة`;
        return `${opening} ${bodyMiddle}${closing}.`;
    }

    // الغياب الثاني بعد شطب سابق وطلب السير: اعتبار الدعوى كأن لم تكن
    const verbTalab = s.gender === 'م' ? 'طلب' : 'طلبت';
    const verbHadara2 = s.gender === 'م' ? 'لم يحضر' : 'لم تحضر';

    const reasons = `الأسباب:\nفبناء على المادة الخامسة والخمسين من نظام المرافعات الشرعية والتي نصها : ((إذا لم يحضر المدعي أي جلسة من جلسات الدعوى ولم يتقدم بعذر تقبله المحكمة فلها أن تقرر شطبها فإذا انقضت (ستون) يوماً ولم يطلب المدعي السير فيها بعد شطبها أو لم يحضر بعد السير فيها عُدًّت كأن لم تكن وإذا طلب المدعي بعد ذلك السير في الدعوى حكمت المحكمة من تلقاء نفسها باعتبار الدعوى كأن لم تكن)) ولأنه سبق أن شطبت هذه القضية ثم ${verbTalab} ${name} السير فيها ثم ${verbHadara2} بعد السير فيها`;

    const ruling = `الحكم:\nحكمت الدائرة باعتبار الدعوى كأن لم تكن، وقررت الدائرة احتساب مدة ثلاثين يومًا للاعتراض عن طريق نظام ناجز من تاريخ صدور الصك، فإن مضت المدة بدون اعتراض سقط الحق في الاعتراض واكتسب الحكم القطعية، وبالله التوفيق.`;

    return `${opening} ${bodyCore}\n\n${reasons}\n\n${ruling}`;
}

// غياب المدعى عليه عن جلسة أداء اليمين رغم تبلُّغه — نكول
function buildOathAbsenceDefendantText(state) {
    const d = state.defendant;
    const p = state.plaintiff;
    const dName = partyLabel('defendant', d.gender);
    const dSuffix = d.gender === 'م' ? 'ه' : 'ها';
    const isRepresented = p.attendance === 'تمثيل';
    const speakerGender = isRepresented ? p.agentGender : p.gender;
    const isFem = speakerGender === 'ف';
    const pName = isRepresented ? `وكيل${p.agentGender === 'ف' ? 'ة' : ''} ${partyLabel('plaintiff', p.gender)}` : partyLabel('plaintiff', p.gender);
    const hasIt = isFem ? 'لديها' : 'لديه';
    const addVerb = isFem ? 'تضيفه' : 'يضيفه';
    const pVerb = isFem ? 'فقررت' : 'فقرر';
    const pVerbEnd = isFem ? 'قررت' : 'قرر';
    const tablighNum = orDots(d.oathTablighNum);

    return `ولم يحضر ${dName} المدون اسم${dSuffix} أعلاه، ولا مَنْ يُمثِّل${dSuffix}، رغم تبلُّغ${dSuffix} لـ( شخص${dSuffix} ) بموعد هذه الجلسة المخصصة لأداء اليمين على نفي الدعوى بناءً على إفادة التبليغ الإلكتروني بالنظام رقم ( ${tablighNum} ) والمتضمنة: "تم التبليغ". ولما نصت عليه الفقرة (2) من المادة (الثالثة بعد المائة) من نظام الإثبات بأن من تخلف عن الحضور لأداء اليمين بغير عذر عُدّ ناكلاً، ثم جرى من الدائرة سؤال ${pName} هل ${hasIt} ما ${addVerb}؟ ${pVerb}: ليس لدي سوى ما قدمت هكذا ${pVerbEnd}، واستناداً للمادة (69) والمادة (159) من نظام المرافعات الشرعية فقد قررت الدائرة قفل باب المرافعة للنطق بالحكم في هذه الجلسة`;
}

// جلسة أداء اليمين: سؤال المدعى عليه الحاضر عن استعداده
function buildOathPerformanceText(state) {
    const s = state.defendant;
    const isFem = s.gender === 'ف';
    const label = partyLabel('defendant', s.gender);
    const readyWord = isFem ? 'مستعدة' : 'مستعد';
    const youWord = isFem ? 'أنتِ' : 'أنت';
    const verb = isFem ? 'قررت' : 'قرر';
    const qa = isFem ? 'قائلة' : 'قائلاً';
    const answer = orDots(s.oathAnswerText);
    return `وبسؤال ${label} الحاضر${isFem ? 'ة' : ''} هل ${youWord} ${readyWord} لأداء اليمين؟ ف${verb} ${qa}: ${answer}. هكذا ${verb}`;
}

// جنس المتكلم عن الطرف (الوكيل إن كان ممثلاً بوكيل)
function speakerGenderOf(s) {
    return s.attendance === 'تمثيل' ? s.agentGender : s.gender;
}

// لقب المخاطب عن الطرف («وكيل المدعي» أو «المدعي»)
function addresseeLabelOf(role, s) {
    const g = speakerGenderOf(s);
    const pName = partyLabel(role, s.gender);
    return s.attendance === 'تمثيل' ? `وكيل${g === 'ف' ? 'ة' : ''} ${pName}` : pName;
}

// قفل باب المرافعة — المادتان (69) و(159) من نظام المرافعات الشرعية
function buildClosingArgumentText(state) {
    const pPresent = state.plaintiff.attendance !== 'لم يحضر';
    const dPresent = state.defendant.attendance !== 'لم يحضر';
    if (!pPresent && !dPresent) return '';

    if (pPresent && dPresent) {
        const bothFem = speakerGenderOf(state.plaintiff) === 'ف' && speakerGenderOf(state.defendant) === 'ف';
        const verbDecide = bothFem ? 'قررتا' : 'قررا';
        return `ثم جرى من الدائرة سؤال أطراف الدعوى هل لديكما ما تضيفانه؟ ف${verbDecide}: ليس لدينا سوى ما قدمنا. هكذا ${verbDecide}، واستناداً للمادة (69) والمادة (159) من نظام المرافعات الشرعية فقد قررت الدائرة قفل باب المرافعة للنطق بالحكم في هذه الجلسة`;
    }

    const role = pPresent ? 'plaintiff' : 'defendant';
    const s = state[role];
    const isFem = speakerGenderOf(s) === 'ف';
    const label = addresseeLabelOf(role, s);
    const hasIt = isFem ? 'لديها' : 'لديه';
    const addVerb = isFem ? 'تضيفه' : 'يضيفه';
    const decideVerb = isFem ? 'قررت' : 'قرر';
    return `ثم جرى من الدائرة سؤال ${label} هل ${hasIt} ما ${addVerb}؟ ف${decideVerb}: ليس لدي سوى ما قدمت. هكذا ${decideVerb}، واستناداً للمادة (69) والمادة (159) من نظام المرافعات الشرعية فقد قررت الدائرة قفل باب المرافعة للنطق بالحكم في هذه الجلسة`;
}

// مصادقة الأطراف على وقائع الجلسات السابقة
function buildPreviousSessionRatification(state) {
    const pPresent = state.plaintiff.attendance !== 'لم يحضر';
    const dPresent = state.defendant.attendance !== 'لم يحضر';

    if (pPresent && dPresent) {
        const bothFem = speakerGenderOf(state.plaintiff) === 'ف' && speakerGenderOf(state.defendant) === 'ف';
        const verb = bothFem ? 'قررتا' : 'قرروا';
        return ` وبعرض وقائع الجلسات السابقة على أطراف الدعوى ${verb} قائلين: نصادق على ما ورد في الجلسات السابقة. هكذا ${verb}.`;
    }

    const presentRole = pPresent ? 'plaintiff' : (dPresent ? 'defendant' : null);
    if (!presentRole) return '';

    const s = state[presentRole];
    const gender = speakerGenderOf(s);
    const name = addresseeLabelOf(presentRole, s);
    const verb = gender === 'م' ? 'قرر' : 'قررت';
    const adj = gender === 'م' ? 'قائلاً' : 'قائلة';
    return ` وبعرض وقائع الجلسات السابقة على ${name} ${verb} ${adj}: نصادق على ما ورد في الجلسات السابقة. هكذا ${verb}.`;
}

// إفهام المدعي بأثر طلب اليمين — المادة (99) من نظام الإثبات
function buildOathSentence(pGender, dName, pName) {
    const heHa = pGender === 'م' ? 'ه' : 'ها';
    const verbTalab = pGender === 'م' ? 'طلب' : 'طلبت';
    const minHu = pGender === 'م' ? 'منه' : 'منها';
    const lahu = pGender === 'م' ? 'له' : 'لها';
    return `وأفهمتُ ${pName} بأن${heHa} إذا ${verbTalab} يمين ${dName} فإن${heHa} تسقط بينات${heHa} ولن تُقبل ${minHu}، وأن اليمين حاسمة للنزاع ولا يجوز ${lahu} إثبات كذبها بعد أدائها استناداً للمادة (التاسعة والتسعين) من نظام الإثبات.`;
}

// كتلة طلب اليمين / الامتناع عنها (مع حالة غياب المدعى عليه)
function buildOathBlock(pGender, dName, addresseeName, requestOath, declineOath, oathStatus, isRepresented, moakkelPhrase) {
    if (declineOath) {
        const verb = pGender === 'م' ? 'قرر' : 'قررت';
        return `، وبسؤاله هل ترغب بتوجيه اليمين للمدعى عليه؟ ${verb}: لا أرغب. هكذا ${verb}.`;
    }
    if (!requestOath) return '.';
    let s;
    if (isRepresented) {
        const clientVerb = pGender === 'م' ? 'يطلب' : 'تطلب';
        s = `، و${clientVerb} ${moakkelPhrase} يمين ${dName}.`;
    } else {
        s = `، وأطلب يمين ${dName}.`;
    }
    s += ' ' + buildOathSentence(pGender, dName, addresseeName);
    if (oathStatus === 'غائب') {
        s += ' قررت الدائرة الاستجابة لطلبه بتوجيه اليمين الحاسمة للمدعى عليه الغائب، وتحديد جلسة قادمة لإبلاغه بالحضور لأدائها، وسيشعر بوجوب حضوره، وأنه إن لم يحضر في الموعد المحدد بغير عذر تقبله المحكمة، أو حضر وامتنع عن أدائها أو ردها؛ عُدَّ ناكلاً، وسيقضى عليه بنكوله، استناداً للفقرتين (1) و (2) من المادة (الثالثة بعد المائة) من نظام الإثبات';
    }
    return s;
}

// قسم سماع الشهود — المادتان (71) و(78) من نظام الإثبات
// ctx: { speakerGender, speakerSuffix, presenterLabel, presenterGender,
//        opposingLabel, opposingPresent, tazkiya, tazkiyaNames, objection, objectionText }
function buildWitnessSection(witnesses, ctx) {
    const decideShort = ctx.speakerGender === 'ف' ? 'قررت' : 'قرر';
    const saying = ctx.speakerGender === 'ف' ? 'قائلة' : 'قائلاً';
    const bringVerb = ctx.presenterGender === 'ف' ? 'أحضرت' : 'أحضر';

    let out = ` وبسؤال${ctx.speakerSuffix} هل من يشهد لك حاضر؟ ف${decideShort} ${saying}: نعم. هكذا ${decideShort}. استناداً إلى نظام الإثبات، وتحديداً المادتين (الحادية والسبعين) و(الثامنة والسبعين) منه، يجب على المحكمة قبل أداء الشهادة أن تطلب من الشاهد الإفصاح عن بياناته الشخصية وعلاقته بالدعوى وتدوين ذلك في الضبط، وقد جرى سماع كل شاهد على انفراد.`;

    witnesses.forEach((w, i) => {
        const ordinal = ordinalWordsMasc[i + 1] || `رقم ${i + 1}`;
        // بيانات الشاهد تُثبت كاملة في الضبط (المادة 71 من نظام الإثبات)، ورقم الجوال وحده لا يُدرج
        out += ` ثم ${bringVerb} ${ctx.presenterLabel} للشهادة الشاهد ${ordinal}، وبسؤاله عن بياناته الشخصية وعلاقته بأطراف الدعوى أجاب قائلاً: اسمي الكامل: ( ${orDots(w.name)} )، وجنسيتي: ( ${orDots(w.nationality)} )، ورقم هويتي: ( ${orDots(w.idNum)} )، وتاريخ ميلادي/عمري: ( ${orDots(w.age)} )، وأعمل في مهنة: ( ${orDots(w.job)} )، وأسكن في: ( ${orDots(w.residence)} )، وصلتي بالمدعي: ( ${orDots(w.relationPlaintiff)} )، وصلتي بالمدعى عليه: ( ${orDots(w.relationDefendant)} )، ومصلحتي في هذه الدعوى هي: ( ${orDots(w.interest)} )، ثم جرى تحليفه اليمين على أن يشهد بالحق، فحلف، ثم جرى سؤاله عما لديه من شهادة، فأجاب قائلاً: ( ${orDots(w.testimony)} ) هكذا أجاب وشهد.`;
    });

    out += buildTazkiyaClause(ctx, bringVerb);
    out += buildWitnessObjectionClause(ctx);
    return out;
}

// تعديل الشهود وتزكيتهم
function buildTazkiyaClause(ctx, bringVerb) {
    if (ctx.tazkiya === 'delay') {
        return ` ثم جرى سؤال ${ctx.presenterLabel} هل لدي${ctx.speakerSuffix} معدِّلون للشهود؟ فأجاب قائلاً: أطلب إمهالي لإحضارهم إلى جلسة أخرى. هكذا أجاب.`;
    }
    if (ctx.tazkiya === 'presented') {
        return ` ثم ${bringVerb} ${ctx.presenterLabel} المعدِّلين: ( ${orDots(ctx.tazkiyaNames)} )، وبسؤالهم عما لديهم شهدوا قائلين: نشهد بالله تعالى أن الشهود المذكورين عدول ثقات مقبولو الشهادة. هكذا شهدوا.`;
    }
    return '';
}

// حق الخصم في الطعن في الشهود أو في شهادتهم — لا يُسأل عنه إن كان غائبًا
function buildWitnessObjectionClause(ctx) {
    if (!ctx.opposingPresent) return '';
    if (ctx.objection === 'نعم') {
        return ` وبعرض الشهود وشهادتهم على ${ctx.opposingLabel} قرر قائلاً: ${orDots(ctx.objectionText)}. هكذا قرر.`;
    }
    return ` وبعرض الشهود وشهادتهم على ${ctx.opposingLabel} قرر قائلاً: لا مطعن لي فيهم ولا في شهادتهم. هكذا قرر.`;
}

// حالة المدعى عليه تجاه اليمين (تُشتق من حالة حضوره)
function getOathDefendantStatus(state) {
    return state.defendant.attendance === 'لم يحضر' ? 'غائب' : 'حاضر';
}

// ==================== التسلسل الإجرائي للدعوى ====================
// الترتيب: عرض الدعوى ← جواب المدعى عليه ← (إقرار: لا بينة | إنكار: تكليف المدعي بالبينة)
// فلا يُسأل المدعي عن بينته قبل عرض الدعوى على خصمه، إلا إذا كان الخصم غائبًا.

// قسم الدعوى والإجابة والبينات (الجلسة التحضيرية)
function buildClaimEvidenceText(state) {
    const pName = partyLabel('plaintiff', state.plaintiff.gender);
    const claimText = state.claim.text.trim() || '.......';
    return ` وبالاطلاع على دعوى ${pName} ونصها: "${claimText}" أ.هـ.` + buildProceedingsAfterClaim(state);
}

// ما يلي عرضَ الدعوى: جواب المدعى عليه ثم ما يترتب عليه من بينات
function buildProceedingsAfterClaim(state) {
    const c = state.claim;

    // غياب المدعى عليه: لا جواب يُعرض، فيُكلَّف المدعي ببينته مباشرة
    if (state.defendant.attendance === 'لم يحضر') {
        return buildPlaintiffEvidenceText(state);
    }

    let out = buildDefendantAnswerClause(state);

    if (c.defendantStance === 'إقرار') {
        return out + buildAdmissionClause(state);
    }
    if (c.defendantStance === 'دفع شكلي') {
        out += buildFormalPleaClause(state);
        if (c.answeredOnMerits !== 'نعم') return out;
    }

    out += buildPlaintiffEvidenceText(state);
    out += buildDefendantEvidenceText(state);
    return out;
}

// عرض الدعوى على المدعى عليه وإثبات جوابه
function buildDefendantAnswerClause(state) {
    const d = state.defendant;
    const pName = partyLabel('plaintiff', state.plaintiff.gender);
    const dName = partyLabel('defendant', d.gender);
    const isRepresented = d.attendance === 'تمثيل' && (d.repType === 'وكيل' || d.repType === 'وكيل شركة');
    const speakerGender = isRepresented ? d.agentGender : d.gender;
    const addressee = isRepresented ? `وكيل${d.agentGender === 'ف' ? 'ة' : ''} ${dName}` : dName;
    const answerVerb = speakerGender === 'م' ? 'أجاب قائلاً' : 'أجابت قائلة';
    const answerVerbEnd = speakerGender === 'م' ? 'أجاب' : 'أجابت';
    return ` وبعرض دعوى ${pName} على ${addressee} ${answerVerb}: ${orDots(state.claim.defendantResponseText)} هكذا ${answerVerbEnd}.`;
}

// الإقرار يُغني عن البينة: «البينة على المدعي» إنما تُطلب عند الإنكار
function buildAdmissionClause(state) {
    const dName = partyLabel('defendant', state.defendant.gender);
    const pName = partyLabel('plaintiff', state.plaintiff.gender);
    return ` ولمَّا كانت إجابة ${dName} إقراراً بما جاء في الدعوى، وكان الإقرار حجةً قاصرةً على المُقِرّ، فلا موجب لتكليف ${pName} بالبينة؛ إذ إنما تُطلب البينة عند الإنكار.`;
}

// الدفع الشكلي وعرضه على المدعي للرد عليه
function buildFormalPleaClause(state) {
    const c = state.claim;
    const d = state.defendant;
    const dName = partyLabel('defendant', d.gender);
    const pName = partyLabel('plaintiff', state.plaintiff.gender);
    const dSuffix = d.gender === 'م' ? 'ه' : 'ها';
    let out = ` وبسؤال ${dName} عن دفع${dSuffix} الشكلي قرر قائلاً: ${orDots(c.formalPleaText)}. هكذا قرر.`;
    out += ` وبعرض هذا الدفع على ${pName} أجاب قائلاً: ${orDots(c.plaintiffReplyText)}. هكذا أجاب.`;
    if (c.answeredOnMerits !== 'نعم') {
        out += ` ولم يُجب ${dName} في الموضوع، وعليه قررت الدائرة النظر في الدفع الشكلي قبل الخوض في موضوع الدعوى.`;
    }
    return out;
}

// تكليف المدعي بالبينة وعرضها — لا يقع إلا بعد إنكار المدعى عليه أو غيابه
function buildPlaintiffEvidenceText(state, opts) {
    const options = opts || {};
    const c = state.claim;
    const s = state.plaintiff;
    const pGender = s.gender;
    const pName = partyLabel('plaintiff', pGender);
    const dName = partyLabel('defendant', state.defendant.gender);
    const pSuffix = pGender === 'م' ? 'ه' : 'ها';

    const isRepresented = s.attendance === 'تمثيل' && (s.repType === 'وكيل' || s.repType === 'وكيل شركة');
    const speakerGender = isRepresented ? s.agentGender : pGender;
    const speakerSuffix = speakerGender === 'م' ? 'ه' : 'ها';
    const decideVerb = speakerGender === 'م' ? 'قرر قائلاً' : 'قررت قائلة';
    // موكلي/موكلتي: ضمير المتكلم (داخل كلام الوكيل المقتبس)
    const moakkelFirst = pGender === 'م' ? 'موكلي' : 'موكلتي';
    // موكلك/موكلتك: ضمير المخاطَب (في السؤال الموجّه للوكيل)
    const moakkelSecond = pGender === 'م' ? 'موكلك' : 'موكلتك';
    // موكله/موكلها/موكلته/موكلتها: ضمير الغائب (في وصف الدائرة قبل الاقتباس)
    const moakkelThird = pGender === 'م'
        ? (speakerGender === 'م' ? 'موكله' : 'موكلها')
        : (speakerGender === 'م' ? 'موكلته' : 'موكلتها');
    const addresseeForOath = isRepresented ? `وكيل${s.agentGender === 'ف' ? 'ة' : ''} ${pName}` : pName;
    const oathStatus = getOathDefendantStatus(state);

    let out = '';
    if (options.lead) {
        out += options.lead;
    } else if (state.defendant.attendance !== 'لم يحضر' && c.defendantStance !== 'إقرار') {
        out += ` ولمَّا كانت إجابة ${dName} إنكاراً لما جاء في الدعوى، وأن البينة على المدعي، فقد جرى تكليف ${pName} بإحضار بينت${pSuffix}.`;
    }

    if (isRepresented) {
        out += ` وبسؤال${speakerSuffix} عن بينة ${moakkelThird} ${decideVerb}: `;
    } else {
        out += ` وبسؤال${pSuffix} عن بينت${pSuffix} ${decideVerb}: `;
    }

    if (c.evidenceChoice === 'none') {
        out += isRepresented ? `لا بينة لدى ${moakkelFirst}` : `لا بينة لدي`;
        out += buildOathBlock(speakerGender, dName, addresseeForOath, c.requestOath, c.declineOath, oathStatus, isRepresented, moakkelFirst);
    } else {
        const items = c.evidenceItems.map(item =>
            item === 'مستند رسمي آخر' ? (c.otherDocumentText.trim() || 'مستند رسمي آخر') : item
        );
        const evText = items.length ? items.join('، و') : MINUTES_PLACEHOLDER;
        out += isRepresented ? `بينة ${moakkelFirst} هي: ${evText}.` : `بينتي هي: ${evText}.`;
        out += ` وقد جرى من الدائرة الاطلاع على بينات ${pName} فوجدتها كما ${speakerGender === 'م' ? 'ذكر' : 'ذكرت'}.`;

        if (c.evidenceItems.includes('شهادة شهود') && c.witnesses.length > 0) {
            out += buildWitnessSection(c.witnesses, {
                speakerGender, speakerSuffix,
                presenterLabel: pName, presenterGender: pGender,
                opposingLabel: addresseeLabelOf('defendant', state.defendant),
                opposingPresent: state.defendant.attendance !== 'لم يحضر',
                tazkiya: c.tazkiya, tazkiyaNames: c.tazkiyaNames,
                objection: c.witnessObjection, objectionText: c.witnessObjectionText
            });
        }

        if (c.hasMoreEvidence === 'لا') {
            if (isRepresented) {
                out += ` وبسؤال${speakerSuffix} هل لدى ${moakkelSecond} مزيد بينة؟ ${decideVerb}: ليس لدى ${moakkelFirst} مزيد بينة`;
            } else {
                out += ` وبسؤال${pSuffix} هل لديك مزيد بينة؟ ${decideVerb}: ليس لدي مزيد بينة`;
            }
            out += buildOathBlock(speakerGender, dName, addresseeForOath, c.requestOath, c.declineOath, oathStatus, isRepresented, moakkelFirst);
        } else if (c.hasMoreEvidence === 'نعم') {
            const moreText = orDots(c.moreEvidenceText);
            if (isRepresented) {
                out += ` وبسؤال${speakerSuffix} هل لدى ${moakkelSecond} مزيد بينة؟ ${decideVerb}: نعم، بينة ${moakkelFirst} الإضافية هي: ${moreText}.`;
            } else {
                out += ` وبسؤال${pSuffix} هل لديك مزيد بينة؟ ${decideVerb}: نعم، بينت${pSuffix} الإضافية هي: ${moreText}.`;
            }
        }
    }

    return out;
}

// دفوع المدعى عليه وبينته — يُسأل عنها بعد بينة المدعي، ولا محل لها مع غيابه
function buildDefendantEvidenceText(state, opts) {
    const options = opts || {};
    const c = state.claim;
    const d = state.defendant;
    if (d.attendance === 'لم يحضر') return '';
    if (!options.force && c.askDefendantEvidence !== 'نعم') return '';

    const e = c.defendantEvidence;
    const dGender = d.gender;
    const dName = partyLabel('defendant', dGender);
    const pName = partyLabel('plaintiff', state.plaintiff.gender);
    const dSuffix = dGender === 'م' ? 'ه' : 'ها';

    const isRepresented = d.attendance === 'تمثيل' && (d.repType === 'وكيل' || d.repType === 'وكيل شركة');
    const speakerGender = isRepresented ? d.agentGender : dGender;
    const speakerSuffix = speakerGender === 'م' ? 'ه' : 'ها';
    const decideVerb = speakerGender === 'م' ? 'قرر قائلاً' : 'قررت قائلة';
    const moakkelFirst = dGender === 'م' ? 'موكلي' : 'موكلتي';
    const moakkelThird = dGender === 'م'
        ? (speakerGender === 'م' ? 'موكله' : 'موكلها')
        : (speakerGender === 'م' ? 'موكلته' : 'موكلتها');

    let out = options.lead || ' ثم انتقلت الدائرة لسماع دفوع المدعى عليه وبينته.';

    if (c.defendantPleasText.trim()) {
        out += ` وبسؤال ${addresseeLabelOf('defendant', d)} عن دفوع${isRepresented ? ` ${moakkelThird}` : dSuffix} ${decideVerb}: ${c.defendantPleasText.trim()}. هكذا ${speakerGender === 'م' ? 'قرر' : 'قررت'}.`;
    }

    if (isRepresented) {
        out += ` وبسؤال${speakerSuffix} عن بينة ${moakkelThird} ${decideVerb}: `;
    } else {
        out += ` وبسؤال${dSuffix} عن بينت${dSuffix} ${decideVerb}: `;
    }

    if (e.choice === 'none') {
        out += isRepresented ? `لا بينة لدى ${moakkelFirst}.` : `لا بينة لدي.`;
        return out;
    }

    const items = e.items.map(item =>
        item === 'مستند رسمي آخر' ? (e.otherDocumentText.trim() || 'مستند رسمي آخر') : item
    );
    const evText = items.length ? items.join('، و') : MINUTES_PLACEHOLDER;
    out += isRepresented ? `بينة ${moakkelFirst} هي: ${evText}.` : `بينتي هي: ${evText}.`;
    out += ` وقد جرى من الدائرة الاطلاع على بينات ${dName} فوجدتها كما ${speakerGender === 'م' ? 'ذكر' : 'ذكرت'}.`;

    if (e.items.includes('شهادة شهود') && e.witnesses.length > 0) {
        out += buildWitnessSection(e.witnesses, {
            speakerGender, speakerSuffix,
            presenterLabel: dName, presenterGender: dGender,
            opposingLabel: addresseeLabelOf('plaintiff', state.plaintiff),
            opposingPresent: state.plaintiff.attendance !== 'لم يحضر',
            tazkiya: e.tazkiya, tazkiyaNames: e.tazkiyaNames,
            objection: e.objection, objectionText: e.objectionText
        });
    }

    return out;
}

// ==================== الأسباب والحكم والإفهام ====================

// الحكم على وقف واجب التدقيق لزومًا، فلا يُترك نوع الإفهام فيه للاختيار
function mandatoryReviewNotice(state) {
    return state.defendant && state.defendant.entityType === ENTITY_TYPES.WAQF;
}

// نوع الإفهام: يُشتق من قيمة المطالبة ما لم يُحدَّد يدويًا
function noticeKindFor(state) {
    const r = state.ruling;
    if (mandatoryReviewNotice(state)) return 'review';
    if (r.noticeKind && r.noticeKind !== 'auto') return r.noticeKind;
    const c = state.claim || {};
    // حدّ الدعاوى اليسيرة قيمةٌ مالية، فالطلب غير المالي لا يدخله ويبقى قابلاً للاستئناف
    if (c.claimType === 'غير مالي') return 'appealable';
    const value = Number(String(c.claimValue).replace(/[^0-9.]/g, ''));
    if (!value) return 'appealable';
    return value <= YASEERA_CLAIM_LIMIT ? 'final' : 'appealable';
}

// نص الإفهام مأخوذ من تصنيف (إفهامات بعد النطق) في مكتبة النماذج
function buildNoticeText(state) {
    const kind = noticeKindFor(state);
    // إفهام واجب التدقيق له صيغتان: رفع الملف بعد مدة الاعتراض، أو إعادة القضية لمحكمة الاستئناف
    if (kind === 'review') {
        const keyword = state.ruling.mandatoryReview === 'إعادة' ? 'واجب التدقيق /إعادة القضية' : 'واجب التدقيق';
        return findTemplate('إفهامات بعد النطق', keyword) || '';
    }
    if (kind === 'sulh') return findTemplate('إفهامات بعد النطق', 'صلح1') || '';
    if (kind === 'final') return findTemplate('إفهامات بعد النطق', 'ف1') || '';
    // ف3 تتضمن إفهام الطرفين وكالةً استناداً للمادة (165)
    const bothByAgent = state.plaintiff.attendance === 'تمثيل' && state.defendant.attendance === 'تمثيل';
    return findTemplate('إفهامات بعد النطق', bothByAgent ? 'ف3' : 'ف2') || '';
}

function buildPresenceClause(state) {
    const dName = partyLabel('defendant', state.defendant.gender);
    return state.ruling.presence === 'غيابي'
        ? `وهذا الحكم غيابي في حق ${dName}.`
        : 'وهذا الحكم حضوري في حق طرفي الدعوى.';
}

// مرحلة الأسباب والحكم — مستقلة عن ضبط الجلسة، تُلحق به عند النطق بالحكم
function buildRulingSection(state) {
    const r = state.ruling;
    if (r.pronounce !== 'نعم') return '';
    let out = `\n\nالأسباب:\n${orDots(r.reasonsText)}`;
    out += `\n\nالحكم:\n${orDots(r.rulingText)}`;
    out += `\n${buildPresenceClause(state)}`;
    const notice = buildNoticeText(state);
    if (notice) out += `\n\n${notice}`;
    return out;
}

// المادة (167) من نظام المرافعات: استمرار الخلف في نظر القضية
const NEW_JUDGE_CLAUSE = ' واستناداً إلى المادة السابعة والستون بعد المائة من نظام المرافعات الشرعية: إذا انتهت ولايـة القـاضي بالنسبة إلى قضيـة ما قبل النطق بالحكم فيها، فلخلفه الاستمرار في نظرها من الحد الذي انتهت إليه إجراءاتها لدى سلفه بعد تلاوة ما تم ضبطه سابقًا على الخصوم، فإن كانت موقعة بتوقيع القاضي السابق على توقيعات المترافعين والشهود فيعتمدها، وإن كان ما تم ضبطه غير موقع من المترافعين أو أحدهم أو القاضي ولم يصدّق المترافعون عليه فإن المرافعة تعاد من جديد.';

// ==================== التوليد الكامل للمحضر ====================

// وقت إغلاق الجلسة كما أدخله القاضي، ومع الحالات القديمة (وقت افتتاح فقط) يُشتق بإضافة 30 دقيقة
function closingTimeParts(state) {
    const c = state.closing;
    if (c && c.hour) return [c.hour, c.minute || 0, c.period || 'ص'];
    const o = state.opening || {};
    const derived = addMinutesToTime(o.hour || 8, o.minute || 0, o.period || 'ص', 30);
    return [derived.hour, derived.minute, derived.period];
}

// إجراءات الجلسة التالية: سماع بينة أحد الخصمين أو كليهما فيها
function buildFollowUpProceedings(state) {
    const f = state.followUp || {};
    let out = '';
    if (f.plaintiffEvidence) {
        out += buildPlaintiffEvidenceText(state, {
            lead: ' وتنفيذاً لما تقرر في الجلسة السابقة من تكليف المدعي بالبينة،'
        });
    }
    if (f.defendantEvidence) {
        out += buildDefendantEvidenceText(state, { force: true });
    }
    return out;
}

function composeMinutes(state) {
    let text;
    let rulingApplies = false;

    if (state.plaintiff.specialCase !== 'none') {
        const opening = buildOpening(state.opening);
        const specialText = buildSpecialCaseText(state, 'plaintiff');
        if (state.defendant.attendance !== 'لم يحضر') {
            text = `${opening} و${buildPartyClause(state, 'defendant')}، و${specialText}`;
        } else {
            text = `${opening} و${specialText}`;
        }
    } else if (state.plaintiff.attendance === 'لم يحضر') {
        if (state.plaintiff.notifyStatus === 'لم يتبلغ') {
            text = `${buildOpening(state.opening)} و${buildNotNotifiedClause('plaintiff', state.plaintiff)}.`;
        } else {
            text = buildShatbText(state);
        }
    } else {
        const opening = buildOpening(state.opening);
        const includePartyData = state.includePartyDataInText === true;
        const pClauses = [buildPartyClause(state, 'plaintiff'), ...state.extraPlaintiffs.map((p, i) => buildExtraClause('plaintiff', i, p, includePartyData))];
        const oathAbsenceActive = state.sessionType === 'previous' && state.defendant.attendance === 'لم يحضر' && state.defendant.oathAbsence === 'نعم';
        const defendantNotNotified = state.defendant.attendance === 'لم يحضر' && state.defendant.notifyStatus === 'لم يتبلغ';

        if (oathAbsenceActive) {
            text = `${opening} و${pClauses.join('، و')}، ${buildOathAbsenceDefendantText(state)}.`;
        } else if (defendantNotNotified) {
            text = `${opening} و${pClauses.join('، و')}، و${buildNotNotifiedClause('defendant', state.defendant)}.`;
        } else {
            const dClauses = [buildPartyClause(state, 'defendant'), ...state.extraDefendants.map((p, i) => buildExtraClause('defendant', i, p, includePartyData))];
            text = `${opening} و${pClauses.join('، و')}، و${dClauses.join('، و')}.`;
            if (state.sessionType === 'previous' && state.defendant.attendance !== 'لم يحضر' && state.defendant.oathPerformanceSession === 'نعم') {
                text = text.replace(/\.\s*$/, '') + '، ' + buildOathPerformanceText(state) + '.';
            }
            if (state.sessionType === 'new') {
                text += buildClaimEvidenceText(state);
            } else if (state.sessionType === 'previous') {
                if (state.sameJudge === 'لا') {
                    text += NEW_JUDGE_CLAUSE;
                    text += buildPreviousSessionRatification(state);
                }
                text += buildFollowUpProceedings(state);
            }
            // لا يُقفل باب المرافعة إذا وُجهت اليمين لمدعى عليه غائب (تُحدد جلسة قادمة)
            const oathNotificationAdjournment = state.sessionType === 'new' && state.claim.requestOath && !state.claim.declineOath && state.defendant.attendance === 'لم يحضر';
            if (!oathNotificationAdjournment) {
                text = text.replace(/\.\s*$/, '') + '، ' + buildClosingArgumentText(state) + '.';
                rulingApplies = true;
            }
        }
    }

    // الأسباب والحكم: مرحلة مستقلة تُلحق بالضبط بعد قفل باب المرافعة
    const rulingSection = rulingApplies ? buildRulingSection(state) : '';

    // ختم الجلسة بالوقت المُدخل في آخر الضبط
    const closingWords = buildTimeArabic(...closingTimeParts(state));
    if (rulingSection) {
        return `${text}${rulingSection}\n\nوأغلقت الجلسة الساعة ${closingWords}.`;
    }
    return text.replace(/\.\s*$/, '') + `، وأغلقت الجلسة الساعة ${closingWords}.`;
}

// ==================== فحص اكتمال الحقول ====================

function collectWarnings(state) {
    const w = [];
    const partyLabelAr = { plaintiff: 'المدعي', defendant: 'المدعى عليه' };
    // بيانات الطرفين (الاسم والهوية) لا تُفحص إلا إذا كانت ستُدرج في نص الضبط
    const includePartyData = state.includePartyDataInText === true;

    if (!String(state.opening.judge || '').trim()) w.push('لم يُدخل اسم القاضي.');
    if (!String(state.opening.court || '').trim()) w.push('لم يُدخل اسم المحكمة.');

    ['plaintiff', 'defendant'].forEach(p => {
        const s = state[p];
        const label = partyLabelAr[p];
        const extraList = p === 'plaintiff' ? state.extraPlaintiffs : state.extraDefendants;

        // في محضر الشطب أو الحالات الاستثنائية لا تُفحص بيانات المدعى عليه
        if (p === 'defendant' && (state.plaintiff.attendance === 'لم يحضر' || state.plaintiff.specialCase !== 'none')) return;

        if (includePartyData && extraList.length > 0 && !s.name.trim()) {
            w.push(`لم يُدخل اسم ${label} الأول (لازم لكل الأطراف عند تعدد ${label === 'المدعي' ? 'المدعين' : 'المدعى عليهم'}).`);
        }
        if (s.attendance === 'أصالة') {
            if (includePartyData) {
                if (!s.name.trim()) w.push(`لم يُدخل اسم ${label}.`);
                if (isCorporateEntity(s)) {
                    if (!s.crNum.trim()) w.push(`لم يُدخل رقم ${corporateDocLabel(s)} لـ${label}.`);
                } else {
                    if (s.nationalityType === 'سعودي' && (!s.saudiId || s.saudiId.length !== 10)) w.push(`رقم هوية ${label} غير مكتمل (10 أرقام).`);
                    if (s.nationalityType === 'غير ذلك') {
                        if (!s.foreignNationality.trim()) w.push(`لم تُحدَّد جنسية ${label}.`);
                        if (!s.iqamaNum || s.iqamaNum.length !== 10) w.push(`رقم إقامة ${label} غير مكتمل (10 أرقام).`);
                    }
                }
            }
        }
        const needsWakeelCheck = s.attendance === 'تمثيل' && (s.repType === 'وكيل' || s.repType === 'وكيل شركة');
        const needsAccompanyCheck = s.attendance === 'أصالة' && s.hasAccompanyingAgent;
        if (needsWakeelCheck || needsAccompanyCheck) {
            if (!s.agentName.trim()) w.push(`لم يُدخل اسم وكيل ${label}.`);
            if (!s.agentId || s.agentId.length !== 10) w.push(`رقم هوية وكيل ${label} غير مكتمل (10 أرقام).`);
            if (!s.wakalaNum.trim()) w.push(`لم يُدخل رقم الوكالة الشرعية لوكيل ${label}.`);
            if (s.repIsLawyer === 'نعم' && !s.licenseNum.trim()) w.push(`لم يُدخل رقم رخصة المحاماة لوكيل ${label}.`);
            if (s.repIsLawyer === 'لا' && s.repType !== 'وكيل شركة') {
                const isAshar = s.kinshipType === 'أصهار';
                const value = isAshar ? s.kinshipAshar : s.kinship;
                const degreeTable = isAshar ? asharDegree : kinshipDegree;
                if (!value) w.push(`لم تُحدَّد صلة القرابة لوكيل ${label} غير المحامي.`);
                else if (value === 'أخرى' || !degreeTable[value]) w.push(`⚠️ صلة قرابة وكيل ${label} غير مقبولة (خارج الدرجات الأربع).`);
            }
        } else if (s.attendance === 'تمثيل' && s.repType === 'ممثل') {
            if (!s.repName.trim()) w.push(`لم يُدخل اسم الممثل النظامي لـ${label}.`);
            if (!s.repNum.trim()) w.push(`لم تُحدَّد صفة الممثل النظامي لـ${label}.`);
        }
        if (s.attendance === 'لم يحضر' && s.notifyStatus === 'تبلغ') {
            const oathAbsenceActive = p === 'defendant' && state.sessionType === 'previous' && s.oathAbsence === 'نعم';
            if (!oathAbsenceActive && !s.tabligh.trim()) w.push(`لم يُدخل رقم مهمّة التبليغ لـ${label}.`);
        }
        if (p === 'defendant' && state.sessionType === 'previous' && s.attendance === 'لم يحضر' && s.oathAbsence === 'نعم' && !s.oathTablighNum.trim()) {
            w.push('لم يُدخل رقم التبليغ باليمين للمدعى عليه.');
        }
    });

    ['plaintiff', 'defendant'].forEach(p => {
        const list = p === 'plaintiff' ? state.extraPlaintiffs : state.extraDefendants;
        list.forEach((ep, i) => {
            const ordinal = ordinalWord(i + 2, ep.gender);
            if (includePartyData && !ep.name.trim()) w.push(`لم يُدخل اسم ${partyLabelAr[p]} ${ordinal}.`);
            if (ep.attendance === 'أصالة' && includePartyData) {
                if (ep.nationalityType === 'غير ذلك') {
                    if (!String(ep.foreignNationality || '').trim()) w.push(`لم تُحدَّد جنسية ${partyLabelAr[p]} ${ordinal}.`);
                    if (!ep.iqamaNum || ep.iqamaNum.length !== 10) w.push(`رقم إقامة ${partyLabelAr[p]} ${ordinal} غير مكتمل (10 أرقام).`);
                } else if (!ep.saudiId || ep.saudiId.length !== 10) {
                    w.push(`رقم هوية ${partyLabelAr[p]} ${ordinal} غير مكتمل (10 أرقام).`);
                }
            }
            if (ep.attendance === 'تمثيل') {
                if (!ep.repName.trim()) w.push(`لم يُدخل اسم وكيل ${partyLabelAr[p]} ${ordinal}.`);
                if (!ep.repNum.trim()) w.push(`لم يُدخل رقم الوكالة لوكيل ${partyLabelAr[p]} ${ordinal}.`);
            }
        });
    });

    const claimApplies = state.sessionType === 'new'
        && state.plaintiff.attendance !== 'لم يحضر'
        && state.plaintiff.specialCase === 'none'
        && !(state.defendant.attendance === 'لم يحضر' && state.defendant.notifyStatus === 'لم يتبلغ');
    const defendantPresent = state.defendant.attendance !== 'لم يحضر';
    const c = state.claim;
    // في الجلسة التحضيرية تُفحص بينة المدعي دائمًا؛ وفي التالية عند سماعها فيها
    const plaintiffEvidenceApplies = claimApplies
        && !(defendantPresent && c.defendantStance === 'إقرار')
        && !(defendantPresent && c.defendantStance === 'دفع شكلي' && c.answeredOnMerits !== 'نعم');
    const followUp = state.followUp || {};

    if (claimApplies) {
        if (!c.text.trim()) w.push('لم يُكتب نص الدعوى.');
        if (defendantPresent && !c.defendantResponseText.trim()) w.push('لم تُكتب إجابة المدعى عليه على الدعوى.');
        if (defendantPresent && c.defendantStance === 'دفع شكلي') {
            if (!c.formalPleaText.trim()) w.push('لم يُكتب نص الدفع الشكلي للمدعى عليه.');
            if (!c.plaintiffReplyText.trim()) w.push('لم يُكتب رد المدعي على الدفع الشكلي.');
        }
    }

    if (plaintiffEvidenceApplies || followUp.plaintiffEvidence) {
        w.push(...evidenceWarnings({
            choice: c.evidenceChoice, items: c.evidenceItems, otherDocumentText: c.otherDocumentText,
            witnesses: c.witnesses, tazkiya: c.tazkiya, tazkiyaNames: c.tazkiyaNames,
            objection: c.witnessObjection, objectionText: c.witnessObjectionText
        }, 'المدعي', defendantPresent));
        if (c.evidenceChoice === 'has' && c.hasMoreEvidence === 'نعم' && !c.moreEvidenceText.trim()) w.push('لم تُكتب البينة الإضافية.');
    }

    const defendantEvidenceApplies = defendantPresent
        && ((claimApplies && c.askDefendantEvidence === 'نعم') || followUp.defendantEvidence);
    if (defendantEvidenceApplies) {
        w.push(...evidenceWarnings(c.defendantEvidence, 'المدعى عليه', state.plaintiff.attendance !== 'لم يحضر'));
    }

    if (state.sessionType === 'previous' && defendantPresent && state.defendant.oathPerformanceSession === 'نعم' && !state.defendant.oathAnswerText.trim()) {
        w.push('لم تُكتب إجابة المدعى عليه عن الاستعداد لأداء اليمين.');
    }

    w.push(...rulingWarnings(state));
    return w;
}

// فحص كتلة بينة (للمدعي أو للمدعى عليه) بنفس القواعد
function evidenceWarnings(block, ownerLabel, opposingPresent) {
    const w = [];
    if (!block || block.choice !== 'has') return w;
    if (block.items.length === 0) w.push(`لم تُحدَّد بينة ${ownerLabel} (اختر نوعًا واحدًا على الأقل).`);
    if (block.items.includes('مستند رسمي آخر') && !block.otherDocumentText.trim()) {
        w.push(`حُدِّد "مستند رسمي آخر" في بينة ${ownerLabel} دون كتابة ماهية المستند.`);
    }
    if (!block.items.includes('شهادة شهود')) return w;

    const fieldLabels = { name: 'اسم الشاهد', nationality: 'جنسية الشاهد', idNum: 'رقم هوية/إقامة الشاهد', age: 'تاريخ الميلاد/العمر', job: 'المهنة', residence: 'مكان سكن الشاهد', relationPlaintiff: 'صلة الشاهد بالمدعي', relationDefendant: 'صلة الشاهد بالمدعى عليه', interest: 'المصلحة', testimony: 'نص الشهادة' };
    block.witnesses.forEach((wit, i) => {
        const ord = ordinalWordsMasc[i + 1] || `رقم ${i + 1}`;
        Object.keys(fieldLabels).forEach(k => {
            if (!String(wit[k] || '').trim()) w.push(`لم يُكمَل حقل "${fieldLabels[k]}" لشاهد ${ownerLabel} ${ord}.`);
        });
        // رقم الهوية أو الإقامة عشرة أرقام كبيانات الطرفين
        const idNum = String(wit.idNum || '').trim();
        if (idNum && idNum.length !== 10) w.push(`رقم هوية/إقامة شاهد ${ownerLabel} ${ord} غير مكتمل (10 أرقام).`);
    });
    if (block.tazkiya === 'presented' && !String(block.tazkiyaNames || '').trim()) {
        w.push(`لم تُكتب أسماء معدِّلي شهود ${ownerLabel}.`);
    }
    if (opposingPresent && block.objection === 'نعم' && !String(block.objectionText || '').trim()) {
        w.push(`حُدِّد وجود مطعن في شهود ${ownerLabel} دون كتابة نص المطعن.`);
    }
    return w;
}

// فحص مرحلة الأسباب والحكم
function rulingWarnings(state) {
    const w = [];
    const r = state.ruling || {};
    if (r.pronounce !== 'نعم') return w;
    if (!String(r.reasonsText || '').trim()) w.push('لم تُكتب أسباب الحكم.');
    if (!String(r.rulingText || '').trim()) w.push('لم يُكتب منطوق الحكم.');
    const claim = state.claim || {};
    if (r.noticeKind === 'auto' && !mandatoryReviewNotice(state) && claim.claimType !== 'غير مالي' && !String(claim.claimValue || '').trim()) {
        w.push('لم تُدخل قيمة المطالبة، فتعذّر اشتقاق نوع الإفهام تلقائيًا (اعتُمد "قابل للاستئناف").');
    }
    if (!buildNoticeText(state)) {
        w.push('⚠️ تعذّر جلب نص الإفهام من مكتبة النماذج (data/templates.js).');
    }
    return w;
}

// تصدير للاختبارات في بيئة Node (لا يؤثر على المتصفح)
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        MINUTES_PLACEHOLDER, VIDEO_CALL_PHRASE, VIDEO_CALL_SHORT, VIDEO_CALL_BRIEF, VIDEO_CALL_BASIS,
        IN_PERSON_SHORT, SESSION_MODES, sessionMode, isInPersonSession,
        NATIONALITY_OPTIONS, WITNESS_NATIONALITY_OPTIONS, EVIDENCE_OPTIONS, YASEERA_CLAIM_LIMIT, ENTITY_TYPES,
        isCorporateEntity, corporateDocLabel, mandatoryReviewNotice, closingTimeParts,
        minutesPhrase, buildTimeArabic, addMinutesToTime, validHoursForPeriod, convertArabicDigits,
        ordinalWord, partyLabel, multiPartyLabel, agentPossessive,
        kinshipDegree, asharDegree, degreeOrdinal, asharTemplate,
        findTemplate, templatesOfCategory,
        normalizeArabicSearch, searchTemplates, curatedReasonsCount, CURATED_REASON_OPENER,
        freshPartyState, freshExtraParty, freshWitness, freshEvidenceBlock,
        freshClaimState, freshRulingState, freshFollowUpState, freshMinutesState,
        buildOpening, buildAgencyVerificationClause, buildKinshipPhrase, buildAccompanyingAgentClause,
        buildIdentityClause, buildPartyClause, buildExtraClause, buildNotNotifiedClause,
        buildSpecialCaseText, buildShatbText, buildOathAbsenceDefendantText, buildOathPerformanceText,
        buildClosingArgumentText, buildPreviousSessionRatification,
        buildOathSentence, buildOathBlock, buildWitnessSection, buildTazkiyaClause, buildWitnessObjectionClause,
        buildClaimEvidenceText, buildProceedingsAfterClaim, buildDefendantAnswerClause,
        buildAdmissionClause, buildFormalPleaClause,
        buildPlaintiffEvidenceText, buildDefendantEvidenceText, buildFollowUpProceedings,
        noticeKindFor, buildNoticeText, buildPresenceClause, buildRulingSection,
        getOathDefendantStatus, composeMinutes, collectWarnings, evidenceWarnings, rulingWarnings
    };
}
