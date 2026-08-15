// ==================== محرك مولّد محضر الجلسة ====================
// هذا الملف يحتوي دوال الصياغة الصافية فقط (بلا أي تعامل مع DOM)
// حتى يمكن اختباره آلياً عبر `npm test` (انظر tests/minutes.test.js)
// واجهة المستخدم في js/minutes-app.js وصفحة minutes.html
//
// فكرة المولّد وإعداد الصياغ القضائية: القاضي/ زياد بن محمد البابطين

// ==================== ثوابت عامة ====================

const MINUTES_PLACEHOLDER = '...........';

// صيغة انعقاد الجلسة عبر الاتصال المرئي (النص الكامل في الافتتاح)
const VIDEO_CALL_PHRASE = 'عبر الاتصال المرئي بأطراف الدَّعوى -من خلال الأنظمة الالكترونية المعتمدة لوزارة العدل، استناداً إلى الفقرة الرابعة من قرار رئيس المجلس الأعلى للقضاء رقم (17388) وتاريخ 5/10/1441هـ والمتضمنة: " تعقد المحكمة جلساتها عن بُعد عبر الأنظمة الالكترونية لوزارة العدل دون الحاجة إلى حضور أطراف الدَّعوى؛ إلا في الحالات التي تقتضي حضورهم " -';

// الصيغة المختصرة لطريقة الانعقاد (تُستخدم في نص الشطب)
const VIDEO_CALL_SHORT = 'بالاتصال المرئي';

const NATIONALITY_OPTIONS = ['آيسلندي', 'أذربيجاني', 'أرجنتيني', 'أردني', 'أرمني', 'أسترالي', 'أفريقي وسطي', 'أفغاني', 'ألباني', 'ألماني', 'أمريكي', 'أنغولي', 'أوروغوياني', 'أوزبكستاني', 'أوغندي', 'أوكراني', 'أيرلندي', 'إثيوبي', 'إريتري', 'إسباني', 'إستوني', 'إسواتيني', 'إكوادوري', 'إماراتي', 'إندونيسي', 'إيراني', 'إيطالي', 'الرأس الأخضر', 'بابوا غينيا الجديدة', 'باراغوياني', 'باكستاني', 'بحريني', 'برازيلي', 'برتغالي', 'بروناوي', 'بريطاني', 'بلجيكي', 'بلغاري', 'بنغلاديشي', 'بنمي', 'بنيني', 'بوتاني', 'بوتسواني', 'بوركينابي', 'بوروندي', 'بوسني', 'بولندي', 'بوليفي', 'بيروفي', 'بيلاروسي', 'تايلاندي', 'تايواني', 'تركمانستاني', 'تركي', 'ترينيدادي', 'تشادي', 'تشيكي', 'تشيلي', 'تنزاني', 'توغولي', 'تونسي', 'تيموري', 'جامايكي', 'جزائري', 'جزر القمر', 'جنوب أفريقي', 'جنوب سوداني', 'جورجي', 'جيبوتي', 'دنماركي', 'دومينيكاني', 'رواندي', 'روسي', 'روماني', 'زامبي', 'زيمبابوي', 'ساحل عاجي', 'ساو تومي وبرينسيبي', 'سريلانكي', 'سلفادوري', 'سلوفاكي', 'سلوفيني', 'سنغافوري', 'سنغالي', 'سوداني', 'سوري', 'سورينامي', 'سويدي', 'سويسري', 'سيراليوني', 'سيشيلي', 'صربي', 'صومالي', 'صيني', 'طاجيكستاني', 'عديم الجنسية', 'عراقي', 'عُماني', 'غابوني', 'غامبي', 'غاني', 'غواتيمالي', 'غياني', 'غيني', 'غيني استوائي', 'غيني بيساوي', 'فرنسي', 'فلبيني', 'فلسطيني', 'فنزويلي', 'فنلندي', 'فيتنامي', 'فيجي', 'قبرصي', 'قطري', 'قيرغيزستاني', 'كازاخستاني', 'كاميروني', 'كرواتي', 'كمبودي', 'كندي', 'كوبي', 'كوري جنوبي', 'كوري شمالي', 'كوستاريكي', 'كولومبي', 'كونغولي', 'كونغولي ديمقراطي', 'كويتي', 'كيني', 'لاتفي', 'لاوسي', 'لبناني', 'لوكسمبورغي', 'ليبي', 'ليبيري', 'ليتواني', 'ليسوتوي', 'مالطي', 'مالي', 'ماليزي', 'مدغشقري', 'مصري', 'مغربي', 'مغولي', 'مقدوني', 'مكسيكي', 'ملاوي', 'موريتاني', 'موريشيوسي', 'موزمبيقي', 'مولدوفي', 'مونتينيغري', 'ميانماري', 'ناميبي', 'نرويجي', 'نمساوي', 'نيبالي', 'نيجري', 'نيجيري', 'نيكاراغوي', 'نيوزيلندي', 'هايتي', 'هندوراسي', 'هندي', 'هنغاري', 'هولندي', 'هونغ كونغي', 'ياباني', 'يمني', 'يوناني', 'أخرى'];

const EVIDENCE_OPTIONS = ['العقد', 'الفاتورة', 'ورقة إقرار', 'سند لأمر', 'إيصال استلام', 'محضر تسليم', 'الحوالة', 'رسالة نصية / واتساب', 'تسجيل صوتي', 'شهادة شهود', 'تقرير خبرة', 'تقدير الأضرار', 'تقرير نجم', 'تقرير المرور', 'عرض سعر', 'صور فوتوغرافية', 'مقاطع فيديو', 'مستند رسمي آخر'];

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
    return { name: '', gender: 'م', attendance: 'أصالة', repName: '', repNum: '' };
}

function freshWitness() {
    return { name: '', age: '', job: '', residence: '', relation: '', interest: '', testimony: '' };
}

function freshClaimState() {
    return {
        text: '', evidenceChoice: 'none', evidenceItems: [], otherDocumentText: '',
        hasMoreEvidence: 'نعم', moreEvidenceText: '',
        requestOath: false, declineOath: false,
        witnesses: [], defendantResponseText: ''
    };
}

// الحالة الكاملة للمحضر (تُمرَّر لكل دوال التوليد)
function freshMinutesState() {
    return {
        sessionType: null,           // 'new' جلسة تحضيرية | 'previous' منظورة سابقًا
        sameJudge: 'نعم',
        opening: { judge: '', court: '', hour: 8, minute: 0, period: 'ص' },
        plaintiff: freshPartyState(),
        defendant: freshPartyState(),
        extraPlaintiffs: [],
        extraDefendants: [],
        claim: freshClaimState()
    };
}

// ==================== بناة الجُمل ====================

function orDots(value) {
    const v = (value == null ? '' : String(value)).trim();
    return v || MINUTES_PLACEHOLDER;
}

// فقرة افتتاح الجلسة
function buildOpening(opening) {
    const judge = orDots(opening.judge);
    const court = orDots(opening.court);
    const time = buildTimeArabic(opening.hour, opening.minute, opening.period);
    return `فلديّ أنا ${judge} القاضي بمحكمة ${court} افتتحتُ الجلسة ${VIDEO_CALL_PHRASE} عند الساعة ${time}،`;
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

// إثبات هوية الطرف الحاضر (هوية وطنية / إقامة / سجل تجاري)
function buildIdentityClause(s) {
    if (s.entityType === 'شركة') {
        return `، بموجب السجل التجاري رقم (${orDots(s.crNum)})`;
    }
    if (s.nationalityType === 'سعودي') {
        return `، بموجب الهوية الوطنية رقم (${orDots(s.saudiId)})`;
    }
    return `، ${orDots(s.foreignNationality)} الجنسية، بموجب الإقامة النظامية رقم (${orDots(s.iqamaNum)})`;
}

// فقرة حضور/غياب الطرف الأساسي (المدعي أو المدعى عليه)
function buildPartyClause(state, role) {
    const s = state[role];
    const extraList = role === 'plaintiff' ? state.extraPlaintiffs : state.extraDefendants;
    const hasMultiple = extraList.length > 0;
    const name = hasMultiple
        ? multiPartyLabel(role, s.gender, 1, s.name)
        : (s.name.trim() ? `${partyLabel(role, s.gender)} ${s.name.trim()}` : partyLabel(role, s.gender));
    const suffix = s.gender === 'م' ? 'ه' : 'ها';

    if (s.attendance === 'أصالة') {
        const verb = s.gender === 'م' ? 'حضر' : 'حضرت';
        let clause = `${verb} ${name} أصالة${buildIdentityClause(s)}`;
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
function buildExtraClause(role, idx, p) {
    const name = multiPartyLabel(role, p.gender, idx + 2, p.name);
    const suffix = p.gender === 'م' ? 'ه' : 'ها';

    if (p.attendance === 'أصالة') {
        const verb = p.gender === 'م' ? 'حضر' : 'حضرت';
        return `${verb} ${name} أصالة`;
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
        return `حيث تبيّن حضور ${subjectName} في النظام الإلكتروني، إلا ${heShe} لم ${verbHadara} الجلسة عبر الاتصال المرئي، ولم تتمكن الدائرة من شطب الجلسة بسبب ثبوت ${possessive} في النظام، وعليه جرى إثبات ذلك في محضر الجلسة`;
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

    const bodyCore = `ولم ${verbHadara} فيها ${name}، ولا مَنْ يُمثِّل${suffix}، ولم ${verbTaqaddama} بعذر تقبله الدَّائرة، وقد جرى تبليغ ${name} بعقد هذه الجلسة - ${VIDEO_CALL_SHORT} - وفق مهمّة التبليغ رقم ( ${tabligh} ) كما جرى بعث وتبليغ ${name} برابط الدخول لغرفة الاتصال المرئي، ولم يكن ${lahuOrLaha} وجود حتَّى انتهاءِ وقت الجلسة المُحدَّد،`;

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
function buildWitnessSection(witnesses, speakerGender, speakerSuffix, pName) {
    const decideShort = speakerGender === 'ف' ? 'قررت' : 'قرر';
    let out = ` وبسؤال${speakerSuffix} هل من يشهد لك حاضر؟ ف${decideShort} قائلاً نعم. هكذا ${decideShort}. استناداً إلى نظام الإثبات، وتحديداً المادتين (الحادية والسبعين) و(الثامنة والسبعين) منه، يجب على المحكمة قبل أداء الشهادة أن تطلب من الشاهد الإفصاح عن بياناته الشخصية وعلاقته بالدعوى وتدوين ذلك في الضبط.`;

    witnesses.forEach((w, i) => {
        const ordinal = ordinalWordsMasc[i + 1] || `رقم ${i + 1}`;
        out += ` ثم أحضر ${pName} للشهادة الشاهد ${ordinal}، وبسؤاله عن بياناته الشخصية وعلاقته بأطراف الدعوى أجاب قائلاً: اسمي الكامل: ( ${orDots(w.name)} )، وتاريخ ميلادي/عمري: ( ${orDots(w.age)} )، وأعمل في مهنة: ( ${orDots(w.job)} )، ومكان إقامتي في: ( ${orDots(w.residence)} )، ووجه اتصالي بأطراف الدعوى هو: ( ${orDots(w.relation)} )، ومصلحتي في هذه الدعوى هي: ( ${orDots(w.interest)} )، ثم جرى سؤاله عما لديه من شهادة، فأجاب قائلاً: أشهد بالله العظيم أن ( ${orDots(w.testimony)} ) هكذا أجاب وشهد.`;
    });
    return out;
}

// حالة المدعى عليه تجاه اليمين (تُشتق من حالة حضوره)
function getOathDefendantStatus(state) {
    return state.defendant.attendance === 'لم يحضر' ? 'غائب' : 'حاضر';
}

// قسم الدعوى والبينة والإجابة (الجلسة التحضيرية)
function buildClaimEvidenceText(state) {
    const c = state.claim;
    const s = state.plaintiff;
    const pGender = s.gender;
    const pName = partyLabel('plaintiff', pGender);
    const dName = partyLabel('defendant', state.defendant.gender);
    const pSuffix = pGender === 'م' ? 'ه' : 'ها';
    const claimText = c.text.trim() || '.......';

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

    let out = ` وبالاطلاع على دعوى ${pName} وجدت نصها: "${claimText}" أ.هـ.`;

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
            out += buildWitnessSection(c.witnesses, speakerGender, speakerSuffix, pName);
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

    if (state.defendant.attendance !== 'لم يحضر') {
        const dState = state.defendant;
        const dIsRepresented = dState.attendance === 'تمثيل' && (dState.repType === 'وكيل' || dState.repType === 'وكيل شركة');
        const dSpeakerGender = dIsRepresented ? dState.agentGender : dState.gender;
        const dAddressee = dIsRepresented ? `وكيل${dState.agentGender === 'ف' ? 'ة' : ''} ${dName}` : dName;
        const dResponseVerb = dSpeakerGender === 'م' ? 'أجاب قائلاً' : 'أجابت قائلة';
        const dResponseVerbEnd = dSpeakerGender === 'م' ? 'أجاب' : 'أجابت';
        out += ` وبعرض دعوى ${pName} على ${dAddressee} ${dResponseVerb}: ${orDots(c.defendantResponseText)} هكذا ${dResponseVerbEnd}.`;
    }

    return out;
}

// المادة (167) من نظام المرافعات: استمرار الخلف في نظر القضية
const NEW_JUDGE_CLAUSE = ' واستناداً إلى المادة السابعة والستون بعد المائة من نظام المرافعات الشرعية: إذا انتهت ولايـة القـاضي بالنسبة إلى قضيـة ما قبل النطق بالحكم فيها، فلخلفه الاستمرار في نظرها من الحد الذي انتهت إليه إجراءاتها لدى سلفه بعد تلاوة ما تم ضبطه سابقًا على الخصوم، فإن كانت موقعة بتوقيع القاضي السابق على توقيعات المترافعين والشهود فيعتمدها، وإن كان ما تم ضبطه غير موقع من المترافعين أو أحدهم أو القاضي ولم يصدّق المترافعون عليه فإن المرافعة تعاد من جديد.';

// ==================== التوليد الكامل للمحضر ====================

function composeMinutes(state) {
    let text;

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
        const pClauses = [buildPartyClause(state, 'plaintiff'), ...state.extraPlaintiffs.map((p, i) => buildExtraClause('plaintiff', i, p))];
        const oathAbsenceActive = state.sessionType === 'previous' && state.defendant.attendance === 'لم يحضر' && state.defendant.oathAbsence === 'نعم';
        const defendantNotNotified = state.defendant.attendance === 'لم يحضر' && state.defendant.notifyStatus === 'لم يتبلغ';

        if (oathAbsenceActive) {
            text = `${opening} و${pClauses.join('، و')}، ${buildOathAbsenceDefendantText(state)}.`;
        } else if (defendantNotNotified) {
            text = `${opening} و${pClauses.join('، و')}، و${buildNotNotifiedClause('defendant', state.defendant)}.`;
        } else {
            const dClauses = [buildPartyClause(state, 'defendant'), ...state.extraDefendants.map((p, i) => buildExtraClause('defendant', i, p))];
            text = `${opening} و${pClauses.join('، و')}، و${dClauses.join('، و')}.`;
            if (state.sessionType === 'previous' && state.defendant.attendance !== 'لم يحضر' && state.defendant.oathPerformanceSession === 'نعم') {
                text = text.replace(/\.\s*$/, '') + '، ' + buildOathPerformanceText(state) + '.';
            }
            if (state.sessionType === 'new') {
                text += buildClaimEvidenceText(state);
            } else if (state.sessionType === 'previous' && state.sameJudge === 'لا') {
                text += NEW_JUDGE_CLAUSE;
                text += buildPreviousSessionRatification(state);
            }
            // لا يُقفل باب المرافعة إذا وُجهت اليمين لمدعى عليه غائب (تُحدد جلسة قادمة)
            const oathNotificationAdjournment = state.sessionType === 'new' && state.claim.requestOath && !state.claim.declineOath && state.defendant.attendance === 'لم يحضر';
            if (!oathNotificationAdjournment) {
                text = text.replace(/\.\s*$/, '') + '، ' + buildClosingArgumentText(state) + '.';
            }
        }
    }

    // ختم الجلسة بعد 30 دقيقة من الافتتاح
    const closing = addMinutesToTime(state.opening.hour, state.opening.minute, state.opening.period, 30);
    const closingWords = buildTimeArabic(closing.hour, closing.minute, closing.period);
    return text.replace(/\.\s*$/, '') + `، وختمت الجلسة عند الساعة ${closingWords}.`;
}

// ==================== فحص اكتمال الحقول ====================

function collectWarnings(state) {
    const w = [];
    const partyLabelAr = { plaintiff: 'المدعي', defendant: 'المدعى عليه' };

    if (!String(state.opening.judge || '').trim()) w.push('لم يُدخل اسم القاضي.');
    if (!String(state.opening.court || '').trim()) w.push('لم يُدخل اسم المحكمة.');

    ['plaintiff', 'defendant'].forEach(p => {
        const s = state[p];
        const label = partyLabelAr[p];
        const extraList = p === 'plaintiff' ? state.extraPlaintiffs : state.extraDefendants;

        // في محضر الشطب أو الحالات الاستثنائية لا تُفحص بيانات المدعى عليه
        if (p === 'defendant' && (state.plaintiff.attendance === 'لم يحضر' || state.plaintiff.specialCase !== 'none')) return;

        if (extraList.length > 0 && !s.name.trim()) {
            w.push(`لم يُدخل اسم ${label} الأول (لازم لكل الأطراف عند تعدد ${label === 'المدعي' ? 'المدعين' : 'المدعى عليهم'}).`);
        }
        if (s.attendance === 'أصالة') {
            if (!s.name.trim()) w.push(`لم يُدخل اسم ${label}.`);
            if (s.entityType === 'شركة') {
                if (!s.crNum.trim()) w.push(`لم يُدخل رقم السجل التجاري لـ${label}.`);
            } else {
                if (s.nationalityType === 'سعودي' && (!s.saudiId || s.saudiId.length !== 10)) w.push(`رقم هوية ${label} غير مكتمل (10 أرقام).`);
                if (s.nationalityType === 'غير ذلك') {
                    if (!s.foreignNationality.trim()) w.push(`لم تُحدَّد جنسية ${label}.`);
                    if (!s.iqamaNum || s.iqamaNum.length !== 10) w.push(`رقم إقامة ${label} غير مكتمل (10 أرقام).`);
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
            if (!ep.name.trim()) w.push(`لم يُدخل اسم ${partyLabelAr[p]} ${ordinal}.`);
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
    if (claimApplies) {
        const c = state.claim;
        if (!c.text.trim()) w.push('لم يُكتب نص الدعوى.');
        if (state.defendant.attendance !== 'لم يحضر' && !c.defendantResponseText.trim()) w.push('لم تُكتب إجابة المدعى عليه على الدعوى.');
        const evidenceCount = c.evidenceItems.length;
        if (c.evidenceChoice === 'has' && evidenceCount === 0) w.push('لم تُحدَّد البينة (اختر نوعًا واحدًا على الأقل).');
        if (c.evidenceChoice === 'has' && c.evidenceItems.includes('مستند رسمي آخر') && !c.otherDocumentText.trim()) w.push('حُدِّد "مستند رسمي آخر" دون كتابة ماهية المستند.');
        if (c.evidenceChoice === 'has' && c.hasMoreEvidence === 'نعم' && !c.moreEvidenceText.trim()) w.push('لم تُكتب البينة الإضافية.');
        if (c.evidenceChoice === 'has' && c.evidenceItems.includes('شهادة شهود')) {
            const fieldLabels = { name: 'الاسم', age: 'تاريخ الميلاد/العمر', job: 'المهنة', residence: 'مكان الإقامة', relation: 'وجه الاتصال', interest: 'المصلحة', testimony: 'نص الشهادة' };
            c.witnesses.forEach((wit, i) => {
                const ord = ordinalWordsMasc[i + 1] || `رقم ${i + 1}`;
                Object.keys(fieldLabels).forEach(k => {
                    if (!wit[k].trim()) w.push(`لم يُكمَل حقل "${fieldLabels[k]}" للشاهد ${ord}.`);
                });
            });
        }
    }
    if (state.sessionType === 'previous' && state.defendant.attendance !== 'لم يحضر' && state.defendant.oathPerformanceSession === 'نعم' && !state.defendant.oathAnswerText.trim()) {
        w.push('لم تُكتب إجابة المدعى عليه عن الاستعداد لأداء اليمين.');
    }

    return w;
}

// تصدير للاختبارات في بيئة Node (لا يؤثر على المتصفح)
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        MINUTES_PLACEHOLDER, VIDEO_CALL_PHRASE, VIDEO_CALL_SHORT,
        NATIONALITY_OPTIONS, EVIDENCE_OPTIONS,
        minutesPhrase, buildTimeArabic, addMinutesToTime, validHoursForPeriod, convertArabicDigits,
        ordinalWord, partyLabel, multiPartyLabel, agentPossessive,
        kinshipDegree, asharDegree, degreeOrdinal, asharTemplate,
        freshPartyState, freshExtraParty, freshWitness, freshClaimState, freshMinutesState,
        buildOpening, buildAgencyVerificationClause, buildKinshipPhrase, buildAccompanyingAgentClause,
        buildIdentityClause, buildPartyClause, buildExtraClause, buildNotNotifiedClause,
        buildSpecialCaseText, buildShatbText, buildOathAbsenceDefendantText, buildOathPerformanceText,
        buildClosingArgumentText, buildPreviousSessionRatification,
        buildOathSentence, buildOathBlock, buildWitnessSection, buildClaimEvidenceText,
        getOathDefendantStatus, composeMinutes, collectWarnings
    };
}
