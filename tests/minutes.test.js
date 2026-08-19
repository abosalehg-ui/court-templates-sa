// اختبارات محرك مولّد محضر الجلسة في js/minutes.js
// التشغيل: npm test  (أو: node --test tests/)
const test = require('node:test');
const assert = require('node:assert/strict');

const M = require('../js/minutes.js');

// حالة جاهزة: جلسة تحضيرية، الطرفان حاضران أصالة ببيانات مكتملة
// مع تفعيل إدراج بيانات الطرفين في النص (المفتاح مُغلق افتراضًا)
function readyState(overrides = {}) {
    const s = M.freshMinutesState();
    s.sessionType = 'new';
    s.includePartyDataInText = true;
    s.opening = { judge: 'فلان بن فلان', court: 'الرياض', hour: 9, minute: 0, period: 'ص' };
    s.plaintiff.name = 'سعد';
    s.plaintiff.saudiId = '1000000001';
    s.defendant.name = 'فهد';
    s.defendant.saudiId = '1000000002';
    s.claim.text = 'أطالب بمبلغ مئة ألف ريال';
    s.claim.defendantResponseText = 'ما ذكره المدعي غير صحيح';
    return Object.assign(s, overrides);
}

// ==================== تحويل الوقت إلى كتابة ====================

test('buildTimeArabic: الساعات الكاملة والكسور', () => {
    assert.equal(M.buildTimeArabic(8, 0, 'ص'), 'الثامنة صباحًا');
    assert.equal(M.buildTimeArabic(9, 15, 'ص'), 'التاسعة والربع صباحًا');
    assert.equal(M.buildTimeArabic(10, 30, 'ص'), 'العاشرة والنصف صباحًا');
    assert.equal(M.buildTimeArabic(1, 0, 'م'), 'الواحدة مساءً');
});

test('buildTimeArabic: إلا ربعًا تنسب للساعة التالية', () => {
    assert.equal(M.buildTimeArabic(8, 45, 'ص'), 'التاسعة إلا ربعًا صباحًا');
    assert.equal(M.buildTimeArabic(12, 45, 'م'), 'الواحدة إلا ربعًا مساءً');
});

test('minutesPhrase: مفرد ومثنى وعقود ومركبات', () => {
    assert.equal(M.minutesPhrase(1), 'ودقيقة واحدة');
    assert.equal(M.minutesPhrase(2), 'ودقيقتين');
    assert.equal(M.minutesPhrase(5), 'وخمس دقائق');
    assert.equal(M.minutesPhrase(11), 'وإحدى عشرة دقيقة');
    assert.equal(M.minutesPhrase(20), 'وعشرين دقيقة');
    assert.equal(M.minutesPhrase(23), 'وثلاث وعشرين دقيقة');
    assert.equal(M.minutesPhrase(55), 'وخمس وخمسين دقيقة');
});

test('addMinutesToTime: داخل نفس الفترة', () => {
    assert.deepEqual(M.addMinutesToTime(9, 10, 'ص', 30), { hour: 9, minute: 40, period: 'ص' });
});

test('addMinutesToTime: العبور من الصباح إلى المساء', () => {
    assert.deepEqual(M.addMinutesToTime(11, 45, 'ص', 30), { hour: 12, minute: 15, period: 'م' });
});

test('addMinutesToTime: الثانية عشرة تُحسب صفراً في نظام 12 ساعة', () => {
    assert.deepEqual(M.addMinutesToTime(12, 40, 'م', 30), { hour: 1, minute: 10, period: 'م' });
});

test('validHoursForPeriod: ساعات الجلسات الصباحية والمسائية', () => {
    assert.deepEqual(M.validHoursForPeriod('ص'), [8, 9, 10, 11]);
    assert.deepEqual(M.validHoursForPeriod('م'), [12, 1, 2, 3]);
});

test('convertArabicDigits: تحويل الأرقام المشرقية', () => {
    assert.equal(M.convertArabicDigits('١٢٣٤٥٦٧٨٩٠'), '1234567890');
    assert.equal(M.convertArabicDigits('رقم ٥ فقط'), 'رقم 5 فقط');
});

// ==================== المطابقة النحوية ====================

test('partyLabel و ordinalWord و multiPartyLabel', () => {
    assert.equal(M.partyLabel('plaintiff', 'م'), 'المدعي');
    assert.equal(M.partyLabel('plaintiff', 'ف'), 'المدعية');
    assert.equal(M.partyLabel('defendant', 'ف'), 'المدعى عليها');
    assert.equal(M.ordinalWord(2, 'م'), 'الثاني');
    assert.equal(M.ordinalWord(2, 'ف'), 'الثانية');
    assert.equal(M.multiPartyLabel('plaintiff', 'م', 1, 'سعد'), 'المدعي الأول سعد');
});

test('agentPossessive: وكيل/وكيلة × مذكر/مؤنث', () => {
    assert.equal(M.agentPossessive('م', 'ه'), 'وكيله');
    assert.equal(M.agentPossessive('م', 'ها'), 'وكيلها');
    assert.equal(M.agentPossessive('ف', 'ه'), 'وكيلته');
    assert.equal(M.agentPossessive('ف', 'ها'), 'وكيلتها');
});

// ==================== درجات القرابة ====================

test('جداول القرابة: نسب وأصهار', () => {
    assert.equal(M.kinshipDegree['الأب'], 1);
    assert.equal(M.kinshipDegree['ابن الخالة'], 4);
    assert.equal(M.asharDegree['أب الزوجة'], 1);
    assert.equal(M.asharDegree['بنت خالة الزوج'], 4);
});

test('buildKinshipPhrase: قرابة مقبولة وخارج الدرجات', () => {
    const s = M.freshPartyState();
    s.repIsLawyer = 'لا';
    s.kinship = 'الأخ';
    assert.match(M.buildKinshipPhrase(s, 'المدعي'), /صلة قرابة: الأخ/);
    s.kinship = 'أخرى';
    const phrase = M.buildKinshipPhrase(s, 'المدعي');
    assert.match(phrase, /المادة \(1\/7\)/);
    assert.match(phrase, /فلا يحق له الاستمرار في الإجابة/);
});

// ==================== بناة الجمل ====================

test('buildOpening: الافتراضي عبر الاتصال المرئي مختصراً بلا نص القرار', () => {
    const text = M.buildOpening({ judge: 'فلان', court: 'الرياض', hour: 8, minute: 30, period: 'ص' });
    assert.match(text, /^فلديّ أنا فلان القاضي بمحكمة الرياض افتتحتُ الجلسة عبر الاتصال المرئي عند الساعة/);
    assert.ok(!text.includes('17388'));
    assert.match(text, /عند الساعة الثامنة والنصف صباحًا،$/);
});

test('buildOpening: خيار «مع المستند» يدرج نص قرار المجلس الأعلى للقضاء', () => {
    const text = M.buildOpening({ judge: 'فلان', court: 'الرياض', hour: 8, minute: 30, period: 'ص', mode: M.SESSION_MODES.VIDEO_FULL });
    assert.match(text, /^فلديّ أنا فلان القاضي بمحكمة الرياض افتتحتُ الجلسة عبر الاتصال المرئي/);
    assert.match(text, /قرار فضيلة رئيس المجلس الأعلى للقضاء رقم 17388/);
    assert.match(text, /المتضمن إطلاق خدمة التقاضي عن بعد/);
    assert.match(text, /عند الساعة الثامنة والنصف صباحًا،$/);
});

test('buildOpening: الجلسة الحضورية تحذف صيغة الاتصال المرئي كاملة', () => {
    const text = M.buildOpening({ judge: 'فلان', court: 'الرياض', hour: 8, minute: 30, period: 'ص', mode: M.SESSION_MODES.IN_PERSON });
    assert.equal(text, 'فلديّ أنا فلان القاضي بمحكمة الرياض افتتحتُ الجلسة عند الساعة الثامنة والنصف صباحًا،');
    assert.ok(!text.includes('المرئي'));
});

test('sessionMode: المختصر افتراضياً لأي قيمة غير معروفة', () => {
    assert.equal(M.sessionMode({}), M.SESSION_MODES.VIDEO);
    assert.equal(M.sessionMode({ mode: 'شيء آخر' }), M.SESSION_MODES.VIDEO);
    assert.equal(M.sessionMode({ mode: 'inPerson' }), M.SESSION_MODES.IN_PERSON);
    assert.equal(M.freshMinutesState().opening.mode, M.SESSION_MODES.VIDEO);
});

test('buildIdentityClause: فرد سعودي / مقيم / شركة', () => {
    const s = M.freshPartyState();
    s.saudiId = '1234567890';
    assert.equal(M.buildIdentityClause(s), '، بموجب الهوية الوطنية رقم (1234567890)');
    s.nationalityType = 'غير ذلك';
    s.foreignNationality = 'مصري';
    s.iqamaNum = '2345678901';
    assert.equal(M.buildIdentityClause(s), '، مصري الجنسية، بموجب الإقامة النظامية رقم (2345678901)');
    s.entityType = 'شركة';
    s.crNum = '55555';
    assert.equal(M.buildIdentityClause(s), '، بموجب السجل التجاري رقم (55555)');
});

test('freshMinutesState: مفتاح إدراج بيانات الطرفين مُغلق افتراضًا', () => {
    assert.equal(M.freshMinutesState().includePartyDataInText, false);
});

test('buildIdentityClause: تسقط العبارة عند إغلاق مفتاح الإدراج', () => {
    const s = M.freshPartyState();
    s.saudiId = '1234567890';
    assert.equal(M.buildIdentityClause(s, false), '');
    s.entityType = 'شركة';
    s.crNum = '55555';
    assert.equal(M.buildIdentityClause(s, false), '');
});

test('buildPartyClause: إغلاق المفتاح يحذف اسم الطرف وهويته ويُبقي لقبه', () => {
    const state = readyState();
    state.includePartyDataInText = false;
    assert.equal(M.buildPartyClause(state, 'plaintiff'), 'حضر المدعي أصالة');
    state.plaintiff.gender = 'ف';
    assert.equal(M.buildPartyClause(state, 'plaintiff'), 'حضرت المدعية أصالة');
    state.defendant.gender = 'ف';
    assert.equal(M.buildPartyClause(state, 'defendant'), 'حضرت المدعى عليها أصالة');
});

test('buildPartyClause: إغلاق المفتاح مع التعدد يُبقي ترقيم الأطراف', () => {
    const state = readyState({ includePartyDataInText: false });
    state.extraPlaintiffs.push(M.freshExtraParty());
    assert.equal(M.buildPartyClause(state, 'plaintiff'), 'حضر المدعي الأول أصالة');
});

test('buildPartyClause: إغلاق المفتاح يُبقي غياب الطرف وتبليغه بلا اسم', () => {
    const state = readyState({ includePartyDataInText: false });
    state.defendant.attendance = 'لم يحضر';
    state.defendant.tabligh = '456';
    assert.equal(
        M.buildPartyClause(state, 'defendant'),
        'لم يحضر المدعى عليه رغم تبليغه بالجلسة وفق مهمّة التبليغ رقم (456)'
    );
});

test('buildPartyClause: إغلاق المفتاح لا يمسّ بيانات الوكيل', () => {
    const state = readyState({ includePartyDataInText: false });
    state.defendant.attendance = 'تمثيل';
    state.defendant.agentName = 'خالد';
    state.defendant.agentId = '3000000003';
    state.defendant.wakalaNum = '777';
    state.defendant.licenseNum = '99';
    const clause = M.buildPartyClause(state, 'defendant');
    assert.match(clause, /^حضر عن المدعى عليه وكيله المحامي خالد/);
    assert.match(clause, /بموجب الوكالة الشرعية رقم \(777\)/);
    assert.match(clause, /للمادَّة \(51\/3\) من نظام المرافعات الشرعية/);
    assert.match(clause, /رخصة مزاولة المحاماة رقم \(99\)/);
    assert.match(clause, /بموجب الهوية رقم \(3000000003\)/);
});

test('buildPartyClause: إغلاق المفتاح مع وكيل مرافق يُبقي بيانات الوكالة وحدها', () => {
    const state = readyState({ includePartyDataInText: false });
    state.plaintiff.hasAccompanyingAgent = true;
    state.plaintiff.agentName = 'خالد';
    state.plaintiff.agentId = '3000000003';
    state.plaintiff.wakalaNum = '777';
    state.plaintiff.licenseNum = '99';
    const clause = M.buildPartyClause(state, 'plaintiff');
    assert.match(clause, /^حضر المدعي أصالة، وحضر معه وكيله المحامي خالد/);
    assert.ok(!clause.includes('الهوية الوطنية'));
    assert.match(clause, /بموجب الوكالة الشرعية رقم \(777\)/);
});

test('buildExtraClause: إغلاق المفتاح يحذف اسم الطرف الإضافي وهويته', () => {
    const p = M.freshExtraParty();
    p.name = 'ناصر';
    p.saudiId = '1010101010';
    assert.equal(M.buildExtraClause('plaintiff', 0, p, false), 'حضر المدعي الثاني أصالة');
    p.gender = 'ف';
    assert.equal(M.buildExtraClause('defendant', 0, p, false), 'حضرت المدعى عليها الثانية أصالة');
});

test('composeMinutes: إغلاق المفتاح يحذف الأسماء والهوية من نص الضبط للأطراف كافة', () => {
    const state = readyState({ includePartyDataInText: false });
    const extra = M.freshExtraParty();
    extra.name = 'ناصر';
    extra.saudiId = '1010101010';
    state.extraPlaintiffs.push(extra);
    const text = M.composeMinutes(state);
    assert.ok(!text.includes('الهوية الوطنية'));
    assert.ok(!text.includes('الإقامة النظامية'));
    assert.ok(!text.includes('سعد'));
    assert.ok(!text.includes('ناصر'));
    assert.match(text, /حضر المدعي الأول أصالة/);
    assert.match(text, /حضر المدعي الثاني أصالة/);
});

test('collectWarnings: إغلاق المفتاح يُسقط تحذيرات الاسم والهوية معًا', () => {
    const state = readyState({ includePartyDataInText: false });
    state.plaintiff.name = '';
    state.plaintiff.saudiId = '';
    state.defendant.name = '';
    state.defendant.saudiId = '';
    state.extraPlaintiffs.push(M.freshExtraParty());   // بلا اسم ولا هوية
    const w = M.collectWarnings(state);
    assert.deepEqual(w.filter(x => x.includes('اسم المدعي') || x.includes('اسم المدعى عليه')), []);
    assert.deepEqual(w.filter(x => x.includes('هوية') || x.includes('إقامة') || x.includes('السجل التجاري')), []);
});

test('collectWarnings: تفعيل المفتاح يعيد المطالبة باسم الطرف وهويته', () => {
    const state = readyState();
    state.plaintiff.name = '';
    state.plaintiff.saudiId = '';
    const w = M.collectWarnings(state);
    assert.ok(w.some(x => x.includes('لم يُدخل اسم المدعي')));
    assert.ok(w.some(x => x.includes('رقم هوية المدعي')));
});

test('buildPartyClause: حضور أصالة', () => {
    const state = readyState();
    const clause = M.buildPartyClause(state, 'plaintiff');
    assert.equal(clause, 'حضر المدعي سعد أصالة، بموجب الهوية الوطنية رقم (1000000001)');
});

test('buildPartyClause: وكيل محامٍ يتضمن المادة (51/3) ورخصة المحاماة', () => {
    const state = readyState();
    state.defendant.attendance = 'تمثيل';
    state.defendant.agentName = 'خالد';
    state.defendant.agentId = '3000000003';
    state.defendant.wakalaNum = '777';
    state.defendant.licenseNum = '99';
    const clause = M.buildPartyClause(state, 'defendant');
    assert.match(clause, /^حضر عن المدعى عليه فهد وكيله المحامي خالد/);
    assert.match(clause, /للمادَّة \(51\/3\) من نظام المرافعات الشرعية/);
    assert.match(clause, /رخصة مزاولة المحاماة رقم \(99\)/);
});

test('buildPartyClause: ممثل نظامي بصفته', () => {
    const state = readyState();
    state.plaintiff.attendance = 'تمثيل';
    state.plaintiff.repType = 'ممثل';
    state.plaintiff.repName = 'ماجد';
    state.plaintiff.repNum = 'المدير العام';
    const clause = M.buildPartyClause(state, 'plaintiff');
    assert.equal(clause, 'حضر عن المدعي سعد ممثله النظامي ماجد، بصفة: المدير العام');
});

test('buildExtraClause: طرف إضافي حاضر وممثل وغائب', () => {
    const p = M.freshExtraParty();
    p.name = 'ناصر';
    p.saudiId = '1010101010';
    assert.equal(M.buildExtraClause('plaintiff', 0, p), 'حضر المدعي الثاني ناصر أصالة، بموجب الهوية الوطنية رقم (1010101010)');
    p.attendance = 'تمثيل';
    p.repName = 'بدر';
    p.repNum = '123';
    assert.equal(M.buildExtraClause('plaintiff', 0, p), 'حضر عن المدعي الثاني ناصر وكيله بدر بموجب الوكالة الشرعية رقم (123)');
    p.attendance = 'لم يحضر';
    assert.equal(M.buildExtraClause('plaintiff', 0, p), 'لم يحضر المدعي الثاني ناصر');
});

test('buildExtraClause: بطاقة هوية الطرف الإضافي كالطرف الأول — سعودي ومقيم', () => {
    const p = M.freshExtraParty();
    p.name = 'ناصر';
    // لم تُعبّأ الهوية بعد: نقاط لا فراغ، كالطرف الأول
    assert.equal(M.buildExtraClause('plaintiff', 0, p), `حضر المدعي الثاني ناصر أصالة، بموجب الهوية الوطنية رقم (${M.MINUTES_PLACEHOLDER})`);

    p.nationalityType = 'غير ذلك';
    p.foreignNationality = 'مصري';
    p.iqamaNum = '2020202020';
    assert.equal(
        M.buildExtraClause('plaintiff', 0, p),
        'حضر المدعي الثاني ناصر أصالة، مصري الجنسية، بموجب الإقامة النظامية رقم (2020202020)'
    );

    const f = M.freshExtraParty();
    f.name = 'نورة';
    f.gender = 'ف';
    f.saudiId = '1122334455';
    assert.equal(M.buildExtraClause('defendant', 0, f), 'حضرت المدعى عليها الثانية نورة أصالة، بموجب الهوية الوطنية رقم (1122334455)');
});

test('collectWarnings: نقص هوية الطرف الإضافي الحاضر أصالة', () => {
    const state = readyState();
    const p = M.freshExtraParty();
    p.name = 'ناصر';
    state.extraPlaintiffs.push(p);
    assert.ok(M.collectWarnings(state).some(x => /رقم هوية المدعي الثاني غير مكتمل/.test(x)));

    p.saudiId = '1010101010';
    assert.ok(!M.collectWarnings(state).some(x => /رقم هوية المدعي الثاني/.test(x)));

    p.nationalityType = 'غير ذلك';
    const w = M.collectWarnings(state);
    assert.ok(w.some(x => /لم تُحدَّد جنسية المدعي الثاني/.test(x)));
    assert.ok(w.some(x => /رقم إقامة المدعي الثاني غير مكتمل/.test(x)));

    p.attendance = 'لم يحضر';
    assert.ok(!M.collectWarnings(state).some(x => /جنسية المدعي الثاني|إقامة المدعي الثاني/.test(x)));
});

test('buildShatbText: الغياب الأول — شطب للمرة الأولى والمادة (55)', () => {
    const state = readyState();
    state.plaintiff.attendance = 'لم يحضر';
    state.plaintiff.tabligh = '456';
    const text = M.buildShatbText(state);
    assert.match(text, /شطب الدَّعوى للمرَّة الأُولى/);
    assert.match(text, /خلال \(60\) يوماً/);
    assert.match(text, /بالمادَّة \(55\) من نظام المرافعات الشَّرعيَّة/);
});

test('buildShatbText: الجلسة الحضورية تحذف رابط غرفة الاتصال المرئي', () => {
    const state = readyState();
    state.opening.mode = M.SESSION_MODES.IN_PERSON;
    state.plaintiff.attendance = 'لم يحضر';
    state.plaintiff.tabligh = '456';
    const text = M.buildShatbText(state);
    assert.ok(!text.includes('غرفة الاتصال المرئي'));
    assert.ok(!text.includes('بالاتصال المرئي'));
    assert.match(text, /حضورياً بمقر المحكمة/);
    assert.match(text, /عبر الوسائل الإلكترونية وفق مهمّة التبليغ رقم \( 456 \)/);
    assert.match(text, /شطب الدَّعوى للمرَّة الأُولى/);
});

test('buildShatbText: الجلسة المرئية المختصرة تبقي فقرة التبليغ المرئي كما هي', () => {
    const state = readyState();
    state.plaintiff.attendance = 'لم يحضر';
    state.plaintiff.tabligh = '456';
    const text = M.buildShatbText(state);
    assert.match(text, /- بالاتصال المرئي -/);
    assert.match(text, /برابط الدخول لغرفة الاتصال المرئي/);
});

test('buildSpecialCaseText: صيغة «حضر بالنظام ولم يحضر» تتبع طريقة الانعقاد', () => {
    const state = readyState();
    state.plaintiff.specialCase = 'systemNoVideo';
    assert.match(M.buildSpecialCaseText(state, 'plaintiff'), /لم يحضر الجلسة عبر الاتصال المرئي/);
    state.opening.mode = M.SESSION_MODES.IN_PERSON;
    assert.match(M.buildSpecialCaseText(state, 'plaintiff'), /لم يحضر الجلسة بمقر المحكمة/);
});

test('buildShatbText: الغياب الثاني — اعتبار الدعوى كأن لم تكن', () => {
    const state = readyState();
    state.plaintiff.attendance = 'لم يحضر';
    state.plaintiff.occurrence = 2;
    const text = M.buildShatbText(state);
    assert.match(text, /الأسباب:/);
    assert.match(text, /الحكم:\nحكمت الدائرة باعتبار الدعوى كأن لم تكن/);
    assert.match(text, /ثلاثين يومًا للاعتراض/);
});

test('buildOathBlock: لا أرغب / أرغب مع مدعى عليه غائب', () => {
    const decline = M.buildOathBlock('م', 'المدعى عليه', 'المدعي', false, true, 'حاضر', false, 'موكلي');
    assert.match(decline, /لا أرغب/);
    const absent = M.buildOathBlock('م', 'المدعى عليه', 'المدعي', true, false, 'غائب', false, 'موكلي');
    assert.match(absent, /وأطلب يمين المدعى عليه/);
    assert.match(absent, /للمادة \(التاسعة والتسعين\) من نظام الإثبات/);
    assert.match(absent, /عُدَّ ناكلاً/);
});

const SAMPLE_WITNESS = { name: 'صالح', age: '30', job: 'موظف', residence: 'الرياض', relation: 'جار', interest: 'لا مصلحة', testimony: 'المبلغ في ذمة المدعى عليه' };

function witnessCtx(overrides = {}) {
    return Object.assign({
        speakerGender: 'م', speakerSuffix: 'ه',
        presenterLabel: 'المدعي', presenterGender: 'م',
        opposingLabel: 'المدعى عليه', opposingPresent: true,
        tazkiya: 'none', tazkiyaNames: '',
        objection: 'لا', objectionText: ''
    }, overrides);
}

test('buildWitnessSection: يتضمن مادتي (71) و(78) وتحليف الشاهد وبياناته', () => {
    const text = M.buildWitnessSection([SAMPLE_WITNESS], witnessCtx());
    assert.match(text, /\(الحادية والسبعين\) و\(الثامنة والسبعين\)/);
    assert.match(text, /جرى سماع كل شاهد على انفراد/);
    assert.match(text, /اسمي الكامل: \( صالح \)/);
    assert.match(text, /ثم جرى تحليفه اليمين على أن يشهد بالحق، فحلف/);
    assert.match(text, /أشهد بالله العظيم أن \( المبلغ في ذمة المدعى عليه \)/);
});

test('buildWitnessSection: مطابقة التذكير والتأنيث في مقدّم الشهود والمتكلم', () => {
    const masc = M.buildWitnessSection([SAMPLE_WITNESS], witnessCtx());
    assert.match(masc, /فقرر قائلاً: نعم/);
    assert.match(masc, /ثم أحضر المدعي للشهادة/);

    const fem = M.buildWitnessSection([SAMPLE_WITNESS], witnessCtx({
        speakerGender: 'ف', speakerSuffix: 'ها', presenterLabel: 'المدعية', presenterGender: 'ف'
    }));
    assert.match(fem, /فقررت قائلة: نعم/);
    assert.match(fem, /ثم أحضرت المدعية للشهادة/);
    assert.ok(!fem.includes('قررت قائلاً'));
    assert.ok(!fem.includes('أحضر المدعية'));
});

test('buildWitnessSection: عرض الشهود على الخصم لسؤاله عن مطعنه', () => {
    const none = M.buildWitnessSection([SAMPLE_WITNESS], witnessCtx());
    assert.match(none, /وبعرض الشهود وشهادتهم على المدعى عليه قرر قائلاً: لا مطعن لي فيهم/);

    const objecting = M.buildWitnessSection([SAMPLE_WITNESS], witnessCtx({ objection: 'نعم', objectionText: 'الشاهد شريك للمدعي' }));
    assert.match(objecting, /قرر قائلاً: الشاهد شريك للمدعي/);

    // الخصم الغائب لا يُسأل عن المطعن
    const absent = M.buildWitnessSection([SAMPLE_WITNESS], witnessCtx({ opposingPresent: false }));
    assert.ok(!absent.includes('وبعرض الشهود وشهادتهم'));
});

test('buildWitnessSection: تعديل الشهود وتزكيتهم', () => {
    const delayed = M.buildWitnessSection([SAMPLE_WITNESS], witnessCtx({ tazkiya: 'delay' }));
    assert.match(delayed, /هل لديه معدِّلون للشهود\؟ فأجاب قائلاً: أطلب إمهالي/);

    const presented = M.buildWitnessSection([SAMPLE_WITNESS], witnessCtx({ tazkiya: 'presented', tazkiyaNames: 'محمد وعبدالله' }));
    assert.match(presented, /ثم أحضر المدعي المعدِّلين: \( محمد وعبدالله \)/);
    assert.match(presented, /عدول ثقات مقبولو الشهادة/);
});

// ==================== الترتيب الإجرائي ====================

test('التسلسل: عرض الدعوى على المدعى عليه يسبق سؤال المدعي عن بينته', () => {
    const text = M.composeMinutes(readyState());
    const answerAt = text.indexOf('وبعرض دعوى المدعي على المدعى عليه');
    const evidenceAt = text.indexOf('عن بينته');
    assert.ok(answerAt > -1 && evidenceAt > -1);
    assert.ok(answerAt < evidenceAt, 'جواب المدعى عليه يجب أن يسبق سؤال المدعي عن البينة');
});

test('التسلسل: الإنكار يُتبع بتكليف المدعي بالبينة', () => {
    const text = M.composeMinutes(readyState());
    assert.match(text, /ولمَّا كانت إجابة المدعى عليه إنكاراً لما جاء في الدعوى، وأن البينة على المدعي، فقد جرى تكليف المدعي بإحضار بينته/);
});

test('التسلسل: الإقرار يُسقط سؤال البينة واليمين', () => {
    const state = readyState();
    state.claim.defendantStance = 'إقرار';
    state.claim.defendantResponseText = 'أقر بما جاء في الدعوى';
    state.claim.requestOath = true;
    const text = M.composeMinutes(state);
    assert.match(text, /فلا موجب لتكليف المدعي بالبينة؛ إذ إنما تُطلب البينة عند الإنكار/);
    assert.ok(!text.includes('عن بينته'));
    assert.ok(!text.includes('وأطلب يمين'));
    assert.match(text, /قفل باب المرافعة/);
});

test('التسلسل: الدفع الشكلي يُعرض على المدعي وقد يُوقف النظر في الموضوع', () => {
    const state = readyState();
    state.claim.defendantStance = 'دفع شكلي';
    state.claim.formalPleaText = 'أدفع بعدم الاختصاص المكاني';
    state.claim.plaintiffReplyText = 'المحكمة مختصة لأن التنفيذ بالرياض';
    state.claim.answeredOnMerits = 'لا';
    const text = M.composeMinutes(state);
    assert.match(text, /عن دفعه الشكلي قرر قائلاً: أدفع بعدم الاختصاص المكاني/);
    assert.match(text, /وبعرض هذا الدفع على المدعي أجاب قائلاً: المحكمة مختصة/);
    assert.match(text, /النظر في الدفع الشكلي قبل الخوض في موضوع الدعوى/);
    assert.ok(!text.includes('عن بينته'));

    state.claim.answeredOnMerits = 'نعم';
    assert.match(M.composeMinutes(state), /عن بينته/);
});

test('التسلسل: غياب المدعى عليه يُبقي سؤال البينة مباشرة بلا عرض', () => {
    const state = readyState();
    state.defendant.attendance = 'لم يحضر';
    state.defendant.tabligh = '789';
    const text = M.composeMinutes(state);
    assert.match(text, /عن بينته/);
    assert.ok(!text.includes('وبعرض دعوى المدعي على'));
    assert.ok(!text.includes('جرى تكليف المدعي بإحضار بينته'));
});

// ==================== بينة المدعى عليه ====================

test('buildDefendantEvidenceText: لا تُدرج إلا عند طلبها ولا مع غياب المدعى عليه', () => {
    const state = readyState();
    assert.equal(M.buildDefendantEvidenceText(state), '');

    state.claim.askDefendantEvidence = 'نعم';
    assert.match(M.buildDefendantEvidenceText(state), /وبسؤاله عن بينته/);

    state.defendant.attendance = 'لم يحضر';
    assert.equal(M.buildDefendantEvidenceText(state), '');
});

test('buildDefendantEvidenceText: الدفوع والبينة وشهود المدعى عليه', () => {
    const state = readyState();
    state.claim.askDefendantEvidence = 'نعم';
    state.claim.defendantPleasText = 'أدفع بالوفاء';
    state.claim.defendantEvidence.choice = 'has';
    state.claim.defendantEvidence.items = ['إيصال استلام', 'شهادة شهود'];
    state.claim.defendantEvidence.witnesses = [SAMPLE_WITNESS];
    const text = M.buildDefendantEvidenceText(state);
    assert.match(text, /وبسؤال المدعى عليه عن دفوعه قرر قائلاً: أدفع بالوفاء/);
    assert.match(text, /بينتي هي: إيصال استلام، وشهادة شهود/);
    assert.match(text, /ثم أحضر المدعى عليه للشهادة الشاهد الأول/);
    // المطعن يُعرض على الخصم الآخر وهو المدعي
    assert.match(text, /وبعرض الشهود وشهادتهم على المدعي قرر/);
});

// ==================== الجلسة التالية ====================

test('composeMinutes: سماع بينة المدعي وشهوده في الجلسة التالية', () => {
    const state = readyState();
    state.sessionType = 'previous';
    state.followUp.plaintiffEvidence = true;
    state.claim.evidenceChoice = 'has';
    state.claim.evidenceItems = ['شهادة شهود'];
    state.claim.hasMoreEvidence = 'لا';
    state.claim.witnesses = [SAMPLE_WITNESS];
    const text = M.composeMinutes(state);
    assert.match(text, /وتنفيذاً لما تقرر في الجلسة السابقة من تكليف المدعي بالبينة/);
    assert.match(text, /ثم أحضر المدعي للشهادة الشاهد الأول/);
    // نص الدعوى لا يُعاد عرضه في الجلسة التالية
    assert.ok(!text.includes('وبالاطلاع على دعوى'));
});

test('composeMinutes: سماع دفوع المدعى عليه وبينته في الجلسة التالية', () => {
    const state = readyState();
    state.sessionType = 'previous';
    state.followUp.defendantEvidence = true;
    state.claim.defendantPleasText = 'أدفع بالإبراء';
    const text = M.composeMinutes(state);
    assert.match(text, /ثم انتقلت الدائرة لسماع دفوع المدعى عليه وبينته/);
    assert.match(text, /أدفع بالإبراء/);
});

// ==================== التوليد الكامل ====================

test('composeMinutes: جلسة تحضيرية مكتملة بطرفين حاضرين', () => {
    const text = M.composeMinutes(readyState());
    assert.match(text, /^فلديّ أنا فلان بن فلان القاضي بمحكمة الرياض/);
    assert.match(text, /حضر المدعي سعد أصالة/);
    assert.match(text, /حضر المدعى عليه فهد أصالة/);
    assert.match(text, /وبالاطلاع على دعوى المدعي وجدت نصها: "أطالب بمبلغ مئة ألف ريال"/);
    assert.match(text, /وبعرض دعوى المدعي على المدعى عليه أجاب قائلاً: ما ذكره المدعي غير صحيح/);
    assert.match(text, /قفل باب المرافعة/);
    assert.match(text, /وختمت الجلسة عند الساعة التاسعة والنصف صباحًا\.$/);
    // لا موضع ناقص في حالة مكتملة
    assert.ok(!text.includes(M.MINUTES_PLACEHOLDER));
});

test('composeMinutes: غياب المدعي المتبلّغ يولّد محضر شطب', () => {
    const state = readyState();
    state.plaintiff.attendance = 'لم يحضر';
    state.plaintiff.tabligh = '456';
    const text = M.composeMinutes(state);
    assert.match(text, /شطب الدَّعوى/);
    assert.match(text, /وختمت الجلسة/);
});

test('composeMinutes: مدعى عليه لم يتبلّغ — رفع الجلسة لإعادة التبليغ', () => {
    const state = readyState();
    state.defendant.attendance = 'لم يحضر';
    state.defendant.notifyStatus = 'لم يتبلغ';
    const text = M.composeMinutes(state);
    assert.match(text, /لم يحضر المدعى عليه، ولم يتبلّغ بالجلسة، وعليه رُفعت الجلسة لإعادة تبليغه بحسب حاله/);
    assert.ok(!text.includes('وبالاطلاع على دعوى'));
});

test('composeMinutes: حالة استثنائية تستبدل المحضر بالكامل', () => {
    const state = readyState();
    state.plaintiff.specialCase = 'systemNoVideo';
    const text = M.composeMinutes(state);
    assert.match(text, /تبيّن حضور المدعي في النظام الإلكتروني/);
    assert.ok(!text.includes('وبالاطلاع على دعوى'));
});

test('composeMinutes: جلسة منظورة سابقًا بقاضٍ مختلف تتضمن المادة (167) والمصادقة', () => {
    const state = readyState();
    state.sessionType = 'previous';
    state.sameJudge = 'لا';
    const text = M.composeMinutes(state);
    assert.match(text, /المادة السابعة والستون بعد المائة/);
    assert.match(text, /نصادق على ما ورد في الجلسات السابقة/);
});

test('composeMinutes: يمين موجهة لمدعى عليه غائب لا تقفل باب المرافعة', () => {
    const state = readyState();
    state.defendant.attendance = 'لم يحضر';
    state.defendant.tabligh = '789';
    state.claim.requestOath = true;
    const text = M.composeMinutes(state);
    assert.match(text, /تحديد جلسة قادمة لإبلاغه بالحضور لأدائها/);
    assert.ok(!text.includes('قفل باب المرافعة'));
});

test('composeMinutes: غياب المدعى عليه عن جلسة أداء اليمين — نكول', () => {
    const state = readyState();
    state.sessionType = 'previous';
    state.defendant.attendance = 'لم يحضر';
    state.defendant.oathAbsence = 'نعم';
    state.defendant.oathTablighNum = '321';
    const text = M.composeMinutes(state);
    assert.match(text, /المخصصة لأداء اليمين على نفي الدعوى/);
    assert.match(text, /عُدّ ناكلاً/);
    assert.match(text, /قفل باب المرافعة/);
});

// ==================== فحص الاكتمال ====================

test('collectWarnings: حالة مكتملة بلا تحذيرات', () => {
    assert.deepEqual(M.collectWarnings(readyState()), []);
});

test('collectWarnings: يرصد الحقول الناقصة', () => {
    const state = M.freshMinutesState();
    state.sessionType = 'new';
    const w = M.collectWarnings(state);
    assert.ok(w.some(x => x.includes('اسم القاضي')));
    assert.ok(w.some(x => x.includes('نص الدعوى')));
    // اسم الطرف لا يُطالب به إلا إذا كان سيُدرج في نص الضبط
    assert.ok(!w.some(x => x.includes('اسم المدعي')));
    state.includePartyDataInText = true;
    assert.ok(M.collectWarnings(state).some(x => x.includes('اسم المدعي')));
});

test('collectWarnings: محضر الشطب لا يطالب ببيانات المدعى عليه', () => {
    const state = readyState();
    state.plaintiff.attendance = 'لم يحضر';
    state.plaintiff.tabligh = '456';
    state.defendant = M.freshPartyState(); // بيانات المدعى عليه فارغة
    const w = M.collectWarnings(state);
    assert.deepEqual(w.filter(x => x.includes('المدعى عليه')), []);
});

// ==================== الأسباب والحكم والإفهام ====================

function rulingState() {
    const s = readyState();
    s.ruling.pronounce = 'نعم';
    s.ruling.reasonsText = 'فبناء على ما تقدم من الدعوى والإجابة';
    s.ruling.rulingText = 'حكمت الدائرة بإلزام المدعى عليه بدفع المبلغ';
    return s;
}

test('findTemplate: الوصول لمكتبة النماذج في data/templates.js', () => {
    assert.ok(M.findTemplate('إفهامات بعد النطق', 'ف1'));
    assert.ok(M.findTemplate('إفهامات بعد النطق', 'واجب التدقيق'));
    assert.equal(M.findTemplate('تصنيف غير موجود', 'ف1'), null);
    assert.ok(M.templatesOfCategory('أسباب الحكم').length > 100);
    assert.equal(M.templatesOfCategory('تصنيف غير موجود').length, 0);
});

test('noticeKindFor: يُشتق نوع الإفهام من قيمة المطالبة', () => {
    const state = rulingState();
    state.ruling.claimValue = '45000';
    assert.equal(M.noticeKindFor(state), 'final');
    state.ruling.claimValue = String(M.YASEERA_CLAIM_LIMIT);
    assert.equal(M.noticeKindFor(state), 'final');
    state.ruling.claimValue = '50001';
    assert.equal(M.noticeKindFor(state), 'appealable');
    // بلا قيمة يُحتاط بالقابلية للاستئناف
    state.ruling.claimValue = '';
    assert.equal(M.noticeKindFor(state), 'appealable');
    // التحديد اليدوي يتقدم على الاشتقاق
    state.ruling.noticeKind = 'sulh';
    assert.equal(M.noticeKindFor(state), 'sulh');
});

test('buildNoticeText: النص مأخوذ من مكتبة النماذج لا مكتوب هنا', () => {
    const state = rulingState();
    state.ruling.claimValue = '10000';
    assert.equal(M.buildNoticeText(state), M.findTemplate('إفهامات بعد النطق', 'ف1'));

    state.ruling.claimValue = '900000';
    assert.equal(M.buildNoticeText(state), M.findTemplate('إفهامات بعد النطق', 'ف2'));

    // حضور الطرفين وكالةً يستدعي صيغة إفهام الوكلاء (المادة 165)
    state.plaintiff.attendance = 'تمثيل';
    state.defendant.attendance = 'تمثيل';
    assert.equal(M.buildNoticeText(state), M.findTemplate('إفهامات بعد النطق', 'ف3'));

    state.ruling.noticeKind = 'sulh';
    assert.equal(M.buildNoticeText(state), M.findTemplate('إفهامات بعد النطق', 'صلح1'));
});

test('buildMandatoryReviewText: نصا التدقيق الوجوبي من المكتبة', () => {
    assert.equal(M.buildMandatoryReviewText({ mandatoryReview: 'لا' }), '');
    assert.equal(M.buildMandatoryReviewText({ mandatoryReview: 'نعم' }), M.findTemplate('إفهامات بعد النطق', 'واجب التدقيق'));
    assert.equal(M.buildMandatoryReviewText({ mandatoryReview: 'إعادة' }), M.findTemplate('إفهامات بعد النطق', 'واجب التدقيق /إعادة القضية'));
});

test('composeMinutes: مرحلة الحكم تُلحق بالضبط بعد قفل باب المرافعة', () => {
    const state = rulingState();
    state.ruling.claimValue = '900000';
    state.ruling.mandatoryReview = 'نعم';
    const text = M.composeMinutes(state);
    const closeAt = text.indexOf('قفل باب المرافعة');
    const reasonsAt = text.indexOf('\n\nالأسباب:\n');
    const rulingAt = text.indexOf('\n\nالحكم:\n');
    const endAt = text.indexOf('وختمت الجلسة');
    assert.ok(closeAt < reasonsAt && reasonsAt < rulingAt && rulingAt < endAt);
    assert.match(text, /وهذا الحكم حضوري في حق طرفي الدعوى/);
    assert.match(text, /قابل للاستئناف/);
    assert.match(text, /واجب التدقيق/);
    assert.match(text, /وختمت الجلسة عند الساعة التاسعة والنصف صباحًا\.$/);
});

test('composeMinutes: بلا نطق بالحكم لا تُلحق الأسباب', () => {
    const text = M.composeMinutes(readyState());
    assert.ok(!text.includes('الأسباب:'));
    assert.ok(!text.includes('الحكم:'));
});

test('composeMinutes: صفة الحكم غيابية في حق المدعى عليه', () => {
    const state = rulingState();
    state.ruling.presence = 'غيابي';
    assert.match(M.composeMinutes(state), /وهذا الحكم غيابي في حق المدعى عليه/);
});

test('collectWarnings: نواقص مرحلة الحكم', () => {
    const state = readyState();
    state.ruling.pronounce = 'نعم';
    const w = M.collectWarnings(state);
    assert.ok(w.some(x => x.includes('أسباب الحكم')));
    assert.ok(w.some(x => x.includes('منطوق الحكم')));
    assert.ok(w.some(x => x.includes('قيمة المطالبة')));
});

test('collectWarnings: نواقص بينة المدعى عليه وشهوده', () => {
    const state = readyState();
    state.claim.askDefendantEvidence = 'نعم';
    state.claim.defendantEvidence.choice = 'has';
    state.claim.defendantEvidence.items = ['شهادة شهود'];
    state.claim.defendantEvidence.witnesses = [M.freshWitness()];
    state.claim.defendantEvidence.tazkiya = 'presented';
    const w = M.collectWarnings(state);
    assert.ok(w.some(x => x.includes('لشاهد المدعى عليه الأول')));
    assert.ok(w.some(x => x.includes('أسماء معدِّلي شهود المدعى عليه')));
});

test('collectWarnings: الإقرار لا يطالب ببينة المدعي', () => {
    const state = readyState();
    state.claim.defendantStance = 'إقرار';
    state.claim.evidenceChoice = 'has'; // بيانات مهملة لا تُفحص
    const w = M.collectWarnings(state);
    assert.deepEqual(w.filter(x => x.includes('بينة المدعي')), []);
});

test('collectWarnings: الدفع الشكلي يستلزم نصه ورد المدعي', () => {
    const state = readyState();
    state.claim.defendantStance = 'دفع شكلي';
    const w = M.collectWarnings(state);
    assert.ok(w.some(x => x.includes('نص الدفع الشكلي')));
    assert.ok(w.some(x => x.includes('رد المدعي على الدفع الشكلي')));
});

test('collectWarnings: وكيل غير محامٍ خارج الدرجات الأربع', () => {
    const state = readyState();
    state.plaintiff.attendance = 'تمثيل';
    state.plaintiff.repIsLawyer = 'لا';
    state.plaintiff.agentName = 'خالد';
    state.plaintiff.agentId = '3000000003';
    state.plaintiff.wakalaNum = '777';
    state.plaintiff.kinship = 'أخرى';
    const w = M.collectWarnings(state);
    assert.ok(w.some(x => x.includes('خارج الدرجات الأربع')));
});

// ==================== بحث نماذج التسبيب ====================

test('normalizeArabicSearch: يسقط التشكيل ويوحّد الهمزات والتاء المربوطة', () => {
    assert.equal(M.normalizeArabicSearch('فبيانُ الدعوى أنّ المدعية'), 'فبيان الدعوي ان المدعيه');
    assert.equal(M.normalizeArabicSearch('  بيّنة   «واتساب»  '), 'بينه واتساب');
    assert.equal(M.normalizeArabicSearch('رقم ١٤٤٧'), 'رقم 1447');
    assert.equal(M.normalizeArabicSearch(null), '');
});

test('searchTemplates: مطابقة مطبَّعة بكلمات متعددة مع تقديم الكلمة المفتاحية', () => {
    const list = [
        { keyword: 'أجرة المثل', content: 'فبيانُ الدعوى أنّ المدعي يطالب بأجرة المثل' },
        { keyword: 'شيكات', content: 'الحكم بقيمة الشيكات وأجرةِ المثل معًا' },
        { keyword: 'دية', content: 'الحكم بالدية' }
    ];
    // بلا استعلام: التصنيف كاملاً
    assert.equal(M.searchTemplates(list, '').length, 3);
    // مكتوب بلا همزات ولا تشكيل، ومع ذلك يُطابق
    const hits = M.searchTemplates(list, 'اجره المثل');
    assert.deepEqual(hits.map(t => t.keyword), ['أجرة المثل', 'شيكات']);
    // كلمتان متباعدتان في النص تُطابقان معًا
    assert.deepEqual(M.searchTemplates(list, 'المثل شيكات').map(t => t.keyword), ['شيكات']);
    assert.deepEqual(M.searchTemplates(list, 'لا يوجد'), []);
});

test('curatedReasonsCount: يُحسب من المكتبة لا برقم ثابت', () => {
    const opener = M.CURATED_REASON_OPENER + ' أنّ المدعي';
    assert.equal(M.curatedReasonsCount([{ content: opener }, { content: opener }, { content: 'فبناء على' }]), 2);
    // لا نماذج منسّقة في المقدمة: يُرجع الاحتياطي محدودًا بطول القائمة
    assert.equal(M.curatedReasonsCount([{ content: 'فبناء على' }, { content: 'و لما' }]), 2);
});

test('نماذج التسبيب في المكتبة: التصنيف موجود ومقدمته منسّقة', () => {
    const all = M.templatesOfCategory('أسباب الحكم');
    assert.ok(all.length > 100);
    const curated = M.curatedReasonsCount(all);
    assert.ok(curated > 0 && curated <= all.length);
    // كل النماذج المنسّقة تبدأ بمطلعها الموحّد
    all.slice(0, curated).forEach(t => {
        assert.ok(M.normalizeArabicSearch(t.content).startsWith(M.normalizeArabicSearch(M.CURATED_REASON_OPENER)));
    });
    // البحث يجد نماذج محدَّثة رغم اختلاف التشكيل والهمزات في الكتابة
    assert.ok(M.searchTemplates(all, 'اجرة المثل').length > 0);
});
