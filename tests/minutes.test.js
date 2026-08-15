// اختبارات محرك مولّد محضر الجلسة في js/minutes.js
// التشغيل: npm test  (أو: node --test tests/)
const test = require('node:test');
const assert = require('node:assert/strict');

const M = require('../js/minutes.js');

// حالة جاهزة: جلسة تحضيرية، الطرفان حاضران أصالة ببيانات مكتملة
function readyState(overrides = {}) {
    const s = M.freshMinutesState();
    s.sessionType = 'new';
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

test('buildOpening: يتضمن القاضي والمحكمة وصيغة الاتصال المرئي والوقت', () => {
    const text = M.buildOpening({ judge: 'فلان', court: 'الرياض', hour: 8, minute: 30, period: 'ص' });
    assert.match(text, /^فلديّ أنا فلان القاضي بمحكمة الرياض افتتحتُ الجلسة عبر الاتصال المرئي/);
    assert.match(text, /قرار رئيس المجلس الأعلى للقضاء رقم \(17388\)/);
    assert.match(text, /عند الساعة الثامنة والنصف صباحًا،$/);
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
    assert.equal(M.buildExtraClause('plaintiff', 0, p), 'حضر المدعي الثاني ناصر أصالة');
    p.attendance = 'تمثيل';
    p.repName = 'بدر';
    p.repNum = '123';
    assert.equal(M.buildExtraClause('plaintiff', 0, p), 'حضر عن المدعي الثاني ناصر وكيله بدر بموجب الوكالة الشرعية رقم (123)');
    p.attendance = 'لم يحضر';
    assert.equal(M.buildExtraClause('plaintiff', 0, p), 'لم يحضر المدعي الثاني ناصر');
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

test('buildWitnessSection: يتضمن مادتي (71) و(78) وبيانات الشاهد', () => {
    const witnesses = [{ name: 'صالح', age: '30', job: 'موظف', residence: 'الرياض', relation: 'جار', interest: 'لا مصلحة', testimony: 'المبلغ في ذمة المدعى عليه' }];
    const text = M.buildWitnessSection(witnesses, 'م', 'ه', 'المدعي');
    assert.match(text, /\(الحادية والسبعين\) و\(الثامنة والسبعين\)/);
    assert.match(text, /اسمي الكامل: \( صالح \)/);
    assert.match(text, /أشهد بالله العظيم أن \( المبلغ في ذمة المدعى عليه \)/);
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
    assert.ok(w.some(x => x.includes('اسم المدعي')));
    assert.ok(w.some(x => x.includes('نص الدعوى')));
});

test('collectWarnings: محضر الشطب لا يطالب ببيانات المدعى عليه', () => {
    const state = readyState();
    state.plaintiff.attendance = 'لم يحضر';
    state.plaintiff.tabligh = '456';
    state.defendant = M.freshPartyState(); // بيانات المدعى عليه فارغة
    const w = M.collectWarnings(state);
    assert.deepEqual(w.filter(x => x.includes('المدعى عليه')), []);
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
