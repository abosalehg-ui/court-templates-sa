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
    s.opening = { judge: 'فلان بن فلان', court: 'الرياض' };
    s.closing = { hour: 9, minute: 30, period: 'ص' };
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

test('buildOpening: الافتتاح بلا ذكر الانعقاد، وصفة القاضي من حقل الاسم', () => {
    const text = M.buildOpening({ judge: 'فلان القاضي في الدائرة الأولى', court: 'الرياض' });
    assert.equal(text, 'لدي أنا فلان القاضي في الدائرة الأولى في المحكمة الرياض،');
    assert.ok(!text.includes('افتتحتُ الجلسة'));
    assert.ok(!text.includes('17388'));
});

test('buildOpening: لا يُذكر وقت الافتتاح في المتن (مثبت في ناجز)', () => {
    const text = M.buildOpening({ judge: 'فلان', court: 'الرياض' });
    assert.ok(!text.includes('الساعة'));
    assert.equal(M.freshMinutesState().opening.hour, undefined);
});

test('buildOpening: طريقة الانعقاد لا تُذكر في الافتتاح مهما كان الخيار', () => {
    [M.SESSION_MODES.VIDEO, M.SESSION_MODES.VIDEO_FULL, M.SESSION_MODES.IN_PERSON].forEach(mode => {
        const text = M.buildOpening({ judge: 'فلان', court: 'الرياض', mode });
        assert.equal(text, 'لدي أنا فلان في المحكمة الرياض،');
        assert.ok(!text.includes('المرئي'));
        assert.ok(!text.includes('17388'));
    });
});

test('closingTimeParts: الوقت المُدخل، ومع الحالات القديمة يُشتق من الافتتاح + 30 دقيقة', () => {
    assert.deepEqual(M.closingTimeParts({ closing: { hour: 10, minute: 15, period: 'ص' } }), [10, 15, 'ص']);
    assert.deepEqual(M.closingTimeParts({ opening: { hour: 9, minute: 0, period: 'ص' } }), [9, 30, 'ص']);
    assert.deepEqual(M.freshMinutesState().closing, { hour: 8, minute: 30, period: 'ص' });
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

test('buildPartyClause: الحضور أصالةً بلا بيانات لا يُثبت في المتن', () => {
    const state = readyState();
    state.includePartyDataInText = false;
    assert.equal(M.buildPartyClause(state, 'plaintiff'), '');
    state.plaintiff.gender = 'ف';
    assert.equal(M.buildPartyClause(state, 'plaintiff'), '');
    state.defendant.gender = 'ف';
    assert.equal(M.buildPartyClause(state, 'defendant'), '');
});

test('buildPartyClause: إسقاط فقرة الأصالة يشمل حال تعدد الأطراف', () => {
    const state = readyState({ includePartyDataInText: false });
    state.extraPlaintiffs.push(M.freshExtraParty());
    assert.equal(M.buildPartyClause(state, 'plaintiff'), '');
});

test('buildPartyClause: الوكيل المرافق يُبقي فقرة الحضور أصالةً ولو أُغلق المفتاح', () => {
    const state = readyState({ includePartyDataInText: false });
    state.plaintiff.hasAccompanyingAgent = true;
    state.plaintiff.agentName = 'خالد';
    state.plaintiff.agentId = '1020304050';
    state.plaintiff.wakalaNum = '123';
    assert.match(M.buildPartyClause(state, 'plaintiff'), /^حضر المدعي أصالة، وحضر معه خالد، بصفته وكيلاً/);
});

test('buildPartyClause: إغلاق المفتاح يُبقي غياب الطرف وتبليغه بلا اسم', () => {
    const state = readyState({ includePartyDataInText: false });
    state.defendant.attendance = 'لم يحضر';
    state.defendant.tabligh = '456';
    assert.equal(
        M.buildPartyClause(state, 'defendant'),
        'قد تبلَّغ المدعى عليه ولم يحضر هو ولا من يمثله، بمهمة التبليغ رقم (456)، ولم يودع مذكرة بدفاعه بناء على ما قررته المادة الخامسة والأربعون من نظام المرافعات الشرعية'
    );
    // المطابقة النحوية للمدعى عليها
    state.defendant.gender = 'ف';
    assert.equal(
        M.buildPartyClause(state, 'defendant'),
        'قد تبلَّغت المدعى عليها ولم تحضر هي ولا من يمثلها، بمهمة التبليغ رقم (456)، ولم تودع مذكرة بدفاعها بناء على ما قررته المادة الخامسة والأربعون من نظام المرافعات الشرعية'
    );
});

test('composeMinutes: فقرة غياب المدعى عليه موضعها بعد رصد الدعوى ومصادقة المدعي', () => {
    const state = readyState({ includePartyDataInText: false });
    state.defendant.attendance = 'لم يحضر';
    state.defendant.tabligh = '456';
    const text = M.composeMinutes(state);
    assert.match(text, /صادق عليها\. وقد تبلَّغ المدعى عليه ولم يحضر هو ولا من يمثله، بمهمة التبليغ رقم \(456\)،/);
    // ولا تُذكر مرتين: أُخرجت من فقرات الحضور التي تلي الافتتاح
    assert.equal(text.match(/قد تبلَّغ المدعى عليه/g).length, 1);
    // وإضافة المدعي إن وُجدت سبقت فقرة الغياب
    state.claim.plaintiffAddition = true;
    state.claim.plaintiffAdditionText = 'أضيف كذا';
    assert.match(M.composeMinutes(state), /هكذا قدَّم\. وقد تبلَّغ المدعى عليه/);
});

test('composeMinutes: الجلسة المنظورة سابقًا تُبقي غياب المدعى عليه في فقرات الحضور', () => {
    const state = readyState({ includePartyDataInText: false });
    state.sessionType = 'previous';
    state.defendant.attendance = 'لم يحضر';
    state.defendant.tabligh = '456';
    assert.match(
        M.composeMinutes(state),
        /^لدي أنا فلان بن فلان في المحكمة الرياض، وقد تبلَّغ المدعى عليه ولم يحضر هو ولا من يمثله، بمهمة التبليغ رقم \(456\)،/
    );
});

test('buildPartyClause: إغلاق المفتاح لا يمسّ بيانات الوكيل', () => {
    const state = readyState({ includePartyDataInText: false });
    state.defendant.attendance = 'تمثيل';
    state.defendant.agentName = 'خالد';
    state.defendant.wakalaIssuer = 'كتابة العدل بالرياض';
    state.defendant.wakalaNum = '777';
    state.defendant.licenseNum = '99';
    assert.equal(
        M.buildPartyClause(state, 'defendant'),
        'حضر عن المدعى عليه خالد، بصفته وكيلاً بموجب الوكالة الصادرة من (كتابة العدل بالرياض)، برقم (777)، ورخصة مزاولة المحاماة رقم (99)'
    );
});

test('buildPartyClause: إغلاق المفتاح مع وكيل مرافق يُبقي بيانات الوكالة وحدها', () => {
    const state = readyState({ includePartyDataInText: false });
    state.plaintiff.hasAccompanyingAgent = true;
    state.plaintiff.agentName = 'خالد';
    state.plaintiff.agentId = '3000000003';
    state.plaintiff.wakalaNum = '777';
    state.plaintiff.licenseNum = '99';
    const clause = M.buildPartyClause(state, 'plaintiff');
    assert.match(clause, /^حضر المدعي أصالة، وحضر معه خالد، بصفته وكيلاً/);
    assert.ok(!clause.includes('الهوية الوطنية'));
    assert.match(clause, /برقم \(777\)/);
});

test('buildExtraClause: إغلاق المفتاح يُسقط فقرة الطرف الإضافي الحاضر أصالةً', () => {
    const p = M.freshExtraParty();
    p.name = 'ناصر';
    p.saudiId = '1010101010';
    assert.equal(M.buildExtraClause('plaintiff', 0, p, false), '');
    p.gender = 'ف';
    assert.equal(M.buildExtraClause('defendant', 0, p, false), '');
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
    // فقرات الحضور أصالةً تسقط كلها، فيتصل الافتتاح بالاطلاع على صحيفة الدعوى
    assert.ok(!text.includes('أصالة'));
    assert.match(text, /^لدي أنا فلان بن فلان في المحكمة الرياض، جرى الاطلاع على صحيفة الدعوى/);
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

test('buildPartyClause: صيغة الوكالة تحمل جهة الإصدار ورقمها ورخصة المحاماة', () => {
    const state = readyState();
    state.defendant.attendance = 'تمثيل';
    state.defendant.agentName = 'خالد';
    state.defendant.wakalaIssuer = 'كتابة العدل بالرياض';
    state.defendant.wakalaNum = '777';
    state.defendant.licenseNum = '99';
    const clause = M.buildPartyClause(state, 'defendant');
    assert.match(clause, /^حضر عن المدعى عليه فهد خالد، بصفته وكيلاً بموجب الوكالة الصادرة من \(كتابة العدل بالرياض\)، برقم \(777\)/);
    assert.match(clause, /رخصة مزاولة المحاماة رقم \(99\)/);
    // عبارة التحقق ورقم هوية الوكيل لم تعودا تُكتبان في المتن
    assert.ok(!clause.includes('51/3'));
    assert.ok(!clause.includes('الهوية'));
});

test('buildAgentCapacityPhrase: الوكيلة غير المحامية تُلحق بها صلة القرابة', () => {
    const s = M.freshPartyState();
    s.agentGender = 'ف';
    s.repIsLawyer = 'لا';
    s.kinship = 'الأخت';
    s.wakalaIssuer = 'الخدمات الإلكترونية بناجز';
    s.wakalaNum = '55';
    assert.equal(
        M.buildAgentCapacityPhrase(s, 'المدعية', { present: true }),
        'الحاضرة بصفتها وكيلةً بموجب الوكالة الصادرة من (الخدمات الإلكترونية بناجز)، برقم (55)، وتربطها بالمدعية صلة قرابة: الأخت'
    );
});

test('buildAgentCredentialsPhrase: وكيل الشركة لا تُذكر له قرابة', () => {
    const s = M.freshPartyState();
    s.repType = 'وكيل شركة';
    s.repIsLawyer = 'لا';
    s.kinship = 'الأخ';
    assert.equal(M.buildAgentCredentialsPhrase(s, 'المدعى عليه'), '');
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
    p.repIssuer = 'كتابة العدل بالرياض';
    p.repNum = '123';
    assert.equal(M.buildExtraClause('plaintiff', 0, p), 'حضر عن المدعي الثاني ناصر بدر، بصفته وكيلاً بموجب الوكالة الصادرة من (كتابة العدل بالرياض)، برقم (123)');
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

const SAMPLE_WITNESS = {
    name: 'صالح', nationality: 'سعودي', idNum: '1099887766', age: '30', job: 'موظف', residence: 'الرياض',
    relationPlaintiff: 'جار', relationDefendant: 'لا صلة', interest: 'لا مصلحة',
    testimony: 'المبلغ في ذمة المدعى عليه', phone: '0500000000'
};

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
    assert.match(text, /وبسؤال الشاهد عن بياناته وما لديه من شهادة\؟/);
    assert.match(text, /وأشهد لله تعالى بأن \( المبلغ في ذمة المدعى عليه \) هذا ما أشهد به هكذا أجاب/);
});

test('buildWitnessSection: صلة الشاهد بكل خصم على حدة', () => {
    const text = M.buildWitnessSection([SAMPLE_WITNESS], witnessCtx());
    assert.match(text, /وصلتي بالمدعي: \( جار \)، وصلتي بالمدعى عليه: \( لا صلة \)/);
    assert.ok(!text.includes('وعلاقتي بأطراف الدعوى هو'));
});

test('buildWitnessSection: جنسية الشاهد وهويته تُثبتان في الضبط دائمًا', () => {
    const text = M.buildWitnessSection([SAMPLE_WITNESS], witnessCtx());
    assert.match(text, /اسمي الكامل: \( صالح \)، وجنسيتي: \( سعودي \)، ورقم هويتي: \( 1099887766 \)، وتاريخ ميلادي/);
});

test('WITNESS_NATIONALITY_OPTIONS: قائمة الطرفين نفسها يتصدرها «سعودي»', () => {
    assert.equal(M.WITNESS_NATIONALITY_OPTIONS[0], 'سعودي');
    assert.equal(M.WITNESS_NATIONALITY_OPTIONS.length, M.NATIONALITY_OPTIONS.length + 1);
    assert.ok(M.WITNESS_NATIONALITY_OPTIONS.includes('مصري'));
    // قائمة الطرفين لا يدخلها «سعودي» لأن له مفتاحًا مستقلًا
    assert.ok(!M.NATIONALITY_OPTIONS.includes('سعودي'));
});

test('buildWitnessSection: رقم جوال الشاهد أرقامٌ بين معقوفتين بلا عبارة تعرّف به', () => {
    const text = M.buildWitnessSection([SAMPLE_WITNESS], witnessCtx());
    assert.match(text, /ومصلحتي في هذه الدعوى هي: \( لا مصلحة \) \[0500000000\]، وأشهد لله تعالى/);
    // لا عبارة تسبقه ولا تلحقه
    assert.ok(!text.includes('جوال'));
    assert.ok(!text.includes('هاتف'));
});

test('buildWitnessSection: لا معقوفتان فارغتان إذا لم يُدخل رقم الجوال', () => {
    const witness = Object.assign({}, SAMPLE_WITNESS, { phone: '' });
    const text = M.buildWitnessSection([witness], witnessCtx());
    assert.match(text, /ومصلحتي في هذه الدعوى هي: \( لا مصلحة \)، وأشهد لله تعالى/);
    assert.ok(!text.includes('['));
});

test('EVIDENCE_OPTIONS: شهادة الشهود ثانيةً بعد العقد', () => {
    assert.equal(M.EVIDENCE_OPTIONS[0], 'العقد');
    assert.equal(M.EVIDENCE_OPTIONS[1], 'شهادة شهود');
    // بقية الخيارات باقية بلا حذف ولا تكرار
    assert.equal(M.EVIDENCE_OPTIONS.length, new Set(M.EVIDENCE_OPTIONS).size);
    assert.equal(M.EVIDENCE_OPTIONS.length, 18);
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
    const answerAt = text.indexOf('وبعرضها على المدعى عليه');
    const evidenceAt = text.indexOf('عن بينته');
    assert.ok(answerAt > -1 && evidenceAt > -1);
    assert.ok(answerAt < evidenceAt, 'جواب المدعى عليه يجب أن يسبق سؤال المدعي عن البينة');
});

// ==================== مصادقة المدعي على صحيفة دعواه وإضافته عليها ====================

test('buildClaimEvidenceText: مصادقة المدعي تتبع صفة حضوره — أصالةً أو وكالة', () => {
    const state = readyState();
    assert.match(M.buildClaimEvidenceText(state), /وبعرضها على المدعي صادق عليها\./);

    state.plaintiff.attendance = 'تمثيل';
    state.plaintiff.repType = 'وكيل';
    state.plaintiff.agentName = 'خالد';
    assert.match(M.buildClaimEvidenceText(state), /وبعرضها على المدعي وكالة صادق عليها\./);

    // الفعل يتبع جنس المصادِق: الوكيلة تُصادق
    state.plaintiff.agentGender = 'ف';
    assert.match(M.buildClaimEvidenceText(state), /وبعرضها على المدعي وكالة صادقت عليها\./);

    // والممثل النظامي صفة لا وكالة، فيبقى اللقب مجردًا
    state.plaintiff.repType = 'ممثل';
    assert.match(M.buildClaimEvidenceText(state), /وبعرضها على المدعي صادق عليها\./);
});

test('buildPlaintiffAdditionClause: إضافة المدعي اختيارية وتلي المصادقة', () => {
    const state = readyState();
    assert.equal(M.buildPlaintiffAdditionClause(state), '');

    state.claim.plaintiffAddition = true;
    state.claim.plaintiffAdditionText = 'أضيف أن المبلغ مئة وعشرون ألفًا';
    assert.equal(
        M.buildPlaintiffAdditionClause(state),
        ' ثم قدم النص التالي: ( أضيف أن المبلغ مئة وعشرون ألفًا ) هكذا قدَّم.'
    );
    assert.match(M.buildClaimEvidenceText(state), /صادق عليها\. ثم قدم النص التالي: \( أضيف أن المبلغ مئة وعشرون ألفًا \) هكذا قدَّم\./);

    // المطابقة النحوية للمدعية
    state.plaintiff.gender = 'ف';
    assert.match(M.buildPlaintiffAdditionClause(state), /^ ثم قدمت النص التالي: \(.*\) هكذا قدَّمت\.$/);

    // وتركُ النص فارغًا يُرصد في التحذيرات ويُطبع نقاطًا
    state.claim.plaintiffAdditionText = '';
    assert.match(M.buildPlaintiffAdditionClause(state), /\( \.+ \)/);
    assert.ok(M.collectWarnings(state).some(x => x.includes('إضافة المدعي')));
});

// ==================== صور جواب المدعى عليه الأربع ====================

test('buildDefendantAnswerClause: مذكرة الدفاع الأولى تُرصد بنصها ويُصادق عليها', () => {
    const state = readyState();
    state.claim.defendantAnswerMode = M.DEFENDANT_ANSWER_MODES.FIRST_MEMO;
    state.claim.defendantFirstMemoText = 'أدفع بعدم صحة الدعوى';
    assert.equal(
        M.buildDefendantAnswerClause(state),
        ' وبالاطلاع على مذكرة الدفاع الأولى المقدَّمة من المدعى عليه ونصها: (( أدفع بعدم صحة الدعوى )) أ. هـ. وبعرضها على المدعى عليه صادق عليها.'
    );
    // النص فارغًا يُنبَّه على نسخه من الطلبات
    state.claim.defendantFirstMemoText = '';
    assert.match(M.buildDefendantAnswerClause(state), /\(\( \.+ تنسخ من الطلبات \.+ \)\)/);
    assert.ok(M.collectWarnings(state).some(x => x.includes('مذكرة الدفاع الأولى')));
});

test('buildDefendantAnswerClause: الإجابة الشفهية هي الصيغة الافتراضية', () => {
    const state = readyState();
    assert.equal(state.claim.defendantAnswerMode, M.DEFENDANT_ANSWER_MODES.ORAL);
    assert.equal(
        M.buildDefendantAnswerClause(state),
        ' وبعرضها على المدعى عليه أجاب قائلاً: ما ذكره المدعي غير صحيح هكذا أجاب.'
    );
});

test('buildDefendantAnswerClause: المذكرة المكتوبة تُقدَّم في الجلسة', () => {
    const state = readyState();
    state.claim.defendantAnswerMode = M.DEFENDANT_ANSWER_MODES.WRITTEN_MEMO;
    state.claim.defendantWrittenMemoText = 'ما ورد في الدعوى غير صحيح';
    assert.equal(
        M.buildDefendantAnswerClause(state),
        ' وبعرضها على المدعى عليه قدم مذكرة مكتوبة نصها: ما ورد في الدعوى غير صحيح، هكذا قدَّم.'
    );
    state.defendant.gender = 'ف';
    assert.match(M.buildDefendantAnswerClause(state), /على المدعى عليها قدمت مذكرة مكتوبة نصها: .*، هكذا قدَّمت\.$/);
    state.claim.defendantWrittenMemoText = '';
    assert.ok(M.collectWarnings(state).some(x => x.includes('المذكرة المكتوبة')));
});

test('buildDefendantAnswerClause: طلب المهلة صيغة موحدة بلا نص يُكتب', () => {
    const state = readyState();
    state.claim.defendantAnswerMode = M.DEFENDANT_ANSWER_MODES.DELAY;
    state.claim.defendantResponseText = '';
    assert.equal(
        M.buildDefendantAnswerClause(state),
        ' وبعرضها على المدعى عليه أجاب قائلاً: اطلب مهلة لتقديم الجواب مفصلا في الجلسة القادمة، هكذا أجاب.'
    );
    // ولا يُطالَب بنص جواب مع هذه الصيغة
    assert.ok(!M.collectWarnings(state).some(x => x.includes('إجابة المدعى عليه على الدعوى')));
});

test('buildDefendantAnswerClause: بيانات الوكيل تُثبت في الجواب بصوره كلها', () => {
    const state = readyState();
    state.defendant.attendance = 'تمثيل';
    state.defendant.agentName = 'خالد';
    state.defendant.wakalaIssuer = 'كتابة العدل بالرياض';
    state.defendant.wakalaNum = '777';
    state.defendant.licenseNum = '99';
    [
        M.DEFENDANT_ANSWER_MODES.FIRST_MEMO, M.DEFENDANT_ANSWER_MODES.ORAL,
        M.DEFENDANT_ANSWER_MODES.WRITTEN_MEMO, M.DEFENDANT_ANSWER_MODES.DELAY
    ].forEach(mode => {
        state.claim.defendantAnswerMode = mode;
        assert.match(
            M.buildDefendantAnswerClause(state),
            /وبعرضها على خالد، الحاضر بصفته وكيلاً بموجب الوكالة الصادرة من \(كتابة العدل بالرياض\)، برقم \(777\)، ورخصة مزاولة المحاماة رقم \(99\)،/,
            `صورة الجواب: ${mode}`
        );
    });
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

test('buildPlaintiffEvidenceText: غياب المدعى عليه يُصرّح بلقب المدعي بدل الضمير', () => {
    const state = readyState();
    state.defendant.attendance = 'لم يحضر';
    state.defendant.tabligh = '789';
    const text = M.composeMinutes(state);
    assert.match(text, / وبسؤال المدعي عن بينته قرر قائلاً:/);
    assert.ok(!text.includes('وبسؤاله عن بينته'));
    // ومع حضور المدعى عليه يبقى الضمير عائدًا على المدعي المذكور في فقرة التكليف بالبينة
    const present = M.composeMinutes(readyState());
    assert.match(present, /فقد جرى تكليف المدعي بإحضار بينته\. وبسؤاله عن بينته/);
});

test('التسلسل: غياب المدعى عليه يُبقي سؤال البينة مباشرة بلا عرض', () => {
    const state = readyState();
    state.defendant.attendance = 'لم يحضر';
    state.defendant.tabligh = '789';
    const text = M.composeMinutes(state);
    assert.match(text, /عن بينته/);
    // لا يُعرض على الغائب جوابٌ، وتبقى مصادقة المدعي على صحيفة دعواه وحدها
    assert.ok(!text.includes('وبعرضها على المدعى عليه'));
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
    assert.ok(!text.includes('جرى الاطلاع على صحيفة الدعوى'));
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

test('composeMinutes: بيانات وكيل المدعى عليه تُثبت في فقرة عرض الدعوى لا في الحضور', () => {
    const state = readyState({ includePartyDataInText: false });
    state.defendant.attendance = 'تمثيل';
    state.defendant.agentName = 'خالد';
    state.defendant.wakalaIssuer = 'كتابة العدل بالرياض';
    state.defendant.wakalaNum = '777';
    state.defendant.licenseNum = '99';
    const text = M.composeMinutes(state);
    assert.match(text, /صادق عليها\. وبعرضها على خالد، الحاضر بصفته وكيلاً بموجب الوكالة الصادرة من \(كتابة العدل بالرياض\)، برقم \(777\)، ورخصة مزاولة المحاماة رقم \(99\)، أجاب قائلاً:/);
    // لا تتكرر بيانات الوكيل في فقرات الحضور بعد الافتتاح
    assert.ok(!text.includes('حضر عن المدعى عليه'));
    assert.equal(text.match(/الوكالة الصادرة من/g).length, 1);
});

test('composeMinutes: الجلسة المنظورة سابقًا تُبقي وكيل المدعى عليه في فقرة الحضور', () => {
    const state = readyState({ includePartyDataInText: false });
    state.sessionType = 'previous';
    state.defendant.attendance = 'تمثيل';
    state.defendant.agentName = 'خالد';
    state.defendant.wakalaIssuer = 'كتابة العدل بالرياض';
    state.defendant.wakalaNum = '777';
    const text = M.composeMinutes(state);
    assert.match(text, /^لدي أنا فلان بن فلان في المحكمة الرياض، وحضر عن المدعى عليه خالد، بصفته وكيلاً بموجب الوكالة الصادرة من \(كتابة العدل بالرياض\)، برقم \(777\)/);
});

test('collectWarnings: جهة إصدار الوكالة لازمة لوكيل الطرف', () => {
    const state = readyState();
    state.defendant.attendance = 'تمثيل';
    state.defendant.agentName = 'خالد';
    state.defendant.wakalaNum = '777';
    state.defendant.licenseNum = '99';
    assert.ok(M.collectWarnings(state).some(x => x.includes('جهة إصدار وكالة وكيل المدعى عليه')));
    state.defendant.wakalaIssuer = 'كتابة العدل بالرياض';
    assert.ok(!M.collectWarnings(state).some(x => x.includes('جهة إصدار وكالة')));
});

// النص المعتمد للجلسة التحضيرية كما أقرَّه ناظر القضية — أي تعديل في الصياغة يجب أن يمر من هنا
test('composeMinutes: النص المعتمد للجلسة التحضيرية بلا بيانات الطرفين', () => {
    const state = M.freshMinutesState();
    state.sessionType = 'new';
    state.claim.defendantStance = 'إنكار';
    state.claim.evidenceChoice = 'none';
    assert.equal(
        M.composeMinutes(state),
        'لدي أنا ........... في المحكمة ...........، جرى الاطلاع على صحيفة الدعوى ونصها : (( ....... تنسخ من نظام ناجز ....... )) أ. هـ، وبعرضها على المدعي صادق عليها. وبعرضها على المدعى عليه أجاب قائلاً: ........... هكذا أجاب. ولمَّا كانت إجابة المدعى عليه إنكاراً لما جاء في الدعوى، وأن البينة على المدعي، فقد جرى تكليف المدعي بإحضار بينته. وبسؤاله عن بينته قرر قائلاً: لا بينة لدي، ثم جرى من الدائرة سؤال أطراف الدعوى هل لديكما ما تضيفانه؟ فقررا: ليس لدينا سوى ما قدمنا. هكذا قررا، واستناداً للمادة (69) والمادة (159) من نظام المرافعات الشرعية فقد قررت الدائرة قفل باب المرافعة للنطق بالحكم في هذه الجلسة، وأغلقت الجلسة الساعة الثامنة والنصف صباحًا.'
    );
});

test('composeMinutes: خلوّ فقرات الحضور لا يُكرِّر الفاصلة بعد الافتتاح', () => {
    const state = readyState({ includePartyDataInText: false });
    state.sessionType = 'previous';
    const text = M.composeMinutes(state);
    assert.ok(!text.includes('،،'));
    assert.match(text, /^لدي أنا فلان بن فلان في المحكمة الرياض، ثم جرى من الدائرة سؤال أطراف الدعوى/);
});

test('composeMinutes: جلسة تحضيرية مكتملة بطرفين حاضرين', () => {
    const text = M.composeMinutes(readyState());
    assert.match(text, /^لدي أنا فلان بن فلان في المحكمة الرياض/);
    assert.match(text, /حضر المدعي سعد أصالة/);
    assert.match(text, /حضر المدعى عليه فهد أصالة/);
    assert.match(text, /جرى الاطلاع على صحيفة الدعوى ونصها : \(\( أطالب بمبلغ مئة ألف ريال \)\) أ\. هـ/);
    assert.match(text, /وبعرضها على المدعى عليه أجاب قائلاً: ما ذكره المدعي غير صحيح/);
    assert.match(text, /قفل باب المرافعة/);
    assert.match(text, /وأغلقت الجلسة الساعة التاسعة والنصف صباحًا\.$/);
    // لا موضع ناقص في حالة مكتملة
    assert.ok(!text.includes(M.MINUTES_PLACEHOLDER));
});

test('composeMinutes: غياب المدعي المتبلّغ يولّد محضر شطب', () => {
    const state = readyState();
    state.plaintiff.attendance = 'لم يحضر';
    state.plaintiff.tabligh = '456';
    const text = M.composeMinutes(state);
    assert.match(text, /شطب الدَّعوى/);
    assert.match(text, /وأغلقت الجلسة/);
});

test('composeMinutes: مدعى عليه لم يتبلّغ — رفع الجلسة لإعادة التبليغ', () => {
    const state = readyState();
    state.defendant.attendance = 'لم يحضر';
    state.defendant.notifyStatus = 'لم يتبلغ';
    const text = M.composeMinutes(state);
    assert.match(text, /لم يحضر المدعى عليه، ولم يتبلّغ بالجلسة، وعليه رُفعت الجلسة لإعادة تبليغه بحسب حاله/);
    assert.ok(!text.includes('جرى الاطلاع على صحيفة الدعوى'));
});

test('composeMinutes: حالة استثنائية تستبدل المحضر بالكامل', () => {
    const state = readyState();
    state.plaintiff.specialCase = 'systemNoVideo';
    const text = M.composeMinutes(state);
    assert.match(text, /تبيّن حضور المدعي في النظام الإلكتروني/);
    assert.ok(!text.includes('جرى الاطلاع على صحيفة الدعوى'));
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

test('noticeKindFor: الطلب غير المالي لا يدخل حدّ الدعاوى اليسيرة', () => {
    const state = rulingState();
    assert.equal(M.freshClaimState().claimType, 'مالي');
    state.claim.claimValue = '10000';
    assert.equal(M.noticeKindFor(state), 'final');
    state.claim.claimType = 'غير مالي';
    assert.equal(M.noticeKindFor(state), 'appealable');
    // التحديد اليدوي يتقدم على الاشتقاق في الحالين
    state.ruling.noticeKind = 'sulh';
    assert.equal(M.noticeKindFor(state), 'sulh');
});

test('rulingWarnings: الطلب غير المالي لا يُطالَب بقيمة المطالبة', () => {
    const state = rulingState();
    assert.ok(M.rulingWarnings(state).some(w => w.includes('قيمة المطالبة')));
    state.claim.claimType = 'غير مالي';
    assert.ok(!M.rulingWarnings(state).some(w => w.includes('قيمة المطالبة')));
});

test('noticeKindFor: يُشتق نوع الإفهام من قيمة المطالبة', () => {
    const state = rulingState();
    state.claim.claimValue = '45000';
    assert.equal(M.noticeKindFor(state), 'final');
    state.claim.claimValue = String(M.YASEERA_CLAIM_LIMIT);
    assert.equal(M.noticeKindFor(state), 'final');
    state.claim.claimValue = '50001';
    assert.equal(M.noticeKindFor(state), 'appealable');
    // بلا قيمة يُحتاط بالقابلية للاستئناف
    state.claim.claimValue = '';
    assert.equal(M.noticeKindFor(state), 'appealable');
    // التحديد اليدوي يتقدم على الاشتقاق
    state.ruling.noticeKind = 'sulh';
    assert.equal(M.noticeKindFor(state), 'sulh');
});

test('buildNoticeText: النص مأخوذ من مكتبة النماذج لا مكتوب هنا', () => {
    const state = rulingState();
    state.claim.claimValue = '10000';
    assert.equal(M.buildNoticeText(state), M.findTemplate('إفهامات بعد النطق', 'ف1'));

    state.claim.claimValue = '900000';
    assert.equal(M.buildNoticeText(state), M.findTemplate('إفهامات بعد النطق', 'ف2'));

    // حضور الطرفين وكالةً يستدعي صيغة إفهام الوكلاء (المادة 165)
    state.plaintiff.attendance = 'تمثيل';
    state.defendant.attendance = 'تمثيل';
    assert.equal(M.buildNoticeText(state), M.findTemplate('إفهامات بعد النطق', 'ف3'));

    state.ruling.noticeKind = 'sulh';
    assert.equal(M.buildNoticeText(state), M.findTemplate('إفهامات بعد النطق', 'صلح1'));
});

test('noticeKindFor: صفة الوقف تُلزم بواجب التدقيق ولا تُترك للاختيار', () => {
    const state = rulingState();
    state.claim.claimValue = '10000';
    assert.equal(M.noticeKindFor(state), 'final');
    state.defendant.entityType = M.ENTITY_TYPES.WAQF;
    assert.equal(M.mandatoryReviewNotice(state), true);
    assert.equal(M.noticeKindFor(state), 'review');
    // حتى مع تحديد نوع آخر يدويًا يبقى واجب التدقيق
    state.ruling.noticeKind = 'sulh';
    assert.equal(M.noticeKindFor(state), 'review');
    assert.equal(M.buildNoticeText(state), M.findTemplate('إفهامات بعد النطق', 'واجب التدقيق'));
});

test('noticeKindFor: صفة الوقف للمدعي لا تُغيّر نوع الإفهام', () => {
    const state = rulingState();
    state.claim.claimValue = '10000';
    state.plaintiff.entityType = M.ENTITY_TYPES.WAQF;
    assert.equal(M.noticeKindFor(state), 'final');
});

test('buildNoticeText: اختيار «واجب التدقيق» يدويًا يجلب نصه من المكتبة', () => {
    const state = rulingState();
    state.ruling.noticeKind = 'review';
    assert.equal(M.buildNoticeText(state), M.findTemplate('إفهامات بعد النطق', 'واجب التدقيق'));
});

test('isCorporateEntity: الوقف شخصية اعتبارية كالشركة، ووثيقته صك الوقفية', () => {
    const s = M.freshPartyState();
    assert.equal(M.isCorporateEntity(s), false);
    s.entityType = M.ENTITY_TYPES.COMPANY;
    assert.equal(M.isCorporateEntity(s), true);
    assert.equal(M.corporateDocLabel(s), 'السجل التجاري');
    s.entityType = M.ENTITY_TYPES.WAQF;
    assert.equal(M.isCorporateEntity(s), true);
    assert.equal(M.corporateDocLabel(s), 'صك الوقفية');
    s.crNum = '77777';
    assert.equal(M.buildIdentityClause(s), '، بموجب صك الوقفية رقم (77777)');
});

test('composeMinutes: حكم على وقف — نص واجب التدقيق مرة واحدة بلا تكرار', () => {
    const state = rulingState();
    state.defendant.entityType = M.ENTITY_TYPES.WAQF;
    state.defendant.attendance = 'تمثيل';
    state.defendant.repType = 'وكيل شركة';
    const text = M.composeMinutes(state);
    const reviewText = M.findTemplate('إفهامات بعد النطق', 'واجب التدقيق');
    assert.equal(text.split(reviewText).length - 1, 1);
    assert.ok(!text.includes('قابل للاستئناف'));
});

test('composeMinutes: حكم على وقف مع إعادة القضية — صيغة واحدة لا صيغتان', () => {
    const state = rulingState();
    state.defendant.entityType = M.ENTITY_TYPES.WAQF;
    state.defendant.attendance = 'تمثيل';
    state.defendant.repType = 'وكيل شركة';
    state.ruling.mandatoryReview = 'إعادة';
    const text = M.composeMinutes(state);
    assert.match(text, /إعادة القضية بعد إنتهاء فترة الإعتراض إلى محكمة الأستئناف/);
    assert.ok(!text.includes('رفع كامل ملف الدعوى'));
});

test('rulingWarnings: لا يُطالَب بقيمة المطالبة مع الوقف واجب التدقيق', () => {
    const state = rulingState();
    assert.ok(M.rulingWarnings(state).some(w => w.includes('قيمة المطالبة')));
    state.defendant.entityType = M.ENTITY_TYPES.WAQF;
    assert.ok(!M.rulingWarnings(state).some(w => w.includes('قيمة المطالبة')));
});

test('buildNoticeText: إفهام واجب التدقيق له صيغتان يختارهما إجراء التدقيق', () => {
    const state = rulingState();
    state.ruling.noticeKind = 'review';
    assert.equal(M.freshRulingState().mandatoryReview, 'نعم');
    assert.equal(M.buildNoticeText(state), M.findTemplate('إفهامات بعد النطق', 'واجب التدقيق'));
    state.ruling.mandatoryReview = 'إعادة';
    assert.equal(M.buildNoticeText(state), M.findTemplate('إفهامات بعد النطق', 'واجب التدقيق /إعادة القضية'));
});

test('buildRulingSection: إجراء التدقيق لا يُلحق فقرة زائدة على إفهام غير واجب التدقيق', () => {
    const state = rulingState();
    state.claim.claimValue = '900000';
    state.ruling.mandatoryReview = 'إعادة';
    const text = M.buildRulingSection(state);
    assert.match(text, /قابل للاستئناف/);
    assert.ok(!text.includes('واجب التدقيق'));
});

test('composeMinutes: مرحلة الحكم تُلحق بالضبط بعد قفل باب المرافعة', () => {
    const state = rulingState();
    state.claim.claimValue = '900000';
    const text = M.composeMinutes(state);
    const closeAt = text.indexOf('قفل باب المرافعة');
    const reasonsAt = text.indexOf('\n\nالأسباب:\n');
    const rulingAt = text.indexOf('\n\nالحكم:\n');
    const endAt = text.indexOf('وأغلقت الجلسة');
    assert.ok(closeAt < reasonsAt && reasonsAt < rulingAt && rulingAt < endAt);
    assert.match(text, /وهذا الحكم حضوري في حق طرفي الدعوى/);
    assert.match(text, /قابل للاستئناف/);
    assert.match(text, /وأغلقت الجلسة الساعة التاسعة والنصف صباحًا\.$/);
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
    assert.ok(w.some(x => x.includes('صلة الشاهد بالمدعي')));
    assert.ok(w.some(x => x.includes('صلة الشاهد بالمدعى عليه')));
    assert.ok(w.some(x => x.includes('جنسية الشاهد')));
    assert.ok(w.some(x => x.includes('رقم هوية/إقامة الشاهد')));
});

test('collectWarnings: رقم هوية الشاهد عشرة أرقام كبيانات الطرفين', () => {
    const state = readyState();
    state.claim.evidenceChoice = 'has';
    state.claim.evidenceItems = ['شهادة شهود'];
    state.claim.witnesses = [Object.assign(M.freshWitness(), SAMPLE_WITNESS, { idNum: '12345' })];
    assert.ok(M.collectWarnings(state).some(x => x.includes('غير مكتمل (10 أرقام)')));

    state.claim.witnesses[0].idNum = '1099887766';
    assert.ok(!M.collectWarnings(state).some(x => x.includes('غير مكتمل (10 أرقام)')));
});

test('freshWitness: رقم الجوال ضمن بيانات الشاهد ولا يُفحص في التحذيرات', () => {
    assert.equal(M.freshWitness().phone, '');
    const state = readyState();
    state.claim.evidenceChoice = 'has';
    state.claim.evidenceItems = ['شهادة شهود'];
    state.claim.witnesses = [M.freshWitness()];
    assert.ok(!M.collectWarnings(state).some(x => x.includes('جوال')));
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
    // كل النماذج المنسّقة تبدأ بأحد مطالعها المعتمدة
    const openers = M.CURATED_REASON_OPENERS.map(M.normalizeArabicSearch);
    all.slice(0, curated).forEach(t => {
        const head = M.normalizeArabicSearch(t.content);
        assert.ok(openers.some(o => head.startsWith(o)));
    });
    // البحث يجد نماذج محدَّثة رغم اختلاف التشكيل والهمزات في الكتابة
    assert.ok(M.searchTemplates(all, 'اجرة المثل').length > 0);
});

// ==================== الغياب بلا سؤال عن البينة — المادة (21/3) من نظام الإثبات ====================

// حالة الغياب: جلسة تحضيرية والمدعى عليه لم يحضر ولا من يمثله
function absentDefendantState(overrides = {}) {
    const s = readyState();
    s.defendant.attendance = 'لم يحضر';
    s.defendant.tabligh = '123456';
    s.claim.evidenceChoice = 'noQuestion';
    return Object.assign(s, overrides);
}

test('خيار (بدون سؤال عن البينة): يُسقط سؤال المدعي ويطبع صيغة قفل المرافعة', () => {
    const text = M.composeMinutes(absentDefendantState());
    assert.ok(text.includes(M.NO_EVIDENCE_QUESTION_CLOSING));
    assert.ok(!text.includes('عن بينته'));
    assert.ok(!text.includes('لا بينة لدي'));
});

test('خيار (بدون سؤال عن البينة): لا يتكرر قفل باب المرافعة', () => {
    const text = M.composeMinutes(absentDefendantState());
    // الفقرة التلقائية (المادتان 69 و159) تُستبدل بصيغة الخيار، فلا يجتمع قفلان
    assert.ok(!text.includes('قفل باب المرافعة'));
    assert.equal(text.split('باب المرافعة').length - 1, 1);
});

test('خيار (بدون سؤال عن البينة): لا يُطبع مع خياري البينة الآخرين', () => {
    ['none', 'has'].forEach(choice => {
        const s = absentDefendantState();
        s.claim.evidenceChoice = choice;
        const text = M.composeMinutes(s);
        assert.ok(!text.includes(M.NO_EVIDENCE_QUESTION_CLOSING));
        assert.ok(text.includes('قفل باب المرافعة'));
    });
});

test('خيار (بدون سؤال عن البينة): يُسقط توجيه اليمين ولو طُلبت', () => {
    const s = absentDefendantState();
    s.claim.requestOath = true;
    const text = M.composeMinutes(s);
    assert.ok(!text.includes('اليمين'));
    assert.ok(text.includes(M.NO_EVIDENCE_QUESTION_CLOSING));
});

test('absenceReasonsVariant: الصيغة تُشتق من صفة المدعى عليه', () => {
    const variantOf = (patch) => {
        const s = absentDefendantState();
        Object.assign(s.defendant, patch);
        return M.absenceReasonsVariant(s);
    };
    assert.equal(variantOf({ entityType: 'فرد', gender: 'م' }), 'male');
    assert.equal(variantOf({ entityType: 'فرد', gender: 'ف' }), 'female');
    assert.equal(variantOf({ entityType: 'شركة', gender: 'ف' }), 'company');
    // الوقف يعامله المولد معاملة المذكر، فيأخذ صيغة المذكر
    assert.equal(variantOf({ entityType: 'وقف', gender: 'م' }), 'male');
});

test('absenceReasonsText: كل صيغة تطابق ألفاظ صفتها', () => {
    const s = absentDefendantState();
    s.defendant.gender = 'م';
    assert.ok(M.absenceReasonsText(s).includes('ولم يحضر أو يحضر من يمثله'));
    s.defendant.gender = 'ف';
    assert.ok(M.absenceReasonsText(s).includes('ولم تحضر في الموعد المحدد، كما لم يحضر من يمثلها'));
    s.defendant.entityType = 'شركة';
    assert.ok(M.absenceReasonsText(s).includes('ولم يحضر من يمثلها في الموعد المحدد'));
    // الصيغ الثلاث تستند جميعها إلى الفقرة الثالثة من المادة الحادية والعشرين
    Object.keys(M.ABSENCE_REASONS_M21).forEach(k => {
        assert.ok(M.ABSENCE_REASONS_M21[k].includes('الفقرة الثالثة من المادة الحادية والعشرين من نظام الإثبات'));
    });
});

test('isAbsenceReasonsText: يميّز النص المولَّد من كتابة القاضي', () => {
    assert.ok(M.isAbsenceReasonsText(M.ABSENCE_REASONS_M21.male));
    assert.ok(M.isAbsenceReasonsText(`\n${M.ABSENCE_REASONS_M21.company}\n`));
    assert.ok(!M.isAbsenceReasonsText(''));
    assert.ok(!M.isAbsenceReasonsText('فبناء على ما تقدم من الدعوى والإجابة'));
    assert.ok(!M.isAbsenceReasonsText(M.ABSENCE_REASONS_M21.female + ' وزيادة من القاضي'));
});

test('صيغ تسبيب الغياب مطابقة لنماذج المكتبة 1-3 في تصنيف (أسباب الحكم)', () => {
    const all = M.templatesOfCategory('أسباب الحكم');
    [['male', 0], ['female', 1], ['company', 2]].forEach(([key, idx]) => {
        assert.equal(M.ABSENCE_REASONS_M21[key], all[idx].content);
    });
});

test('نص الأسباب المولَّد يظهر في قسم الأسباب من المحضر', () => {
    const s = absentDefendantState();
    s.ruling.pronounce = 'نعم';
    s.ruling.reasonsText = M.absenceReasonsText(s);
    s.ruling.rulingText = 'إلزام المدعى عليه بأداء مئة ألف ريال';
    const text = M.composeMinutes(s);
    assert.ok(text.includes('الأسباب:'));
    assert.ok(text.includes(M.ABSENCE_REASONS_M21.male));
});
