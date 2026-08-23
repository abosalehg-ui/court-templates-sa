// اختبارات دوال الحساب الصافية في js/calc.js
// التشغيل: npm test  (أو: node --test tests/)
const test = require('node:test');
const assert = require('node:assert/strict');

const {
    numberToArabicWords,
    tafqitAmount,
    tafqitSAR,
    computeInheritance,
    computeInheritanceAmounts,
    dateDiffYMD,
    computeEOSCore
} = require('../js/calc.js');

// قالب ورثة فارغ — كل الحقول صفر
function emptyHeirs(overrides = {}) {
    return {
        sons: 0, daughters: 0, grandsonsSon: 0, granddaughtersSon: 0,
        father: 0, mother: 0, grandfather: 0,
        grandmotherMother: 0, grandmotherFather: 0,
        husband: 0, wives: 0,
        brothersFull: 0, sistersFull: 0,
        brothersFather: 0, sistersFather: 0,
        brothersMother: 0, sistersMother: 0,
        nephewsFull: 0, nephewsFather: 0,
        unclesFull: 0, unclesFather: 0,
        cousinsFull: 0, cousinsFather: 0,
        ...overrides
    };
}

function shareOf(result, heirName) {
    return result.shares.find(s => s.heir === heirName);
}

// ==================== التفقيط ====================

test('التفقيط: أعداد أساسية', () => {
    assert.equal(numberToArabicWords(0), 'صفر');
    assert.equal(numberToArabicWords(1), 'واحد');
    assert.equal(numberToArabicWords(11), 'أحد عشر');
    assert.equal(numberToArabicWords(25), 'خمسة وعشرون');
    assert.equal(numberToArabicWords(100), 'مائة');
    assert.equal(numberToArabicWords(345), 'ثلاثمائة وخمسة وأربعون');
});

test('التفقيط: الآلاف والملايين', () => {
    assert.equal(numberToArabicWords(1000), 'ألف');
    assert.equal(numberToArabicWords(2000), 'ألفان');
    assert.equal(numberToArabicWords(5000), 'خمسة آلاف');
    assert.equal(numberToArabicWords(1000000), 'مليون');
    assert.equal(numberToArabicWords(400000), 'أربعمائة ألف');
    assert.equal(numberToArabicWords(1250), 'ألف ومائتان وخمسون');
});

test('التفقيط: الأعداد السالبة', () => {
    assert.equal(numberToArabicWords(-5), 'سالب خمسة');
});

test('التفقيط بأسلوب «مئة»: تشمل المائتين لا المئات المركبة وحدها', () => {
    assert.equal(numberToArabicWords(300, 'مئة'), 'ثلاثمئة');
    assert.equal(numberToArabicWords(200, 'مئة'), 'مئتان');
});

// ==================== تفقيط المبالغ بالعملة ====================

test('tafqitSAR: لاحقة الريال الموحدة للحاسبات', () => {
    assert.equal(tafqitSAR(1), 'واحد ريال سعودي فقط لا غير');
    assert.equal(tafqitSAR(2), 'ريالان سعوديان فقط لا غير');
    assert.equal(tafqitSAR(5), 'خمسة ريالات سعودية فقط لا غير');
    assert.equal(tafqitSAR(150), 'مائة وخمسون ريال سعودي فقط لا غير');
    // الحاسبات تُسقط الكسور كما كانت تفعل كتلها المكررة
    assert.equal(tafqitSAR(150.75), 'مائة وخمسون ريال سعودي فقط لا غير');
});

test('tafqitAmount: تثنية الوحدة الفرعية صحيحة (هللتان لا هللةان)', () => {
    assert.equal(tafqitAmount(5.02, 'SAR'), 'خمسة ريالات سعودية وهللتان فقط لا غير');
    assert.equal(tafqitAmount(1.25, 'SAR'), 'واحد ريال سعودي وخمسة وعشرون هللة فقط لا غير');
    assert.equal(tafqitAmount(3.02, 'OMR'), 'ثلاثة ريالات عمانية وبيستان فقط لا غير');
});

test('tafqitAmount: بدون عملة يخرج العدد وحده بلا خاتمة', () => {
    assert.equal(tafqitAmount(7, 'NONE'), 'سبعة');
});

// ==================== المواريث: العمريتان (البند 2) ====================

test('العمرية الأولى: زوج + أم + أب — للأم ثلث الباقي (1/6) لا ثلث التركة', () => {
    const r = computeInheritance(emptyHeirs({ husband: 1, mother: 1, father: 1 }), 0);
    const husband = shareOf(r, 'الزوج');
    const mother = shareOf(r, 'الأم');
    const father = shareOf(r, 'الأب');

    assert.equal(husband.shareValue, 1 / 2);
    assert.equal(mother.share, '1/3 الباقي');
    // ثلث الباقي بعد النصف = (1 - 1/2) / 3 = 1/6
    assert.ok(Math.abs(mother.shareValue - 1 / 6) < 1e-9,
        `نصيب الأم يجب أن يكون 1/6 وليس ${mother.shareValue}`);
    assert.equal(father.type, 'تعصيب');
    // لا عول في العمرية: 1/2 + 1/6 = 2/3 < 1
    assert.equal(r.awl, null);
});

test('العمرية الثانية: زوجة + أم + أب — للأم ثلث الباقي (1/4)', () => {
    const r = computeInheritance(emptyHeirs({ wives: 1, mother: 1, father: 1 }), 0);
    const mother = shareOf(r, 'الأم');
    // ثلث الباقي بعد الربع = (1 - 1/4) / 3 = 1/4
    assert.ok(Math.abs(mother.shareValue - 1 / 4) < 1e-9,
        `نصيب الأم يجب أن يكون 1/4 وليس ${mother.shareValue}`);
});

test('أم بلا زوج ولا فرع ولا جمع إخوة: تأخذ الثلث كاملاً', () => {
    const r = computeInheritance(emptyHeirs({ mother: 1, father: 1 }), 0);
    const mother = shareOf(r, 'الأم');
    assert.equal(mother.share, '1/3');
    assert.ok(Math.abs(mother.shareValue - 1 / 3) < 1e-9);
});

// ==================== المواريث: الجد والإخوة (البند 1) ====================

test('جد + إخوة أشقاء: لا ازدواج في الباقي — الإخوة محجوبون والجد وحده عاصب', () => {
    const r = computeInheritance(emptyHeirs({ grandfather: 1, brothersFull: 2 }), 0);

    // الجد وحده يأخذ الباقي تعصيباً
    const asabaShares = r.shares.filter(s => s.type === 'تعصيب');
    assert.equal(asabaShares.length, 1, 'يجب أن يكون هناك عاصب واحد فقط');
    assert.equal(asabaShares[0].heir, 'الجد');

    // الإخوة محجوبون بالجد ولا يظهرون في الأنصبة
    assert.equal(shareOf(r, 'الإخوة الأشقاء'), undefined);
    const blockedBrothers = r.blocked.find(b => b.heir === 'الإخوة الأشقاء');
    assert.ok(blockedBrothers, 'يجب أن يظهر الإخوة الأشقاء في المحجوبين');
    assert.equal(blockedBrothers.blockedBy, 'الجد');
});

test('جد + أخوات شقيقات وإخوة لأب: الجميع محجوبون بالجد', () => {
    const r = computeInheritance(emptyHeirs({ grandfather: 1, sistersFull: 1, brothersFather: 1 }), 0);
    assert.equal(shareOf(r, 'الأخت الشقيقة'), undefined);
    assert.ok(r.blocked.some(b => b.heir === 'الأخوات الشقيقات' && b.blockedBy === 'الجد'));
    assert.ok(r.blocked.some(b => b.heir === 'الإخوة لأب' && b.blockedBy === 'الجد'));
});

test('إخوة أشقاء بلا جد ولا أب: يرثون الباقي تعصيباً', () => {
    const r = computeInheritance(emptyHeirs({ brothersFull: 2, mother: 1 }), 0);
    const brothers = shareOf(r, 'الإخوة الأشقاء');
    assert.ok(brothers, 'الإخوة الأشقاء يجب أن يرثوا عند عدم الأب والجد');
    assert.equal(brothers.type, 'تعصيب');
});

// ==================== المواريث: الحجب الأساسي ====================

test('الأب يحجب الجد والإخوة', () => {
    const r = computeInheritance(emptyHeirs({ father: 1, grandfather: 1, brothersFull: 1 }), 0);
    assert.ok(r.blocked.some(b => b.heir === 'الجد' && b.blockedBy === 'الأب'));
    assert.ok(r.blocked.some(b => b.heir === 'الإخوة الأشقاء' && b.blockedBy === 'الأب'));
    assert.equal(shareOf(r, 'الجد'), undefined);
});

test('الابن يحجب أبناء الابن', () => {
    const r = computeInheritance(emptyHeirs({ sons: 1, grandsonsSon: 1 }), 0);
    assert.ok(r.blocked.some(b => b.heir === 'أبناء الابن'));
});

test('الفرع الوارث يحجب الإخوة لأم', () => {
    const r = computeInheritance(emptyHeirs({ daughters: 1, brothersMother: 2 }), 0);
    assert.ok(r.blocked.some(b => b.heir === 'الإخوة لأم'));
    assert.equal(shareOf(r, 'الإخوة والأخوات لأم'), undefined);
});

// ==================== المواريث: الفروض والعول ====================

test('بنت واحدة: النصف', () => {
    const r = computeInheritance(emptyHeirs({ daughters: 1 }), 0);
    assert.equal(shareOf(r, 'البنت').shareValue, 1 / 2);
});

test('عول: زوج + أختان شقيقتان (7/6)', () => {
    const r = computeInheritance(emptyHeirs({ husband: 1, sistersFull: 2 }), 0);
    assert.ok(r.awl, 'المسألة يجب أن تكون عائلة');
    assert.ok(Math.abs(r.totalShares - 7 / 6) < 1e-9);
});

test('زوج مع فرع وارث: الربع', () => {
    const r = computeInheritance(emptyHeirs({ husband: 1, sons: 1 }), 0);
    assert.equal(shareOf(r, 'الزوج').shareValue, 1 / 4);
});

test('رد: يُكتشف عند نقص الفروض عن الواحد بلا عصبة', () => {
    // بنت فقط: 1/2 والباقي بلا عاصب
    const r = computeInheritance(emptyHeirs({ daughters: 1 }), 0);
    assert.ok(r.radd, 'يجب اكتشاف الرد');
    assert.ok(Math.abs(r.radd.remaining - 1 / 2) < 1e-9);
});

// ==================== نهاية الخدمة ====================

test('نهاية الخدمة: 10 سنوات بأجر 10000 (المادة 84)', () => {
    const r = computeEOSCore(10000, 'employer', { years: 10, months: 0, days: 0 });
    assert.equal(r.firstFive, 25000);   // 5 × 0.5 × 10000
    assert.equal(r.afterFive, 50000);   // 5 × 10000
    assert.equal(r.fullAward, 75000);
    assert.equal(r.dueAward, 75000);
    assert.equal(r.ratio, 1);
});

test('استقالة أقل من سنتين: لا يستحق', () => {
    const r = computeEOSCore(10000, 'resignation', { years: 1, months: 6, days: 0 });
    assert.equal(r.ratio, 0);
    assert.equal(r.dueAward, 0);
});

test('استقالة 3 سنوات: الثلث', () => {
    const r = computeEOSCore(12000, 'resignation', { years: 3, months: 0, days: 0 });
    assert.ok(Math.abs(r.ratio - 1 / 3) < 1e-9);
    // المكافأة الكاملة: 3 × 0.5 × 12000 = 18000 → الثلث = 6000
    assert.ok(Math.abs(r.dueAward - 6000) < 1e-6);
});

test('استقالة 7 سنوات: الثلثان', () => {
    const r = computeEOSCore(10000, 'resignation', { years: 7, months: 0, days: 0 });
    assert.ok(Math.abs(r.ratio - 2 / 3) < 1e-9);
});

test('استقالة 12 سنة: المكافأة كاملة', () => {
    const r = computeEOSCore(10000, 'resignation', { years: 12, months: 0, days: 0 });
    assert.equal(r.ratio, 1);
});

test('المادة 80: لا مكافأة', () => {
    const r = computeEOSCore(10000, 'art80', { years: 8, months: 0, days: 0 });
    assert.equal(r.dueAward, 0);
});

test('كسر السنة يُحتسب بنسبته', () => {
    const r = computeEOSCore(12000, 'employer', { years: 2, months: 6, days: 0 });
    // 2.5 سنة × نصف شهر × 12000 = 15000
    assert.ok(Math.abs(r.fullAward - 15000) < 1e-6);
});

test('مدخلات ناقصة تُعيد null', () => {
    assert.equal(computeEOSCore(0, 'employer', { years: 5, months: 0, days: 0 }), null);
    assert.equal(computeEOSCore(10000, 'employer', null), null);
});

// ==================== فرق التواريخ ====================

test('dateDiffYMD: فرق بسيط', () => {
    const d = dateDiffYMD('2020-01-01', '2025-01-01');
    assert.deepEqual(d, { years: 5, months: 0, days: 0 });
});

test('dateDiffYMD: نهاية قبل البداية تُعيد خطأ', () => {
    const d = dateDiffYMD('2025-01-01', '2020-01-01');
    assert.ok(d.error);
});

// ==================== توزيع مبالغ التركة (computeInheritanceAmounts) ====================

// مبلغ مجموعة وارث معين من نتيجة التوزيع
function amountOf(result, amounts, heirName) {
    const idx = result.shares.findIndex(s => s.heir === heirName);
    return idx > -1 ? amounts[idx] : undefined;
}

test('فرض + تعصيب: الأب مع بنت يأخذ السدس والباقي معاً', () => {
    const r = computeInheritance(emptyHeirs({ father: 1, daughters: 1 }), 60000);
    const amounts = computeInheritanceAmounts(r, 60000);
    assert.equal(Math.round(amountOf(r, amounts, 'البنت')), 30000);
    // الأب: 10000 فرضاً (السدس) + 20000 الباقي تعصيباً
    assert.equal(Math.round(amountOf(r, amounts, 'الأب')), 30000);
    const total = amounts.reduce((a, b) => a + b, 0);
    assert.equal(Math.round(total), 60000);
});

test('الرد: بنت + أم بلا عصبة — الباقي يُرد بنسبة الفرضين ويستغرق التركة', () => {
    const r = computeInheritance(emptyHeirs({ daughters: 1, mother: 1 }), 80000);
    assert.ok(r.radd, 'المسألة ردّية');
    const amounts = computeInheritanceAmounts(r, 80000);
    // الفرضان 1/2 و1/6 (نسبة 3:1)، فالرد يوصلهما إلى 3/4 و1/4
    assert.equal(Math.round(amountOf(r, amounts, 'البنت')), 60000);
    assert.equal(Math.round(amountOf(r, amounts, 'الأم')), 20000);
});

test('الرد لا يشمل الزوجين: زوج + أم — الرد كله للأم', () => {
    const r = computeInheritance(emptyHeirs({ husband: 1, mother: 1 }), 60000);
    const amounts = computeInheritanceAmounts(r, 60000);
    // الزوج نصفه فرضاً بلا زيادة، والأم ثلثها + الباقي رداً = النصف
    assert.equal(Math.round(amountOf(r, amounts, 'الزوج')), 30000);
    assert.equal(Math.round(amountOf(r, amounts, 'الأم')), 30000);
});

test('العول: زوج + شقيقتان — تُنقص الفروض بنسبها ويستغرق المجموع التركة', () => {
    const r = computeInheritance(emptyHeirs({ husband: 1, sistersFull: 2 }), 70000);
    assert.ok(r.awl, 'المسألة عائلة');
    const amounts = computeInheritanceAmounts(r, 70000);
    // الفروض 1/2 و2/3 (عالت إلى 7/6): الزوج 3/7 والأخوات 4/7
    assert.equal(Math.round(amountOf(r, amounts, 'الزوج')), 30000);
    assert.equal(Math.round(amountOf(r, amounts, 'الأخوات الشقيقات')), 40000);
});

test('التعصيب المحض: زوجة + ابن — الباقي بعد الثمن للعصبة', () => {
    const r = computeInheritance(emptyHeirs({ wives: 1, sons: 1 }), 80000);
    const amounts = computeInheritanceAmounts(r, 80000);
    assert.equal(Math.round(amountOf(r, amounts, 'الزوجة')), 10000);
    assert.equal(Math.round(amountOf(r, amounts, 'الأبناء')), 70000);
});

test('بلا قيمة تركة: المبالغ كلها صفر', () => {
    const r = computeInheritance(emptyHeirs({ daughters: 1 }), 0);
    const amounts = computeInheritanceAmounts(r, 0);
    assert.ok(amounts.every(a => a === 0));
});

test('computeInheritance لا تعدّل كائن المدخلات وتعيد الورثة الأصليين', () => {
    const heirs = emptyHeirs({ father: 1, grandfather: 1, brothersFull: 2 });
    const r = computeInheritance(heirs, 0);
    // الحجب يصفّر الجد والإخوة داخلياً، لكن مدخلات المستدعي تبقى كما هي
    assert.equal(heirs.grandfather, 1);
    assert.equal(heirs.brothersFull, 2);
    assert.equal(r.originalHeirs.brothersFull, 2);
});

test('تسميات الجمع سليمة: الزوجات والجدات', () => {
    const withWives = computeInheritance(emptyHeirs({ wives: 2, sons: 1 }), 0);
    assert.ok(shareOf(withWives, 'الزوجات'), 'زوجتان تُعرضان باسم «الزوجات»');
    const withGrandmas = computeInheritance(emptyHeirs({ grandmotherMother: 1, grandmotherFather: 1, sons: 1 }), 0);
    assert.ok(shareOf(withGrandmas, 'الجدات'), 'جدتان تُعرضان باسم «الجدات»');
});
