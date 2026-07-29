// اختبارات دوال الحساب الصافية في js/calc.js
// التشغيل: npm test  (أو: node --test tests/)
const test = require('node:test');
const assert = require('node:assert/strict');

const {
    numberToArabicWords,
    computeInheritance,
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
