// اختبارات محرك معالجة الصكوك القضائية في js/sakk.js
// التشغيل: npm test  (أو: node --test tests/)
//
// القيم المتوقعة هنا مثبَّتة من مخرجات المعالج الأصلي (sakk-processor) على النصوص نفسها،
// فهي عقد سلوكي: أي تغيير فيها يعني انحرافاً عن سلوك المصدر لا مجرد فشل اختبار.
const test = require('node:test');
const assert = require('node:assert/strict');

const S = require('../js/sakk.js');

// اختصار: معالجة بلا تصحيح إملائي (لعزل أثر المرحلة المدروسة)
function raw(text) {
    return S.processSakk(text, { enableCorrection: false });
}

// ==================== التنظيف الأساسي (المراحل 1–3) ====================

test('المسافات المتكررة تُستبدل بمسافة واحدة مع عدّها', () => {
    const r = raw('الحمد  لله  وحده');
    assert.equal(r.text, 'الحمد لله وحده');
    assert.equal(r.stats.spaces, 2);
});

test('المسافة قبل النقطتين تُحذف مع عدّها', () => {
    const r = raw('وبعد : فلدي أنا');
    assert.equal(r.text, 'وبعد: فلدي أنا');
    assert.equal(r.stats.colons, 1);
});

test('الأسطر الفارغة تُحذف مع عدّها', () => {
    const r = raw('السطر الأول\n\nالسطر الثاني');
    assert.equal(r.text, 'السطر الأول\nالسطر الثاني');
    assert.equal(r.stats.emptyLines, 1);
});

test('المراحل الثلاث تعمل مجتمعة', () => {
    const r = raw('الحمد  لله وبعد : كذا\n\n\nنص  الدعوى : كذا');
    assert.equal(r.text, 'الحمد لله وبعد: كذا\nنص الدعوى: كذا');
    assert.deepEqual(
        { spaces: r.stats.spaces, colons: r.stats.colons, emptyLines: r.stats.emptyLines },
        { spaces: 2, colons: 2, emptyLines: 1 }
    );
});

// ==================== تعقيم علامات التتبع من المدخل ====================

test('علامات التتبع الواردة في المدخل تُعقَّم فلا تُفسد المخرجات', () => {
    const dirty = 'نص فيه ' + S.MERGE_START + 'علامة' + S.MERGE_END + ' و' + S.DELETE_MARK + 'أخرى';
    const r = raw(dirty);
    assert.equal(r.text, 'نص فيه علامة وأخرى');
    assert.match(r.annotated, /^[^\u0001-\u0003]*$/);
    assert.equal(r.stats.sessionsMerged, 0);
});

// ==================== دمج الجلسات: الأنماط الأحد عشر (المرحلة 4) ====================

// كل حالة تمثّل نمطاً من مصفوفة sessionMergePatterns بترتيبها
const MERGE_CASES = [
    ['1- الختام بالساعة + افتتاح الجلسة',
        'وكان ختامها الساعة 11:30 صباحاً. افتتحت الجلسة المرئية عن بعد وفيها حضر المدعي',
        'وفي جلسة أخرى حضر المدعي'],
    ['2- الختام بالساعة + في هذه الجلسة',
        'وكان ختامها الساعة 9:15 مساءً. وفي هذه الجلسة حضر الطرفان',
        'وفي جلسة أخرى حضر الطرفان'],
    ['3- رفع الجلسة للنطق بالحكم (يحفظ المقدمة)',
        'رفعت الجلسة للنطق بالحكم. افتتحت الجلسة وفيها حضر المدعى عليه',
        'رفعت الجلسة للنطق بالحكم. وفي جلسة أخرى حضر المدعى عليه'],
    ['4- النكول (يحفظ المقدمة)',
        'وإذا لم يحضر يحكم عليه بالنكول. وفي هذه الجلسة حضر المدعي',
        'وإذا لم يحضر يحكم عليه بالنكول. وفي جلسة أخرى حضر المدعي'],
    ['5- للدراسة والاطلاع (يحفظ المقدمة)',
        'رفعت الجلسة للدراسة والاطلاع وفي هذه الجلسة وعبر الاتصال المرئي حضر الطرفان',
        'رفعت الجلسة للدراسة والاطلاع وفي جلسة أخرى حضر الطرفان'],
    ['6- التوفيق + قرار المجلس الأعلى للقضاء (الأطول)',
        'هذا وبالله التوفيق. وفي هذا اليوم عقدت الدائرة هذه الجلسة عن بعد بناء على قرار المجلس الأعلى للقضاء رقم (17388) بتاريخ 5/10/1441 هـ والمبلغ بتعميم معالي رئيس المجلس الأعلى للقضاء – وفقه الله – برقم (1505) بتاريخ 5/5/1441 هـ حضر المدعي',
        'وفي جلسة أخرى حضر المدعي'],
    ['7- التوفيق + افتتاح الجلسة',
        'هذا وبالله التوفيق. افتتحت الجلسة عبر الاتصال المرئي وفيها حضر المدعي',
        'وفي جلسة أخرى حضر المدعي'],
    ['8- التوفيق + عقدت هذه الجلسة',
        'وبالله التوفيق. وعقدت هذه الجلسة حضر المدعي',
        'وفي جلسة أخرى حضر المدعي'],
    ['9- رفعت الجلسة + افتتاح (يحفظ المقدمة)',
        'رفعت الجلسة. افتتحت الجلسة بالترافع المرئي وفيها حضر المدعي',
        'رفعت الجلسة. وفي جلسة أخرى حضر المدعي'],
    ['10- رفعت الجلسة + في هذه الجلسة (يحفظ المقدمة)',
        'رفعت الجلسة. وفي هذه الجلسة حضر المدعي',
        'رفعت الجلسة. وفي جلسة أخرى حضر المدعي'],
    ['11- صباحاً/مساءً + في هذه الجلسة (النمط العام)',
        'تمت في تمام العاشرة صباحاً. وفي هذه الجلسة حضر المدعي',
        'تمت في تمام العاشرة وفي جلسة أخرى حضر المدعي']
];

for (const [name, input, expected] of MERGE_CASES) {
    test('دمج الجلسات — ' + name, () => {
        const r = raw(input);
        assert.equal(r.text, expected);
        assert.equal(r.stats.sessionsMerged, 1);
    });
}

test('عدد أنماط الدمج المحددة أحد عشر نمطاً', () => {
    assert.equal(S.buildSessionMergePatterns().length, 11);
});

test('أنماط الدمج تُبنى جديدة في كل استدعاء فلا تحمل حالة lastIndex', () => {
    const input = 'رفعت الجلسة. وفي هذه الجلسة حضر المدعي';
    assert.equal(raw(input).text, raw(input).text);
    assert.equal(raw(input).stats.sessionsMerged, 1);
});

test('جلسات متتابعة بأنماط مختلطة تُدمج كلها', () => {
    const r = raw('رفعت الجلسة. وفي هذه الجلسة حضر المدعي. هذا وبالله التوفيق. افتتحت الجلسة وفيها حضر المدعى عليه. وكان ختامها الساعة 9:15 مساءً. وفي هذه الجلسة قُفل باب المرافعة');
    assert.equal(r.text, 'رفعت الجلسة. وفي جلسة أخرى حضر المدعي. وفي جلسة أخرى حضر المدعى عليه. وفي جلسة أخرى قُفل باب المرافعة');
    assert.equal(r.stats.sessionsMerged, 3);
});

// ==================== التصحيح الإملائي والتشكيل (المرحلة 5) ====================

test('التصحيح يطبَّق ويُحصى ويُسجَّل في correctedWords', () => {
    const r = S.processSakk('ان المدعي اقر بذلك اثناء الجلسة، وقد احيل الى محكمة التمييز اليوم');
    assert.equal(r.text, 'أنَّ المدعي أقرَّ بذلك أثناء الجلسة، وقد أُحِيْلَ إلى محكمة الاستئناف اليوم');
    assert.equal(r.stats.corrections, 6);
    assert.ok(r.correctedWords.some((w) => w.wrong === 'اقر' && w.correct === 'أقرَّ'));
});

test('إطفاء التصحيح يترك النص كما هو بلا أي تصحيح', () => {
    const input = 'ان المدعي اقر بذلك اثناء الجلسة، وقد احيل الى محكمة التمييز اليوم';
    const r = raw(input);
    assert.equal(r.text, input);
    assert.equal(r.stats.corrections, 0);
    assert.deepEqual(r.correctedWords, []);
});

test('التصحيح محصور بحدود الكلمة العربية فلا يقع داخل كلمة أطول', () => {
    // «ان» مدخل في القاموس، ولا يجوز أن يلتقطه داخل «كان» أو «هناك»
    const r = S.processSakk('كان المدعي هناك');
    assert.equal(r.text, 'كان المدعي هناك');
    assert.equal(r.stats.corrections, 0);
});

test('القاموس مرتَّب تنازلياً بطول الكلمة الخاطئة', () => {
    const sorted = S.sortedCorrections();
    for (let i = 1; i < sorted.length; i++) {
        assert.ok(sorted[i - 1][1].length >= sorted[i][1].length);
    }
});

test('ترتيب القاموس لا يمسّ المصفوفة المشتركة', () => {
    const data = require('../data/sakk-corrections.js');
    const firstBefore = data[0];
    S.sortedCorrections();
    assert.equal(data[0], firstBefore);
});

test('القاموس يحوي 218 مدخلاً كلها أزواج نصية', () => {
    const data = require('../data/sakk-corrections.js');
    assert.equal(data.length, 218);
    assert.ok(data.every((p) => Array.isArray(p) && p.length === 2 &&
        typeof p[0] === 'string' && typeof p[1] === 'string'));
});

// ==================== الدمج الاحتياطي (المرحلة 5ب) ====================

test('الدمج الاحتياطي يعمل تحت عتبة 17 كلمة مع عبارة ختام معروفة', () => {
    const r = raw('رفعت الجلسة وبالله التوفيق حضر المدعي وكيلاً عن موكله');
    assert.equal(r.text, 'رفعت الجلسة وفي جلسة أخرى حضر المدعي وكيلاً عن موكله');
    assert.equal(r.stats.sessionsMerged, 1);
});

test('الدمج الاحتياطي لا يعمل فوق عتبة 17 كلمة', () => {
    const input = 'رفعت الجلسة وبالله التوفيق ' + 'كلمة '.repeat(20) + 'حضر المدعي';
    const r = raw(input);
    assert.equal(r.text, input);
    assert.equal(r.stats.sessionsMerged, 0);
});

test('الدمج الاحتياطي لا يعمل بلا عبارة ختام معروفة', () => {
    const input = 'رفعت الجلسة ثم حضر المدعي';
    const r = raw(input);
    assert.equal(r.text, input);
    assert.equal(r.stats.sessionsMerged, 0);
});

test('الدمج الاحتياطي يتجاهل ما عالجته الأنماط المحددة فلا يُعدّ مرتين', () => {
    // النمط 9 يعالج هذا المقطع، فلا يعود الاحتياطي يعدّه مرة أخرى
    const r = raw('رفعت الجلسة. افتتحت الجلسة وفيها حضر المدعي');
    assert.equal(r.stats.sessionsMerged, 1);
});

test('ازدواج الواو الناتج عن الدمج يُنظَّف', () => {
    assert.ok(!raw('ورفعت الجلسة. وفي هذه الجلسة حضر المدعي').text.includes('ووفي جلسة أخرى'));
});

// ==================== حذف جلسات الشطب (المرحلة 6) ====================

test('جلسة الشطب تُحذف كاملة بين بداية الجلسة وصيغة الختام', () => {
    const r = raw('افتتحت الجلسة وفيها حضر المدعي. لذا فقد قررت الدائرة شطب الدعوى وعلى ذلك أغلقت الجلسة، وبالله التوفيق. افتتحت الجلسة وفيها حضر الطرفان مرة أخرى');
    assert.equal(r.stats.shatbDeleted, 1);
    assert.ok(!r.text.includes('شطب الدعوى'));
    assert.ok(r.text.includes('حضر الطرفان مرة أخرى'));
});

test('الجلسة الصحيحة بعد الشطب لا تُلتهم (السقف الآمن)', () => {
    const r = raw('افتتحت الجلسة وفيها حضر المدعي. لذا فقد قررت الدائرة شطب الدعوى وعلى ذلك أغلقت الجلسة، وبالله التوفيق. افتتحت الجلسة وفيها حضر الطرفان مرة أخرى');
    assert.ok(r.text.includes('مرة أخرى'), 'الجلسة التالية للشطب يجب أن تبقى');
});

test('تعذّر تحديد نهاية جلسة الشطب يرفع shatbSkipped بلا حذف (فشل آمن)', () => {
    const input = 'افتتحت الجلسة وفيها حضر المدعي. لذا قررت الدائرة شطب الدعوى ولم تُختم بصيغة معروفة';
    const r = raw(input);
    assert.equal(r.stats.shatbDeleted, 0);
    assert.equal(r.stats.shatbSkipped, 1);
    assert.equal(r.text, input);
    assert.equal(r.warnings.length, 1);
});

test('عبارة الشطب مع تشكيل على «الدائرة» تُلتقط', () => {
    const r = raw('افتتحت الجلسة وفيها حضر المدعي. لذا فقد قررت الدَّائرة شطب الدعوى وصحبه أجمعين.');
    assert.equal(r.stats.shatbDeleted, 1);
});

test('جلستا شطب في نص واحد تُحذفان معاً', () => {
    const r = raw('افتتحت الجلسة وفيها حضر المدعي. لذا قررت الدائرة شطب الدعوى وصحبه أجمعين. افتتحت الجلسة وفيها حضر الطرفان. لذا فقد قررت الدائرة شطب الدعوى وعلى ذلك أغلقت الجلسة، وبالله التوفيق. افتتحت الجلسة وفيها استؤنف النظر');
    assert.equal(r.stats.shatbDeleted, 2);
    assert.ok(!r.text.includes('شطب الدعوى'));
    assert.ok(r.text.includes('استؤنف النظر'));
});

test('شطب بلا بداية جلسة سابقة لا يُحذف ولا يُتجاوز', () => {
    const input = 'لذا فقد قررت الدائرة شطب الدعوى وصحبه أجمعين.';
    const r = raw(input);
    assert.equal(r.stats.shatbDeleted, 0);
    assert.equal(r.stats.shatbSkipped, 0);
    assert.equal(r.text, input);
});

// ==================== العقد العام للمخرجات ====================

test('text خالٍ من علامات التتبع و annotated يحملها', () => {
    const r = raw('رفعت الجلسة. وفي هذه الجلسة حضر المدعي');
    assert.ok(!/[\u0001-\u0003]/.test(r.text));
    assert.ok(/[\u0001-\u0003]/.test(r.annotated));
    assert.equal(r.annotated.replace(/[\u0001-\u0003]/g, ''), r.text);
});

test('النص السليم يمر بلا تغيير وبإحصائيات صفرية', () => {
    const input = 'نص سليم بلا مشاكل.';
    const r = raw(input);
    assert.equal(r.text, input);
    assert.deepEqual(r.stats, {
        spaces: 0, colons: 0, emptyLines: 0,
        sessionsMerged: 0, shatbDeleted: 0, shatbSkipped: 0, corrections: 0
    });
    assert.deepEqual(r.warnings, []);
});

test('المدخل الفارغ أو المعدوم لا يُسقط المحرك', () => {
    for (const input of ['', '   ', null, undefined]) {
        const r = raw(input);
        assert.equal(typeof r.text, 'string');
        assert.equal(r.stats.sessionsMerged, 0);
    }
});

test('المعالجة لا تعدّل المدخل ولا تحمل حالة بين الاستدعاءات', () => {
    const input = 'رفعت الجلسة. وفي هذه الجلسة حضر المدعي';
    const first = S.processSakk(input);
    const second = S.processSakk(input);
    assert.equal(first.text, second.text);
    assert.deepEqual(first.stats, second.stats);
});

test('countChars يعدّ الأحرف بلا مسافات', () => {
    assert.equal(S.countChars('أ ب\nج\tد'), 4);
    assert.equal(S.countChars(''), 0);
    assert.equal(S.countChars(null), 0);
});

// ==================== ضبط مركّب واقعي (مطابقة شاملة) ====================

test('ضبط مركّب: جلسات متعددة وشطب وتصحيح ومسافات وأسطر فارغة', () => {
    const input = [
        'الحمد لله وحده وبعد : فلدي أنا فلان بن فلان القاضي بالمحكمة العامة',
        '',
        'افتتحت الجلسة  وفيها حضر المدعي أصالة، وقد اقر بان الدعوى صحيحة اثناء نظرها.',
        'وكان ختامها الساعة 11:30 صباحاً. افتتحت الجلسة المرئية عن بعد وفيها حضر المدعى عليه',
        'رفعت الجلسة للنطق بالحكم. افتتحت الجلسة وفيها حضر الطرفان',
        '',
        'لذا فقد قررت الدائرة شطب الدعوى وعلى ذلك أغلقت الجلسة، وبالله التوفيق.',
        'افتتحت الجلسة وفيها حضر المدعي مرة اخرى وقد احيل الامر الى محكمة التمييز :  وانتهت'
    ].join('\n');

    const r = S.processSakk(input);

    assert.equal(
        r.text,
        'الحمد لله وحده وبعد: فلدي أنا فلان بن فلان القاضي بالمحكمة العامة\n وفي جلسة أخرى حضر المدعي مرة اخرى وقد أُحِيْلَ الامر إلى محكمة الاستئناف: وانتهت'
    );
    assert.deepEqual(r.stats, {
        spaces: 2, colons: 2, emptyLines: 2,
        sessionsMerged: 3, shatbDeleted: 1, shatbSkipped: 0, corrections: 6
    });
});
