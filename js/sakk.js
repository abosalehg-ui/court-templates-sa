// ==================== محرك معالجة الصكوك القضائية ====================
// هذا الملف يحتوي منطق المعالجة الصافي فقط (بلا أي تعامل مع DOM)
// حتى يمكن اختباره آلياً عبر `npm test` (انظر tests/sakk.test.js)
// واجهة المستخدم في js/sakk-app.js وصفحة sakk.html
//
// المصدر: معالج الصكوك القضائية — https://github.com/abosalehg-ui/sakk-processor (رخصة MIT)
// نُقل المنطق والأنماط كما هي حرفياً، وأُزيل تشابكها مع الواجهة فقط.

// ==================== علامات التتبع الداخلية ====================
// تُحقن أثناء المعالجة لتعليم مواضع الدمج والحذف، فتتحول إلى تظليل في العرض
// وتُحذف من النص النهائي. مأخوذة من نطاق المحارف التحكمية حتى لا تصادف نصاً قضائياً.
const MERGE_START = '\u0001';
const MERGE_END = '\u0002';
const DELETE_MARK = '\u0003';
const MERGED_PHRASE = MERGE_START + 'وفي جلسة أخرى' + MERGE_END;

const SAKK_MARKS = { MERGE_START, MERGE_END, DELETE_MARK, MERGED_PHRASE };

// ==================== مصدر قاموس التصحيح ====================
// يعمل في المتصفح (متغير عام من data/sakk-corrections.js) وفي Node (require)
function correctionsSource() {
    if (typeof sakkCorrections !== 'undefined') return sakkCorrections;
    if (typeof require !== 'undefined') {
        try {
            return require('../data/sakk-corrections.js');
        } catch (e) {
            return [];
        }
    }
    return [];
}

// القاموس مرتباً تنازلياً بطول الكلمة الخاطئة، حتى لا تلتقط كلمة قصيرة جزءاً من كلمة أطول.
// يُرتَّب على نسخة حتى لا يُمسّ ترتيب المصفوفة المشتركة.
function sortedCorrections() {
    return correctionsSource().slice().sort((a, b) => b[1].length - a[1].length);
}

// ==================== الأنماط الفرعية الموحّدة ====================
// عبارة افتتاح الجلسة بكل صيغها (عادية / هذه / المرئية / عن بعد / عبر الاتصال المرئي)
const SESSION_OPEN = String.raw`افتتحت\s+(?:هذه\s+)?الجلسة(?:\s+المرئية)?(?:\s+(?:عبر\s+الترافع\s+المرئي|بالترافع\s+المرئي|عبر\s+الاتصال\s+المرئي|بالاتصال\s+المرئي))?(?:\s+عن\s+بُ?عد)?`;
// ذيل «وفيها» يُحذف مع عبارة الافتتاح
const TAIL = String.raw`(?:\s+وفيها)?(?:\s+و(?=\s|$))?`;
// عبارة ختام الجلسة بالساعة
const KHITAM = String.raw`وكان\s+ختامها\s+الساعة\s+\d{1,2}:\d{2}\s+(?:صباحاً|صباحًا|مساءً|مساءًا|[صم])(?![\u0600-\u06FF])`;
// عبارة التوفيق بكل صيغها (هذا وبالله / هذا بالله / وبالله / بالله)
// عبارة الصلاة على النبي التي قد تلي «وبالله التوفيق» (اختيارية)
const SALAT = String.raw`(?:\s*[.،,]?\s*وصلى\s+الله\s+(?:وسلم\s+)?على\s+نبينا\s+محمد\s+و[آاأ]له\s+وصحبه(?:\s+وسلم)?(?:\s+[اأ]جمعين)?)?`;
const TAWFIQ = String.raw`(?:هذا\s+)?و?بالله\s+التوفيق` + SALAT;

// ==================== أنماط دمج الجلسات ====================
// الترتيب مقصود: الأطول والأخص أولاً، حتى لا يبتلع نمط عام ما يخص نمطاً أدق.
// تُبنى داخل دالة لأن الأنماط عامة (g) فتحمل حالة lastIndex بين الاستدعاءات.
function buildSessionMergePatterns() {
    return [

        // ========== أنماط الختام بالساعة ==========

        // 1- وكان ختامها الساعة XX:XX صباحاً/مساءً. افتتحت الجلسة [المرئية عن بعد] [وفيها]
        {
            pattern: new RegExp(KHITAM + String.raw`\s*\.?\s*` + SESSION_OPEN + TAIL, 'g'),
            replacement: 'وفي جلسة أخرى'
        },
        // 2- وكان ختامها الساعة XX:XX صباحاً/مساءً. [و]في هذه الجلسة
        {
            pattern: new RegExp(KHITAM + String.raw`\s*\.?\s*و?في\s+هذه\s+الجلسة`, 'g'),
            replacement: 'وفي جلسة أخرى'
        },

        // ========== أنماط متخصصة (تحافظ على المقدمة) ==========

        // 3- رفع الجلسة للنطق بالحكم. افتتحت الجلسة [المرئية عن بعد] [وفيها]
        {
            pattern: new RegExp(String.raw`(رفعت?\s+الجلسة\s+للنطق\s+بالحكم)\s*\.?\s*` + SESSION_OPEN + TAIL, 'g'),
            replacement: '$1. وفي جلسة أخرى'
        },
        // 4- وإذا لم يحضر يحكم عليه بالنكول. [و]في هذه الجلسة
        {
            pattern: /(وإذا\s+لم\s+يحضر\s+يحكم\s+عليه\s+بالنكول)\s*\.?\s*و?في\s+هذه\s+الجلسة/g,
            replacement: '$1. وفي جلسة أخرى'
        },
        // 5- للدراسة والاطلاع وفي هذه الجلسة وعبر الاتصال المرئي
        {
            pattern: /(للدراسة\s+والاطلاع)\s+و?في\s+هذه\s+الجلسة\s+وعبر\s+الاتصال\s+المرئي/g,
            replacement: '$1 وفي جلسة أخرى'
        },

        // ========== أنماط التوفيق (الأطول أولاً) ==========

        // 6- هذا وبالله التوفيق. وفي هذا اليوم عقدت الدائرة هذه الجلسة عن بعد بقرار المجلس الأعلى للقضاء
        {
            pattern: new RegExp(TAWFIQ + String.raw`\s*\.?\s*وفي\s+هذا\s+اليوم\s+عقدت\s+الدائرة\s+هذه\s+الجلسة\s+عن\s+بعد\s+بناء\s+على\s+قرار\s+المجلس\s+الأعلى\s+للقضاء\s+رقم\s*\([^)]*\)\s*بتاريخ\s*[\d\s/]+\s*هـ\s*والمبلغ\s+بتعميم\s+معالي\s+رئيس\s+المجلس\s+الأعلى\s+للقضاء\s*[–—-]\s*وفقه\s+الله\s*[–—-]\s*برقم\s*\([^)]*\)\s*بتاريخ\s*[\d\s/]+\s*هـ`, 'g'),
            replacement: 'وفي جلسة أخرى'
        },
        // 7- [هذا] و/بالله التوفيق. افتتحت الجلسة [المرئية عن بعد] [وفيها]
        {
            pattern: new RegExp(TAWFIQ + String.raw`\s*\.?\s*` + SESSION_OPEN + TAIL, 'g'),
            replacement: 'وفي جلسة أخرى'
        },
        // 8- [هذا] و/بالله التوفيق. [و]في هذه الجلسة / [و]عقدت هذه الجلسة
        {
            pattern: new RegExp(TAWFIQ + String.raw`\s*\.?\s*(?:و?في\s+هذه\s+الجلسة|و?عقدت\s+هذه\s+الجلسة)`, 'g'),
            replacement: 'وفي جلسة أخرى'
        },

        // ========== أنماط رفع الجلسة (تحافظ على سبب الرفع) ==========

        // 9- رفع/رفعت الجلسة [.] افتتحت الجلسة [المرئية عن بعد] [وفيها]
        {
            pattern: new RegExp(String.raw`(رفعت?\s+الجلسة(?:\s+(?:لذلك|لذا))?)\s*\.?\s*` + SESSION_OPEN + TAIL, 'g'),
            replacement: '$1. وفي جلسة أخرى'
        },
        // 10- رفع/رفعت الجلسة [.] [و]في هذه الجلسة
        {
            pattern: /(رفعت?\s+الجلسة(?:\s+(?:لذلك|لذا))?)\s*\.?\s*و?في\s+هذه\s+الجلسة/g,
            replacement: '$1. وفي جلسة أخرى'
        },

        // ========== أنماط عامة (تأتي آخراً) ==========

        // 11- صباحاً/مساءً. [و]في هذه الجلسة (نمط مختصر)
        {
            pattern: /(?:صباحاً|صباحًا|مساءً|مساءًا)\s*\.\s*و?في\s+هذه\s+الجلسة/g,
            replacement: 'وفي جلسة أخرى'
        }
    ];
}

// ==================== أنماط جلسات الشطب ====================
const SHATB_PATTERN_SOURCE = String.raw`لذا\s+(?:فقد\s+)?قررت\s+الد[\u064B-\u0652]*ائرة\s+شطب\s+الدعوى`;

// أنماط بداية الجلسة — تُستخدم لتحديد بداية نطاق الحذف ونهايته الآمنة
const SESSION_START_SOURCES = [
    'افتتحت\\s+الجلسة',
    'فلدي\\s+أنا',
    'لدي\\s+أنا',
    'وفي\\s+هذا\\s+اليوم\\s+عقدت\\s+الد[\\u064B-\\u0652]*ائرة\\s+هذه\\s+الجلسة'
];

// صيغة ختام الجلسة التي ينتهي عندها نطاق حذف جلسة الشطب
const SHATB_END_PATTERN = /(?:وصحبه\s+(?:أجمعين|وسل[\u064B-\u0652]*م)|(?:و\s*)?على\s+ذلك\s+أ?غلقت\s+الجلسة(?:\s*[،,]\s*(?:هذا\s+)?و?بالله\s+التوفيق(?:\s+وصلى\s+الله[\s\S]{0,80}?وسل[\u064B-\u0652]*م)?)?)\s*[.،]?/;

// ==================== عدد الأحرف بلا مسافات ====================
function countChars(text) {
    return String(text == null ? '' : text).replace(/\s/g, '').length;
}

// ==================== المعالجة ====================
// processSakk(input, options) → نتيجة كاملة بلا أي أثر على DOM
//   options.enableCorrection — تفعيل التشكيل والتصحيح الإملائي (افتراضياً مفعّل)
// المخرجات:
//   text           النص النظيف النهائي (بلا علامات تتبع) — هو ما يُنسخ ويُحمَّل
//   annotated      النص نفسه مع علامات التتبع — تبني منه الواجهة عرض المقارنة
//   stats          إحصائيات العمليات الست
//   correctedWords قائمة {wrong, correct, count} لتظليل الكلمات المصححة
//   warnings       رسائل للمستخدم (بديل showToast الذي كان داخل المحرك)
function processSakk(input, options) {
    const opts = options || {};
    const enableCorrection = opts.enableCorrection !== false;
    const inputText = String(input == null ? '' : input);

    const stats = {
        spaces: 0,
        colons: 0,
        emptyLines: 0,
        sessionsMerged: 0,
        shatbDeleted: 0,
        shatbSkipped: 0,
        corrections: 0
    };
    const correctedWords = [];
    const warnings = [];

    // تعقيم علامات التتبع من المدخل حتى لا تختلط بما يحقنه المحرك
    let text = inputText.replace(/[\u0001-\u0003]/g, '');

    // 1. استبدال المسافتين بمسافة واحدة
    const beforeSpaces = text;
    text = text.replace(/  +/g, ' ');
    stats.spaces = (beforeSpaces.match(/  +/g) || []).length;

    // 2. استبدال " :" ب ":"
    const beforeColons = text;
    text = text.replace(/ :/g, ':');
    stats.colons = (beforeColons.match(/ :/g) || []).length;

    // 3. حذف الأسطر الفارغة
    const beforeEmptyLines = text;
    text = text.replace(/\n\s*\n/g, '\n');
    stats.emptyLines = (beforeEmptyLines.match(/\n\s*\n/g) || []).length;

    // 4. دمج الجلسات بالأنماط المحددة
    buildSessionMergePatterns().forEach(function (entry) {
        const matches = text.match(entry.pattern);
        if (matches) {
            stats.sessionsMerged += matches.length;
            text = text.replace(entry.pattern, entry.replacement.replace('وفي جلسة أخرى', MERGED_PHRASE));
        }
    });

    // 5. التصحيح اللغوي (إذا كان مفعّلاً)
    // حدود الكلمة العربية مكتوبة يدوياً لأن \b لا تعمل مع الحروف العربية
    if (enableCorrection) {
        sortedCorrections().forEach(function (pair) {
            const correct = pair[0];
            const wrong = pair[1];
            const escapedWrong = wrong.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
            const regex = new RegExp('(^|[^\\u0600-\\u06FF])' + escapedWrong + '(?![\\u0600-\\u06FF])', 'g');
            const foundMatches = text.match(regex);
            if (foundMatches) {
                stats.corrections += foundMatches.length;
                correctedWords.push({ wrong: wrong, correct: correct, count: foundMatches.length });
                text = text.replace(regex, '$1' + correct);
            }
        });
    }

    // آلية الدمج الاحتياطية (العامة):
    // تلتقط فواصل جلسات لم تغطها الأنماط المحددة في الخطوة 4. حدودها الموثقة:
    //   - تعمل فقط على مقاطع أقصر من 17 كلمة بين «رفعت الجلسة» و«حضر»
    //     وتحوي إحدى عبارات الختام المعروفة (وبالله التوفيق / ختامها / للدراسة)
    //   - تتجاهل أي مقطع سبق أن عولج بالأنماط المحددة (يحوي علامات داخلية)
    //     منعاً للدمج المزدوج والعد المزدوج في الإحصائيات
    const sessionRegex = /(رفعت?\s+الجلسة)([\s\S]*?)(حضرت?(?:\s|$))/g;
    let match;
    const sessionMatches = [];
    while ((match = sessionRegex.exec(text)) !== null) {
        sessionMatches.push({
            fullMatch: match[0],
            start: match[1],
            middle: match[2],
            end: match[3],
            index: match.index
        });
    }

    // مرور عكسي حتى لا تنزاح فهارس المطابقات الباقية بعد كل استبدال
    for (let i = sessionMatches.length - 1; i >= 0; i--) {
        const m = sessionMatches[i];
        const middleText = m.middle;

        if (/[\u0001-\u0003]/.test(middleText)) {
            continue;
        }

        const wordCount = middleText.trim().split(/\s+/).length;

        if (wordCount < 17) {
            const hasTawfiq = middleText.includes('وبالله التوفيق');
            const hasKhitam = middleText.includes('ختامها');
            const hasDirasa = middleText.includes('للدراسة');

            if (hasTawfiq || hasKhitam || hasDirasa) {
                const replacement = m.start + ' ' + MERGED_PHRASE + ' ' + m.end;
                text = text.substring(0, m.index) + replacement + text.substring(m.index + m.fullMatch.length);
                stats.sessionsMerged++;
            }
        }
    }

    // إزالة ازدواج الواو الناتج عن الدمج (مثل: «ورفعت الجلسة. افتتحت الجلسة» → «ووفي جلسة أخرى»)
    text = text.replace(new RegExp('و' + MERGE_START + 'وفي جلسة أخرى', 'g'), MERGE_START + 'وفي جلسة أخرى');
    text = text.replace(/ووفي\s+جلسة\s+أخرى/g, 'وفي جلسة أخرى');

    // 6. حذف جلسات الشطب
    // يُقيَّد البحث عن صيغة الختام بنطاق الجلسة نفسها (حتى بداية الجلسة التالية إن وُجدت)
    // حتى لا يلتهم الحذف جلسات صحيحة إذا لم تُختم جلسة الشطب بالصيغة المتوقعة.
    // عند تعذر تحديد النهاية بأمان يُتجاوز الحذف ويُنبَّه المستخدم
    const shatbPattern = new RegExp(SHATB_PATTERN_SOURCE, 'g');
    let shatbMatch;
    const shatbPositions = [];

    while ((shatbMatch = shatbPattern.exec(text)) !== null) {
        shatbPositions.push(shatbMatch.index);
    }

    for (let i = shatbPositions.length - 1; i >= 0; i--) {
        const shatbPos = shatbPositions[i];
        const textBeforeShatb = text.substring(0, shatbPos);

        // آخر بداية جلسة قبل عبارة الشطب = بداية نطاق الحذف
        let lastStartPos = -1;

        for (const source of SESSION_START_SOURCES) {
            const pattern = new RegExp(source, 'g');
            let startMatch;
            while ((startMatch = pattern.exec(textBeforeShatb)) !== null) {
                if (startMatch.index > lastStartPos) {
                    lastStartPos = startMatch.index;
                }
            }
        }

        if (lastStartPos !== -1) {
            const textAfterShatb = text.substring(shatbPos);

            // أول بداية جلسة تالية بعد عبارة الشطب = الحد الأقصى الآمن لنطاق الحذف
            let nextStartPos = textAfterShatb.length;
            for (const source of SESSION_START_SOURCES) {
                const nextMatch = new RegExp(source, 'g').exec(textAfterShatb);
                if (nextMatch && nextMatch.index < nextStartPos) {
                    nextStartPos = nextMatch.index;
                }
            }

            const endMatch = SHATB_END_PATTERN.exec(textAfterShatb.substring(0, nextStartPos));

            if (endMatch) {
                const endPos = shatbPos + endMatch.index + endMatch[0].length;
                text = text.substring(0, lastStartPos) + DELETE_MARK + text.substring(endPos);
                stats.shatbDeleted++;
            } else {
                stats.shatbSkipped++;
            }
        }
    }

    if (stats.shatbSkipped > 0) {
        warnings.push('تم تجاوز حذف ' + stats.shatbSkipped + ' جلسة شطب لتعذر تحديد نهايتها بأمان');
    }

    return {
        text: text.replace(/[\u0001-\u0003]/g, ''),
        annotated: text,
        stats: stats,
        correctedWords: correctedWords,
        warnings: warnings
    };
}

// تصدير للاختبارات في بيئة Node (لا يؤثر على المتصفح)
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        SAKK_MARKS,
        MERGE_START, MERGE_END, DELETE_MARK, MERGED_PHRASE,
        SESSION_OPEN, TAIL, KHITAM, SALAT, TAWFIQ,
        SESSION_START_SOURCES, SHATB_PATTERN_SOURCE, SHATB_END_PATTERN,
        buildSessionMergePatterns, sortedCorrections,
        countChars, processSakk
    };
}
