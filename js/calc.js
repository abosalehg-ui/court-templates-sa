// ==================== دوال الحساب الصافية ====================
// هذا الملف يحتوي منطق الحساب فقط (بلا أي تعامل مع DOM)
// حتى يمكن اختباره آلياً عبر `npm test` (انظر مجلد tests/)

// ==================== TAFQIT (Number to Text) ====================
const currencies = {
    SAR: { singular: 'ريال سعودي', dual: 'ريالان سعوديان', plural: 'ريالات سعودية', gender: 'm', subunit: 'هللة', subunitDual: 'هللتان', subunitPlural: 'هللات' },
    USD: { singular: 'دولار أمريكي', dual: 'دولاران أمريكيان', plural: 'دولارات أمريكية', gender: 'm', subunit: 'سنت', subunitDual: 'سنتان', subunitPlural: 'سنتات' },
    EUR: { singular: 'يورو', dual: 'يورو', plural: 'يورو', gender: 'm', subunit: 'سنت', subunitDual: 'سنتان', subunitPlural: 'سنتات' },
    GBP: { singular: 'جنيه إسترليني', dual: 'جنيهان إسترلينيان', plural: 'جنيهات إسترلينية', gender: 'm', subunit: 'بنس', subunitDual: 'بنسان', subunitPlural: 'بنسات' },
    AED: { singular: 'درهم إماراتي', dual: 'درهمان إماراتيان', plural: 'دراهم إماراتية', gender: 'm', subunit: 'فلس', subunitDual: 'فلسان', subunitPlural: 'فلوس' },
    KWD: { singular: 'دينار كويتي', dual: 'ديناران كويتيان', plural: 'دنانير كويتية', gender: 'm', subunit: 'فلس', subunitDual: 'فلسان', subunitPlural: 'فلوس' },
    QAR: { singular: 'ريال قطري', dual: 'ريالان قطريان', plural: 'ريالات قطرية', gender: 'm', subunit: 'درهم', subunitDual: 'درهمان', subunitPlural: 'دراهم' },
    BHD: { singular: 'دينار بحريني', dual: 'ديناران بحرينيان', plural: 'دنانير بحرينية', gender: 'm', subunit: 'فلس', subunitDual: 'فلسان', subunitPlural: 'فلوس' },
    OMR: { singular: 'ريال عماني', dual: 'ريالان عمانيان', plural: 'ريالات عمانية', gender: 'm', subunit: 'بيسة', subunitDual: 'بيستان', subunitPlural: 'بيسات' },
    EGP: { singular: 'جنيه مصري', dual: 'جنيهان مصريان', plural: 'جنيهات مصرية', gender: 'm', subunit: 'قرش', subunitDual: 'قرشان', subunitPlural: 'قروش' },
    JOD: { singular: 'دينار أردني', dual: 'ديناران أردنيان', plural: 'دنانير أردنية', gender: 'm', subunit: 'قرش', subunitDual: 'قرشان', subunitPlural: 'قروش' },
    NONE: { singular: '', dual: '', plural: '', gender: 'm', subunit: '', subunitDual: '', subunitPlural: '' }
};

const ones = ['', 'واحد', 'اثنان', 'ثلاثة', 'أربعة', 'خمسة', 'ستة', 'سبعة', 'ثمانية', 'تسعة', 'عشرة', 'أحد عشر', 'اثنا عشر', 'ثلاثة عشر', 'أربعة عشر', 'خمسة عشر', 'ستة عشر', 'سبعة عشر', 'ثمانية عشر', 'تسعة عشر'];
const tens = ['', '', 'عشرون', 'ثلاثون', 'أربعون', 'خمسون', 'ستون', 'سبعون', 'ثمانون', 'تسعون'];
const hundreds = ['', 'مائة', 'مائتان', 'ثلاثمائة', 'أربعمائة', 'خمسمائة', 'ستمائة', 'سبعمائة', 'ثمانمائة', 'تسعمائة'];

function numberToArabicWords(num, style = 'مائة') {
    if (num === 0) return 'صفر';
    if (num < 0) return 'سالب ' + numberToArabicWords(-num, style);

    let words = '';

    // Billions
    if (num >= 1000000000) {
        const billions = Math.floor(num / 1000000000);
        if (billions === 1) words += 'مليار ';
        else if (billions === 2) words += 'ملياران ';
        else if (billions >= 3 && billions <= 10) words += numberToArabicWords(billions, style) + ' مليارات ';
        else words += numberToArabicWords(billions, style) + ' مليار ';
        num %= 1000000000;
        if (num > 0) words += 'و';
    }

    // Millions
    if (num >= 1000000) {
        const millions = Math.floor(num / 1000000);
        if (millions === 1) words += 'مليون ';
        else if (millions === 2) words += 'مليونان ';
        else if (millions >= 3 && millions <= 10) words += numberToArabicWords(millions, style) + ' ملايين ';
        else words += numberToArabicWords(millions, style) + ' مليون ';
        num %= 1000000;
        if (num > 0) words += 'و';
    }

    // Thousands
    if (num >= 1000) {
        const thousands = Math.floor(num / 1000);
        if (thousands === 1) words += 'ألف ';
        else if (thousands === 2) words += 'ألفان ';
        else if (thousands >= 3 && thousands <= 10) words += numberToArabicWords(thousands, style) + ' آلاف ';
        else words += numberToArabicWords(thousands, style) + ' ألف ';
        num %= 1000;
        if (num > 0) words += 'و';
    }

    // Hundreds
    if (num >= 100) {
        let h = hundreds[Math.floor(num / 100)];
        // «مائتان» لا تحوي «مائة» حرفيًا فتُعالج على حدة، وإلا خرج النص مختلط الأسلوب عند 200
        if (style === 'مئة') h = h.replace('مائتان', 'مئتان').replace('مائة', 'مئة');
        words += h + ' ';
        num %= 100;
        if (num > 0) words += 'و';
    }

    // Tens and ones
    if (num >= 20) {
        const t = Math.floor(num / 10);
        const o = num % 10;
        if (o > 0) {
            words += ones[o] + ' و' + tens[t] + ' ';
        } else {
            words += tens[t] + ' ';
        }
    } else if (num > 0) {
        words += ones[num] + ' ';
    }

    return words.trim();
}

// صياغة التفقيط الكاملة لمبلغ بعملة: الجزء الصحيح، وكسوره بالوحدة الفرعية، وخاتمة «فقط لا غير».
// هذه هي الصياغة الوحيدة في المشروع — كانت مكررة في ست حاسبات داخل js/app.js.
function tafqitAmount(num, currencyCode = 'SAR', style = 'مائة') {
    const value = Number(num);
    if (isNaN(value) || value < 0) return '';

    const intPart = Math.floor(value);
    const decPart = Math.round((value - intPart) * 100);
    let result = numberToArabicWords(intPart, style);

    const curr = currencies[currencyCode];
    if (!curr || currencyCode === 'NONE') return result;

    if (intPart === 1) result += ' ' + curr.singular;
    else if (intPart === 2) result = curr.dual;
    else if (intPart >= 3 && intPart <= 10) result += ' ' + curr.plural;
    else result += ' ' + curr.singular;

    if (decPart > 0) {
        if (decPart === 2) {
            // المثنى صيغة قائمة بذاتها («هللتان») لا عدد + مفرد («اثنان هللة»)
            result += ' و' + curr.subunitDual;
        } else {
            result += ' و' + numberToArabicWords(decPart, style);
            if (decPart === 1) result += ' ' + curr.subunit;
            else if (decPart >= 3 && decPart <= 10) result += ' ' + curr.subunitPlural;
            else result += ' ' + curr.subunit;
        }
    }

    return result + ' فقط لا غير';
}

// اللاحقة السعودية الموحدة للحاسبات (تُسقط الكسور كما كانت تفعل الحاسبات كلها)
function tafqitSAR(num) {
    return tafqitAmount(Math.floor(Number(num) || 0), 'SAR');
}

// ==================== INHERITANCE ====================
function computeInheritance(heirs, estate) {
    // نسخة محلية حتى لا تتغير مدخلات المستدعي (الحجب يصفّر بعض الحقول أثناء الحساب)
    const originalHeirs = { ...heirs };
    heirs = { ...heirs };
    const shares = [];
    const blocked = [];
    let totalShares = 0;
    let baseValue = 24; // أصل المسألة الافتراضي

    const hasBranch = heirs.sons > 0 || heirs.daughters > 0;
    const hasMaleBranch = heirs.sons > 0;
    const hasGrandsonsBranch = heirs.grandsonsSon > 0;

    // Determine blocking (الحجب)
    // Sons block grandsons
    if (heirs.sons > 0 && heirs.grandsonsSon > 0) {
        blocked.push({ heir: 'أبناء الابن', blockedBy: 'الأبناء', reason: 'الفرع الأقرب يحجب الأبعد' });
    }
    if (heirs.sons > 0 && heirs.granddaughtersSon > 0) {
        blocked.push({ heir: 'بنات الابن', blockedBy: 'الأبناء', reason: 'الابن يحجب بنات الابن' });
    }

    // Father blocks grandfather
    if (heirs.father > 0 && heirs.grandfather > 0) {
        blocked.push({ heir: 'الجد', blockedBy: 'الأب', reason: 'الأب يحجب الجد' });
        heirs.grandfather = 0;
    }

    // Father blocks siblings
    if (heirs.father > 0) {
        if (heirs.brothersFull > 0) blocked.push({ heir: 'الإخوة الأشقاء', blockedBy: 'الأب', reason: 'الأب يحجب الإخوة' });
        if (heirs.sistersFull > 0) blocked.push({ heir: 'الأخوات الشقيقات', blockedBy: 'الأب', reason: 'الأب يحجب الأخوات' });
        if (heirs.brothersFather > 0) blocked.push({ heir: 'الإخوة لأب', blockedBy: 'الأب', reason: 'الأب يحجب الإخوة لأب' });
        if (heirs.sistersFather > 0) blocked.push({ heir: 'الأخوات لأب', blockedBy: 'الأب', reason: 'الأب يحجب الأخوات لأب' });
        heirs.brothersFull = heirs.sistersFull = heirs.brothersFather = heirs.sistersFather = 0;
    }

    // Grandfather blocks siblings when there is no father
    // تعتمد الحاسبة القول بأن الجد بمنزلة الأب في حجب الإخوة والأخوات
    // (وهو ما أخذ به نظام الأحوال الشخصية السعودي؛ ومسألة الجد والإخوة موضع خلاف فقهي)
    if (heirs.father === 0 && heirs.grandfather > 0) {
        if (heirs.brothersFull > 0) blocked.push({ heir: 'الإخوة الأشقاء', blockedBy: 'الجد', reason: 'الجد يحجب الإخوة عند عدم الأب (والمسألة خلافية)' });
        if (heirs.sistersFull > 0) blocked.push({ heir: 'الأخوات الشقيقات', blockedBy: 'الجد', reason: 'الجد يحجب الأخوات عند عدم الأب (والمسألة خلافية)' });
        if (heirs.brothersFather > 0) blocked.push({ heir: 'الإخوة لأب', blockedBy: 'الجد', reason: 'الجد يحجب الإخوة لأب عند عدم الأب (والمسألة خلافية)' });
        if (heirs.sistersFather > 0) blocked.push({ heir: 'الأخوات لأب', blockedBy: 'الجد', reason: 'الجد يحجب الأخوات لأب عند عدم الأب (والمسألة خلافية)' });
        heirs.brothersFull = heirs.sistersFull = heirs.brothersFather = heirs.sistersFather = 0;
    }

    // Branch blocks siblings for mother
    if (hasBranch) {
        if (heirs.brothersMother > 0) blocked.push({ heir: 'الإخوة لأم', blockedBy: 'الفرع الوارث', reason: 'الفرع الوارث يحجب الإخوة لأم' });
        if (heirs.sistersMother > 0) blocked.push({ heir: 'الأخوات لأم', blockedBy: 'الفرع الوارث', reason: 'الفرع الوارث يحجب الأخوات لأم' });
        heirs.brothersMother = heirs.sistersMother = 0;
    }

    // 1. الزوج / الزوجة (أصحاب فروض لا يحجبون)
    if (heirs.husband > 0) {
        const share = hasBranch ? '1/4' : '1/2';
        const shareValue = hasBranch ? 1 / 4 : 1 / 2;
        shares.push({
            heir: 'الزوج',
            count: 1,
            share: share,
            shareValue: shareValue,
            type: 'فرض',
            evidence: hasBranch
                ? 'قال تعالى: {فَإِن كَانَ لَهُنَّ وَلَدٌ فَلَكُمُ الرُّبُعُ مِمَّا تَرَكْنَ}'
                : 'قال تعالى: {وَلَكُمْ نِصْفُ مَا تَرَكَ أَزْوَاجُكُمْ إِن لَّمْ يَكُن لَّهُنَّ وَلَدٌ}'
        });
        totalShares += shareValue;
    }

    if (heirs.wives > 0) {
        const share = hasBranch ? '1/8' : '1/4';
        const shareValue = hasBranch ? 1 / 8 : 1 / 4;
        shares.push({
            heir: heirs.wives > 1 ? 'الزوجات' : 'الزوجة',
            count: heirs.wives,
            share: share + (heirs.wives > 1 ? ' (يقتسمن)' : ''),
            shareValue: shareValue,
            type: 'فرض',
            evidence: hasBranch
                ? 'قال تعالى: {فَإِن كَانَ لَكُمْ وَلَدٌ فَلَهُنَّ الثُّمُنُ مِمَّا تَرَكْتُم}'
                : 'قال تعالى: {وَلَهُنَّ الرُّبُعُ مِمَّا تَرَكْتُمْ إِن لَّمْ يَكُن لَّكُمْ وَلَدٌ}'
        });
        totalShares += shareValue;
    }

    // 2. الأب
    if (heirs.father > 0) {
        if (hasMaleBranch) {
            // فرض فقط (السدس)
            shares.push({
                heir: 'الأب',
                count: 1,
                share: '1/6',
                shareValue: 1 / 6,
                type: 'فرض',
                evidence: 'قال تعالى: {وَلِأَبَوَيْهِ لِكُلِّ وَاحِدٍ مِّنْهُمَا السُّدُسُ مِمَّا تَرَكَ إِن كَانَ لَهُ وَلَدٌ}'
            });
            totalShares += 1 / 6;
        } else if (heirs.daughters > 0 && !hasMaleBranch) {
            // فرض + تعصيب
            shares.push({
                heir: 'الأب',
                count: 1,
                share: '1/6 + الباقي تعصيباً',
                shareValue: 1 / 6,
                type: 'فرض + تعصيب',
                evidence: 'يأخذ السدس فرضاً لوجود الفرع الوارث، والباقي تعصيباً لعدم وجود الذكور من الفرع'
            });
            totalShares += 1 / 6;
        } else {
            // تعصيب فقط (يأخذ الباقي)
            shares.push({
                heir: 'الأب',
                count: 1,
                share: 'الباقي تعصيباً',
                shareValue: 0,
                type: 'تعصيب',
                evidence: 'قال ﷺ: "ألحقوا الفرائض بأهلها فما بقي فلأولى رجل ذكر"'
            });
        }
    }

    // 3. الأم
    if (heirs.mother > 0) {
        let share, shareValue, evidence;
        const hasMultipleSiblings = (heirs.brothersFull + heirs.sistersFull + heirs.brothersFather + heirs.sistersFather + heirs.brothersMother + heirs.sistersMother) >= 2;

        if (hasBranch || hasMultipleSiblings) {
            share = '1/6';
            shareValue = 1 / 6;
            evidence = 'قال تعالى: {فَإِن كَانَ لَهُ إِخْوَةٌ فَلِأُمِّهِ السُّدُسُ}';
        } else if (!hasBranch && (heirs.husband > 0 || heirs.wives > 0) && heirs.father > 0) {
            // العمريتان - للأم ثلث الباقي بعد فرض الزوج/الزوجة (وليس ثلث كامل التركة)
            const spouseShare = heirs.husband > 0 ? 1 / 2 : 1 / 4;
            share = '1/3 الباقي';
            shareValue = (1 - spouseShare) / 3;
            evidence = 'هذه إحدى العمريتين: للأم ثلث الباقي بعد فرض الزوج/الزوجة، قضى بها عمر رضي الله عنه';
        } else {
            share = '1/3';
            shareValue = 1 / 3;
            evidence = 'قال تعالى: {فَإِن لَّمْ يَكُن لَّهُ وَلَدٌ وَوَرِثَهُ أَبَوَاهُ فَلِأُمِّهِ الثُّلُثُ}';
        }

        shares.push({
            heir: 'الأم',
            count: 1,
            share: share,
            shareValue: shareValue,
            type: 'فرض',
            evidence: evidence
        });
        totalShares += shareValue;
    }

    // 4. الجد (إذا لم يُحجب)
    if (heirs.grandfather > 0 && heirs.father === 0) {
        if (hasMaleBranch) {
            shares.push({
                heir: 'الجد',
                count: 1,
                share: '1/6',
                shareValue: 1 / 6,
                type: 'فرض',
                evidence: 'الجد يقوم مقام الأب عند عدمه، فله السدس مع الفرع الوارث الذكر'
            });
            totalShares += 1 / 6;
        } else if (heirs.daughters > 0) {
            shares.push({
                heir: 'الجد',
                count: 1,
                share: '1/6 + الباقي تعصيباً',
                shareValue: 1 / 6,
                type: 'فرض + تعصيب',
                evidence: 'الجد يأخذ السدس فرضاً مع البنات، والباقي تعصيباً'
            });
            totalShares += 1 / 6;
        } else {
            shares.push({
                heir: 'الجد',
                count: 1,
                share: 'الباقي تعصيباً',
                shareValue: 0,
                type: 'تعصيب',
                evidence: 'الجد يأخذ الباقي تعصيباً عند عدم الفرع الوارث'
            });
        }
    }

    // 5. الجدات
    const grandmothers = heirs.grandmotherMother + heirs.grandmotherFather;
    if (grandmothers > 0 && heirs.mother === 0) {
        shares.push({
            heir: grandmothers > 1 ? 'الجدات' : 'الجدة',
            count: grandmothers,
            share: '1/6' + (grandmothers > 1 ? ' (يقتسمن)' : ''),
            shareValue: 1 / 6,
            type: 'فرض',
            evidence: 'للجدة أو الجدات السدس، قضى به أبو بكر الصديق رضي الله عنه'
        });
        totalShares += 1 / 6;
    } else if (grandmothers > 0 && heirs.mother > 0) {
        blocked.push({ heir: 'الجدة/الجدات', blockedBy: 'الأم', reason: 'الأم تحجب الجدات' });
    }

    // 6. البنات
    if (heirs.daughters > 0 && heirs.sons === 0) {
        if (heirs.daughters === 1) {
            shares.push({
                heir: 'البنت',
                count: 1,
                share: '1/2',
                shareValue: 1 / 2,
                type: 'فرض',
                evidence: 'قال تعالى: {وَإِن كَانَتْ وَاحِدَةً فَلَهَا النِّصْفُ}'
            });
            totalShares += 1 / 2;
        } else {
            shares.push({
                heir: 'البنات',
                count: heirs.daughters,
                share: '2/3 (يقتسمن)',
                shareValue: 2 / 3,
                type: 'فرض',
                evidence: 'قال تعالى: {فَإِن كُنَّ نِسَاءً فَوْقَ اثْنَتَيْنِ فَلَهُنَّ ثُلُثَا مَا تَرَكَ}'
            });
            totalShares += 2 / 3;
        }
    }

    // 7. الأبناء والبنات (تعصيب)
    if (heirs.sons > 0) {
        const totalChildren = heirs.sons * 2 + heirs.daughters;
        shares.push({
            heir: 'الأبناء' + (heirs.daughters > 0 ? ' والبنات' : ''),
            count: heirs.sons + heirs.daughters,
            share: 'الباقي تعصيباً (للذكر مثل حظ الأنثيين)',
            shareValue: 0,
            type: 'تعصيب',
            evidence: 'قال تعالى: {يُوصِيكُمُ اللَّهُ فِي أَوْلَادِكُمْ لِلذَّكَرِ مِثْلُ حَظِّ الْأُنثَيَيْنِ}'
        });
    }

    // 8. بنات الابن
    if (heirs.granddaughtersSon > 0 && heirs.sons === 0) {
        if (heirs.daughters === 0) {
            if (heirs.granddaughtersSon === 1) {
                shares.push({
                    heir: 'بنت الابن',
                    count: 1,
                    share: '1/2',
                    shareValue: 1 / 2,
                    type: 'فرض',
                    evidence: 'بنت الابن تأخذ النصف عند انفرادها وعدم وجود البنت'
                });
                totalShares += 1 / 2;
            } else {
                shares.push({
                    heir: 'بنات الابن',
                    count: heirs.granddaughtersSon,
                    share: '2/3 (يقتسمن)',
                    shareValue: 2 / 3,
                    type: 'فرض',
                    evidence: 'بنات الابن لهن الثلثان عند عدم وجود البنات'
                });
                totalShares += 2 / 3;
            }
        } else if (heirs.daughters === 1) {
            shares.push({
                heir: 'بنت/بنات الابن',
                count: heirs.granddaughtersSon,
                share: '1/6 (تكملة الثلثين)',
                shareValue: 1 / 6,
                type: 'فرض',
                evidence: 'بنات الابن لهن السدس تكملة الثلثين مع البنت الواحدة'
            });
            totalShares += 1 / 6;
        }
    }

    // 9. أبناء الابن مع بنات الابن (تعصيب)
    if (heirs.grandsonsSon > 0 && heirs.sons === 0) {
        shares.push({
            heir: 'أبناء الابن' + (heirs.granddaughtersSon > 0 ? ' وبنات الابن' : ''),
            count: heirs.grandsonsSon + heirs.granddaughtersSon,
            share: 'الباقي تعصيباً',
            shareValue: 0,
            type: 'تعصيب',
            evidence: 'أبناء الابن يعصبون بنات الابن، للذكر مثل حظ الأنثيين'
        });
        // Remove daughters of son from fard if they were added
        const dgsIndex = shares.findIndex(s => s.heir.includes('بنت') && s.heir.includes('الابن'));
        if (dgsIndex > -1) {
            totalShares -= shares[dgsIndex].shareValue;
            shares.splice(dgsIndex, 1);
        }
    }

    // 10. الأخوات الشقيقات
    if (heirs.sistersFull > 0 && heirs.father === 0 && !hasBranch) {
        if (heirs.brothersFull > 0) {
            // تعصيب بالغير
        } else {
            if (heirs.sistersFull === 1) {
                shares.push({
                    heir: 'الأخت الشقيقة',
                    count: 1,
                    share: '1/2',
                    shareValue: 1 / 2,
                    type: 'فرض',
                    evidence: 'قال تعالى: {إِنِ امْرُؤٌ هَلَكَ لَيْسَ لَهُ وَلَدٌ وَلَهُ أُخْتٌ فَلَهَا نِصْفُ مَا تَرَكَ}'
                });
                totalShares += 1 / 2;
            } else {
                shares.push({
                    heir: 'الأخوات الشقيقات',
                    count: heirs.sistersFull,
                    share: '2/3 (يقتسمن)',
                    shareValue: 2 / 3,
                    type: 'فرض',
                    evidence: 'قال تعالى: {فَإِن كَانَتَا اثْنَتَيْنِ فَلَهُمَا الثُّلُثَانِ مِمَّا تَرَكَ}'
                });
                totalShares += 2 / 3;
            }
        }
    }

    // 11. الإخوة والأخوات الأشقاء (تعصيب)
    if (heirs.brothersFull > 0 && heirs.father === 0 && !hasBranch) {
        shares.push({
            heir: 'الإخوة الأشقاء' + (heirs.sistersFull > 0 ? ' والأخوات الشقيقات' : ''),
            count: heirs.brothersFull + heirs.sistersFull,
            share: 'الباقي تعصيباً',
            shareValue: 0,
            type: 'تعصيب',
            evidence: 'الإخوة الأشقاء عصبة بالنفس، يعصبون أخواتهم'
        });
    }

    // 12. الأخوات لأب
    if (heirs.sistersFather > 0 && heirs.father === 0 && !hasBranch && heirs.brothersFull === 0) {
        if (heirs.sistersFull === 0 && heirs.brothersFather === 0) {
            if (heirs.sistersFather === 1) {
                shares.push({
                    heir: 'الأخت لأب',
                    count: 1,
                    share: '1/2',
                    shareValue: 1 / 2,
                    type: 'فرض',
                    evidence: 'الأخت لأب تأخذ النصف عند عدم وجود الأشقاء'
                });
                totalShares += 1 / 2;
            } else {
                shares.push({
                    heir: 'الأخوات لأب',
                    count: heirs.sistersFather,
                    share: '2/3 (يقتسمن)',
                    shareValue: 2 / 3,
                    type: 'فرض',
                    evidence: 'الأخوات لأب لهن الثلثان عند عدم وجود الأشقاء'
                });
                totalShares += 2 / 3;
            }
        } else if (heirs.sistersFull === 1) {
            shares.push({
                heir: 'الأخت/الأخوات لأب',
                count: heirs.sistersFather,
                share: '1/6 (تكملة الثلثين)',
                shareValue: 1 / 6,
                type: 'فرض',
                evidence: 'الأخوات لأب لهن السدس تكملة الثلثين مع الأخت الشقيقة'
            });
            totalShares += 1 / 6;
        }
    }

    // 13. الإخوة والأخوات لأب (تعصيب)
    if (heirs.brothersFather > 0 && heirs.father === 0 && !hasBranch && heirs.brothersFull === 0 && heirs.sistersFull < 2) {
        shares.push({
            heir: 'الإخوة لأب' + (heirs.sistersFather > 0 ? ' والأخوات لأب' : ''),
            count: heirs.brothersFather + heirs.sistersFather,
            share: 'الباقي تعصيباً',
            shareValue: 0,
            type: 'تعصيب',
            evidence: 'الإخوة لأب عصبة، يعصبون أخواتهم لأب'
        });
    }

    // 14. الإخوة والأخوات لأم
    const siblingsMotherCount = heirs.brothersMother + heirs.sistersMother;
    if (siblingsMotherCount > 0 && heirs.father === 0 && heirs.grandfather === 0 && !hasBranch) {
        if (siblingsMotherCount === 1) {
            shares.push({
                heir: heirs.brothersMother > 0 ? 'الأخ لأم' : 'الأخت لأم',
                count: 1,
                share: '1/6',
                shareValue: 1 / 6,
                type: 'فرض',
                evidence: 'قال تعالى: {وَإِن كَانَ رَجُلٌ يُورَثُ كَلَالَةً أَوِ امْرَأَةٌ وَلَهُ أَخٌ أَوْ أُخْتٌ فَلِكُلِّ وَاحِدٍ مِّنْهُمَا السُّدُسُ}'
            });
            totalShares += 1 / 6;
        } else {
            shares.push({
                heir: 'الإخوة والأخوات لأم',
                count: siblingsMotherCount,
                share: '1/3 (يقتسمون بالسوية)',
                shareValue: 1 / 3,
                type: 'فرض',
                evidence: 'قال تعالى: {فَإِن كَانُوا أَكْثَرَ مِن ذَٰلِكَ فَهُمْ شُرَكَاءُ فِي الثُّلُثِ}'
            });
            totalShares += 1 / 3;
        }
    }

    // باقي العصبات (أبناء الإخوة، الأعمام، أبناء الأعمام)
    const remainingAsaba = [];
    if (heirs.nephewsFull > 0) remainingAsaba.push({ name: 'أبناء الأخ الشقيق', count: heirs.nephewsFull, order: 1 });
    if (heirs.nephewsFather > 0) remainingAsaba.push({ name: 'أبناء الأخ لأب', count: heirs.nephewsFather, order: 2 });
    if (heirs.unclesFull > 0) remainingAsaba.push({ name: 'العم الشقيق', count: heirs.unclesFull, order: 3 });
    if (heirs.unclesFather > 0) remainingAsaba.push({ name: 'العم لأب', count: heirs.unclesFather, order: 4 });
    if (heirs.cousinsFull > 0) remainingAsaba.push({ name: 'ابن العم الشقيق', count: heirs.cousinsFull, order: 5 });
    if (heirs.cousinsFather > 0) remainingAsaba.push({ name: 'ابن العم لأب', count: heirs.cousinsFather, order: 6 });

    // Add nearest asaba if no closer asaba exists
    const hasCloserAsaba = heirs.sons > 0 || heirs.grandsonsSon > 0 || heirs.father > 0 || heirs.grandfather > 0 || heirs.brothersFull > 0 || heirs.brothersFather > 0;
    if (!hasCloserAsaba && remainingAsaba.length > 0) {
        remainingAsaba.sort((a, b) => a.order - b.order);
        const nearest = remainingAsaba[0];
        shares.push({
            heir: nearest.name,
            count: nearest.count,
            share: 'الباقي تعصيباً',
            shareValue: 0,
            type: 'تعصيب',
            evidence: 'عصبة بالنفس يأخذ الباقي بعد أصحاب الفروض'
        });

        // Block further asaba
        for (let i = 1; i < remainingAsaba.length; i++) {
            blocked.push({ heir: remainingAsaba[i].name, blockedBy: nearest.name, reason: 'العاصب الأقرب يحجب الأبعد' });
        }
    }

    // Check for العول (if shares exceed 1)
    let awl = null;
    if (totalShares > 1) {
        awl = {
            original: 1,
            adjusted: totalShares,
            explanation: 'المسألة عائلة: مجموع الفروض يزيد عن أصل المسألة، فيُنقص من كل ذي فرض بقدر نسبته'
        };
    }

    // Check for الرد (if shares less than 1 and no asaba)
    let radd = null;
    const hasAsaba = shares.some(s => s.type === 'تعصيب' || s.type === 'فرض + تعصيب');
    if (totalShares < 1 && !hasAsaba) {
        radd = {
            remaining: 1 - totalShares,
            explanation: 'الرد: الباقي يُرد على أصحاب الفروض (عدا الزوجين) بنسبة فروضهم'
        };
    }

    return {
        shares,
        blocked,
        totalShares,
        awl,
        radd,
        baseValue,
        originalHeirs
    };
}

// ==================== توزيع مبالغ التركة ====================
// يعيد لكل مجموعة في result.shares مبلغها الإجمالي بالريال (بمحاذاة الفهارس)، ويعالج:
// - العول: إنقاص كل الفروض بنسبها عند تجاوز مجموعها الواحد.
// - الباقي التعصيبي: للعصبة المحضة، وإلا فلصاحب «فرض + تعصيب» (كالأب مع البنات).
// - الرد: عند غياب العصبة يُرد الباقي على أصحاب الفروض (عدا الزوجين) بنسبة فروضهم.
function computeInheritanceAmounts(result, estateValue) {
    const shares = result && Array.isArray(result.shares) ? result.shares : [];
    const estate = Number(estateValue) || 0;
    if (estate <= 0) return shares.map(() => 0);

    const fardTotal = shares.reduce((sum, s) => sum + (s.shareValue > 0 ? s.shareValue : 0), 0);

    // العول: لا باقي بعده، وتُنقص الفروض كلها بنسبها
    if (fardTotal > 1) {
        return shares.map(s => (s.shareValue > 0 ? estate * (s.shareValue / fardTotal) : 0));
    }

    const amounts = shares.map(s => (s.shareValue > 0 ? estate * s.shareValue : 0));
    const remainder = estate * (1 - fardTotal);
    if (remainder <= 0) return amounts;

    // العصبة المحضة أولى بالباقي (الفرع الوارث)، ثم صاحب الفرض والتعصيب
    let asabaIdx = shares.findIndex(s => s.type === 'تعصيب');
    if (asabaIdx === -1) asabaIdx = shares.findIndex(s => s.type === 'فرض + تعصيب');
    if (asabaIdx > -1) {
        amounts[asabaIdx] += remainder;
        return amounts;
    }

    // الرد على أصحاب الفروض عدا الزوجين؛ فإن لم يوجد غيرهما بقي الباقي دون توزيع
    const isSpouse = s => s.heir === 'الزوج' || s.heir === 'الزوجة' || s.heir === 'الزوجات';
    const raddBase = shares.reduce((sum, s) => sum + (s.shareValue > 0 && !isSpouse(s) ? s.shareValue : 0), 0);
    if (raddBase > 0) {
        shares.forEach((s, i) => {
            if (s.shareValue > 0 && !isSpouse(s)) amounts[i] += remainder * (s.shareValue / raddBase);
        });
    }
    return amounts;
}

// ==================== END OF SERVICE ====================
// حساب فرق مدة الخدمة بين تاريخين (سنوات/أشهر/أيام)
function dateDiffYMD(startVal, endVal) {
    const start = new Date(startVal);
    const end = new Date(endVal);
    if (end <= start) return { error: 'تاريخ انتهاء الخدمة يجب أن يكون بعد تاريخ بدايتها' };
    let y = end.getFullYear() - start.getFullYear();
    let m = end.getMonth() - start.getMonth();
    let d = end.getDate() - start.getDate();
    if (d < 0) { d += 30; m--; }
    if (m < 0) { m += 12; y--; }
    return { years: y, months: m, days: d };
}

// حساب مكافأة نهاية الخدمة (المواد 84، 85، 87، 80 من نظام العمل)
function computeEOSCore(wage, reason, duration) {
    if (!wage || wage <= 0 || !duration) return null;
    if (duration.error) return { error: duration.error };

    // إجمالي المدة بالسنوات (كسر السنة بنسبته: الشهر 1/12 واليوم 1/360)
    const totalYears = duration.years + duration.months / 12 + duration.days / 360;

    // المادة 84: نصف شهر عن كل سنة من الخمس الأولى، وشهر عن كل سنة تالية
    const firstFive = Math.min(totalYears, 5) * 0.5 * wage;
    const afterFive = Math.max(totalYears - 5, 0) * wage;
    const fullAward = firstFive + afterFive;

    let ratio = 1;
    let ratioLabel = 'كاملة (100%)';
    let ruleNote = 'يستحق العامل المكافأة كاملة (المادة 84 من نظام العمل)';

    if (reason === 'resignation') {
        if (totalYears < 2) {
            ratio = 0;
            ratioLabel = 'لا يستحق';
            ruleNote = 'لا يستحق العامل المستقيل مكافأة لأن مدة خدمته أقل من سنتين (المادة 85)';
        } else if (totalYears < 5) {
            ratio = 1 / 3;
            ratioLabel = 'الثلث (33.33%)';
            ruleNote = 'يستحق العامل المستقيل ثلث المكافأة لأن مدة خدمته من سنتين إلى خمس سنوات (المادة 85)';
        } else if (totalYears < 10) {
            ratio = 2 / 3;
            ratioLabel = 'الثلثان (66.67%)';
            ruleNote = 'يستحق العامل المستقيل ثلثي المكافأة لأن مدة خدمته من خمس إلى عشر سنوات (المادة 85)';
        } else {
            ruleNote = 'يستحق العامل المستقيل المكافأة كاملة لأن مدة خدمته عشر سنوات فأكثر (المادة 85)';
        }
    } else if (reason === 'art80') {
        ratio = 0;
        ratioLabel = 'لا يستحق';
        ruleNote = 'فسخ صاحب العمل للعقد لإحدى حالات المادة (80) يكون دون مكافأة';
    } else if (reason === 'art81') {
        ruleNote = 'يستحق العامل المكافأة كاملة لفسخه العقد لإحدى حالات المادة (81) — المادة 87';
    } else if (reason === 'force') {
        ruleNote = 'يستحق العامل المكافأة كاملة لتركه العمل نتيجة قوة قاهرة (المادة 87)';
    } else if (reason === 'marriage') {
        ruleNote = 'تستحق العاملة المكافأة كاملة لإنهائها العقد خلال ستة أشهر من الزواج أو ثلاثة أشهر من الوضع (المادة 87)';
    }

    return {
        wage, duration, totalYears, firstFive, afterFive, fullAward,
        ratio, ratioLabel, ruleNote,
        dueAward: fullAward * ratio
    };
}

// تصدير للاختبارات في بيئة Node (لا يؤثر على المتصفح)
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        currencies,
        numberToArabicWords,
        tafqitAmount,
        tafqitSAR,
        computeInheritance,
        computeInheritanceAmounts,
        dateDiffYMD,
        computeEOSCore
    };
}
