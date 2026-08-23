// ==================== أدوات نصية مشتركة بين صفحات التطبيق ====================
// يُحمَّل هذا الملف قبل بقية السكربتات في كل صفحة (index / minutes / sakk)،
// ويُستورد في Node عبر require للاختبارات ولاستهلاك js/minutes.js له.

// حماية من التأطير (Clickjacking): GitHub Pages لا يتيح رؤوس استجابة،
// وتوجيه frame-ancestors لا يعمل من وسم <meta>، فيُكتفى بهذا الفحص —
// نسخة Vercel محمية أصلًا برأس CSP في vercel.json.
if (typeof window !== 'undefined' && window.top !== window.self) {
    try {
        window.top.location.replace(window.location.href);
    } catch (e) {
        // إطار من أصل آخر يمنع التوجيه: تُفرَّغ الصفحة وتُعرض رسالة (بـ textContent لا HTML)
        document.documentElement.textContent = '';
        const msg = document.createElement('p');
        msg.dir = 'rtl';
        msg.style.cssText = 'font-family: sans-serif; text-align: center; padding: 3rem;';
        msg.textContent = 'لا يمكن عرض هذه الصفحة داخل إطار خارجي — افتحها مباشرة من الرابط: ' + window.location.href;
        document.documentElement.appendChild(msg);
    }
}

// تحويل الأرقام العربية المشرقية إلى أرقام لاتينية
function convertArabicDigits(str) {
    const arabicDigits = '٠١٢٣٤٥٦٧٨٩';
    return String(str).replace(/[٠-٩]/g, d => String(arabicDigits.indexOf(d)));
}

const TASHKEEL_RE = /[\u064B-\u065F\u0670\u0640]/g;

// تطبيع نص عربي للبحث: إسقاط التشكيل وتوحيد الهمزات والتاء المربوطة
// والألف المقصورة والأرقام العربية — تستهلكه صفحة النماذج ومولّد الجلسة معًا
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

// تصدير للاختبارات ولبقية المحركات في بيئة Node (لا يؤثر على المتصفح)
if (typeof module !== 'undefined' && module.exports) {
    module.exports = { convertArabicDigits, normalizeArabicSearch };
}
