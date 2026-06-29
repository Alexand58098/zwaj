# ⚽ TifoTV

موقع لمشاهدة beIN Sports + SSC + الرياضية + 50+ قناة رياضية 24/7

## 🚀 ثلاث طرق للتشغيل

---

### 🥇 الطريقة الأسهل: رفع على Netlify (مجاناً، أوصي بها)

**الميزة الكبرى**: لا تحتاج Worker منفصل! كل شيء يعمل تلقائياً.

#### الخطوات (3 دقائق):

1. **سجّل في Netlify مجاناً**: https://app.netlify.com/signup
   - يمكنك التسجيل بـ GitHub / GitLab / Email

2. **اسحب وأفلت المجلد**:
   - افتح: https://app.netlify.com/drop
   - **اسحب كامل مجلد المشروع** (يحتوي `index.html`, `netlify.toml`, `netlify/`) إلى الصفحة
   - انتظر 10-20 ثانية حتى تنتهي عملية النشر

3. **خلاص!** ستحصل على رابط مثل `https://amazing-name-123.netlify.app`
   - افتحه → كل القنوات تعمل تلقائياً ✓
   - لا حاجة لإعداد أي Worker، الـ Function الداخلية تتولى الأمر

---

### 🥈 الطريقة الثانية: Cloudflare Worker (إذا لا تريد رفع الموقع)

استخدم هذه الطريقة إذا تريد الموقع محلياً على جهازك.

#### الخطوات:
1. اتبع التعليمات في `SETUP-WORKER.md`
2. أنشئ Worker على cloudflare.com
3. الصق `worker.js` فيه
4. ضع رابط الـ Worker في الشريط الأصفر بالموقع

---

### 🥉 الطريقة الثالثة: بدون أي إعداد (محدودة)

افتح `index.html` مباشرة → ستعمل **6 قنوات حرة** فقط:
- ⭐ **Arryadia HD** (الناقل الرسمي للمباراة!)
- Al Aoula HD, Al Maghribia HD, Tamazight HD
- beIN SPORTS XTRA HD
- beIN Sports XTRA Español

---

## 📁 ملفات المشروع

| الملف | الوصف |
|------|--------|
| `index.html` | الموقع الرئيسي |
| `netlify.toml` | إعدادات Netlify |
| `netlify/functions/bein.js` | الـ Function التي تتجاوز الحجب (Netlify) |
| `worker.js` | كود Cloudflare Worker (بديل لـ Netlify) |
| `SETUP-WORKER.md` | دليل إنشاء Worker على Cloudflare |
| `README.md` | هذا الملف |

---

## 📺 القنوات المتوفرة (54+ قناة)

- **beIN Sports Arabic**: 12 قناة (beIN 1-9, XTRA, MAX AR, HD Qatar)
- **beIN Sports France**: 10 قنوات (beIN 1-3 + MAX 4-10)
- **beIN Sports MENA English**: 2
- **beIN Sports Türkiye**: 5
- **beIN International**: USA, Español, Australia, Malaysia
- **SSC السعودية**: SSC 1-5 + Extra 1-3
- **خليجي**: Dubai 1-3, Abu Dhabi 1-2, Premium 1-2
- **مغربية حرة**: Arryadia, Al Aoula, Al Maghribia, Tamazight

---

## ⚡ خصائص

- 🎨 تصميم احترافي بـ glassmorphism
- 🔍 بحث فوري وتبويبات حسب اللغة
- 📊 منتقي جودة (Auto/1080p/720p/480p)
- ⌨️ اختصارات لوحة المفاتيح (`F` ملء شاشة، `M` كتم، `/` بحث)
- 📱 متجاوب كلياً (موبايل + تابلت + ديسكتوب)
- 🌙 وضع مظلم بألوان المنتخبين
- ⏰ عدّاد تنازلي للمباراة
