# تثبيت PPTK للماك __PPTK_VERSION__ يدويًا (لفريق تقنية المعلومات)

هذا الملف المضغوط يثبّت PPTK على جهاز ماك واحد. فيه **نفس حزمة PKG اللي ينشرها Intune**، فالجهاز اللي يتثبّت يدويًا يستقبل كل التحديثات بعدين من Intune، بدون تثبيت مكرر وبدون ما تعيد التثبيت يدويًا.

## وش يتثبّت

| العنصر | المكان |
|---|---|
| مساعد الاختصارات | `/Applications/PPTK Keys.app` |
| تشغيل المساعد عند دخول أي مستخدم | `/Library/LaunchAgents/com.pptk.keys.plist` |
| أداة PowerPoint، وتنحمّل لكل المستخدمين | `/Library/Application Support/Microsoft/Office365/User Content.localized/Startup/PowerPoint/PPTK.ppam` |
| ملف الإصدار وسكربت الإيميل وسكربتات الإزالة والفحص | `/Library/Application Support/PPTK/` |

## المتطلبات

- macOS 13 أو أحدث (Apple silicon أو Intel).
- Microsoft 365 PowerPoint في مجلد `/Applications`.
- حساب مسؤول (Administrator).

## التثبيت

1. انقل الملف المضغوط للجهاز عن طريق OneDrive أو Teams أو فلاش، وحطه في `/Users/Shared`.
2. اضغط عليه مرتين عشان ينفك الضغط.
3. اضغط مرتين على **Install PPTK.command**.
   - إذا قال macOS إنه ما يقدر يفتحه: روح لـ System Settings ← Privacy & Security، ودوّر على الرسالة الخاصة بـ «Install PPTK.command»، واضغط **Open Anyway** وأكّد، ثم افتح الملف مرة ثانية. هذي الطريقة الرسمية من Apple، ولا تحذف علامة التحميل من الملفات.
   - أو من Terminal: `cd "/Users/Shared/PPTK-Mac-__PPTK_VERSION__-Installer" && bash "Install PPTK.command"`
4. اكتب كلمة مرور المسؤول. السكربت يتحقق من الحزمة ثم يثبّتها ثم يفحص التثبيت.
5. الموظف يوافق مرة وحدة في System Settings ← Privacy & Security:
   - **Accessibility**: فعّل **PPTK Keys**.
   - **Input Monitoring**: فعّل **PPTK Keys**.
6. **اقفل PowerPoint بـ ⌘Q وافتحه من جديد**، ويطلع تبويب **PPTK**. أول مرة يستخدم فيها اختصار، يسأله macOS إذا يسمح لـ PPTK Keys بالتحكم في PowerPoint، ويضغط **OK**.

## الفحص

```bash
sudo "/Library/Application Support/PPTK/pptk-detect.sh"
```

النتيجة `PPTK __PPTK_VERSION__ installed` معناها التثبيت كامل.

## التحديثات

التحديثات تجي من Intune (التفاصيل في `INTUNE-MAC.md`).
- PowerPoint ما ينقفل أبدًا: إذا كان مفتوح يكمل بالنسخة الحالية، والأداة الجديدة تشتغل أول ما يُفتح PowerPoint مرة ثانية.
- تثبيت إصدار أقدم فوق إصدار أحدث يُرفض بدون أي تغيير.
- الصلاحيات تبقى إذا كان التحديث موقّع من نفس فريق Developer ID.

## الإزالة

اضغط مرتين على **Uninstall PPTK.command**. يحذف المساعد والأداة والملفات بدون ما يلمس أي إضافة ثانية في PowerPoint.
