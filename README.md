# لوحة مؤشرات المنشآت الاقتصادية — فلسطين

تطبيق Dash تفاعلي باللغة العربية لعرض مؤشرات المنشآت الاقتصادية في فلسطين.

## تشغيل محلي

```bash
pip install -r requirements.txt
python dashboard.py
```

ثم افتح التطبيق على المنفذ المحلي الظاهر في الطرفية.

## النشر العام على Render

- **Language:** Python 3
- **Build Command:** `pip install -r requirements.txt`
- **Start Command:** `gunicorn dashboard:server`
- **Root Directory:** اتركه فارغًا إذا كانت الملفات في جذر المستودع.

ملف Excel `Book6.10.xlsx` يجب أن يبقى في نفس مجلد `dashboard.py`.
