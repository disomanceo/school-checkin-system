# การเชื่อม GitHub

ตอนนี้โปรเจกต์มีไฟล์สำหรับระบบ School Checkin แล้ว แต่ Git repository ในเครื่องยังตั้งค่าไม่สำเร็จ เพราะโฟลเดอร์ `.git` ถูกจำกัดสิทธิ์เขียนในสภาพแวดล้อมนี้

## ทางเลือกที่ 1: เชื่อมผ่าน Git ปกติ

ใช้เมื่อเปิด Terminal ในเครื่องของคุณเอง:

```text
git init
git add .
git commit -m "Initial school checkin integration"
git branch -M main
git remote add origin https://github.com/OWNER/REPOSITORY.git
git push -u origin main
```

ให้เปลี่ยน `OWNER/REPOSITORY` เป็น repo จริง เช่น:

```text
https://github.com/your-name/school-checkin-system.git
```

## ทางเลือกที่ 2: ให้ Codex อัปโหลดผ่าน GitHub connector

ส่งชื่อ repo มาในรูปแบบ:

```text
OWNER/REPOSITORY
```

หรือส่งลิงก์:

```text
https://github.com/OWNER/REPOSITORY
```

จากนั้น Codex จะสร้างไฟล์เหล่านี้ใน repo ให้:

- `README.md`
- `.gitignore`
- `.clasp.json`
- `GITHUB_SETUP.md`
- `google-apps-script/Code.gs`
- `google-apps-script/appsscript.json`
- `dashboard-example/index.html`

## สิ่งที่ต้องระวัง

- ตรวจว่า repo เป็น private ถ้าข้อมูลระบบโรงเรียนยังไม่พร้อมเปิดเผย
- ห้าม commit ค่า token, password, หรือไฟล์ `.env`
- `.clasp.json` มี Script ID ซึ่งไม่ใช่รหัสผ่าน แต่ควรใช้ใน repo ที่คนเกี่ยวข้องเข้าถึงได้เท่านั้น
- Web App URL ใน Dashboard เป็น endpoint ที่ deploy แล้ว ถ้าปรับสิทธิ์ Apps Script ต้องทดสอบใหม่ทุกครั้ง
