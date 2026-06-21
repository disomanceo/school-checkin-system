# ดึงโค้ดจริงจาก Google Apps Script ลงเครื่อง

เป้าหมายของขั้นตอนนี้คือดึงไฟล์ที่กำลังใช้งานจริงบน GAS ลงมาไว้ที่โฟลเดอร์ `gas-current/` ก่อน แล้วค่อยนำเข้า GitHub ต่อ

## Script ID

```text
1COLxqmeEreUywkaq5HeW3QJBzR-hLvDyVh_xbqb2lv6UMy7w6Zq6a8rq
```

## ขั้นตอน

เปิด Terminal หรือ PowerShell ที่โฟลเดอร์นี้:

```text
D:\kan la online
```

แล้วรันคำสั่งนี้:

```powershell
$env:npm_config_cache = "D:\kan la online\.npm-cache"
npx.cmd @google/clasp login --no-localhost --auth "D:\kan la online\.clasp-auth"
```

ระบบจะแสดงลิงก์ Google ให้เปิดใน browser:

1. เปิดลิงก์ที่แสดง
2. ล็อกอินด้วยบัญชี Google ที่มีสิทธิ์แก้ Apps Script นี้
3. อนุญาตสิทธิ์
4. คัดลอกโค้ดที่ Google ให้กลับมา
5. วางกลับใน Terminal

จากนั้นดึงโค้ดจริง:

```powershell
$env:npm_config_cache = "D:\kan la online\.npm-cache"
npx.cmd @google/clasp pull --auth "D:\kan la online\.clasp-auth" --project "D:\kan la online\gas-current"
```

เมื่อสำเร็จ ไฟล์จริงจาก Apps Script จะอยู่ใน:

```text
gas-current/
```

## ตรวจสอบหลัง pull

ควรเห็นไฟล์เช่น:

- `appsscript.json`
- ไฟล์ `.gs`
- ไฟล์ `.html` ถ้าโปรเจกต์ GAS มีหน้าเว็บ

## หมายเหตุ

- โฟลเดอร์ `google-apps-script/` คือโค้ดตัวอย่างที่สร้างไว้ก่อนหน้า
- โฟลเดอร์ `gas-current/` คือโค้ดจริงที่ดึงจาก GAS
- ไฟล์ login/cache เช่น `.clasp-auth/` และ `.npm-cache/` ถูกใส่ใน `.gitignore` แล้ว ไม่ควรส่งขึ้น GitHub
