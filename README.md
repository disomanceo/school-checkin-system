# เชื่อม Google Sheets กับระบบ School Checkin

ชุดไฟล์นี้ใช้สำหรับดึงและจัดเก็บโค้ด Google Apps Script ของระบบ `School Checkin System` ลงเครื่อง ก่อนนำไปจัดการต่อด้วย Git/GitHub

## โฟลเดอร์สำคัญ

- `gas-current/` คือโค้ดจริงที่ดึงจาก Google Apps Script ที่กำลังใช้งานอยู่
- `google-apps-script/` คือโค้ด API ตัวอย่างที่สร้างไว้ก่อนหน้า ใช้เป็น reference เท่านั้น
- `dashboard-example/` คือหน้า dashboard ตัวอย่างแยกจากระบบจริง

## โปรเจกต์ Apps Script

Script ID ที่เชื่อมไว้:

```text
1COLxqmeEreUywkaq5HeW3QJBzR-hLvDyVh_xbqb2lv6UMy7w6Zq6a8rq
```

มีไฟล์ `.clasp.json` สำหรับผูกโฟลเดอร์นี้กับโปรเจกต์ Apps Script แล้ว ถ้าเครื่องติดตั้ง `clasp` และล็อกอิน Google ไว้ สามารถอัปโหลดโค้ดด้วยคำสั่ง:

```text
clasp push
```

หมายเหตุ: Script ID คือรหัสโปรเจกต์ Apps Script ยังไม่ใช่ Web App URL สำหรับให้หน้าเว็บเรียกข้อมูล ต้อง Deploy เป็น Web app ก่อนจึงจะได้ URL รูปแบบ `https://script.google.com/macros/s/.../exec`

Web App URL ที่ได้รับ:

```text
https://script.google.com/macros/s/AKfycbzB0SKfVjK95eSOxbgRwNOEjv_wIJiHoh5xI972UGghcmay8v9aBE5Uqyj-YkwJKJLfVA/exec
```

ผลทดสอบล่าสุด: URL เปิดได้ แต่ deployment ปัจจุบันยังตอบกลับเป็นหน้าเดิมที่มีข้อความ `QR Code ไม่ถูกต้อง` แปลว่ายังไม่ได้ deploy โค้ด API จาก `google-apps-script/Code.gs` หรือ deployment ยังชี้เวอร์ชันเก่าอยู่

## โครงสร้างข้อมูลที่พบ

- `Users` เก็บผู้ใช้ บทบาท ตำแหน่ง สถานะบัญชี รูปโปรไฟล์ ลายเซ็น และเวลาเข้าระบบล่าสุด
- `CheckinData` เก็บประวัติเช็กอิน พิกัด ระยะ เวลาเข้า เวลาออกอัตโนมัติ สถานะเวลา และสถานะอนุมัติ
- `PermissionRequests` เก็บคำขอลา มาสาย ไปราชการ สถานะอนุมัติ ผู้อนุมัติ PDF และหลักฐาน
- `Settings` เก็บค่าระบบ เช่น พิกัดโรงเรียน รัศมี เวลาเริ่มงาน ปีการศึกษา ภาคเรียน และเวอร์ชันแอป
- `PositionSettings` เก็บเวลาออกอัตโนมัติตามตำแหน่ง

## วิธีติดตั้ง API ใน Google Apps Script

1. เปิดไฟล์ Google Sheets ที่ให้มา
2. ไปที่ `Extensions` > `Apps Script`
3. สร้างไฟล์ `Code.gs`
4. นำโค้ดจาก `google-apps-script/Code.gs` ไปวาง
5. กด `Deploy` > `New deployment`
6. เลือกชนิดเป็น `Web app`
7. ตั้งค่า `Execute as` เป็นบัญชีของคุณ
8. ตั้งค่า `Who has access` เป็น `Anyone with the link` หรือจำกัดตามนโยบายโรงเรียน
9. กด `Deploy` แล้วคัดลอก Web App URL

## วิธีอัปโหลดด้วย clasp ถ้าต้องการ

ถ้าต้องการส่งไฟล์จากโฟลเดอร์นี้ขึ้น Apps Script โดยตรง:

```text
clasp login
clasp push
```

หลังจาก push แล้วให้กลับไปที่หน้า Apps Script เพื่อ Deploy หรือ Manage deployments อีกครั้ง

## วิธีทดสอบดึงข้อมูล

เปิด Web App URL แล้วเติมพารามิเตอร์ท้ายลิงก์:

```text
?action=summary
?action=users&limit=20
?action=checkins&limit=50
?action=requests&limit=50
?action=requests&status=pending
?action=settings
```

ตัวอย่าง:

```text
https://script.google.com/macros/s/DEPLOYMENT_ID/exec?action=summary
```

## วิธีใช้หน้า Dashboard ตัวอย่าง

1. เปิดไฟล์ `dashboard-example/index.html`
2. หา `PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE`
3. แทนด้วย Web App URL ที่ได้จาก Apps Script
4. เปิดไฟล์นี้ในเบราว์เซอร์

หน้าเว็บจะเรียก `?action=summary` แล้วแสดงจำนวนผู้ใช้ รายการเช็กอิน คำขอทั้งหมด คำขอรออนุมัติ และรายการล่าสุด

## ข้อควรระวัง

- อย่าส่ง `PinHash` และ `PinSalt` ไปหน้าเว็บ ตัวอย่างนี้ตัดออกจาก API ผู้ใช้แล้ว
- ถ้าระบบใช้งานจริง ควรเพิ่ม token หรือระบบตรวจสิทธิ์ก่อนให้แก้ไขข้อมูล
- ถ้าต้องเพิ่มฟังก์ชันอนุมัติคำขอ ให้ทำเป็น `doPost` แยกจาก API อ่านข้อมูล
- ควรใช้ `ApprovalStatusCode` และ `TimeStatusCode` เป็นค่าหลักในระบบ เพราะอ่านง่ายกว่า emoji/status ภาษาไทย

## GitHub

ดูขั้นตอนเชื่อม GitHub เพิ่มเติมใน `GITHUB_SETUP.md`

## ดึงโค้ดจริงจาก GAS

ดูขั้นตอนใน `PULL_FROM_GAS.md`
