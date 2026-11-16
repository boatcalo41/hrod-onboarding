# HROD — Onboarding & Offboarding System (Allied-Tek)

ระบบดูแลพนักงานใหม่ตั้งแต่วันแรกจนครบทดลองงาน (Onboarding) และกระบวนการเมื่อพนักงานออกจากงาน (Offboarding) — หน้าเว็บเดียว ไม่มี backend ของตัวเอง เก็บข้อมูลจริงไว้ใน **Google Sheet ของคุณเอง** ผ่าน **Google Apps Script**

```
[GitHub Pages: หน้าเว็บ index.html]  --fetch-->  [Google Apps Script Web App]  --read/write-->  [Google Sheet ของคุณ]
```

ข้อมูลพนักงานทั้งหมด (เช็คลิสต์ onboarding/offboarding, แบบสะท้อน ฯลฯ) อยู่ใน Google Sheet ในบัญชี Google ของคุณ — ไม่มีที่ไหนอื่นเก็บสำเนาไว้

## การตั้งค่าครั้งแรก (ทำครั้งเดียว)

### 1. สร้าง Google Sheet + Apps Script

1. ไปที่ [sheets.google.com](https://sheets.google.com) สร้างสเปรดชีตใหม่ ตั้งชื่ออะไรก็ได้ เช่น "HROD Database"
2. เมนู **Extensions (ส่วนขยาย) → Apps Script**
3. จะเปิดหน้าต่าง Apps Script editor พร้อมไฟล์ `Code.gs` ว่างๆ — ลบโค้ดเดิมทิ้งทั้งหมด แล้ววางโค้ดทั้งหมดจากไฟล์ [`apps-script/Code.gs`](apps-script/Code.gs) ในโปรเจกต์นี้ลงไปแทน
4. กด **Save** (ไอคอนแผ่นดิสก์ หรือ Ctrl+S)

### 2. Deploy เป็น Web App

1. มุมขวาบน กด **Deploy → New deployment**
2. ที่ "Select type" กดไอคอนเฟือง เลือก **Web app**
3. ตั้งค่า:
   - **Execute as:** Me (บัญชีของคุณ)
   - **Who has access:** **Anyone**
4. กด **Deploy**
5. ระบบจะขอ **Authorize access** — กด Authorize แล้วเลือกบัญชี Google ของคุณ จะมีหน้าเตือน "Google hasn't verified this app" (เพราะเป็นสคริปต์ที่คุณเขียนเอง ไม่ใช่แอปสาธารณะ) ให้กด **Advanced → Go to (ชื่อโปรเจกต์) (unsafe)** แล้ว Allow ได้เลย ปลอดภัยเพราะเป็นโค้ดของคุณเองที่รันในบัญชีคุณเอง
6. จะได้ **Web app URL** หน้าตาประมาณ:
   ```
   https://script.google.com/macros/s/AKfycby.../exec
   ```
   **คัดลอก URL นี้เก็บไว้** — ต้องใช้ในขั้นตอนถัดไป

> ทุกครั้งที่แก้โค้ดใน Code.gs แล้วอยากให้มีผล ต้องกด **Deploy → Manage deployments → แก้ไข (ไอคอนดินสอ) → Version: New version → Deploy** ใหม่ (แก้โค้ดเฉยๆ ไม่ deploy ใหม่จะไม่มีผล)

### 3. เปิดใช้งาน GitHub Pages

1. ไปที่ repo นี้บน GitHub → **Settings → Pages**
2. Source เลือก **Deploy from a branch** → Branch: **main** / Folder: **/(root)** → Save
3. รอสัก 1-2 นาที จะได้ลิงก์เว็บ เช่น `https://<username>.github.io/hrod-onboarding/`

### 4. เชื่อมเว็บเข้ากับ Apps Script

1. เปิดเว็บ HROD (ลิงก์จาก GitHub Pages)
2. กด **"⚙️ ตั้งค่าฐานข้อมูล"** ที่แถบด้านซ้าย
3. วาง **Web app URL** จากขั้นตอนที่ 2 ลงไป แล้วกด **บันทึกและเชื่อมต่อ**
4. ถ้าเชื่อมต่อสำเร็จ ไอคอนจะเปลี่ยนเป็น ✅ และจะเห็นพนักงานตัวอย่างที่เพิ่มเข้าไป (ถ้ามี) หรือหน้าว่างพร้อมให้กด "+ เพิ่มพนักงานใหม่"

> URL นี้เก็บไว้ใน localStorage ของเบราว์เซอร์แต่ละเครื่อง — **ทุกคนที่จะใช้ระบบ (HR, หัวหน้างาน, พนักงาน) ต้องกด "⚙️ ตั้งค่าฐานข้อมูล" แล้ววาง URL เดียวกันนี้ ในเบราว์เซอร์ของตัวเองครั้งแรกที่เปิดใช้** (เพราะเป็น static site ไม่มี server เก็บ config กลาง)

## ข้อจำกัดที่ควรรู้

- **ไม่ realtime แบบ 100%** — หน้าเว็บจะดึงข้อมูลใหม่ทุก ~4 วินาที (ไม่ใช่ push ทันทีเหมือนฐานข้อมูลแบบ Firestore) ดังนั้นถ้าคน A แก้ข้อมูล คน B จะเห็นการเปลี่ยนแปลงภายในไม่กี่วินาที ไม่ใช่ทันทีเป๊ะๆ
- **ไม่มีระบบ login/สิทธิ์จริง** — ใครก็ตามที่มี Web app URL จะอ่าน/เขียนข้อมูลได้ทั้งหมด (ตัวเลือก role ในแอปเป็นแค่ UI mode ช่วยจัดระเบียบหน้าจอ ไม่ใช่การยืนยันตัวตน) เหมาะกับทีมเล็กที่ไว้ใจกัน ไม่ควรแชร์ URL นี้แบบสาธารณะ
- **Google Apps Script มี quota** — บัญชี Google ฟรีมีโควตาการเรียกสคริปต์ต่อวันจำกัด (ปกติหลักพันครั้ง/วัน) สำหรับทีมขนาดเล็ก-กลางใช้งานปกติเพียงพอ
- ถ้าต้องการความปลอดภัย/เรียลไทม์แบบเข้มงวดกว่านี้ในอนาคต ควรย้ายไปใช้ฐานข้อมูลจริง (Firebase, Supabase ฯลฯ) แทน Google Sheets

## โครงสร้างโปรเจกต์

```
index.html              หน้าเว็บทั้งหมด (HTML/CSS/JS ในไฟล์เดียว)
assets/                 โลโก้บริษัทและรูปนโยบาย/สวัสดิการที่ดึงมาจากเอกสาร Onboarding Passport ต้นฉบับ
apps-script/Code.gs     โค้ด backend ที่ต้องเอาไปวางใน Google Apps Script (ดูขั้นตอนด้านบน)
```

## พัฒนาต่อ / แก้ไขหน้าเว็บ

แก้ไข `index.html` ได้โดยตรง (เป็น vanilla JS ไม่มี build step) แล้ว commit + push ขึ้น `main` — GitHub Pages จะอัปเดตเว็บให้อัตโนมัติภายในไม่กี่นาที
