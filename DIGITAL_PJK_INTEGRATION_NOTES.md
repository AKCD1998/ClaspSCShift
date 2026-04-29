# DigitalPJK Integration Notes

บันทึกสำหรับต่อ API จากโปรเจค `ClaspSCShift` ไปยัง repo:

`C:\Users\scgro\Desktop\Webapp training project\digitalPJKform`

เป้าหมายคือหลังจากตารางเวรใน Google Apps Script ถูกบันทึกเสร็จแล้ว ให้มีปุ่มสั่งสร้างเอกสาร preview จาก `digitalPJKform` ก่อนพิมพ์ โดยใช้รูปแบบ PDF template และพิกัดข้อความจาก repo นั้น

## ภาพรวม digitalPJKform

- Stack: React + Vite frontend, Express backend, PostgreSQL
- PDF generation ใช้ `pdf-lib`
- endpoint หลักสำหรับสร้าง PDF:
  - `POST /api/documents/generate`
  - route อยู่ที่ `backend/src/routes/documents.routes.js`
  - controller อยู่ที่ `backend/src/controllers/documents.controller.js`
- endpoint ต้องใช้ `Authorization: Bearer <JWT>` ผ่าน `authRequired`
- template ปัจจุบัน:
  - PDF: `backend/src/assets/templates/form_gor_gor_1.pdf`
  - field coordinates: `backend/src/assets/templates/form_gor_gor_1.fields.json`
- ข้อมูลสาขาอยู่ใน:
  - `backend/src/data/branches/001.json`
  - `backend/src/data/branches/003.json`
  - `backend/src/data/branches/004.json`
  - `backend/src/data/branches/005.json`

## ข้อมูลสาขาที่พบใน digitalPJKform

| Branch | สาขา | มีใน repo |
| --- | --- | --- |
| `001` | ตลาดแม่กลอง | มี |
| `003` | วัดช่องลม | มี |
| `004` | ตลาดบางน้อย | มี |
| `005` | ถนนเอกชัยสมุทรสาคร | มี |

ตัวอย่าง field ของ branch JSON:

```json
{
  "pharmacy_name_th": "ศิริชัยเภสัช",
  "branch_name_th": "วัดช่องลม",
  "address_no": "60/11",
  "soi": "-",
  "district": "เมืองสมุทรสงคราม",
  "province": "สมุทรสงคราม",
  "postcode": "75000",
  "phone": "098-861-5900",
  "license_no": "สส.5/2565",
  "location_text": "บ้านปรก",
  "operator_title": "ภญ.",
  "operator_work_hours": "08:00-22:00"
}
```

## Payload ที่ digitalPJKform รับ

Frontend เดิมเรียก `generateDocumentPdfWithOptions()` แล้วส่ง body ประมาณนี้:

```json
{
  "templateKey": "form_gor_gor_1",
  "branchId": "<branch uuid>",
  "formData": {
    "operatorTitle": "ภญ. ศุภิสรา ศิริมงคล"
  },
  "subPharmacistSlots": [
    {
      "pharmacistName": "ภก. ธนัท สังขรักษ์",
      "pharmacistLicense": "35084",
      "dateStart": "2026-01-31",
      "dateEnd": "2026-02-03",
      "timeStart": "09:00",
      "timeEnd": "18:00"
    }
  ]
}
```

`subPharmacistSlots` คือส่วนนี้ในเอกสาร:

> ขอแจ้งชื่อเภสัชกรผู้มีหน้าที่ปฏิบัติการแทนผู้มีหน้าที่ปฏิบัติการซึ่งไม่อาจปฏิบัติหน้าที่เป็นการชั่วคราว

field coordinates ใน `form_gor_gor_1.fields.json` ตอนนี้รองรับ slot 1-3:

- `slot1_pharmacistName`
- `slot1_pharmacistLicense`
- `slot1_dateStart`
- `slot1_dateEnd`
- `slot1_timeRange`
- `slot2_*`
- `slot3_*`

ดังนั้นชุดแรกควรรองรับสูงสุด 3 รายการต่อเอกสาร ถ้ามากกว่านี้ต้องตัดเป็นหลาย PDF หรือขยาย template/mapping

## Pattern จาก ClaspSCShift

ฝั่ง Google Apps Script มี normalized schedule model อยู่แล้ว:

- `buildEmployeeModels_()` อ่านข้อมูลพนักงานจากแถวตาราง
- `buildScheduleEntryModels_()` สร้างรายการเวรต่อวัน
- `getDigitalPjkRequiredScheduleEntries()` คืนรายการ cell ที่ต้องออก ภจก.1 ตามกฎ `requiresDigitalPjk`
- schedule entry มีข้อมูลสำคัญ:
  - `employeeName`
  - `roleGroup`
  - `workDate`
  - `rawCode`
  - `locationCode`
  - `start`
  - `end`
  - `segmentType`
  - `displayMode`
- `locationCode` มาจาก Shift Dictionary หรือ fallback เช่น `001`, `003`, `004`
- `roleGroup` ที่เกี่ยวข้อง:
  - `senior_pharmacist`
  - `operation_pharmacist`

## Mapping ที่ต้องการ

### เภสัชกรประจำสาขา

ในเอกสารอยู่แถวประมาณ:

`รหัสไปรษณีย์ ... โทรศัพท์ ... มีนาย/นาง/นางสาว ... เป็นผู้มีหน้าที่ปฏิบัติการ เวลาปฏิบัติการ ...`

ควร map จากข้อมูล branch/user ของ `digitalPJKform`:

- branch profile ใช้ข้อมูลร้าน, เลขที่, ตำบล/อำเภอ, จังหวัด, โทรศัพท์, ใบอนุญาตร้าน
- operator หรือเภสัชกรประจำสาขา ใช้ `operatorTitle` / `operator_display_name_th`
- เวลาร้านใช้ `operatorWorkHours`

ต้องยืนยันว่าข้อมูลชื่อเภสัชกรประจำสาขาใน digitalPJKform ตรงกับตาราง GAS และมีคำนำหน้าเรียบร้อย

### เภสัชกรชั่วคราว / Part-time

ในเอกสารอยู่ส่วน:

`ขอแจ้งชื่อเภสัชกรผู้มีหน้าที่ปฏิบัติการแทน...`

ควรสร้างจาก schedule entries ใน GAS ที่เป็นเหตุการณ์เภสัชกรชั่วคราวไปอยู่ร้านแทนเภสัชกรประจำสาขา โดยดูจาก raw code เป็นหลัก:

| จำนวนชั่วโมง | สาขาที่ทำ | Raw code | ต้องออก ภจก.1 |
| --- | --- | --- | --- |
| 12H | 001 | `12H-1` | ใช่ |
| 11H | 003 | `11H-3` | ใช่ |
| 11H | 004 | `11H-4` | ใช่ |
| 11H | 005 | `11H-5` | ใช่ |
| 8H | 001 | `8H-1` | ใช่ |
| 8H | 003 | `8H-3` | ใช่ |
| 8H | 004 | `8H-4` | ใช่ |
| 8H | 005 | `8H-5` | ใช่ |
| 4H | 001 | `4H-1` | ใช่ |
| 3H | 003 | `3H-3` | ใช่ |
| 3H | 004 | `3H-4` | ใช่ |
| 3H | 005 | `3H-5` | ใช่ |
| - | สาขาตัวเอง | `1` | ไม่ใช่ |
| - | สาขาตัวเอง | `1/3H` | ไม่ใช่ |

ระบบจึงควร filter รายการที่:

- เป็นเภสัชกร (`roleGroup` เป็น `senior_pharmacist` หรือ `operation_pharmacist`)
- มี `locationCode` ตรงกับ branch ที่จะออกเอกสาร
- มี `start` และ `end`
- ไม่ใช่ leave/off/empty
- มี `requiresDigitalPjk = true` จาก Shift Dictionary หรือ fallback dictionary

จากนั้น group เป็น slot:

- group by `branchCode + pharmacistName + pharmacistLicense + timeStart + timeEnd`
- รวมวันที่ที่ต่อเนื่องเป็นช่วง `dateStart` ถึง `dateEnd`
- ถ้าไม่ต่อเนื่องให้แยก slot ใหม่

## Branch Code Mapping

| GAS locationCode | digitalPJK branch |
| --- | --- |
| `001` | สาขาตลาดแม่กลอง |
| `003` | สาขาวัดช่องลม |
| `004` | สาขาตลาดบางน้อย |
| `005` | สาขาถนนเอกชัยสมุทรสาคร |

## API Design ที่แนะนำ

ไม่ควรให้ GAS เรียก `/api/documents/generate` โดยตรงด้วย login user ปกติในระยะยาว เพราะต้องจัดการ JWT หมดอายุ

สถานะปัจจุบันใน `ClaspSCShift`:

- Apps Script login ด้วย admin ของ DigitalPJK เพื่อเรียก `/api/branches`, `/api/pharmacists/part-time`, และ `/api/documents/generate`
- `prepareDigitalPjkBatchPreview(sheetName)` เตรียม jobs แยกตาม branch `001`, `003`, `004`, `005`
- หน้าเว็บเรียก `generateDigitalPjkBatchPreviewJob(job)` ทีละ job เพื่ออัปเดต progress bar ได้จริง
- ถ้า branch มีมากกว่า 3 slots ระบบจะแยกเป็นหลาย PDF เพราะ template `form_gor_gor_1` รองรับ slot 1-3

แนะนำเพิ่ม endpoint integration เฉพาะใน `digitalPJKform`:

```text
POST /api/integrations/sc-shift/document-preview
```

ใช้ header:

```text
X-Integration-Key: <secret>
```

payload:

```json
{
  "branchCode": "003",
  "templateKey": "form_gor_gor_1",
  "source": {
    "spreadsheetId": "...",
    "sheetName": "แก้ไข เมษายน 69",
    "generatedAt": "2026-04-28T10:00:00.000Z"
  },
  "operator": {
    "displayNameTh": "ภญ. ศุภิสรา ศิริมงคล"
  },
  "subPharmacistSlots": [
    {
      "pharmacistName": "ภก. ธนัท สังขรักษ์",
      "pharmacistLicense": "35084",
      "dateStart": "2026-01-31",
      "dateEnd": "2026-02-03",
      "timeStart": "09:00",
      "timeEnd": "18:00"
    }
  ]
}
```

ฝั่ง digitalPJKform endpoint นี้ควร:

1. ตรวจ `X-Integration-Key`
2. หา branch ด้วย `branchCode`
3. สร้าง payload แบบเดียวกับ `generateDocumentPdf`
4. เรียก `stampTemplatePdf`
5. return PDF binary เป็น `inline` เพื่อให้ GAS เปิด preview ได้

## งานที่ต้องทำก่อน implementation

1. เพิ่มหรือยืนยันข้อมูลเภสัชกรประจำแต่ละสาขา
2. เพิ่ม endpoint integration สำหรับ GAS ใน digitalPJKform
3. เพิ่ม secret config เช่น `SC_SHIFT_INTEGRATION_KEY`
4. ต่อยอด function ฝั่ง GAS จาก `getDigitalPjkRequiredScheduleEntries()` ให้สร้าง `subPharmacistSlots` จาก normalized schedule
5. เพิ่มปุ่มในหน้า GAS หลังบันทึกตารางสำเร็จ หรือใน toolbar เพื่อสร้าง preview
6. ถ้ามีมากกว่า 3 slots ต้องตัดสินใจว่าจะ:
   - แยกหลาย PDF
   - หรือเพิ่ม fields ใน template สำหรับ slot 4+

## จุดเสี่ยง / ต้องระวัง

- `digitalPJKform` ใช้ `branchId` ภายใน DB แต่ GAS รู้จัก `branchCode` มากกว่า จึงควรให้ integration endpoint รับ `branchCode`
- JWT ปัจจุบันหมดอายุตาม `JWT_EXPIRES_IN` default `1h`; integration key จะเหมาะกับ GAS มากกว่า
- Shift Dictionary ต้องมี `locationCode`, `start`, `end` ครบ ไม่อย่างนั้นสร้างช่วงเวลาปฏิบัติการไม่ได้
- ข้อมูลใบอนุญาตเภสัชกร part-time มีใน `digitalPJKform/backend/src/db/syncPartTimePharmacists.js`; GAS ต้อง map ชื่อให้ตรง หรือ digitalPJK endpoint ต้อง resolve ชื่อ/เลขใบอนุญาตให้
- Template ปัจจุบันรองรับ 3 slots เท่านั้น
