# Shift Symbol Concept

เอกสารนี้สรุปแนวคิดการออกแบบสัญลักษณ์ตารางเวรสำหรับระบบที่อ่านข้อมูลจาก Google Sheet มาแสดงบน Web App และในอนาคตจะรองรับการแก้ไข/บันทึกกลับไปยัง Google Sheet แบบ CRUD

## Problem

ตารางเวรเดิมใช้ cell ขนาดเล็กใน Google Sheet เพื่อแสดงรหัสเวรแบบสั้น เช่น `1`, `2`, `11H-3`, `4H-1`, `3H-4`, `Off1` ทำให้ดูภาพรวมเร็ว แต่รหัสบางตัวมีความหมายซ้อนกันหลายชั้น โดยเฉพาะกลุ่มเภสัชกรอาวุโส

ตัวอย่างปัญหา:

| Raw code | สิ่งที่แสดงใน cell | ความหมายจริงที่ควรสื่อ |
| --- | --- | --- |
| `4H-1` | ทำงาน 4 ชั่วโมงที่สาขา 001 | ทำงานสำนักงาน `O-1` 08.00-16.30 แล้วไปสาขา 001 17.00-21.00 |
| `3H-4` | ทำงาน 3 ชั่วโมงที่สาขา 004 | ทำงานสำนักงาน `O-1` 08.00-16.30 แล้วไปสาขา 004 17.00-20.00 |
| `11H-3` | ทำงาน 11 ชั่วโมงที่สาขา 003 | ทำงานที่สาขา 003 รวม 11 ชั่วโมง |

ดังนั้น raw code ใน spreadsheet ไม่ควรเป็นความหมายทั้งหมดของ shift แต่ควรเป็นเพียง "รหัสย่อ" ที่ web app สามารถแปลเป็นข้อมูลโครงสร้างได้

## Core Concept

แนวคิดหลักคือแยก 3 สิ่งออกจากกัน:

1. **Raw Code**
   รหัสที่อยู่ใน Google Sheet เช่น `1`, `2`, `4H-1`, `3H-4`

2. **Structured Meaning**
   ความหมายจริงของรหัส เช่น ช่วงเวลา, สถานที่, ประเภทงาน, งานหลัก, งานเสริม

3. **Display Layer**
   วิธีแสดงผลบน web app เช่น badge, tooltip, detail panel, edit drawer

Google Sheet ยังเก็บข้อมูลแบบ compact ได้เหมือนเดิม แต่ web app ต้องเป็นตัวแปลความหมายและแสดงรายละเอียดที่ cell เล็ก ๆ แสดงไม่พอ

## Shift As Segments

หนึ่ง cell ในตารางเวรไม่จำเป็นต้องหมายถึง shift เดี่ยวเสมอไป แต่ควรถูกมองเป็นชุดของช่วงเวลาทำงาน หรือ `segments`

ตัวอย่าง `4H-1` สำหรับเภสัชกรอาวุโส:

```js
{
  rawCode: "4H-1",
  roleGroup: "senior_pharmacist",
  segments: [
    {
      type: "office",
      locationCode: "office",
      label: "O-1",
      start: "08:00",
      end: "16:30"
    },
    {
      type: "branch_ot",
      locationCode: "001",
      label: "4H-1",
      start: "17:00",
      end: "21:00"
    }
  ]
}
```

ตัวอย่าง `3H-4`:

```js
{
  rawCode: "3H-4",
  roleGroup: "senior_pharmacist",
  segments: [
    {
      type: "office",
      locationCode: "office",
      label: "O-1",
      start: "08:00",
      end: "16:30"
    },
    {
      type: "branch_ot",
      locationCode: "004",
      label: "3H-4",
      start: "17:00",
      end: "20:00"
    }
  ]
}
```

## Role Groups

การแปลรหัสต้องอิงตามสายงาน เพราะรหัสเดียวกันอาจมีบริบทต่างกัน

### เภสัชกรอาวุโส

| Raw code | ความหมาย |
| --- | --- |
| `11H-3` | ทำงานที่สาขา 003 เป็นเวลา 11 ชั่วโมง |
| `4H-1` | สำนักงาน `O-1` 08.00-16.30 + สาขา 001 17.00-21.00 |
| `3H-4` | สำนักงาน `O-1` 08.00-16.30 + สาขา 004 17.00-20.00 |

รหัสกลุ่มนี้บางตัวเป็น **composite shift** คือมีงานหลักและงานต่อท้ายในวันเดียวกัน

### เภสัชปฏิบัติการ / ผู้ช่วยเภสัชกร

| Raw code | ความหมาย |
| --- | --- |
| `1` | ทำงาน 08.00-17.00 |
| `2` | ทำงาน 11.00-20.00 |
| `1/3H` | ทำงาน OT 17.00-20.00 |
| `Off1`, `Off2` | วันหยุด |

รหัสกลุ่มนี้เป็นมาตรฐานกว่า และเหมาะกับ pattern แบบ base shift / add-on shift

## Display Design

ไม่ควรพยายามยัดรายละเอียดทั้งหมดลงใน cell เดียว เพราะ cell ใน spreadsheet และ web table มีพื้นที่จำกัด

ควรออกแบบการแสดงผลเป็น 3 ระดับ:

## 1. Compact Cell

ใช้สำหรับดูภาพรวมเร็วในตาราง

ตัวอย่าง:

```text
O-1
4H-1
```

หรือ:

```text
O-1 + 4H-1
```

สำหรับ shift ปกติ:

```text
1
```

```text
2
```

## 2. Tooltip / Hover Detail

ใช้สำหรับอ่านความหมายโดยไม่ต้องออกจากหน้าตาราง

ตัวอย่าง tooltip ของ `4H-1`:

```text
สำนักงาน 08.00-16.30
สาขา 001 17.00-21.00
```

ตัวอย่าง tooltip ของ `3H-4`:

```text
สำนักงาน 08.00-16.30
สาขา 004 17.00-20.00
```

## 3. Click Detail / Edit Drawer

ใช้สำหรับ CRUD ในอนาคต

ตัวอย่าง:

```text
วันที่: 14 มีนาคม 2569
พนักงาน: ...
Raw code: 4H-1

ช่วงงาน:
1. สำนักงาน 08.00-16.30
2. สาขา 001 17.00-21.00

[แก้ไข] [บันทึกลง Google Sheet]
```

## Color Concept

สีมีความสำคัญเพราะ spreadsheet ใช้สีช่วยสื่อความหมายของงาน จึงควร preserve สีจาก Google Sheet ให้มากที่สุด

แนวคิดสีที่ควรใช้ใน web app:

| ประเภท | แนวทางสี |
| --- | --- |
| เวรปกติ | ใช้สีจาก Google Sheet |
| งานสำนักงาน | badge หรือพื้นสีอ่อนที่อ่านง่าย เช่น `O-1` |
| งานสาขา | ใช้สีตามสาขาหรือสีที่ spreadsheet กำหนด |
| OT / งานต่อท้าย | แสดงเป็น badge เพิ่ม หรือขอบ/ชั้นที่ต่างจาก base shift |
| Off | ใช้สีตาม spreadsheet เช่น ฟ้า |
| ลา / BL / RxM | ใช้สีเฉพาะตาม spreadsheet |

หลักสำคัญ:

- web app ควรใช้ cell background จาก Google Sheet เป็นหลัก
- font color จาก Google Sheet ใช้ได้ถ้าอ่านออก
- ถ้า font color อ่านยาก ให้ web app fallback เป็นสีดำ/ขาวที่ contrast เพียงพอ
- ไม่ควรปรับสี cell ให้จางลงเอง เพราะสีเป็นข้อมูลเชิงความหมาย

## Shift Dictionary

ควรเพิ่ม worksheet สำหรับเก็บพจนานุกรมรหัสเวร เช่น `Shift Dictionary` หรือ `รหัสเวร`

ตัวอย่าง schema:

| column | ความหมาย |
| --- | --- |
| `rawCode` | รหัสที่อยู่ใน Google Sheet |
| `roleGroup` | กลุ่มสายงาน เช่น `senior_pharmacist`, `operation_pharmacist`, `assistant` |
| `displayLabel` | label ที่ใช้แสดงใน web app |
| `displayMode` | `single`, `base`, `addon`, `composite`, `off`, `leave` |
| `segmentType` | `regular`, `office`, `branch`, `branch_ot`, `off`, `leave` |
| `locationCode` | สาขาหรือสถานที่ เช่น `001`, `003`, `004`, `office` |
| `start` | เวลาเริ่ม |
| `end` | เวลาจบ |
| `impliedBaseCode` | shift พื้นฐานที่ต้องถือว่ามีด้วย เช่น `O-1` |
| `color` | สี fallback ถ้าไม่ใช้สีจาก cell |
| `description` | คำอธิบายมนุษย์อ่าน |

ตัวอย่างข้อมูล:

| rawCode | roleGroup | displayLabel | displayMode | segmentType | locationCode | start | end | impliedBaseCode | description |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| `1` | `operation_pharmacist` | `1` | `single` | `regular` |  | `08:00` | `17:00` |  | ทำงาน 08.00-17.00 |
| `2` | `operation_pharmacist` | `2` | `single` | `regular` |  | `11:00` | `20:00` |  | ทำงาน 11.00-20.00 |
| `1/3H` | `operation_pharmacist` | `1/3H` | `addon` | `ot` |  | `17:00` | `20:00` | `1` | OT หลังเวร 1 |
| `O-1` | `senior_pharmacist` | `O-1` | `base` | `office` | `office` | `08:00` | `16:30` |  | งานสำนักงาน |
| `4H-1` | `senior_pharmacist` | `4H-1` | `composite` | `branch_ot` | `001` | `17:00` | `21:00` | `O-1` | สำนักงาน + สาขา 001 |
| `3H-4` | `senior_pharmacist` | `3H-4` | `composite` | `branch_ot` | `004` | `17:00` | `20:00` | `O-1` | สำนักงาน + สาขา 004 |
| `11H-3` | `senior_pharmacist` | `11H-3` | `single` | `branch` | `003` |  |  |  | ทำงานที่สาขา 003 รวม 11 ชั่วโมง |

## Current Implementation Status

ตอนนี้ repo เริ่มมี foundation สำหรับ edit mode และ pseudo-relational layer แล้ว โดยยังคงให้ worksheet รายเดือนเป็น source หลักเหมือนเดิม

สิ่งที่ทำแล้ว:

- อ่านข้อมูลจาก Google Sheet พร้อม values, สีพื้น, สีตัวอักษร, merge metadata และ shift metadata
- มี `Shift Dictionary` / `รหัสเวร` เป็น optional worksheet และมี fallback dictionary ใน Apps Script
- มี normalized object layer สำหรับ employees, schedule entries, shift dictionary items, draft changes และ audit log
- มี Edit Mode ใน web app: เลือก cell, แก้ raw code แบบ local, ล้าง cell, review pending changes และ save กลับ Google Sheet
- save flow เขียน raw code กลับ cell เดิม และ append audit row ไปที่ `Schedule Audit Log`
- sidebar catalog ใช้ metadata จาก Shift Dictionary เพื่อแสดง label, รายละเอียด และสีของสัญลักษณ์
- regular monthly Off ไม่ให้ผู้ใช้เลือก `Off1`, `Off2` เอง แต่ให้กด `ตั้งเป็น Off` แล้วระบบเรียงเลข `Off1`-`Off6` ต่อพนักงานต่อเดือน
- Annual Leave ไม่ให้ผู้ใช้กำหนด `AL1`-`AL6` เอง แต่ให้เลือก `ตั้งลาพักร้อน` จาก catalog แล้วระบบเรียงเลขทั้งปี
- Public Holiday / นักขัตฤกษ์ ใช้ catalog option `ตั้งนักขัตฤกษ์` แล้วระบบเรียงเลข `Off1`-`Off13` ต่อพนักงานต่อปี พร้อม metadata/color แยกจาก regular Off

ข้อสำคัญของ Public Holiday:

- raw code ที่บันทึกลงชีทยังเป็น `Off1`-`Off13` ตามรูปแบบชีทเดิม
- web app แยก Public Holiday ออกจาก regular Off ด้วย metadata และสีของ cell ไม่ใช่ดูจาก raw code อย่างเดียว
- การลบ Public Holiday ใช้ปุ่ม `ล้างช่องนี้` เดิม แล้วระบบ recompute ลำดับที่เหลือทั้งปี
- ยังไม่มีการ add/remove employee row หรือ add/remove date column จาก web app

## CRUD Direction

อนาคต web app ควรแก้ข้อมูลแบบ structured ไม่ใช่ให้ผู้ใช้พิมพ์ raw code อย่างเดียว

แนวทาง:

1. อ่าน raw code จาก Google Sheet
2. อ่าน `Shift Dictionary`
3. แปลง raw code + role group เป็น segments
4. แสดงผลแบบ compact ในตาราง
5. เมื่อคลิก cell ให้เปิด detail/edit drawer
6. ผู้ใช้แก้ segments เช่น เวลา, สาขา, ประเภทงาน
7. ระบบ serialize กลับเป็น raw code หรือบันทึก structured data ลง hidden sheet เพิ่ม
8. Google Sheet ยังอ่านได้แบบเดิม แต่ web app มีความหมายครบกว่า

## Recommended Roadmap

### Phase 1: Display Parser

- เก็บ Google Sheet แบบเดิม
- เพิ่ม parser ใน web app สำหรับ code สำคัญ เช่น `4H-1`, `3H-4`, `11H-3`, `1`, `2`, `1/3H`
- แสดง tooltip หรือ title ที่อธิบายความหมายจริง
- preserve สีจาก Google Sheet

### Phase 2: Shift Dictionary

- เพิ่ม worksheet `Shift Dictionary`
- ย้าย mapping จาก code ไปอยู่ใน Google Sheet
- web app อ่าน dictionary เพื่อแปลรหัสแบบ dynamic
- ลดการ hardcode ใน `Code.js` / `Index.html`

### Phase 3: Structured Edit

- คลิก cell แล้วเปิด edit drawer
- แก้ base shift / add-on shift / location / time ได้
- validate เวลาและกฎของแต่ละสายงาน
- บันทึกกลับ Google Sheet

### Phase 4: Full CRUD

- เพิ่ม/แก้/ลบ shift ได้จาก web app
- บันทึก audit trail หรือ draft changes
- รองรับ preview ก่อนเขียนทับ Google Sheet จริง
- รองรับหลายสายงานและเวลายืดหยุ่น

## Design Principle

อย่าบังคับให้ cell เล็ก ๆ ใน spreadsheet สื่อทุกอย่างให้ครบ

ให้ Google Sheet เป็น compact source และให้ web app เป็นตัวแปลความหมาย แสดงรายละเอียด และแก้ไขข้อมูลแบบ structured

แนวนี้จะทำให้ระบบ:

- ใช้งานร่วมกันได้หลายสายงาน
- รองรับรหัสเวรที่ซับซ้อน
- ไม่เสีย compatibility กับ Google Sheet เดิม
- รองรับ CRUD ในอนาคตได้
- ลดความสับสนของสัญลักษณ์เฉพาะกลุ่ม เช่น เภสัชกรอาวุโส
