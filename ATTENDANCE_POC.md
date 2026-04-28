# Attendance Comparison POC

เอกสารนี้สรุป proof of concept สำหรับเทียบตารางเวรกับข้อมูลแตะนิ้วจริง โดยรอบแรกจำกัดเฉพาะพนักงาน pilot:

- พนักงาน: `ณัฐธนนท์ มานะกุล`
- ชื่อใน CSV: `ณัฐธนนท์`
- รหัสพนักงานใน CSV: `26001`
- แผนกใน CSV: `OMO`
- เดือนที่เทียบ: เมษายน 2026 / เมษายน 2569
- worksheet: `แก้ไข เมษายน 69`

## What Was Added

- เพิ่ม Apps Script helper `inspectPilotAttendanceEmployee(sheetName)` สำหรับตรวจ row พนักงานแบบ read-only
- เพิ่ม Apps Script helper `getPilotAttendancePocData(sheetName)` สำหรับส่ง master expected schedule ของพนักงาน pilot ไปให้หน้า web app
- เพิ่ม `buildPilotExpectedSchedule_()` เพื่อสร้าง expected schedule จาก normalized `scheduleEntries`
- เพิ่ม `getShiftTimingWithBreak_()` เพื่อแปลเฉพาะกฎ pilot:
  - `O-1` = ทำงาน 08:00-17:00
  - พักกลางวัน 12:00-13:00
  - expected work minutes = 480
  - `Off*` = วันหยุด
  - blank / unsupported code ถูกแยกสถานะชัดเจน
- เพิ่ม break metadata ให้ Shift Dictionary/fallback โดยไม่เปลี่ยน flow edit/save เดิม
- เพิ่มส่วน UI `ตรวจเทียบเวลา POC` ใน `Index.html`
  - upload CSV ด้วย `FileReader`
  - parse CSV ฝั่ง client
  - normalize punch records ของ `ณัฐธนนท์`
  - run comparison กับ master schedule จาก Google Sheet
  - แสดงตาราง expected vs actual

## Data Flow

1. Web app render ตารางเวรตาม flow เดิมผ่าน `getWorksheetWebAppData()`
2. ผู้ใช้เลือก CSV บันทึกเวลาใน section POC
3. Browser parse CSV เป็น records ด้วย `parseAttendanceCsvClientSide()`
4. `normalizePilotAttendanceRecords()` filter เฉพาะ:
   - `หมายเลขพนักงาน = 26001` หรือ
   - `ชื่อ = ณัฐธนนท์`
5. ระบบ group records ตามวันที่ และ deduplicate timestamp ซ้ำในวันเดียวกัน
6. เมื่อกดรันเทียบเวลา หน้าเว็บเรียก `getPilotAttendancePocData('แก้ไข เมษายน 69')`
7. Server อ่าน normalized schedule entries ของ `ณัฐธนนท์ มานะกุล`
8. Client ใช้ `comparePilotAttendance()` สร้างผลเทียบ expected vs actual

## Output Fields

Comparison table แสดง field หลัก:

- `date`
- `employeeName`
- `rawCode`
- `expectedStatus`
- `expectedStart`
- `actualStart`
- `breakStart`
- `breakOut`
- `breakEnd`
- `breakIn`
- `expectedEnd`
- `actualEnd`
- `actualPunchCount`
- `flags`

## Assumptions

- รอบนี้เทียบเฉพาะวันที่ไม่เกิน `2026-04-28`
- `2026-04-28 11:13` เป็นเวลาที่ export CSV ดังนั้นถ้ามีแค่ punch เช้าในวันที่ 28 ให้ flag `PARTIAL_EXPORT_DAY`
- CSV ปัจจุบันทุกแถวมี `สถานะการลงเวลา = เช็คอิน` จึง infer punch role จากลำดับเวลา
- ถ้าวันหนึ่งมี 4 unique times:
  - time[0] = actualStart
  - time[1] = breakOut
  - time[2] = breakIn
  - time[3] หรือเวลาสุดท้าย = actualEnd
- ถ้าวันหนึ่งมี 3 unique times ให้ถือว่า lunch punch ไม่ครบและ flag `MISSING_BREAK_PUNCH`
- ถ้าวันหนึ่งมี duplicate timestamp ให้ deduplicate และ flag `DUPLICATE_PUNCH_DEDUPED`

## Flags

- `OFF_DAY`
- `UNSUPPORTED_SHIFT`
- `BLANK_SCHEDULE`
- `ABSENT`
- `MISSING_BREAK_PUNCH`
- `MISSING_CLOCK_OUT`
- `DUPLICATE_PUNCH_DEDUPED`
- `PARTIAL_EXPORT_DAY`

## Unsupported Cases

- ยังไม่รองรับ payroll calculation
- ยังไม่คำนวณ late / early leave / overtime
- ยังไม่รองรับ shift อื่นนอกจาก `O-1` และ `Off*` ใน POC
- ยังไม่รองรับการ upload CSV ไปเก็บใน Google Sheet
- ยังไม่รองรับการ match employee แบบถาวรด้วย employee master table
- ยังไม่รองรับหลายคน หลายเดือน หรือหลายแผนก

## Next Steps

1. เพิ่ม employee mapping sheet ระหว่างชื่อในตารางเวร, ชื่อใน CSV, employee number และ department
2. ย้าย shift timing/break rules ไปไว้ใน `Shift Dictionary` ให้ครบทุก code
3. เพิ่ม tolerance rules เช่น สายได้กี่นาที, ลืมแตะพัก, ทำ OT
4. เพิ่ม attendance import repository หรือ hidden sheet สำหรับเก็บ CSV upload
5. แยก comparison result เป็น audit/report ที่ export ได้
6. ค่อยต่อ payroll calculation หลัง expected vs actual rules เสถียรแล้ว
