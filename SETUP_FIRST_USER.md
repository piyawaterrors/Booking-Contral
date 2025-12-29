# วิธีสร้าง User แรกในระบบ

## วิธีที่ 1: เพิ่มข้อมูลใน Google Sheets โดยตรง (แนะนำ)

### ขั้นตอน:

1. เปิด Google Spreadsheet ของคุณ
2. ไปที่ Sheet "Users"
3. ใส่ Header ในแถวแรก (Row 1):

```
A1: รหัสผู้ใช้
B1: อีเมล
C1: ชื่อผู้ใช้
D1: รหัสผ่าน
E1: ชื่อ-นามสกุล
F1: บทบาท
G1: สถานะการใช้งาน
H1: วันที่สร้าง
I1: วันที่แก้ไขล่าสุด
```

4. ใส่ข้อมูล Owner ในแถวที่ 2 (Row 2):

```
A2: USR001
B2: owner@example.com
C2: admin
D2: 5e884898da28047151d0e56f8dc6292773603d0d6aabbdd62a11ef721d1542d8
E2: ผู้ดูแลระบบ
F2: Owner
G2: Active
H2: 2025-12-29 19:00:00
I2: 2025-12-29 19:00:00
```

### ข้อมูล Login:

- **Username:** admin
- **Password:** password123

**หมายเหตุ:** รหัสผ่านในคอลัมน์ D เป็น SHA-256 hash ของ "password123"

---

## วิธีที่ 2: สร้างฟังก์ชัน Setup ใน Code.gs

เพิ่มฟังก์ชันนี้ใน Code.gs และรันครั้งเดียว:

```javascript
/**
 * Setup Initial Owner User (Run this once)
 */
function setupInitialOwner() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.USERS);

    // Check if sheet is empty (only header or no data)
    const data = sheet.getDataRange().getValues();
    if (data.length > 1) {
      Logger.log("มี User อยู่ในระบบแล้ว");
      return {
        success: false,
        message: "มี User อยู่ในระบบแล้ว ไม่สามารถสร้าง Owner ใหม่ได้",
      };
    }

    // Create header if not exists
    if (data.length === 0) {
      const headers = [
        "รหัสผู้ใช้",
        "อีเมล",
        "ชื่อผู้ใช้",
        "รหัสผ่าน",
        "ชื่อ-นามสกุล",
        "บทบาท",
        "สถานะการใช้งาน",
        "วันที่สร้าง",
        "วันที่แก้ไขล่าสุด",
      ];
      sheet.appendRow(headers);
    }

    // Create Owner user
    const userId = "USR001";
    const hashedPassword = hashPassword("password123");
    const timestamp = getCurrentTimestamp();

    const ownerRow = [
      userId,
      "owner@example.com",
      "owner",
      hashedPassword,
      "เจ้าของ",
      "Owner",
      "เปิดใช้งาน",
      timestamp,
      timestamp,
    ];

    sheet.appendRow(ownerRow);

    Logger.log("สร้าง Owner สำเร็จ!");
    Logger.log("Username: owner");
    Logger.log("Password: password123");

    return {
      success: true,
      message: "สร้าง Owner สำเร็จ!\nUsername: owner\nPassword: password123",
      data: {
        username: "admin",
        password: "password123",
      },
    };
  } catch (error) {
    Logger.log("Error: " + error.message);
    return {
      success: false,
      message: "เกิดข้อผิดพลาด: " + error.message,
    };
  }
}
```

### วิธีรันฟังก์ชัน:

1. เปิด Apps Script Editor
2. เลือกฟังก์ชัน `setupInitialOwner` จาก dropdown
3. คลิก Run (▶️)
4. อนุญาตสิทธิ์ที่ขอ
5. ดูผลลัพธ์ใน Execution log

---

## วิธีที่ 3: สร้างหน้า Setup สำหรับครั้งแรก

สร้างหน้า setup.html ที่ตรวจสอบว่ามี User หรือยัง ถ้าไม่มีให้สร้าง Owner อัตโนมัติ

---

## แนะนำ: ใช้วิธีที่ 1

วิธีที่ 1 ง่ายที่สุดและรวดเร็วที่สุด เพียงแค่ copy-paste ข้อมูลลงใน Google Sheets

หลังจากสร้าง User แรกแล้ว คุณสามารถ:

1. Login ด้วย username: admin, password: password123
2. ไปที่เมนู "จัดการพนักงาน"
3. สร้าง User อื่นๆ ผ่านระบบได้เลย
4. เปลี่ยนรหัสผ่านของ admin ในภายหลัง
