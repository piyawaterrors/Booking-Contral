# 📋 Phase 4: ระบบจัดการสิทธิ์ผู้ใช้ (User Management)

## 🎯 วัตถุประสงค์

พัฒนาระบบจัดการพนักงานและสิทธิ์การใช้งาน สำหรับ **Owner** เท่านั้น

---

## ✨ Features ที่ต้องพัฒนา

### 1. หน้า Employee Management (PageEmployees.html)

- ✅ แสดงรายการพนักงานทั้งหมด
- ✅ ค้นหาพนักงาน (ชื่อ, อีเมล, ชื่อผู้ใช้)
- ✅ กรองตามตำแหน่ง (Sale/OP/Admin/AR_AP/Cost/Owner)
- ✅ กรองตามสถานะ (Active/Inactive)
- ✅ เพิ่มพนักงานใหม่
- ✅ แก้ไขข้อมูลพนักงาน
- ✅ เปลี่ยนสถานะพนักงาน
- ✅ รีเซ็ตรหัสผ่าน
- ✅ ลบพนักงาน (Soft delete)

### 2. Backend Functions (Code.gs)

- ✅ `getAllUsers(sessionToken)` - ดึงข้อมูลพนักงานทั้งหมด
- ✅ `createUser(sessionToken, userData)` - สร้างพนักงานใหม่
- ✅ `updateUser(sessionToken, userId, userData)` - แก้ไขข้อมูลพนักงาน
- ✅ `deleteUser(sessionToken, userId)` - ลบพนักงาน (Soft delete)
- ✅ `resetUserPassword(sessionToken, userId)` - รีเซ็ตรหัสผ่าน

---

## 📊 Database Schema

### ตาราง Users (มีอยู่แล้ว)

| Column | Field Name        | Type     | Description                             |
| ------ | ----------------- | -------- | --------------------------------------- |
| A      | รหัสผู้ใช้        | String   | รหัสผู้ใช้ (Primary Key, Auto-generate) |
| B      | อีเมล             | String   | อีเมลผู้ใช้                             |
| C      | ชื่อผู้ใช้        | String   | ชื่อผู้ใช้สำหรับ Login                  |
| D      | รหัสผ่าน          | String   | รหัสผ่าน (เข้ารหัสด้วย SHA-256)         |
| E      | ชื่อ-นามสกุล      | String   | ชื่อ-นามสกุลเต็ม                        |
| F      | บทบาท             | String   | บทบาท (Sale/OP/Admin/AR_AP/Cost/Owner)  |
| G      | สถานะการใช้งาน    | String   | สถานะ (Active/Inactive)                 |
| H      | วันที่สร้าง       | DateTime | วันที่และเวลาที่สร้างผู้ใช้             |
| I      | วันที่แก้ไขล่าสุด | DateTime | วันที่และเวลาที่แก้ไขล่าสุด             |

---

## 🔧 Backend Functions ที่ต้องสร้าง

### 1. getAllUsers(sessionToken)

```javascript
/**
 * Get All Users (Owner only)
 */
function getAllUsers(sessionToken) {
  try {
    // Validate session and check Owner role
    const session = validateSession(sessionToken);
    if (!session) {
      return {
        success: false,
        message: "Session หมดอายุ กรุณาเข้าสู่ระบบใหม่",
      };
    }

    if (session.role !== CONFIG.ROLES.OWNER) {
      return { success: false, message: "คุณไม่มีสิทธิ์เข้าถึงข้อมูลนี้" };
    }

    const sheet = getSheet(CONFIG.SHEETS.USERS);
    const data = sheet.getDataRange().getValues();

    const users = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      users.push({
        userId: row[0],
        email: row[1],
        username: row[2],
        fullName: row[4],
        role: row[5],
        status: row[6],
        createdAt: row[7],
        updatedAt: row[8],
      });
    }

    return {
      success: true,
      data: users,
    };
  } catch (error) {
    return {
      success: false,
      message: "เกิดข้อผิดพลาด: " + error.message,
    };
  }
}
```

### 2. createUser(sessionToken, userData)

```javascript
/**
 * Create New User (Owner only)
 */
function createUser(sessionToken, userData) {
  try {
    // Validate session
    const session = validateSession(sessionToken);
    if (!session || session.role !== CONFIG.ROLES.OWNER) {
      return { success: false, message: "คุณไม่มีสิทธิ์ดำเนินการนี้" };
    }

    // Validate required fields
    if (
      !userData.fullName ||
      !userData.email ||
      !userData.username ||
      !userData.role
    ) {
      return { success: false, message: "กรุณากรอกข้อมูลให้ครบถ้วน" };
    }

    // Validate email format
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(userData.email)) {
      return { success: false, message: "รูปแบบอีเมลไม่ถูกต้อง" };
    }

    // Check duplicate email and username
    const sheet = getSheet(CONFIG.SHEETS.USERS);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === userData.email) {
        return { success: false, message: "อีเมลนี้ถูกใช้งานแล้ว" };
      }
      if (data[i][2] === userData.username) {
        return { success: false, message: "ชื่อผู้ใช้นี้ถูกใช้งานแล้ว" };
      }
    }

    // Generate user ID
    const userId = generateUserId();

    // Default password
    const defaultPassword = "password123";
    const hashedPassword = hashPassword(defaultPassword);

    // Prepare data
    const now = new Date();
    const newRow = [
      userId,
      userData.email,
      userData.username,
      hashedPassword,
      userData.fullName,
      userData.role,
      "Active",
      now,
      now,
    ];

    // Append to sheet
    sheet.appendRow(newRow);

    return {
      success: true,
      message: "เพิ่มพนักงานสำเร็จ",
      data: { userId: userId },
    };
  } catch (error) {
    return {
      success: false,
      message: "เกิดข้อผิดพลาด: " + error.message,
    };
  }
}
```

### 3. updateUser(sessionToken, userId, userData)

```javascript
/**
 * Update User (Owner only)
 */
function updateUser(sessionToken, userId, userData) {
  try {
    // Validate session
    const session = validateSession(sessionToken);
    if (!session || session.role !== CONFIG.ROLES.OWNER) {
      return { success: false, message: "คุณไม่มีสิทธิ์ดำเนินการนี้" };
    }

    const sheet = getSheet(CONFIG.SHEETS.USERS);
    const data = sheet.getDataRange().getValues();

    // Find user
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === userId) {
        rowIndex = i + 1; // Sheet row is 1-indexed
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, message: "ไม่พบข้อมูลพนักงาน" };
    }

    // Check duplicate email and username (exclude current user)
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] !== userId) {
        if (data[i][1] === userData.email) {
          return { success: false, message: "อีเมลนี้ถูกใช้งานแล้ว" };
        }
        if (data[i][2] === userData.username) {
          return { success: false, message: "ชื่อผู้ใช้นี้ถูกใช้งานแล้ว" };
        }
      }
    }

    // Update data
    sheet.getRange(rowIndex, 2).setValue(userData.email);
    sheet.getRange(rowIndex, 3).setValue(userData.username);
    sheet.getRange(rowIndex, 5).setValue(userData.fullName);
    sheet.getRange(rowIndex, 6).setValue(userData.role);
    sheet.getRange(rowIndex, 7).setValue(userData.status);
    sheet.getRange(rowIndex, 9).setValue(new Date());

    return {
      success: true,
      message: "แก้ไขข้อมูลสำเร็จ",
    };
  } catch (error) {
    return {
      success: false,
      message: "เกิดข้อผิดพลาด: " + error.message,
    };
  }
}
```

### 4. deleteUser(sessionToken, userId)

```javascript
/**
 * Delete User (Soft delete - Owner only)
 */
function deleteUser(sessionToken, userId) {
  try {
    // Validate session
    const session = validateSession(sessionToken);
    if (!session || session.role !== CONFIG.ROLES.OWNER) {
      return { success: false, message: "คุณไม่มีสิทธิ์ดำเนินการนี้" };
    }

    const sheet = getSheet(CONFIG.SHEETS.USERS);
    const data = sheet.getDataRange().getValues();

    // Find user
    let rowIndex = -1;
    let userRole = null;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === userId) {
        rowIndex = i + 1;
        userRole = data[i][5];
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, message: "ไม่พบข้อมูลพนักงาน" };
    }

    // Prevent deleting Owner
    if (userRole === CONFIG.ROLES.OWNER) {
      return { success: false, message: "ไม่สามารถลบ Owner ได้" };
    }

    // Soft delete - change status to Inactive
    sheet.getRange(rowIndex, 7).setValue("Inactive");
    sheet.getRange(rowIndex, 9).setValue(new Date());

    return {
      success: true,
      message: "ลบพนักงานสำเร็จ",
    };
  } catch (error) {
    return {
      success: false,
      message: "เกิดข้อผิดพลาด: " + error.message,
    };
  }
}
```

### 5. resetUserPassword(sessionToken, userId)

```javascript
/**
 * Reset User Password (Owner only)
 */
function resetUserPassword(sessionToken, userId) {
  try {
    // Validate session
    const session = validateSession(sessionToken);
    if (!session || session.role !== CONFIG.ROLES.OWNER) {
      return { success: false, message: "คุณไม่มีสิทธิ์ดำเนินการนี้" };
    }

    const sheet = getSheet(CONFIG.SHEETS.USERS);
    const data = sheet.getDataRange().getValues();

    // Find user
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === userId) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, message: "ไม่พบข้อมูลพนักงาน" };
    }

    // Reset password to default
    const defaultPassword = "password123";
    const hashedPassword = hashPassword(defaultPassword);

    sheet.getRange(rowIndex, 4).setValue(hashedPassword);
    sheet.getRange(rowIndex, 9).setValue(new Date());

    return {
      success: true,
      message: "รีเซ็ตรหัสผ่านสำเร็จ",
    };
  } catch (error) {
    return {
      success: false,
      message: "เกิดข้อผิดพลาด: " + error.message,
    };
  }
}
```

### 6. generateUserId()

```javascript
/**
 * Generate User ID
 */
function generateUserId() {
  const sheet = getSheet(CONFIG.SHEETS.USERS);
  const lastRow = sheet.getLastRow();

  if (lastRow === 1) {
    return "USR001";
  }

  const lastId = sheet.getRange(lastRow, 1).getValue();
  const number = parseInt(lastId.replace("USR", "")) + 1;
  return "USR" + number.toString().padStart(3, "0");
}
```

---

## 🎨 UI Components

### 1. Employee Table

- แสดงข้อมูล: รหัส, ชื่อ-นามสกุล, ชื่อผู้ใช้, อีเมล, ตำแหน่ง, สถานะ
- ปุ่มจัดการ: แก้ไข, รีเซ็ตรหัสผ่าน, ลบ
- Badge สีตามตำแหน่ง
- Badge สีตามสถานะ

### 2. Search & Filter

- Search box: ค้นหาชื่อ, อีเมล, ชื่อผู้ใช้
- Filter ตำแหน่ง: Dropdown
- Filter สถานะ: Dropdown

### 3. Add/Edit Modal

- ฟอร์มกรอกข้อมูล
- Validation
- แสดงข้อความรหัสผ่านเริ่มต้น

### 4. Reset Password Modal

- ยืนยันการรีเซ็ต
- แสดงรหัสผ่านใหม่

---

## ✅ Business Rules

1. **สิทธิ์การเข้าถึง**

   - เฉพาะ Owner เท่านั้น
   - ตรวจสอบสิทธิ์ทุกครั้งที่เรียก API

2. **การสร้างพนักงาน**

   - รหัสผ่านเริ่มต้น: `password123`
   - สถานะเริ่มต้น: `Active`
   - อีเมลและชื่อผู้ใช้ต้องไม่ซ้ำ

3. **การลบพนักงาน**

   - Soft delete (เปลี่ยนสถานะเป็น Inactive)
   - ไม่สามารถลบ Owner ได้

4. **การรีเซ็ตรหัสผ่าน**
   - รหัสผ่านใหม่: `password123`
   - Hash ด้วย SHA-256

---

## 🔒 Security

1. **Authentication**

   - ตรวจสอบ Session Token ทุกครั้ง
   - ตรวจสอบ Role = Owner

2. **Validation**

   - Validate email format
   - Check duplicate email/username
   - Validate required fields

3. **Password**
   - Hash ด้วย SHA-256
   - ไม่แสดงรหัสผ่านจริง

---

## 📝 Testing Checklist

- [ ] Owner สามารถดูรายการพนักงานทั้งหมดได้
- [ ] Owner สามารถเพิ่มพนักงานใหม่ได้
- [ ] Owner สามารถแก้ไขข้อมูลพนักงานได้
- [ ] Owner สามารถเปลี่ยนสถานะพนักงานได้
- [ ] Owner สามารถรีเซ็ตรหัสผ่านได้
- [ ] Owner สามารถลบพนักงานได้ (ยกเว้น Owner)
- [ ] ไม่สามารถลบ Owner ได้
- [ ] อีเมลและชื่อผู้ใช้ต้องไม่ซ้ำ
- [ ] Search และ Filter ทำงานถูกต้อง
- [ ] Role อื่นไม่สามารถเข้าถึงหน้านี้ได้
- [ ] Toast notification แสดงถูกต้อง

---

## 🚀 Deployment Steps

1. เพิ่มฟังก์ชันทั้งหมดใน `Code.gs`
2. อัพเดท `PageEmployees.html` ให้สมบูรณ์
3. ทดสอบการทำงานทุก feature
4. Deploy version ใหม่
5. ทดสอบบน Production

---

**สถานะ:** 🟡 In Progress  
**ผู้รับผิดชอบ:** Antigravity AI  
**วันที่:** 29 ธันวาคม 2568
