// นำเข้าโมดูล app และ BrowserWindow จาก Electron
const { app, BrowserWindow } = require('electron');
// นำเข้าโมดูล path
const path = require('path');

// ฟังก์ชันสร้างหน้าต่างหลักของแอปพลิเคชัน
function createWindow() {
    // สร้างหน้าต่างหลัก (mainWindow) พร้อมกำหนดขนาดและการตั้งค่า
    const mainWindow = new BrowserWindow({
        width: 800, // ความกว้างของหน้าต่าง
        height: 600, // ความสูงของหน้าต่าง
        webPreferences: {
            preload: path.join(__dirname, 'preload.js'), // กำหนดไฟล์ preload.js เพื่อโหลดก่อน
            nodeIntegration: true,  // เปิดใช้งานการรวม Node.js
            contextIsolation: false,  // ปิดใช้งานการแยกบริบท (context isolation)
        }
    });

    // โหลดไฟล์ HTML ที่จะใช้แสดงในหน้าต่างหลัก
    mainWindow.loadFile('P3Auto.html');
}

// เมื่อแอปพร้อมทำงาน ให้สร้างหน้าต่างหลัก
app.whenReady().then(() => {
    createWindow();

    // ถ้าแอปถูกเปิดใหม่และไม่มีหน้าต่างใดๆ ให้สร้างหน้าต่างใหม่
    app.on('activate', () => {
        if (BrowserWindow.getAllWindows().length === 0) {
            createWindow();
        }
    });
});

// เมื่อหน้าต่างทั้งหมดถูกปิด
app.on('window-all-closed', () => {
    // ถ้าแพลตฟอร์มไม่ใช่ macOS (darwin) ให้ปิดแอปพลิเคชัน
    if (process.platform !== 'darwin') {
        app.quit();
    }
});
