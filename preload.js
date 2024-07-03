// เพิ่ม event listener ที่จะทำงานเมื่อเนื้อหาของหน้าต่างถูกโหลดเสร็จสมบูรณ์
window.addEventListener('DOMContentLoaded', () => {
    // ฟังก์ชัน replaceText สำหรับการแทนที่ข้อความในองค์ประกอบที่ระบุด้วยตัวเลือก (selector) และข้อความ (text)
    const replaceText = (selector, text) => {
        // ดึงองค์ประกอบจาก DOM ด้วยตัวเลือกที่ระบุ
        const element = document.getElementById(selector);
        // ถ้าองค์ประกอบถูกพบ ให้แทนที่ข้อความภายในด้วยข้อความใหม่
        if (element) element.innerText = text;
    };

    // สำหรับแต่ละประเภทใน ['chrome', 'node', 'electron']
    for (const type of ['chrome', 'node', 'electron']) {
        // เรียกใช้ฟังก์ชัน replaceText โดยใช้ ID ขององค์ประกอบและเวอร์ชันของประเภทนั้น ๆ
        replaceText(`${type}-version`, process.versions[type]);
    }
});
