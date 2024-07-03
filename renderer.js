const images = [
    "file:///E:/Python/Report/images/B1.jpg",
    "file:///E:/Python/Report/images/B2.jpg",
    "file:///E:/Python/Report/images/B3.jpg",
    "file:///E:/Python/Report/images/B4.jpg",
    "file:///E:/Python/Report/images/B5.jpg",
    "file:///E:/Python/Report/images/B6.jpg",
    "file:///E:/Python/Report/images/B7.jpg",
    "file:///E:/Python/Report/images/B8.jpg",
    "file:///E:/Python/Report/images/B9.jpg"
];
let currentImageIndex = 0;

function changeBackgroundImage() {
    document.body.style.backgroundImage = `url(${images[currentImageIndex]})`;
    currentImageIndex = (currentImageIndex + 1) % images.length;
}

setInterval(changeBackgroundImage, 5000); // เปลี่ยนทุก 5 วินาที
changeBackgroundImage(); // เริ่มการเปลี่ยนภาพทันทีที่โหลดหน้าเว็บ

document.getElementById("saveBtn").addEventListener("click", function() {
    const formData = new FormData(document.getElementById("dataForm"));
    const data = {
        User: formData.get("user"),
        Shift: formData.get("shift"),
        Date: formData.get("date"),
        "Extruder No": formData.get("extruder-no"),
        "Start Time": formData.get("start-time"),
        "End Time": formData.get("end-time"),
        "Name Food": formData.get("name-food"),
        "Code Lot": formData.get("code-lot"),
        Shape: formData.get("shape"),
        Size: formData.get("size"),
        Color: formData.get("color"),
        "Number Dir": formData.get("number-dir"),
        "Number of Blades": formData.get("number-blades"),
        Oil: formData.get("oil"),
        "Lot color oil": formData.get("lot-color-oil"),
        "Weight color oil": formData.get("weight-color-oil")
    };

    fetch('http://127.0.0.1:5000/save_data', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify(data)
    })
    .then(response => response.json())
    .then(data => {
        console.log('Success:', data);
    })
    .catch((error) => {
        console.error('Error:', error);
    });
});

// เพิ่มฟังก์ชันการบันทึกข้อมูลลงใน Excel และเปิดหน้า P3manual.html
document.getElementById("saveToExcelBtn").addEventListener("click", function() {
    const formData = new FormData(document.getElementById("dataForm"));
    const data = {
        User: formData.get("user"),
        Shift: formData.get("shift"),
        Date: formData.get("date"),
        "Extruder No": formData.get("extruder-no"),
        "Start Time": formData.get("start-time"),
        "End Time": formData.get("end-time"),
        "Name Food": formData.get("name-food"),
        "Code Lot": formData.get("code-lot"),
        Shape: formData.get("shape"),
        Size: formData.get("size"),
        Color: formData.get("color"),
        "Number Dir": formData.get("number-dir"),
        "Number of Blades": formData.get("number-blades"),
        Oil: formData.get("oil"),
        "Lot color oil": formData.get("lot-color-oil"),
        "Weight color oil": formData.get("weight-color-oil")
    };

    fetch('http://127.0.0.1:5000/save_to_excel', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify(data)
    })
    .then(response => response.json())
    .then(data => {
        console.log('Success:', data);
        // เปิดหน้า P3manual.html
        window.open('P3manual.html', '_blank');
    })
    .catch((error) => {
        console.error('Error:', error);
    });
});
