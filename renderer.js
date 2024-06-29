const XLSX = require('xlsx');
const { writeFile } = require('fs');
const path = require('path');

document.addEventListener('DOMContentLoaded', () => {
    const form = document.getElementById('form');
    const previewButton = document.getElementById('previewButton');
    const preview = document.getElementById('preview');
    const saveButton = document.getElementById('save');

    previewButton.addEventListener('click', () => {
        const name = document.getElementById('name').value;
        const email = document.getElementById('email').value;
        const phone = document.getElementById('phone').value;

        if (name && email && phone) {
            preview.innerHTML = `
                <h3>Preview:</h3>
                <p>Name: ${name}</p>
                <p>Email: ${email}</p>
                <p>Phone: ${phone}</p>
            `;
            saveButton.style.display = 'block';
        } else {
            preview.innerHTML = '<p>Please fill in all fields.</p>';
            saveButton.style.display = 'none';
        }
    });

    saveButton.addEventListener('click', () => {
        const name = document.getElementById('name').value;
        const email = document.getElementById('email').value;
        const phone = document.getElementById('phone').value;

        const data = [
            { Name: name, Email: email, Phone: phone }
        ];

        const ws = XLSX.utils.json_to_sheet(data);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

        const filePath = path.join(__dirname, 'details.xlsx');
        XLSX.writeFile(wb, filePath);

        preview.innerHTML = '<p>Details saved to Excel.</p>';
        saveButton.style.display = 'none';
    });
});
