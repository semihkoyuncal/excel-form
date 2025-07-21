document.getElementById('excelFile').addEventListener('change', function(e) {
    const file = e.target.files[0];
    const reader = new FileReader();

    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        
        const jsonData = XLSX.utils.sheet_to_json(firstSheet)[0];

        document.getElementById('name').value = jsonData['Ad Soyad'] || '';
        document.getElementById('email').value = jsonData['E-posta'] || '';
        document.getElementById('phone').value = jsonData['Telefon'] || '';
        document.getElementById('address').value = jsonData['Adres'] || '';
    };

    reader.readAsArrayBuffer(file);
}); 