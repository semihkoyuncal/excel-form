let allPersonsData = [];
let currentPersonIndex = -1;


document.getElementById('excelFile').addEventListener('change', function(e) {
    const file = e.target.files[0];
    
    if (!file) {
        showMessage('Lütfen bir Excel dosyası seçin.', 'error');
        return;
    }
    
    const reader = new FileReader();

    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            
            
            const jsonData = XLSX.utils.sheet_to_json(firstSheet);
            
            if (jsonData.length === 0) {
                showMessage('Excel dosyası boş veya geçersiz format.', 'error');
                return;
            }
            
            
            allPersonsData = jsonData;
            
            
            populatePersonSelector();
            
            
            if (allPersonsData.length > 0) {
                currentPersonIndex = 0;
                document.getElementById('personSelect').selectedIndex = 1; // İlk gerçek seçeneği seç
                fillFormWithPersonData(0);
            }
            
            showMessage(`${allPersonsData.length} kişi başarıyla yüklendi!`, 'success');
            
        } catch (error) {
            console.error('Excel okuma hatası:', error);
            showMessage('Excel dosyası okunamadı. Dosya formatını kontrol edin.', 'error');
        }
    };

    reader.readAsArrayBuffer(file);
});


function populatePersonSelector() {
    const personSelect = document.getElementById('personSelect');
    const personSelector = document.getElementById('personSelector');
    
    
    personSelect.innerHTML = '<option value="">-- Kişi Seçin --</option>';
    
    
    allPersonsData.forEach((person, index) => {
        const option = document.createElement('option');
        option.value = index;
        option.textContent = person['Ad Soyad'] || `Kişi ${index + 1}`;
        personSelect.appendChild(option);
    });
    
    
    personSelector.style.display = 'block';
    
    
    personSelect.addEventListener('change', function() {
        const selectedIndex = parseInt(this.value);
        if (!isNaN(selectedIndex)) {
            currentPersonIndex = selectedIndex;
            fillFormWithPersonData(selectedIndex);
            updatePersonInfo();
            updateNavigationButtons();
        } else {
            clearForm();
            currentPersonIndex = -1;
        }
    });
    
    updatePersonInfo();
    updateNavigationButtons();
}


function fillFormWithPersonData(index) {
    if (index < 0 || index >= allPersonsData.length) return;
    
    const person = allPersonsData[index];
    
    document.getElementById('name').value = person['Ad Soyad'] || '';
    document.getElementById('email').value = person['E-posta'] || '';
    document.getElementById('phone').value = person['Telefon'] || '';
    document.getElementById('address').value = person['Adres'] || '';
}

function navigatePerson(direction) {
    const newIndex = currentPersonIndex + direction;
    
    if (newIndex >= 0 && newIndex < allPersonsData.length) {
        currentPersonIndex = newIndex;
        document.getElementById('personSelect').selectedIndex = newIndex + 1;
        fillFormWithPersonData(newIndex);
        updatePersonInfo();
        updateNavigationButtons();
    }
}


function updatePersonInfo() {
    const personInfo = document.getElementById('personInfo');
    
    if (currentPersonIndex >= 0 && currentPersonIndex < allPersonsData.length) {
        personInfo.textContent = `${currentPersonIndex + 1} / ${allPersonsData.length} kişi`;
    } else {
        personInfo.textContent = `Toplam ${allPersonsData.length} kişi`;
    }
}

function updateNavigationButtons() {
    const prevBtn = document.getElementById('prevBtn');
    const nextBtn = document.getElementById('nextBtn');
    
    prevBtn.disabled = currentPersonIndex <= 0;
    nextBtn.disabled = currentPersonIndex >= allPersonsData.length - 1;
}


function clearForm() {
    document.getElementById('name').value = '';
    document.getElementById('email').value = '';
    document.getElementById('phone').value = '';
    document.getElementById('address').value = '';
}


function exportCurrentPerson() {
    if (currentPersonIndex >= 0 && currentPersonIndex < allPersonsData.length) {
        const person = allPersonsData[currentPersonIndex];
        const jsonStr = JSON.stringify(person, null, 2);
        
        
        console.log('Seçilen kişi:', person);
        
        
        const blob = new Blob([jsonStr], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `${person['Ad Soyad'] || 'kisi'}.json`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        
        showMessage(`${person['Ad Soyad'] || 'Kişi'} bilgileri dışa aktarıldı!`, 'success');
    } else {
        showMessage('Lütfen önce bir kişi seçin.', 'error');
    }
}


function showMessage(message, type = 'info') {
    const messageArea = document.getElementById('messageArea');
    messageArea.innerHTML = `<div class="message ${type}">${message}</div>`;
    
    
    setTimeout(() => {
        messageArea.innerHTML = '';
    }, 5000);
}


document.addEventListener('keydown', function(e) {
    if (allPersonsData.length > 0) {
        if (e.key === 'ArrowLeft' && e.ctrlKey) {
            e.preventDefault();
            navigatePerson(-1);
        } else if (e.key === 'ArrowRight' && e.ctrlKey) {
            e.preventDefault();
            navigatePerson(1);
        }
    }
});