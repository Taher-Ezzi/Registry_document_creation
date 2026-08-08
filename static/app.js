// ==========================================================================
// REGISTRY ARCHITECT // ULTRA-REFRESHING CLIENT CONTROLLER
// Real-time Deed Inspector, One-Click Presets, Dynamic Multi-Buyer Manager
// ==========================================================================

let extractedData = null;
let currentBuyers = [];

const HINDI_ONES = {
    0: '', 1: 'एक', 2: 'दो', 3: 'तीन', 4: 'चार', 5: 'पाँच', 6: 'छह', 7: 'सात', 8: 'आठ', 9: 'नौ',
    10: 'दस', 11: 'ग्यारह', 12: 'बारह', 13: 'तेरह', 14: 'चौदह', 15: 'पंद्रह', 16: 'सोलह', 17: 'सत्रह', 18: 'अठारह', 19: 'उन्नीस',
    20: 'बीस', 21: 'इक्कीस', 22: 'बाईस', 23: 'तेईस', 24: 'चौबीस', 25: 'पच्चीस', 26: 'छब्बीस', 27: 'सत्ताईस', 28: 'अट्ठाईस', 29: 'उनत्तीस',
    30: 'तीस', 31: 'इकत्तीस', 32: 'बत्तीस', 33: 'तैंतीस', 34: 'चौंतीस', 35: 'पैंतीस', 36: 'छत्तीस', 37: 'सैंतीस', 38: 'अड़तीस', 39: 'उनतालीस',
    40: 'चालीस', 41: 'इकतालीस', 42: 'बयालीस', 43: 'तिरालीस', 44: 'चौवालिस', 45: 'पैंतालीस', 46: 'छियालीस', 47: 'सैंतालीस', 48: 'अड़तालीस', 49: 'उनचास',
    50: 'पचास', 51: 'इक्कावन', 52: 'बावन', 53: 'तिरेपन', 54: 'चौवन', 55: 'पचपन', 56: 'छप्पन', 57: 'सत्तावन', 58: 'अट्टावन', 59: 'उनसठ',
    60: 'साठ', 61: 'इक्सठ', 62: 'बासठ', 63: 'तिर्सठ', 64: 'चौंसठ', 65: 'पैंसठ', 66: 'छियासठ', 67: 'सरसठ', 68: 'अड़सठ', 69: 'उनसत्तर',
    70: 'सत्तर', 71: 'इकहत्तर', 72: 'बहत्तर', 73: 'तिहत्तर', 74: 'चौहत्तर', 75: 'पचहत्तर', 76: 'छियात्तर', 77: 'सतहत्तर', 78: 'इठ्‌योत्तर', 79: 'उनासी',
    80: 'अस्सी', 81: 'इक्यासी', 82: 'बयासी', 83: 'तिरासी', 84: 'चौरासी', 85: 'पचासी', 86: 'छियासी', 87: 'सत्तासी', 88: 'अ्ठासी', 89: 'नवासी',
    90: 'नब्बे', 91: 'इक्यान्वे', 92: 'बान्वे', 93: 'तिरान्वे', 94: 'चौरान्वे', 95: 'पन्चान्वे', 96: 'छियान्वे', 97: 'संतान्वे', 98: 'अन्ठान्वे', 99: 'निन्यानवे'
};

function numberToHindiWords(n) {
    n = parseInt(n) || 0;
    if (n === 0) return '';
    const parts = [];
    const crores = Math.floor(n / 10000000);
    n %= 10000000;
    if (crores > 0) parts.push(`${HINDI_ONES[crores] || crores} करोड़`);
    
    const lakhs = Math.floor(n / 100000);
    n %= 100000;
    if (lakhs > 0) parts.push(`${HINDI_ONES[lakhs] || lakhs} लाख`);
    
    const thousands = Math.floor(n / 1000);
    n %= 1000;
    if (thousands > 0) parts.push(`${HINDI_ONES[thousands] || thousands} हज़ार`);
    
    const hundreds = Math.floor(n / 100);
    n %= 100;
    if (hundreds > 0) parts.push(`${HINDI_ONES[hundreds] || hundreds} सौ`);
    
    if (n > 0) parts.push(HINDI_ONES[n] || n);
    
    return parts.join(' ') + ' मात्र';
}

function formatIndianCurrency(n) {
    if (!n) return '';
    const s = Math.round(Number(n)).toString();
    if (s.length <= 3) return s;
    const last3 = s.substring(s.length - 3);
    let other = s.substring(0, s.length - 3);
    let res = '';
    while (other.length > 2) {
        res = ',' + other.substring(other.length - 2) + res;
        other = other.substring(0, other.length - 2);
    }
    return other + res + ',' + last3;
}

// Toast Notification Engine
function showToast(message, type = 'info') {
    const container = document.getElementById('toast-container');
    if (!container) return;
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    toast.innerHTML = `<span>${type === 'success' ? '✓' : '⚡'}</span> ${message}`;
    container.appendChild(toast);
    setTimeout(() => {
        toast.style.opacity = '0';
        toast.style.transform = 'translateY(12px)';
        setTimeout(() => toast.remove(), 300);
    }, 4000);
}

// Multi-Buyer Card Management
function addNewBuyer() {
    currentBuyers.push({
        name: '',
        relation_title: 'पति श्री',
        relation_name: '',
        pan_no: '',
        aadhar_no: '',
        address: currentBuyers.length > 0 ? currentBuyers[0].address : ''
    });
    renderBuyers(currentBuyers);
    showToast(`Added Co-Buyer #${currentBuyers.length}`, 'info');
    updateLiveInspector();
}

function removeBuyer(idx) {
    if (currentBuyers.length > 1) {
        currentBuyers.splice(idx, 1);
        renderBuyers(currentBuyers);
        showToast('Co-buyer removed', 'info');
        updateLiveInspector();
    }
}

function renderBuyers(buyersList) {
    const container = document.getElementById('buyers-container');
    container.innerHTML = '';
    const badge = document.getElementById('badge-buyer-count');
    if (badge) {
        badge.textContent = `${buyersList.length} Buyer${buyersList.length > 1 ? 's (Joint Applicants)' : ''}`;
    }

    buyersList.forEach((b, idx) => {
        const div = document.createElement('div');
        div.className = 'buyer-entry-box';

        const initials = b.name ? b.name.replace(/[^A-Za-z\u0900-\u097F]/g, '').slice(0, 2).toUpperCase() : `B${idx + 1}`;

        div.innerHTML = `
            <div class="buyer-entry-header">
                <span class="buyer-num">
                    <span class="buyer-avatar-pill">${initials}</span> 
                    क्रेता / आवेदक क्रमांक ${idx + 1} ${idx === 0 ? '(मुख्य क्रेता / Primary)' : '(सह-क्रेता / Co-Buyer)'}
                </span>
                ${buyersList.length > 1 ? `<button type="button" class="btn-remove-buyer" onclick="removeBuyer(${idx})">✕ Remove Buyer</button>` : ''}
            </div>
            <div class="field-grid">
                <div class="field-group" style="grid-column: span 2;">
                    <label>क्रेता का नाम (Buyer Full Name)</label>
                    <input type="text" id="buyer-name-${idx}" value="${b.name || ''}" placeholder="श्री / श्रीमती...">
                </div>
                <div class="field-group">
                    <label>सम्बन्ध (Relation)</label>
                    <select id="buyer-rel-title-${idx}">
                        <option value="पति श्री" ${b.relation_title === 'पति श्री' ? 'selected' : ''}>पति श्री (Husband)</option>
                        <option value="पिता श्री" ${b.relation_title === 'पिता श्री' ? 'selected' : ''}>पिता श्री (Father)</option>
                        <option value="आत्मज श्री" ${b.relation_title === 'आत्मज श्री' ? 'selected' : ''}>आत्मज श्री (Son of)</option>
                        <option value="पत्नी श्रीमती" ${b.relation_title === 'पत्नी श्रीमती' ? 'selected' : ''}>पत्नी श्रीमती (Wife of)</option>
                        <option value="माता श्रीमती" ${b.relation_title === 'माता श्रीमती' ? 'selected' : ''}>माता श्रीमती (Mother)</option>
                    </select>
                </div>
                <div class="field-group">
                    <label>सम्बन्धी का नाम (Relative Name)</label>
                    <input type="text" id="buyer-rel-name-${idx}" value="${b.relation_name || ''}" placeholder="रिश्तेदार का नाम...">
                </div>
            </div>
            <div class="field-grid">
                <div class="field-group">
                    <label>PAN Number (स्थाई लेखा संख्या)</label>
                    <input type="text" id="buyer-pan-${idx}" value="${b.pan_no || ''}" placeholder="ABCDE1234F" style="text-transform: uppercase;">
                </div>
                <div class="field-group">
                    <label>Aadhaar Number (12 Digit)</label>
                    <input type="text" id="buyer-aadhar-${idx}" value="${b.aadhar_no || ''}" placeholder="XXXX XXXX XXXX">
                </div>
            </div>
            <div class="field-group full-width">
                <label>स्थाई पता (Residential Address)</label>
                <input type="text" id="buyer-addr-${idx}" value="${b.address || ''}" placeholder="मकान नम्बर, कॉलोनी, शहर, पिनकोड...">
            </div>
        `;
        container.appendChild(div);

        // Bind live preview listeners
        ['name', 'rel-name', 'pan', 'aadhar', 'addr'].forEach(f => {
            const el = div.querySelector(`#buyer-${f}-${idx}`);
            if (el) el.addEventListener('input', () => updateLiveInspector());
        });
        const sel = div.querySelector(`#buyer-rel-title-${idx}`);
        if (sel) sel.addEventListener('change', () => updateLiveInspector());
    });
}

const HINDI_MONTHS_LIST = ['', 'जनवरी', 'फरवरी', 'मार्च', 'अप्रैल', 'मई', 'जून', 'जुलाई', 'अगस्त', 'सितम्बर', 'अक्टूबर', 'नवम्बर', 'दिसम्बर'];

function getTodayHindiDate() {
    const today = new Date();
    const day = today.getDate();
    const month = HINDI_MONTHS_LIST[today.getMonth() + 1];
    const year = today.getFullYear();
    return `${day} ${month}, ${year}.`;
}

// Live Real-Time Deed Inspector Sync
function updateLiveInspector() {
    let plotNo = document.getElementById('input-plot-no').value || 'E9';
    plotNo = plotNo.replace(/\s*\([^)]*\)/g, '').trim();
    const plotNoHindi = document.getElementById('input-plot-hindi').value || 'ई-नौ';
    const colony = (extractedData && extractedData.colony_name) ? extractedData.colony_name : 'एमराल्ड आश्रय';
    const allotment = document.getElementById('input-allotment').value;
    const formattedPrice = formatIndianCurrency(allotment);
    const wordsPrice = allotment ? numberToHindiWords(allotment) : 'सात लाख तीस हज़ार पाँच सौ संतान्वे मात्र';

    const east = document.getElementById('input-east').value || 'भूखण्ड क्रमांक L1, एमराल्ड आश्रय';
    const west = document.getElementById('input-west').value || 'भूखण्ड क्रमांक ई-8, एमराल्ड आश्रय';
    const north = document.getElementById('input-north').value || 'भूखण्ड क्रमांक 85 एवं 86, एमराल्ड आश्रय';
    const south = document.getElementById('input-south').value || '9.00 मीटर चौड़ी सड़क';

    // 1. Top Banner
    const elBanner = document.getElementById('preview-banner-text');
    if (elBanner) {
        elBanner.innerHTML = `<mark>भूखण्ड क्रमांक ${plotNo} (${plotNoHindi})</mark>, ${colony}, ग्राम-सोनवाय, तहसील-राऊ, ज़िला-इन्दौर`;
    }

    // 2. Preamble Clause
    const elPreamble = document.getElementById('preview-preamble-text');
    if (elPreamble) {
        elPreamble.innerHTML = `...‘${colony}’ में ‘आर्थिक रूप से कमज़ोर श्रेणी’ (EWS) हेतु आरक्षित भूखण्डों में से भूखण्ड क्रमांक <mark>${plotNo} (${plotNoHindi})</mark>, जिसका विक्रय व्यवहार मूल्य <mark>रुपये ${formattedPrice}/- (अक्षरी रुपये ${wordsPrice})</mark> है, का विक्रय-पत्र`;
    }

    // 3. 2-Column Boundaries
    const elBounds = document.getElementById('preview-boundary-text');
    if (elBounds) {
        elBounds.innerHTML = `<strong>पूर्व:</strong> ${east} | <strong>पश्चिम:</strong> ${west} <br><strong>उत्तर:</strong> ${north} | <strong>दक्षिण:</strong> ${south}`;
    }

    // 4. Closing Price Clause
    const elClosing = document.getElementById('preview-closing-text');
    if (elClosing) {
        elClosing.innerHTML = `...उपरोक्त वर्णित अनुसार चुकता विक्रय मूल्य <mark>रुपये ${formattedPrice}/- (अक्षरी रुपये ${wordsPrice})</mark> प्राप्त करने के उपरान्त विक्रय कर दिया है... (दिनांक : ${getTodayHindiDate()})`;
    }
}

// Preset Handlers for 1-Click Verification
const PRESETS = {
    e9: {
        colony_name: 'एमराल्ड आश्रय',
        plot_data: { plot_no: 'E9', plot_no_hindi: 'ई-नौ', width: '', length: '', area_sqm: 39.67, area_sqft: 427 },
        allotment_val: 730597,
        execution_date: getTodayHindiDate(),
        boundaries: {
            east: 'भूखण्ड क्रमांक L1, एमराल्ड आश्रय',
            west: 'भूखण्ड क्रमांक ई-8, एमराल्ड आश्रय',
            north: 'भूखण्ड क्रमांक 85 एवं 86, एमराल्ड आश्रय',
            south: '9.00 मीटर चौड़ी सड़क'
        },
        kyc_data: {
            buyers: [{
                name: 'श्रीमती खुशबू यादव',
                relation_title: 'पति श्री',
                relation_name: 'पवन यादव',
                pan_no: 'LUKPK3635A',
                aadhar_no: '9913 9334 4052',
                address: 'मकान नं. 1737, द्वारकापुरी, फूटी कोठी रोड, सुदामा नगर, इंदौर, मध्य प्रदेश - 452009'
            }]
        },
        payments: [
            { s_no: 1, amount: 200000, bank_hindi: 'कोटक महिन्द्रा बैंक', date_hindi: '30 जुलाई, 2026', mode: 'Online', ref_no: 'KKBKH26211830056', is_utr: false },
            { s_no: 2, amount: 500000, bank_hindi: 'कोटक महिन्द्रा बैंक', date_hindi: '01 अगस्त, 2026', mode: 'Online', ref_no: 'KKBKH26213830998', is_utr: false },
            { s_no: 3, amount: 30597, bank_hindi: 'कोटक महिन्द्रा बैंक', date_hindi: '04 अगस्त, 2026', mode: 'Online', ref_no: 'KKBKH26216629311', is_utr: false }
        ],
        tds_info: null
    },
    e8: {
        colony_name: 'एमराल्ड आश्रय',
        plot_data: { plot_no: 'E8', plot_no_hindi: 'ई-आठ', width: '', length: '', area_sqm: 39.67, area_sqft: 427 },
        allotment_val: 750000,
        execution_date: getTodayHindiDate(),
        boundaries: {
            east: 'भूखण्ड क्रमांक ई-9, एमराल्ड आश्रय',
            west: 'भूखण्ड क्रमांक ई-7, एमराल्ड आश्रय',
            north: 'भूखण्ड क्रमांक 85, एमराल्ड आश्रय',
            south: '9.00 मीटर चौड़ी सड़क'
        },
        kyc_data: {
            buyers: [{
                name: 'श्री तनुज परिहार',
                relation_title: 'पिता श्री',
                relation_name: 'राजू परिहार',
                pan_no: 'EGGPP7796Q',
                aadhar_no: '3544 3057 7863',
                address: 'मकान नं. 497-ऋषी पैलेस, कॉलोनी, इंदौर, मध्य प्रदेश - 452009'
            }]
        },
        payments: [
            { s_no: 1, amount: 375000, bank_hindi: 'पंजाब नेशनल बैंक', date_hindi: '15 जुलाई, 2026', mode: 'RTGS', ref_no: 'PUNBH26201994821', is_utr: true },
            { s_no: 2, amount: 375000, bank_hindi: 'पंजाब नेशनल बैंक', date_hindi: '22 जुलाई, 2026', mode: 'RTGS', ref_no: 'PUNBH26208492019', is_utr: true }
        ],
        tds_info: null
    },
    l1: {
        colony_name: 'एमराल्ड गेटवे',
        plot_data: { plot_no: 'L-1', plot_no_hindi: 'एल-एक', width: 4.00, length: 9.61, area_sqm: 38.44, area_sqft: 414 },
        allotment_val: 825000,
        execution_date: getTodayHindiDate(),
        boundaries: {
            east: 'भूखण्ड क्रमांक 11, एमराल्ड गेटवे',
            west: 'भूखण्ड क्रमांक L2, एमराल्ड गेटवे',
            north: 'गार्डन, एमराल्ड गेटवे',
            south: '12.00 मीटर चौड़ी सड़क'
        },
        kyc_data: {
            buyers: [{
                name: 'श्रीमती सरिका सेठी',
                relation_title: 'पति श्री',
                relation_name: 'अमित कुमार सेठी',
                pan_no: 'BUFPS3575N',
                aadhar_no: '2432 9492 5984',
                address: 'मकान नम्बर-22, झंडा बाज़ार के पास, विनोभा भवन वार्ड-6, सिहोर, जबलपुर-483225 (म.प्र.)'
            }]
        },
        payments: [
            { s_no: 1, amount: 200000, bank_hindi: 'एक्सीस बैंक', date_hindi: '10 जून, 2026', mode: 'RTGS', ref_no: 'UTIBH26161839201', is_utr: true },
            { s_no: 2, amount: 300000, bank_hindi: 'एक्सीस बैंक', date_hindi: '18 जून, 2026', mode: 'RTGS', ref_no: 'UTIBH26169948201', is_utr: true },
            { s_no: 3, amount: 317750, bank_hindi: 'एक्सीस बैंक', date_hindi: '25 जून, 2026', mode: 'RTGS', ref_no: 'UTIBH26176629401', is_utr: true }
        ],
        tds_info: {
            amount: 7250,
            bank_hindi: 'एक्सीस बैंक',
            date_hindi: '28 जून, 2026',
            challan_no: '01568',
            bsr_code: '6360014'
        }
    }
};

function loadPreset(key) {
    const data = PRESETS[key];
    if (!data) return;
    extractedData = data;
    const reviewSec = document.getElementById('review-section');
    reviewSec.style.display = 'block';
    populateReviewForm(data);
    showToast(`Loaded preset: ${data.plot_data.plot_no} (${data.kyc_data.buyers[0].name})`, 'success');
    
    // Highlight active preset chip
    document.querySelectorAll('.preset-chip').forEach(c => c.classList.remove('active'));
    const activeChip = document.getElementById(`preset-${key}`);
    if (activeChip) activeChip.classList.add('active');

    // Stepper navigation update
    document.getElementById('step-nav-1').classList.add('completed');
    document.getElementById('step-nav-2').classList.add('active');

    reviewSec.scrollIntoView({ behavior: 'smooth' });
}

function clearAllFields() {
    extractedData = null;
    currentBuyers = [];
    document.querySelectorAll('input').forEach(i => i.value = '');
    document.querySelectorAll('.dropzone').forEach(d => d.classList.remove('filled'));
    document.querySelectorAll('.dropzone-hint').forEach(h => h.textContent = 'Drag & drop or select file');
    document.getElementById('payments-tbody').innerHTML = '';
    document.getElementById('buyers-container').innerHTML = '';
    document.getElementById('preview-hindi-amount').textContent = '';
    document.getElementById('download-card').style.display = 'none';
    document.getElementById('review-section').style.display = 'none';
    document.querySelectorAll('.preset-chip').forEach(c => c.classList.remove('active'));
    showToast('All fields cleared and reset', 'info');
}

document.addEventListener('DOMContentLoaded', () => {
    // Dropzone File Selectors
    const fileInputs = ['template', 'excel', 'kyc', 'png'];
    fileInputs.forEach(id => {
        const input = document.getElementById(`file-${id}`);
        const hint = document.getElementById(`hint-${id}`);
        const zone = document.getElementById(`dropzone-${id}`);
        if (input && hint) {
            input.addEventListener('change', () => {
                if (input.files && input.files.length > 0) {
                    if (input.files.length === 1) {
                        hint.textContent = `✓ ${input.files[0].name}`;
                    } else {
                        hint.textContent = `✓ ${input.files.length} files selected`;
                    }
                    zone.classList.add('filled');
                    showToast(`Loaded ${input.files.length} file(s) into ${id.toUpperCase()}`, 'success');
                }
            });
        }
    });

    // Preset Chips
    document.getElementById('preset-e9').addEventListener('click', () => loadPreset('e9'));
    document.getElementById('preset-e8').addEventListener('click', () => loadPreset('e8'));
    document.getElementById('preset-l1').addEventListener('click', () => loadPreset('l1'));
    document.getElementById('btn-reset-all').addEventListener('click', clearAllFields);

    // Live Price in Words
    const allotmentInput = document.getElementById('input-allotment');
    const hindiPreview = document.getElementById('preview-hindi-amount');
    if (allotmentInput && hindiPreview) {
        allotmentInput.addEventListener('input', () => {
            hindiPreview.textContent = numberToHindiWords(allotmentInput.value);
            updateLiveInspector();
        });
    }

    // Live Input Listeners for Real-Time Deed Inspector
    ['plot-no', 'plot-hindi', 'width', 'length', 'sqm', 'sqft', 'east', 'west', 'north', 'south'].forEach(id => {
        const el = document.getElementById(`input-${id}`);
        if (el) el.addEventListener('input', () => updateLiveInspector());
    });

    // Action Buttons
    document.getElementById('btn-extract').addEventListener('click', () => performExtraction());
    document.getElementById('btn-load-defaults').addEventListener('click', () => performExtraction(true));
    document.getElementById('btn-generate').addEventListener('click', () => generateRegistryDoc());
    
    const btnAddTop = document.getElementById('btn-add-buyer-top');
    const btnAddBottom = document.getElementById('btn-add-buyer-bottom');
    if (btnAddTop) btnAddTop.addEventListener('click', addNewBuyer);
    if (btnAddBottom) btnAddBottom.addEventListener('click', addNewBuyer);
});

async function performExtraction(useDefaults = false) {
    const btn = document.getElementById('btn-extract');
    const oldText = btn.innerHTML;
    btn.innerHTML = '⏳ Auto-Extracting Data...';
    btn.disabled = true;

    try {
        const formData = new FormData();
        if (!useDefaults) {
            const templateFile = document.getElementById('file-template').files[0];
            const excelFile = document.getElementById('file-excel').files[0];
            const pngFile = document.getElementById('file-png').files[0];
            const kycFiles = document.getElementById('file-kyc').files;

            if (templateFile) formData.append('template_file', templateFile);
            if (excelFile) formData.append('excel_file', excelFile);
            if (pngFile) formData.append('png_file', pngFile);
            for (let i = 0; i < kycFiles.length; i++) {
                formData.append('kyc_files', kycFiles[i]);
            }
        }

        const res = await fetch('/api/extract', {
            method: 'POST',
            body: formData
        });

        extractedData = await res.json();
        
        const reviewSec = document.getElementById('review-section');
        reviewSec.style.display = 'block';
        
        populateReviewForm(extractedData);
        showToast('✓ Auto-extracted all plot, payment, and KYC details!', 'success');

        document.getElementById('step-nav-1').classList.add('completed');
        document.getElementById('step-nav-2').classList.add('active');

        reviewSec.scrollIntoView({ behavior: 'smooth' });

    } catch (err) {
        showToast('Extraction failed: ' + err.message, 'info');
    } finally {
        btn.innerHTML = oldText;
        btn.disabled = false;
    }
}

function populateReviewForm(data) {
    // 1. Multi-Buyers
    if (data.kyc_data && data.kyc_data.buyers && data.kyc_data.buyers.length > 0) {
        currentBuyers = data.kyc_data.buyers;
    } else {
        currentBuyers = [{
            name: data.kyc_data ? data.kyc_data.primary_name : '',
            relation_title: data.kyc_data ? data.kyc_data.relation_title : 'पति श्री',
            relation_name: data.kyc_data ? data.kyc_data.relation_name : '',
            pan_no: data.kyc_data ? data.kyc_data.pan_no : '',
            aadhar_no: data.kyc_data ? data.kyc_data.aadhar_no : '',
            address: data.kyc_data ? data.kyc_data.address : ''
        }];
    }
    renderBuyers(currentBuyers);

    // 2. Plot Details
    if (data.plot_data) {
        document.getElementById('input-plot-no').value = data.plot_data.plot_no || '';
        document.getElementById('input-plot-hindi').value = data.plot_data.plot_no_hindi || '';
        document.getElementById('input-width').value = data.plot_data.width || '';
        document.getElementById('input-length').value = data.plot_data.length || '';
        document.getElementById('input-sqm').value = data.plot_data.area_sqm || '';
        document.getElementById('input-sqft').value = data.plot_data.area_sqft || '';
    }

    if (data.colony_name) {
        const badge = document.getElementById('badge-colony-name');
        if (badge) badge.textContent = data.colony_name;
    }

    document.getElementById('input-allotment').value = data.allotment_val || '';
    document.getElementById('preview-hindi-amount').textContent = data.allotment_val_words ? data.allotment_val_words : (data.allotment_val ? numberToHindiWords(data.allotment_val) : '');

    // 3. Boundaries (from Excel)
    if (data.boundaries) {
        document.getElementById('input-east').value = data.boundaries.east || '';
        document.getElementById('input-west').value = data.boundaries.west || '';
        document.getElementById('input-north').value = data.boundaries.north || '';
        document.getElementById('input-south').value = data.boundaries.south || '';
    }

    // 4. Payments Table (from Excel)
    const tbody = document.getElementById('payments-tbody');
    tbody.innerHTML = '';
    const payments = data.payments || [];
    document.getElementById('badge-payment-count').textContent = `${payments.length + (data.tds_info ? 1 : 0)} Payments from Excel`;

    payments.forEach((p, idx) => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>#${p.s_no || (idx + 1)}</td>
            <td><strong>₹${Number(p.amount).toLocaleString('en-IN')}/-</strong></td>
            <td>${p.bank_hindi}</td>
            <td>${p.date_hindi}</td>
            <td><span class="code-pill">${p.mode}</span></td>
            <td><code>${p.ref_no}</code></td>
        `;
        tbody.appendChild(tr);
    });

    if (data.tds_info) {
        const tr = document.createElement('tr');
        tr.style.background = 'rgba(6, 182, 212, 0.05)';
        tr.innerHTML = `
            <td><span class="badge-pill">TDS</span></td>
            <td><strong>₹${Number(data.tds_info.amount).toLocaleString('en-IN')}/-</strong></td>
            <td>${data.tds_info.bank_hindi} (BSR: ${data.tds_info.bsr_code})</td>
            <td>${data.tds_info.date_hindi}</td>
            <td><span class="code-pill">Challan 393(1)</span></td>
            <td><code>Challan: ${data.tds_info.challan_no}</code></td>
        `;
        tbody.appendChild(tr);
    }

    updateLiveInspector();
}

async function generateRegistryDoc() {
    const btn = document.getElementById('btn-generate');
    const oldText = btn.innerHTML;
    btn.innerHTML = '⚙️ Compiling Legal Word Deed (.docx)...';
    btn.disabled = true;

    try {
        const buyersPayload = currentBuyers.map((_, idx) => ({
            name: document.getElementById(`buyer-name-${idx}`).value,
            relation_title: document.getElementById(`buyer-rel-title-${idx}`).value,
            relation_name: document.getElementById(`buyer-rel-name-${idx}`).value,
            pan_no: document.getElementById(`buyer-pan-${idx}`).value,
            aadhar_no: document.getElementById(`buyer-aadhar-${idx}`).value,
            address: document.getElementById(`buyer-addr-${idx}`).value
        }));

        const payload = {
            template_path: extractedData ? extractedData.template_path : null,
            colony_name: extractedData ? extractedData.colony_name : 'एमराल्ड आश्रय',
            buyers: buyersPayload,
            plot_data: {
                plot_no: document.getElementById('input-plot-no').value,
                plot_no_hindi: document.getElementById('input-plot-hindi').value,
                width: document.getElementById('input-width').value,
                length: document.getElementById('input-length').value,
                area_sqm: parseFloat(document.getElementById('input-sqm').value) || 0,
                area_sqft: parseInt(document.getElementById('input-sqft').value) || 0
            },
            allotment_val: parseInt(document.getElementById('input-allotment').value) || 0,
            boundaries: {
                east: document.getElementById('input-east').value,
                west: document.getElementById('input-west').value,
                north: document.getElementById('input-north').value,
                south: document.getElementById('input-south').value
            },
            payments: extractedData && extractedData.payments ? extractedData.payments : [],
            tds_info: extractedData ? extractedData.tds_info : null,
            execution_date: (extractedData && extractedData.execution_date) ? extractedData.execution_date : getTodayHindiDate()
        };

        const res = await fetch('/api/generate', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });

        const result = await res.json();
        if (result.status === 'success') {
            const card = document.getElementById('download-card');
            const link = document.getElementById('btn-download');
            link.href = result.download_url;
            card.style.display = 'flex';
            
            document.getElementById('step-nav-2').classList.add('completed');
            document.getElementById('step-nav-3').classList.add('active');

            showToast('🎉 Legal Deed Word document compiled with 100% precision!', 'success');
            card.scrollIntoView({ behavior: 'smooth' });
        }
    } catch (err) {
        showToast('Generation failed: ' + err.message, 'info');
    } finally {
        btn.innerHTML = oldText;
        btn.disabled = false;
    }
}
