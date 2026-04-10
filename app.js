/* ============================================
   RTICCE-2K26 CERTIFICATE GENERATOR — ENGINE
   Canvas-based certificate rendering with
   Excel parsing, photo mapping, and batch export
   ============================================ */

// ===== STATE =====
const state = {
    participants: [],
    photos: {},          // filename -> dataURL
    certificates: [],    // { name, college, dataURL }
};

// ===== DOM REFS =====
const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => document.querySelectorAll(sel);

const els = {
    excelInput: $('#excel-input'),
    photosInput: $('#photos-input'),
    excelDrop: $('#excel-drop-zone'),
    photosDrop: $('#photos-drop-zone'),
    excelStatus: $('#excel-status'),
    photosStatus: $('#photos-status'),
    dataSection: $('#data-section'),
    dataCount: $('#data-count'),
    dataTbody: $('#data-tbody'),
    previewSection: $('#preview-section'),
    certGrid: $('#certificate-grid'),
    btnDemo: $('#btn-demo'),
    btnGenerate: $('#btn-generate'),
    btnDownloadAll: $('#btn-download-all'),
    modal: $('#preview-modal'),
    modalImg: $('#modal-img'),
    modalClose: $('#modal-close'),
    modalDownload: $('#modal-download'),
    toast: $('#toast'),
    toastMsg: $('#toast-msg'),
    progressOverlay: $('#progress-overlay'),
    progressBar: $('#progress-bar'),
    progressText: $('#progress-text'),
    progressTitle: $('#progress-title'),
    canvas: $('#cert-canvas'),
};

// ===== DEMO DATA =====
const DEMO_PARTICIPANTS = [
    { name: 'Arun Kumar S', college: 'PT Lee CNCET, Kanchipuram', paperTitle: 'IoT-Based Smart Agriculture System Using LoRa', photo: 'demo1' },
    { name: 'Priya Dharshini M', college: 'Anna University, Chennai', paperTitle: '5G Network Optimization Using Deep Learning', photo: 'demo2' },
    { name: 'Karthikeyan R', college: 'SRM Institute of Science, Chennai', paperTitle: 'VLSI Design of Low-Power ALU Architecture', photo: 'demo3' },
    { name: 'Divya Lakshmi T', college: 'VIT University, Vellore', paperTitle: 'Machine Learning in Medical Image Analysis', photo: 'demo4' },
    { name: 'Mohammed Rafi A', college: 'MIT Campus, Anna University', paperTitle: 'Embedded System for Autonomous Drone Control', photo: 'demo5' },
    { name: 'Swetha Kumari V', college: 'KVCET, Kanchipuram', paperTitle: 'Cloud Computing Security Framework Design', photo: 'demo6' },
];

// ===== GENERATE DEMO AVATAR =====
function generateDemoAvatar(name, index) {
    const c = document.createElement('canvas');
    c.width = 300;
    c.height = 300;
    const ctx = c.getContext('2d');

    const colors = [
        ['#667eea', '#764ba2'],
        ['#f093fb', '#f5576c'],
        ['#4facfe', '#00f2fe'],
        ['#43e97b', '#38f9d7'],
        ['#fa709a', '#fee140'],
        ['#a18cd1', '#fbc2eb'],
    ];

    const [c1, c2] = colors[index % colors.length];
    const grad = ctx.createLinearGradient(0, 0, 300, 300);
    grad.addColorStop(0, c1);
    grad.addColorStop(1, c2);
    ctx.fillStyle = grad;
    ctx.fillRect(0, 0, 300, 300);

    // Draw initials
    const initials = name.split(' ').map(w => w[0]).join('').substring(0, 2).toUpperCase();
    ctx.fillStyle = 'rgba(255,255,255,0.9)';
    ctx.font = 'bold 100px Inter, sans-serif';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText(initials, 150, 155);

    return c.toDataURL('image/png');
}

// ===== TOAST =====
function showToast(msg) {
    els.toastMsg.textContent = msg;
    els.toast.classList.remove('hidden');
    els.toast.classList.add('show');
    setTimeout(() => {
        els.toast.classList.remove('show');
        setTimeout(() => els.toast.classList.add('hidden'), 400);
    }, 3000);
}

// ===== PROGRESS =====
function showProgress(title) {
    els.progressTitle.textContent = title;
    els.progressBar.style.width = '0%';
    els.progressText.textContent = '0 / 0';
    els.progressOverlay.classList.remove('hidden');
}

function updateProgress(current, total) {
    const pct = Math.round((current / total) * 100);
    els.progressBar.style.width = pct + '%';
    els.progressText.textContent = `${current} / ${total}`;
}

function hideProgress() {
    els.progressOverlay.classList.add('hidden');
}

// ===== FILE HANDLERS =====
function setupDragDrop(zone, input) {
    zone.addEventListener('click', () => input.click());

    zone.addEventListener('dragover', (e) => {
        e.preventDefault();
        zone.classList.add('drag-over');
    });

    zone.addEventListener('dragleave', () => {
        zone.classList.remove('drag-over');
    });

    zone.addEventListener('drop', (e) => {
        e.preventDefault();
        zone.classList.remove('drag-over');
        const files = e.dataTransfer.files;
        if (input === els.excelInput) {
            handleExcel(files[0]);
        } else {
            handlePhotos(files);
        }
    });
}

function handleExcel(file) {
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            const json = XLSX.utils.sheet_to_json(sheet);

            if (json.length === 0) {
                showToast('Excel file is empty!');
                return;
            }

            // Map columns flexibly
            state.participants = json.map((row) => {
                const keys = Object.keys(row);
                return {
                    name: row['Name'] || row['name'] || row['FULL NAME'] || row['Full Name'] || row[keys[0]] || '',
                    college: row['College'] || row['college'] || row['COLLEGE'] || row['Institution'] || row[keys[1]] || '',
                    paperTitle: row['Paper Title'] || row['Paper'] || row['paper_title'] || row['PAPER TITLE'] || row['Title'] || row[keys[2]] || '',
                    photo: row['Photo'] || row['photo'] || row['PHOTO'] || row['Photo Filename'] || row['Image'] || row[keys[3]] || '',
                };
            });

            els.excelStatus.textContent = `✓ ${state.participants.length} participants loaded from "${file.name}"`;
            els.excelDrop.classList.add('uploaded');
            renderDataTable();
            showToast(`${state.participants.length} participants loaded successfully!`);
        } catch (err) {
            console.error(err);
            showToast('Error reading Excel file. Please check the format.');
        }
    };
    reader.readAsArrayBuffer(file);
}

function handlePhotos(files) {
    if (!files || files.length === 0) return;

    let loaded = 0;
    const total = files.length;

    Array.from(files).forEach((file) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            const nameWithoutExt = file.name.replace(/\.[^.]+$/, '');
            state.photos[file.name] = e.target.result;
            state.photos[nameWithoutExt] = e.target.result;
            state.photos[file.name.toLowerCase()] = e.target.result;
            state.photos[nameWithoutExt.toLowerCase()] = e.target.result;
            loaded++;

            if (loaded === total) {
                els.photosStatus.textContent = `✓ ${total} photos loaded`;
                els.photosDrop.classList.add('uploaded');
                renderDataTable();
                showToast(`${total} photos loaded successfully!`);
            }
        };
        reader.readAsDataURL(file);
    });
}

// ===== DATA TABLE =====
function renderDataTable() {
    if (state.participants.length === 0) return;

    els.dataSection.classList.remove('hidden');
    els.dataSection.classList.add('fade-in-up');
    els.dataCount.textContent = `${state.participants.length} participants loaded`;

    els.dataTbody.innerHTML = state.participants.map((p, i) => {
        const photoSrc = getPhotoForParticipant(p, i);
        const photoHTML = photoSrc
            ? `<img class="table-photo" src="${photoSrc}" alt="${p.name}">`
            : `<div class="table-photo-placeholder">${getInitials(p.name)}</div>`;

        return `
            <tr>
                <td>${i + 1}</td>
                <td>${photoHTML}</td>
                <td><strong>${p.name}</strong></td>
                <td>${p.college}</td>
                <td>${p.paperTitle}</td>
                <td>
                    <div class="td-actions">
                        <button class="btn-icon" onclick="previewSingle(${i})" title="Preview Certificate">
                            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="16" height="16">
                                <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/>
                            </svg>
                        </button>
                        <button class="btn-icon" onclick="downloadSingle(${i})" title="Download Certificate">
                            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="16" height="16">
                                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                                <polyline points="7 10 12 15 17 10"/>
                                <line x1="12" y1="15" x2="12" y2="3"/>
                            </svg>
                        </button>
                    </div>
                </td>
            </tr>
        `;
    }).join('');
}

function getInitials(name) {
    return name.split(' ').map(w => w[0]).join('').substring(0, 2).toUpperCase();
}

function getPhotoForParticipant(participant, index) {
    const photoKey = participant.photo;
    if (photoKey && photoKey.startsWith('demo')) {
        return generateDemoAvatar(participant.name, index);
    }
    if (state.photos[photoKey]) return state.photos[photoKey];
    if (state.photos[photoKey?.toLowerCase()]) return state.photos[photoKey?.toLowerCase()];
    const withoutExt = photoKey?.replace(/\.[^.]+$/, '');
    if (state.photos[withoutExt]) return state.photos[withoutExt];
    return null;
}

// ===== CERTIFICATE RENDERER =====
async function renderCertificate(participant, index) {
    const canvas = els.canvas;
    const ctx = canvas.getContext('2d');
    const W = 1600;
    const H = 1130;

    ctx.clearRect(0, 0, W, H);

    // ========= BACKGROUND =========
    const bgGrad = ctx.createLinearGradient(0, 0, W, H);
    bgGrad.addColorStop(0, '#0a0a14');
    bgGrad.addColorStop(0.5, '#0d0d1a');
    bgGrad.addColorStop(1, '#08081a');
    ctx.fillStyle = bgGrad;
    ctx.fillRect(0, 0, W, H);

    // Subtle radial glow at center
    const radGlow = ctx.createRadialGradient(W / 2, H / 2, 100, W / 2, H / 2, 600);
    radGlow.addColorStop(0, 'rgba(255, 215, 0, 0.03)');
    radGlow.addColorStop(1, 'transparent');
    ctx.fillStyle = radGlow;
    ctx.fillRect(0, 0, W, H);

    // ========= BORDER — TRIPLE LINE =========
    // Outer border
    ctx.strokeStyle = 'rgba(255, 215, 0, 0.3)';
    ctx.lineWidth = 3;
    roundRect(ctx, 20, 20, W - 40, H - 40, 16);
    ctx.stroke();

    // Middle border
    ctx.strokeStyle = 'rgba(255, 215, 0, 0.15)';
    ctx.lineWidth = 1;
    roundRect(ctx, 32, 32, W - 64, H - 64, 12);
    ctx.stroke();

    // Inner border
    ctx.strokeStyle = 'rgba(255, 215, 0, 0.4)';
    ctx.lineWidth = 2;
    roundRect(ctx, 44, 44, W - 88, H - 88, 10);
    ctx.stroke();

    // ========= CORNER DECORATIONS =========
    drawCornerDecor(ctx, 44, 44, 1, 1);
    drawCornerDecor(ctx, W - 44, 44, -1, 1);
    drawCornerDecor(ctx, 44, H - 44, 1, -1);
    drawCornerDecor(ctx, W - 44, H - 44, -1, -1);

    // ========= TOP DECORATIVE LINE =========
    const topLineY = 100;
    ctx.strokeStyle = 'rgba(255, 215, 0, 0.2)';
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(200, topLineY);
    ctx.lineTo(W - 200, topLineY);
    ctx.stroke();

    // Diamond at center of top line
    drawDiamond(ctx, W / 2, topLineY, 8, 'rgba(255, 215, 0, 0.5)');

    // ========= HEADER: COLLEGE NAME =========
    ctx.textAlign = 'center';
    ctx.fillStyle = 'rgba(255, 215, 0, 0.6)';
    ctx.font = '600 16px "Inter", sans-serif';
    ctx.letterSpacing = '4px';
    ctx.fillText('P.T. LEE CHENGALVARAYA NAICKER COLLEGE OF ENGINEERING & TECHNOLOGY', W / 2, 140);

    ctx.fillStyle = 'rgba(240, 240, 245, 0.35)';
    ctx.font = '400 13px "Inter", sans-serif';
    ctx.fillText('Department of Electronics & Communication Engineering', W / 2, 165);

    // ========= CONFERENCE NAME =========
    ctx.fillStyle = 'rgba(255, 215, 0, 0.08)';
    ctx.font = '900 120px "Cinzel", serif';
    ctx.fillText('RTICCE', W / 2, 290);

    ctx.fillStyle = 'rgba(255, 215, 0, 0.85)';
    ctx.font = '800 28px "Cinzel", serif';
    ctx.fillText('RTICCE-2K26', W / 2, 210);

    ctx.fillStyle = 'rgba(240, 240, 245, 0.45)';
    ctx.font = '400 14px "Inter", sans-serif';
    ctx.fillText('National Conference on Recent Trends in Information & Computer Communication Engineering', W / 2, 240);

    ctx.fillStyle = 'rgba(255, 215, 0, 0.4)';
    ctx.font = '500 13px "Inter", sans-serif';
    ctx.fillText('April 17, 2026 • Hybrid Mode', W / 2, 263);

    // ========= LINE SEPARATOR =========
    const sepY = 285;
    ctx.strokeStyle = 'rgba(255, 215, 0, 0.15)';
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(300, sepY);
    ctx.lineTo(W - 300, sepY);
    ctx.stroke();
    drawDiamond(ctx, W / 2, sepY, 6, 'rgba(255, 215, 0, 0.4)');

    // ========= CERTIFICATE OF TITLE =========
    ctx.fillStyle = 'rgba(255, 215, 0, 0.9)';
    ctx.font = '700 42px "Cinzel", serif';
    ctx.fillText('CERTIFICATE', W / 2, 340);

    ctx.fillStyle = 'rgba(240, 240, 245, 0.35)';
    ctx.font = '400 16px "Inter", sans-serif';
    ctx.fillText('of Participation', W / 2, 370);

    // ========= PHOTO SECTION =========
    const photoSize = 160;
    const photoCenterX = W / 2;
    const photoCenterY = 480;

    // Photo glow ring
    const photoGlow = ctx.createRadialGradient(photoCenterX, photoCenterY, photoSize / 2 - 10, photoCenterX, photoCenterY, photoSize / 2 + 20);
    photoGlow.addColorStop(0, 'transparent');
    photoGlow.addColorStop(0.8, 'rgba(255, 215, 0, 0.08)');
    photoGlow.addColorStop(1, 'transparent');
    ctx.fillStyle = photoGlow;
    ctx.fillRect(photoCenterX - photoSize, photoCenterY - photoSize, photoSize * 2, photoSize * 2);

    // Photo border ring
    ctx.beginPath();
    ctx.arc(photoCenterX, photoCenterY, photoSize / 2 + 4, 0, Math.PI * 2);
    ctx.strokeStyle = 'rgba(255, 215, 0, 0.4)';
    ctx.lineWidth = 2;
    ctx.stroke();

    ctx.beginPath();
    ctx.arc(photoCenterX, photoCenterY, photoSize / 2 + 8, 0, Math.PI * 2);
    ctx.strokeStyle = 'rgba(255, 215, 0, 0.15)';
    ctx.lineWidth = 1;
    ctx.stroke();

    // Clip and draw photo
    const photoSrc = getPhotoForParticipant(participant, index);
    if (photoSrc) {
        const img = await loadImage(photoSrc);
        ctx.save();
        ctx.beginPath();
        ctx.arc(photoCenterX, photoCenterY, photoSize / 2, 0, Math.PI * 2);
        ctx.clip();

        // Cover-fit image into circle
        const minDim = Math.min(img.width, img.height);
        const sx = (img.width - minDim) / 2;
        const sy = (img.height - minDim) / 2;
        ctx.drawImage(img, sx, sy, minDim, minDim,
            photoCenterX - photoSize / 2, photoCenterY - photoSize / 2, photoSize, photoSize);
        ctx.restore();
    } else {
        // Placeholder circle
        ctx.beginPath();
        ctx.arc(photoCenterX, photoCenterY, photoSize / 2, 0, Math.PI * 2);
        ctx.fillStyle = 'rgba(255, 215, 0, 0.06)';
        ctx.fill();
        ctx.fillStyle = 'rgba(255, 215, 0, 0.5)';
        ctx.font = 'bold 52px "Cinzel", serif';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.fillText(getInitials(participant.name), photoCenterX, photoCenterY);
        ctx.textBaseline = 'alphabetic';
    }

    // ========= PRESENTED TO =========
    ctx.textAlign = 'center';
    ctx.fillStyle = 'rgba(240, 240, 245, 0.4)';
    ctx.font = '400 15px "Inter", sans-serif';
    ctx.fillText('This certificate is proudly presented to', W / 2, 590);

    // ========= PARTICIPANT NAME =========
    ctx.fillStyle = '#ffd700';
    ctx.font = '700 38px "Cinzel", serif';
    const nameText = participant.name.toUpperCase();
    ctx.fillText(nameText, W / 2, 640);

    // Name underline
    const nameWidth = ctx.measureText(nameText).width;
    const nameGrad = ctx.createLinearGradient(W / 2 - nameWidth / 2, 0, W / 2 + nameWidth / 2, 0);
    nameGrad.addColorStop(0, 'transparent');
    nameGrad.addColorStop(0.2, 'rgba(255, 215, 0, 0.4)');
    nameGrad.addColorStop(0.8, 'rgba(255, 215, 0, 0.4)');
    nameGrad.addColorStop(1, 'transparent');
    ctx.strokeStyle = nameGrad;
    ctx.lineWidth = 1.5;
    ctx.beginPath();
    ctx.moveTo(W / 2 - nameWidth / 2 - 20, 650);
    ctx.lineTo(W / 2 + nameWidth / 2 + 20, 650);
    ctx.stroke();

    // ========= COLLEGE =========
    ctx.fillStyle = 'rgba(240, 240, 245, 0.5)';
    ctx.font = '500 16px "Inter", sans-serif';
    ctx.fillText(participant.college, W / 2, 680);

    // ========= FOR PRESENTING PAPER =========
    ctx.fillStyle = 'rgba(240, 240, 245, 0.35)';
    ctx.font = '400 14px "Inter", sans-serif';
    ctx.fillText('for presenting the paper entitled', W / 2, 720);

    // ========= PAPER TITLE =========
    ctx.fillStyle = 'rgba(255, 215, 0, 0.7)';
    ctx.font = 'italic 600 20px "Inter", sans-serif';

    // Word wrap if too long
    const maxWidth = W - 300;
    const paperLines = wrapText(ctx, `"${participant.paperTitle}"`, maxWidth);
    let paperY = 755;
    paperLines.forEach((line) => {
        ctx.fillText(line, W / 2, paperY);
        paperY += 28;
    });

    // ========= AT THE CONFERENCE =========
    const confY = paperY + 15;
    ctx.fillStyle = 'rgba(240, 240, 245, 0.3)';
    ctx.font = '400 13px "Inter", sans-serif';
    ctx.fillText('at the National Conference RTICCE-2K26, held on April 17, 2026', W / 2, confY);
    ctx.fillText('at P.T. Lee Chengalvaraya Naicker College of Engineering & Technology, Kanchipuram', W / 2, confY + 20);

    // ========= BOTTOM DECORATIVE LINE =========
    const botLineY = H - 200;
    ctx.strokeStyle = 'rgba(255, 215, 0, 0.15)';
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(100, botLineY);
    ctx.lineTo(W - 100, botLineY);
    ctx.stroke();

    // ========= SIGNATURES =========
    const sigY = H - 140;

    // Left signature - HOD
    ctx.fillStyle = 'rgba(255, 215, 0, 0.5)';
    ctx.font = '600 15px "Inter", sans-serif';
    ctx.fillText('Dr. A. Karthikayen', 300, sigY);
    ctx.fillStyle = 'rgba(240, 240, 245, 0.3)';
    ctx.font = '400 12px "Inter", sans-serif';
    ctx.fillText('HOD / ECE', 300, sigY + 20);

    // Signature line left
    ctx.strokeStyle = 'rgba(255, 215, 0, 0.25)';
    ctx.beginPath();
    ctx.moveTo(200, sigY - 15);
    ctx.lineTo(400, sigY - 15);
    ctx.stroke();

    // Center - Principal
    ctx.fillStyle = 'rgba(255, 215, 0, 0.5)';
    ctx.font = '600 15px "Inter", sans-serif';
    ctx.fillText('Dr. P. Palanisamy', W / 2, sigY);
    ctx.fillStyle = 'rgba(240, 240, 245, 0.3)';
    ctx.font = '400 12px "Inter", sans-serif';
    ctx.fillText('Principal', W / 2, sigY + 20);

    ctx.strokeStyle = 'rgba(255, 215, 0, 0.25)';
    ctx.beginPath();
    ctx.moveTo(W / 2 - 100, sigY - 15);
    ctx.lineTo(W / 2 + 100, sigY - 15);
    ctx.stroke();

    // Right - Convenor
    ctx.fillStyle = 'rgba(255, 215, 0, 0.5)';
    ctx.font = '600 15px "Inter", sans-serif';
    ctx.fillText('Dr. S. Parasuraman', W - 300, sigY);
    ctx.fillStyle = 'rgba(240, 240, 245, 0.3)';
    ctx.font = '400 12px "Inter", sans-serif';
    ctx.fillText('Co-Convenor', W - 300, sigY + 20);

    ctx.strokeStyle = 'rgba(255, 215, 0, 0.25)';
    ctx.beginPath();
    ctx.moveTo(W - 400, sigY - 15);
    ctx.lineTo(W - 200, sigY - 15);
    ctx.stroke();

    // ========= BOTTOM CORNER ITEMS =========
    ctx.fillStyle = 'rgba(240, 240, 245, 0.15)';
    ctx.font = '400 10px "Inter", sans-serif';
    ctx.textAlign = 'left';
    ctx.fillText(`Cert. No: RTICCE-2K26/${String(index + 1).padStart(3, '0')}`, 65, H - 65);
    ctx.textAlign = 'right';
    ctx.fillText('April 17, 2026 • Kanchipuram, Tamil Nadu', W - 65, H - 65);
    ctx.textAlign = 'center';

    return canvas.toDataURL('image/png', 1.0);
}

// ===== CANVAS HELPERS =====
function roundRect(ctx, x, y, w, h, r) {
    ctx.beginPath();
    ctx.moveTo(x + r, y);
    ctx.lineTo(x + w - r, y);
    ctx.arcTo(x + w, y, x + w, y + r, r);
    ctx.lineTo(x + w, y + h - r);
    ctx.arcTo(x + w, y + h, x + w - r, y + h, r);
    ctx.lineTo(x + r, y + h);
    ctx.arcTo(x, y + h, x, y + h - r, r);
    ctx.lineTo(x, y + r);
    ctx.arcTo(x, y, x + r, y, r);
    ctx.closePath();
}

function drawCornerDecor(ctx, x, y, dx, dy) {
    const len = 40;
    ctx.strokeStyle = 'rgba(255, 215, 0, 0.5)';
    ctx.lineWidth = 2.5;
    ctx.lineCap = 'round';

    ctx.beginPath();
    ctx.moveTo(x, y + dy * len);
    ctx.lineTo(x, y);
    ctx.lineTo(x + dx * len, y);
    ctx.stroke();

    // Inner shorter accent
    ctx.strokeStyle = 'rgba(255, 215, 0, 0.25)';
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(x + dx * 6, y + dy * (len - 8));
    ctx.lineTo(x + dx * 6, y + dy * 6);
    ctx.lineTo(x + dx * (len - 8), y + dy * 6);
    ctx.stroke();
}

function drawDiamond(ctx, x, y, size, color) {
    ctx.fillStyle = color;
    ctx.beginPath();
    ctx.moveTo(x, y - size);
    ctx.lineTo(x + size, y);
    ctx.lineTo(x, y + size);
    ctx.lineTo(x - size, y);
    ctx.closePath();
    ctx.fill();
}

function wrapText(ctx, text, maxWidth) {
    const words = text.split(' ');
    const lines = [];
    let currentLine = '';

    for (const word of words) {
        const testLine = currentLine ? currentLine + ' ' + word : word;
        if (ctx.measureText(testLine).width > maxWidth) {
            if (currentLine) lines.push(currentLine);
            currentLine = word;
        } else {
            currentLine = testLine;
        }
    }
    if (currentLine) lines.push(currentLine);
    return lines;
}

function loadImage(src) {
    return new Promise((resolve, reject) => {
        const img = new Image();
        img.crossOrigin = 'anonymous';
        img.onload = () => resolve(img);
        img.onerror = reject;
        img.src = src;
    });
}

// ===== GENERATE CERTIFICATES =====
async function generateAllCertificates() {
    if (state.participants.length === 0) {
        showToast('No participants loaded!');
        return;
    }

    showProgress('Generating Certificates...');
    state.certificates = [];
    els.certGrid.innerHTML = '';

    for (let i = 0; i < state.participants.length; i++) {
        const p = state.participants[i];
        const dataURL = await renderCertificate(p, i);
        state.certificates.push({ name: p.name, college: p.college, dataURL });
        updateProgress(i + 1, state.participants.length);

        // Yield to UI
        await new Promise(r => setTimeout(r, 50));
    }

    hideProgress();
    renderCertificateGrid();
    showToast(`${state.certificates.length} certificates generated!`);
}

function renderCertificateGrid() {
    els.previewSection.classList.remove('hidden');
    els.previewSection.classList.add('fade-in-up');

    els.certGrid.innerHTML = state.certificates.map((cert, i) => `
        <div class="cert-card fade-in-up" style="animation-delay: ${i * 80}ms">
            <img class="cert-card-img" src="${cert.dataURL}" alt="Certificate - ${cert.name}" onclick="openPreview(${i})">
            <div class="cert-card-info">
                <div>
                    <div class="cert-card-name">${cert.name}</div>
                    <div class="cert-card-college">${cert.college}</div>
                </div>
                <button class="btn-icon" onclick="downloadCert(${i})" title="Download">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="16" height="16">
                        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                        <polyline points="7 10 12 15 17 10"/>
                        <line x1="12" y1="15" x2="12" y2="3"/>
                    </svg>
                </button>
            </div>
        </div>
    `).join('');

    // Scroll to preview
    els.previewSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// ===== PREVIEW & DOWNLOAD =====
let currentPreviewIndex = -1;

function openPreview(index) {
    currentPreviewIndex = index;
    els.modalImg.src = state.certificates[index].dataURL;
    els.modal.classList.remove('hidden');
}

function closePreview() {
    els.modal.classList.add('hidden');
    currentPreviewIndex = -1;
}

function downloadCert(index) {
    const cert = state.certificates[index];
    const link = document.createElement('a');
    link.download = `Certificate_${cert.name.replace(/\s+/g, '_')}.png`;
    link.href = cert.dataURL;
    link.click();
    showToast(`Downloaded certificate for ${cert.name}`);
}

async function previewSingle(index) {
    showProgress('Generating Preview...');
    const p = state.participants[index];
    const dataURL = await renderCertificate(p, index);
    hideProgress();

    // Show in modal
    currentPreviewIndex = -1;
    els.modalImg.src = dataURL;
    els.modal.classList.remove('hidden');

    // Store temporarily for download
    els.modalDownload.onclick = () => {
        const link = document.createElement('a');
        link.download = `Certificate_${p.name.replace(/\s+/g, '_')}.png`;
        link.href = dataURL;
        link.click();
        showToast(`Downloaded certificate for ${p.name}`);
    };
}

async function downloadSingle(index) {
    showProgress('Generating Certificate...');
    const p = state.participants[index];
    const dataURL = await renderCertificate(p, index);
    hideProgress();

    const link = document.createElement('a');
    link.download = `Certificate_${p.name.replace(/\s+/g, '_')}.png`;
    link.href = dataURL;
    link.click();
    showToast(`Downloaded certificate for ${p.name}`);
}

async function downloadAllAsZip() {
    if (state.certificates.length === 0) {
        showToast('Generate certificates first!');
        return;
    }

    showProgress('Creating ZIP file...');
    const zip = new JSZip();

    state.certificates.forEach((cert, i) => {
        const base64Data = cert.dataURL.split(',')[1];
        const fileName = `Certificate_${String(i + 1).padStart(3, '0')}_${cert.name.replace(/\s+/g, '_')}.png`;
        zip.file(fileName, base64Data, { base64: true });
        updateProgress(i + 1, state.certificates.length);
    });

    const blob = await zip.generateAsync({ type: 'blob' });
    saveAs(blob, 'RTICCE-2K26_Certificates.zip');
    hideProgress();
    showToast('All certificates downloaded as ZIP!');
}

// ===== DEMO MODE =====
function loadDemoData() {
    state.participants = DEMO_PARTICIPANTS.map((p, i) => ({
        ...p,
    }));

    // Generate demo avatars as photos
    DEMO_PARTICIPANTS.forEach((p, i) => {
        state.photos['demo' + (i + 1)] = generateDemoAvatar(p.name, i);
    });

    els.excelStatus.textContent = '✓ Demo data loaded (6 sample participants)';
    els.excelDrop.classList.add('uploaded');
    els.photosStatus.textContent = '✓ Demo avatars generated';
    els.photosDrop.classList.add('uploaded');

    renderDataTable();
    showToast('Demo data loaded! Click "Generate All Certificates" to see the magic ✨');
}

// ===== EVENT LISTENERS =====
function init() {
    // Drag & drop zones
    setupDragDrop(els.excelDrop, els.excelInput);
    setupDragDrop(els.photosDrop, els.photosInput);

    // File inputs
    els.excelInput.addEventListener('change', (e) => handleExcel(e.target.files[0]));
    els.photosInput.addEventListener('change', (e) => handlePhotos(e.target.files));

    // Buttons
    els.btnDemo.addEventListener('click', loadDemoData);
    els.btnGenerate.addEventListener('click', generateAllCertificates);
    els.btnDownloadAll.addEventListener('click', downloadAllAsZip);
    els.modalClose.addEventListener('click', closePreview);
    els.modalDownload.addEventListener('click', () => {
        if (currentPreviewIndex >= 0) downloadCert(currentPreviewIndex);
    });

    // Close modal on overlay click
    els.modal.addEventListener('click', (e) => {
        if (e.target === els.modal) closePreview();
    });

    // Keyboard
    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') closePreview();
    });
}

// Make functions globally available
window.previewSingle = previewSingle;
window.downloadSingle = downloadSingle;
window.downloadCert = downloadCert;
window.openPreview = openPreview;

// Go!
init();
