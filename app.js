/**
 * Maha AI Olympiad - Attendance Register to Student Template Scanner
 * Handles AI Vision OCR, Marathi Transliteration, Interactive Review, and Excel Export.
 */

// Application State
const state = {
    apiKey: localStorage.getItem('gemini_api_key') || '',
    model: localStorage.getItem('gemini_model') || '',  // Will be auto-discovered
    _discoveredModels: null,  // Cache for discovered models
    _abortController: null,   // For cancelling in-flight requests
    defaultSection: 'A',
    nameFormat: 'clean', // 'clean' (Vedika Nehare) | 'with_father' (Vedika Rajkumar Nehare)
    rollNumberFormat: 'sequential_01', // 'sequential_01' | 'sequential_1' | 'adm_no'
    excludeStruckOut: true,
    
    images: [], // Image and rendered PDF pages: { id, name, url, base64, classHint, classOverride, scanned }
    activeImageId: null,
    
    students: [], // Extracted or demo students
    filterClass: 'all',
    searchQuery: '',
    isProcessing: false,
    step: 'upload',
    scanRequested: false,
    
    // Zoom state
    zoomScale: 1,
    zoomRotate: 0,
    panX: 0,
    panY: 0,
    isPanning: false,
    startX: 0,
    startY: 0
};

// DOM Elements
const elements = {
    dropzone: document.getElementById('dropzone'),
    fileInput: document.getElementById('file-input'),
    cameraInput: document.getElementById('camera-input'),
    uploadPrompt: document.getElementById('upload-prompt'),
    workspace: document.getElementById('workspace'),
    
    // Image viewer
    viewerImage: document.getElementById('viewer-image'),
    viewerContainer: document.getElementById('viewer-container'),
    imageSelector: document.getElementById('image-selector'),
    zoomInBtn: document.getElementById('zoom-in'),
    zoomOutBtn: document.getElementById('zoom-out'),
    zoomResetBtn: document.getElementById('zoom-reset'),
    rotateBtn: document.getElementById('zoom-rotate'),
    
    // Roster view
    rosterTableBody: document.getElementById('roster-tbody'),
    studentCountBadge: document.getElementById('student-count-badge'),
    classFilterContainer: document.getElementById('class-filter-container'),
    searchInput: document.getElementById('search-input'),
    nameFormatToggle: document.getElementById('name-format-toggle'),
    rollFormatSelect: document.getElementById('roll-format-select'),
    defaultSectionInput: document.getElementById('default-section-input'),
    
    // Action Buttons
    exportXlsxBtn: document.getElementById('export-xlsx-btn'),
    exportCsvBtn: document.getElementById('export-csv-btn'),
    copyClipboardBtn: document.getElementById('copy-clipboard-btn'),
    addStudentBtn: document.getElementById('add-student-btn'),
    loadDemoBtn: document.getElementById('load-demo-btn'),
    settingsBtn: document.getElementById('settings-btn'),
    
    // Modals
    settingsModal: document.getElementById('settings-modal'),
    saveSettingsBtn: document.getElementById('save-settings-btn'),
    closeSettingsBtn: document.getElementById('close-settings-btn'),
    apiKeyInput: document.getElementById('api-key-input'),
    modelSelect: document.getElementById('model-select'),
    processingOverlay: document.getElementById('processing-overlay'),
    processingText: document.getElementById('processing-text')
};

// Initialize Application
function init() {
    setupEventListeners();
    setupDropzone();
    setupImageViewer();
    
    // Sync settings in inputs
    elements.apiKeyInput.value = state.apiKey;
    elements.defaultSectionInput.value = state.defaultSection;
    
    // If no data loaded yet, show prompt
    updateViewMode();
    
    // Auto-discover available models on startup if API key exists
    if (state.apiKey) {
        discoverModels(state.apiKey).then(() => {
            syncModelDropdown();
        });
    }
}

// ==========================================
// MODEL DISCOVERY (the key fix)
// ==========================================

/**
 * Queries the Gemini API to discover which models are actually available
 * for this API key. This eliminates ALL hardcoded model name guessing.
 */
async function discoverModels(apiKey) {
    try {
        const res = await fetch(`https://generativelanguage.googleapis.com/v1beta/models?key=${apiKey}`);
        if (!res.ok) {
            console.warn('Model discovery failed:', res.status);
            return [];
        }
        const data = await res.json();
        const models = (data.models || [])
            .filter(m => m.supportedGenerationMethods && m.supportedGenerationMethods.includes('generateContent'))
            .map(m => ({
                id: m.name.replace(/^models\//, ''),
                displayName: m.displayName || m.name.replace(/^models\//, ''),
                description: m.description || ''
            }));
        
        // Sort: prefer FAST flash models for OCR (not heavy thinking models)
        // Priority: 2.5-flash > 3.5-flash-lite > flash-lite > flash > pro
        models.sort((a, b) => {
            // Score each model: lower = faster = better for OCR
            function speed(m) {
                const id = m.id;
                if (id.includes('flash-lite')) return 1;      // Ultra-fast
                if (id === 'gemini-2.5-flash') return 2;      // Fast & proven
                if (id.includes('2.5-flash')) return 3;       // 2.5 flash variants
                if (id.includes('flash') && !id.includes('3.7')) return 4; // Other flash
                if (id.includes('3.7-flash')) return 8;       // Slow thinking model
                if (id.includes('pro')) return 9;              // Slowest
                return 5;
            }
            return speed(a) - speed(b);
        });
        
        state._discoveredModels = models;
        
        // Auto-select best model if current one is empty or not in the list
        if (models.length > 0) {
            const currentValid = models.find(m => m.id === state.model);
            if (!currentValid) {
                state.model = models[0].id;
                localStorage.setItem('gemini_model', state.model);
            }
        }
        
        console.log('Discovered models:', models.map(m => m.id));
        return models;
    } catch (err) {
        console.warn('Model discovery error:', err);
        return [];
    }
}

/**
 * Populate the model <select> dropdown with actually available models
 */
function syncModelDropdown() {
    const select = elements.modelSelect;
    if (!select || !state._discoveredModels || state._discoveredModels.length === 0) return;
    
    select.innerHTML = '';
    state._discoveredModels.forEach(m => {
        const opt = document.createElement('option');
        opt.value = m.id;
        opt.textContent = m.displayName || m.id;
        if (m.id === state.model) opt.selected = true;
        select.appendChild(opt);
    });
}

// Event Listeners
function setupEventListeners() {
    document.querySelectorAll('[data-step]').forEach(button => button.addEventListener('click', () => {
        state.step = button.dataset.step;
        updateViewMode();
        window.scrollTo({ top: 0 });
    }));
    document.getElementById('choose-files-btn').addEventListener('click', () => elements.fileInput.click());
    document.getElementById('take-photo-btn').addEventListener('click', () => elements.cameraInput.click());
    ['back-review-btn', 'return-review-btn'].forEach(id => document.getElementById(id).addEventListener('click', () => {
        state.step = 'review';
        updateViewMode();
        window.scrollTo({ top: 0 });
    }));
    document.getElementById('go-download-btn').addEventListener('click', () => {
        state.step = 'download';
        updateViewMode();
        window.scrollTo({ top: 0 });
    });
    document.getElementById('scan-pending-btn').addEventListener('click', () => {
        if (!state.apiKey) {
            state.scanRequested = true;
            elements.settingsModal.classList.remove('hidden');
        }
        else processImagesWithAI();
    });
    // Demo button
    elements.loadDemoBtn.addEventListener('click', loadDemoData);

    // Upload buttons (Header and Viewer toolbar)
    const headerUploadBtn = document.getElementById('header-upload-btn');
    if (headerUploadBtn) {
        headerUploadBtn.addEventListener('click', () => {
            elements.fileInput.click();
        });
    }

    const addPageBtn = document.getElementById('add-page-btn');
    if (addPageBtn) {
        addPageBtn.addEventListener('click', () => {
            elements.fileInput.click();
        });
    }

    // Reset workspace button
    const resetWorkspaceBtn = document.getElementById('reset-workspace-btn');
    if (resetWorkspaceBtn) {
        resetWorkspaceBtn.addEventListener('click', () => {
            if (confirm('Clear current data and upload new register images or PDFs?')) {
                state.students = [];
                state.images = [];
                state.activeImageId = null;
                state.step = 'upload';
                state.filterClass = 'all';
                state.searchQuery = '';
                elements.searchInput.value = '';
                state.scanRequested = false;
                document.querySelector('.app-menu').open = false;
                elements.viewerImage.src = '';
                elements.imageSelector.innerHTML = '';
                updateViewMode();
                showToast('Workspace reset. Drop new images or PDFs to begin!', 'info');
            }
        });
    }

    // Cancel scan button on overlay
    const cancelProcessingBtn = document.getElementById('cancel-processing-btn');
    if (cancelProcessingBtn) {
        cancelProcessingBtn.addEventListener('click', () => {
            if (state._abortController) {
                state._abortController.abort();
                state._abortController = null;
            }
            showToast('Scan cancelled.', 'info');
        });
    }

    // Global drag-and-drop onto workspace
    window.addEventListener('dragover', (e) => {
        e.preventDefault();
    });
    window.addEventListener('drop', (e) => {
        e.preventDefault();
        const files = e.dataTransfer?.files;
        if (files && files.length > 0) {
            handleFiles(Array.from(files));
        }
    });
    
    // File inputs
    elements.fileInput.addEventListener('change', handleFileSelect);
    elements.cameraInput.addEventListener('change', handleFileSelect);
    
    // Image dropdown selector
    elements.imageSelector.addEventListener('change', (e) => {
        setActiveImage(e.target.value);
    });
    
    // Settings modal
    elements.settingsBtn.addEventListener('click', () => {
        elements.apiKeyInput.value = state.apiKey;
        syncModelDropdown();
        elements.settingsModal.classList.remove('hidden');
    });
    
    elements.closeSettingsBtn.addEventListener('click', () => {
        state.scanRequested = false;
        elements.settingsModal.classList.add('hidden');
    });

    const cancelBtn = document.getElementById('cancel-settings-btn');
    if (cancelBtn) {
        cancelBtn.addEventListener('click', () => {
            state.scanRequested = false;
            elements.settingsModal.classList.add('hidden');
        });
    }

    const modalTryDemoBtn = document.getElementById('modal-try-demo-btn');
    if (modalTryDemoBtn) {
        modalTryDemoBtn.addEventListener('click', () => {
            elements.settingsModal.classList.add('hidden');
            loadDemoData();
        });
    }

    // Close modal on click outside
    elements.settingsModal.addEventListener('click', (e) => {
        if (e.target === elements.settingsModal) {
            state.scanRequested = false;
            elements.settingsModal.classList.add('hidden');
        }
    });

    // Close on Escape key
    window.addEventListener('keydown', (e) => {
        if (e.key === 'Escape' && !elements.settingsModal.classList.contains('hidden')) {
            state.scanRequested = false;
            elements.settingsModal.classList.add('hidden');
        }
    });
    
    elements.saveSettingsBtn.addEventListener('click', async () => {
        const newKey = elements.apiKeyInput.value.trim();
        const newModel = elements.modelSelect.value;
        
        state.apiKey = newKey;
        state.model = newModel;
        localStorage.setItem('gemini_api_key', newKey);
        localStorage.setItem('gemini_model', newModel);
        elements.settingsModal.classList.add('hidden');
        showToast('Settings saved!', 'success');
        
        // Re-discover models whenever API key changes
        if (newKey) {
            await discoverModels(newKey);
            syncModelDropdown();
            if (state.scanRequested && !state.isProcessing) {
                state.scanRequested = false;
                await processImagesWithAI();
            }
        }
    });
    
    // Search input
    elements.searchInput.addEventListener('input', (e) => {
        state.searchQuery = e.target.value.toLowerCase().trim();
        renderRosterTable();
    });
    
    // Formatting Toggles
    elements.nameFormatToggle.addEventListener('change', (e) => {
        state.nameFormat = e.target.checked ? 'with_father' : 'clean';
        renderRosterTable();
    });
    
    elements.rollFormatSelect.addEventListener('change', (e) => {
        state.rollNumberFormat = e.target.value;
        recalculateRollNumbers();
        renderRosterTable();
    });
    
    elements.defaultSectionInput.addEventListener('input', (e) => {
        const val = e.target.value.toUpperCase().trim() || 'A';
        state.defaultSection = val;
        state.students.forEach(s => s.section = val);
        renderRosterTable();
    });

    // Bulk Class Change (Toolbar)
    const bulkClassSelect = document.getElementById('bulk-class-select');
    if (bulkClassSelect) {
        bulkClassSelect.addEventListener('change', (e) => {
            applyBulkClass(e.target.value);
        });
    }

    // Quick Class Change (Table Header)
    const thClassSelect = document.getElementById('th-class-select');
    if (thClassSelect) {
        thClassSelect.addEventListener('change', (e) => {
            applyBulkClass(e.target.value);
        });
    }

    // Page Class Quick-Bar (Image Viewer)
    const applyPageClassBtn = document.getElementById('apply-page-class-btn');
    const pageClassSelect = document.getElementById('page-class-select');
    if (applyPageClassBtn && pageClassSelect) {
        applyPageClassBtn.addEventListener('click', () => {
            applyClassToActivePage(pageClassSelect.value);
        });
    }
    
    // Auto-fill DOB buttons
    const toolbarAutoDobBtn = document.getElementById('toolbar-auto-dob-btn');
    if (toolbarAutoDobBtn) {
        toolbarAutoDobBtn.addEventListener('click', () => fillMissingDobs(false));
    }
    const thFillDobBtn = document.getElementById('th-fill-dob-btn');
    if (thFillDobBtn) {
        thFillDobBtn.addEventListener('click', () => fillMissingDobs(false));
    }
    
    // Add student button
    elements.addStudentBtn.addEventListener('click', () => {
        const targetClass = state.filterClass !== 'all' ? state.filterClass : '6';
        const newStudent = {
            id: 'manual_' + Date.now(),
            firstName: 'New',
            fatherName: '',
            lastName: 'Student',
            marathiName: 'नवीन विद्यार्थी',
            dob: generateRandomDobForClass(targetClass),
            studentClass: targetClass,
            section: state.defaultSection,
            rollNumber: (state.students.length + 1).toString().padStart(2, '0'),
            admNo: '',
            category: 'OPEN',
            aadhaar: '',
            isStruckOut: false
        };
        state.students.push(newStudent);
        renderRosterTable();
        showToast('New student added to the bottom of the list.', 'info');
    });
    
    // Exports
    elements.exportXlsxBtn.addEventListener('click', exportToExcel);
    elements.exportCsvBtn.addEventListener('click', exportToCsv);
    elements.copyClipboardBtn.addEventListener('click', copyToClipboard);
}

// Dropzone setup
function setupDropzone() {
    const dropzone = elements.dropzone;
    
    ['dragenter', 'dragover'].forEach(eventName => {
        dropzone.addEventListener(eventName, (e) => {
            e.preventDefault();
            e.stopPropagation();
            dropzone.classList.add('dragover');
        }, false);
    });
    
    ['dragleave', 'drop'].forEach(eventName => {
        dropzone.addEventListener(eventName, (e) => {
            e.preventDefault();
            e.stopPropagation();
            dropzone.classList.remove('dragover');
        }, false);
    });
    
    dropzone.addEventListener('drop', (e) => {
        const files = e.dataTransfer.files;
        if (files && files.length > 0) {
            handleFiles(Array.from(files));
        }
    });
}

// Handle file selection from input
function handleFileSelect(e) {
    if (e.target.files && e.target.files.length > 0) {
        handleFiles(Array.from(e.target.files));
    }
}

// Normalize Marathi numerals before interpreting class labels.
function normalizeDigits(value) {
    return String(value ?? '').replace(/[०-९]/g, digit => '०१२३४५६७८९'.indexOf(digit));
}

function normalizeClass(value) {
    const text = normalizeDigits(value).trim();
    const roman = text.match(/^(?:class|std\.?|standard|grade)\s+(XII|XI|IX|VIII|VII|VI|IV|III|II|X|V|I)$/i);
    if (roman) return String({ I: 1, II: 2, III: 3, IV: 4, V: 5, VI: 6, VII: 7, VIII: 8, IX: 9, X: 10, XI: 11, XII: 12 }[roman[1].toUpperCase()]);
    const match = text.match(/^(?:(?:class|std\.?|standard|grade|वर्ग|इयत्ता)\s*[:._-]?\s*)?(1[0-2]|[1-9])(?:\s*(?:st|nd|rd|th|वी|वि|वा|री|रा|ली|ला))?$/i);
    if (match) return match[1];
    const words = { पहिली: '1', दुसरी: '2', तिसरी: '3', चौथी: '4', पाचवी: '5', सहावी: '6', सातवी: '7', आठवी: '8', नववी: '9', दहावी: '10', अकरावी: '11', बारावी: '12' };
    return words[text.replace(/^(?:वर्ग|इयत्ता)\s*/, '')] || '';
}

function detectClassFromFilename(name) {
    // Require a class label: dates, image counters and phone numbers are not grades.
    const text = normalizeDigits(name);
    const matches = [...text.matchAll(/(?:^|[\s_.-])(?:class|standard|std|grade|वर्ग|इयत्ता)[\s_.-]*(1[0-2]|[1-9])(?:st|nd|rd|th|वी|वि|वा)?(?=$|[\s_.-])/gi)];
    if (matches.length !== 1) return null;
    // A filename describing multiple classes cannot supply a single page class.
    if (/(?:[&_+]|\band\b|\bto\b|-)\s*[०-९0-9]/i.test(text.slice(matches[0].index + matches[0][0].length))) return null;
    return matches[0][1];
}

function resolveStudentClass(item, page) {
    return page.classOverride || normalizeClass(item.classHeader) || page.classHint || normalizeClass(item.studentClass);
}

// Bulk Class Change Functions
function applyClassToActivePage(newClass) {
    if (!newClass) return;
    const activeImg = state.images.find(i => i.id === state.activeImageId);
    if (activeImg) activeImg.classOverride = newClass.toString();
    let count = 0;
    
    state.students.forEach(s => {
        const matchesImage = activeImg && (s.imageId === activeImg.id || s.imageSource === activeImg.name);
        if (state.images.length <= 1 || matchesImage) {
            s.studentClass = newClass.toString();
            count++;
        }
    });
    
    
    recalculateRollNumbers();
    renderClassFilters();
    renderRosterTable();
    showToast(`Updated ${count} student(s) on this page to Class ${newClass}!`, 'success');
}

function applyBulkClass(newClass) {
    if (!newClass) return;
    let count = 0;
    
    state.students.forEach(s => {
        if (state.filterClass === 'all' || s.studentClass === state.filterClass) {
            s.studentClass = newClass.toString();
            count++;
        }
    });
    
    if (state.filterClass !== 'all') {
        state.filterClass = newClass.toString();
    }
    
    recalculateRollNumbers();
    renderClassFilters();
    renderRosterTable();
    showToast(`Changed Class to ${newClass} for ${count} student(s)!`, 'success');
    
    const bulkSelect = document.getElementById('bulk-class-select');
    if (bulkSelect) bulkSelect.value = '';
    const thSelect = document.getElementById('th-class-select');
    if (thSelect) thSelect.value = '';
}

// Generate realistic random Date of Birth based on student class
function generateRandomDobForClass(studentClass) {
    const cls = parseInt(studentClass, 10) || 6;
    // Current academic year 2026 -> Class 1 is ~6 yrs old (2020), Class 6 is ~11-12 yrs old (2014-2015)
    const baseYear = 2026 - 5 - cls;
    const year = Math.random() > 0.5 ? baseYear : baseYear - 1;
    const month = Math.floor(Math.random() * 12) + 1;
    const maxDays = new Date(year, month, 0).getDate();
    const day = Math.floor(Math.random() * maxDays) + 1;
    
    const mm = month.toString().padStart(2, '0');
    const dd = day.toString().padStart(2, '0');
    return `${year}-${mm}-${dd}`;
}

// Auto-fill missing DOBs across all or filtered students
function fillMissingDobs(forceAll = false) {
    if (state.students.length === 0) {
        showToast('No students loaded to fill birthdates.', 'warning');
        return;
    }
    
    let filledCount = 0;
    state.students.forEach(s => {
        const isMissing = !s.dob || s.dob === 'dd/mm/yyyy' || s.dob.trim() === '' || s.dob.length < 8;
        if (forceAll || isMissing) {
            s.dob = generateRandomDobForClass(s.studentClass || '6');
            filledCount++;
        }
    });
    
    renderRosterTable();
    if (filledCount > 0) {
        showToast(`Generated realistic random birthdates for ${filledCount} student(s)!`, 'success');
    } else {
        showToast('All students already have valid birthdates!', 'info');
    }
}

// Process uploaded files
async function handleFiles(files) {
    if (state.isProcessing) {
        showToast('Please wait for the current import to finish.', 'warning');
        return;
    }
    const accepted = files.filter(f => f.type.startsWith('image/') || isPdf(f));
    if (accepted.length !== files.length) showToast('Only images and PDF documents are supported.', 'warning');
    if (!accepted.length) return;
    state.step = 'upload';
    const controller = new AbortController();
    state._abortController = controller;
    showProcessing(true, 'Preparing register pages...');
    try {
        for (const file of accepted) {
            controller.signal.throwIfAborted();
            try {
                const pages = isPdf(file)
                    ? await pdfToPages(file, controller.signal)
                    : [makePage(file.name, await fileToBase64(file), detectClassFromFilename(file.name))];
                controller.signal.throwIfAborted();
                state.images.push(...pages);
            } catch (err) {
                if (controller.signal.aborted) throw err;
                showToast(`Could not open ${file.name}: ${err.message}`, 'error');
            }
        }
    } catch (err) {
        if (!controller.signal.aborted) showToast(err.message, 'error');
    } finally {
        state._abortController = null;
        elements.fileInput.value = '';
        elements.cameraInput.value = '';
        showProcessing(false);
        updateImageSelector();
        if (state.images.length) setActiveImage(state.images[state.images.length - 1].id);
        updateViewMode();
    }
    // Scanning starts only after the user has reviewed their uploaded pages.
}

function isPdf(file) {
    return file.type === 'application/pdf' || /\.pdf$/i.test(file.name);
}

function makePage(name, base64, classHint) {
    return { id: crypto.randomUUID(), name, url: base64, base64, classHint, scanned: false };
}

let pdfLibraryPromise;
async function loadPdfLibrary() {
    if (!pdfLibraryPromise) {
        pdfLibraryPromise = import('https://cdn.jsdelivr.net/npm/pdfjs-dist@4.10.38/build/pdf.min.mjs')
            .then(lib => {
                lib.GlobalWorkerOptions.workerSrc = 'https://cdn.jsdelivr.net/npm/pdfjs-dist@4.10.38/build/pdf.worker.min.mjs';
                return lib;
            }).catch(err => { pdfLibraryPromise = null; throw new Error('PDF reader could not load. Check your internet connection and retry.'); });
    }
    return pdfLibraryPromise;
}

async function pdfToPages(file, signal) {
    if (file.size > 50 * 1024 * 1024) throw new Error('PDFs must be smaller than 50 MB. Split this document and retry.');
    const lib = await loadPdfLibrary();
    signal.throwIfAborted();
    const loading = lib.getDocument({
        data: new Uint8Array(await file.arrayBuffer()), isEvalSupported: false,
        cMapUrl: 'https://cdn.jsdelivr.net/npm/pdfjs-dist@4.10.38/cmaps/', cMapPacked: true,
        standardFontDataUrl: 'https://cdn.jsdelivr.net/npm/pdfjs-dist@4.10.38/standard_fonts/'
    });
    const abort = () => { loading.destroy().catch(() => {}); };
    signal.addEventListener('abort', abort, { once: true });
    try {
        signal.throwIfAborted();
        const pdf = await loading.promise;
        if (pdf.numPages > 100) throw new Error('Please split PDFs longer than 100 pages into smaller documents.');
        const pages = [];
        for (let number = 1; number <= pdf.numPages; number++) {
            signal.throwIfAborted();
            showProcessing(true, `Preparing ${file.name}: page ${number} of ${pdf.numPages}...`);
            const page = await pdf.getPage(number);
            const original = page.getViewport({ scale: 1 });
            const viewport = page.getViewport({ scale: 2400 / Math.max(original.width, original.height) });
            const canvas = document.createElement('canvas');
            canvas.width = Math.ceil(viewport.width);
            canvas.height = Math.ceil(viewport.height);
            try {
                await page.render({ canvasContext: canvas.getContext('2d'), viewport, background: 'white' }).promise;
                signal.throwIfAborted();
                pages.push(makePage(`${file.name} — page ${number}`, canvas.toDataURL('image/jpeg', 0.92), detectClassFromFilename(file.name)));
            } finally {
                canvas.width = canvas.height = 0;
                page.cleanup();
            }
        }
        return pages;
    } catch (err) {
        if (signal.aborted) signal.throwIfAborted();
        if (err.name === 'PasswordException') throw new Error('This PDF is password protected. Upload an unlocked copy.');
        throw err;
    } finally {
        signal.removeEventListener('abort', abort);
        await loading.destroy();
    }
}

function fileToBase64(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            const img = new Image();
            img.onload = () => {
                // Preserve small Marathi characters and handwritten class headings.
                const maxDim = 2400;
                let w = img.width;
                let h = img.height;
                if (w > maxDim || h > maxDim) {
                    if (w > h) {
                        h = Math.round((h * maxDim) / w);
                        w = maxDim;
                    } else {
                        w = Math.round((w * maxDim) / h);
                        h = maxDim;
                    }
                }
                const canvas = document.createElement('canvas');
                canvas.width = w;
                canvas.height = h;
                const ctx = canvas.getContext('2d');
                ctx.fillStyle = '#fff';
                ctx.fillRect(0, 0, w, h);
                ctx.drawImage(img, 0, 0, w, h);
                // Use high-quality JPEG so small letters remain legible for OCR.
                const compressed = canvas.toDataURL('image/jpeg', 0.92);
                resolve(compressed);
            };
            img.onerror = () => reject(new Error('Image format could not be decoded. Try JPG or PNG.'));
            img.src = e.target.result;
        };
        reader.onerror = reject;
        reader.readAsDataURL(file);
    });
}

// Update Image Dropdown in viewer
function updateImageSelector() {
    elements.imageSelector.innerHTML = '';
    state.images.forEach((img, idx) => {
        const opt = document.createElement('option');
        opt.value = img.id;
        opt.textContent = `Page ${idx + 1}: ${img.name}`;
        elements.imageSelector.appendChild(opt);
    });
}

function setActiveImage(imageId) {
    state.activeImageId = imageId;
    elements.imageSelector.value = imageId;
    const img = state.images.find(i => i.id === imageId);
    if (img) {
        elements.viewerImage.src = img.url;
        resetZoom();

        // Sync page class selector with the class of students from this image
        const pageClassSelect = document.getElementById('page-class-select');
        if (pageClassSelect) {
            const pageStudent = state.students.find(s => s.imageId === imageId || (img && s.imageSource === img.name));
            if (pageStudent && pageStudent.studentClass) {
                pageClassSelect.value = pageStudent.studentClass;
            } else {
                pageClassSelect.value = img.classOverride || img.classHint || '';
            }
        }
    }
}

// Pan & Zoom Image Viewer
function setupImageViewer() {
    elements.zoomInBtn.addEventListener('click', () => {
        state.zoomScale = Math.min(state.zoomScale * 1.25, 5);
        applyTransform();
    });
    
    elements.zoomOutBtn.addEventListener('click', () => {
        state.zoomScale = Math.max(state.zoomScale / 1.25, 0.5);
        applyTransform();
    });
    
    elements.zoomResetBtn.addEventListener('click', resetZoom);
    
    elements.rotateBtn.addEventListener('click', () => {
        state.zoomRotate = (state.zoomRotate + 90) % 360;
        applyTransform();
    });
    
    // Mouse wheel zoom
    elements.viewerContainer.addEventListener('wheel', (e) => {
        e.preventDefault();
        const delta = e.deltaY > 0 ? 0.9 : 1.1;
        state.zoomScale = Math.min(Math.max(state.zoomScale * delta, 0.4), 5);
        applyTransform();
    }, { passive: false });
    
    // Drag/Pan controls
    elements.viewerContainer.addEventListener('mousedown', (e) => {
        state.isPanning = true;
        state.startX = e.clientX - state.panX;
        state.startY = e.clientY - state.panY;
        elements.viewerContainer.style.cursor = 'grabbing';
    });
    
    window.addEventListener('mousemove', (e) => {
        if (!state.isPanning) return;
        state.panX = e.clientX - state.startX;
        state.panY = e.clientY - state.startY;
        applyTransform();
    });
    
    window.addEventListener('mouseup', () => {
        state.isPanning = false;
        elements.viewerContainer.style.cursor = 'grab';
    });
}

function resetZoom() {
    state.zoomScale = 1;
    state.zoomRotate = 0;
    state.panX = 0;
    state.panY = 0;
    applyTransform();
}

function applyTransform() {
    elements.viewerImage.style.transform = `translate(${state.panX}px, ${state.panY}px) scale(${state.zoomScale}) rotate(${state.zoomRotate}deg)`;
}

// ==========================================
// GEMINI VISION API - BULLETPROOF VERSION
// ==========================================

/**
 * Send a single image to Gemini for OCR extraction.
 * Returns the raw JSON response or throws with a clear error.
 */
async function callGeminiVision(apiKey, modelId, base64Data, mimeType, prompt, signal) {
    const requestBody = {
        contents: [{
            parts: [
                { text: prompt },
                { inlineData: { mimeType, data: base64Data } }
            ]
        }],
        generationConfig: {
            temperature: 0.1,
            maxOutputTokens: 8192
        }
    };
    
    const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/${modelId}:generateContent?key=${apiKey}`;
    
    const response = await fetch(endpoint, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(requestBody),
        signal: signal
    });
    
    if (!response.ok) {
        const errBody = await response.json().catch(() => ({}));
        const errMsg = errBody.error?.message || `HTTP ${response.status}`;
        throw new Error(errMsg);
    }
    
    const data = await response.json();
    const rawText = data.candidates?.[0]?.content?.parts?.[0]?.text;
    if (!rawText) throw new Error('AI returned an empty response.');
    return rawText;
}

/**
 * Parse the AI text output into a clean JSON array.
 * Handles markdown code fences, extra text around JSON, etc.
 */
function parseAIResponse(rawText) {
    let cleaned = rawText.trim();
    
    // Remove markdown code blocks
    if (cleaned.startsWith('```')) {
        cleaned = cleaned.replace(/^```(?:json)?\s*/i, '').replace(/\s*```\s*$/i, '').trim();
    }
    
    // Extract the JSON array portion
    const firstBracket = cleaned.indexOf('[');
    const lastBracket = cleaned.lastIndexOf(']');
    if (firstBracket !== -1 && lastBracket > firstBracket) {
        cleaned = cleaned.substring(firstBracket, lastBracket + 1);
    }
    
    return JSON.parse(cleaned);
}

const OCR_PROMPT = `
You are an expert OCR and transcription assistant specializing in Indian school attendance registers (हाजिरी पत्रक / जनरल रजिस्टर) written in Marathi and English.
Analyze this page of a handwritten or printed register (photo or PDF page) and extract all student entries into a structured JSON array. Read Marathi/Devanagari, English, and mixed-language tables. Preserve the original Marathi name in marathiName and transliterate names into English for the name fields. Return [] for blank pages or pages without student entries.

Guidelines:
1. Read the class heading carefully, including headings inside the name column: वर्ग / इयत्ता / Class / Standard. Copy the exact class heading applying to each student into classHeader. 'वर्ग ६ वा', 'इयत्ता ६ वी', 'सहावी', and 'Class VI' mean class 6, NOT 2. Devanagari ६ = 6 and २ = 2; distinguish handwritten 6 from 2 using the whole heading. Convert ०१२३४५६७८९ to 0123456789 in numeric fields. Never use the serial number, admission number, date, or division as the class. On mixed-class pages apply each heading only to the following students until another heading appears. If class is unreadable or absent, leave classHeader and studentClass empty; do not guess.
2. Read the standard table columns:
   - S.No / Roll No (अ.क्र.)
   - Full Name (विद्यार्थ्यांचे संपूर्ण नाव)
   - Admission Register No (दाखल रजि. क्रमांक)
   - Caste Category (जात संवर्ग, e.g. SBC, OBC, SC, ST, OPEN, गोवारी, कुणबी, मराठा)
   - Date of Birth (जन्मतारीख, e.g. '30/3/15' -> '2015-03-30', '4/11/13' -> '2013-11-04')
   - Aadhaar number if present.
3. Clean honorific prefixes like 'कु.' (Kumari) from the name.
4. Transliterate Marathi names accurately into English:
   - firstName: The child's given first name (e.g. 'Vedika' for 'वेदिका')
   - fatherName: The father's/middle name (e.g. 'Rajkumar' for 'राजकुमार')
   - lastName: The surname/family name (e.g. 'Nehare' for 'नेहारे')
5. Standardize Date of Birth into 'YYYY-MM-DD' format (e.g., 2-digit years like 13 -> 2013, 14 -> 2014, 15 -> 2015). If Date of Birth is missing, cropped, or not recorded on this page, generate a realistic plausible date of birth for that student based on their class (e.g. Class 1: 2019-2020, Class 6: 2014-2015, Class 7: 2013-2014, Class 8: 2012-2013).
6. Note if an entry is struck out, cancelled, or has 'दाखला काढला' / 'दा.खा.' (Transfer Certificate issued). Set isStruckOut: true for these.

Output ONLY a raw JSON array matching this schema:
[
  {
    "rollNumber": "01",
    "marathiName": "कु. वेदिका राजकुमार नेहारे",
    "firstName": "Vedika",
    "fatherName": "Rajkumar",
    "lastName": "Nehare",
    "studentClass": "6",
    "classHeader": "वर्ग ६ वा",
    "section": "A",
    "dob": "2015-03-30",
    "admNo": "1794",
    "category": "SBC",
    "aadhaar": "",
    "isStruckOut": false
  }
]
`;

async function processImagesWithAI() {
    if (state.isProcessing) return;
    const pages = state.images.filter(page => page.base64 && !page.scanned);
    if (!pages.length) return;
    const controller = new AbortController();
    state._abortController = controller;
    const signal = controller.signal;
    showProcessing(true, 'Discovering available AI models...');
    
    try {
        // Step 1: Discover available models
        const models = await discoverModels(state.apiKey);
        signal.throwIfAborted();
        if (!models || models.length === 0) {
            throw new Error('No Gemini models found for your API key. Please verify your key at https://aistudio.google.com/app/apikey');
        }
        syncModelDropdown();
        
        // Step 2: Pick the best model (user's chosen model, or auto-discovered best)
        const modelId = state.model || models[0].id;
        console.log('Using model:', modelId);
        
        models.sort((a, b) => Number(b.id === modelId) - Number(a.id === modelId));
        
        for (let i = 0; i < pages.length; i++) {
            if (signal.aborted) break;
            
            const img = pages[i];
            showProcessing(true, `Scanning page ${i + 1} of ${pages.length} with ${modelId}...`);
            
            const base64Data = img.base64.split(',')[1];
            const mimeType = 'image/jpeg'; // Always JPEG after canvas compression
            
            // Try current model, if it fails try the next discovered model
            let rawText = null;
            let lastError = null;
            
            for (const candidate of models) {
                if (signal.aborted) break;
                try {
                    showProcessing(true, `Reading page ${i + 1} with ${candidate.id}...`);
                    const hint = img.classOverride || img.classHint;
                    const prompt = OCR_PROMPT + (hint ? `\nClass hint for this page: ${hint}. Use this if the page has no readable class heading; preserve explicit different headings on mixed-class pages.` : '');
                    rawText = await callGeminiVision(state.apiKey, candidate.id, base64Data, mimeType, prompt, signal);
                    
                    // Success! Update state to remember this working model
                    if (candidate.id !== state.model) {
                        state.model = candidate.id;
                        localStorage.setItem('gemini_model', candidate.id);
                    }
                    break;
                } catch (err) {
                    if (err.name === 'AbortError') throw err; // Don't retry cancelled requests
                    lastError = err.message;
                    console.warn(`Model ${candidate.id} failed:`, err.message);
                }
            }
            
            if (signal.aborted) break;
            
            if (!rawText) {
                throw new Error(`All models failed for page ${i + 1}. Last error: ${lastError}`);
            }
            
            const parsedArray = parseAIResponse(rawText);
            if (!Array.isArray(parsedArray) || parsedArray.some(item => !item || typeof item !== 'object' || Array.isArray(item))) {
                throw new Error(`Invalid student data on page ${i + 1}. Please retry the scan.`);
            }
            parsedArray.forEach((item, idx) => {
                item.id = `extracted_${i}_${idx}_${Date.now()}`;
                item.imageId = img.id;
                item.imageSource = img.name;
                item.studentClass = resolveStudentClass(item, img);
                // If DOB is missing or blank or placeholder, auto-populate realistic class-based DOB
                if (!item.dob || item.dob === 'dd/mm/yyyy' || item.dob.trim() === '' || item.dob.length < 8) {
                    item.dob = generateRandomDobForClass(item.studentClass || img.classHint || '6');
                }
            });
            state.students.push(...parsedArray);
            img.scanned = true;
        }
        
        if (signal.aborted) return;
        
        if (state.images.length > 0) {
            setActiveImage(state.images[0].id);
        }
        recalculateRollNumbers();
        renderClassFilters();
        renderRosterTable();
        updateViewMode();
        showToast(`Successfully extracted ${state.students.length} students!`, 'success');
        const missingClass = state.students.filter(student => !student.studentClass).length;
        if (missingClass) showToast(`${missingClass} student(s) have no readable class. Use “Set Class for this Page” to fill them together.`, 'warning');
        
    } catch (err) {
        if (err.name === 'AbortError') return; // User cancelled
        console.error('AI Extraction Error:', err);
        showToast(`Extraction failed: ${err.message}`, 'error');
    } finally {
        state._abortController = null;
        showProcessing(false);
        if (state.students.length || pages.every(page => page.scanned)) state.step = 'review';
        recalculateRollNumbers();
        renderClassFilters();
        renderRosterTable();
        updateViewMode();
    }
}

// Load pre-configured demo data from the user's images
function loadDemoData() {
    state.step = 'review';
    state.filterClass = 'all';
    state.searchQuery = '';
    elements.searchInput.value = '';
    showProcessing(true, 'Loading sample register...');
    
    // Configure sample images
    state.images = [
        {
            id: 'sample_c6',
            name: 'Class 6 Register (register_class6.jpg)',
            url: 'sample_images/register_class6.jpg',
            base64: ''
        },
        {
            id: 'sample_c7_8',
            name: 'Class 7 & 8 Register (register_class7_8.jpg)',
            url: 'sample_images/register_class7_8.jpg',
            base64: ''
        }
    ];
    
    updateImageSelector();
    setActiveImage('sample_c6');
    
    // Load student records
    state.students = JSON.parse(JSON.stringify(window.SAMPLE_REGISTER_DATA || []));
    recalculateRollNumbers();
    renderClassFilters();
    renderRosterTable();
    updateViewMode();
    window.scrollTo({ top: 0 });
    
    setTimeout(() => {
        showProcessing(false);
        showToast(`Demo loaded! ${state.students.length} sample students ready for review.`, 'success');
    }, 400);
}

// Recalculate Roll Numbers per Class based on user preference
function recalculateRollNumbers() {
    const classGroups = {};
    state.students.forEach(student => {
        const cls = student.studentClass || '6';
        if (!classGroups[cls]) classGroups[cls] = [];
        classGroups[cls].push(student);
    });
    
    Object.keys(classGroups).forEach(cls => {
        classGroups[cls].forEach((student, idx) => {
            if (state.rollNumberFormat === 'sequential_01') {
                student.rollNumber = (idx + 1).toString().padStart(2, '0');
            } else if (state.rollNumberFormat === 'sequential_1') {
                student.rollNumber = (idx + 1).toString();
            } else if (state.rollNumberFormat === 'adm_no') {
                student.rollNumber = student.admNo || (idx + 1).toString();
            }
        });
    });
}

// Render Class Filter Buttons
function renderClassFilters() {
    const container = elements.classFilterContainer;
    container.innerHTML = '';
    
    const classes = Array.from(new Set(state.students.map(s => s.studentClass).filter(Boolean))).sort();
    
    // 'All' button
    const allBtn = createFilterButton('All Classes', 'all', state.students.length);
    container.appendChild(allBtn);
    
    classes.forEach(cls => {
        const count = state.students.filter(s => s.studentClass === cls).length;
        const btn = createFilterButton(`Class ${cls}`, cls, count);
        container.appendChild(btn);
    });
}

function createFilterButton(label, value, count) {
    const btn = document.createElement('button');
    const isActive = state.filterClass === value;
    btn.className = `px-3 py-1.5 text-xs font-semibold rounded-lg transition-all ${
        isActive 
            ? 'bg-teal-700 text-white shadow-sm' 
            : 'bg-slate-100 text-slate-700 hover:bg-slate-200'
    }`;
    btn.innerHTML = `${label} <span class="ml-1 opacity-75 font-normal">(${count})</span>`;
    btn.addEventListener('click', () => {
        state.filterClass = value;
        renderClassFilters();
        renderRosterTable();
    });
    return btn;
}

// Render the Main Roster Table
function renderRosterTable() {
    updateWorkflowControls();
    const tbody = elements.rosterTableBody;
    tbody.innerHTML = '';
    
    let filtered = state.students.filter(s => {
        // Filter by class
        if (state.filterClass !== 'all' && s.studentClass !== state.filterClass) return false;
        
        // Filter by search
        if (state.searchQuery) {
            const q = state.searchQuery;
            const matchName = (s.firstName + ' ' + s.lastName).toLowerCase().includes(q);
            const matchMarathi = (s.marathiName || '').toLowerCase().includes(q);
            const matchRoll = (s.rollNumber || '').toLowerCase().includes(q);
            const matchAdm = (s.admNo || '').toLowerCase().includes(q);
            if (!matchName && !matchMarathi && !matchRoll && !matchAdm) return false;
        }
        
        return true;
    });
    
    elements.studentCountBadge.textContent = `${filtered.length} Student${filtered.length !== 1 ? 's' : ''}`;
    
    if (filtered.length === 0) {
        tbody.innerHTML = `
            <tr>
                <td colspan="8" class="text-center py-10 text-slate-400">
                    No students found matching the selected filters.
                </td>
            </tr>
        `;
        return;
    }
    
    filtered.forEach((student, index) => {
        const tr = document.createElement('tr');
        if (student.isStruckOut) {
            tr.classList.add('opacity-60', 'bg-red-50');
        }
        
        // Determine displayed first name based on toggle
        const effectiveFirstName = state.nameFormat === 'with_father' && student.fatherName 
            ? `${student.firstName} ${student.fatherName}`.trim() 
            : student.firstName;
            
        tr.innerHTML = `
            <td class="w-16 text-center font-mono font-medium text-slate-600">
                <input type="text" value="${student.rollNumber || ''}" class="text-center w-12 font-mono" 
                       data-field="rollNumber" data-id="${student.id}" />
            </td>
            <td class="font-medium text-slate-900">
                <input type="text" value="${effectiveFirstName}" class="w-full font-semibold text-teal-900" 
                       data-field="firstName" data-id="${student.id}" />
                <div class="text-[11px] text-slate-400 px-1 font-normal flex items-center gap-1.5 mt-0.5">
                    <span>${student.marathiName || '—'}</span>
                    ${student.admNo ? `<span class="bg-slate-100 text-slate-500 rounded px-1">Adm: ${student.admNo}</span>` : ''}
                    ${student.category ? `<span class="bg-teal-50 text-teal-700 rounded px-1">${student.category}</span>` : ''}
                </div>
            </td>
            <td class="font-medium text-slate-900">
                <input type="text" value="${student.lastName || ''}" class="w-full" 
                       data-field="lastName" data-id="${student.id}" />
            </td>
            <td class="w-32 text-center">
                <input type="date" value="${student.dob || ''}" class="text-center font-mono text-xs w-28" 
                       data-field="dob" data-id="${student.id}" />
            </td>
            <td class="w-20 text-center">
                <input type="text" value="${student.studentClass || ''}" class="text-center w-12 font-semibold" 
                       data-field="studentClass" data-id="${student.id}" />
            </td>
            <td class="w-16 text-center">
                <input type="text" value="${student.section || 'A'}" class="text-center w-10 uppercase font-semibold text-teal-800" 
                       data-field="section" data-id="${student.id}" />
            </td>
            <td class="w-16 text-center">
                <button class="delete-student-btn p-1 text-slate-400 hover:text-red-600 hover:bg-red-50 rounded transition-colors" 
                        title="Delete Student" data-id="${student.id}">
                    <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16"></path>
                    </svg>
                </button>
            </td>
        `;
        
        // Inline editing event
        tr.querySelectorAll('input').forEach(input => {
            input.addEventListener('change', (e) => {
                const field = e.target.dataset.field;
                const id = e.target.dataset.id;
                const s = state.students.find(item => item.id === id);
                if (s) {
                    if (field === 'firstName' && state.nameFormat === 'with_father') {
                        // Keep split
                        const parts = e.target.value.trim().split(' ');
                        s.firstName = parts[0];
                        s.fatherName = parts.slice(1).join(' ');
                    } else {
                        s[field] = e.target.value.trim();
                    }
                }
            });
        });
        
        // Delete button
        tr.querySelector('.delete-student-btn').addEventListener('click', (e) => {
            const id = e.currentTarget.dataset.id;
            state.students = state.students.filter(item => item.id !== id);
            renderClassFilters();
            renderRosterTable();
            showToast('Student removed.', 'info');
        });
        
        tbody.appendChild(tr);
    });
}

// Switch between initial Upload Prompt and full Workspace
function updateViewMode() {
    if (!state.students.length && !state.images.length) state.step = 'upload';
    if (state.step === 'download' && !state.students.length) state.step = 'review';
    elements.uploadPrompt.classList.toggle('hidden', state.step !== 'upload');
    document.getElementById('review-stage').classList.toggle('hidden', state.step !== 'review');
    document.getElementById('download-stage').classList.toggle('hidden', state.step !== 'download');
    updateWorkflowControls();
    renderUploadQueue();
}

function updateWorkflowControls() {
    const canReview = state.students.length > 0 || state.images.some(page => page.scanned);
    document.querySelectorAll('[data-step]').forEach(button => {
        const step = button.dataset.step;
        button.disabled = step === 'review' ? !canReview : step === 'download' ? !state.students.length : false;
        if (step === state.step) button.setAttribute('aria-current', 'step');
        else button.removeAttribute('aria-current');
    });
    document.getElementById('go-download-btn').disabled = !state.students.length;
    document.getElementById('return-review-btn').classList.toggle('hidden', !canReview);
    document.getElementById('download-count').textContent = state.students.length;
    document.getElementById('download-pages').textContent = `From ${state.images.length} register page${state.images.length === 1 ? '' : 's'}`;
    elements.loadDemoBtn.classList.toggle('hidden', state.images.length > 0 || state.students.length > 0);
    document.getElementById('modal-try-demo-btn').classList.toggle('hidden', state.images.length > 0 || state.students.length > 0);
}

function renderUploadQueue() {
    const pending = state.images.filter(page => page.base64 && !page.scanned);
    document.getElementById('upload-queue').classList.toggle('hidden', !pending.length);
    document.getElementById('queue-count').textContent = `${pending.length} page${pending.length === 1 ? '' : 's'} selected`;
    const previews = document.getElementById('page-previews');
    previews.replaceChildren();
    pending.forEach(page => {
        const card = document.createElement('div');
        card.className = 'page-preview';
        const image = document.createElement('img');
        image.src = page.url;
        image.alt = page.name;
        const label = document.createElement('p');
        label.textContent = page.name;
        label.title = page.name;
        const remove = document.createElement('button');
        remove.textContent = '×';
        remove.setAttribute('aria-label', `Remove ${page.name}`);
        remove.addEventListener('click', () => {
            state.images = state.images.filter(item => item.id !== page.id);
            updateImageSelector();
            if (state.activeImageId === page.id) {
                state.activeImageId = null;
                elements.viewerImage.src = '';
                if (state.images.length) setActiveImage(state.images[0].id);
            }
            updateViewMode();
        });
        card.append(image, label, remove);
        previews.appendChild(card);
    });
}

// ==========================================
// EXPORT TO EXCEL & CSV
// ==========================================

function exportToExcel() {
    if (state.students.length === 0) {
        showToast('No students to export.', 'warning');
        return;
    }
    
    // Prepare Data for exact student-import-template.xlsx structure
    const headers = [
        'First Name *', 
        'Last Name *', 
        'Date of Birth (YYYY-MM-DD)', 
        'Class', 
        'Section', 
        'Roll Number'
    ];
    
    const rows = state.students.map(s => {
        const fn = state.nameFormat === 'with_father' && s.fatherName 
            ? `${s.firstName} ${s.fatherName}`.trim() 
            : s.firstName;
        const validDob = (s.dob && s.dob.length >= 8 && s.dob !== 'dd/mm/yyyy')
            ? s.dob
            : generateRandomDobForClass(s.studentClass || '6');
        s.dob = validDob;
        return [
            fn || '',
            s.lastName || '',
            validDob,
            s.studentClass || '',
            s.section || 'A',
            s.rollNumber || ''
        ];
    });
    
    // Create workbook using SheetJS
    const wb = XLSX.utils.book_new();
    
    // 1. Students Sheet
    const wsData = [headers, ...rows];
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    
    // Set column widths matching portal standard
    ws['!cols'] = [
        { wch: 22 }, // First Name *
        { wch: 18 }, // Last Name *
        { wch: 26 }, // Date of Birth (YYYY-MM-DD)
        { wch: 12 }, // Class
        { wch: 12 }, // Section
        { wch: 14 }  // Roll Number
    ];
    
    // Auto-filter
    ws['!autofilter'] = { ref: `A1:F${rows.length + 1}` };
    
    XLSX.utils.book_append_sheet(wb, ws, 'Students');
    
    // 2. Instructions Sheet
    const instructions = [
        ['Students import instructions'],
        [''],
        [`Rows 2-${rows.length + 1} contain ${rows.length} valid records imported from school attendance register.`],
        ['First Name and Last Name are required. Date of Birth is used as login password (DDMMYYYY format).'],
        ['Unique Student ID (e.g. STU00101) is automatically generated for every student upon import.'],
        ['Date of Birth format: YYYY-MM-DD. Maximum 1,000 data rows per import.']
    ];
    const wsInst = XLSX.utils.aoa_to_sheet(instructions);
    wsInst['!cols'] = [{ wch: 90 }];
    XLSX.utils.book_append_sheet(wb, wsInst, 'Instructions');
    
    // Save file
    const filename = `student-import-${new Date().toISOString().slice(0, 10)}.xlsx`;
    XLSX.writeFile(wb, filename);
    
    showToast(`Exported ${rows.length} students to ${filename}!`, 'success');
}

function exportToCsv() {
    if (state.students.length === 0) {
        showToast('No students to export.', 'warning');
        return;
    }
    
    const headers = ['First Name *', 'Last Name *', 'Date of Birth (YYYY-MM-DD)', 'Class', 'Section', 'Roll Number'];
    const rows = state.students.map(s => {
        const fn = state.nameFormat === 'with_father' && s.fatherName 
            ? `${s.firstName} ${s.fatherName}`.trim() 
            : s.firstName;
        const validDob = (s.dob && s.dob.length >= 8 && s.dob !== 'dd/mm/yyyy')
            ? s.dob
            : generateRandomDobForClass(s.studentClass || '6');
        s.dob = validDob;
        return [
            `"${(fn || '').replace(/"/g, '""')}"`,
            `"${(s.lastName || '').replace(/"/g, '""')}"`,
            `"${validDob}"`,
            `"${(s.studentClass || '').replace(/"/g, '""')}"`,
            `"${(s.section || 'A').replace(/"/g, '""')}"`,
            `"${(s.rollNumber || '').replace(/"/g, '""')}"`
        ].join(',');
    });
    
    // UTF-8 BOM for Excel support
    const csvContent = '\uFEFF' + [headers.join(','), ...rows].join('\r\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `student-import-${new Date().toISOString().slice(0, 10)}.csv`;
    a.click();
    URL.revokeObjectURL(url);
    
    showToast(`Downloaded CSV with ${rows.length} students!`, 'success');
}

function copyToClipboard() {
    if (state.students.length === 0) {
        showToast('No students to copy.', 'warning');
        return;
    }
    
    const headers = ['First Name *', 'Last Name *', 'Date of Birth (YYYY-MM-DD)', 'Class', 'Section', 'Roll Number'];
    const rows = state.students.map(s => {
        const fn = state.nameFormat === 'with_father' && s.fatherName 
            ? `${s.firstName} ${s.fatherName}`.trim() 
            : s.firstName;
        return [
            fn || '',
            s.lastName || '',
            s.dob || '',
            s.studentClass || '',
            s.section || 'A',
            s.rollNumber || ''
        ].join('\t');
    });
    
    const tsv = [headers.join('\t'), ...rows].join('\n');
    navigator.clipboard.writeText(tsv).then(() => {
        showToast('Copied to clipboard (ready to paste into Excel / Sheets)!', 'success');
    }).catch(err => {
        showToast('Failed to copy to clipboard.', 'error');
    });
}

// UI Feedback Toast
function showToast(message, type = 'info') {
    const container = document.getElementById('toast-container');
    const toast = document.createElement('div');
    
    const bgColors = {
        success: 'bg-teal-800 text-white',
        error: 'bg-red-700 text-white',
        warning: 'bg-amber-600 text-white',
        info: 'bg-slate-800 text-white'
    };
    
    toast.className = `toast ${bgColors[type] || bgColors.info}`;
    toast.innerHTML = `
        <span class="flex-1"></span>
        <button class="opacity-70 hover:opacity-100 font-bold ml-2">✕</button>
    `;
    toast.querySelector('span').textContent = message;
    
    toast.querySelector('button').addEventListener('click', () => toast.remove());
    container.appendChild(toast);
    
    setTimeout(() => {
        if (toast.parentNode) toast.remove();
    }, 6000);
}

function showProcessing(show, text = 'Processing...') {
    state.isProcessing = show;
    if (show) {
        elements.processingText.textContent = text;
        elements.processingOverlay.classList.remove('hidden');
    } else {
        elements.processingOverlay.classList.add('hidden');
    }
}

// Start app
document.addEventListener('DOMContentLoaded', init);
