const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const { webcrypto } = require('node:crypto');

function app() {
    const context = vm.createContext({
        localStorage: { getItem: () => '', setItem() {} },
        document: { getElementById: () => ({ value: '', classList: { remove() {}, add() {} } }), addEventListener() {} },
        crypto: webcrypto, AbortController, console: { log() {}, warn() {}, error() {} },
    });
    vm.runInContext(fs.readFileSync('app.js', 'utf8'), context);
    vm.runInContext(fs.readFileSync('excel-import.js', 'utf8'), context);
    vm.runInContext(`
        const notices = [];
        showToast = (message, type) => notices.push({message, type});
        showProcessing = value => { state.isProcessing = value; };
        updateImageSelector = setActiveImage = updateViewMode = renderClassFilters = renderRosterTable = syncModelDropdown = () => {};
    `, context);
    return { context, run: code => vm.runInContext(code, context) };
}

test('filename dates and camera counters never become class 2', () => {
    const { context, run } = app();
    for (const name of ['IMG_20260917_123456.jpg', '2026-09-17.pdf', 'WhatsApp Image 2026-09-17.jpeg', 'scan_2.jpg', 'shirpur7.jpg']) {
        context.name = name;
        assert.equal(run('detectClassFromFilename(name)'), null, name);
    }
    for (const [name, expected] of [['register_class6.jpg', '6'], ['Std १२.pdf', '12'], ['class10.pdf', '10'], ['वर्ग ६.pdf', '6'], ['register_class7_8.jpg', null], ['class 7 & 8.pdf', null]]) {
        context.name = name;
        assert.equal(run('detectClassFromFilename(name)'), expected, name);
    }
});

test('Marathi headings correct 6 versus 2 while preserving actual class 2 and mixed classes', () => {
    const { context, run } = app();
    for (const heading of ['वर्ग ६ वा', 'इयत्ता ६ वी', 'सहावी', 'Class VI', '६']) {
        context.heading = heading;
        assert.equal(run(`resolveStudentClass({studentClass:'2', classHeader:heading}, {})`), '6');
    }
    assert.equal(run(`resolveStudentClass({studentClass:'2'}, {classHint:'6'})`), '6');
    assert.equal(run(`resolveStudentClass({studentClass:'2', classHeader:'वर्ग २ री'}, {classHint:'6'})`), '2');
    assert.equal(run(`resolveStudentClass({studentClass:'8', classHeader:'वर्ग ८ वा'}, {classHint:'7'})`), '8');
    assert.equal(run(`resolveStudentClass({}, {})`), '');
    assert.equal(run(`resolveStudentClass({classHeader:'वर्ग २ री'}, {classOverride:'6'})`), '6');
});

test('PDF pages render sequentially with source labels and release resources', async () => {
    const { run } = app();
    await run(`
        const rendered = [], cleaned = [];
        let destroyed = 0;
        document.createElement = () => ({getContext: () => ({}), toDataURL: () => 'data:image/jpeg;base64,abc'});
        loadPdfLibrary = async () => ({getDocument: () => ({
            promise: Promise.resolve({numPages: 2, getPage: async n => ({
                getViewport: ({scale}) => ({width: 600*scale, height: 800*scale}),
                render: ({viewport}) => { rendered.push([n, viewport.height]); return {promise: Promise.resolve()}; },
                cleanup: () => cleaned.push(n)
            })}), destroy: async () => { destroyed++; }
        })});
        let pages;
        pdfToPages({name:'class6.PDF',size:100,arrayBuffer:async()=>new ArrayBuffer(1)}, new AbortController().signal).then(result => {pages=result;});
    `);
    assert.equal(run('pages.length'), 2);
    assert.equal(run('pages[1].name'), 'class6.PDF — page 2');
    assert.equal(run('pages[1].classHint'), '6');
    assert.equal(run('rendered[0][1]'), 2400);
    assert.equal(run('cleaned.length'), 2);
    assert.equal(run('destroyed'), 1);
    assert.notEqual(run('pages[0].id'), run('pages[1].id'));
});

test('PDF password errors and cancellation release the document', async () => {
    const { run } = app();
    run(`let destroyed = 0; loadPdfLibrary = async () => ({getDocument: () => ({promise: Promise.reject(Object.assign(new Error('password'), {name:'PasswordException'})), destroy: async () => { destroyed++; }})});`);
    await assert.rejects(run(`pdfToPages({name:'locked.pdf',size:1,arrayBuffer:async()=>new ArrayBuffer(1)},new AbortController().signal)`), /password protected/);
    assert.equal(run('destroyed'), 1);
    await assert.rejects(run(`pdfToPages({size:51*1024*1024},new AbortController().signal)`), /50 MB/);
    await assert.rejects(run(`const cancelled = new AbortController(); cancelled.abort(); pdfToPages({size:1},cancelled.signal)`), { name: 'AbortError' });
});

test('adding pages preserves edits and retry does not duplicate completed pages', async () => {
    const { run } = app();
    run(`
        state.apiKey = 'test';
        state.students = [{id:'existing', firstName:'Edited',studentClass:'6'}];
        state.images = [makePage('class6.pdf — page 1','data:image/jpeg;base64,one','6'),makePage('class6.pdf — page 2','data:image/jpeg;base64,two','6')];
        discoverModels = async () => [{id:'test'}];
        let calls = 0;
        callGeminiVision = async () => {
            calls++;
            if(calls === 2) throw new Error('Temporary failure');
            return JSON.stringify([{firstName:'Student',studentClass:'2',dob:'2015-01-01'}]);
        };
    `);
    await run('processImagesWithAI()');
    assert.equal(run('state.students.length'), 2);
    assert.equal(run('state.students[1].studentClass'), '6');
    assert.equal(run('state.images[0].scanned'), true);
    assert.equal(run('state.images[1].scanned'), false);
    await run('processImagesWithAI()');
    assert.equal(run('state.students.length'), 3);
    assert.equal(run('state.students[0].firstName'), 'Edited');
    assert.equal(run('calls'), 3);
    await run('processImagesWithAI()');
    assert.equal(run('calls'), 3);
});

test('cancellation during discovery sends no OCR requests', async () => {
    const { run } = app();
    run(`state.images=[makePage('page','data:image/jpeg;base64,one',null)];
         discoverModels=async()=>{state._abortController.abort();return [{id:'test'}];};
         let calls=0;callGeminiVision=async()=>{calls++;};`);
    await run('processImagesWithAI()');
    assert.equal(run('calls'), 0);
    assert.equal(run('state.isProcessing'), false);
});

test('upload prepares previews without starting OCR or replacing edited students', async () => {
    const { run } = app();
    run(`state.apiKey='configured'; state.step='review';
        state.students=[{id:'existing',firstName:'Edited',studentClass:'6'}];
        fileToBase64=async()=> 'data:image/jpeg;base64,test';
        let scanCalls=0; processImagesWithAI=async()=>{scanCalls++;};`);
    await run(`handleFiles([{name:'class6.jpg',type:'image/jpeg'}])`);
    assert.equal(run('scanCalls'), 0);
    assert.equal(run('state.step'), 'upload');
    assert.equal(run('state.images.length'), 1);
    assert.equal(run('state.images[0].scanned'), false);
    assert.equal(run('state.students[0].firstName'), 'Edited');
    assert.equal(run('state.isProcessing'), false);
});

test('excel rows: father name dropped, sheets merged by sheet name, only classes 6-10, sorted', () => {
    const { run } = app();
    const out = run(`(() => {
        const head = ['Sr No', 'Student Name', 'Date of Birth', 'Father Name Extra'];
        const a = studentsFromSheet([['School X'], [], head,
            [2, 'KU. SNEHA RAMESH PATIL', '15/03/2012', ''], [1, 'Amit Suresh Kale', '2011-07-04', ''], ['', 'Total', '', '']], 'Class 7', 'A');
        const b = studentsFromSheet([['First Name', 'Middle Name', 'Surname', 'Std', 'DOB'],
            ['Riya', 'Mohan', 'Shah', 'VIII', '01-Jan-13'], ['Tom', 'X', 'Y', '3', ''], ['Zed', 'Q', 'W', '11', '']], 'Sheet1', 'A');
        const c = studentsFromSheet([['Foo']], 'Summary', 'A');
        const all = sortRoster([...b.students, ...a.students]);
        return JSON.stringify({ all: all.map(s => [s.firstName, s.fatherName, s.lastName, s.dob, s.studentClass, s.rollNumber]), out: b.outOfRange, skipped: c.skipped });
    })()`);
    const r = JSON.parse(out);
    assert.deepEqual(r.all, [
        ['Amit', 'Suresh', 'Kale', '2011-07-04', '7', '1'],
        ['Sneha', 'Ramesh', 'Patil', '2012-03-15', '7', '2'],
        ['Riya', 'Mohan', 'Shah', '2013-01-01', '8', ''],
    ]);
    assert.equal(r.out, 2);
    assert.ok(r.skipped);
});

test('parseDobCell and parseClassCell handle common school formats', () => {
    const { run } = app();
    const d = JSON.parse(run(`JSON.stringify(['5/9/2012', '2012.09.05', '5 Sep 2012', '05092012', 'junk'].map(parseDobCell))`));
    assert.deepEqual(d, ['2012-09-05', '2012-09-05', '2012-09-05', '2012-09-05', '']);
    const c = JSON.parse(run(`JSON.stringify(['6', '6th', 'Class 7', 'IX', '8-B', '10A', 'सहावी'].map(parseClassCell))`));
    assert.deepEqual(c.map(x => x.cls), [6, 6, 7, 9, 8, 10, 6]);
    assert.equal(c[4].section, 'B');
});

test('excel headers: "Student Full Name" is a full-name column, not a surname column', () => {
    const { context, run } = app();
    const cases = { 'Student Full Name': 'full', 'Full Name': 'full', 'Name of Student': 'full', 'Surname': 'last', 'Last Name *': 'last',
        'First Name *': 'first', 'FName': 'first', 'LName': 'last', "Father's Name": 'middle', 'Sr No': 'roll', 'D.O.B.': 'dob' };
    for (const [header, role] of Object.entries(cases)) {
        context.header = header;
        assert.equal(run('detectColumn(header)'), role, header);
    }
});

test('excel AI fallback: Marathi detection, chunking, output cleanup and duplicates', () => {
    const { run } = app();
    const r = JSON.parse(run(`JSON.stringify({
        marathi: hasMarathiNames([{ firstName: 'वेदिका', lastName: 'नेहारे' }, { firstName: 'Amit', lastName: 'Kale' }]),
        english: hasMarathiNames([{ firstName: 'Amit', lastName: 'Kale' }]),
        chunks: sheetCsvChunks(Array.from({ length: 170 }, (_, i) => 'row' + i).join(String.fromCharCode(10))).map(c => [c.context.split(String.fromCharCode(10))[0], c.rows.split(String.fromCharCode(10)).length]),
        ai: normalizeAiStudents([
            { firstName: 'Ku. VEDIKA', fatherName: 'Rajkumar', lastName: 'Nehare', studentClass: 'VI', section: '', dob: '30/03/2014', rollNumber: 2 },
            { firstName: 'Tom', lastName: 'Y', studentClass: '4' },
            { firstName: '', lastName: 'Nobody', studentClass: '7' },
            'junk'
        ], 'Std 8 B', 'A'),
        noClass: normalizeAiStudents([{ firstName: 'Riya', lastName: 'Shah', studentClass: '' }], 'Std 8 B', 'A').students[0].studentClass,
        dupes: finalizeRoster([
            { firstName: 'A', lastName: 'B', dob: '2014-01-01', studentClass: '6', section: 'A', rollNumber: '' },
            { firstName: 'a', lastName: 'b', dob: '2014-01-01', studentClass: '6', section: 'A', rollNumber: '' },
            { firstName: 'A', lastName: 'B', dob: '', studentClass: '6', section: 'A', rollNumber: '' }
        ]).duplicates
    })`));
    assert.equal(r.marathi, true);
    assert.equal(r.english, false);
    assert.deepEqual(r.chunks, [['', 80], ['row0', 80], ['row0', 10]]);
    assert.equal(r.ai.students.length, 1);
    const s = r.ai.students[0];
    assert.deepEqual([s.firstName, s.fatherName, s.lastName, s.studentClass, s.section, s.dob, s.rollNumber], ['Vedika', 'Rajkumar', 'Nehare', '6', 'A', '2014-03-30', '2']);
    assert.equal(r.ai.outOfRange, 1);
    assert.equal(r.noClass, '8');
    assert.equal(r.dupes, 1);
});
