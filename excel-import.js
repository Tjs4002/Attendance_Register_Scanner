/**
 * Excel / CSV roster import.
 * Reads every sheet of a workbook, merges them, keeps classes 6-10 only,
 * drops fathers' names and returns students shaped like the AI-extracted ones.
 * Pure functions (plus the SheetJS `XLSX` global) so they are easy to test.
 */

const EXCEL_MIN_CLASS = 6;
const EXCEL_MAX_CLASS = 10;
const ROMAN_CLASSES = { i: 1, ii: 2, iii: 3, iv: 4, v: 5, vi: 6, vii: 7, viii: 8, ix: 9, x: 10, xi: 11, xii: 12 };
const NAME_PREFIXES = /^(?:kumari|kum|ku|km|master|mast|mr|mrs|ms|miss|shri|smt|chi|कु|कुमारी|चि|श्री)\.?\s+/i;

function isExcelFile(file) {
    return /\.(xlsx|xlsm|xls|csv)$/i.test(file.name) ||
        /spreadsheet|ms-excel|text\/csv/.test(file.type);
}

function excelText(value) {
    return normalizeDigits(value).replace(/\s+/g, ' ').trim();
}

// Lower-case header with everything but letters/digits removed ("First Name *" -> "firstname").
function compactHeader(value) {
    return excelText(value).toLowerCase().replace(/[^\p{L}\p{N}\p{M}]/gu, '');
}

function detectColumn(header) {
    const h = compactHeader(header);
    if (!h) return null;
    if (/mother|आई/.test(h)) return null;
    if (/father|middle|guardian|parent|mname|वडील|पित्याचे/.test(h)) return 'middle';
    if (/first|given|fname/.test(h)) return 'first';
    if (/last|sur|family|lname|आडनाव/.test(h)) return 'last';
    if (/dob|birth|जन्म/.test(h)) return 'dob';
    if (/section|division|^div|^sec$|तुकडी/.test(h)) return 'section';
    if (/class|^std|standard|grade|वर्ग|इयत्ता/.test(h) && !/teacher/.test(h)) return 'class';
    if (/roll|^srno|^sno$|serial|अक्र/.test(h)) return 'roll'; // serial numbers only fix the ordering
    if (/admission|^admno|^adm|grno|regno|saral|दाखल/.test(h)) return 'adm';
    if (/name|नाव/.test(h) && !/school|teacher|शाळा/.test(h)) return 'full';
    return null;
}

// Find the header row (within the first 25 rows) and map roles to column indexes.
function findHeader(rows) {
    for (let r = 0; r < Math.min(rows.length, 25); r++) {
        const columns = {};
        rows[r].forEach((cell, c) => {
            const role = detectColumn(cell);
            if (role && !(role in columns)) columns[role] = c;
        });
        if (columns.full !== undefined || columns.first !== undefined) return { row: r, columns };
    }
    return null;
}

// Parses "6", "6th", "Class 6", "VI", "6-A", "6A", "सहावी" into { cls, section }.
function parseClassCell(value) {
    const text = excelText(value).toLowerCase().replace(/\b(class|std|standard|grade)\b\.?/g, ' ').replace(/[._]/g, ' ').trim();
    if (!text) return null;
    let m = text.match(/^(\d{1,2})\s*(?:st|nd|rd|th|वी|वा)?\s*[-/ ]?\s*([a-z])?$/i);
    if (m) return { cls: Number(m[1]), section: (m[2] || '').toUpperCase() };
    m = text.match(/^(xii|xi|x|ix|viii|vii|vi|v|iv|iii|ii|i)\s*[-/ ]?\s*([a-z])?$/i);
    if (m) return { cls: ROMAN_CLASSES[m[1].toLowerCase()], section: (m[2] || '').toUpperCase() };
    const word = normalizeClass(text);
    return word ? { cls: Number(word), section: '' } : null;
}

// Looser version for sheet names such as "Class 7 - A", "8th std", "IX B", "Sheet 6A".
function parseClassFromSheetName(name) {
    const direct = parseClassCell(name);
    if (direct) return direct;
    const text = excelText(name).toLowerCase();
    let m = text.match(/(?:^|[^a-z\d])(\d{1,2})\s*(?:st|nd|rd|th)?\s*[-_ ]?\s*([a-z])?(?![a-z\d])/);
    if (m) return { cls: Number(m[1]), section: (m[2] || '').toUpperCase() };
    m = text.match(/(?:^|[^a-z])(xii|xi|x|ix|viii|vii|vi|v|iv|iii|ii|i)(?![a-z])\s*[-_ ]?\s*([a-z])?(?![a-z])/);
    if (m) return { cls: ROMAN_CLASSES[m[1]], section: (m[2] || '').toUpperCase() };
    return null;
}

function validYmd(y, m, d) {
    if (y < 1990 || y > 2030 || m < 1 || m > 12 || d < 1) return '';
    if (d > new Date(Date.UTC(y, m, 0)).getUTCDate()) return '';
    return `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
}

const MONTHS = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];

function fullYear(y) {
    return y >= 100 ? y : (y <= 30 ? 2000 + y : 1900 + y);
}

// Returns YYYY-MM-DD, or '' when the value is not a recognisable date.
function parseDobCell(value) {
    if (value instanceof Date) return validYmd(value.getFullYear(), value.getMonth() + 1, value.getDate());
    if (typeof value === 'number') {
        if (value > 1000 && value < 80000 && typeof XLSX !== 'undefined') {
            const p = XLSX.SSF.parse_date_code(value);
            return p ? validYmd(p.y, p.m, p.d) : '';
        }
        value = String(value);
    }
    const text = excelText(value);
    let m = text.match(/^(\d{4})[/\-. ](\d{1,2})[/\-. ](\d{1,2})/);
    if (m) return validYmd(+m[1], +m[2], +m[3]);
    m = text.match(/^(\d{1,2})[/\-. ](\d{1,2})[/\-. ](\d{2,4})/);
    if (m) {
        let d = +m[1], mo = +m[2];
        if (mo > 12 && d <= 12) [d, mo] = [mo, d]; // tolerate US-style month/day
        return validYmd(fullYear(+m[3]), mo, d);
    }
    m = text.match(/^(\d{1,2})[\s\-/.,]*([a-z]{3,9})[\s\-/.,]*(\d{2,4})/i);
    if (m) {
        const mo = MONTHS.indexOf(m[2].slice(0, 3).toLowerCase()) + 1;
        return mo ? validYmd(fullYear(+m[3]), mo, +m[1]) : '';
    }
    m = text.match(/^(\d{2})(\d{2})(\d{4})$/); // ddmmyyyy
    if (m) return validYmd(+m[3], +m[2], +m[1]);
    return '';
}

function tidyNamePart(part) {
    const cleaned = part.replace(/[^\p{L}\p{M}'-]/gu, '');
    if (/^[A-Za-z'-]+$/.test(cleaned) && (cleaned === cleaned.toUpperCase() || cleaned === cleaned.toLowerCase())) {
        return cleaned.charAt(0).toUpperCase() + cleaned.slice(1).toLowerCase();
    }
    return cleaned;
}

function nameTokens(value) {
    let text = excelText(value).replace(/[,]/g, ' ');
    while (NAME_PREFIXES.test(text)) text = text.replace(NAME_PREFIXES, '');
    return text.split(' ').map(tidyNamePart).filter(Boolean);
}

// "Vedika Rajkumar Nehare" -> first "Vedika", father "Rajkumar", last "Nehare".
function splitStudentName({ full, first, middle, last }) {
    const fullTokens = nameTokens(full);
    const firstTokens = nameTokens(first);
    const lastTokens = nameTokens(last);
    const middleTokens = nameTokens(middle);
    let firstName = firstTokens[0];
    let lastName = lastTokens.length ? lastTokens.join(' ') : '';
    let fatherName = middleTokens.join(' ');
    if (firstTokens.length > 1 && !fatherName) fatherName = firstTokens.slice(1).join(' ');
    if (firstName === undefined) {
        firstName = fullTokens[0] || '';
        if (fullTokens.length > 1) {
            if (!lastName) lastName = fullTokens[fullTokens.length - 1];
            if (!fatherName) fatherName = fullTokens.slice(1, -1).join(' ');
        }
    }
    return { firstName, fatherName, lastName };
}

function studentsFromSheet(rows, sheetName, defaultSection) {
    const header = findHeader(rows);
    if (!header) return { students: [], skipped: 'no student-name column found' };
    const { columns } = header;
    const sheetClass = parseClassFromSheetName(sheetName);
    const cell = (row, role) => (columns[role] === undefined ? '' : row[columns[role]] ?? '');
    const students = [];
    let outOfRange = 0;

    for (const row of rows.slice(header.row + 1)) {
        const parts = {
            full: cell(row, 'full'), first: cell(row, 'first'),
            middle: cell(row, 'middle'), last: cell(row, 'last')
        };
        const name = splitStudentName(parts);
        if (!name.firstName || /^(total|grand|boys|girls)$/i.test(name.firstName)) continue;
        if (columns.full !== undefined && compactHeader(parts.full) === compactHeader(rows[header.row][columns.full])) continue;

        const classInfo = (columns.class !== undefined && parseClassCell(cell(row, 'class'))) || sheetClass;
        const cls = classInfo ? classInfo.cls : null;
        if (cls !== null && (cls < EXCEL_MIN_CLASS || cls > EXCEL_MAX_CLASS)) { outOfRange++; continue; }

        const sectionText = excelText(cell(row, 'section')).toUpperCase();
        const section = sectionText.slice(0, 2) || classInfo?.section || (sheetClass && sheetClass.section) || defaultSection;
        const rollText = excelText(cell(row, 'roll'));
        students.push({
            id: `excel_${crypto.randomUUID()}`,
            firstName: name.firstName,
            fatherName: name.fatherName,
            lastName: name.lastName,
            marathiName: '',
            dob: parseDobCell(cell(row, 'dob')),
            studentClass: cls === null ? '' : String(cls),
            section,
            rollNumber: rollText,
            admNo: excelText(cell(row, 'adm')),
            category: '',
            aadhaar: '',
            isStruckOut: false,
            imageSource: sheetName
        });
    }
    return { students, outOfRange };
}

function sortRoster(students) {
    const num = s => (/^\d+$/.test(s.rollNumber) ? Number(s.rollNumber) : Infinity);
    const text = (a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' });
    return students.sort((a, b) =>
        (Number(a.studentClass) || 99) - (Number(b.studentClass) || 99) ||
        text(a.section, b.section) ||
        num(a) - num(b) ||
        text(a.firstName, b.firstName) || text(a.lastName, b.lastName));
}

/**
 * Reads a workbook file, merges all its sheets and returns
 * { students, sheets, outOfRange, duplicates, skippedSheets }.
 */
async function importExcelFile(file, defaultSection = 'A') {
    if (typeof XLSX === 'undefined') throw new Error('Excel reader could not load. Check your internet connection and retry.');
    if (file.size > 20 * 1024 * 1024) throw new Error('Spreadsheets must be smaller than 20 MB.');
    const workbook = XLSX.read(await file.arrayBuffer(), { type: 'array' });
    const result = { students: [], sheets: workbook.SheetNames.length, outOfRange: 0, duplicates: 0, skippedSheets: [] };
    const seen = new Set();

    for (const sheetName of workbook.SheetNames) {
        const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, raw: true, defval: '' });
        const parsed = studentsFromSheet(rows, sheetName, defaultSection);
        if (parsed.skipped) { result.skippedSheets.push(sheetName); continue; }
        result.outOfRange += parsed.outOfRange;
        for (const student of parsed.students) {
            // Same child listed on two sheets of one workbook is imported once.
            const key = student.dob && [student.firstName, student.lastName, student.dob, student.studentClass].join('|').toLowerCase();
            if (key && seen.has(key)) { result.duplicates++; continue; }
            if (key) seen.add(key);
            result.students.push(student);
        }
    }
    sortRoster(result.students);
    return result;
}
