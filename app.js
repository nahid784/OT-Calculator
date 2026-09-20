const state = {
  rows: [],
  basicSalary: 0,
  showBonus: false,
};

// ot logic

const fileInput = document.getElementById('attendance-file');
const basicSalaryInput = document.getElementById('basic-salary');
const totalOtHoursEl = document.getElementById('total-ot-hours');
const totalOtBonusEl = document.getElementById('total-ot-bonus');
const otRateEl = document.getElementById('ot-rate');
const rowsParsedEl = document.getElementById('rows-parsed');
const eligibleDaysEl = document.getElementById('eligible-days');
const attendanceBody = document.getElementById('attendance-body');
const toggleBonusBtn = document.getElementById('toggle-bonus');
const recalcBtn = document.getElementById('recalculate-btn');

function formatHours(value) {
  if (!Number.isFinite(value) || value <= 0) return '0.00';
  return Number(value).toFixed(2);
}

function formatTaka(value) {
  const amount = Number(value) || 0;
  return `৳ ${amount.toLocaleString('en-BD', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
}

function safeNumber(value) {
  if (value === null || value === undefined || value === '') {
    return 0;
  }

  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : 0;
  }

  const cleaned = String(value)
    .replace(/,/g, '')
    .replace(/[^0-9.\-]/g, '');

  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : 0;
}

function normalizeStatus(raw) {
  const value = String(raw || '').trim().toLowerCase();
  if (!value) return 'unknown';

  if (value.includes('holiday')) return 'holiday';
  if (value.includes('absent') || value.includes('not present')) return 'absent';
  if (value.includes('late') && value.includes('early')) return 'late-early';
  if (value.includes('late')) return 'late';
  if (value.includes('early')) return 'early';
  if (value.includes('present') || value.includes('intime') || value.includes('outtime')) return 'present';
  return 'other';
}

function formatStatus(status) {
  const key = normalizeStatus(status);
  const map = {
    present: 'Present',
    late: 'Late In',
    early: 'Early Out',
    'late-early': 'Late In + Early Out',
    holiday: 'Holiday',
    absent: 'Absent',
    other: 'Other',
    unknown: 'Unknown',
  };
  return map[key] || 'Unknown';
}

function statusClass(status) {
  const key = normalizeStatus(status);
  const map = {
    present: 'status-present',
    late: 'status-present',
    early: 'status-present',
    'late-early': 'status-present',
    holiday: 'status-holiday',
    absent: 'status-absent',
    other: 'status-other',
    unknown: 'status-other',
  };
  return map[key] || 'status-other';
}

function timeToMinutes(value) {
  if (!value) return 0;

  if (typeof value === 'number') {
    const totalDays = value % 1;
    const minutes = totalDays * 24 * 60;
    return Math.round(minutes);
  }

  const text = String(value).trim();

  if (!text) return 0;

  const lower = text.toLowerCase();
  const match = lower.match(/(\d{1,2})(?::|\.)(\d{1,2})(?::|\.)(\d{1,2})\s*(am|pm)?/i)
    || lower.match(/(\d{1,2})(?::|\.)(\d{1,2})\s*(am|pm)?/i);

  if (!match) {
    const decimal = Number(text);
    if (Number.isFinite(decimal)) {
      return Math.round((decimal % 1) * 24 * 60);
    }
    return 0;
  }

  let hours = Number(match[1]);
  let minutes = Number(match[2] || 0);
  const sec = Number(match[3] || 0);
  const meridian = (match[4] || '').toLowerCase();

  if (meridian === 'pm' && hours < 12) hours += 12;
  if (meridian === 'am' && hours === 12) hours = 0;

  return hours * 60 + minutes + Math.round(sec / 60);
}

function jsDateFromValue(value) {
  if (!value) return null;

  if (value instanceof Date) return value;

  if (typeof value === 'number') {
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    return new Date(excelEpoch.getTime() + value * 86400000);
  }

  if (typeof value === 'string') {
    const trimmed = value.trim();
    const directDate = new Date(trimmed);
    if (!Number.isNaN(directDate.getTime())) return directDate;

    const datetime = trimmed.match(/(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})\s*(\d{1,2}:\d{2}(?::\d{2})?\s*(?:AM|PM|am|pm)?)?/i);
    if (datetime) {
      const parsed = new Date(datetime[1]);
      if (datetime[2]) {
        const time = new Date(`2000-01-01 ${datetime[2]}`);
        parsed.setHours(time.getHours(), time.getMinutes(), time.getSeconds());
      }
      return parsed;
    }

    const timeOnly = trimmed.match(/(\d{1,2}:\d{2}(?::\d{2})?\s*(?:AM|PM|am|pm)?)/i);
    if (timeOnly) {
      const fake = new Date('2000-01-01');
      const d = new Date(`2000-01-01 ${timeOnly[1]}`);
      fake.setHours(d.getHours(), d.getMinutes(), d.getSeconds());
      return fake;
    }
  }

  return null;
}

function getCell(row, aliases) {
  const keys = Object.keys(row || {});
  for (const alias of aliases) {
    const match = keys.find((k) => String(k).trim().toLowerCase() === alias.toLowerCase());
    if (match) return row[match];
  }
  return '';
}

function parseWorkbookFile(file) {
  const reader = new FileReader();

  reader.onload = function (event) {
    try {
      const data = event.target.result;
      const workbook = XLSX.read(data, { type: 'array', cellDates: true });
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(firstSheet, { defval: '', raw: false });
      const parsedRows = rows
        .map((entry) => normalizeAttendanceRow(entry))
        .filter((entry) => Object.keys(entry).length > 0);

      state.rows = parsedRows;
      renderTable();
      updateSummary();
    } catch (error) {
      console.error(error);
      attendanceBody.innerHTML = `<tr><td colspan="7" class="empty-state">Unable to parse the file. Please upload a valid Excel or CSV attendance sheet.</td></tr>`;
    }
  };

  reader.readAsArrayBuffer(file);
}

async function parsePdfFile(file) {
  if (!window.pdfjsLib) {
    throw new Error('PDF parser is unavailable. Check the internet connection and reload the page.');
  }

  pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
  const data = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data }).promise;
  const records = [];
  let currentRecord = null;
  let previousDateColumnX = null;

  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber);
    const content = await page.getTextContent();
    const items = content.items
      .map((item) => ({
        text: String(item.str || '').trim(),
        x: item.transform[4],
        y: item.transform[5],
      }))
      .filter((item) => item.text)
      .sort((first, second) => second.y - first.y || first.x - second.x);

    const dateHeader = items.find((item) => item.text.toUpperCase() === 'DATE');
    const dateColumnX = dateHeader ? dateHeader.x : previousDateColumnX;
    if (dateColumnX !== null && dateColumnX !== undefined) {
      previousDateColumnX = dateColumnX;
    }

    const baselines = new Map();
    for (const item of items) {
      const baseline = Math.round(item.y * 100) / 100;
      const group = baselines.get(baseline) || [];
      group.push(item);
      baselines.set(baseline, group);
    }

    for (const group of baselines.values()) {
      if (group.some((item) => isPdfDateStart(item, dateColumnX))) {
        if (currentRecord) records.push(currentRecord);
        currentRecord = [];
      }

      if (currentRecord) {
        group.forEach((item) => {
          if (!isPdfHeaderItem(item.text)) currentRecord.push(item);
        });
      }
    }
  }

  if (currentRecord) records.push(currentRecord);

  return records
    .map((record) => parsePdfAttendanceLine(record))
    .filter((row) => row && row.date && row.completed !== '');
}

function isPdfHeaderItem(text) {
  return /^(attendance report|emp id|name|designation|dept|date|start|time|in|end|out|status|working|hours|completed)$/i.test(text);
}

function isPdfDateStart(item, dateColumnX) {
  const text = item.text;
  const isDateText = /^\d{1,2}-$/.test(text)
    || /^\d{1,2}[-/]\w{3}[-/]\d{2,4}$/i.test(text)
    || /^\d{1,2}[-/]\d{1,2}[-/]\d{2,4}$/.test(text);

  return isDateText
    && (dateColumnX === null || dateColumnX === undefined || Math.abs(item.x - dateColumnX) < 25);
}

function parsePdfAttendanceLine(line) {
  const recordItems = Array.isArray(line) ? line : null;
  const lineText = recordItems ? recordItems.map((item) => item.text).join(' ') : line;
  const dateMatch = lineText.match(/\b(\d{1,2})-\s*([A-Za-z]{3})-(\d{2,4})\b/i)
    || lineText.match(/\b\d{1,2}[-/]\d{1,2}[-/]\d{2,4}\b/);
  const clockMatches = [...lineText.matchAll(/\d{1,2}:\d{2}(?::\d{2})?/g)];
  const meridianMatches = [...lineText.matchAll(/\b(?:AM|PM)\b/gi)];
  const timeMatches = clockMatches.map((match, index) => ({
    index: match.index,
    0: `${match[0]}${meridianMatches[index] ? ` ${meridianMatches[index][0]}` : ''}`,
  }));
  if (!dateMatch || timeMatches.length < 1) return null;

  const statusMatch = lineText.match(/holiday|medical\s+leave|outtime\s+missing|late\s+in|early\s+out|present|absent/ig);
  const inTimeMatch = timeMatches.length >= 2 ? timeMatches[1] : timeMatches[0];
  const outTimeMatch = timeMatches.length >= 4 ? timeMatches[3] : null;

  const date = dateMatch[1]
    ? `${dateMatch[1]}-${dateMatch[2]}-${dateMatch[3]}`
    : dateMatch[0];
  let workingHours = '';
  let completed = '';

  if (recordItems) {
    const isNumericCell = (item) => /^-?\d+(?:\.\d+)?$/.test(item.text);
    const workingCell = recordItems.find((item) => item.x >= 660 && item.x < 720 && isNumericCell(item));
    const completedCell = recordItems.find((item) => item.x >= 720 && isNumericCell(item));
    workingHours = workingCell ? workingCell.text : '';
    completed = completedCell ? completedCell.text : '';
  } else {
    const afterStatus = statusMatch
      ? lineText.slice(statusMatch.index + statusMatch[0].length)
      : lineText;
    const numericValues = afterStatus.match(/-?\d+(?:\.\d+)?/g) || [];
    if (numericValues.length < 2) return null;
    workingHours = numericValues[0];
    completed = numericValues[1];
  }

  if (completed === '') return null;

  return {
    employeeId: '',
    name: '',
    date,
    startTime: '',
    inTime: `${date} ${inTimeMatch[0]}`,
    endTime: '',
    outTime: outTimeMatch ? `${date} ${outTimeMatch[0]}` : '',
    status: statusMatch ? statusMatch[0] : 'Present',
    workingHours,
    completed,
    normalizedStatus: normalizeStatus(statusMatch ? statusMatch[0] : 'Present'),
  };
}

function normalizeAttendanceRow(entry) {
  const row = {};
  row.employeeId = getCell(entry, ['EMP ID', 'Employee ID', 'Emp ID', 'ID']);
  row.name = getCell(entry, ['NAME', 'Employee Name', 'Name']);
  row.date = getCell(entry, ['DATE', 'Date']);
  row.startTime = getCell(entry, ['START TIME', 'Start Time', 'Time In', 'Shift Start']);
  row.inTime = getCell(entry, ['IN TIME', 'In Time', 'Punch In', 'Check In']);
  row.endTime = getCell(entry, ['END TIME', 'End Time', 'Shift End']);
  row.outTime = getCell(entry, ['OUT TIME', 'Out Time', 'Punch Out', 'Check Out']);
  row.status = getCell(entry, ['STATUS', 'Status']);
  row.workingHours = getCell(entry, ['WORKING HOURS', 'Working Hours', 'Hours']);
  row.completed = getCell(entry, ['COMPLETED', 'Completed', 'Completed Hours', 'Total Hours']);

  const statusValue = row.status || '';
  row.normalizedStatus = normalizeStatus(statusValue);
  return row;
}

function computeBreakTimeMinutes(startTimeValue, endTimeValue) {
  const start = jsDateFromValue(startTimeValue);
  const end = jsDateFromValue(endTimeValue);
  if (!start || !end) return 0;

  const startMinutes = (start.getHours() * 60) + start.getMinutes();
  const endMinutes = (end.getHours() * 60) + end.getMinutes();

  let total = 0;
  const breakWindows = [
    { start: 17 * 60 + 30, end: 18 * 60 },
    { start: 21 * 60 + 30, end: 22 * 60 },
  ];

  const otStart = Math.max(startMinutes, 16 * 60);
  const otEnd = Math.max(endMinutes, otStart);

  for (const window of breakWindows) {
    const overlapStart = Math.max(otStart, window.start);
    const overlapEnd = Math.min(otEnd, window.end);
    if (overlapEnd > overlapStart) {
      total += overlapEnd - overlapStart;
    }
  }

  return total;
}

function getHolidayBreakDeduction(row) {
  const date = jsDateFromValue(row.date);
  const isFriday = !!date && date.getDay() === 5;
  const endMinutes = Math.max(timeToMinutes(row.outTime || row.endTime), timeToMinutes(row.inTime || row.startTime));

  if (isFriday && endMinutes <= 16 * 60 + 30) return 1.5;
  if (!isFriday && endMinutes <= 16 * 60 + 30) return 1.0;
  if (endMinutes <= 19 * 60 + 30) return 1.5;
  return 2.0;
}

function getHolidayCalculationHours(row) {
  const startMinutes = timeToMinutes(row.inTime || row.startTime);
  const endMinutes = timeToMinutes(row.outTime || row.endTime);

  if (!startMinutes || !endMinutes || endMinutes <= startMinutes) {
    return safeNumber(row.completed);
  }

  const effectiveStart = Math.max(startMinutes, 7 * 60);
  return Math.max(0, (endMinutes - effectiveStart) / 60);
}

function floorQuarter(value) {
  if (!Number.isFinite(value)) return 0;
  return Math.floor(value * 4) / 4;
}

function calculateRowOt(row) {
  const completed = safeNumber(row.completed);
  const status = normalizeStatus(row.status);

  if (!row.inTime) return 0;

  if (status === 'holiday') {
    const holidayBreakDeduction = getHolidayBreakDeduction(row);
    const holidayCompleted = getHolidayCalculationHours(row);
    const adjustedOt = Math.max(0, holidayCompleted - holidayBreakDeduction);
    return floorQuarter(adjustedOt);
  }

  if (completed < 12) return 0;

  const baseHours = Math.max(0, completed - 9);
  const breakMinutes = computeBreakTimeMinutes(row.inTime || row.startTime, row.outTime || row.endTime);
  const breakHours = breakMinutes / 60;
  const adjustedOt = Math.max(0, baseHours - breakHours);
  return floorQuarter(adjustedOt);
}

function calculateBonusForRow(otHours, basicSalary) {
  if (!basicSalary || otHours <= 0) return 0;
  const otRate = basicSalary / 104;
  return otHours * otRate;
}

function updateSummary() {
  const basicSalary = Number(basicSalaryInput.value || 0);
  state.basicSalary = basicSalary;

  const otRate = basicSalary > 0 ? basicSalary / 104 : 0;
  const totalOtHours = state.rows.reduce((sum, row) => {
    const hours = calculateRowOt(row);
    return sum + hours;
  }, 0);
  const totalOtBonus = totalOtHours * otRate;

  totalOtHoursEl.textContent = formatHours(totalOtHours);
  totalOtBonusEl.textContent = formatTaka(totalOtBonus);
  otRateEl.textContent = formatTaka(otRate);
  rowsParsedEl.textContent = String(state.rows.length);

  const eligibleRows = state.rows.filter((row) => calculateRowOt(row) > 0);
  eligibleDaysEl.textContent = String(eligibleRows.length);

  const bonusCard = document.querySelector('.bonus-card');
  bonusCard.classList.toggle('revealed', state.showBonus);
  totalOtBonusEl.style.filter = state.showBonus ? 'blur(0)' : 'blur(6px)';
}

function renderTable() {
  if (!state.rows.length) {
    attendanceBody.innerHTML = '<tr><td colspan="7" class="empty-state">Upload an attendance report to begin.</td></tr>';
    return;
  }

  const basicSalary = Number(basicSalaryInput.value || 0);
  const sortedRows = [...state.rows].filter((row) => {
    const status = normalizeStatus(row.status);
    return status !== 'unknown' || row.completed || row.date;
  });

  attendanceBody.innerHTML = sortedRows
    .map((row) => {
      const otHours = calculateRowOt(row);
      const shouldShow = otHours > 0;

      if (!shouldShow) {
        return '';
      }

      return `
        <tr>
          <td>${row.date || '-'}</td>
          <td><span class="status-pill ${statusClass(row.status)}">${formatStatus(row.status)}</span></td>
          <td>${row.inTime || '-'}</td>
          <td>${row.outTime || '-'}</td>
          <td>${row.workingHours || '-'}</td>
          <td>${row.completed || '-'}</td>
          <td class="ot-value ${otHours === 0 ? 'ot-zero' : ''}">${formatHours(otHours)}</td>
        </tr>
      `;
    })
    .join('');

  if (!attendanceBody.innerHTML.trim()) {
    attendanceBody.innerHTML = '<tr><td colspan="7" class="empty-state">No overtime rows were found for the uploaded attendance report.</td></tr>';
  }

  updateSummary();
}

fileInput.addEventListener('change', (event) => {
  const file = event.target.files && event.target.files[0];
  if (!file) return;

  const extension = file.name.split('.').pop().toLowerCase();
  if (['xlsx', 'xls', 'csv'].includes(extension)) {
    parseWorkbookFile(file);
    return;
  }

  if (extension === 'pdf') {
    parsePdfFile(file)
      .then((parsedRows) => {
        state.rows = parsedRows;
        renderTable();
        updateSummary();
      })
      .catch((error) => {
        console.error(error);
        attendanceBody.innerHTML = `<tr><td colspan="7" class="empty-state">Unable to parse this PDF. Please upload a text-based attendance PDF or use Excel/CSV.</td></tr>`;
      });
  }
});

basicSalaryInput.addEventListener('input', () => {
  updateSummary();
  renderTable();
});

toggleBonusBtn.addEventListener('click', () => {
  state.showBonus = !state.showBonus;
  toggleBonusBtn.textContent = state.showBonus ? 'Hide' : 'Display';
  updateSummary();
});

recalcBtn.addEventListener('click', () => {
  renderTable();
});

updateSummary();
renderTable();
