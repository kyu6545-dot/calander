const CONFIG = {
  PLAN_SHEET: 'Plan',
  MONTHLY_SHEET: 'MonthlyView',
  DEFAULT_CAL_NAME: 'Work Plan',
  TZ: 'Asia/Seoul',
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('WorkPlan')
    .addItem('Setup Template', 'setupTemplate_')
    .addItem('Sync', 'syncPlanToCalendar')
    .addItem('Export Monthly PDF', 'exportMonthlyPdf_')
    .addToUi();
}

function setupTemplate_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(CONFIG.PLAN_SHEET);
  if (!sh) sh = ss.insertSheet(CONFIG.PLAN_SHEET);

  const header = [
    'Date',
    'Start',
    'End',
    'Org',
    'Owner',
    'Type',
    'Title',
    'Priority',
    'Notes',
    'Status',
    'CalendarName',
    'EventId',
  ];
  sh.clear();
  sh.getRange(1, 1, 1, header.length).setValues([header]);
  sh.setFrozenRows(1);

  const orgRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['KR', 'US'], true)
    .build();
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['PLAN', 'UPDATE', 'DONE'], true)
    .build();

  sh.getRange(2, 4, sh.getMaxRows() - 1, 1).setDataValidation(orgRule);
  sh.getRange(2, 10, sh.getMaxRows() - 1, 1).setDataValidation(statusRule);
}

function syncPlanToCalendar() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CONFIG.PLAN_SHEET);
  if (!sh) throw new Error(`Sheet not found: ${CONFIG.PLAN_SHEET}`);

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return;

  const header = values[0].map((h) => String(h).trim());
  const idx = Object.fromEntries(header.map((h, i) => [h, i]));

  const must = ['Date', 'Org', 'Owner', 'Title', 'Status', 'CalendarName', 'EventId'];
  must.forEach((k) => {
    if (idx[k] === undefined) throw new Error(`Missing column: ${k}`);
  });

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const dateStr = str_(row[idx.Date]);
    const startStr = str_(row[idx.Start]);
    const endStr = str_(row[idx.End]);
    const org = normalizeOrg_(str_(row[idx.Org]));
    const owner = str_(row[idx.Owner]);
    const type = str_(row[idx.Type]);
    const title = str_(row[idx.Title]);
    const priority = str_(row[idx.Priority]);
    const notes = str_(row[idx.Notes]);
    const status = str_(row[idx.Status]).toUpperCase();
    const calName = str_(row[idx.CalendarName]) || CONFIG.DEFAULT_CAL_NAME;
    const eventId = str_(row[idx.EventId]);

    if (!status) continue;
    if (!org || !owner || !title) {
      Logger.log(`Row ${r + 1}: missing org/owner/title`);
      continue;
    }

    const cal = getOrCreateCalendar_(calName);

    const eventTitle = `[${org}] ${owner} - ${title}`;
    const description = [
      `Type: ${type}`,
      `Priority: ${priority}`,
      `Notes: ${notes}`,
      `SourceRow: ${r + 1}`,
      `Sheet: ${CONFIG.PLAN_SHEET}`,
    ].join('\n');

    try {
      if (status === 'DONE') {
        if (eventId) {
          const ev = safeGetEventById_(cal, eventId);
          if (ev) ev.deleteEvent();
          sh.getRange(r + 1, idx.EventId + 1).setValue('');
        }
        continue;
      }

      if (status === 'PLAN') {
        if (eventId) continue;
        const { start, end, isAllDay } = parseWhen_(dateStr, startStr, endStr);
        const ev = isAllDay
          ? cal.createAllDayEvent(eventTitle, start, { description })
          : cal.createEvent(eventTitle, start, end, { description });
        sh.getRange(r + 1, idx.EventId + 1).setValue(ev.getId());
        continue;
      }

      if (status === 'UPDATE') {
        if (!eventId) {
          Logger.log(`Row ${r + 1}: UPDATE but no EventId`);
          continue;
        }
        const ev = safeGetEventById_(cal, eventId);
        if (!ev) {
          Logger.log(`Row ${r + 1}: event not found for EventId ${eventId}`);
          continue;
        }

        const { start, end, isAllDay } = parseWhen_(dateStr, startStr, endStr);
        ev.setTitle(eventTitle);
        ev.setDescription(description);
        if (isAllDay) {
          ev.setAllDayDate(start);
        } else {
          ev.setTime(start, end);
        }
      }
    } catch (e) {
      Logger.log(`Row ${r + 1}: error ${e}`);
    }
  }
}

function exportMonthlyPdf_() {
  const ui = SpreadsheetApp.getUi();
  const prompt = ui.prompt('Export Monthly PDF', 'Enter month (YYYY-MM)', ui.ButtonSet.OK_CANCEL);
  if (prompt.getSelectedButton() !== ui.Button.OK) return;
  const monthValue = String(prompt.getResponseText() || '').trim();
  if (!/^\d{4}-\d{2}$/.test(monthValue)) {
    ui.alert('Invalid month. Use YYYY-MM format.');
    return;
  }

  try {
    const ss = SpreadsheetApp.getActive();
    const plan = ss.getSheetByName(CONFIG.PLAN_SHEET);
    if (!plan) throw new Error(`Sheet not found: ${CONFIG.PLAN_SHEET}`);

    const rows = plan.getDataRange().getValues();
    if (rows.length < 2) throw new Error('No plan data to export.');

    const header = rows[0].map((h) => String(h).trim());
    const idx = Object.fromEntries(header.map((h, i) => [h, i]));

    const statusIndex = idx.Status;
    const dateIndex = idx.Date;

    const filtered = rows
      .slice(1)
      .map((row, index) => ({ row, index: index + 2 }))
      .filter(({ row }) => {
        const status = str_(row[statusIndex]).toUpperCase();
        if (!(status === 'PLAN' || status === 'UPDATE')) return false;
        const dateStr = str_(row[dateIndex]);
        return dateStr.startsWith(`${monthValue}-`);
      })
      .map(({ row }) => row);

    let monthly = ss.getSheetByName(CONFIG.MONTHLY_SHEET);
    if (!monthly) monthly = ss.insertSheet(CONFIG.MONTHLY_SHEET);
    monthly.clear();
    monthly.getRange(1, 1, 1, header.length).setValues([header]);

    if (filtered.length) {
      monthly.getRange(2, 1, filtered.length, header.length).setValues(filtered);
    }
    monthly.setFrozenRows(1);

    const sheetId = monthly.getSheetId();
    const exportUrl = [
      `https://docs.google.com/spreadsheets/d/${ss.getId()}/export`,
      '?format=pdf',
      '&portrait=true',
      '&fitw=true',
      '&sheetnames=false',
      '&printtitle=false',
      '&pagenumbers=false',
      '&gridlines=false',
      '&fzr=false',
      `&gid=${sheetId}`,
    ].join('');

    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(exportUrl, {
      headers: { Authorization: `Bearer ${token}` },
      muteHttpExceptions: true,
    });

    if (response.getResponseCode() !== 200) {
      throw new Error(`PDF export failed: ${response.getContentText()}`);
    }

    const fileName = `WorkPlan_${monthValue}.pdf`;
    const blob = response.getBlob().setName(fileName);
    DriveApp.createFile(blob);
    ui.alert(`PDF saved to Drive: ${fileName}`);
  } catch (e) {
    Logger.log(`Export error: ${e}`);
    ui.alert(`Export failed: ${e.message || e}`);
  }
}

function getOrCreateCalendar_(name) {
  const cals = CalendarApp.getCalendarsByName(name);
  if (cals && cals.length) return cals[0];
  return CalendarApp.createCalendar(name);
}

function safeGetEventById_(cal, eventId) {
  try {
    return cal.getEventById(eventId) || null;
  } catch (e) {
    return null;
  }
}

function parseWhen_(dateStr, startStr, endStr) {
  if (!dateStr) throw new Error('Missing Date');
  const [y, m, d] = dateStr.split('-').map(Number);
  if (!(y && m && d)) throw new Error(`Invalid Date: ${dateStr}`);

  const allDay = !startStr && !endStr;
  if (allDay) {
    const day = new Date(y, m - 1, d);
    return { start: day, end: day, isAllDay: true };
  }

  const start = parseDateTime_(y, m, d, startStr);
  const end = parseDateTime_(y, m, d, endStr);
  if (end.getTime() <= start.getTime()) throw new Error('End <= Start');
  return { start, end, isAllDay: false };
}

function parseDateTime_(y, m, d, hm) {
  const [hh, mm] = hm.split(':').map(Number);
  if (Number.isNaN(hh) || Number.isNaN(mm)) throw new Error(`Invalid time: ${hm}`);
  return new Date(y, m - 1, d, hh, mm, 0);
}

function normalizeOrg_(org) {
  const v = (org || '').toUpperCase();
  if (v === 'HQ') return 'KR';
  if (v === 'KR' || v === 'US') return v;
  return '';
}

function str_(v) {
  return v === null || v === undefined ? '' : String(v).trim();
}
