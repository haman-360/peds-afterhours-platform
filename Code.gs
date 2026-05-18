const CONFIG = {
  spreadsheetIdProperty: 'SPREADSHEET_ID',
  doctorEmailProperty: 'DOCTOR_EMAIL',
  receptionPinProperty: 'RECEPTION_PIN',
  sheetName: 'consultations',
  authorizedPatientsSheetName: 'authorized_patients',
  authorizedPatientsSheetNameAliases: ['authorized_paptients'],
  defaultStatus: '未対応',
  timezone: 'Asia/Tokyo'
};

const HEADERS = [
  '受付日時',
  '患者ID',
  '相談者続柄',
  '返信用電話番号',
  '主訴',
  '発症日時',
  '体温',
  '呼吸状態',
  '水分摂取',
  '尿回数',
  '意識状態',
  'けいれん',
  '発疹',
  '嘔吐',
  '下痢',
  '服薬状況',
  '保護者の希望',
  '緊急フラグ',
  '対応状況',
  '医師メモ',
  'カルテ転記済み'
];

const AUTHORIZED_PATIENT_HEADERS = [
  '患者ID',
  '相談コード',
  '有効',
  'メモ',
  '有効化日時',
  '有効化担当'
];

function doGet(e) {
  const page = e && e.parameter && e.parameter.page;
  const pageSettings = {
    hospitals: {
      templateName: 'hospital',
      title: '泉州地域 小児救急輪番病院'
    },
    reception: {
      templateName: 'reception',
      title: '相談コード交付 受付管理'
    }
  };
  const selectedPage = pageSettings[page] || {
    templateName: 'index',
    title: '小児かかりつけ夜間休日相談'
  };

  return HtmlService
    .createTemplateFromFile(selectedPage.templateName)
    .evaluate()
    .setTitle(selectedPage.title)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    const payload = parsePostPayload_(e);
    const result = handleConsultationSubmission(payload);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    console.error(error);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, message: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function submitConsultation(formData) {
  return handleConsultationSubmission(formData);
}

function handleConsultationSubmission(formData) {
  const normalized = normalizeFormData_(formData);
  validateRequiredFields_(normalized);
  validateAuthorizedPatient_(normalized);

  const emergency = judgeEmergencyFlag(normalized);
  const aiDraft = createAiSummaryDraft(normalized);
  const receivedAt = new Date();

  saveConsultationToSheet(normalized, emergency, receivedAt);
  notifyDoctor(normalized, emergency, receivedAt, aiDraft);

  return {
    receivedAt: Utilities.formatDate(receivedAt, CONFIG.timezone, 'yyyy/MM/dd HH:mm'),
    emergencyFlag: emergency.flag,
    emergencyReasons: emergency.reasons
  };
}

function saveConsultationToSheet(data, emergency, receivedAt) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sheet = getConsultationSheet_();
    ensureHeaderRow_(sheet);
    ensureColumnFormats_(sheet);

    sheet.appendRow([
      receivedAt,
      data.patientId,
      data.relationship,
      data.phone,
      data.chiefComplaint,
      data.onsetAt,
      data.temperature,
      data.breathing,
      data.hydration,
      data.urination,
      data.consciousness,
      data.convulsion,
      data.rash,
      data.vomiting,
      data.diarrhea,
      data.medication,
      data.parentRequest,
      emergency.flag ? '要注意' : '通常',
      CONFIG.defaultStatus,
      '',
      false
    ]);
  } finally {
    lock.releaseLock();
  }
}

function notifyDoctor(data, emergency, receivedAt, aiDraft) {
  const doctorEmail = getDoctorEmail_();
  if (!doctorEmail) {
    console.warn('DOCTOR_EMAIL is not configured. Email notification skipped.');
    return;
  }

  const receivedAtText = Utilities.formatDate(receivedAt, CONFIG.timezone, 'yyyy/MM/dd HH:mm');
  const subjectPrefix = emergency.flag ? '【要注意】' : '【相談】';
  const subject = `${subjectPrefix}小児かかりつけ夜間休日相談 患者ID:${data.patientId}`;
  const body = [
    '小児かかりつけ夜間休日相談フォームに送信がありました。',
    '',
    `受付日時: ${receivedAtText}`,
    `緊急フラグ: ${emergency.flag ? '要注意' : '通常'}`,
    `判定理由: ${emergency.reasons.length ? emergency.reasons.join(' / ') : '該当なし'}`,
    '',
    `患者ID: ${data.patientId}`,
    `相談者続柄: ${data.relationship}`,
    `返信用電話番号: ${data.phone}`,
    `主訴: ${data.chiefComplaint}`,
    `発症日時: ${data.onsetAt}`,
    `体温: ${data.temperature}`,
    `呼吸状態: ${data.breathing}`,
    `水分摂取: ${data.hydration}`,
    `尿回数: ${data.urination}`,
    `意識状態: ${data.consciousness}`,
    `けいれん: ${data.convulsion}`,
    `発疹: ${data.rash}`,
    `嘔吐: ${data.vomiting}`,
    `下痢: ${data.diarrhea}`,
    `服薬状況: ${data.medication}`,
    `保護者の希望: ${data.parentRequest}`,
    '',
    `AI要約欄: ${aiDraft.summary}`,
    '',
    'このフォームは緊急通報ではありません。状態悪化時は119または救急受診の案内を優先してください。'
  ].join('\n');

  MailApp.sendEmail({
    to: doctorEmail,
    subject,
    body
  });
}

function judgeEmergencyFlag(data) {
  const reasons = [];
  const temp = Number(String(data.temperature).replace(/[^\d.]/g, ''));
  const emergencyBreathing = ['肩で息をしている', '唇が紫'];
  const emergencyConsciousness = ['ぐったりしている', '呼びかけに反応しにくい'];
  const emergencyConvulsion = ['あり', '現在も続いている'];

  if (emergencyBreathing.indexOf(data.breathing) !== -1) {
    reasons.push(`呼吸状態: ${data.breathing}`);
  }

  if (emergencyConsciousness.indexOf(data.consciousness) !== -1) {
    reasons.push(`意識状態: ${data.consciousness}`);
  }

  if (emergencyConvulsion.indexOf(data.convulsion) !== -1) {
    reasons.push(`けいれん: ${data.convulsion}`);
  }

  if (data.hydration === 'ほとんど取れない' || data.urination === '半日以上なし') {
    reasons.push('脱水リスク');
  }

  if (data.vomiting === '繰り返している' || data.diarrhea === '血便あり') {
    reasons.push('消化器症状の要注意項目');
  }

  if (!Number.isNaN(temp) && temp >= 40) {
    reasons.push(`高体温: ${data.temperature}`);
  }

  const complaintText = `${data.chiefComplaint} ${data.parentRequest}`;
  const urgentWords = ['呼吸困難', '意識がない', '顔色が悪い', 'チアノーゼ', 'ぐったり', 'けいれん', '痙攣', '止まらない'];
  urgentWords.forEach((word) => {
    if (complaintText.indexOf(word) !== -1) {
      reasons.push(`自由記載: ${word}`);
    }
  });

  return {
    flag: reasons.length > 0,
    reasons: Array.from(new Set(reasons))
  };
}

function createAiSummaryDraft(data) {
  return {
    summary: '未実装',
    source: buildAiSummaryInput(data)
  };
}

function buildAiSummaryInput(data) {
  return {
    patientId: data.patientId,
    chiefComplaint: data.chiefComplaint,
    onsetAt: data.onsetAt,
    temperature: data.temperature,
    breathing: data.breathing,
    hydration: data.hydration,
    urination: data.urination,
    consciousness: data.consciousness,
    convulsion: data.convulsion,
    rash: data.rash,
    vomiting: data.vomiting,
    diarrhea: data.diarrhea,
    medication: data.medication,
    parentRequest: data.parentRequest
  };
}

function parsePostPayload_(e) {
  if (!e) return {};

  if (e.postData && e.postData.contents) {
    const contentType = e.postData.type || '';
    if (contentType.indexOf('application/json') !== -1) {
      return JSON.parse(e.postData.contents);
    }
  }

  return e.parameter || {};
}

function normalizeFormData_(formData) {
  const data = formData || {};
  return {
    patientId: normalizePatientId_(data.patientId),
    consultationCode: normalizeConsultationCode_(data.consultationCode),
    relationship: cleanText_(data.relationship),
    phone: cleanText_(data.phone),
    chiefComplaint: cleanText_(data.chiefComplaint),
    onsetAt: cleanText_(data.onsetAt),
    temperature: cleanText_(data.temperature),
    breathing: cleanText_(data.breathing),
    hydration: cleanText_(data.hydration),
    urination: cleanText_(data.urination),
    consciousness: cleanText_(data.consciousness),
    convulsion: cleanText_(data.convulsion),
    rash: cleanText_(data.rash),
    vomiting: cleanText_(data.vomiting),
    diarrhea: cleanText_(data.diarrhea),
    medication: cleanText_(data.medication),
    parentRequest: cleanText_(data.parentRequest)
  };
}

function validateRequiredFields_(data) {
  const required = {
    patientId: '患者ID',
    consultationCode: 'かかりつけ相談コード',
    relationship: '相談者続柄',
    phone: '返信用電話番号',
    chiefComplaint: '主訴',
    breathing: '呼吸状態',
    hydration: '水分摂取',
    urination: '尿回数',
    consciousness: '意識状態',
    convulsion: 'けいれん',
    vomiting: '嘔吐',
    diarrhea: '下痢',
    parentRequest: '保護者の希望'
  };

  const missing = Object.keys(required).filter((key) => !data[key]);
  if (missing.length) {
    throw new Error(`未入力の項目があります: ${missing.map((key) => required[key]).join('、')}`);
  }

  validateInputFormats_(data);
}

function validateInputFormats_(data) {
  if (!/^\d{5}$/.test(data.patientId)) {
    throw new Error('患者IDは、先頭の0を含めて5桁の数字で入力してください。');
  }

  if (!/^[A-Z0-9]{4}-[A-Z0-9]{4}-[A-Z0-9]{4}$/.test(data.consultationCode)) {
    throw new Error('かかりつけ相談コードは、XXXX-XXXX-XXXXの形式で入力してください。');
  }
}

function validateAuthorizedPatient_(data) {
  const sheet = getAuthorizedPatientsSheet_();
  ensureAuthorizedPatientsHeaderRow_(sheet);
  ensureAuthorizedPatientsColumnFormats_(sheet);
  const headerMap = getAuthorizedPatientsHeaderMap_(sheet);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    throw new Error('かかりつけ相談コードの登録がまだありません。医院へお問い合わせください。');
  }

  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const matched = values.some((row) => {
    const patientId = normalizePatientId_(row[headerMap['患者ID']]);
    const consultationCode = normalizeConsultationCode_(row[headerMap['相談コード']]);
    const enabled = isEnabledValue_(row[headerMap['有効']]);

    return enabled &&
      patientId === data.patientId &&
      consultationCode === data.consultationCode;
  });

  if (!matched) {
    throw new Error('患者IDまたはかかりつけ相談コードを確認できませんでした。入力内容をご確認ください。');
  }
}

function getConsultationSheet_() {
  const spreadsheet = getSpreadsheet_();
  return spreadsheet.getSheetByName(CONFIG.sheetName) || spreadsheet.insertSheet(CONFIG.sheetName);
}

function getAuthorizedPatientsSheet_() {
  const spreadsheet = getSpreadsheet_();
  const sheetNames = [CONFIG.authorizedPatientsSheetName].concat(CONFIG.authorizedPatientsSheetNameAliases || []);
  const existingSheet = sheetNames
    .map((sheetName) => spreadsheet.getSheetByName(sheetName))
    .find((sheet) => sheet);

  return existingSheet || spreadsheet.insertSheet(CONFIG.authorizedPatientsSheetName);
}

function getSpreadsheet_() {
  const spreadsheetId = getConfiguredValue_(CONFIG.spreadsheetIdProperty, looksLikeSpreadsheetId_);
  const spreadsheet = spreadsheetId
    ? SpreadsheetApp.openById(spreadsheetId)
    : SpreadsheetApp.getActiveSpreadsheet();

  if (!spreadsheet) {
    throw new Error('保存先スプレッドシートが見つかりません。Script PropertiesにSPREADSHEET_IDを設定してください。');
  }

  return spreadsheet;
}

function ensureHeaderRow_(sheet) {
  const currentHeaders = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  const hasHeaders = currentHeaders.some((value) => value);

  if (!hasHeaders) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
    return;
  }

  removeDeprecatedConsultationColumns_(sheet);
}

function removeDeprecatedConsultationColumns_(sheet) {
  const deprecatedHeaders = ['写真添付有無'];
  const lastColumn = sheet.getLastColumn();
  if (lastColumn < 1) return;

  const currentHeaders = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  for (let index = currentHeaders.length - 1; index >= 0; index--) {
    if (deprecatedHeaders.indexOf(currentHeaders[index]) !== -1) {
      sheet.deleteColumn(index + 1);
    }
  }
}

function ensureColumnFormats_(sheet) {
  const patientIdColumn = HEADERS.indexOf('患者ID') + 1;
  const phoneColumn = HEADERS.indexOf('返信用電話番号') + 1;

  sheet.getRange(1, patientIdColumn, sheet.getMaxRows(), 1).setNumberFormat('@');
  sheet.getRange(1, phoneColumn, sheet.getMaxRows(), 1).setNumberFormat('@');
}

function ensureAuthorizedPatientsHeaderRow_(sheet) {
  const headerWidth = Math.max(sheet.getLastColumn(), AUTHORIZED_PATIENT_HEADERS.length);
  const currentHeaders = sheet.getRange(1, 1, 1, headerWidth).getValues()[0].map(cleanText_);
  const hasHeaders = currentHeaders.some((value) => value);

  if (!hasHeaders) {
    sheet.getRange(1, 1, 1, AUTHORIZED_PATIENT_HEADERS.length).setValues([AUTHORIZED_PATIENT_HEADERS]);
    sheet.setFrozenRows(1);
    return;
  }

  AUTHORIZED_PATIENT_HEADERS.forEach((header) => {
    if (currentHeaders.indexOf(header) === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(header);
    }
  });

  const refreshedHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(cleanText_);
  ['患者ID', '相談コード', '有効'].forEach((requiredHeader, index) => {
    if (refreshedHeaders.indexOf(requiredHeader) === -1) {
      sheet.getRange(1, index + 1).setValue(requiredHeader);
    }
  });
}

function getAuthorizedPatientsHeaderMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(cleanText_);
  return headers.reduce((map, header, index) => {
    if (header) {
      map[header] = index;
    }
    return map;
  }, {});
}

function getReceptionPatientList(auth) {
  validateReceptionAuth_(auth);

  const sheet = getAuthorizedPatientsSheet_();
  ensureAuthorizedPatientsHeaderRow_(sheet);
  ensureAuthorizedPatientsColumnFormats_(sheet);

  const headerMap = getAuthorizedPatientsHeaderMap_(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return {
      generatedAt: formatDateTime_(new Date()),
      patients: []
    };
  }

  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const patients = values
    .map((row, index) => {
      const enabledAt = row[headerMap['有効化日時']];
      return {
        rowNumber: index + 2,
        patientId: normalizePatientId_(row[headerMap['患者ID']]),
        consultationCode: cleanText_(row[headerMap['相談コード']]),
        enabled: isEnabledValue_(row[headerMap['有効']]),
        enabledAt: enabledAt ? formatDateTime_(enabledAt) : '',
        enabledBy: cleanText_(row[headerMap['有効化担当']]),
        memo: cleanText_(row[headerMap['メモ']])
      };
    })
    .filter((patient) => patient.patientId || patient.consultationCode);

  return {
    generatedAt: formatDateTime_(new Date()),
    patients
  };
}

function activateReceptionPatient(payload) {
  validateReceptionAuth_(payload && payload.auth);

  const rowNumber = Number(payload && payload.rowNumber);
  const staffName = cleanText_(payload && payload.staffName);
  if (!Number.isInteger(rowNumber) || rowNumber < 2) {
    throw new Error('更新対象が正しくありません。画面を再読み込みしてください。');
  }
  if (!staffName) {
    throw new Error('担当者名を入力してください。');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sheet = getAuthorizedPatientsSheet_();
    ensureAuthorizedPatientsHeaderRow_(sheet);
    ensureAuthorizedPatientsColumnFormats_(sheet);
    const headerMap = getAuthorizedPatientsHeaderMap_(sheet);

    if (rowNumber > sheet.getLastRow()) {
      throw new Error('対象患者が見つかりません。画面を再読み込みしてください。');
    }

    const row = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (!normalizePatientId_(row[headerMap['患者ID']]) || !normalizeConsultationCode_(row[headerMap['相談コード']])) {
      throw new Error('患者IDまたは相談コードが未登録の行は有効化できません。');
    }

    if (isEnabledValue_(row[headerMap['有効']])) {
      const enabledAt = row[headerMap['有効化日時']];
      return {
        rowNumber,
        enabled: true,
        enabledAt: enabledAt ? formatDateTime_(enabledAt) : '',
        enabledBy: cleanText_(row[headerMap['有効化担当']])
      };
    }

    const now = new Date();
    sheet.getRange(rowNumber, headerMap['有効'] + 1).setValue(true);
    sheet.getRange(rowNumber, headerMap['有効化日時'] + 1).setValue(now);
    sheet.getRange(rowNumber, headerMap['有効化担当'] + 1).setValue(staffName);

    return {
      rowNumber,
      enabled: true,
      enabledAt: formatDateTime_(now),
      enabledBy: staffName
    };
  } finally {
    lock.releaseLock();
  }
}

function validateReceptionAuth_(auth) {
  const configuredPin = cleanText_(PropertiesService.getScriptProperties().getProperty(CONFIG.receptionPinProperty));
  if (!configuredPin) {
    throw new Error('受付管理用PINが未設定です。Script PropertiesにRECEPTION_PINを設定してください。');
  }

  const submittedPin = cleanText_(auth && auth.pin);
  if (!submittedPin || submittedPin !== configuredPin) {
    throw new Error('受付管理用PINが正しくありません。');
  }
}

function formatDateTime_(value) {
  const date = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(date.getTime())) return '';
  return Utilities.formatDate(date, CONFIG.timezone, 'yyyy/MM/dd HH:mm');
}

function ensureAuthorizedPatientsColumnFormats_(sheet) {
  sheet.getRange(1, 1, sheet.getMaxRows(), 2).setNumberFormat('@');

  const headerMap = getAuthorizedPatientsHeaderMap_(sheet);
  if (headerMap['有効化日時'] !== undefined) {
    sheet.getRange(1, headerMap['有効化日時'] + 1, sheet.getMaxRows(), 1).setNumberFormat('yyyy/mm/dd hh:mm');
  }
}

function setupAuthorizedPatientsSheet() {
  const sheet = getAuthorizedPatientsSheet_();
  ensureAuthorizedPatientsHeaderRow_(sheet);
  ensureAuthorizedPatientsColumnFormats_(sheet);
}

function getDoctorEmail_() {
  return getConfiguredValue_(CONFIG.doctorEmailProperty, looksLikeEmail_);
}

function getConfiguredValue_(propertyNameOrValue, directValueDetector) {
  if (!propertyNameOrValue) return '';

  const directValue = cleanText_(propertyNameOrValue);
  if (directValueDetector(directValue)) {
    return directValue;
  }

  return cleanText_(PropertiesService.getScriptProperties().getProperty(directValue));
}

function looksLikeEmail_(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value);
}

function looksLikeSpreadsheetId_(value) {
  return /^[a-zA-Z0-9_-]{25,}$/.test(value);
}

function cleanText_(value) {
  return String(value || '').trim();
}

function normalizePatientId_(value) {
  const text = cleanText_(value);
  if (/^\d{1,5}$/.test(text)) {
    return text.padStart(5, '0');
  }

  return text;
}

function normalizeConsultationCode_(value) {
  return cleanText_(value).toUpperCase().replace(/\s+/g, '');
}

function isEnabledValue_(value) {
  if (value === true) return true;

  const text = cleanText_(value).toLowerCase();
  return text === 'true' ||
    text === '有効' ||
    text === 'yes' ||
    text === 'y' ||
    text === '1';
}
