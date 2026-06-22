/**
 * ИУС 2.0 — бесплатная база работ: Google Таблица + Apps Script + Google Drive
 * Версия START OVER CLEAN
 *
 * Как работает:
 * - works: лист с работами учеников
 * - tasks: лист с заданиями, которые можно менять на странице учителя
 * - audio: файлы сохраняются в папку Google Drive, ID папки хранится в Script Properties
 */

const IUS_VERSION = 'task-image-2026-06-22';
const WORKS_SHEET = 'works';
const TASKS_SHEET = 'tasks';
const AUDIO_FOLDER_NAME = 'IUS-2.0 audio';
const AUDIO_FOLDER_ID_PROP = 'IUS_AUDIO_FOLDER_ID';
const TASK_IMAGE_FOLDER_NAME = 'IUS-2.0 task images';
const TASK_IMAGE_FOLDER_ID_PROP = 'IUS_TASK_IMAGE_FOLDER_ID';

const WORK_HEADERS = [
  'id',
  'sentAt',
  'status',
  'studentName',
  'className',
  'taskId',
  'taskTitle',
  'note',
  'audioFileId',
  'audioFileUrl',
  'audioMimeType',
  'audioFileName',
  'audioError',
  'payloadJson',
  'resultJson'
];

const TASK_HEADERS = [
  'id',
  'tabTitle',
  'title',
  'prepSeconds',
  'answerSeconds',
  'instruction',
  'content',
  'imageFileId',
  'imageFileUrl',
  'imageMimeType',
  'imageFileName',
  'imageError'
];

const DEFAULT_TASKS = [
  {
    id: '1',
    tabTitle: '1 сорудах\nАаҕыы',
    title: '1 сорудах. Тиэкиһи ааҕыы',
    prepSeconds: 60,
    answerSeconds: 120,
    instruction: 'Тиэкиһи болҕойон аах. 1 мүнүүтэ бэлэмнэн.',
    content: 'Саха сылгыта – тымныыны тулуйар дьикти боруода.\n\nСаха сылгыта – аан дойдуга саамай тымныыны тулуйар дьиэ кыыла. Бу сылгылар тоҕус ый устата аһаҕас халлаан анныгар сылдьаллар, сиртэн-буортан астарын булуналлар.\n\nКыһын халыҥ хаары туйахтарынан хаһан, хагдарыйбыт оту булан аһыыллар.\n\nСаха сылгыта намыһах уҥуохтаах, түүтэ олус хойуу, сиэлэ-кутуруга уһун. Тириитин анныгар халыҥ сыа мунньуллар, ол тымныыны туоруурга көмөлөһөр.\n\nСылгы саха киһитигэр былыр былыргаттан тыын суолталаах этэ. Сахалар сылгыны Күн Дьөһөгөй оҕото диэн ааттыыллар. Сылгы – саха олоҕун, сиэрин-туомун арахсыбат чааһа.'
  },
  {
    id: '2',
    tabTitle: '2 сорудах\nКэпсээһин',
    title: '2 сорудах. Тиэкиһи кэпсээһин',
    prepSeconds: 60,
    answerSeconds: 180,
    instruction: 'Аахпыт тиэкискин сиһилии кэпсээ. 1 мүнүүтэ иһигэр бэлэмнэн. Наада буоллаҕына сурун.',
    content: 'Бэлиэтэн: сүрүн санааны, сылгы уратытын, саха олоҕор суолтатын умнуман кэпсээ.'
  },
  {
    id: '3',
    tabTitle: '3 сорудах\nСэһэргээһии',
    title: '3 сорудах. Сэһэргээһии көрүҥүн талыы',
    prepSeconds: 60,
    answerSeconds: 180,
    instruction: 'Сэһэргэһии тиэмэтиттэн биири талан кэпсээҥ. 1 мүнүүтэ бэлэмнэн, 3 мүнүүтэ кэпсээ.',
    content: 'Тиэмэлэр:\n1. Хартыынаны ойуулааһын.\n2. Сүрүн боппуруоһу тойоннооһун: Саха тыла — төрүт тыл, кинини харыстыыр эбэтэр хаалларар кэм кэллэ.\n3. Олоххо буолбут түгэнтэн сэһэргээһин.\n\nКэпсииргэр умнума:\n- Саха тыла — ийэ тылым...\n- Мин санаабар, ...\n- Тоҕо диэтэххэ, ...\n- Холобур, ...\n- Түмүккэ, ол иһин...'
  },
  {
    id: '4',
    tabTitle: '4 сорудах\nКэпсэтии',
    title: '4 сорудах. Ыйытыыларга эппиэттээһин',
    prepSeconds: 0,
    answerSeconds: 180,
    instruction: 'Кэпсэтии. Ыйытыыларга толору хоруйу биэр.',
    content: 'Ыйытыылар:\n1. Саха тылын харыстыырга туохха кыһаллыахтаахпыт?\n2. Оскуола саха тылын сайдыытыгар хайдах көмөлөһүөн сөбүй?\n3. Эн бэйэҥ саха тылын туттууга тугу оҥоруоххун сөбүй?'
  }
];

/**
 * Первый запуск. Создаёт листы works/tasks и заполняет задания.
 * Не требует Google Drive.
 */
function setupIUS() {
  const ss = getSpreadsheet_();
  const works = getOrCreateSheet_(ss, WORKS_SHEET, WORK_HEADERS);
  ensureHeaders_(works, WORK_HEADERS);
  const tasks = getOrCreateSheet_(ss, TASKS_SHEET, TASK_HEADERS);
  ensureHeaders_(tasks, TASK_HEADERS);
  if (tasks.getLastRow() < 2) {
    const values = DEFAULT_TASKS.map(taskToRow_);
    tasks.getRange(2, 1, values.length, TASK_HEADERS.length).setValues(values);
  }
  Logger.log('setupIUS завершён. Таблица: ' + ss.getUrl());
}

/**
 * Запусти отдельно, чтобы Google запросил доступ к DriveApp.
 * После успешного запуска появится папка IUS-2.0 audio и тестовый файл будет удалён.
 */
function authorizeIUS() {
  setupIUS();
  const audioFolder = getOrCreateAudioFolder_();
  const imageFolder = getOrCreateTaskImageFolder_();
  const blob = Utilities.newBlob('IUS Drive test ' + new Date().toISOString(), 'text/plain', 'ius-drive-test.txt');
  const file = audioFolder.createFile(blob);
  const url = file.getUrl();
  file.setTrashed(true);
  Logger.log('DriveApp разрешён. Тестовый файл создан и удалён: ' + url);
  Logger.log('Папка аудио: ' + audioFolder.getUrl());
  Logger.log('Папка картинок заданий: ' + imageFolder.getUrl());
}

/**
 * Диагностика. Запускай из Apps Script, если что-то не работает.
 */
function diagnoseIUS() {
  const result = {
    version: IUS_VERSION,
    spreadsheet: null,
    worksSheet: false,
    tasksSheet: false,
    drive: false,
    audioFolderId: null,
    taskImageFolderId: null,
    error: null
  };
  try {
    const ss = getSpreadsheet_();
    result.spreadsheet = ss.getUrl();
    result.worksSheet = Boolean(ss.getSheetByName(WORKS_SHEET));
    result.tasksSheet = Boolean(ss.getSheetByName(TASKS_SHEET));
    const folder = getOrCreateAudioFolder_();
    const imageFolder = getOrCreateTaskImageFolder_();
    result.drive = true;
    result.audioFolderId = folder.getId();
    result.taskImageFolderId = imageFolder.getId();
  } catch (error) {
    result.error = String(error && error.message ? error.message : error);
  }
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

/**
 * Тестовая строка без аудио. Проверяет запись в таблицу.
 */
function testAddRow() {
  setupIUS();
  const submission = {
    id: 'test-' + Date.now(),
    sentAt: new Date().toISOString(),
    status: 'sent',
    student: { fullName: 'Тестовый ученик', className: '8А', login: 'test' },
    task: { id: '1', title: 'Тестовое задание' },
    note: 'Проверка таблицы',
    audio: null
  };
  const saved = saveSubmission_(submission);
  Logger.log(JSON.stringify(saved, null, 2));
}

function doGet(e) {
  const action = String((e && e.parameter && e.parameter.action) || 'ping');
  try {
    setupIUS();
    let data;
    if (action === 'ping') data = { ok: true, version: IUS_VERSION, message: 'Apps Script работает' };
    else if (action === 'diagnose') data = { ok: true, diagnose: diagnoseIUS() };
    else if (action === 'list') data = { ok: true, items: listSubmissions_(Number(e.parameter.limit || 50)) };
    else if (action === 'get') data = { ok: true, item: getSubmissionById_(String(e.parameter.id || '')) };
    else if (action === 'getAudio') data = getAudioByWorkId_(String(e.parameter.id || ''));
    else if (action === 'getTaskImage') data = getTaskImageByTaskId_(String(e.parameter.id || ''));
    else if (action === 'listTasks') data = { ok: true, tasks: listTasks_() };
    else data = { ok: false, error: 'Неизвестное action: ' + action };
    return output_(e, data);
  } catch (error) {
    return output_(e, { ok: false, error: String(error && error.message ? error.message : error) });
  }
}

function doPost(e) {
  try {
    setupIUS();
    const payload = parsePayload_(e);
    const action = String(payload.action || 'submit');
    let data;
    if (action === 'submit') data = { ok: true, item: saveSubmission_(payload.submission || payload) };
    else if (action === 'saveResult') data = { ok: true, item: saveResult_(payload.submissionId, payload.result) };
    else if (action === 'saveTask') data = { ok: true, tasks: saveTask_(payload.task) };
    else if (action === 'uploadTaskImage') data = { ok: true, task: saveTaskImage_(payload.taskId, payload.image) };
    else if (action === 'clearTaskImage') data = { ok: true, task: clearTaskImage_(payload.taskId) };
    else data = { ok: false, error: 'Неизвестное action: ' + action };
    return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(error && error.message ? error.message : error) })).setMimeType(ContentService.MimeType.JSON);
  }
}

function output_(e, data) {
  const callback = String((e && e.parameter && e.parameter.callback) || '').trim();
  const json = JSON.stringify(data);
  if (callback && /^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
    return ContentService.createTextOutput(callback + '(' + json + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function parsePayload_(e) {
  if (e && e.parameter && e.parameter.payload) return JSON.parse(e.parameter.payload);
  if (e && e.postData && e.postData.contents) return JSON.parse(e.postData.contents);
  return {};
}

function getSpreadsheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Скрипт должен быть открыт через Google Таблицу: Расширения → Apps Script');
  return ss;
}

function getOrCreateSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  ensureHeaders_(sheet, headers);
  return sheet;
}

function ensureHeaders_(sheet, headers) {
  if (sheet.getMaxColumns() < headers.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  }
  const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  let changed = false;
  for (let i = 0; i < headers.length; i++) {
    if (current[i] !== headers[i]) changed = true;
  }
  if (changed) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
}

function getOrCreateAudioFolder_() {
  const props = PropertiesService.getScriptProperties();
  const oldId = props.getProperty(AUDIO_FOLDER_ID_PROP);
  if (oldId) {
    try {
      return DriveApp.getFolderById(oldId);
    } catch (error) {
      props.deleteProperty(AUDIO_FOLDER_ID_PROP);
    }
  }
  // Не используем getFoldersByName, чтобы не было ошибок на старых разрешениях.
  const folder = DriveApp.createFolder(AUDIO_FOLDER_NAME);
  props.setProperty(AUDIO_FOLDER_ID_PROP, folder.getId());
  return folder;
}

function getOrCreateTaskImageFolder_() {
  const props = PropertiesService.getScriptProperties();
  const oldId = props.getProperty(TASK_IMAGE_FOLDER_ID_PROP);
  if (oldId) {
    try {
      return DriveApp.getFolderById(oldId);
    } catch (error) {
      props.deleteProperty(TASK_IMAGE_FOLDER_ID_PROP);
    }
  }
  const folder = DriveApp.createFolder(TASK_IMAGE_FOLDER_NAME);
  props.setProperty(TASK_IMAGE_FOLDER_ID_PROP, folder.getId());
  return folder;
}

function saveSubmission_(submission) {
  if (!submission) throw new Error('Пустая работа ученика');
  const ss = getSpreadsheet_();
  const sheet = getOrCreateSheet_(ss, WORKS_SHEET, WORK_HEADERS);

  const id = String(submission.id || ('work-' + Date.now()));
  submission.id = id;
  submission.sentAt = submission.sentAt || new Date().toISOString();

  const result = {
    id,
    sentAt: submission.sentAt,
    status: 'sent',
    studentName: (submission.student && submission.student.fullName) || '',
    className: (submission.student && submission.student.className) || '',
    taskId: (submission.task && submission.task.id) || '',
    taskTitle: (submission.task && submission.task.title) || '',
    note: submission.note || '',
    audioFileId: '',
    audioFileUrl: '',
    audioMimeType: (submission.audio && submission.audio.mimeType) || '',
    audioFileName: (submission.audio && submission.audio.fileName) || '',
    audioError: '',
    payloadJson: JSON.stringify(cleanSubmissionForStorage_(submission)),
    resultJson: ''
  };

  // Сначала записываем строку в таблицу, чтобы работа не потерялась даже при ошибке аудио.
  const rowIndex = sheet.getLastRow() + 1;
  sheet.getRange(rowIndex, 1, 1, WORK_HEADERS.length).setValues([workToRow_(result)]);

  if (submission.audio && submission.audio.dataUrl) {
    try {
      const audioInfo = saveAudio_(submission, id);
      result.audioFileId = audioInfo.fileId;
      result.audioFileUrl = audioInfo.url;
      result.audioMimeType = audioInfo.mimeType;
      result.audioFileName = audioInfo.fileName;
      result.status = 'sent';
    } catch (error) {
      result.status = 'audio_error';
      result.audioError = String(error && error.message ? error.message : error);
    }
  } else {
    result.status = 'no_audio';
    result.audioError = 'Аудио не передано';
  }

  // Обновляем строку с audioFileId или ошибкой.
  sheet.getRange(rowIndex, 1, 1, WORK_HEADERS.length).setValues([workToRow_(result)]);
  return result;
}

function saveAudio_(submission, workId) {
  const audio = submission.audio;
  const parsed = parseDataUrl_(audio.dataUrl);
  const folder = getOrCreateAudioFolder_();
  const studentName = sanitizeFileName_((submission.student && submission.student.fullName) || 'student');
  const taskId = sanitizeFileName_((submission.task && submission.task.id) || 'task');
  const ext = extensionFromMime_(parsed.mimeType || audio.mimeType || 'audio/webm');
  const fileName = studentName + '_task-' + taskId + '_' + workId + '.' + ext;
  const bytes = Utilities.base64Decode(parsed.base64);
  const blob = Utilities.newBlob(bytes, parsed.mimeType || audio.mimeType || 'audio/webm', fileName);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return {
    fileId: file.getId(),
    url: file.getUrl(),
    mimeType: parsed.mimeType || audio.mimeType || 'audio/webm',
    fileName
  };
}

function parseDataUrl_(dataUrl) {
  const value = String(dataUrl || '').trim();
  // Принимает data:audio/webm;base64,... и data:audio/webm;codecs=opus;base64,...
  const match = value.match(/^data:([^;,]+)(?:;[^,]*)?;base64,(.*)$/s);
  if (!match) throw new Error('Неверный формат аудио dataUrl');
  return { mimeType: match[1], base64: match[2].replace(/\s/g, '') };
}

function cleanSubmissionForStorage_(submission) {
  const copy = JSON.parse(JSON.stringify(submission));
  if (copy.audio && copy.audio.dataUrl) {
    copy.audio.dataUrl = '[saved-to-drive]';
  }
  return copy;
}

function listSubmissions_(limit) {
  const sheet = getOrCreateSheet_(getSpreadsheet_(), WORKS_SHEET, WORK_HEADERS);
  const rows = readObjects_(sheet, WORK_HEADERS);
  return rows.slice(-Math.max(1, limit || 50)).reverse().map(rowToWork_);
}

function getSubmissionById_(id) {
  if (!id) throw new Error('Не указан id работы');
  const sheet = getOrCreateSheet_(getSpreadsheet_(), WORKS_SHEET, WORK_HEADERS);
  const rows = readObjects_(sheet, WORK_HEADERS);
  const found = rows.find(row => String(row.id) === String(id));
  if (!found) throw new Error('Работа не найдена: ' + id);
  return rowToWork_(found);
}

function getAudioByWorkId_(id) {
  const item = getSubmissionById_(id);
  if (!item.audio || !item.audio.fileId) {
    return { ok: false, error: item.audioError || 'У работы нет audioFileId' };
  }
  const file = DriveApp.getFileById(item.audio.fileId);
  const blob = file.getBlob();
  const mimeType = item.audio.mimeType || blob.getContentType() || 'audio/webm';
  const base64 = Utilities.base64Encode(blob.getBytes());
  return { ok: true, id, mimeType, dataUrl: 'data:' + mimeType + ';base64,' + base64 };
}

function rowToWork_(row) {
  let payload = null;
  try { payload = row.payloadJson ? JSON.parse(row.payloadJson) : null; } catch (e) {}
  return {
    id: row.id,
    sentAt: row.sentAt,
    createdAt: row.sentAt,
    status: row.status,
    student: {
      fullName: row.studentName,
      className: row.className
    },
    task: {
      id: row.taskId,
      title: row.taskTitle
    },
    note: row.note,
    audio: {
      fileId: row.audioFileId,
      fileUrl: row.audioFileUrl,
      directUrl: row.audioFileId ? 'https://drive.google.com/uc?export=download&id=' + row.audioFileId : '',
      mimeType: row.audioMimeType,
      fileName: row.audioFileName
    },
    audioError: row.audioError,
    payload
  };
}

function workToRow_(work) {
  return WORK_HEADERS.map(h => work[h] == null ? '' : work[h]);
}

function readObjects_(sheet, headers) {
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const values = sheet.getRange(2, 1, last - 1, headers.length).getValues();
  return values.map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

function saveResult_(submissionId, result) {
  if (!submissionId) throw new Error('Не указан submissionId');
  const sheet = getOrCreateSheet_(getSpreadsheet_(), WORKS_SHEET, WORK_HEADERS);
  const rows = readObjects_(sheet, WORK_HEADERS);
  const index = rows.findIndex(row => String(row.id) === String(submissionId));
  if (index < 0) throw new Error('Работа не найдена: ' + submissionId);
  const rowNumber = index + 2;
  const statusCol = WORK_HEADERS.indexOf('status') + 1;
  const resultCol = WORK_HEADERS.indexOf('resultJson') + 1;
  sheet.getRange(rowNumber, statusCol).setValue('checked');
  sheet.getRange(rowNumber, resultCol).setValue(JSON.stringify(result || {}));
  return getSubmissionById_(submissionId);
}

function listTasks_() {
  const sheet = getOrCreateSheet_(getSpreadsheet_(), TASKS_SHEET, TASK_HEADERS);
  const rows = readObjects_(sheet, TASK_HEADERS);
  if (!rows.length) return DEFAULT_TASKS;
  return rows.map(rowToTask_);
}

function saveTask_(task) {
  if (!task || !task.id) throw new Error('Не указано задание');
  const sheet = getOrCreateSheet_(getSpreadsheet_(), TASKS_SHEET, TASK_HEADERS);
  const rows = readObjects_(sheet, TASK_HEADERS);
  const index = rows.findIndex(row => String(row.id) === String(task.id));
  const current = index >= 0 ? rowToTask_(rows[index]) : {};
  // Сохраняем существующую картинку, если учитель меняет только текст задания.
  const merged = Object.assign({}, current, task);
  const row = taskToRow_(merged);
  if (index >= 0) sheet.getRange(index + 2, 1, 1, TASK_HEADERS.length).setValues([row]);
  else sheet.appendRow(row);
  return listTasks_();
}

function saveTaskImage_(taskId, image) {
  if (!taskId) throw new Error('Не указан ID задания');
  if (!image || !image.dataUrl) throw new Error('Картинка не передана');
  const parsed = parseDataUrl_(image.dataUrl);
  if (!String(parsed.mimeType || '').toLowerCase().startsWith('image/')) {
    throw new Error('Файл должен быть изображением');
  }
  const folder = getOrCreateTaskImageFolder_();
  const ext = extensionFromMime_(parsed.mimeType || image.mimeType || 'image/jpeg');
  const fileName = 'task-' + sanitizeFileName_(taskId) + '_' + Date.now() + '.' + ext;
  const bytes = Utilities.base64Decode(parsed.base64);
  const blob = Utilities.newBlob(bytes, parsed.mimeType || image.mimeType || 'image/jpeg', fileName);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return updateTaskImageFields_(taskId, {
    imageFileId: file.getId(),
    imageFileUrl: file.getUrl(),
    imageMimeType: parsed.mimeType || image.mimeType || 'image/jpeg',
    imageFileName: image.fileName || fileName,
    imageError: ''
  });
}

function clearTaskImage_(taskId) {
  if (!taskId) throw new Error('Не указан ID задания');
  return updateTaskImageFields_(taskId, {
    imageFileId: '',
    imageFileUrl: '',
    imageMimeType: '',
    imageFileName: '',
    imageError: ''
  });
}

function updateTaskImageFields_(taskId, fields) {
  const sheet = getOrCreateSheet_(getSpreadsheet_(), TASKS_SHEET, TASK_HEADERS);
  const rows = readObjects_(sheet, TASK_HEADERS);
  const index = rows.findIndex(row => String(row.id) === String(taskId));
  if (index < 0) throw new Error('Задание не найдено: ' + taskId);
  const updated = Object.assign({}, rowToTask_(rows[index]), fields || {});
  sheet.getRange(index + 2, 1, 1, TASK_HEADERS.length).setValues([taskToRow_(updated)]);
  return updated;
}

function getTaskImageByTaskId_(taskId) {
  if (!taskId) return { ok: false, error: 'Не указан ID задания' };
  const task = listTasks_().find(t => String(t.id) === String(taskId));
  if (!task || !task.image || !task.image.fileId) {
    return { ok: false, error: 'У задания нет картинки' };
  }
  const file = DriveApp.getFileById(task.image.fileId);
  const blob = file.getBlob();
  const mimeType = task.image.mimeType || blob.getContentType() || 'image/jpeg';
  const base64 = Utilities.base64Encode(blob.getBytes());
  return { ok: true, id: taskId, mimeType, dataUrl: 'data:' + mimeType + ';base64,' + base64 };
}

function taskToRow_(task) {
  return TASK_HEADERS.map(h => task[h] == null ? '' : task[h]);
}

function rowToTask_(row) {
  return {
    id: String(row.id || ''),
    tabTitle: String(row.tabTitle || ''),
    title: String(row.title || ''),
    prepSeconds: Number(row.prepSeconds || 0),
    answerSeconds: Number(row.answerSeconds || 180),
    instruction: String(row.instruction || ''),
    content: String(row.content || ''),
    imageFileId: String(row.imageFileId || ''),
    imageFileUrl: String(row.imageFileUrl || ''),
    imageMimeType: String(row.imageMimeType || ''),
    imageFileName: String(row.imageFileName || ''),
    imageError: String(row.imageError || ''),
    image: {
      fileId: String(row.imageFileId || ''),
      fileUrl: String(row.imageFileUrl || ''),
      directUrl: row.imageFileId ? 'https://drive.google.com/uc?export=view&id=' + row.imageFileId : '',
      mimeType: String(row.imageMimeType || ''),
      fileName: String(row.imageFileName || ''),
      error: String(row.imageError || '')
    }
  };
}

function sanitizeFileName_(value) {
  return String(value || 'file').replace(/[^А-Яа-яA-Za-z0-9_-]+/g, '_').replace(/^_+|_+$/g, '').slice(0, 80) || 'file';
}

function extensionFromMime_(mimeType) {
  const type = String(mimeType || '').toLowerCase();
  if (type.includes('png')) return 'png';
  if (type.includes('jpeg') || type.includes('jpg')) return 'jpg';
  if (type.includes('webp')) return 'webp';
  if (type.includes('gif')) return 'gif';
  if (type.includes('ogg')) return 'ogg';
  if (type.includes('mp4')) return 'm4a';
  if (type.includes('mpeg')) return 'mp3';
  if (type.includes('wav')) return 'wav';
  if (type.startsWith('image/')) return 'jpg';
  return 'webm';
}
