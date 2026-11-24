const SPREADSHEET_ID_KEY = 'AGEX902_SPREADSHEET_ID';
const LECTURER_KEY = 'AGEX902_LECTURER_KEY';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('AGEX 902 | Advanced Econometrics Revision')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function setupPortal() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty(SPREADSHEET_ID_KEY);
  let ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.create('AGEX 902 Revision Portal Data');
  props.setProperty(SPREADSHEET_ID_KEY, ss.getId());
  const sheets = {
    Students: ['Timestamp','Student ID','Full Name','Registration Number','Email','Last Login'],
    Attempts: ['Timestamp','Student ID','Registration Number','Module','Theory Score','Theory Total','Calculation Score','Calculation Total','Overall Percent'],
    Activity: ['Timestamp','Student ID','Registration Number','Module','Activity']
  };
  Object.keys(sheets).forEach(name => {
    let sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0) sh.appendRow(sheets[name]);
    sh.setFrozenRows(1);
  });
  if (!props.getProperty(LECTURER_KEY)) {
    props.setProperty(LECTURER_KEY, Utilities.getUuid());
  }
  Logger.log('Lecturer access key (keep private): ' + props.getProperty(LECTURER_KEY));
  return {spreadsheetUrl:ss.getUrl()};
}

function getSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty(SPREADSHEET_ID_KEY);
  if (!id) throw new Error('Portal is not initialized. Run setupPortal() once from the Apps Script editor.');
  return SpreadsheetApp.openById(id);
}

function studentLogin(data) {
  if (!data || !data.name || !data.reg || !data.email) throw new Error('Enter your full name, registration number and email.');
  const ss = getSpreadsheet_(), sh = ss.getSheetByName('Students');
  const reg = String(data.reg).trim().toUpperCase();
  const email = String(data.email).trim().toLowerCase();
  const rows = sh.getDataRange().getValues();
  let studentId = '';
  for (let i=1;i<rows.length;i++) {
    if (String(rows[i][3]).trim().toUpperCase() === reg) {
      if (String(rows[i][4]).trim().toLowerCase() !== email) throw new Error('The email does not match this registration number. Contact your lecturer.');
      studentId = rows[i][1];
      sh.getRange(i+1,6).setValue(new Date());
      sh.getRange(i+1,3).setValue(data.name);
      break;
    }
  }
  if (!studentId) {
    studentId = Utilities.getUuid();
    sh.appendRow([new Date(),studentId,data.name,reg,email,new Date()]);
  }
  return {id:studentId,name:data.name,reg:reg,email:email,progress:getStudentProgress_(studentId)};
}

function recordActivity(student, module, activity) {
  const ss=getSpreadsheet_();
  ss.getSheetByName('Activity').appendRow([new Date(),student.id,student.reg,module,activity]);
  return true;
}

function submitAttempt(student, result) {
  const ss=getSpreadsheet_();
  const percent = result.total ? 100*result.score/result.total : 0;
  ss.getSheetByName('Attempts').appendRow([
    new Date(),student.id,student.reg,result.module,
    result.theoryScore,result.theoryTotal,result.calcScore,result.calcTotal,percent
  ]);
  recordActivity(student,result.module,'Assessment submitted');
  return {percent:Math.round(percent*100)/100};
}

function getStudentProgress(studentId, registrationNumber) {
  if (!studentId || !registrationNumber) throw new Error('Student session details are required.');
  const students=getSpreadsheet_().getSheetByName('Students').getDataRange().getValues();
  const valid=students.slice(1).some(r=>r[1]===studentId && String(r[3]).toUpperCase()===String(registrationNumber).trim().toUpperCase());
  if (!valid) throw new Error('Student record could not be verified. Please sign in again.');
  return getStudentProgress_(studentId);
}

function getStudentProgress_(studentId) {
  const rows=getSpreadsheet_().getSheetByName('Attempts').getDataRange().getValues();
  return rows.slice(1).filter(r=>r[1]===studentId).map(r=>({
    module:r[3], theoryScore:r[4], theoryTotal:r[5],
    calcScore:r[6], calcTotal:r[7], percent:r[8],
    timestamp:r[0] instanceof Date ? r[0].toISOString() : r[0]
  }));
}

function lecturerLogin(key) {
  const saved=PropertiesService.getScriptProperties().getProperty(LECTURER_KEY);
  if (!saved || String(key)!==saved) throw new Error('Invalid lecturer access key.');
  const token=Utilities.getUuid();
  CacheService.getScriptCache().put('AGEX902_LECTURER_SESSION_'+token,'1',21600);
  return {token:token,dashboard:lecturerDashboardData_()};
}

function getLecturerDashboard(token) {
  if (!token || CacheService.getScriptCache().get('AGEX902_LECTURER_SESSION_'+token)!=='1') {
    throw new Error('Lecturer session expired. Please sign in again.');
  }
  return lecturerDashboardData_();
}

function lecturerDashboardData_() {
  const ss=getSpreadsheet_();
  const students=ss.getSheetByName('Students').getDataRange().getValues().slice(1);
  const attempts=ss.getSheetByName('Attempts').getDataRange().getValues().slice(1);
  const activity=ss.getSheetByName('Activity').getDataRange().getValues().slice(1);
  const latest={}, best={};
  attempts.forEach(r=>{
    const key=r[1]+'|'+r[3];
    if(!latest[key] || new Date(r[0])>new Date(latest[key][0])) latest[key]=r;
    if(!best[key] || Number(r[8]||0)>Number(best[key][8]||0)) best[key]=r;
  });
  const latestRows=Object.values(latest), bestRows=Object.values(best);
  const avg=latestRows.length ? latestRows.reduce((s,r)=>s+Number(r[8]||0),0)/latestRows.length : 0;
  const roster=students.map(s=>{
    const latestForStudent=latestRows.filter(r=>r[1]===s[1]);
    const bestForStudent=bestRows.filter(r=>r[1]===s[1]);
    const completed=new Set(latestForStudent.map(r=>r[3]));
    const latestAverage=latestForStudent.length ? latestForStudent.reduce((v,r)=>v+Number(r[8]||0),0)/latestForStudent.length : 0;
    const bestAverage=bestForStudent.length ? bestForStudent.reduce((v,r)=>v+Number(r[8]||0),0)/bestForStudent.length : 0;
    return {name:s[2],reg:s[3],email:s[4],lastLogin:s[5],completed:completed.size,
      average:Math.round(latestAverage*100)/100,bestAverage:Math.round(bestAverage*100)/100,
      attempts:attempts.filter(r=>r[1]===s[1]).length};
  });
  const scoreDetails=latestRows.map(r=>{
    const key=r[1]+'|'+r[3], bestRow=best[key];
    const student=students.find(s=>s[1]===r[1]);
    return {name:student?student[2]:'',reg:r[2],module:r[3],
      latestPercent:Math.round(Number(r[8]||0)*100)/100,
      bestPercent:Math.round(Number(bestRow[8]||0)*100)/100};
  });
  return {
    totalStudents:students.length,totalAttempts:attempts.length,
    average:Math.round(avg*100)/100,
    activeStudents:new Set(activity.map(r=>r[1])).size,
    roster:roster,scoreDetails:scoreDetails,
    attempts:attempts.slice(-150).reverse().map(r=>({date:r[0],reg:r[2],module:r[3],theory:r[4]+'/'+r[5],calculation:r[6]+'/'+r[7],percent:Math.round(Number(r[8]||0)*100)/100}))
  };
}