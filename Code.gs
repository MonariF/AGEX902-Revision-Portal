const SPREADSHEET_ID_KEY = 'AGEX902_SPREADSHEET_ID';
const LECTURER_KEY = 'AGEX902_LECTURER_KEY';
function doGet(){return HtmlService.createHtmlOutputFromFile('Index').setTitle('AGEX 902 | Advanced Econometrics Revision').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);}
function setupPortal(){
 const p=PropertiesService.getScriptProperties();let id=p.getProperty(SPREADSHEET_ID_KEY);
 const ss=id?SpreadsheetApp.openById(id):SpreadsheetApp.create('AGEX 902 Revision Portal Data');
 p.setProperty(SPREADSHEET_ID_KEY,ss.getId());
 const defs={Students:['Timestamp','Student ID','Full Name','Registration Number','Email','Last Login'],Attempts:['Timestamp','Student ID','Registration Number','Module','Theory Score','Theory Total','Calculation Score','Calculation Total','Overall Percent'],Activity:['Timestamp','Student ID','Registration Number','Module','Activity']};
 Object.keys(defs).forEach(n=>{let s=ss.getSheetByName(n)||ss.insertSheet(n);if(s.getLastRow()===0)s.appendRow(defs[n]);s.setFrozenRows(1);});
 if(!p.getProperty(LECTURER_KEY))p.setProperty(LECTURER_KEY,Utilities.getUuid());
 return {spreadsheetUrl:ss.getUrl(),lecturerKey:p.getProperty(LECTURER_KEY)};
}
function getSS_(){const id=PropertiesService.getScriptProperties().getProperty(SPREADSHEET_ID_KEY);if(!id)throw new Error('Run setupPortal() once from the Apps Script editor.');return SpreadsheetApp.openById(id);}
function studentLogin(d){
 if(!d||!d.name||!d.reg||!d.email)throw new Error('Enter your full name, registration number and email.');
 const sh=getSS_().getSheetByName('Students'),reg=String(d.reg).trim().toUpperCase(),email=String(d.email).trim().toLowerCase(),rows=sh.getDataRange().getValues();let id='';
 for(let i=1;i<rows.length;i++)if(String(rows[i][3]).trim().toUpperCase()===reg){if(String(rows[i][4]).trim().toLowerCase()!==email)throw new Error('Email does not match this registration number.');id=rows[i][1];sh.getRange(i+1,6).setValue(new Date());sh.getRange(i+1,3).setValue(d.name);break;}
 if(!id){id=Utilities.getUuid();sh.appendRow([new Date(),id,d.name,reg,email,new Date()]);}
 return {id:id,name:d.name,reg:reg,email:email};
}
function recordActivity(st,m,a){getSS_().getSheetByName('Activity').appendRow([new Date(),st.id,st.reg,m,a]);return true;}
function submitAttempt(st,r){
 const pct=r.total?100*r.score/r.total:0;
 getSS_().getSheetByName('Attempts').appendRow([new Date(),st.id,st.reg,r.module,r.theoryScore,r.theoryTotal,r.calcScore,r.calcTotal,pct]);
 recordActivity(st,r.module,'Assessment submitted');return {percent:Math.round(pct*100)/100};
}
// Keep the actual lecturer key private in Apps Script > Project Settings > Script Properties.
// Property name: AGEX902_LECTURER_KEY. Do not put the key in this public repository.
function lecturerLogin(key){const saved=PropertiesService.getScriptProperties().getProperty(LECTURER_KEY);if(!saved||String(key)!==saved)throw new Error('Invalid lecturer access key.');return lecturerDashboardData();}
function lecturerDashboardData(){
 const ss=getSS_(),students=ss.getSheetByName('Students').getDataRange().getValues().slice(1),attempts=ss.getSheetByName('Attempts').getDataRange().getValues().slice(1),activity=ss.getSheetByName('Activity').getDataRange().getValues().slice(1),latest={};
 attempts.forEach(r=>{const k=r[1]+'|'+r[3];if(!latest[k]||new Date(r[0])>new Date(latest[k][0]))latest[k]=r;});
 const la=Object.values(latest),avg=la.length?la.reduce((s,r)=>s+Number(r[8]||0),0)/la.length:0;
 return {totalStudents:students.length,totalAttempts:attempts.length,average:Math.round(avg*100)/100,activeStudents:new Set(activity.map(r=>r[1])).size,
 roster:students.map(s=>{const a=la.filter(r=>r[1]===s[1]);return {name:s[2],reg:s[3],email:s[4],completed:new Set(a.map(r=>r[3])).size,average:a.length?Math.round(a.reduce((v,r)=>v+Number(r[8]||0),0)/a.length*100)/100:0,attempts:a.length};}),
 attempts:attempts.slice(-150).reverse().map(r=>({date:r[0],reg:r[2],module:r[3],theory:r[4]+'/'+r[5],calculation:r[6]+'/'+r[7],percent:Math.round(Number(r[8]||0)*100)/100}))};
}