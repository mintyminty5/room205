/**
 * 205호 타임캡슐 · 공유판 (익명 제출 → 교사 승인 → 우리 반 게시판)
 * - 학생 글은 이 스크립트를 배포한 선생님의 구글 시트에만 저장됩니다.
 * - 이름은 받지 않습니다. 반 코드(반코드 시트에 등록된 값)가 맞아야 제출됩니다.
 * - 시트의 '승인' 칸을 체크한 글만 게시판에 보입니다.
 */
var SUB='제출', CODES='반코드';
var MAX_LEN=150, MIN_LEN=10, MAX_ROWS_PER_CODE=400;

function setup(){
  var props=PropertiesService.getScriptProperties();
  var ss;
  var id=props.getProperty('SHEET_ID');
  if(id){ ss=SpreadsheetApp.openById(id); }
  else { ss=SpreadsheetApp.create('205호 공유판 제출'); props.setProperty('SHEET_ID',ss.getId()); }
  var s=ss.getSheetByName(SUB)||ss.insertSheet(SUB);
  if(s.getLastRow()===0){ s.appendRow(['시간','반 코드','내일 내가 할 것','우리 반이 다음 주에 해볼 것','승인','비고']); s.setFrozenRows(1); s.setColumnWidth(3,380); s.setColumnWidth(4,380); }
  var c=ss.getSheetByName(CODES)||ss.insertSheet(CODES);
  if(c.getLastRow()===0){ c.appendRow(['반 코드(여기에 사용할 코드를 한 줄에 하나씩 적으세요)']); c.appendRow(['TEST205']); c.setColumnWidth(1,420); }
  var def=ss.getSheetByName('Sheet1')||ss.getSheetByName('시트1'); if(def&&ss.getSheets().length>1){ ss.deleteSheet(def); }
  Logger.log('시트 주소: '+ss.getUrl());
  return ss.getUrl();
}

function ss_(){ var id=PropertiesService.getScriptProperties().getProperty('SHEET_ID'); if(!id) throw new Error('setup()을 먼저 실행하세요'); return SpreadsheetApp.openById(id); }
function json_(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function clean_(t){ return String(t==null?'':t).replace(/[\u0000-\u001f\u007f]/g,' ').replace(/\s+/g,' ').trim().slice(0,MAX_LEN); }
function validCode_(ss,code){
  if(!code) return false;
  var c=ss.getSheetByName(CODES); if(!c) return false;
  var vals=c.getRange(2,1,Math.max(1,c.getLastRow()-1),1).getValues();
  for(var i=0;i<vals.length;i++){ if(String(vals[i][0]).trim().toLowerCase()===code.toLowerCase()) return true; }
  return false;
}

function doPost(e){
  var lock=LockService.getScriptLock();
  try{
    lock.waitLock(10000);
    var body=JSON.parse((e&&e.postData&&e.postData.contents)||'{}');
    if(body.hp) return json_({ok:true}); // 허니팟: 봇은 성공한 것처럼 처리하고 저장하지 않음
    var code=clean_(body.code).slice(0,30), s3=clean_(body.step3), s4=clean_(body.step4);
    var ss=ss_();
    if(!validCode_(ss,code)) return json_({ok:false,error:'code'});
    if(s3.length<MIN_LEN||s4.length<MIN_LEN) return json_({ok:false,error:'short'});
    var s=ss.getSheetByName(SUB);
    var count=0, last=s.getLastRow();
    if(last>1){ var codes=s.getRange(2,2,last-1,1).getValues(); for(var i=0;i<codes.length;i++){ if(String(codes[i][0]).toLowerCase()===code.toLowerCase()) count++; } }
    if(count>=MAX_ROWS_PER_CODE) return json_({ok:false,error:'full'});
    s.appendRow([new Date(),code,s3,s4,false,'']);
    var r=s.getLastRow();
    s.getRange(r,5).insertCheckboxes();
    return json_({ok:true});
  }catch(err){
    return json_({ok:false,error:'server'});
  }finally{ try{lock.releaseLock();}catch(x){} }
}

function doGet(e){
  try{
    var p=(e&&e.parameter)||{};
    if(p.action==='ping') return json_({ok:true});
    if(p.action==='board'){
      var code=clean_(p.code).slice(0,30);
      var ss=ss_();
      if(!validCode_(ss,code)) return json_({ok:false,error:'code'});
      var s=ss.getSheetByName(SUB), last=s.getLastRow(), items=[];
      if(last>1){
        var vals=s.getRange(2,2,last-1,4).getValues();
        for(var i=0;i<vals.length;i++){
          if(String(vals[i][0]).toLowerCase()===code.toLowerCase() && vals[i][3]===true){ items.push({s3:String(vals[i][1]),s4:String(vals[i][2])}); }
        }
      }
      return json_({ok:true,items:items});
    }
    return json_({ok:true,service:'room205-share'});
  }catch(err){ return json_({ok:false,error:'server'}); }
}
