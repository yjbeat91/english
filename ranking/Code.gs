/* ===== 영단어 놀이터 랭킹 서버 (Google Apps Script) =====
   이 코드는 점수를 이 스프레드시트의 'scores' 시트에 적고, 랭킹을 계산해서 돌려줘요.
   학생들은 이 페이지를 직접 열지 않아요. 앱이 뒤에서 조용히 주고받기만 해요.
   장난 기록은 'scores' 시트에서 그 줄을 지우면 랭킹에서 바로 빠져요. */

var SHEET_NAME='scores';

function sheet_(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sh=ss.getSheetByName(SHEET_NAME);
  if(!sh){sh=ss.insertSheet(SHEET_NAME);sh.appendRow(['시간','플레이어ID','닉네임','칭호','게임','점수','시간(초)','기록']);sh.setFrozenRows(1)}
  return sh;
}
function rows_(){
  var v=sheet_().getDataRange().getValues();var out=[];
  for(var i=1;i<v.length;i++){var r=v[i];if(!r[1])continue;
    var ts=r[0] instanceof Date?r[0].getTime():Date.parse(r[0]);if(isNaN(ts))continue;
    out.push({ts:ts,pid:String(r[1]),nick:String(r[2]),title:String(r[3]||''),game:String(r[4]),score:Number(r[5])||0,time:Number(r[6])||0,label:String(r[7]||'')})}
  return out;
}
function json_(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON)}
function doGet(e){
  var p=(e&&e.parameter)||{};
  if(p.action==='top'){
    if(GAME_IDS.indexOf(p.game)<0)return json_({ok:false,error:'game'});
    var period=['today','week','all'].indexOf(p.period)>=0?p.period:'week';
    var b=makeBoard(rows_(),p.game,period,String(p.pid||''),Date.now(),20);b.ok=true;return json_(b);
  }
  return json_({ok:true,service:'영단어 놀이터 랭킹'});
}
function doPost(e){
  var p={};try{p=JSON.parse(e.postData.contents)}catch(err){return json_({ok:false,error:'format'})}
  if(p.action!=='submit')return json_({ok:false,error:'action'});
  if(GAME_IDS.indexOf(p.game)<0)return json_({ok:false,error:'game'});
  var pid=String(p.pid||'').replace(/[^a-z0-9]/gi,'').slice(0,24);if(!pid)return json_({ok:false,error:'pid'});
  var nick=cleanNick(p.nick);if(!nick||isBadNick(nick))nick='익명';
  var title=cleanNick(p.title);
  var score=Math.max(-100000,Math.min(100000,Number(p.score)||0)),time=Math.max(0,Math.min(36000,Math.round(Number(p.time)||0)));
  var label=String(p.label||'').replace(/[<>]/g,'').slice(0,40);
  var lock=LockService.getScriptLock();lock.waitLock(10000);
  try{sheet_().appendRow([new Date(),pid,nick,title,p.game,score,time,label])}finally{lock.releaseLock()}
  var rows=rows_(),now=Date.now(),res={ok:true,ranks:{}};
  ['today','week','all'].forEach(function(per){var b=makeBoard(rows,p.game,per,pid,now,0);res.ranks[per]={rank:b.me?b.me.rank:null,total:b.total}});
  return json_(res);
}

/* ===== 공용 규칙 (앱과 구글 Apps Script가 똑같이 사용해요) ===== */
var GAME_IDS=['cards','memory','connect','cloze','spell','speed'];
var BAD_WORDS=['시발','씨발','씨빨','씨바','쓰발','쉬발','ㅅㅂ','ㅆㅂ','ㅄ','병신','븅신','빙신','ㅂㅅ','개새','개세','개색','새끼','섀끼','쌔끼','ㅅㄲ','좆','존나','졸라','ㅈㄴ','지랄','ㅈㄹ','미친','ㅁㅊ','닥쳐','꺼져','엠창','느금','니애미','니미','애미','애비','썅','쌍년','등신','찐따','fuck','fuk','shit','bitch','sex','섹스','porn','dick','asshole','penis','pussy','nigger','idiot'];
function cleanNick(s){return String(s==null?'':s).replace(/[<>"'`\\]/g,'').replace(/\s+/g,' ').trim().slice(0,10)}
function isBadNick(s){var t=String(s).toLowerCase().replace(/[^a-z가-힣ㄱ-ㅎㅏ-ㅣ]/g,'');
  for(var i=0;i<BAD_WORDS.length;i++){if(t.indexOf(BAD_WORDS[i])>=0)return true}return false}
function kstDay(ms){return new Date(ms+9*3600000).toISOString().slice(0,10)}
function kstWeekStart(ms){var d=new Date(ms+9*3600000);var dow=(d.getUTCDay()+6)%7;return new Date(d.getTime()-dow*86400000).toISOString().slice(0,10)}
function betterRec(a,b){if(a.score!==b.score)return a.score>b.score;if(a.time!==b.time)return a.time<b.time;return a.ts<b.ts}
/* rows: [{ts,pid,nick,title,game,score,time,label}] → 한 사람당 최고 기록 하나로 순위 */
function makeBoard(rows,game,period,pid,now,limit){
  var t=kstDay(now),w=kstWeekStart(now),best={};
  rows.forEach(function(r){if(r.game!==game)return;var d=kstDay(r.ts);
    if(period==='today'&&d!==t)return;if(period==='week'&&d<w)return;
    var b=best[r.pid];if(!b||betterRec(r,b))best[r.pid]=r});
  var list=Object.keys(best).map(function(k){return best[k]}).sort(function(a,b){return betterRec(a,b)?-1:betterRec(b,a)?1:0});
  var out=[],me=null,rank=0;
  list.forEach(function(r,i){if(i===0||list[i-1].score!==r.score||list[i-1].time!==r.time)rank=i+1;
    var e={rank:rank,nick:r.nick,title:r.title||'',label:r.label,me:r.pid===pid};
    if(i<(limit||20))out.push(e);if(r.pid===pid)me=e});
  return {total:list.length,top:out,me:me};
}
