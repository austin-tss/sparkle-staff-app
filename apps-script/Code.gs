// Sparkle Staff activity log receiver. Deploy: Deploy > New deployment > Web app,
// Execute as: Me, Who has access: Anyone. Paste the /exec URL into LOG_URL in app.js.
const SHEET_ID='18c4tjPypvgJRZBLqqywivVUlZzl2Iuv8G-AWaiosk54';
function doPost(e){
  const pre=(()=>{try{return JSON.parse(e.postData.contents||'{}')}catch(_){return {}}})();
  if(pre.action==='tryon')return tryon(pre);
  const lock=LockService.getScriptLock();lock.waitLock(10000);
  try{
    const sh=SpreadsheetApp.openById(SHEET_ID).getSheetByName('Log');
    const body=JSON.parse(e.postData.contents||'{}');
    const evs=(body.events||[]).slice(0,100);
    const seen=new Set(sh.getRange(2,7,Math.max(sh.getLastRow()-1,1),1).getValues().flat());
    const now=Utilities.formatDate(new Date(),'America/New_York','M/d/yyyy h:mm:ss a')+' ET';
    const rows=evs.filter(x=>x&&x.id&&!seen.has(x.id)).map(x=>[x.time,x.who,x.event,x.detail,x.device,x.session,x.id,now].map(v=>String(v||'').slice(0,500)));
    if(rows.length)sh.getRange(sh.getLastRow()+1,1,rows.length,8).setValues(rows);
    return ContentService.createTextOutput(JSON.stringify({ok:true,added:rows.length})).setMimeType(ContentService.MimeType.JSON);
  }finally{lock.releaseLock()}
}
function doGet(){return ContentService.createTextOutput('Sparkle Staff log endpoint OK')}
// Placeholder: virtual try-on via FASHN. Key will live in Script Properties as FASHN_API_KEY.
// Not wired yet; will need the script.external_request scope when implemented.
function tryon(body){
  const key=PropertiesService.getScriptProperties().getProperty('FASHN_API_KEY');
  const out=key?{ok:false,error:'tryon not implemented yet'}:{ok:false,error:'FASHN_API_KEY not set'};
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
