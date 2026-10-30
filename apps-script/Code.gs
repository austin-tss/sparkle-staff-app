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
// Virtual try-on via FASHN v1.6. Key lives in Script Properties as FASHN_API_KEY.
const J=o=>ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
function tryonFolder_(){const it=DriveApp.getFoldersByName('Sparkle Try-On');return it.hasNext()?it.next():DriveApp.createFolder('Sparkle Try-On')}
function dataUrlBlob_(d,name){const m=/^data:([^;]+);base64,(.*)$/.exec(d||'');if(!m)return null;return Utilities.newBlob(Utilities.base64Decode(m[2]),m[1],name)}
function tryon(body){
  const key=PropertiesService.getScriptProperties().getProperty('FASHN_API_KEY');
  if(!key)return J({ok:false,error:'Try-on not set up yet'});
  const person=body.person,garment=body.garment;
  if(!/^data:image\//.test(person||'')||!/^data:image\//.test(garment||''))return J({ok:false,error:'Need a customer photo and an item photo'});
  const cat=['auto','tops','bottoms','one-pieces'].includes(body.category)?body.category:'auto';
  const H={Authorization:'Bearer '+key};
  try{
    const r=UrlFetchApp.fetch('https://api.fashn.ai/v1/run',{method:'post',contentType:'application/json',headers:H,muteHttpExceptions:true,
      payload:JSON.stringify({model_name:'tryon-v1.6',inputs:{model_image:person,garment_image:garment,category:cat,mode:'balanced',output_format:'jpeg',return_base64:true}})});
    const run=JSON.parse(r.getContentText()||'{}');
    if(r.getResponseCode()>=300||!run.id)return J({ok:false,error:'Try-on service error: '+(run.error&&(run.error.message||run.error)||r.getResponseCode())});
    let st={},t0=Date.now();
    while(Date.now()-t0<90000){
      Utilities.sleep(2500);
      st=JSON.parse(UrlFetchApp.fetch('https://api.fashn.ai/v1/status/'+run.id,{headers:H,muteHttpExceptions:true}).getContentText()||'{}');
      if(st.status==='completed'||st.status==='failed'||st.status==='canceled')break;
    }
    if(st.status!=='completed'||!(st.output||[])[0])return J({ok:false,error:st.error?(st.error.message||st.error.name||String(st.error)):'Try-on took too long, try again'});
    let out=st.output[0];
    if(!/^data:/.test(out)){const b=UrlFetchApp.fetch(out).getBlob();out='data:'+b.getContentType()+';base64,'+Utilities.base64Encode(b.getBytes())}
    const stamp=Utilities.formatDate(new Date(),'America/New_York','yyyy-MM-dd HHmmss');
    const who=String(body.who||'').slice(0,60),item=String(body.item||'').slice(0,100);
    try{const f=tryonFolder_();const tag=stamp+' '+(item||'item');
      [[person,'customer'],[out,'result']].forEach(([d,n])=>{const b=dataUrlBlob_(d,tag+' '+n+'.jpg');if(b)f.createFile(b)})}catch(e){}
    try{SpreadsheetApp.openById(SHEET_ID).getSheetByName('Log').appendRow([Utilities.formatDate(new Date(),'America/New_York','M/d/yyyy h:mm:ss a')+' ET',who,'tryon',(item||'item')+' · '+cat,'',String(body.session||''),'tryon-'+run.id,Utilities.formatDate(new Date(),'America/New_York','M/d/yyyy h:mm:ss a')+' ET'])}catch(e){}
    return J({ok:true,image:out});
  }catch(e){return J({ok:false,error:'Try-on failed: '+e.message})}
}
