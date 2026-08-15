const SHEET_ID = '1Rba07Brfx8DG550CWlS0hwg1kLTSzwClKt237jD1Tds';
const HEADERS = ['Ημερομηνία','Visitor ID','Ονοματεπώνυμο','RSVP','Άτομα','Μήνυμα / Διατροφή','Δώρο','IBAN','IRIS','Τελευταίο κλικ','Τελευταία ενημέρωση'];
function doGet(){return ContentService.createTextOutput('Wedding endpoint is active');}
function doPost(e){
  const lock=LockService.getScriptLock(); lock.waitLock(10000);
  try{
    const p=e.parameter||{}, ss=SpreadsheetApp.openById(SHEET_ID), sh=ss.getSheetByName('Απαντήσεις')||ss.insertSheet('Απαντήσεις');
    if(sh.getLastRow()===0)sh.appendRow(HEADERS);
    const id=p.visitorId||Utilities.getUuid(), data=sh.getDataRange().getValues(); let row=0;
    for(let i=data.length-1;i>0;i--)if(String(data[i][1])===id){row=i+1;break}
    if(!row){sh.appendRow([new Date(),id,'','','','','Όχι','Όχι','Όχι','',new Date()]);row=sh.getLastRow()}
    if(p.action==='rsvp'){sh.getRange(row,3,1,4).setValues([[p.name||'',p.attendance||'',p.guests||'',p.message||'']])}
    if(p.action==='gift_click'){
      const gift=p.gift||'', isNo=gift==='ΟΧΙ'; sh.getRange(row,7).setValue(isNo?'Όχι':'Ναι');
      if(gift==='IBAN')sh.getRange(row,8).setValue('Ναι'); if(gift==='IRIS')sh.getRange(row,9).setValue('Ναι');
      sh.getRange(row,10).setValue(gift); 
    }
    sh.getRange(row,11).setValue(new Date()); return ContentService.createTextOutput(JSON.stringify({ok:true})).setMimeType(ContentService.MimeType.JSON);
  } finally {lock.releaseLock()}
}
