/**
 * GÜÇ ELEKTRONİĞİ LABORATUVARI TAKİP SİSTEMİ
 * Tam Sürüm (Şık E-Posta Şablonları, Dış Lab İptali, İsimli Mail Bildirimleri, Lokal Rapor)
 */

const SPREADSHEET_ID = "1B-5oZuEBnktm-UBcyt4qbWpyU28CO3zxMy6OpgCRIgk"; 
const KALITE_SS_ID = "1xFGZQSVpWkWdU6YHUNfmFMtGCv4UBCiPtuhXrFjaFNw";
const MEKANIK_SS_ID = "1MQpi0pLpTs_NLsxAon5QmJa7qG5yt7BMaF7Wr9diWjM";
const KALITE_APP_URL = "https://script.google.com/a/macros/anova.com.tr/s/AKfycbxZ4AbZyYM0YcDX3pBR3yFombp15kokE_E7PzfXuWXYFP-Z5KoK-tsGVL2nHY0yLUc/exec";

function getAppUrl() { return ScriptApp.getService().getUrl(); }
function getSpreadsheet() { return SpreadsheetApp.openById(SPREADSHEET_ID); }
function clearLabCache() { CacheService.getScriptCache().remove("guc_lab_data"); }

function parseIncomingDate(val) {
   if (!val) return new Date();
   let num = Number(val);
   if (!isNaN(num)) return new Date(num);
   return new Date(val);
}

function getAdminSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName("Yönetici Şifreleri") || ss.getSheetByName("yonetici sifreleri");
  if (!sheet) {
    const allSheets = ss.getSheets();
    for (let i = 0; i < allSheets.length; i++) {
        let sName = allSheets[i].getName().toLowerCase().trim();
        if(sName.includes("yonet") && sName.includes("sifre")) return allSheets[i];
    }
  }
  return sheet;
}

function checkIfSuperAdmin(email) {
  if (!email) return false;
  let sheet = getAdminSheet();
  if (!sheet) return false;
  let e = email.toLowerCase().trim();
  let data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase().trim() === e) {
       return (data[i][3] === true || String(data[i][3]).toLowerCase() === "true");
    }
  }
  return false;
}

function checkIfAdmin(email) {
  if (!email) return false;
  let e = email.toLowerCase().trim(), sheet = getAdminSheet();
  if (!sheet) return false;
  let data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) { if (String(data[i][0]).toLowerCase().trim() === e) return true; }
  return false;
}

function getDurumColIndex(sheet) {
  let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let idx = headers.indexOf("Durum"); 
  return idx !== -1 ? idx + 1 : 6; 
}

function updatePersonColumn(invSheet, rowIndex, personName) {
  let headers = invSheet.getRange(1, 1, 1, invSheet.getLastColumn()).getValues()[0];
  let targetCol = -1;
  for (let i = 0; i < headers.length; i++) {
    let hL = String(headers[i]).toLowerCase().trim();
    if (hL.includes("kişi") || hL.includes("personel") || hL.includes("zimmet")) { targetCol = i + 1; break; }
  }
  if (targetCol === -1) targetCol = 14; 
  invSheet.getRange(rowIndex, targetCol).setValue(personName);
}

function getInvDetails(invData, code) {
    let headers = invData[0].map(h => String(h).toLowerCase().trim());
    let idxMarka = headers.findIndex(h => h.includes('marka')); if(idxMarka===-1) idxMarka=2;
    let idxModel = headers.findIndex(h => h.includes('model')); if(idxModel===-1) idxModel=3;
    for(let i=1; i<invData.length; i++) {
        if(String(invData[i][0]).trim() === String(code).trim()) {
            return { marka: String(invData[i][idxMarka]||"").trim(), model: String(invData[i][idxModel]||"").trim() };
        }
    }
    return { marka: "", model: "" };
}

function doGet(e) {
  const userEmail = Session.getActiveUser().getEmail() || "Bilinmeyen Kullanıcı";
  if (e && e.parameter && e.parameter.action) {
    let action = e.parameter.action;
    if (action === "approve_transfer" || action === "reject_transfer") return handleTransferAction(action, e.parameter.code, e.parameter.from, e.parameter.to, userEmail, e.parameter.token);
  }
  const template = HtmlService.createTemplateFromFile('Index');
  template.userEmail = userEmail; 
  return template.evaluate().setTitle('Güç Elektroniği Laboratuvarı').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

function verifyAdminLogin(password) {
  try {
    const userEmail = Session.getActiveUser().getEmail().toLowerCase().trim(), adminSheet = getAdminSheet();
    if (!adminSheet) throw new Error("Yönetici sayfası bulunamadı!");
    const data = adminSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase().trim() === userEmail && String(data[i][1]).trim() === password) {
          let isSuper = (data[i][3] === true || String(data[i][3]).toLowerCase() === "true");
          return { success: true, role: isSuper ? "super" : "admin" };
      }
    }
    throw new Error("Hatalı şifre veya yetkisiz erişim!");
  } catch (err) { throw new Error(err.message); }
}

function getAdminStats() {
  if (!checkIfAdmin(Session.getActiveUser().getEmail())) return null;
  let stats = { quarantine: 0, outgoing: 0, calRed: 0, calYellow: 0 };
  const ss = getSpreadsheet();
  const invData = ss.getSheetByName("Envanter").getDataRange().getValues();
  let durumIdx = invData[0].findIndex(h => String(h).trim() === "Durum");
  if(durumIdx === -1) durumIdx = 5;

  for(let i=1; i<invData.length; i++) {
      let d = String(invData[i][durumIdx]).trim();
      if (d.includes("Onayı Bekliyor") || d.includes("Biriminden Devir") || d.includes("Transfer Onayı")) stats.quarantine++;
      if (d.includes("Dış Lab. Devir Bekliyor")) stats.outgoing++;
  }
  
  let cals = getUpcomingCalsServer();
  cals.forEach(c => { if(c.diffDays <= 10) stats.calRed++; else if(c.diffDays <= 30) stats.calYellow++; });
  return stats;
}

function getUpcomingCalsServer() {
  if (!checkIfAdmin(Session.getActiveUser().getEmail())) throw new Error("Yetkiniz yok!");
  let u = []; const t = new Date(); t.setHours(0,0,0,0); const nM = new Date(t); nM.setDate(t.getDate() + 30);
  try {
     let targetSS = getSpreadsheet(); let invSheet = targetSS.getSheetByName("Envanter");
     if (invSheet) {
         let data = invSheet.getDataRange().getValues();
         if (data.length > 1) {
             let headers = data[0].map(h => String(h).toLowerCase().trim());
             let idxKodu = 0, idxIsmi = 1, idxMarka = headers.findIndex(h => h.includes('marka')); if(idxMarka===-1) idxMarka=2;
             let idxModel = headers.findIndex(h => h.includes('model')); if(idxModel===-1) idxModel=3;
             let idxCal = headers.findIndex(h => h.includes('kalibrasyon') && h.includes('son')); if(idxCal===-1) idxCal=4;
             let idxArizali = headers.findIndex(h => h.includes('arızalı') || h.includes('arizali')); if(idxArizali===-1) idxArizali=6;
             let idxKalibda = headers.findIndex(h => h.includes('kalibrasyonda')); if(idxKalibda===-1) idxKalibda=7;

             for(let i=1; i<data.length; i++) {
                if (!data[i][0]) continue;
                if (String(data[i][idxArizali]).toLowerCase() === 'true' || String(data[i][idxKalibda]).toLowerCase() === 'true') continue;

                let calStr = data[i][idxCal], cDate = null;
                if (calStr instanceof Date) cDate = calStr; else if(calStr) { let p = String(calStr).split('.'); if(p.length === 3) cDate = new Date(p[2], p[1]-1, p[0]); }

                if (cDate && cDate <= nM) {
                   const cT = new Date(cDate); cT.setHours(0,0,0,0);
                   let diff = Math.round((cT.getTime() - t.getTime()) / (1000 * 60 * 60 * 24));
                   u.push({ code: String(data[i][idxKodu]).trim(), name: String(data[i][idxIsmi]).trim(), brand: String(data[i][idxMarka] || '-').trim(), model: String(data[i][idxModel] || '-').trim(), cal: Utilities.formatDate(cT, "GMT+3", "dd.MM.yyyy"), diffDays: diff });
                }
             }
         }
     }
  } catch(e) {}
  return u.sort((a,b) => a.diffDays - b.diffDays);
}

function notifyLabAdmins(targetSS_ID, subject, htmlBody) {
  try {
    let tSS = SpreadsheetApp.openById(targetSS_ID);
    let adminSheet = tSS.getSheetByName("Yönetici Şifreleri") || tSS.getSheetByName("yonetici sifreleri");
    if (!adminSheet) return;
    
    let data = adminSheet.getDataRange().getValues();
    let emails = [];
    
    for (let i = 1; i < data.length; i++) {
      let email = String(data[i][0]).trim();
      let notify = data[i][2]; 
      if (email && (notify === true || String(notify).toLowerCase() === "true")) {
        emails.push(email);
      }
    }
    
    if (emails.length > 0) {
      MailApp.sendEmail({ bcc: emails.join(","), subject: subject, htmlBody: htmlBody });
    }
  } catch (e) {}
}

function addNewDevice(deviceData) {
  if (!checkIfAdmin(Session.getActiveUser().getEmail())) throw new Error("Yetkiniz yok!");
  const invSheet = getSpreadsheet().getSheetByName("Envanter"); 
  const data = invSheet.getDataRange().getValues(), headers = data[0];
  for (let i = 1; i < data.length; i++) { if (String(data[i][0]).trim() === String(deviceData.code).trim()) throw new Error("Bu koda sahip cihaz var!"); }
  
  let baslangicDurumu = "Müsait";
  let newRow = new Array(headers.length).fill("");
  let map = { "Demirbaş Kodu": deviceData.code, "Cihaz İsmi": deviceData.name, "Cihaz Markası": deviceData.brand, "Cihaz Modeli": deviceData.model, "Son Kalibrasyon Tarihi": deviceData.calDate, "Durum": baslangicDurumu, "Arızalı": false, "Kalibrasyonda": false, "Klasörleme": deviceData.folder, "Kategori": deviceData.folder, "Kalibrasyon Tarihini Güncelleyen Kişi": Session.getActiveUser().getEmail(), "Arıza Notu": "", "Datasheet Linki": deviceData.datasheet, "Kalibrasyon Sertifikası": deviceData.cert };
  
  for(let i = 0; i < headers.length; i++) { 
    let head = String(headers[i]).trim();
    let headLower = head.toLowerCase();
    if(map[head] !== undefined) newRow[i] = map[head];
    else {
      if(headLower.includes("marka")) newRow[i] = deviceData.brand;
      else if(headLower.includes("model")) newRow[i] = deviceData.model;
      else if(headLower.includes("kalibrasyon") && headLower.includes("son")) newRow[i] = deviceData.calDate;
      else if(headLower.includes("sertifika")) newRow[i] = deviceData.cert;
      else if(headLower.includes("güncelle")) newRow[i] = Session.getActiveUser().getEmail();
      else if(headLower.includes("klasör") || headLower.includes("kategori")) newRow[i] = deviceData.folder;
    }
  }
  invSheet.appendRow(newRow); clearLabCache();
  return "Başarıyla eklendi!";
}

// 📌 GÜNCELLENDİ: PERSONEL ATAMA E-POSTA ŞABLONU (ŞIK VE SİYAH ÜST BİLGİ, SARI ÇİZGİ)
function adminAssignDevice(code, targetEmail, note) {
  if (!checkIfAdmin(Session.getActiveUser().getEmail())) throw new Error("Yetkiniz yok!");
  const ss = getSpreadsheet(), invSheet = ss.getSheetByName("Envanter"), logSheet = ss.getSheetByName("Kayitlar");
  let invData = invSheet.getDataRange().getValues(), durumCol = getDurumColIndex(invSheet), devName = "Cihaz", foundInv = false;
  let targetName = targetEmail.split('@')[0].toUpperCase();
  let details = getInvDetails(invData, code);

  for (let i = 1; i < invData.length; i++) {
    if (String(invData[i][0]).trim() === String(code).trim()) {
      invSheet.getRange(i + 1, durumCol).setValue("Kullanımda"); updatePersonColumn(invSheet, i + 1, targetName); 
      devName = String(invData[i][1]).trim(); foundInv = true; break;
    }
  }
  if (!foundInv) throw new Error("Cihaz bulunamadı!");
  
  let logData = logSheet.getDataRange().getValues();
  for (let i = logData.length - 1; i >= 1; i--) {
    if (String(logData[i][0]).trim() === String(code).trim() && String(logData[i][9]).trim() === "Kullanımda") {
       logSheet.getRange(i + 1, 9).setValue(new Date()); logSheet.getRange(i + 1, 10).setValue(`Yönetici Tarafından Devredildi -> ${targetEmail}`); break;
    }
  }
  logSheet.appendRow([String(code).trim(), devName, details.marka, details.model, targetName, targetEmail, new Date(), new Date(new Date().setFullYear(new Date().getFullYear() + 1)), "", "Kullanımda", note || "Yönetici Tarafından Atandı"]);
  clearLabCache(); 
  
  let emailHtml = `
  <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; max-width: 600px; margin: 0 auto; border: 1px solid #e9ecef; border-radius: 10px; overflow: hidden; box-shadow: 0 4px 10px rgba(0,0,0,0.05); background-color: #ffffff;">
      <div style="background-color: #212529; border-bottom: 4px solid #f4c430; padding: 20px; text-align: center;">
          <img src="https://anova.com.tr/wp-content/uploads/2022/06/Anova_Logo_Beyaz_00.webp" alt="Anova Logo" style="height: 45px; display: inline-block;">
      </div>
      <div style="padding: 30px;">
          <h2 style="color: #212529; margin-top: 0; font-size: 22px;">Yeni Cihaz Zimmeti 📦</h2>
          <p style="color: #555; font-size: 15px; line-height: 1.6;">Merhaba <b>${targetName}</b>,<br>Güç Elektroniği Laboratuvarı yöneticisi tarafından adınıza yeni bir cihaz zimmetlenmiştir.</p>
          <div style="background-color: #f8f9fa; border-left: 4px solid #f4c430; padding: 15px; margin: 20px 0; border-radius: 4px;">
              <div style="margin-bottom: 8px;"><span style="color: #6c757d; font-size: 12px; text-transform: uppercase; font-weight: bold;">Demirbaş Kodu</span><br><b style="color: #212529; font-size: 15px;">${code}</b></div>
              <div style="margin-bottom: 8px;"><span style="color: #6c757d; font-size: 12px; text-transform: uppercase; font-weight: bold;">Cihaz İsmi</span><br><b style="color: #212529; font-size: 15px;">${devName}</b></div>
              <div><span style="color: #6c757d; font-size: 12px; text-transform: uppercase; font-weight: bold;">Atama Notu</span><br><b style="color: #212529; font-size: 14px;">${note || "Yönetici Tarafından Atandı"}</b></div>
          </div>
          <p style="color: #555; font-size: 14px; line-height: 1.6; margin-bottom: 0;">Cihaz otomatik olarak üzerinize kaydedilmiştir. İşiniz bittiğinde sistem üzerinden iade etmeyi unutmayınız.</p>
      </div>
  </div>`;
  try { MailApp.sendEmail({ to: targetEmail, subject: "Yeni Cihaz Zimmeti: " + code, htmlBody: emailHtml }); } catch(e) {}
  
  return "Cihaz başarıyla personele atandı ve bilgilendirme maili gönderildi!";
}

// 📌 GÜNCELLENDİ: KARŞI LABORATUVARA GİDEN TRANSFER E-POSTA ŞABLONU (SARI ÇİZGİ)
function transferToLab(code, labKey, note) {
  if (!checkIfAdmin(Session.getActiveUser().getEmail())) throw new Error("Yetkiniz yok!");
  const ssGuc = getSpreadsheet(), invGuc = ssGuc.getSheetByName("Envanter"), dataGuc = invGuc.getDataRange().getValues();
  let rowIdx = -1, rowData = null, headersGuc = dataGuc[0];

  for (let i = 1; i < dataGuc.length; i++) {
    if (String(dataGuc[i][0]).trim() === String(code).trim()) { rowIdx = i + 1; rowData = dataGuc[i]; break; }
  }
  if (rowIdx === -1) throw new Error("Cihaz envanterde bulunamadı!");

  let currentDurum = String(rowData[getDurumColIndex(invGuc)-1]).trim();
  if(currentDurum === "Arızalı" || currentDurum === "Kalibrasyonda" || currentDurum.includes("Bekliyor")) throw new Error("Bu cihazın durumu transfere uygun değil (" + currentDurum + ").");

  let targetSS_ID = "", targetLabName = "";
  if (labKey === "kalite") { targetSS_ID = KALITE_SS_ID; targetLabName = "Kalite Birimi"; }
  else if (labKey === "mekanik") { targetSS_ID = MEKANIK_SS_ID; targetLabName = "Mekanik Lab."; }
  else throw new Error("Geçersiz laboratuvar seçimi.");

  invGuc.getRange(rowIdx, getDurumColIndex(invGuc)).setValue(`Dış Lab. Devir Bekliyor:${targetLabName}`);

  let idxM = headersGuc.findIndex(h => String(h).toLowerCase().includes("marka"));
  let idxMd = headersGuc.findIndex(h => String(h).toLowerCase().includes("model"));
  let marka = idxM !== -1 ? String(rowData[idxM]) : "";
  let model = idxMd !== -1 ? String(rowData[idxMd]) : "";

  const targetSS = SpreadsheetApp.openById(targetSS_ID), targetInv = targetSS.getSheetByName("Envanter"), headersTarget = targetInv.getDataRange().getValues()[0];
  let newRow = new Array(headersTarget.length).fill("");
  for(let i=0; i<headersTarget.length; i++) {
     let hName = String(headersTarget[i]).trim();
     let gucIdx = headersGuc.findIndex(h => String(h).trim() === hName);
     if (gucIdx !== -1) newRow[i] = rowData[gucIdx];
     if (hName.toLowerCase() === "durum") { newRow[i] = "Transfer Onayı Bekliyor"; }
  }

  let logEmail = Session.getActiveUser().getEmail();
  let senderName = logEmail.split('@')[0].replace(/\./g, ' ').toUpperCase();
  let senderLabel = senderName + " (Güç Elektroniği)";

  let notCol = headersTarget.findIndex(h => h.toLowerCase().includes("arıza notu") || h.toLowerCase().includes("ariza notu"));
  if (notCol !== -1) newRow[notCol] = "[Güç Lab. Devri] " + (note || "");

  targetInv.appendRow(newRow); 

  let logSheet = ssGuc.getSheetByName("Kayitlar");
  if(logSheet) logSheet.appendRow([String(code).trim(), String(rowData[1]).trim(), marka, model, "DIŞ BİRİM", logEmail, new Date(), new Date(), "", `Dış Lab. Devir Bekliyor:${targetLabName}`, note || "Laboratuvarlar Arası Transfer Başlatıldı"]);

  try {
      let targetLogSheet = targetSS.getSheetByName("Kayitlar");
      if(targetLogSheet) { 
          targetLogSheet.appendRow([String(code).trim(), String(rowData[1]).trim(), marka, model, senderLabel, logEmail, new Date(), new Date(), "", "Transfer Onayı Bekliyor", note || "Güç Elektroniğinden Transfer Edildi"]); 
      }
  } catch(e) {}

  let targetAdminSubject = "Yeni Laboratuvar Transferi: " + code;
  let targetAdminHtml = `
  <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; max-width: 600px; margin: 0 auto; border: 1px solid #e9ecef; border-radius: 10px; overflow: hidden; box-shadow: 0 4px 10px rgba(0,0,0,0.05); background-color: #ffffff;">
      <div style="background-color: #212529; border-bottom: 4px solid #f4c430; padding: 20px; text-align: center;">
          <img src="https://anova.com.tr/wp-content/uploads/2022/06/Anova_Logo_Beyaz_00.webp" alt="Anova Logo" style="height: 45px; display: inline-block;">
      </div>
      <div style="padding: 30px;">
          <h2 style="color: #212529; margin-top: 0; font-size: 22px;">Onay Bekleyen Transfer 📦</h2>
          <p style="color: #555; font-size: 15px; line-height: 1.6;"><b>${senderLabel}</b> tarafından laboratuvarınıza yeni bir cihaz transfer edilmiştir.</p>
          <div style="background-color: #f8f9fa; border-left: 4px solid #f4c430; padding: 15px; margin: 20px 0; border-radius: 4px;">
              <div style="margin-bottom: 8px;"><span style="color: #6c757d; font-size: 12px; text-transform: uppercase; font-weight: bold;">Demirbaş Kodu</span><br><b style="color: #212529; font-size: 15px;">${code}</b></div>
              <div style="margin-bottom: 8px;"><span style="color: #6c757d; font-size: 12px; text-transform: uppercase; font-weight: bold;">Cihaz İsmi</span><br><b style="color: #212529; font-size: 15px;">${rowData[1]}</b></div>
              <div><span style="color: #6c757d; font-size: 12px; text-transform: uppercase; font-weight: bold;">Transfer Notu</span><br><b style="color: #212529; font-size: 14px;">${note || "Belirtilmedi"}</b></div>
          </div>
          <p style="color: #555; font-size: 14px; line-height: 1.6; margin-bottom: 0;">Lütfen panelinizden <b>Onay Bekleyenler (Karantina)</b> sekmesine girerek işlemi onaylayın veya reddedin.</p>
      </div>
  </div>`;
  notifyLabAdmins(targetSS_ID, targetAdminSubject, targetAdminHtml);

  clearLabCache();
  return `Cihaz başarıyla ${targetLabName} onay havuzuna gönderildi.`;
}

function loadOutgoingTransfers() {
  const ss = getSpreadsheet(), invSheet = ss.getSheetByName("Envanter"), data = invSheet.getDataRange().getValues(), durumIdx = getDurumColIndex(invSheet) - 1;
  let results = [];
  for(let i=1; i<data.length; i++) {
     let d = String(data[i][durumIdx]).trim();
     if(d.startsWith("Dış Lab. Devir Bekliyor") || d.startsWith("Transfer Onayı Bekliyor")) results.push({ code: data[i][0], name: data[i][1], durum: d });
  }
  return results;
}

function cancelOutgoingTransfer(code) {
  if (!checkIfAdmin(Session.getActiveUser().getEmail())) throw new Error("Yetkiniz yok!");
  const ss = getSpreadsheet(), invSheet = ss.getSheetByName("Envanter"), logSheet = ss.getSheetByName("Kayitlar");
  let durumCol = getDurumColIndex(invSheet), data = invSheet.getDataRange().getValues(), durumVal = "", devName = "";

  for(let i=1; i<data.length; i++){
     if(String(data[i][0]).trim() === String(code).trim()){
        durumVal = String(data[i][durumCol-1]).trim(); devName = data[i][1]; invSheet.getRange(i+1, durumCol).setValue("Müsait"); break;
     }
  }

  if(durumVal.includes("Kalite") || durumVal.includes("Mekanik")) {
      let targetSS_ID = durumVal.includes("Kalite") ? KALITE_SS_ID : MEKANIK_SS_ID;
      try {
          let tSheet = SpreadsheetApp.openById(targetSS_ID).getSheetByName("Envanter"), tData = tSheet.getDataRange().getValues();
          for(let k=1; k<tData.length; k++){ if(String(tData[k][0]).trim() === String(code).trim()) { tSheet.deleteRow(k+1); break; } }
      } catch(e) {}
  }
  let details = getInvDetails(data, code);
  logSheet.appendRow([code, devName, details.marka, details.model, "Sistem", Session.getActiveUser().getEmail(), new Date(), new Date(), "", "Transfer İptal", "Dış laboratuvar transferi geri çekildi."]);
  clearLabCache(); return "Transfer başarıyla iptal edildi ve cihaz envantere geri döndü.";
}

function processDueReservations(ss) {
  const rezSheet = ss.getSheetByName("RezervasyonListesi");
  if (!rezSheet) return false;
  const rezData = rezSheet.getDataRange().getValues();
  if (rezData.length <= 1) return false;
  
  const invSheet = ss.getSheetByName("Envanter"), logSheet = ss.getSheetByName("Kayitlar"), invData = invSheet.getDataRange().getValues(), durumCol = getDurumColIndex(invSheet);
  let now = new Date(), rowsToDelete = [], updated = false;
  
  for (let i = rezData.length - 1; i >= 1; i--) {
    let rowStart = new Date(rezData[i][4]);
    if (rowStart <= now) {
      let code = String(rezData[i][0]).trim(), devName = rezData[i][1], userName = rezData[i][2], userEmail = rezData[i][3], rowEnd = new Date(rezData[i][5]), note = (rezData[i][6] || "") + " [Otomatik: Rzv. Başladı]";
      let isMusait = false, invRowIdx = -1;
      
      for (let j = 1; j < invData.length; j++) {
        if (String(invData[j][0]).trim() === code) { invRowIdx = j + 1; if (String(invData[j][durumCol - 1]).trim() === "Müsait") isMusait = true; break; }
      }
      
      if (isMusait && invRowIdx !== -1) {
        invSheet.getRange(invRowIdx, durumCol).setValue("Kullanımda"); updatePersonColumn(invSheet, invRowIdx, userName); 
        let details = getInvDetails(invData, code);
        logSheet.appendRow([code, devName, details.marka, details.model, userName, userEmail, rowStart, rowEnd, "", "Kullanımda", note]);
        rowsToDelete.push(i + 1); updated = true;
      }
    }
  }
  rowsToDelete.sort((a, b) => b - a).forEach(rowIdx => { rezSheet.deleteRow(rowIdx); });
  return updated;
}

function getLabData() {
  try {
    const ss = getSpreadsheet();
    if(processDueReservations(ss)) clearLabCache();

    const cache = CacheService.getScriptCache();
    const cachedData = cache.get("guc_lab_data");
    let results = [];
    
    if (cachedData) {
        results = JSON.parse(cachedData);
    } else {
        const invSheet = ss.getSheetByName("Envanter");
        const logSheet = ss.getSheetByName("Kayitlar");
        const rezSheet = ss.getSheetByName("RezervasyonListesi");
        const invData = invSheet.getDataRange().getValues();
        const headers = invData[0] || [];
        
        let activeLogs = {}, upcomingRez = {};
        let hL = headers.map(h => String(h).toLowerCase().trim());
        let idxKodu = 0, idxIsmi = 1;
        let idxMarka = hL.findIndex(h => h.includes('marka')); if(idxMarka===-1) idxMarka=2;
        let idxModel = hL.findIndex(h => h.includes('model')); if(idxModel===-1) idxModel=3;
        let idxCal = hL.findIndex(h => h.includes('kalibrasyon') && h.includes('son')); if(idxCal===-1) idxCal=4;
        let idxDurum = hL.findIndex(h => h === 'durum'); if(idxDurum===-1) idxDurum=5;
        let idxArizali = hL.findIndex(h => h.includes('arızalı') || h.includes('arizali')); if(idxArizali===-1) idxArizali=6;
        let idxKalibda = hL.findIndex(h => h.includes('kalibrasyonda')); if(idxKalibda===-1) idxKalibda=7;
        let idxKlasor = hL.findIndex(h => h.includes('klasör') || h.includes('kategori')); if(idxKlasor===-1) idxKlasor=8;
        let idxNote = hL.findIndex(h => h.includes('arıza notu') || h.includes('ariza notu')); if(idxNote===-1) idxNote=10;
        let idxDs = hL.findIndex(h => h.includes('datasheet')); if(idxDs===-1) idxDs=11;
        let idxCert = hL.findIndex(h => h.includes('sertifika')); if(idxCert===-1) idxCert=12;

        if (logSheet) {
          let logData = logSheet.getDataRange().getValues();
          for (let i = 1; i < logData.length; i++) {
            if (String(logData[i][9] || "").trim() === "Kullanımda") { 
              activeLogs[String(logData[i][0]).trim()] = { kisi: logData[i][4], email: logData[i][5], alindi: logData[i][6] ? Utilities.formatDate(new Date(logData[i][6]), "GMT+3", "HH:mm dd.MM.yyyy") : "-", teslim: logData[i][7] ? Utilities.formatDate(new Date(logData[i][7]), "GMT+3", "HH:mm dd.MM.yyyy") : "-" };
            }
          }
        }
        
        if (rezSheet) {
          let rData = rezSheet.getDataRange().getValues();
          for(let i = 1; i < rData.length; i++) {
            let code = String(rData[i][0]).trim();
            if(!upcomingRez[code]) upcomingRez[code] = [];
            upcomingRez[code].push({ kisi: rData[i][2], email: rData[i][3], start: rData[i][4] ? Utilities.formatDate(new Date(rData[i][4]), "GMT+3", "HH:mm dd.MM.yyyy") : "", end: rData[i][5] ? Utilities.formatDate(new Date(rData[i][5]), "GMT+3", "HH:mm dd.MM.yyyy") : "" });
          }
        }

        for (let i = 1; i < invData.length; i++) {
            let row = invData[i];
            if(!row || row.length === 0 || String(row[0]).trim() === "") continue; 
            
            let drm = String(row[idxDurum]).trim();
            if (drm.includes("Onayı Bekliyor") || drm.includes("Transfer Onayı Bekliyor") || drm.includes("Biriminden Devir") || drm.includes("Dış Lab. Devir")) continue;

            let obj = {};
            for (let j = 0; j < headers.length; j++) { 
              let hName = String(headers[j]).trim();
              obj[hName] = row[j] instanceof Date ? Utilities.formatDate(row[j], "GMT+3", "HH:mm dd.MM.yyyy") : (row[j] != null ? String(row[j]).trim() : ""); 
            }
            
            let invCode = String(row[idxKodu]).trim();
            obj['ActiveDetails'] = activeLogs[invCode] || null; 
            obj['UpcomingRez'] = upcomingRez[invCode] || null; 
            obj['ArizaNotu'] = row.length > idxNote ? String(row[idxNote]).trim() : "";
            obj['_searchString'] = Object.values(obj).join(" ").toLowerCase(); 
            
            obj['Demirbaş Kodu'] = invCode; obj['Cihaz İsmi'] = String(row[idxIsmi] || "").trim(); obj['Marka'] = String(row[idxMarka] || "").trim(); obj['Model'] = String(row[idxModel] || "").trim();
            obj['Son Kalibrasyon'] = row[idxCal] instanceof Date ? Utilities.formatDate(row[idxCal], "GMT+3", "dd.MM.yyyy") : String(row[idxCal] || "").trim();
            obj['Durum'] = drm; obj['Arızalı'] = headers.length > idxArizali ? String(row[idxArizali]).trim() : "false"; obj['Kalibrasyonda'] = headers.length > idxKalibda ? String(row[idxKalibda]).trim() : "false";
            
            let folderVal = String(row[idxKlasor] || "").trim();
            obj['Klasörleme'] = (folderVal !== "" && folderVal.toUpperCase() !== "FALSE" && folderVal.toUpperCase() !== "TRUE") ? folderVal : "Genel";
            obj['Datasheet Linki'] = headers.length > idxDs ? String(row[idxDs]).trim() : ""; obj['Sertifika Linki'] = headers.length > idxCert ? String(row[idxCert]).trim() : "";
            
            results.push(obj);
        }
        cache.put("guc_lab_data", JSON.stringify(results), 900);
    }
    return results;
  } catch(e) { throw new Error(e.message); }
}

function reserveDevice(reserveData) {
  try {
    const ss = getSpreadsheet();
    const userEmail = Session.getActiveUser().getEmail() || "Bilinmiyor";
    const userName = userEmail.includes('@') ? userEmail.split('@')[0].toUpperCase() : "TEST"; 
    
    let sDate = parseIncomingDate(reserveData.startDate);
    let rDate = parseIncomingDate(reserveData.returnDate);
    let now = new Date();
    
    let diffMinutes = (sDate.getTime() - now.getTime()) / (1000 * 60);
    let finalNote = (reserveData.project ? `[Proje: ${reserveData.project.trim()}] ` : "") + (reserveData.note || "");
    
    if (diffMinutes > 15) {
      const rezSheet = ss.getSheetByName("RezervasyonListesi");
      rezSheet.appendRow([reserveData.code, reserveData.name, userName, userEmail, sDate, rDate, finalNote]);
      clearLabCache(); 
      return "İleri tarihli rezervasyonunuz başarıyla kaydedildi!";
    } else {
      const inventorySheet = ss.getSheetByName("Envanter");
      const logSheet = ss.getSheetByName("Kayitlar");
      let invData = inventorySheet.getDataRange().getValues();
      let durumCol = getDurumColIndex(inventorySheet);
      let details = getInvDetails(invData, reserveData.code);

      for (let i = 1; i < invData.length; i++) { 
        if (String(invData[i][0]).trim() == String(reserveData.code).trim()) { 
          inventorySheet.getRange(i + 1, durumCol).setValue("Kullanımda"); 
          updatePersonColumn(inventorySheet, i + 1, userName); break; 
        } 
      }
      logSheet.appendRow([reserveData.code, reserveData.name, details.marka, details.model, userName, userEmail, now, rDate, "", "Kullanımda", finalNote]);
      clearLabCache(); 
      return "Cihaz başarıyla üzerinize kaydedildi.";
    }
  } catch(e) { throw new Error(e.message); }
}

function findDeviceForMulti(code) {
   const invData = getSpreadsheet().getSheetByName("Envanter").getDataRange().getValues();
   let headLower = String(invData[0]).toLowerCase().trim().split(',');
   let idxArizali = headLower.findIndex(h => h.includes('arızalı') || h.includes('arizali')); if(idxArizali===-1) idxArizali=6;
   let idxKalibda = headLower.findIndex(h => h.includes('kalibrasyonda')); if(idxKalibda===-1) idxKalibda=7;

   for(let i=1; i<invData.length; i++) {
      if(String(invData[i][0]).trim() === String(code).trim()) {
         let isA = String(invData[i][idxArizali]).toLowerCase() === 'true';
         let isC = String(invData[i][idxKalibda]).toLowerCase() === 'true';
         if(isA || isC) throw new Error("Bu cihaz arızalı veya kalibrasyonda, eklenemez!");
         return { code: String(invData[i][0]).trim(), name: String(invData[i][1]).trim() };
      }
   }
   throw new Error("Cihaz bulunamadı: " + code);
}

function bulkReserveDevices(bulkData) {
  try {
    const ss = getSpreadsheet(), userEmail = Session.getActiveUser().getEmail() || "Bilinmiyor";
    const userName = userEmail.includes('@') ? userEmail.split('@')[0].toUpperCase() : "TEST";
    
    let sDate = parseIncomingDate(bulkData.startDate);
    let rDate = parseIncomingDate(bulkData.returnDate);
    let now = new Date();
    let diffMinutes = (sDate.getTime() - now.getTime()) / (1000 * 60);
    
    let finalNote = (bulkData.project ? `[Proje: ${bulkData.project.trim()}] ` : "") + (bulkData.note || "");
    let successCount = 0;

    const invData = ss.getSheetByName("Envanter").getDataRange().getValues();
    let headers = invData[0].map(h => String(h).toLowerCase().trim());
    let idxMarka = headers.findIndex(h => h.includes('marka')); if(idxMarka===-1) idxMarka=2;
    let idxModel = headers.findIndex(h => h.includes('model')); if(idxModel===-1) idxModel=3;

    if (diffMinutes > 15) {
       const rezSheet = ss.getSheetByName("RezervasyonListesi");
       let devNames = {};
       for(let i=1; i<invData.length; i++) devNames[String(invData[i][0]).trim()] = String(invData[i][1]).trim();
       
       bulkData.codes.forEach(code => {
         let codeStr = String(code).trim();
         let devName = devNames[codeStr] || "Cihaz";
         rezSheet.appendRow([codeStr, devName, userName, userEmail, sDate, rDate, finalNote]);
         successCount++;
       });
       clearLabCache(); return `${successCount} adet cihaz için ileri tarihli rezervasyon oluşturuldu.`;
    } else {
      const inventorySheet = ss.getSheetByName("Envanter"), logSheet = ss.getSheetByName("Kayitlar");
      let durumCol = getDurumColIndex(inventorySheet);

      bulkData.codes.forEach(code => {
        let codeStr = String(code).trim(), devName = "Cihaz", marka = "", model = "", found = false;
        for (let i = 1; i < invData.length; i++) {
          if (String(invData[i][0]).trim() === codeStr) {
            inventorySheet.getRange(i + 1, durumCol).setValue("Kullanımda"); updatePersonColumn(inventorySheet, i + 1, userName); 
            devName = String(invData[i][1]).trim(); marka = String(invData[i][idxMarka]||"").trim(); model = String(invData[i][idxModel]||"").trim(); found = true; break;
          }
        }
        if (found) { logSheet.appendRow([codeStr, devName, marka, model, userName, userEmail, now, rDate, "", "Kullanımda", finalNote]); successCount++; }
      });
      clearLabCache(); return `${successCount} adet cihaz başarıyla üzerinize kaydedildi.`;
    }
  } catch(e) { throw new Error(e.message); }
}

function bulkReturnDevices(assetCodes, returnNote) {
  try {
    const ss = getSpreadsheet(), logSheet = ss.getSheetByName("Kayitlar"), invSheet = ss.getSheetByName("Envanter");
    const logData = logSheet.getDataRange().getValues(), invData = invSheet.getDataRange().getValues(), durumCol = getDurumColIndex(invSheet);

    let successCount = 0;
    assetCodes.forEach(code => {
      let codeStr = String(code).trim(), logRowIndex = -1;
      for (let i = logData.length - 1; i >= 1; i--) { if (String(logData[i][0]).trim() === codeStr && String(logData[i][9]).trim() === "Kullanımda") { logRowIndex = i + 1; break; } }

      if (logRowIndex !== -1) {
        for (let j = 1; j < invData.length; j++) {
          if (String(invData[j][0]).trim() === codeStr) { invSheet.getRange(j + 1, durumCol).setValue("Müsait"); updatePersonColumn(invSheet, j + 1, ""); break; }
        }
        logSheet.getRange(logRowIndex, 9).setValue(new Date()); 
        logSheet.getRange(logRowIndex, 10).setValue("Teslim Edildi");
        if(returnNote) { let extNote = logSheet.getRange(logRowIndex, 11).getValue(); logSheet.getRange(logRowIndex, 11).setValue(extNote ? extNote + " | " + returnNote : returnNote); }
        successCount++;
      }
    });

    clearLabCache(); return `${successCount} adet cihaz başarıyla iade edildi.`;
  } catch(e) { throw new Error(e.message); }
}

function cancelReservation(code) {
   const ss = getSpreadsheet(), rezSheet = ss.getSheetByName("RezervasyonListesi"), data = rezSheet.getDataRange().getValues();
   const userEmail = Session.getActiveUser().getEmail().toLowerCase().trim(); let isAdmin = checkIfAdmin(userEmail);

   for(let i = data.length - 1; i >= 1; i--) {
      if(String(data[i][0]).trim() === String(code).trim()) {
         if(isAdmin || String(data[i][3]).toLowerCase().trim() === userEmail) { rezSheet.deleteRow(i + 1); clearLabCache(); return "Rezervasyon başarıyla iptal edildi."; } 
         else { throw new Error("Bu rezervasyonu iptal etme yetkiniz bulunmuyor."); }
      }
   }
   throw new Error("Açık rezervasyon bulunamadı.");
}

// 📌 GÜNCELLENDİ: PERSONEL-PERSONEL ARASI DEVİR TALEBİ E-POSTASI (YEŞİL/SARI)
function initiateTransfer(code, targetEmail) {
  const userEmail = Session.getActiveUser().getEmail();
  let logSheet = getSpreadsheet().getSheetByName("Kayitlar"), logData = logSheet.getDataRange().getValues(), devName = "Cihaz";
  for (let i = logData.length - 1; i >= 1; i--) { if (String(logData[i][0]).trim() === String(code).trim() && String(logData[i][9]).trim() === "Kullanımda" && logData[i][5] === userEmail) { devName = logData[i][1]; break; } }
  
  const transferToken = Utilities.getUuid();
  CacheService.getScriptCache().put('trf_' + code + '_' + targetEmail.toLowerCase(), transferToken, 21600); // 6 saat geçerli, tek kullanımlık

  let appUrl = `${getAppUrl()}?action=approve_transfer&code=${encodeURIComponent(code)}&from=${encodeURIComponent(userEmail)}&to=${encodeURIComponent(targetEmail)}&token=${transferToken}`;
  let rejectUrl = `${getAppUrl()}?action=reject_transfer&code=${encodeURIComponent(code)}&from=${encodeURIComponent(userEmail)}&to=${encodeURIComponent(targetEmail)}&token=${transferToken}`;
  
  let emailHtml = `
  <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; max-width: 600px; margin: 0 auto; border: 1px solid #e9ecef; border-radius: 10px; overflow: hidden; box-shadow: 0 4px 10px rgba(0,0,0,0.05); background-color: #ffffff;">
      <div style="background-color: #212529; border-bottom: 4px solid #f4c430; padding: 20px; text-align: center;">
          <img src="https://anova.com.tr/wp-content/uploads/2022/06/Anova_Logo_Beyaz_00.webp" alt="Anova Logo" style="height: 45px; display: inline-block;">
      </div>
      <div style="padding: 30px;">
          <h2 style="color: #212529; margin-top: 0; font-size: 22px;">Zimmet Devir Talebi 🤝</h2>
          <p style="color: #555; font-size: 15px; line-height: 1.6;"><b>${userEmail}</b>, size aşağıdaki cihazı devretmek istiyor.</p>
          <div style="background-color: #f8f9fa; border-left: 4px solid #f4c430; padding: 15px; margin: 20px 0; border-radius: 4px;">
              <div style="margin-bottom: 8px;"><span style="color: #6c757d; font-size: 12px; text-transform: uppercase; font-weight: bold;">Demirbaş Kodu</span><br><b style="color: #212529; font-size: 15px;">${code}</b></div>
              <div><span style="color: #6c757d; font-size: 12px; text-transform: uppercase; font-weight: bold;">Cihaz İsmi</span><br><b style="color: #212529; font-size: 15px;">${devName}</b></div>
          </div>
          <p style="color: #555; font-size: 14px; line-height: 1.6; margin-bottom: 20px;">Cihazı fiziksel olarak teslim aldıysanız lütfen onaylayın.</p>
          <div style="text-align: center;">
              <a href="${appUrl}" style="background-color: #198754; color: white; text-decoration: none; font-weight: bold; padding: 12px 25px; border-radius: 6px; display: inline-block; margin-right: 10px;">✅ TESLİM ALDIM</a>
              <a href="${rejectUrl}" style="background-color: #dc3545; color: white; text-decoration: none; font-weight: bold; padding: 12px 25px; border-radius: 6px; display: inline-block;">❌ ALMADIM</a>
          </div>
      </div>
  </div>`;

  try { MailApp.sendEmail({ to: targetEmail, subject: "Zimmet Devir Talebi: " + code, htmlBody: emailHtml }); } catch(e) {}
  return "Devir onayı için mail gönderildi.";
}

function handleTransferAction(action, code, fromEmail, toEmail, currentUser, token) {
  if (currentUser !== toEmail) return HtmlService.createHtmlOutput("<h2>Yetkisiz işlem!</h2>");

  const cacheKey = 'trf_' + code + '_' + toEmail.toLowerCase();
  const cache = CacheService.getScriptCache();
  const validToken = cache.get(cacheKey);
  if (!validToken || !token || validToken !== token) {
    return HtmlService.createHtmlOutput("<h2 style='color:#c62828;'>Bu link geçersiz veya süresi dolmuş. Lütfen gönderen kişiden devri tekrar başlatmasını isteyin.</h2>");
  }
  cache.remove(cacheKey); // Tek kullanımlık: aynı link ikinci kez işlem yapamaz

  if (action === "reject_transfer") return HtmlService.createHtmlOutput("<h3 style='color:red;'>Reddettiniz.</h3>");

  const ss = getSpreadsheet(), logSheet = ss.getSheetByName("Kayitlar"), invSheet = ss.getSheetByName("Envanter");
  let logData = logSheet.getDataRange().getValues(), invData = invSheet.getDataRange().getValues(), devName = "Cihaz", found = false;
  let details = getInvDetails(invData, code);

  for (let i = logData.length - 1; i >= 1; i--) {
      if (String(logData[i][0]).trim() === String(code).trim() && String(logData[i][9]).trim() === "Kullanımda" && logData[i][5] === fromEmail) {
          logSheet.getRange(i + 1, 9).setValue(new Date()); logSheet.getRange(i + 1, 10).setValue(`Devredildi -> ${toEmail}`); 
          devName = logData[i][1]; found = true; break;
      }
  }
  if(!found) return HtmlService.createHtmlOutput("<h2>Zimmet bulunamadı.</h2>");
  
  let toName = toEmail.split('@')[0].toUpperCase();
  for (let j = 1; j < invData.length; j++) { if(String(invData[j][0]).trim() === String(code).trim()) { updatePersonColumn(invSheet, j + 1, toName); break; } }
  
  logSheet.appendRow([code, devName, details.marka, details.model, toName, toEmail, new Date(), new Date(new Date().setFullYear(new Date().getFullYear() + 1)), "", "Kullanımda", "Devir Yoluyla Alındı"]);
  clearLabCache(); return HtmlService.createHtmlOutput("<h2 style='color:green;'>Devir Başarılı!</h2>");
}

function reportFault(code, note) {
  const ss = getSpreadsheet(), invSheet = ss.getSheetByName("Envanter");
  const headers = invSheet.getDataRange().getValues()[0], invData = invSheet.getDataRange().getValues();
  let arizaliIdx = headers.findIndex(h => String(h).trim().includes("Arızalı")), notIdx = headers.findIndex(h => String(h).trim().includes("Arıza Notu"));
  if(arizaliIdx === -1) arizaliIdx = 6; if(notIdx === -1) notIdx = 10;
  
  for (let i = 1; i < invData.length; i++) {
    if (String(invData[i][0]).trim() === String(code).trim()) { invSheet.getRange(i + 1, arizaliIdx + 1).setValue(true); invSheet.getRange(i + 1, notIdx + 1).setValue(note); break; }
  }
  clearLabCache(); return "Bildirim iletildi.";
}

function returnDevice(assetCode, returnNote) { return bulkReturnDevices([assetCode], returnNote); }

function getUserDevices(targetUser) {
    let searchTarget = targetUser ? String(targetUser).toLowerCase().trim() : (Session.getActiveUser().getEmail() || "").toLowerCase();
    let allDevices = [], ss = getSpreadsheet(), logData = ss.getSheetByName("Kayitlar").getDataRange().getValues(), rezSheet = ss.getSheetByName("RezervasyonListesi");
    const foundCodes = new Set();
    
    for (let i = logData.length - 1; i >= 1; i--) {
      let logCode = String(logData[i][0]).trim();
      if (String(logData[i][9]).trim() === "Kullanımda" && (String(logData[i][5]).toLowerCase().includes(searchTarget) || String(logData[i][4]).toLowerCase().includes(searchTarget)) && !foundCodes.has(logCode)) {
        allDevices.push({ code: logCode, name: logData[i][1], date: logData[i][6] ? Utilities.formatDate(new Date(logData[i][6]), "GMT+3", "HH:mm dd.MM.yyyy") : "-", returnDate: logData[i][7] ? Utilities.formatDate(new Date(logData[i][7]), "GMT+3", "HH:mm dd.MM.yyyy") : "-", type: 'active' }); foundCodes.add(logCode);
      }
    }
    
    if (rezSheet) {
      const rezData = rezSheet.getDataRange().getValues();
      for (let i = 1; i < rezData.length; i++) {
        if ((String(rezData[i][3]).toLowerCase().includes(searchTarget) || String(rezData[i][2]).toLowerCase().includes(searchTarget))) {
          allDevices.push({ code: String(rezData[i][0]), name: rezData[i][1], date: rezData[i][4] ? Utilities.formatDate(new Date(rezData[i][4]), "GMT+3", "HH:mm dd.MM.yyyy") : "-", returnDate: rezData[i][5] ? Utilities.formatDate(new Date(rezData[i][5]), "GMT+3", "HH:mm dd.MM.yyyy") : "-", note: rezData[i][6] || "", type: 'reservation' });
        }
      }
    }
    return allDevices;
}

function updateCalibration(code, newDateStr, updaterEmail) {
  if (!checkIfAdmin(Session.getActiveUser().getEmail())) throw new Error("Yetkiniz yok!");
  const invSheet = getSpreadsheet().getSheetByName("Envanter"), headers = invSheet.getDataRange().getValues()[0], data = invSheet.getDataRange().getValues();
  let calColIdx = headers.findIndex(h => String(h).trim().includes("Son Kalibrasyon")), updColIdx = headers.findIndex(h => String(h).trim().includes("Güncelleyen"));
  if (calColIdx === -1) calColIdx = 4; if (updColIdx === -1) updColIdx = 9;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(code).trim()) { 
      invSheet.getRange(i + 1, calColIdx + 1).setValue(newDateStr); invSheet.getRange(i + 1, updColIdx + 1).setValue(updaterEmail || Session.getActiveUser().getEmail()); 
      clearLabCache(); return "Kalibrasyon Tarihi Güncellendi!"; 
    }
  }
  throw new Error("Cihaz Bulunamadı!");
}

function getPlannedReservations() {
  if (!checkIfAdmin(Session.getActiveUser().getEmail())) throw new Error("Yetkiniz yok!");
  const ss = getSpreadsheet(), sheet = ss.getSheetByName("RezervasyonListesi"); if (!sheet) return [];
  const data = sheet.getDataRange().getValues(); let results = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) results.push({ code: data[i][0], name: data[i][1], user: data[i][2], email: data[i][3], start: data[i][4] ? Utilities.formatDate(new Date(data[i][4]), "GMT+3", "HH:mm dd.MM.yyyy") : "-", end: data[i][5] ? Utilities.formatDate(new Date(data[i][5]), "GMT+3", "HH:mm dd.MM.yyyy") : "-" });
  }
  return results;
}

function updateDeviceStatus(code, status, note) {
  if (!checkIfAdmin(Session.getActiveUser().getEmail())) throw new Error("Yetkiniz yok!");
  const invSheet = getSpreadsheet().getSheetByName("Envanter"), headers = invSheet.getDataRange().getValues()[0], data = invSheet.getDataRange().getValues();
  let arizaliIdx = headers.findIndex(h => String(h).trim().includes("Arızalı")), kalibIdx = headers.findIndex(h => String(h).trim().includes("Kalibrasyonda")), durumIdx = headers.findIndex(h => String(h).trim().includes("Durum")), notIdx = headers.findIndex(h => String(h).trim().includes("Arıza Notu"));
  if(arizaliIdx === -1) arizaliIdx = 6; if(kalibIdx === -1) kalibIdx = 7; if(durumIdx === -1) durumIdx = 5; if(notIdx === -1) notIdx = 10;
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(code).trim()) {
      if (arizaliIdx !== -1) invSheet.getRange(i + 1, arizaliIdx + 1).setValue(status === "Arızalı");
      if (kalibIdx !== -1) invSheet.getRange(i + 1, kalibIdx + 1).setValue(status === "Kalibrasyonda");
      if (durumIdx !== -1 && (status === "Müsait" || status === "Karantina Onaylandı")) invSheet.getRange(i + 1, durumIdx + 1).setValue("Müsait");
      if (notIdx !== -1) invSheet.getRange(i + 1, notIdx + 1).setValue(note || ""); 
      clearLabCache(); return `Güncellendi!`;
    }
  }
  throw new Error("Bulunamadı!");
}

function getDeviceHistory(query) {
  try {
      const logSheet = getSpreadsheet().getSheetByName("Kayitlar"); 
      if(!logSheet) return [];
      const data = logSheet.getDataRange().getValues(); 
      let results = [];
      let q = query ? String(query).toLowerCase().trim() : "";

      for (let i = data.length - 1; i >= 1; i--) { 
          if (!data[i] || !data[i][0]) continue;
          
          let c = String(data[i][0] || "").trim();
          let n = String(data[i][1] || "").trim();
          let m = String(data[i][2] || "").trim();
          let md = String(data[i][3] || "").trim();
          let u = String(data[i][4] || "").trim();
          let em = String(data[i][5] || "").trim();
          
          if (q === "" || c.toLowerCase().includes(q) || n.toLowerCase().includes(q) || u.toLowerCase().includes(q) || m.toLowerCase().includes(q) || md.toLowerCase().includes(q) || em.toLowerCase().includes(q)) {
              let alis = data[i][6] ? (data[i][6] instanceof Date ? Utilities.formatDate(data[i][6], "GMT+3", "HH:mm dd.MM.yyyy") : String(data[i][6])) : "-";
              let tahmini = data[i][7] ? (data[i][7] instanceof Date ? Utilities.formatDate(data[i][7], "GMT+3", "HH:mm dd.MM.yyyy") : String(data[i][7])) : "-";
              let gercek = data[i][8] ? (data[i][8] instanceof Date ? Utilities.formatDate(data[i][8], "GMT+3", "HH:mm dd.MM.yyyy") : String(data[i][8])) : "Hala Kullanımda";
              
              results.push({ 
                  code: c, name: n, marka: m, model: md, kisi: u, email: em,
                  alis: alis, tahmini: tahmini, gercek: gercek, 
                  durum: String(data[i][9] || ""), not: String(data[i][10] || "") 
              });
          }
      }
      return results;
  } catch(e) { return []; }
}

function getAdminList() {
  if (!checkIfSuperAdmin(Session.getActiveUser().getEmail())) throw new Error("Sadece kurucu görebilir!");
  const adminSheet = getAdminSheet(); if(!adminSheet) return [];
  const data = adminSheet.getDataRange().getValues(); let list = [];
  for (let i = 1; i < data.length; i++) {
    let e = String(data[i][0]).toLowerCase().trim();
    if (e) {
        let isSuper = (data[i][3] === true || String(data[i][3]).toLowerCase() === "true");
        list.push({ email: e, password: String(data[i][1]).trim(), role: isSuper ? "super" : "admin", notify: (data[i][2] === true || String(data[i][2]).toLowerCase() === "true") });
    }
  }
  return list;
}

function removeAdmin(email) {
  if (!checkIfSuperAdmin(Session.getActiveUser().getEmail())) throw new Error("Yetkiniz yok!");
  let adminSheet = getAdminSheet(), target = email.toLowerCase().trim(), data = adminSheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 3; i--) { if (String(data[i][0]).toLowerCase().trim() === target) { adminSheet.deleteRow(i + 1); return "Silindi!"; } }
  throw new Error("Bulunamadı veya Kurucu Silinemez!");
}

function toggleAdminNotification(email, isChecked) {
  if (!checkIfSuperAdmin(Session.getActiveUser().getEmail())) throw new Error("Yetkiniz yok!");
  let adminSheet = getAdminSheet(), data = adminSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) { if (String(data[i][0]).toLowerCase().trim() === email.toLowerCase().trim()) { adminSheet.getRange(i + 1, 3).setValue(isChecked); return "Güncellendi!"; } }
  throw new Error("Bulunamadı!");
}

function toggleSuperAdmin(email, isChecked) {
  if (!checkIfSuperAdmin(Session.getActiveUser().getEmail())) throw new Error("Yetkiniz yok!");
  let adminSheet = getAdminSheet(), data = adminSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase().trim() === email.toLowerCase().trim()) {
          adminSheet.getRange(i + 1, 4).setValue(isChecked); return "Yetki Güncellendi!";
      }
  }
  throw new Error("Bulunamadı!");
}

function addOrUpdateAdmin(email, password, isSuper) {
  if (!checkIfSuperAdmin(Session.getActiveUser().getEmail())) throw new Error("Yetkiniz yok!");
  email = email.toLowerCase().trim();
  let adminSheet = getAdminSheet(), data = adminSheet.getDataRange().getValues(), updated = false, targetRowIndex = -1;
  for (let i = 1; i < data.length; i++) { 
     if (String(data[i][0]).toLowerCase().trim() === email) { adminSheet.getRange(i + 1, 2).setValue(password); adminSheet.getRange(i + 1, 4).setValue(isSuper); updated = true; break; } 
  }
  if(!updated) { adminSheet.appendRow([email, password, false, isSuper]); targetRowIndex = adminSheet.getLastRow(); adminSheet.getRange(targetRowIndex, 3, 1, 2).insertCheckboxes(); }
  return updated ? "Yönetici Şifresi Güncellendi!" : "Yeni Yönetici Eklendi!";
}

function loadQuarantine() {
  const ss = getSpreadsheet(), invSheet = ss.getSheetByName("Envanter"), data = invSheet.getDataRange().getValues(), headers = data[0];
  const logSheet = ss.getSheetByName("Kayitlar"); let logData = logSheet ? logSheet.getDataRange().getValues() : [];

  let durumIdx = headers.findIndex(h => String(h).trim().includes("Durum")); if (durumIdx === -1) durumIdx = 5;
  let idxName = 1, idxBrand = headers.findIndex(h => String(h).toLowerCase().trim().includes("marka")), idxModel = headers.findIndex(h => String(h).toLowerCase().trim().includes("model")), idxCal = headers.findIndex(h => String(h).toLowerCase().trim().includes("son kalibrasyon"));
  if(idxBrand===-1) idxBrand = 2; if(idxModel===-1) idxModel = 3; if(idxCal===-1) idxCal = 4;

  let pending = [];
  for (let i = 1; i < data.length; i++) {
    let d = String(data[i][durumIdx]).trim();
    if (d.includes("Onayı Bekliyor") || d.includes("Transfer Onayı Bekliyor") || d.includes("Biriminden Devir")) {
        let calVal = data[i][idxCal];
        if (Object.prototype.toString.call(calVal) === '[object Date]') { calVal = Utilities.formatDate(calVal, "GMT+3", "yyyy-MM-dd"); } 
        else if (calVal) { let p = String(calVal).split('.'); if(p.length === 3) calVal = `${p[2]}-${p[1]}-${p[0]}`; else calVal = String(calVal); }

        let code = String(data[i][0]).trim(), lastUser = "Sistem / Dış Laboratuvar", lastAction = "Bekliyor", lastNote = "-";
        
        for (let j = logData.length - 1; j >= 1; j--) {
            if (String(logData[j][0]).trim() === code) {
                lastUser = String(logData[j][4] || logData[j][5] || "Sistem"); lastAction = String(logData[j][9] || "İşlem"); lastNote = String(logData[j][10] || "-"); break;
            }
        }
        pending.push({ code: String(code), name: String(data[i][idxName] || ""), brand: String(data[i][idxBrand] || ""), model: String(data[i][idxModel] || ""), calDate: String(calVal || ""), durum: String(d), contextUser: String(lastUser), contextAction: String(lastAction), contextNote: String(lastNote) });
    }
  }
  return pending;
}

function approveQuarantine(code, updatedData) { 
  if (!checkIfAdmin(Session.getActiveUser().getEmail())) throw new Error("Yetkiniz yok!");
  const ss = getSpreadsheet(), invSheet = ss.getSheetByName("Envanter"), logSheet = ss.getSheetByName("Kayitlar");
  let durumCol = getDurumColIndex(invSheet), invData = invSheet.getDataRange().getValues(), headers = invData[0];
  let currentDurum = "", invRowIdx = -1, devName="Cihaz", details = getInvDetails(invData, code);
  let userEmail = Session.getActiveUser().getEmail(), userName = userEmail.split('@')[0].replace(/\./g, ' ').toUpperCase() + " (Güç Elektroniği)";

  for (let i = 1; i < invData.length; i++) {
    if (String(invData[i][0]).trim() === String(code).trim()) { 
       currentDurum = String(invData[i][durumCol-1]).trim(); invRowIdx = i + 1; devName = invData[i][1]; 
       if (updatedData) {
         let idxName = headers.findIndex(h => String(h).trim() === "Cihaz İsmi"); if(idxName!==-1 && updatedData.name) { invSheet.getRange(invRowIdx, idxName+1).setValue(updatedData.name); devName = updatedData.name; }
         let idxBrand = headers.findIndex(h => String(h).toLowerCase().trim().includes("marka")); if(idxBrand!==-1 && updatedData.brand) { invSheet.getRange(invRowIdx, idxBrand+1).setValue(updatedData.brand); details.marka = updatedData.brand; }
         let idxModel = headers.findIndex(h => String(h).toLowerCase().trim().includes("model")); if(idxModel!==-1 && updatedData.model) { invSheet.getRange(invRowIdx, idxModel+1).setValue(updatedData.model); details.model = updatedData.model; }
         let idxCal = headers.findIndex(h => String(h).toLowerCase().trim().includes("son kalibrasyon")); 
         if(idxCal!==-1 && updatedData.calDate) { let p = updatedData.calDate.split('-'); if(p.length === 3) invSheet.getRange(invRowIdx, idxCal+1).setValue(`${p[2]}.${p[1]}.${p[0]}`); }
      }
       break; 
    }
  }

  if (currentDurum.includes("İade Onayı Bekliyor")) {
     invSheet.getRange(invRowIdx, durumCol).setValue("Müsait"); updatePersonColumn(invSheet, invRowIdx, ""); 
     let logData = logSheet.getDataRange().getValues();
     for (let i = logData.length - 1; i >= 1; i--) {
        if (String(logData[i][0]).trim() === String(code).trim() && String(logData[i][9]).trim() === "İade Onayına Gönderildi") { logSheet.getRange(i + 1, 9).setValue(new Date()); logSheet.getRange(i + 1, 10).setValue("Teslim Edildi"); break; }
     }
     clearLabCache(); return "Cihaz teslim alındı ve sisteme Müsait olarak kaydedildi.";
  } else if (currentDurum.includes("IT Onayı Bekliyor")) {
     invSheet.getRange(invRowIdx, durumCol).setValue("Kullanımda");
     let kisiCol = invData[0].findIndex(h => h.toLowerCase().includes("kişi") || h.toLowerCase().includes("zimmet")), targetName = kisiCol !== -1 ? String(invData[invRowIdx-1][kisiCol]).trim() : "Bilinmiyor";
     logSheet.appendRow([code, devName, details.marka, details.model, targetName, targetName.toLowerCase()+"@anova.com.tr", new Date(), new Date(new Date().setFullYear(new Date().getFullYear() + 1)), "", "Kullanımda", "Zimmet Onaylandı"]);
     clearLabCache(); return "Transfer onaylandı. Cihaz artık personelin zimmetindedir.";
  } else {
     invSheet.getRange(invRowIdx, durumCol).setValue("Müsait"); 
     if (currentDurum.includes("Biriminden Devir") || currentDurum.includes("Transfer Onayı Bekliyor")) {
          let targetSS_ID = null;
          let logData = logSheet.getDataRange().getValues();
          for (let j = logData.length - 1; j >= 1; j--) {
              if (String(logData[j][0]).trim() === String(code).trim()) {
                  let who = String(logData[j][4]).toUpperCase(); let noteStr = String(logData[j][10]).toUpperCase();
                  if (who.includes("KALİTE") || noteStr.includes("KALİTE") || who.includes("KALITE")) { targetSS_ID = KALITE_SS_ID; break; }
                  if (who.includes("MEKANİK") || noteStr.includes("MEKANİK") || who.includes("MEKANIK")) { targetSS_ID = MEKANIK_SS_ID; break; }
              }
          }
          if(targetSS_ID) {
              try {
                  let tSheet = SpreadsheetApp.openById(targetSS_ID).getSheetByName("Envanter"), tData = tSheet.getDataRange().getValues();
                  for(let k = 1; k < tData.length; k++){ if(String(tData[k][0]).trim() === String(code).trim()){ tSheet.deleteRow(k + 1); break; } }
                  let tLog = SpreadsheetApp.openById(targetSS_ID).getSheetByName("Kayitlar");
                  if(tLog) tLog.appendRow([code, devName, details.marka, details.model, "Güç Elektroniği", userEmail, new Date(), new Date(), "", "Transfer Onaylandı", "Karşı laboratuvar cihazı kabul etti."]);
              } catch(e) {}
          }
          logSheet.appendRow([code, devName, details.marka, details.model, userName, userEmail, new Date(), new Date(), "", "Transfer Onaylandı", "Cihaz Müsait olarak havuza alındı."]);
     } else {
          logSheet.appendRow([code, devName, details.marka, details.model, userName, userEmail, new Date(), new Date(), "", "Onaylandı", "Cihaz onaylanarak envantere dahil edildi."]);
     }
     clearLabCache(); return "Cihaz onaylandı ve havuza eklendi.";
  }
}

// 📌 GÜNCELLENDİ: REDDEDİLME E-POSTA ŞABLONU (KIRMIZI ÇİZGİ, SIFIR ÇERÇEVELİ LOGO)
function rejectQuarantine(code, rejectNote) {
  let userEmail = Session.getActiveUser().getEmail();
  let userName = userEmail.split('@')[0].replace(/\./g, ' ').toUpperCase() + " (Güç Elektroniği)";

  if (!checkIfAdmin(userEmail)) throw new Error("Yetkiniz yok!");
  const ss = getSpreadsheet(), invSheet = ss.getSheetByName("Envanter"), durumCol = getDurumColIndex(invSheet), invData = invSheet.getDataRange().getValues();

  for (let i = 1; i < invData.length; i++) {
    if (String(invData[i][0]).trim() === String(code).trim()) {
      let curDurum = String(invData[i][durumCol-1]).trim(), devName = String(invData[i][1]).trim();
      let logSheet = ss.getSheetByName("Kayitlar"), details = getInvDetails(invData, code);
      
      if (curDurum.includes("Biriminden Devir") || curDurum.includes("Transfer Onayı Bekliyor")) {
          let originLab = null; let targetSS_ID = null;
          let logData = logSheet.getDataRange().getValues();
          for (let j = logData.length - 1; j >= 1; j--) {
              if (String(logData[j][0]).trim() === String(code).trim()) {
                  let who = String(logData[j][4]).toUpperCase(); let noteStr = String(logData[j][10]).toUpperCase();
                  if (who.includes("KALİTE") || noteStr.includes("KALİTE") || who.includes("KALITE")) { originLab = "Kalite Birimi"; targetSS_ID = KALITE_SS_ID; break; }
                  if (who.includes("MEKANİK") || noteStr.includes("MEKANİK") || who.includes("MEKANIK")) { originLab = "Mekanik Lab."; targetSS_ID = MEKANIK_SS_ID; break; }
              }
          }

          if (targetSS_ID) {
              try {
                  let eSS = SpreadsheetApp.openById(targetSS_ID), eSheet = eSS.getSheetByName("Envanter"), eData = eSheet.getDataRange().getValues(), eDurumCol = getDurumColIndex(eSheet);
                  for(let k=1; k<eData.length; k++) {
                      if(String(eData[k][0]).trim() === String(code).trim()) {
                          eSheet.getRange(k+1, eDurumCol).setValue("İade Onayı Bekliyor");
                          let eLog = eSS.getSheetByName("Kayitlar");
                          if(eLog) eLog.appendRow([code, devName, details.marka, details.model, "Güç Elektroniği", userEmail, new Date(), new Date(), "", "Transfer Reddedildi", `Ret Sebebi: ${rejectNote || "Belirtilmedi"}`]);
                          
                          let rejectAdminSubject = "Transfer Reddedildi: " + code;
                          let rejectAdminHtml = `
                          <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; max-width: 600px; margin: 0 auto; border: 1px solid #e9ecef; border-radius: 10px; overflow: hidden; box-shadow: 0 4px 10px rgba(0,0,0,0.05); background-color: #ffffff;">
                              <div style="background-color: #212529; border-bottom: 4px solid #dc3545; padding: 20px; text-align: center;">
                                  <img src="https://anova.com.tr/wp-content/uploads/2022/06/Anova_Logo_Beyaz_00.webp" alt="Anova Logo" style="height: 45px; display: inline-block;">
                              </div>
                              <div style="padding: 30px;">
                                  <h2 style="color: #212529; margin-top: 0; font-size: 22px;">Transfer Reddedildi ❌</h2>
                                  <p style="color: #555; font-size: 15px; line-height: 1.6;">Göndermiş olduğunuz cihaz <b>Güç Elektroniği</b> departmanı tarafından reddedildi ve onay havuzunuza iade statüsünde geri döndü.</p>
                                  <div style="background-color: #fff5f5; border-left: 4px solid #dc3545; padding: 15px; margin: 20px 0; border-radius: 4px;">
                                      <div style="margin-bottom: 8px;"><span style="color: #6c757d; font-size: 12px; text-transform: uppercase; font-weight: bold;">Demirbaş Kodu</span><br><b style="color: #212529; font-size: 15px;">${code}</b></div>
                                      <div style="margin-bottom: 8px;"><span style="color: #6c757d; font-size: 12px; text-transform: uppercase; font-weight: bold;">Cihaz İsmi</span><br><b style="color: #212529; font-size: 15px;">${devName}</b></div>
                                     <div><span style="color: #6c757d; font-size: 12px; text-transform: uppercase; font-weight: bold;">Ret Sebebi</span><br><b style="color: #dc3545; font-size: 14px;">${rejectNote || "Belirtilmedi"}</b></div>
                                  </div>
                                  <p style="color: #555; font-size: 14px; line-height: 1.6; margin-bottom: 0;">Lütfen panelinizden iadeyi onaylayarak cihazı tekrar envanterinize (Müsait) alın.</p>
                              </div>
                          </div>`;
                          notifyLabAdmins(targetSS_ID, rejectAdminSubject, rejectAdminHtml);
                          break;
                      }
                  }
              } catch(e) {}
          }
          invSheet.deleteRow(i + 1); clearLabCache(); return "Cihaz reddedildi ve ilk laboratuvara iade edildi.";
      } else if (curDurum.includes("İade Onayı Bekliyor")) {
          invSheet.getRange(i+1, durumCol).setValue("Kullanımda");
          let logData = logSheet.getDataRange().getValues();
          for (let j = logData.length - 1; j >= 1; j--) {
            if (String(logData[j][0]).trim() === String(code).trim() && String(logData[j][9]).trim() === "İade Onayına Gönderildi") {
               logSheet.getRange(j + 1, 10).setValue("Kullanımda"); 
               let oldNote = logSheet.getRange(j + 1, 11).getValue();
               logSheet.getRange(j + 1, 11).setValue(oldNote + " | İade Reddedildi: " + rejectNote); break;
            }
          }
          clearLabCache(); return "İade reddedildi, cihaz personelin üzerinde bırakıldı.";
      } else {
          invSheet.deleteRow(i + 1); 
          logSheet.appendRow([code, devName, details.marka, details.model, userName, userEmail, new Date(), new Date(), "", "Reddedildi", `Kayıt/Devir Reddedildi. Sebep: ${rejectNote || "Belirtilmedi"}`]);
          clearLabCache(); return "İşlem reddedildi ve cihaz kaydı silindi.";
      }
    }
  }
  throw new Error("Cihaz bulunamadı!");
}

function getLabStatistics() {
  if (!checkIfSuperAdmin(Session.getActiveUser().getEmail())) throw new Error("Bu alanı sadece Süper Adminler görebilir!");
  const ss = getSpreadsheet(), logSheet = ss.getSheetByName("Kayitlar"), invSheet = ss.getSheetByName("Envanter");
  if (!logSheet || !invSheet) return { deviceStats: [], userStats: [] };
  
  const invData = invSheet.getDataRange().getValues(), headers = invData[0].map(h => String(h).toLowerCase().trim());
  let idxMarka = headers.findIndex(h => h.includes('marka')); if (idxMarka === -1) idxMarka = 2;
  let idxModel = headers.findIndex(h => h.includes('model')); if (idxModel === -1) idxModel = 3;
  
  let brandModelMap = {};
  for (let i = 1; i < invData.length; i++) {
     let code = String(invData[i][0]).trim();
     if (code) {
         let marka = String(invData[i][idxMarka] || "").trim(), model = String(invData[i][idxModel] || "").trim(), displayName = "";
         if (marka && model) displayName = marka + " " + model; else if (marka) displayName = marka; else if (model) displayName = model;
         brandModelMap[code] = displayName.toUpperCase();
     }
  }
  
  const data = logSheet.getDataRange().getValues();
  let deviceCount = {}, userCount = {}, totalBorrows = 0;
  
  for (let i = 1; i < data.length; i++) {
    let devCode = String(data[i][0]).trim(), devRawName = String(data[i][1]).trim(), userName = String(data[i][4]).trim(), action = String(data[i][9]).trim(); 
    let devName = brandModelMap[devCode] ? brandModelMap[devCode] : devRawName;
    let finalLabel = devCode ? (devCode + " - " + devName.toUpperCase()) : devName.toUpperCase();
    
    if (devName && userName && !userName.toLowerCase().includes("sistem") && !userName.toLowerCase().includes("güç")) {
       if (action === "Kullanımda" || action.includes("Teslim") || action.includes("Devredildi") || action.includes("İade")) {
           deviceCount[finalLabel] = (deviceCount[finalLabel] || 0) + 1; userCount[userName] = (userCount[userName] || 0) + 1; totalBorrows++;
       }
    }
  }
  
  let devArr = Object.keys(deviceCount).map(k => { let count = deviceCount[k]; return { label: k, count: count, percent: totalBorrows > 0 ? ((count / totalBorrows) * 100).toFixed(1) : 0 }; }).sort((a, b) => b.count - a.count).slice(0, 10);
  let userArr = Object.keys(userCount).map(k => { return { label: k, count: userCount[k] }; }).sort((a, b) => b.count - a.count).slice(0, 10);
  return { deviceStats: devArr, userStats: userArr };
}
