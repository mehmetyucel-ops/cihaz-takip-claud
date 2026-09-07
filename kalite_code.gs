/**
 * KALİTE VE ZİMMET YÖNETİMİ TAKİP SİSTEMİ
 * Tam Sürüm (Şık E-Posta Şablonları, Dış Lab İptali, İsimli Mail Bildirimleri, TDK İmla Düzeltmesi)
 */

const SPREADSHEET_ID = "1xFGZQSVpWkWdU6YHUNfmFMtGCv4UBCiPtuhXrFjaFNw"; 
const GUC_SS_ID = "1B-5oZuEBnktm-UBcyt4qbWpyU28CO3zxMy6OpgCRIgk";
const MEKANIK_SS_ID = "1MQpi0pLpTs_NLsxAon5QmJa7qG5yt7BMaF7Wr9diWjM";
const KALITE_APP_URL = "https://script.google.com/a/macros/anova.com.tr/s/AKfycbxZ4AbZyYM0YcDX3pBR3yFombp15kokE_E7PzfXuWXYFP-Z5KoK-tsGVL2nHY0yLUc/exec";

function getAppUrl() { return ScriptApp.getService().getUrl(); }
function getSpreadsheet() { return SpreadsheetApp.openById(SPREADSHEET_ID); }
function clearLabCache() { CacheService.getScriptCache().remove("kalite_lab_data"); }
function parseIncomingDate(val) { if (!val) return new Date(); let num = Number(val); if (!isNaN(num)) return new Date(num); return new Date(val); }

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
  let sheet = getAdminSheet(); if (!sheet) return false;
  let e = email.toLowerCase().trim(), data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) { if (String(data[i][0]).toLowerCase().trim() === e) return (data[i][3] === true || String(data[i][3]).toLowerCase() === "true"); }
  return false;
}

function checkIfAdmin(email) {
  if (!email) return false;
  let e = email.toLowerCase().trim(), sheet = getAdminSheet(); if (!sheet) return false;
  let data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) { if (String(data[i][0]).toLowerCase().trim() === e) return true; }
  return false;
}

function getDurumColIndex(sheet) {
  let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let idx = headers.findIndex(h => String(h).trim() === "Durum"); return idx !== -1 ? idx + 1 : 7; 
}

function updatePersonColumn(invSheet, rowIndex, personName) {
  let headers = invSheet.getRange(1, 1, 1, invSheet.getLastColumn()).getValues()[0], targetCol = -1;
  for (let i = 0; i < headers.length; i++) { let hL = String(headers[i]).toLowerCase().trim(); if (hL.includes("kişi") || hL.includes("personel") || hL.includes("zimmet")) { targetCol = i + 1; break; } }
  if (targetCol === -1) targetCol = 14; 
  invSheet.getRange(rowIndex, targetCol).setValue(personName);
}

function getInvDetails(invData, code) {
    let headers = invData[0].map(h => String(h).toLowerCase().trim());
    let idxMarka = headers.findIndex(h => h.includes('marka')); if(idxMarka===-1) idxMarka=2;
    let idxModel = headers.findIndex(h => h.includes('model')); if(idxModel===-1) idxModel=3;
    for(let i=1; i<invData.length; i++) { if(String(invData[i][0]).trim() === String(code).trim()) return { marka: String(invData[i][idxMarka]||"").trim(), model: String(invData[i][idxModel]||"").trim() }; }
    return { marka: "", model: "" };
}

function doGet(e) {
  const userEmail = Session.getActiveUser().getEmail() || "Bilinmeyen Kullanıcı";
  if (e && e.parameter && e.parameter.action) { if (e.parameter.action === "accept_quality_transfer") return acceptQualityTransfer(e.parameter.code, e.parameter.email, userEmail, e.parameter.token); }
  const template = HtmlService.createTemplateFromFile('Index'); template.userEmail = userEmail; 
  return template.evaluate().setTitle('Kalite & Zimmet Yönetimi').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

function acceptQualityTransfer(code, expectedEmail, currentUserEmail, token) {
    if (currentUserEmail.toLowerCase() !== expectedEmail.toLowerCase()) return HtmlService.createHtmlOutput("<h2 style='color:red; text-align:center; font-family:sans-serif; margin-top:50px;'>Hata: Bu cihazı onaylama yetkiniz yok!</h2>");
    const cacheKey = 'trf_' + code + '_' + expectedEmail.toLowerCase();
    const cache = CacheService.getScriptCache();
    const validToken = cache.get(cacheKey);
    if (!validToken || !token || validToken !== token) {
        return HtmlService.createHtmlOutput("<h2 style='color:#c62828; text-align:center; font-family:sans-serif; margin-top:50px;'>Bu link geçersiz veya süresi dolmuş.</h2>");
    }
    cache.remove(cacheKey);
    try {
        const ss = getSpreadsheet(), invSheet = ss.getSheetByName("Envanter"), logSheet = ss.getSheetByName("Kayitlar");
        const invData = invSheet.getDataRange().getValues(), durumCol = getDurumColIndex(invSheet);
        let found = false, devName = "Cihaz", details = getInvDetails(invData, code);

        for (let i = 1; i < invData.length; i++) { if (String(invData[i][0]).trim() === String(code).trim()) { invSheet.getRange(i + 1, durumCol).setValue("Kullanımda"); devName = invData[i][1]; found = true; break; } }
        if (found) {
            let userName = currentUserEmail.split('@')[0].toUpperCase();
            logSheet.appendRow([code, devName, details.marka, details.model, userName, currentUserEmail, new Date(), new Date(new Date().setFullYear(new Date().getFullYear() + 1)), "", "Kullanımda", "Zimmet Fiziksel Olarak Teslim Alındı"]);
            clearLabCache(); return HtmlService.createHtmlOutput("<h2 style='color:green; text-align:center; font-family:sans-serif; margin-top:50px;'>✅ Teslimat Onaylandı! Cihaz artık zimmetinizdedir. Ekranı kapatabilirsiniz.</h2>");
        } else { return HtmlService.createHtmlOutput("<h2 style='color:orange; text-align:center; font-family:sans-serif; margin-top:50px;'>Cihaz bulunamadı veya zaten onaylanmış.</h2>"); }
    } catch(e) { return HtmlService.createHtmlOutput(`<h2>Hata oluştu: ${e.message}</h2>`); }
}

function verifyAdminLogin(password) {
  try {
    const userEmail = Session.getActiveUser().getEmail().toLowerCase().trim(), adminSheet = getAdminSheet();
    if (!adminSheet) throw new Error("Yönetici sayfası bulunamadı!");
    const data = adminSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase().trim() === userEmail && String(data[i][1]).trim() === password) { let isSuper = (data[i][3] === true || String(data[i][3]).toLowerCase() === "true"); return { success: true, role: isSuper ? "super" : "admin" }; }
    }
    throw new Error("Hatalı şifre veya yetkisiz erişim!");
  } catch (err) { throw new Error(err.message); }
}

function getAdminStats() {
  if (!checkIfAdmin(Session.getActiveUser().getEmail())) return null;
  let stats = { quarantine: 0, outgoing: 0, calRed: 0, calYellow: 0 };
  const ss = getSpreadsheet(), invData = ss.getSheetByName("Envanter").getDataRange().getValues();
  let durumIdx = invData[0].findIndex(h => String(h).trim() === "Durum"); if(durumIdx === -1) durumIdx = 6;

  for(let i=1; i<invData.length; i++) {
      let d = String(invData[i][durumIdx]).trim();
      if (d.includes("Onayı Bekliyor") || d.includes("Transfer Onayı Bekliyor") || d.includes("Biriminden Devir")) stats.quarantine++;
      if (d.includes("Dış Lab. Devir Bekliyor")) stats.outgoing++;
  }
  let cals = getUpcomingCalsServer();
  cals.forEach(c => { if(c.diffDays <= 10) stats.calRed++; else if(c.diffDays <= 30) stats.calYellow++; });
  return stats;
}

function getUpcomingCalsServer() {
  if (!checkIfAdmin(Session.getActiveUser().getEmail())) throw new Error("Yetkiniz yok!");
  let u = []; const t = new Date(); t.setHours(0,0,0,0); const nM = new Date(t); nM.setDate(t.getDate() + 30); 
  
  const labsToScan = [
      { id: SPREADSHEET_ID, name: "Kalite" },
      { id: GUC_SS_ID, name: "Güç Elektroniği" },
      { id: MEKANIK_SS_ID, name: "Mekanik Lab." }
  ];

  labsToScan.forEach(lab => {
      try {
         let targetSS = SpreadsheetApp.openById(lab.id);
         let invSheet = targetSS.getSheetByName("Envanter");
         if (invSheet) {
             let data = invSheet.getDataRange().getValues();
             if (data.length > 1) {
                 let headers = data[0].map(h => String(h).toLowerCase().trim());
                 let idxKodu = 0, idxIsmi = 1, idxMarka = headers.findIndex(h => h.includes('marka')), idxModel = headers.findIndex(h => h.includes('model')), idxCal = headers.findIndex(h => h.includes('kalibrasyon') && h.includes('son')), idxArizali = headers.findIndex(h => h.includes('arızalı') || h.includes('arizali')), idxKalibda = headers.findIndex(h => h.includes('kalibrasyonda'));
                 if(idxMarka===-1) idxMarka=2; if(idxModel===-1) idxModel=3; if(idxCal===-1) idxCal=4; if(idxArizali===-1) idxArizali=6; if(idxKalibda===-1) idxKalibda=7;

                 for(let i=1; i<data.length; i++) {
                    if (!data[i][0]) continue;
                    if (String(data[i][idxArizali]).toLowerCase() === 'true' || String(data[i][idxKalibda]).toLowerCase() === 'true') continue; 

                    let calStr = data[i][idxCal], cDate = null;
                    if (calStr instanceof Date) cDate = calStr; else if(calStr) { let p = String(calStr).split('.'); if(p.length === 3) cDate = new Date(p[2], p[1]-1, p[0]); }

                    if (cDate && cDate <= nM) {
                       const cT = new Date(cDate); cT.setHours(0,0,0,0);
                       let diff = Math.round((cT.getTime() - t.getTime()) / (1000 * 60 * 60 * 24));
                       u.push({ code: String(data[i][idxKodu]).trim(), name: String(data[i][idxIsmi]).trim(), brand: String(data[i][idxMarka] || '-').trim(), model: String(data[i][idxModel] || '-').trim(), cal: Utilities.formatDate(cT, "GMT+3", "dd.MM.yyyy"), labName: lab.name, diffDays: diff });
                    }
                 }
             }
         }
      } catch(e) {}
  });

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
  const invSheet = getSpreadsheet().getSheetByName("Envanter"), data = invSheet.getDataRange().getValues(), headers = data[0];
  for (let i = 1; i < data.length; i++) { if (String(data[i][0]).trim() === String(deviceData.code).trim()) throw new Error("Bu koda sahip cihaz var!"); }
  
  let newRow = new Array(headers.length).fill("");
  let map = { "Demirbaş Kodu": deviceData.code, "Cihaz İsmi": deviceData.name, "Cihaz Markası": deviceData.brand, "Cihaz Modeli": deviceData.model, "Son Kalibrasyon Tarihi": deviceData.calDate, "Durum": "Müsait", "Arızalı": false, "Kalibrasyonda": false, "Klasörleme": deviceData.folder, "Kategori": deviceData.folder, "Kalibrasyon Tarihini Güncelleyen Kişi": Session.getActiveUser().getEmail(), "Arıza Notu": "", "Datasheet Linki": deviceData.datasheet, "Kalibrasyon Sertifikası": deviceData.cert };
  
  for(let i = 0; i < headers.length; i++) { 
    let head = String(headers[i]).trim(), headLower = head.toLowerCase();
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
  invSheet.appendRow(newRow); clearLabCache(); return "Cihaz Kalite envanterine eklendi!";
}

function adminAssignDevice(code, targetEmail, note) {
  if (!checkIfAdmin(Session.getActiveUser().getEmail())) throw new Error("Yetkiniz yok!");
  const ss = getSpreadsheet(), invSheet = ss.getSheetByName("Envanter"), logSheet = ss.getSheetByName("Kayitlar");
  let invData = invSheet.getDataRange().getValues(), durumCol = getDurumColIndex(invSheet), devName = "Cihaz", foundInv = false;
  let targetName = targetEmail.split('@')[0].toUpperCase(), details = getInvDetails(invData, code);

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
       logSheet.getRange(i + 1, 9).setValue(new Date()); logSheet.getRange(i + 1, 10).setValue(`Yönetici Devri -> ${targetEmail}`); break;
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
          <p style="color: #555; font-size: 15px; line-height: 1.6;">Merhaba <b>${targetName}</b>,<br>Kalite departmanı yöneticisi tarafından adınıza yeni bir cihaz zimmetlenmiştir.</p>
          <div style="background-color: #f8f9fa; border-left: 4px solid #f4c430; padding: 15px; margin: 20px 0; border-radius: 4px;">
              <div style="margin-bottom: 8px;"><span style="color: #6c757d; font-size: 12px; text-transform: uppercase; font-weight: bold;">Demirbaş Kodu</span><br><b style="color: #212529; font-size: 15px;">${code}</b></div>
              <div style="margin-bottom: 8px;"><span style="color: #6c757d; font-size: 12px; text-transform: uppercase; font-weight: bold;">Cihaz İsmi</span><br><b style="color: #212529; font-size: 15px;">${devName}</b></div>
              <div><span style="color: #6c757d; font-size: 12px; text-transform: uppercase; font-weight: bold;">Atama Notu</span><br><b style="color: #212529; font-size: 14px;">${note || "Yönetici Tarafından Atandı"}</b></div>
          </div>
          <p style="color: #555; font-size: 14px; line-height: 1.6; margin-bottom: 0;">Cihaz otomatik olarak üzerinize kaydedilmiştir. İşiniz bittiğinde sistem üzerinden iade etmeyi unutmayınız.</p>
      </div>
  </div>`;
  try { MailApp.sendEmail({ to: targetEmail, subject: "Yeni Cihaz Zimmeti: " + code, htmlBody: emailHtml }); } catch(e) {}
  
  return "Cihaz başarıyla personele zimmetlendi!";
}

function transferToLab(code, labKey, note) {
  if (!checkIfAdmin(Session.getActiveUser().getEmail())) throw new Error("Yetkiniz yok!");
  
  const ss = getSpreadsheet(), invSheet = ss.getSheetByName("Envanter"), data = invSheet.getDataRange().getValues(), headers = data[0];
  let rowIdx = -1, rowData = null;

  for (let i = 1; i < data.length; i++) { if (String(data[i][0]).trim() === String(code).trim()) { rowIdx = i + 1; rowData = data[i]; break; } }
  if (rowIdx === -1) throw new Error("Cihaz Kalite envanterinde bulunamadı!");

  let durumCol = getDurumColIndex(invSheet);
  let currentDurum = String(rowData[durumCol-1]).trim();
  if(currentDurum === "Arızalı" || currentDurum === "Kalibrasyonda" || currentDurum.includes("Bekliyor")) throw new Error("Bu cihazın durumu transfere uygun değil (" + currentDurum + ").");

  let targetSS_ID = "", targetLabName = "";
  if (labKey === "guc_elektronigi") { targetSS_ID = GUC_SS_ID; targetLabName = "Güç Elektroniği"; }
  else if (labKey === "mekanik") { targetSS_ID = MEKANIK_SS_ID; targetLabName = "Mekanik Lab."; }
  else { throw new Error("Seçilen laboratuvar desteklenmiyor."); }

  invSheet.getRange(rowIdx, durumCol).setValue(`Dış Lab. Devir Bekliyor:${targetLabName}`);
  
  let idxM = headers.findIndex(h => String(h).toLowerCase().includes("marka"));
  let idxMd = headers.findIndex(h => String(h).toLowerCase().includes("model"));
  let marka = idxM !== -1 ? String(rowData[idxM]) : "";
  let model = idxMd !== -1 ? String(rowData[idxMd]) : "";

  const targetSS = SpreadsheetApp.openById(targetSS_ID), targetInv = targetSS.getSheetByName("Envanter"), targetHeaders = targetInv.getDataRange().getValues()[0];
  let newRow = new Array(targetHeaders.length).fill("");
  for(let i=0; i<targetHeaders.length; i++) {
     let hName = String(targetHeaders[i]).trim(), sourceIdx = headers.findIndex(h => String(h).trim() === hName);
     if (sourceIdx !== -1) newRow[i] = rowData[sourceIdx];
     if (hName.toLowerCase() === "durum") newRow[i] = "Transfer Onayı Bekliyor";
  }
  
  let notCol = targetHeaders.findIndex(h => h.toLowerCase().includes("arıza notu") || h.toLowerCase().includes("ariza notu"));
  if (notCol !== -1) newRow[notCol] = "[Kalite Devri] " + (note || "");
  
  targetInv.appendRow(newRow);
  
  let logEmail = Session.getActiveUser().getEmail();
  let senderName = logEmail.split('@')[0].replace(/\./g, ' ').toUpperCase();
  let senderLabel = senderName + " (Kalite Birimi)";

  let logSheet = ss.getSheetByName("Kayitlar");
  logSheet.appendRow([code, rowData[1], marka, model, "DIŞ BİRİM", logEmail, new Date(), new Date(), "", `Dış Lab. Devir Bekliyor:${targetLabName}`, note || "Laboratuvara transfer başlatıldı"]);
  
  try {
      let targetLogSheet = targetSS.getSheetByName("Kayitlar");
      if(targetLogSheet) { 
          targetLogSheet.appendRow([String(code).trim(), String(rowData[1]).trim(), marka, model, senderLabel, logEmail, new Date(), new Date(), "", "Transfer Onayı Bekliyor", note || "Kalite Biriminden Transfer Edildi"]); 
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
  return `Cihaz başarıyla ${targetLabName} onay havuzuna gönderildi. Karşı taraf onayladığında Kalite'den tamamen düşecektir.`;
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

  if(durumVal.includes("Güç") || durumVal.includes("Mekanik")) {
      let targetSS_ID = durumVal.includes("Güç") ? GUC_SS_ID : MEKANIK_SS_ID;
      try {
          let tSheet = SpreadsheetApp.openById(targetSS_ID).getSheetByName("Envanter"), tData = tSheet.getDataRange().getValues();
          for(let k=1; k<tData.length; k++){ if(String(tData[k][0]).trim() === String(code).trim()) { tSheet.deleteRow(k+1); break; } }
      } catch(e) {}
  }
  let details = getInvDetails(data, code);
  logSheet.appendRow([code, devName, details.marka, details.model, "Sistem", Session.getActiveUser().getEmail(), new Date(), new Date(), "", "Transfer İptal", "Dış laboratuvar transferi geri çekildi."]);
  clearLabCache(); return "Transfer başarıyla iptal edildi ve cihaz Kalite envanterine geri döndü.";
}

function getLabData() {
  try {
    const cache = CacheService.getScriptCache();
    const cachedData = cache.get("kalite_lab_data");
    if (cachedData) return JSON.parse(cachedData);

    const ss = getSpreadsheet(), invSheet = ss.getSheetByName("Envanter"), logSheet = ss.getSheetByName("Kayitlar");
    const invData = invSheet.getDataRange().getValues(), headers = invData[0] || [];
    
    let activeLogs = {};
    let hL = headers.map(h => String(h).toLowerCase().trim());
    let idxKodu = 0, idxIsmi = 1;
    let idxMarka = hL.findIndex(h => h.includes('marka')); if(idxMarka===-1) idxMarka=2;
    let idxModel = hL.findIndex(h => h.includes('model')); if(idxModel===-1) idxModel=3;
    let idxCal = hL.findIndex(h => h.includes('kalibrasyon') && h.includes('son')); if(idxCal===-1) idxCal=4;
    let idxDurum = hL.findIndex(h => h === 'durum'); if(idxDurum===-1) idxDurum=6;
    let idxArizali = hL.findIndex(h => h.includes('arızalı') || h.includes('arizali')); if(idxArizali===-1) idxArizali=7;
    let idxKalibda = hL.findIndex(h => h.includes('kalibrasyonda')); if(idxKalibda===-1) idxKalibda=8;
    let idxKlasor = hL.findIndex(h => h.includes('klasör') || h.includes('kategori')); if(idxKlasor===-1) idxKlasor=9;
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

    let results = [];
    for (let i = 1; i < invData.length; i++) {
        let row = invData[i];
        if(!row || row.length === 0 || String(row[0]).trim() === "") continue; 
        
        let drm = String(row[idxDurum]).trim();
        if (drm.includes("Dış Lab. Devir Bekliyor") || drm.includes("Transfer Onayı Bekliyor")) continue;

        let obj = {};
        for (let j = 0; j < headers.length; j++) { 
          let hName = String(headers[j]).trim();
          obj[hName] = row[j] instanceof Date ? Utilities.formatDate(row[j], "GMT+3", "HH:mm dd.MM.yyyy") : (row[j] != null ? String(row[j]).trim() : ""); 
        }
        
        let invCode = String(row[idxKodu]).trim();
        obj['ActiveDetails'] = activeLogs[invCode] || null; 
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
    cache.put("kalite_lab_data", JSON.stringify(results), 900);
    return results;
  } catch(e) { throw new Error(e.message); }
}

function returnDevice(assetCode, returnNote) {
  try {
    const ss = getSpreadsheet(), logSheet = ss.getSheetByName("Kayitlar"), invSheet = ss.getSheetByName("Envanter");
    const logData = logSheet.getDataRange().getValues(), invData = invSheet.getDataRange().getValues(), durumCol = getDurumColIndex(invSheet);
    let logRowIndex = -1;

    for (let i = logData.length - 1; i >= 1; i--) { 
        if (String(logData[i][0]).trim() === String(assetCode).trim() && String(logData[i][9]).trim() === "Kullanımda") { logRowIndex = i + 1; break; } 
    }

    if (logRowIndex !== -1) {
      for (let j = 1; j < invData.length; j++) {
        if (String(invData[j][0]).trim() === String(assetCode).trim()) { invSheet.getRange(j + 1, durumCol).setValue("İade Onayı Bekliyor"); break; }
      }
      logSheet.getRange(logRowIndex, 10).setValue("İade Onayına Gönderildi");
      if(returnNote) { let extNote = logSheet.getRange(logRowIndex, 11).getValue(); logSheet.getRange(logRowIndex, 11).setValue(extNote ? extNote + " | İade: " + returnNote : "İade: " + returnNote); }
      clearLabCache(); return `İade talebiniz yöneticiye iletildi.`;
    }
    throw new Error("Aktif kullanım kaydı bulunamadı!");
  } catch(e) { throw new Error(e.message); }
}

function reportFault(code, note) {
  const userEmail = Session.getActiveUser().getEmail(), ss = getSpreadsheet(), invSheet = ss.getSheetByName("Envanter");
  const headers = invSheet.getDataRange().getValues()[0], invData = invSheet.getDataRange().getValues();
  let arizaliIdx = headers.findIndex(h => String(h).trim().includes("Arızalı")), notIdx = headers.findIndex(h => String(h).trim().includes("Arıza Notu"));
  if(arizaliIdx === -1) arizaliIdx = 6; if(notIdx === -1) notIdx = 10;
  
  for (let i = 1; i < invData.length; i++) {
    if (String(invData[i][0]).trim() === String(code).trim()) { invSheet.getRange(i + 1, arizaliIdx + 1).setValue(true); invSheet.getRange(i + 1, notIdx + 1).setValue(note); break; }
  }
  clearLabCache(); return "Bildirim iletildi.";
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

function updateDeviceStatus(code, status, note) {
  if (!checkIfAdmin(Session.getActiveUser().getEmail())) throw new Error("Yetkiniz yok!");
  const invSheet = getSpreadsheet().getSheetByName("Envanter"), headers = invSheet.getDataRange().getValues()[0], data = invSheet.getDataRange().getValues();
  let arizaliIdx = headers.findIndex(h => String(h).trim().includes("Arızalı")), kalibIdx = headers.findIndex(h => String(h).trim().includes("Kalibrasyonda")), durumIdx = headers.findIndex(h => String(h).trim().includes("Durum")), notIdx = headers.findIndex(h => String(h).trim().includes("Arıza Notu"));
  if(arizaliIdx === -1) arizaliIdx = 6; if(kalibIdx === -1) kalibIdx = 7; if(durumIdx === -1) durumIdx = 5; if(notIdx === -1) notIdx = 10;
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(code).trim()) {
      if (arizaliIdx !== -1) invSheet.getRange(i + 1, arizaliIdx + 1).setValue(status === "Arızalı");
      if (kalibIdx !== -1) invSheet.getRange(i + 1, kalibIdx + 1).setValue(status === "Kalibrasyonda");
      if (durumIdx !== -1 && status === "Müsait") invSheet.getRange(i + 1, durumIdx + 1).setValue("Müsait");
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

  let durumIdx = headers.findIndex(h => String(h).trim().includes("Durum")); if (durumIdx === -1) durumIdx = 6; 
  let idxName = 1, idxBrand = headers.findIndex(h => String(h).toLowerCase().trim().includes("marka")), idxModel = headers.findIndex(h => String(h).toLowerCase().trim().includes("model")), idxCal = headers.findIndex(h => String(h).toLowerCase().trim().includes("son kalibrasyon"));
  if(idxBrand===-1) idxBrand = 2; if(idxModel===-1) idxModel = 3; if(idxCal===-1) idxCal = 5;

  let pending = [];
  for (let i = 1; i < data.length; i++) {
    let d = String(data[i][durumIdx]).trim();
    if (d.includes("Onayı Bekliyor") || d.includes("Transfer Onayı Bekliyor") || d.includes("Biriminden Devir")) {
        let calVal = data[i][idxCal];
        if (Object.prototype.toString.call(calVal) === '[object Date]') {
            calVal = Utilities.formatDate(calVal, "GMT+3", "yyyy-MM-dd");
        } else if (calVal) { 
            let p = String(calVal).split('.'); 
            if(p.length === 3) calVal = `${p[2]}-${p[1]}-${p[0]}`; 
            else calVal = String(calVal);
        }

        let code = String(data[i][0]).trim();
        let lastUser = "Sistem / Dış Laboratuvar", lastAction = "Bekliyor", lastNote = "-";
        
        for (let j = logData.length - 1; j >= 1; j--) {
            if (String(logData[j][0]).trim() === code) {
                lastUser = String(logData[j][4] || logData[j][5] || "Sistem");
                lastAction = String(logData[j][9] || "İşlem");
                lastNote = String(logData[j][10] || "-");
                break;
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
  let userEmail = Session.getActiveUser().getEmail(), userName = userEmail.split('@')[0].replace(/\./g, ' ').toUpperCase() + " (Kalite)";

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
  } else {
     invSheet.getRange(invRowIdx, durumCol).setValue("Müsait"); 
     if (currentDurum.includes("Biriminden Devir") || currentDurum.includes("Transfer Onayı Bekliyor")) {
          let targetSS_ID = null;
          let logData = logSheet.getDataRange().getValues();
          for (let j = logData.length - 1; j >= 1; j--) {
              if (String(logData[j][0]).trim() === String(code).trim()) {
                  let who = String(logData[j][4]).toUpperCase(); let noteStr = String(logData[j][10]).toUpperCase();
                  if (who.includes("GÜÇ") || noteStr.includes("GÜÇ") || who.includes("GUC")) { targetSS_ID = GUC_SS_ID; break; }
                  if (who.includes("MEKANİK") || noteStr.includes("MEKANİK") || who.includes("MEKANIK")) { targetSS_ID = MEKANIK_SS_ID; break; }
              }
          }
          if(targetSS_ID) {
              try {
                  let tSheet = SpreadsheetApp.openById(targetSS_ID).getSheetByName("Envanter"), tData = tSheet.getDataRange().getValues();
                  for(let k = 1; k < tData.length; k++){ if(String(tData[k][0]).trim() === String(code).trim()){ tSheet.deleteRow(k + 1); break; } }
                  let tLog = SpreadsheetApp.openById(targetSS_ID).getSheetByName("Kayitlar");
                  if(tLog) tLog.appendRow([code, devName, details.marka, details.model, "Kalite Birimi", userEmail, new Date(), new Date(), "", "Transfer Onaylandı", "Karşı departman cihazı kabul etti."]);
              } catch(e) {}
          }
          logSheet.appendRow([code, devName, details.marka, details.model, userName, userEmail, new Date(), new Date(), "", "Transfer Onaylandı", "Cihaz Müsait olarak havuza alındı."]);
     } else {
          logSheet.appendRow([code, devName, details.marka, details.model, userName, userEmail, new Date(), new Date(), "", "Onaylandı", "Cihaz onaylanarak envantere dahil edildi."]);
     }
     clearLabCache(); return "Cihaz onaylandı ve havuza eklendi.";
  }
}

function rejectQuarantine(code, rejectNote) {
  let userEmail = Session.getActiveUser().getEmail();
  let userName = userEmail.split('@')[0].replace(/\./g, ' ').toUpperCase() + " (Kalite Birimi)";

  if (!checkIfAdmin(userEmail)) throw new Error("Yetkiniz yok!");
  const ss = getSpreadsheet(), invSheet = ss.getSheetByName("Envanter"), durumCol = getDurumColIndex(invSheet), invData = invSheet.getDataRange().getValues();

  for (let i = 1; i < invData.length; i++) {
    if (String(invData[i][0]).trim() === String(code).trim()) {
      let curDurum = String(invData[i][durumCol-1]).trim(), devName = String(invData[i][1]).trim();
      let logSheet = ss.getSheetByName("Kayitlar"), details = getInvDetails(invData, code);

      if (curDurum.includes("İade Onayı Bekliyor")) {
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
          let logData = logSheet.getDataRange().getValues();
          let originLab = null; let targetSS_ID = null;
          
          for (let j = logData.length - 1; j >= 1; j--) {
              if (String(logData[j][0]).trim() === String(code).trim()) {
                  let who = String(logData[j][4]).toUpperCase(); let noteStr = String(logData[j][10]).toUpperCase();
                  if (who.includes("GÜÇ") || noteStr.includes("GÜÇ") || who.includes("GUC")) { originLab = "Güç Elektroniği"; targetSS_ID = GUC_SS_ID; break; }
                  if (who.includes("MEKANİK") || noteStr.includes("MEKANİK") || who.includes("MEKANIK") || noteStr.includes("MEKANIK")) { originLab = "Mekanik Lab."; targetSS_ID = MEKANIK_SS_ID; break; }
                  break; 
              }
          }

          if (targetSS_ID) {
              try {
                  let eSS = SpreadsheetApp.openById(targetSS_ID), eSheet = eSS.getSheetByName("Envanter"), eData = eSheet.getDataRange().getValues(), eDurumCol = getDurumColIndex(eSheet);
                  for(let k=1; k<eData.length; k++) {
                      if(String(eData[k][0]).trim() === String(code).trim()) {
                          eSheet.getRange(k+1, eDurumCol).setValue("İade Onayı Bekliyor");
                          let eLog = eSS.getSheetByName("Kayitlar");
                          if(eLog) eLog.appendRow([code, devName, details.marka, details.model, userName, userEmail, new Date(), new Date(), "", "Transfer Reddedildi", `Ret Sebebi: ${rejectNote || "Belirtilmedi"}`]);
                          
                          let rejectAdminSubject = "Transfer Reddedildi: " + code;
                          let rejectAdminHtml = `
                          <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; max-width: 600px; margin: 0 auto; border: 1px solid #e9ecef; border-radius: 10px; overflow: hidden; box-shadow: 0 4px 10px rgba(0,0,0,0.05); background-color: #ffffff;">
                              <div style="background-color: #212529; border-bottom: 4px solid #dc3545; padding: 20px; text-align: center;">
                                  <img src="https://anova.com.tr/wp-content/uploads/2022/06/Anova_Logo_Beyaz_00.webp" alt="Anova Logo" style="height: 45px; display: inline-block;">
                              </div>
                              <div style="padding: 30px;">
                                  <h2 style="color: #212529; margin-top: 0; font-size: 22px;">Transfer Reddedildi ❌</h2>
                                  <p style="color: #555; font-size: 15px; line-height: 1.6;">Göndermiş olduğunuz cihaz <b>${userName}</b> tarafından reddedildi ve onay havuzunuza iade statüsünde geri döndü.</p>
                                  <div style="background-color: #fff5f5; border-left: 4px solid #dc3545; padding: 15px; margin: 20px 0; border-radius: 4px;">
                                      <div style="margin-bottom: 8px;"><span style="color: #6c757d; font-size: 12px; text-transform: uppercase; font-weight: bold;">Demirbaş Kodu</span><br><b style="color: #212529; font-size: 15px;">${code}</b></div>
                                      <div style="margin-bottom: 8px;"><span style="color: #6c757d; font-size: 12px; text-transform: uppercase; font-weight: bold;">Cihaz İsmi</span><br><b style="color: #212529; font-size: 15px;">${devName}</b></div>
                                      <div style="margin-bottom: 8px;"><span style="color: #6c757d; font-size: 12px; text-transform: uppercase; font-weight: bold;">Reddeden Kişi</span><br><b style="color: #212529; font-size: 15px;">${userName}</b></div>
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

          invSheet.deleteRow(i + 1); 
          logSheet.appendRow([code, devName, details.marka, details.model, userName, userEmail, new Date(), new Date(), "", "Reddedildi", `Transfer Reddedildi. Sebep: ${rejectNote || "Belirtilmedi"}`]);
          clearLabCache(); 
          return targetSS_ID ? "İşlem reddedildi ve cihaz gönderen laboratuvara iade edildi." : "İşlem reddedildi ve cihaz kaydı silindi.";
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
    
    if (devName && userName && !userName.toLowerCase().includes("sistem") && !userName.toLowerCase().includes("kalite") && !userName.toLowerCase().includes("güç") && !userName.toLowerCase().includes("mekanik")) {
       if (action === "Kullanımda" || action.includes("Teslim") || action.includes("Devredildi") || action.includes("İade") || action.includes("Onay")) {
           deviceCount[finalLabel] = (deviceCount[finalLabel] || 0) + 1; userCount[userName] = (userCount[userName] || 0) + 1; totalBorrows++;
       }
    }
  }
  
  let devArr = Object.keys(deviceCount).map(k => { let count = deviceCount[k]; return { label: k, count: count, percent: totalBorrows > 0 ? ((count / totalBorrows) * 100).toFixed(1) : 0 }; }).sort((a, b) => b.count - a.count).slice(0, 10);
  let userArr = Object.keys(userCount).map(k => { return { label: k, count: userCount[k] }; }).sort((a, b) => b.count - a.count).slice(0, 10);
  return { deviceStats: devArr, userStats: userArr };
}
