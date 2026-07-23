/**
 * PTC Payroll — Google Apps Script backend
 * -----------------------------------------------------------------------
 * วิธีติดตั้ง (ทำครั้งเดียว):
 *  1. ไปที่ https://script.google.com/ -> "โปรเจกต์ใหม่" (New project)
 *  2. ลบโค้ดเริ่มต้นทั้งหมด แล้ววางไฟล์นี้ทั้งหมดแทน
 *  3. กด "บันทึก" (ตั้งชื่อโปรเจกต์ เช่น "PTC Payroll Backend")
 *  4. กด Deploy -> New deployment
 *       - Select type: Web app
 *       - Execute as: Me (บัญชี Google ของคุณ)
 *       - Who has access: Anyone
 *  5. กด Deploy แล้วอนุญาตสิทธิ์ (Authorize access) ตามที่ Google ถามทุกขั้นตอน
 *  6. คัดลอก Web app URL ที่ได้ (ลงท้ายด้วย /exec)
 *  7. นำ URL ไปวางในหน้าเว็บ Payroll -> ปุ่ม "⚙ ตั้งค่า"
 *
 * สคริปต์นี้จะสร้าง Google Sheet ชื่อ "PTC Payroll Data" ไว้ใน Google Drive
 * ของบัญชีที่ Deploy โดยอัตโนมัติในการเรียกใช้งานครั้งแรก และเก็บข้อมูล
 * พนักงาน + ข้อมูลเงินเดือนแต่ละงวดไว้ในนั้น ทุกครั้งที่กด "บันทึกลง Google
 * Drive" บนหน้าเว็บ ข้อมูลทั้งหมดจะถูกเขียนทับ (sync) ลงชีตนี้
 * -----------------------------------------------------------------------
 */

var SHEET_NAME = 'PTC Payroll Data';
var PROP_KEY = 'PTC_PAYROLL_SHEET_ID';

function getOrCreateSpreadsheet_(){
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty(PROP_KEY);
  var ss;
  if(id){
    try{ ss = SpreadsheetApp.openById(id); return ss; }catch(e){ /* fallthrough: recreate */ }
  }
  ss = SpreadsheetApp.create(SHEET_NAME);
  props.setProperty(PROP_KEY, ss.getId());
  return ss;
}

function getSheet_(ss, name, headers){
  var sh = ss.getSheetByName(name);
  if(!sh){
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
  }
  return sh;
}

function doGet(e){
  var action = e.parameter.action || 'load';
  if(action === 'load'){
    return loadPayload_();
  }
  return jsonOutput_({status:'error', message:'unknown action'});
}

function doPost(e){
  var body;
  try{
    body = JSON.parse(e.postData.contents);
  }catch(err){
    return jsonOutput_({status:'error', message:'invalid JSON body'});
  }
  if(body.action === 'save'){
    return savePayload_(body.payload);
  }
  return jsonOutput_({status:'error', message:'unknown action'});
}

function savePayload_(payload){
  try{
    var ss = getOrCreateSpreadsheet_();

    // --- Meta sheet (company name, current period) ---
    var meta = getSheet_(ss, 'Meta', ['key','value']);
    meta.clearContents();
    meta.appendRow(['key','value']);
    meta.appendRow(['company', payload.company || '']);
    meta.appendRow(['currentPeriod', payload.currentPeriod || '']);

    // --- Employees sheet ---
    var empSheet = getSheet_(ss, 'Employees', ['id','name','position','department','rate']);
    empSheet.clearContents();
    empSheet.appendRow(['id','name','position','department','rate']);
    (payload.employees||[]).forEach(function(emp){
      empSheet.appendRow([emp.id, emp.name, emp.position, emp.department, emp.rate]);
    });

    // --- PayrollEntries sheet (flattened: one row per employee per pay period) ---
    var entrySheet = getSheet_(ss, 'PayrollEntries', [
      'payPeriod','employeeId','positionAllowance','overtime','commission','bonus',
      'otherIncome','ssoRate','incomeTax','pf','loanDeduction','wht'
    ]);
    entrySheet.clearContents();
    entrySheet.appendRow([
      'payPeriod','employeeId','positionAllowance','overtime','commission','bonus',
      'otherIncome','ssoRate','incomeTax','pf','loanDeduction','wht'
    ]);
    var periods = payload.payPeriods || {};
    Object.keys(periods).forEach(function(period){
      var entries = periods[period].entries || {};
      Object.keys(entries).forEach(function(empId){
        var en = entries[empId];
        entrySheet.appendRow([
          period, empId, en.positionAllowance||0, en.overtime||0, en.commission||0, en.bonus||0,
          en.otherIncome||0, en.ssoRate||0, en.incomeTax||0, en.pf||0, en.loanDeduction||0, en.wht||0
        ]);
      });
    });

    return jsonOutput_({status:'ok'});
  }catch(err){
    return jsonOutput_({status:'error', message: String(err)});
  }
}

function loadPayload_(){
  try{
    var ss = getOrCreateSpreadsheet_();

    var meta = ss.getSheetByName('Meta');
    var company = '', currentPeriod = '';
    if(meta){
      var metaRows = meta.getDataRange().getValues();
      metaRows.forEach(function(row){
        if(row[0]==='company') company = row[1];
        if(row[0]==='currentPeriod') currentPeriod = row[1];
      });
    }

    var employees = [];
    var empSheet = ss.getSheetByName('Employees');
    if(empSheet){
      var empRows = empSheet.getDataRange().getValues();
      for(var i=1;i<empRows.length;i++){
        var row = empRows[i];
        if(!row[0]) continue;
        employees.push({id:row[0], name:row[1], position:row[2], department:row[3], rate:Number(row[4])||0});
      }
    }

    var payPeriods = {};
    var entrySheet = ss.getSheetByName('PayrollEntries');
    if(entrySheet){
      var entryRows = entrySheet.getDataRange().getValues();
      for(var j=1;j<entryRows.length;j++){
        var r = entryRows[j];
        if(!r[0]) continue;
        var period = r[0], empId = r[1];
        if(!payPeriods[period]) payPeriods[period] = {entries:{}};
        payPeriods[period].entries[empId] = {
          positionAllowance:Number(r[2])||0, overtime:Number(r[3])||0, commission:Number(r[4])||0,
          bonus:Number(r[5])||0, otherIncome:Number(r[6])||0, ssoRate:Number(r[7])||0,
          incomeTax:Number(r[8])||0, pf:Number(r[9])||0, loanDeduction:Number(r[10])||0, wht:Number(r[11])||0
        };
      }
    }

    var payload = {
      company: company || 'บริษัท พัทธา คอร์ปอเรชั่น จำกัด',
      currentPeriod: currentPeriod,
      employees: employees,
      payPeriods: payPeriods
    };
    return jsonOutput_({status:'ok', payload: payload});
  }catch(err){
    return jsonOutput_({status:'error', message: String(err)});
  }
}

function jsonOutput_(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
