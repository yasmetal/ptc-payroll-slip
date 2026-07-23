/* =========================================================================
   ระบบเงินเดือน PTC — Payroll & Slip System
   สูตรคำนวณทั้งหมดอิงจากไฟล์ต้นฉบับ "ตาราง Payroll พร้อม Slip_PTC.xlsx"
   (แผ่น Payroll + แผ่น Slip) มีการแก้ไข 2 จุดที่พบว่าเป็นข้อผิดพลาดในสูตรต้นฉบับ:
     1) เดิม "รวมเงินได้" (M) = SUM(G:K) ไม่ได้รวมคอลัมน์ L "เงินได้อื่นๆ"
        -> แก้เป็นรวม L ด้วย เพื่อให้ยอดถูกต้องเมื่อมีการกรอกเงินได้อื่นๆ
     2) เดิมในชีต Slip ช่อง "เงินอื่นๆ" ใช้สูตร =-INDEX(...) (ติดลบ)
        -> แก้เป็นค่าบวกตามปกติ
   ========================================================================= */

const STORAGE_KEY = 'ptc_payroll_state_v1';
const GAS_URL_KEY = 'ptc_payroll_gas_url';

const THAI_MONTHS = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน',
  'กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];

// อีเมลเริ่มต้นของพนักงานแต่ละคน (แก้ไขได้ภายหลังในตารางเงินเดือน หรือแท็บสลิป)
const DEFAULT_EMAILS = {
  'Emp.001': 'kesara385860@gmail.com',
  'Emp.002': 'nuriyah2655@gmail.com',
  'Emp.003': 'kaewngamkhakayyaratn@gmail.com',
  'Emp.004': 'jane.korada@gmail.com',
  'Emp.005': 'khattiyanee25122549@gmail.com',
  'Emp.006': 'sakunrat200646@gmail.com'
};

/* ---------------------------------------------------------------------
   Default / seed data — real sample data copied from the source workbook
   --------------------------------------------------------------------- */
function seedState(){
  return {
    company: 'บริษัท พัทธา คอร์ปอเรชั่น จำกัด',
    employees: [
      {id:'Emp.001', name:'เกศรา นาก้อนทอง (เฟิน)',      position:'หัวหน้าคลังและนำเข้า', department:'นำเข้าสินค้า', rate:15000, email:DEFAULT_EMAILS['Emp.001']},
      {id:'Emp.002', name:'นูรียะห์ บาราเฮง (ยะห์)',        position:'เจ้าหน้า QC',          department:'QC สินค้า',   rate:11000, email:DEFAULT_EMAILS['Emp.002']},
      {id:'Emp.003', name:'กัญญารัตน์ แก้วงามขำ (กุ้ย)',    position:'แพ็คสินค้า',           department:'แพ็คสินค้า',  rate:10500, email:DEFAULT_EMAILS['Emp.003']},
      {id:'Emp.004', name:'กรฎา พรหมสุวรรณ์ (เจน)',        position:'แพ็คสินค้า',           department:'แพ็คสินค้า',  rate:10500, email:DEFAULT_EMAILS['Emp.004']},
      {id:'Emp.005', name:'ขัตติยานี สาริศรี (นุ่น)',       position:'นำเข้าสินค้า',         department:'นำเข้าสินค้า', rate:10500, email:DEFAULT_EMAILS['Emp.005']},
      {id:'Emp.006', name:'สกุลรัตน์ ประดิษสุวรรณ์ (ต้า)',  position:'เจ้าหน้า QC',          department:'เจ้าหน้า QC', rate:10500, email:DEFAULT_EMAILS['Emp.006']}
    ],
    currentPeriod: '2026-06-30',
    payPeriods: {
      '2026-06-30': {
        entries: {
          'Emp.001': mkEntry({ssoRate:5}),
          'Emp.002': mkEntry({ssoRate:5}),
          'Emp.003': mkEntry({ssoRate:5}),
          'Emp.004': mkEntry({ssoRate:5}),
          'Emp.005': mkEntry({ssoRate:0, wht:315}),
          'Emp.006': mkEntry({ssoRate:0, wht:315})
        }
      }
    }
  };
}

function mkEntry(overrides){
  return Object.assign({
    positionAllowance:0, overtime:0, commission:0, bonus:0, otherIncome:0,
    ssoRate:5, incomeTax:0, pf:0, loanDeduction:0, wht:0
  }, overrides || {});
}

/* ---------------------------------------------------------------------
   State load / save (localStorage as local cache; Google Drive via GAS)
   --------------------------------------------------------------------- */
let state = loadLocalState();

function loadLocalState(){
  try{
    const raw = localStorage.getItem(STORAGE_KEY);
    if(raw){
      const loaded = JSON.parse(raw);
      backfillEmails(loaded);
      return loaded;
    }
  }catch(e){ console.warn('load state failed', e); }
  return seedState();
}

// เติมอีเมลเริ่มต้นให้พนักงานที่เคยบันทึกไว้ก่อนมีฟีเจอร์นี้ (ไม่ทับอีเมลที่แก้ไขไว้แล้ว)
function backfillEmails(s){
  (s.employees||[]).forEach(emp=>{
    if(!emp.email && DEFAULT_EMAILS[emp.id]) emp.email = DEFAULT_EMAILS[emp.id];
    if(emp.email==null) emp.email = '';
  });
}
function saveLocalState(){
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

/* ---------------------------------------------------------------------
   Calculation engine — mirrors the Excel formulas
   --------------------------------------------------------------------- */
function round0(n){ return Math.round((n + Number.EPSILON) * 100) / 100; }

function calcRow(emp, entry){
  const salary = Number(emp.rate) || 0; // "เงินเดือน" (G) == "อัตราเงินเดือน" (F) เหมือนไฟล์ต้นฉบับ
  const positionAllowance = Number(entry.positionAllowance)||0;
  const overtime          = Number(entry.overtime)||0;
  const commission        = Number(entry.commission)||0;
  const bonus             = Number(entry.bonus)||0;
  const otherIncome       = Number(entry.otherIncome)||0;

  // M = SUM(G:K)+L  (แก้บั๊กจากต้นฉบับให้รวม "เงินได้อื่นๆ" ด้วย)
  const totalIncome = round0(salary+positionAllowance+overtime+commission+bonus+otherIncome);

  // N = IF(G<=1650,1650*rate%, IF(G<=15000,G*rate%,15000*rate%))  — ฐานประกันสังคมตามกฎ สปส. (1,650–15,000)
  const ssoRate = Number(entry.ssoRate)||0;
  const ssoBase = Math.min(Math.max(salary,1650),15000);
  const sso = round0(ssoBase * ssoRate/100);

  const incomeTax   = Number(entry.incomeTax)||0;      // O — กรอกเอง (ต้นฉบับไม่มีสูตร)
  const pf          = Number(entry.pf)||0;             // สำรองเลี้ยงชีพ (มีเฉพาะในชีต Slip ของต้นฉบับ)
  const loanDeduction = Number(entry.loanDeduction)||0; // P — กรอกเอง
  const wht         = Number(entry.wht)||0;            // Q — หัก ณ ที่จ่าย 3% (กรอกเอง/กดปุ่มคำนวณ 3%)

  const totalDeduct = round0(sso+incomeTax+pf+loanDeduction+wht); // R = SUM(N:Q)+PF
  const netPay = round0(totalIncome-totalDeduct);                  // S = M-R

  return {salary,positionAllowance,overtime,commission,bonus,otherIncome,totalIncome,
    ssoRate,sso,incomeTax,pf,loanDeduction,wht,totalDeduct,netPay};
}

function currentEntries(){
  if(!state.payPeriods[state.currentPeriod]){
    state.payPeriods[state.currentPeriod] = {entries:{}};
  }
  return state.payPeriods[state.currentPeriod].entries;
}
function getEntry(empId){
  const entries = currentEntries();
  if(!entries[empId]) entries[empId] = mkEntry();
  return entries[empId];
}

function fmt(n){
  return (Math.round(n*100)/100).toLocaleString('th-TH', {minimumFractionDigits:0, maximumFractionDigits:2});
}

/* ---------------------------------------------------------------------
   Rendering — Payroll table
   --------------------------------------------------------------------- */
function renderPayrollTable(){
  const tbody = document.getElementById('payrollBody');
  tbody.innerHTML = '';

  const totals = {income:0,pos:0,ot:0,comm:0,bonus:0,other:0,sso:0,tax:0,pf:0,loan:0,wht:0,deduct:0,net:0};

  state.employees.forEach((emp, idx)=>{
    const entry = getEntry(emp.id);
    const r = calcRow(emp, entry);

    totals.income += r.totalIncome; totals.pos += r.positionAllowance; totals.ot += r.overtime;
    totals.comm += r.commission; totals.bonus += r.bonus; totals.other += r.otherIncome;
    totals.sso += r.sso; totals.tax += r.incomeTax; totals.pf += r.pf; totals.loan += r.loanDeduction;
    totals.wht += r.wht; totals.deduct += r.totalDeduct; totals.net += r.netPay;

    const tr = document.createElement('tr');
    tr.dataset.empId = emp.id;
    tr.innerHTML = `
      <td>${idx+1}</td>
      <td><input class="text" data-field="id" value="${escapeHtml(emp.id)}" style="width:80px"></td>
      <td class="col-name"><input class="text" data-field="name" value="${escapeHtml(emp.name)}"></td>
      <td class="col-pos"><input class="text" data-field="position" value="${escapeHtml(emp.position)}"></td>
      <td class="col-dept"><input class="text" data-field="department" value="${escapeHtml(emp.department)}"></td>
      <td class="col-email"><input class="text" type="email" data-field="email" value="${escapeHtml(emp.email||'')}" placeholder="name@example.com" style="width:170px"></td>
      <td><input type="number" data-field="rate" value="${emp.rate}"></td>
      <td class="readonly">${fmt(r.salary)}</td>
      <td><input type="number" data-field="positionAllowance" value="${entry.positionAllowance}"></td>
      <td><input type="number" data-field="overtime" value="${entry.overtime}"></td>
      <td><input type="number" data-field="commission" value="${entry.commission}"></td>
      <td><input type="number" data-field="bonus" value="${entry.bonus}"></td>
      <td><input type="number" data-field="otherIncome" value="${entry.otherIncome}"></td>
      <td><input type="number" data-field="ssoRate" value="${entry.ssoRate}" title="อัตรา % (ค่าเริ่มต้น 5%)" style="width:55px"> %</td>
      <td><input type="number" data-field="incomeTax" value="${entry.incomeTax}"></td>
      <td><input type="number" data-field="pf" value="${entry.pf}"></td>
      <td><input type="number" data-field="loanDeduction" value="${entry.loanDeduction}"></td>
      <td>
        <input type="number" data-field="wht" value="${entry.wht}" style="width:65px">
        <button type="button" class="btn-mini" data-action="wht3" title="คำนวณ 3% ของรวมเงินได้อัตโนมัติ">3%</button>
      </td>
      <td class="readonly">${fmt(r.totalDeduct)}</td>
      <td class="col-net">${fmt(r.netPay)}</td>
      <td><button class="btn-remove" data-action="remove-emp">ลบ</button></td>
    `;
    tbody.appendChild(tr);
  });

  document.getElementById('tot_income').textContent = fmt(totals.income);
  document.getElementById('tot_pos').textContent = fmt(totals.pos);
  document.getElementById('tot_ot').textContent = fmt(totals.ot);
  document.getElementById('tot_comm').textContent = fmt(totals.comm);
  document.getElementById('tot_bonus').textContent = fmt(totals.bonus);
  document.getElementById('tot_other').textContent = fmt(totals.other);
  document.getElementById('tot_sso').textContent = fmt(totals.sso);
  document.getElementById('tot_tax').textContent = fmt(totals.tax);
  document.getElementById('tot_pf').textContent = fmt(totals.pf);
  document.getElementById('tot_loan').textContent = fmt(totals.loan);
  document.getElementById('tot_wht').textContent = fmt(totals.wht);
  document.getElementById('tot_deduct').textContent = fmt(totals.deduct);
  document.getElementById('tot_net').textContent = fmt(totals.net);

  renderSlipEmployeeOptions();
}

function escapeHtml(s){
  return String(s??'').replace(/[&<>"']/g, c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

/* ---------------------------------------------------------------------
   Payroll table event handling (delegated)
   --------------------------------------------------------------------- */
document.getElementById('payrollBody').addEventListener('input', (e)=>{
  const field = e.target.dataset.field;
  if(!field) return;
  const tr = e.target.closest('tr');
  const empId = tr.dataset.empId;
  const emp = state.employees.find(x=>x.id===empId);
  const entry = getEntry(empId);

  const employeeFields = ['id','name','position','department','rate','email'];
  if(employeeFields.includes(field)){
    if(field==='id'){
      const newId = e.target.value.trim();
      if(newId && newId!==empId){
        // rename id across employee + all period entries
        state.employees.forEach(x=>{});
        emp.id = newId;
        Object.values(state.payPeriods).forEach(p=>{
          if(p.entries[empId]){ p.entries[newId]=p.entries[empId]; delete p.entries[empId]; }
        });
        tr.dataset.empId = newId;
      }
    } else if(field==='rate'){
      emp.rate = Number(e.target.value)||0;
    } else {
      emp[field] = e.target.value;
    }
  } else {
    entry[field] = Number(e.target.value)||0;
  }
  saveLocalState();
  renderPayrollTable();
  refreshFocus(tr.dataset.empId, field);
});

// keep input focus after full re-render (simple approach: re-focus same field of same emp)
function refreshFocus(empId, field){
  const tr = document.querySelector(`#payrollBody tr[data-emp-id="${CSS.escape(empId)}"]`);
  if(!tr) return;
  const input = tr.querySelector(`[data-field="${field}"]`);
  if(input){ input.focus(); const v=input.value; input.value=''; input.value=v; }
}

document.getElementById('payrollBody').addEventListener('click', (e)=>{
  const action = e.target.dataset.action;
  if(!action) return;
  const tr = e.target.closest('tr');
  const empId = tr.dataset.empId;

  if(action==='remove-emp'){
    if(!confirm('ลบพนักงานคนนี้ออกจากตาราง? (ข้อมูลงวดนี้ของพนักงานจะถูกลบด้วย)')) return;
    state.employees = state.employees.filter(x=>x.id!==empId);
    Object.values(state.payPeriods).forEach(p=>{ delete p.entries[empId]; });
    saveLocalState();
    renderPayrollTable();
  }
  if(action==='wht3'){
    const emp = state.employees.find(x=>x.id===empId);
    const entry = getEntry(empId);
    const r = calcRow(emp, entry);
    entry.wht = round0(r.totalIncome*0.03);
    saveLocalState();
    renderPayrollTable();
  }
});

document.getElementById('btnAddEmployee').addEventListener('click', ()=>{
  let n = state.employees.length+1;
  let newId = 'Emp.'+String(n).padStart(3,'0');
  while(state.employees.some(e=>e.id===newId)){ n++; newId='Emp.'+String(n).padStart(3,'0'); }
  state.employees.push({id:newId, name:'', position:'', department:'', rate:0});
  saveLocalState();
  renderPayrollTable();
});

/* ---------------------------------------------------------------------
   Pay date control
   --------------------------------------------------------------------- */
const payDateInput = document.getElementById('payDate');
payDateInput.value = state.currentPeriod;
payDateInput.addEventListener('change', ()=>{
  state.currentPeriod = payDateInput.value;
  if(!state.payPeriods[state.currentPeriod]) state.payPeriods[state.currentPeriod] = {entries:{}};
  saveLocalState();
  renderPayrollTable();
  renderSlip();
});

const companyNameInput = document.getElementById('companyNameInput');
companyNameInput.value = state.company;
companyNameInput.addEventListener('input', ()=>{
  state.company = companyNameInput.value;
  document.getElementById('companyNameHeader').textContent = state.company;
  document.querySelector('.slip-company').textContent = state.company;
  saveLocalState();
});

/* ---------------------------------------------------------------------
   Tabs
   --------------------------------------------------------------------- */
document.querySelectorAll('.tab-btn').forEach(btn=>{
  btn.addEventListener('click', ()=>{
    document.querySelectorAll('.tab-btn').forEach(b=>b.classList.remove('active'));
    document.querySelectorAll('.tab-panel').forEach(p=>p.classList.remove('active'));
    btn.classList.add('active');
    document.getElementById('tab-'+btn.dataset.tab).classList.add('active');
    if(btn.dataset.tab==='slip'){ renderSlipEmployeeOptions(); renderSlip(); }
  });
});

/* ---------------------------------------------------------------------
   Slip rendering
   --------------------------------------------------------------------- */
function renderSlipEmployeeOptions(){
  const sel = document.getElementById('slipEmployeeSelect');
  const prev = sel.value;
  sel.innerHTML = state.employees.map(e=>`<option value="${escapeHtml(e.id)}">${escapeHtml(e.id)} - ${escapeHtml(e.name)}</option>`).join('');
  if(prev && state.employees.some(e=>e.id===prev)) sel.value = prev;
}
document.getElementById('slipEmployeeSelect').addEventListener('change', renderSlip);

function thaiDateStr(dateStr){
  if(!dateStr) return '';
  const [y,m,d] = dateStr.split('-').map(Number);
  return `${d} ${THAI_MONTHS[m-1]} ${y+543}`;
}
function thaiMonthStr(dateStr){
  if(!dateStr) return '';
  const [y,m] = dateStr.split('-').map(Number);
  return THAI_MONTHS[m-1];
}

function renderSlip(){
  const sel = document.getElementById('slipEmployeeSelect');
  const empId = sel.value;
  const emp = state.employees.find(e=>e.id===empId);
  if(!emp) return;
  const entry = getEntry(empId);
  const r = calcRow(emp, entry);

  document.getElementById('s_empid').textContent = emp.id;
  document.getElementById('s_pos').textContent = emp.position;
  document.getElementById('s_name').textContent = emp.name;
  document.getElementById('s_dept').textContent = emp.department;
  document.getElementById('s_paydate').textContent = thaiDateStr(state.currentPeriod);
  document.getElementById('s_month').textContent = thaiMonthStr(state.currentPeriod);

  document.getElementById('s_salary').textContent = fmt(r.salary)+' บาท';
  document.getElementById('s_posallow').textContent = fmt(r.positionAllowance)+' บาท';
  document.getElementById('s_ot').textContent = fmt(r.overtime)+' บาท';
  document.getElementById('s_comm').textContent = fmt(r.commission)+' บาท';
  document.getElementById('s_bonus').textContent = fmt(r.bonus)+' บาท';
  document.getElementById('s_otherinc').textContent = fmt(r.otherIncome)+' บาท';
  document.getElementById('s_totalincome').textContent = fmt(r.totalIncome)+' บาท';

  document.getElementById('s_sso').textContent = fmt(r.sso)+' บาท';
  document.getElementById('s_tax').textContent = fmt(r.incomeTax)+' บาท';
  document.getElementById('s_pf').textContent = fmt(r.pf)+' บาท';
  document.getElementById('s_loan').textContent = fmt(r.loanDeduction)+' บาท';
  document.getElementById('s_wht').textContent = fmt(r.wht)+' บาท';
  document.getElementById('s_totaldeduct').textContent = fmt(r.totalDeduct)+' บาท';

  document.getElementById('s_netpay').textContent = fmt(r.netPay)+' บาท';

  document.querySelector('.slip-company').textContent = state.company;

  document.getElementById('slipEmployeeEmail').value = emp.email || '';
}

// อัปเดตอีเมลของพนักงานที่เลือกอยู่ในแท็บสลิป (แก้ไขได้ภายหลัง, sync กับตารางเงินเดือน)
document.getElementById('slipEmployeeEmail').addEventListener('input', ()=>{
  const empId = document.getElementById('slipEmployeeSelect').value;
  const emp = state.employees.find(e=>e.id===empId);
  if(!emp) return;
  emp.email = document.getElementById('slipEmployeeEmail').value.trim();
  saveLocalState();
});

/* ---------------------------------------------------------------------
   PDF build (per employee) — html2canvas + jsPDF, shared by download & email
   --------------------------------------------------------------------- */
async function buildSlipPdf(empId){
  const emp = state.employees.find(e=>e.id===empId);
  if(!emp) return null;
  const prevSelected = document.getElementById('slipEmployeeSelect').value;
  document.getElementById('slipEmployeeSelect').value = empId;
  renderSlip();
  await new Promise(r=>setTimeout(r,50)); // let DOM paint

  const node = document.getElementById('slipDoc');
  const canvas = await html2canvas(node, {scale:2, backgroundColor:'#ffffff'});
  const imgData = canvas.toDataURL('image/png');
  const { jsPDF } = window.jspdf;
  const pdf = new jsPDF({orientation:'portrait', unit:'pt', format:'a4'});
  const pageWidth = pdf.internal.pageSize.getWidth();
  const imgW = pageWidth-40;
  const imgH = canvas.height * (imgW/canvas.width);
  pdf.addImage(imgData,'PNG',20,20,imgW,imgH);
  const fname = `Slip_${emp.id}_${state.currentPeriod}.pdf`.replace(/[^\w.\-]+/g,'_');

  document.getElementById('slipEmployeeSelect').value = prevSelected;
  renderSlip();

  return {pdf, filename:fname, emp};
}

async function exportSlipToPdf(empId){
  const built = await buildSlipPdf(empId);
  if(!built) return;
  built.pdf.save(built.filename);
}

document.getElementById('btnDownloadSlip').addEventListener('click', ()=>{
  const empId = document.getElementById('slipEmployeeSelect').value;
  exportSlipToPdf(empId);
});

document.getElementById('btnExportAllPdf').addEventListener('click', async ()=>{
  document.querySelectorAll('.tab-btn').forEach(b=>b.classList.remove('active'));
  document.querySelectorAll('.tab-panel').forEach(p=>p.classList.remove('active'));
  document.querySelector('[data-tab="slip"]').classList.add('active');
  document.getElementById('tab-slip').classList.add('active');
  for(const emp of state.employees){
    await exportSlipToPdf(emp.id);
    await new Promise(r=>setTimeout(r,300));
  }
});

/* ---------------------------------------------------------------------
   Email slip (per employee) — sends the PDF as an attachment via the
   Google Apps Script backend (Gmail account that deployed Code.gs)
   --------------------------------------------------------------------- */
function pdfToBase64(pdf){
  // pdf.output('datauristring') -> "data:application/pdf;filename=...;base64,XXXXX"
  const datauri = pdf.output('datauristring');
  const idx = datauri.indexOf('base64,');
  return idx>=0 ? datauri.slice(idx+7) : '';
}

async function emailSlipToEmployee(empId){
  const gasUrl = getGasUrl();
  if(!gasUrl){ alert('กรุณาตั้งค่า Apps Script Web App URL ก่อน (ปุ่ม ⚙ ตั้งค่า)'); return {ok:false, skipped:true}; }

  const emp = state.employees.find(e=>e.id===empId);
  if(!emp) return {ok:false};
  if(!emp.email){ return {ok:false, noEmail:true, emp}; }

  const built = await buildSlipPdf(empId);
  if(!built) return {ok:false};

  const pdfBase64 = pdfToBase64(built.pdf);
  const subject = `สลิปเงินเดือน ${thaiDateStr(state.currentPeriod)} - ${emp.name}`;
  const body = `เรียนคุณ ${emp.name}\n\nแนบไฟล์สลิปเงินเดือนประจำงวดวันที่ ${thaiDateStr(state.currentPeriod)} มาด้วยแล้ว\n\nขอบคุณครับ/ค่ะ\n${state.company}`;

  try{
    const res = await fetch(gasUrl, {
      method:'POST',
      headers:{'Content-Type':'text/plain;charset=utf-8'},
      body: JSON.stringify({
        action:'sendEmail',
        payload:{ to: emp.email, subject, body, filename: built.filename, pdfBase64 }
      })
    });
    const data = await res.json();
    return {ok: data.status==='ok', message: data.message, emp};
  }catch(err){
    console.error(err);
    return {ok:false, error:err, emp};
  }
}

document.getElementById('btnEmailSlip').addEventListener('click', async ()=>{
  const empId = document.getElementById('slipEmployeeSelect').value;
  const btn = document.getElementById('btnEmailSlip');
  const original = btn.textContent;
  btn.textContent = 'กำลังส่ง...'; btn.disabled = true;
  const result = await emailSlipToEmployee(empId);
  btn.textContent = original; btn.disabled = false;

  if(result.noEmail){ alert('พนักงานคนนี้ยังไม่มีอีเมล กรุณากรอกอีเมลก่อนส่ง'); return; }
  if(result.ok){ alert(`ส่งสลิปไปที่ ${result.emp.email} สำเร็จ`); }
  else { alert('ส่งอีเมลไม่สำเร็จ: '+(result.message||'ตรวจสอบ URL และการ Deploy Apps Script (ดู README.md)')); }
});

document.getElementById('btnEmailAllSlips').addEventListener('click', async ()=>{
  if(!getGasUrl()){ alert('กรุณาตั้งค่า Apps Script Web App URL ก่อน (ปุ่ม ⚙ ตั้งค่า)'); return; }
  if(!confirm(`จะส่งสลิปเงินเดือนงวด ${thaiDateStr(state.currentPeriod)} ให้พนักงานทุกคนที่มีอีเมลทาง E-mail ดำเนินการต่อ?`)) return;

  document.querySelectorAll('.tab-btn').forEach(b=>b.classList.remove('active'));
  document.querySelectorAll('.tab-panel').forEach(p=>p.classList.remove('active'));
  document.querySelector('[data-tab="slip"]').classList.add('active');
  document.getElementById('tab-slip').classList.add('active');

  const btn = document.getElementById('btnEmailAllSlips');
  const original = btn.textContent;
  btn.disabled = true;
  let sent = 0, failed = [], skipped = [];
  for(const emp of state.employees){
    btn.textContent = `กำลังส่ง... (${sent+failed.length+skipped.length+1}/${state.employees.length})`;
    const result = await emailSlipToEmployee(emp.id);
    if(result.noEmail) skipped.push(emp.name);
    else if(result.ok) sent++;
    else failed.push(emp.name);
    await new Promise(r=>setTimeout(r,300));
  }
  btn.textContent = original; btn.disabled = false;

  let msg = `ส่งสำเร็จ ${sent} คน`;
  if(skipped.length) msg += `\nข้ามเพราะไม่มีอีเมล: ${skipped.join(', ')}`;
  if(failed.length) msg += `\nส่งไม่สำเร็จ: ${failed.join(', ')}`;
  alert(msg);
});

/* ---------------------------------------------------------------------
   Google Drive sync via Google Apps Script Web App
   --------------------------------------------------------------------- */
function getGasUrl(){ return localStorage.getItem(GAS_URL_KEY) || ''; }
function setGasUrl(url){ localStorage.setItem(GAS_URL_KEY, url); }

function updateSyncStatus(){
  const el = document.getElementById('syncStatus');
  const url = getGasUrl();
  el.textContent = url ? '✅ เชื่อมต่อ Google Drive แล้ว' : 'ยังไม่ได้เชื่อมต่อ Google Drive';
}

document.getElementById('btnSettings').addEventListener('click', ()=>{
  document.getElementById('gasUrlInput').value = getGasUrl();
  document.getElementById('settingsModal').classList.add('open');
});
document.getElementById('btnCloseSettings').addEventListener('click', ()=>{
  document.getElementById('settingsModal').classList.remove('open');
});
document.getElementById('btnSaveSettings').addEventListener('click', ()=>{
  setGasUrl(document.getElementById('gasUrlInput').value.trim());
  updateSyncStatus();
  document.getElementById('settingsModal').classList.remove('open');
});

document.getElementById('btnSaveDrive').addEventListener('click', async ()=>{
  const url = getGasUrl();
  if(!url){ alert('กรุณาตั้งค่า Apps Script Web App URL ก่อน (ปุ่ม ⚙ ตั้งค่า)'); return; }
  const btn = document.getElementById('btnSaveDrive');
  const original = btn.textContent;
  btn.textContent = 'กำลังบันทึก...'; btn.disabled = true;
  try{
    const res = await fetch(url, {
      method:'POST',
      headers:{'Content-Type':'text/plain;charset=utf-8'}, // avoids CORS preflight to Apps Script
      body: JSON.stringify({action:'save', payload: state})
    });
    const data = await res.json();
    if(data.status==='ok'){ alert('บันทึกข้อมูลลง Google Drive สำเร็จ'); }
    else { alert('บันทึกไม่สำเร็จ: '+(data.message||'unknown error')); }
  }catch(err){
    console.error(err);
    alert('เชื่อมต่อ Google Drive ไม่สำเร็จ ตรวจสอบ URL และการ Deploy Apps Script (ดู README.md)');
  }finally{
    btn.textContent = original; btn.disabled = false;
  }
});

document.getElementById('btnLoadDrive').addEventListener('click', async ()=>{
  const url = getGasUrl();
  if(!url){ alert('กรุณาตั้งค่า Apps Script Web App URL ก่อน (ปุ่ม ⚙ ตั้งค่า)'); return; }
  if(!confirm('การโหลดจะแทนที่ข้อมูลปัจจุบันในเบราว์เซอร์นี้ ดำเนินการต่อ?')) return;
  const btn = document.getElementById('btnLoadDrive');
  const original = btn.textContent;
  btn.textContent = 'กำลังโหลด...'; btn.disabled = true;
  try{
    const res = await fetch(url+'?action=load');
    const data = await res.json();
    if(data.status==='ok' && data.payload){
      state = data.payload;
      if(!state.currentPeriod) state.currentPeriod = Object.keys(state.payPeriods||{})[0] || '2026-06-30';
      saveLocalState();
      payDateInput.value = state.currentPeriod;
      companyNameInput.value = state.company;
      document.getElementById('companyNameHeader').textContent = state.company;
      renderPayrollTable();
      renderSlip();
      alert('โหลดข้อมูลจาก Google Drive สำเร็จ');
    } else {
      alert('ไม่พบข้อมูล หรือโหลดไม่สำเร็จ: '+(data.message||''));
    }
  }catch(err){
    console.error(err);
    alert('เชื่อมต่อ Google Drive ไม่สำเร็จ ตรวจสอบ URL และการ Deploy Apps Script (ดู README.md)');
  }finally{
    btn.textContent = original; btn.disabled = false;
  }
});

/* ---------------------------------------------------------------------
   Init
   --------------------------------------------------------------------- */
document.getElementById('companyNameHeader').textContent = state.company;
updateSyncStatus();
renderPayrollTable();
renderSlip();
