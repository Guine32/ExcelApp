// app.js
// Colunas (em ordem): Ordens, Operação, Trab.real, Unidade, Cen Trab, Centro,
// Nº Pessoal, Data Inic, Inicio, Data Fim, Término, Data Lanca, Texto de confirmação

const cenOptions = ["SP04IPVS","SP04IVSO","SP04IMBI","SP04IBUT","SP04IPIN","SP04IFAL","SP04IFRA","SP04IFRE","SP04IPTA","SP04IHIG","SP04IREP","SP04ILUZ"];

const tableBody = document.getElementById('tableBody');
const addRowBtn = document.getElementById('addRowBtn');
const exportBtn = document.getElementById('exportBtn');

addRowBtn.addEventListener('click', addEditableRow);
exportBtn.addEventListener('click', exportToExcel);

document.getElementById('createCollabInputs').addEventListener('click', createCollabFields);
document.getElementById('applyCollabs').addEventListener('click', applyCollaborators);

// init with one row
addEditableRow();

function addEditableRow(prefill = {}) {
  const tr = document.createElement('tr');

  // helpers to create cells
  function tdInput(value = '', placeholder = '') {
    const td = document.createElement('td');
    td.className = 'border px-2 py-1';
    const input = document.createElement('input');
    input.type = 'text';
    input.value = value;
    input.placeholder = placeholder;
    input.className = 'w-full';
    td.appendChild(input);
    return { td, input };
  }

  // 1 Ordens (user)
  const { td: tdOrd, input: inpOrd } = tdInput(prefill.ordens || '', 'Ordens');
  tr.appendChild(tdOrd);

  // 2 Operação (0060 auto)
  const tdOper = document.createElement('td'); tdOper.className = 'border px-2 py-1';
  const inpOper = document.createElement('input'); inpOper.type = 'text'; inpOper.value = '0060'; inpOper.readOnly = true; inpOper.className='w-full bg-gray-50';
  tdOper.appendChild(inpOper); tr.appendChild(tdOper);

  // 3 Trab.real (computed)
  const tdTrab = document.createElement('td'); tdTrab.className = 'border px-2 py-1';
  const inpTrab = document.createElement('input'); inpTrab.type = 'text'; inpTrab.readOnly = true; inpTrab.className='w-full bg-gray-50';
  tdTrab.appendChild(inpTrab); tr.appendChild(tdTrab);

  // 4 Unidade HRS
  const tdUnid = document.createElement('td'); tdUnid.className = 'border px-2 py-1';
  const inpUnid = document.createElement('input'); inpUnid.type = 'text'; inpUnid.value = 'HRS'; inpUnid.readOnly = true; inpUnid.className='w-full bg-gray-50';
  tdUnid.appendChild(inpUnid); tr.appendChild(tdUnid);

  // 5 Cen Trab (select)
  const tdCen = document.createElement('td'); tdCen.className = 'border px-2 py-1';
  const selCen = document.createElement('select'); selCen.className='w-full';
  const emptyOpt = document.createElement('option'); emptyOpt.value=''; emptyOpt.textContent='--';
  selCen.appendChild(emptyOpt);
  cenOptions.forEach(c => {
    const o = document.createElement('option'); o.value = c; o.textContent = c; selCen.appendChild(o);
  });
  if(prefill.cen) selCen.value = prefill.cen;
  tdCen.appendChild(selCen); tr.appendChild(tdCen);

  // 6 Centro VQ01
  const tdCentro = document.createElement('td'); tdCentro.className = 'border px-2 py-1';
  const inpCentro = document.createElement('input'); inpCentro.type='text'; inpCentro.value='VQ01'; inpCentro.readOnly = true; inpCentro.className='w-full bg-gray-50';
  tdCentro.appendChild(inpCentro); tr.appendChild(tdCentro);

  // 7 Nº Pessoal (user will fill later but allow input now)
  const { td: tdNp, input: inpNp } = tdInput(prefill.np || '', 'Nº Pessoal');
  tr.appendChild(tdNp);

  // 8 Data Inic (date)
  const tdDataIn = document.createElement('td'); tdDataIn.className = 'border px-2 py-1';
  const inDataIn = document.createElement('input'); inDataIn.type = 'text'; inDataIn.placeholder='DD.MM.AAAA'; inDataIn.className='w-full date-input';
  if(prefill.dataIn) inDataIn.value = prefill.dataIn;
  tdDataIn.appendChild(inDataIn); tr.appendChild(tdDataIn);

  // 9 Inicio (time)
  const tdInicio = document.createElement('td'); tdInicio.className = 'border px-2 py-1';
  const inInicio = document.createElement('input'); inInicio.type='text'; inInicio.placeholder='HH:MM:SS'; inInicio.className='w-full time-input';
  if(prefill.inicio) inInicio.value = prefill.inicio;
  tdInicio.appendChild(inInicio); tr.appendChild(tdInicio);

  // 10 Data Fim (date)
  const tdDataFim = document.createElement('td'); tdDataFim.className = 'border px-2 py-1';
  const inDataFim = document.createElement('input'); inDataFim.type='text'; inDataFim.placeholder='DD.MM.AAAA'; inDataFim.className='w-full date-input';
  if(prefill.dataFim) inDataFim.value = prefill.dataFim;
  tdDataFim.appendChild(inDataFim); tr.appendChild(tdDataFim);

  // 11 Término (time)
  const tdTerm = document.createElement('td'); tdTerm.className = 'border px-2 py-1';
  const inTerm = document.createElement('input'); inTerm.type='text'; inTerm.placeholder='HH:MM:SS'; inTerm.className='w-full time-input';
  if(prefill.termino) inTerm.value = prefill.termino;
  tdTerm.appendChild(inTerm); tr.appendChild(tdTerm);

  // 12 Data Lanca (blank)
  const tdDL = document.createElement('td'); tdDL.className = 'border px-2 py-1';
  const inpDL = document.createElement('input'); inpDL.type='text'; inpDL.value = ''; inpDL.className='w-full';
  tdDL.appendChild(inpDL); tr.appendChild(tdDL);

  // 13 Texto de confirmação (user)
  const { td: tdTexto, input: inpTexto } = tdInput(prefill.texto || '', 'Texto de confirmação');
  tr.appendChild(tdTexto);

  // actions (delete)
  const tdActions = document.createElement('td'); tdActions.className = 'border px-2 py-1';
  const delBtn = document.createElement('button'); delBtn.className='text-red-600 px-2'; delBtn.textContent='Remover';
  delBtn.addEventListener('click', () => { tr.remove(); });
  tdActions.appendChild(delBtn); tr.appendChild(tdActions);

  // append
  tableBody.appendChild(tr);

  // Initialize flatpickr on the date and time fields
  flatpickr(inDataIn, {
    dateFormat: 'd.m.Y',
    allowInput: true,
    defaultDate: inDataIn.value || null
  });
  flatpickr(inDataFim, {
    dateFormat: 'd.m.Y',
    allowInput: true,
    defaultDate: inDataFim.value || null
  });

  flatpickr(inInicio, {
    enableTime: true,
    noCalendar: true,
    enableSeconds: true,
    time_24hr: true,
    dateFormat: 'H:i:S',
    defaultDate: inInicio.value || null
  });
  flatpickr(inTerm, {
    enableTime: true,
    noCalendar: true,
    enableSeconds: true,
    time_24hr: true,
    dateFormat: 'H:i:S',
    defaultDate: inTerm.value || null
  });

  const computeAndRender = () => {
    const dataIn = inDataIn.value.trim();
    const dataFim = inDataFim.value.trim();
    const inicio = inInicio.value.trim();
    const termino = inTerm.value.trim();

    // Update preview fields (inputs already show)
    // compute Trab.real only if both date & times available
    if(dataIn && dataFim && inicio && termino) {
      const diff = computeDuration(dataIn, inicio, dataFim, termino);
      inpTrab.value = diff;
    } else {
      inpTrab.value = '';
    }
  };

  // Add listeners to relevant fields
  [inDataIn, inDataFim, inInicio, inTerm].forEach(el => el.addEventListener('change', computeAndRender));
  // For other inputs that affect preview nothing special needed (native inputs show current value)

  // Return tr in case needed
  return tr;
}

// computeDuration: data strings formatted DD.MM.YYYY, time strings HH:MM:SS
function computeDuration(d1, t1, d2, t2) {
  // parse dd.mm.yyyy
  function parseDate(dstr, tstr) {
    const [dd, mm, yyyy] = dstr.split('.');
    const [hh, min, ss] = tstr.split(':');
    if(!dd || !mm || !yyyy || hh===undefined) return null;
    return new Date(
      parseInt(yyyy), parseInt(mm)-1, parseInt(dd),
      parseInt(hh), parseInt(min), parseInt(ss)
    );
  }
  const start = parseDate(d1, t1);
  let end = parseDate(d2, t2);
  if(!start || !end) return '';
  // If end < start, assume end is next day (or later) — keep as is though if user meant another date
  if(end < start) {
    // add one day
    end = new Date(end.getTime() + 24*3600*1000);
  }
  let ms = end - start;
  if(ms < 0) return '';
  const seconds = Math.floor(ms/1000);
  const hh = Math.floor(seconds/3600);
  const mm = Math.floor((seconds%3600)/60);
  const ss = seconds%60;
  const pad = n => String(n).padStart(2,'0');
  return `${pad(hh)}:${pad(mm)}:${pad(ss)}`;
}

// EXPORT to Excel using SheetJS
function exportToExcel() {
  // build data array of arrays with header
  const header = ["Ordens","Operação","Trab.real","Unidade","Cen Trab","Centro","Nº Pessoal","Data Inic","Inicio","Data Fim","Término","Data Lanca","Texto de confirmação"];
  const rows = [header];

  // iterate over table rows
  for(const tr of tableBody.querySelectorAll('tr')) {
    // skip empty-row markers (we will allow them - produce empty row in excel)
    if(tr.classList.contains('empty-row')) {
      rows.push(new Array(header.length).fill(''));
      continue;
    }
    const cells = tr.querySelectorAll('td');
    if(cells.length < 13) continue;
    const getVal = (cellIndex) => {
      const el = cells[cellIndex].querySelector('input,select');
      return el ? (el.value ?? '') : '';
    };
    const row = [
      getVal(0), // Ordens
      getVal(1), // Operação
      getVal(2), // Trab.real
      getVal(3), // Unidade
      getVal(4), // Cen Trab
      getVal(5), // Centro
      getVal(6), // Nº Pessoal
      getVal(7), // Data Inic
      getVal(8), // Inicio
      getVal(9), // Data Fim
      getVal(10),// Término
      getVal(11),// Data Lanca
      getVal(12) // Texto
    ];
    rows.push(row);
  }

  // create worksheet
  const ws = XLSX.utils.aoa_to_sheet(rows);

  // Optionally: set column widths
  const colWidths = [
    {wpx:110},{wpx:70},{wpx:90},{wpx:50},{wpx:110},{wpx:60},{wpx:90},{wpx:90},{wpx:80},{wpx:90},{wpx:80},{wpx:80},{wpx:220}
  ];
  ws['!cols'] = colWidths;

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Planilha");

  // write file
  XLSX.writeFile(wb, "planilha_gerada.xlsx");
}

// ------- Colaboradores handling -------

function createCollabFields() {
  const num = parseInt(document.getElementById('numCollab').value || '0', 10);
  const container = document.getElementById('collabInputs');
  container.innerHTML = '';
  if(!num || num <= 0) return;
  for(let i=0;i<num;i++){
    const div = document.createElement('div');
    div.className = 'flex gap-2 items-center';
    const label = document.createElement('label');
    label.textContent = `Nº Pessoal ${i+1}:`;
    const inp = document.createElement('input');
    inp.type = 'text';
    inp.className = 'border px-2 py-1';
    inp.placeholder = 'Nº Pessoal';
    div.appendChild(label);
    div.appendChild(inp);
    container.appendChild(div);
  }
}

function applyCollaborators() {
  // collect personnel numbers
  const container = document.getElementById('collabInputs');
  const inputs = Array.from(container.querySelectorAll('input'));
  if(inputs.length === 0) {
    alert('Crie os campos de colaboradores primeiro (informe quantos e clique em "Criar campos").');
    return;
  }
  const personnels = inputs.map(i => i.value.trim()).filter(v => v !== '');
  if(personnels.length === 0) {
    alert('Informe ao menos um Nº Pessoal válido.');
    return;
  }

  // Collect the current rows to duplicate (snapshot)
  const currentRows = Array.from(tableBody.querySelectorAll('tr')).filter(tr => !tr.classList.contains('empty-row'));
  if(currentRows.length === 0) {
    alert('Não há linhas para duplicar.');
    return;
  }

  // We'll build new rows: for each personnel, append copies of currentRows with Nº Pessoal replaced
  const newRows = [];

  personnels.forEach((pers, idx) => {
    currentRows.forEach(tr => {
      // read each value from tr
      const cells = tr.querySelectorAll('td');
      const getVal = (index) => {
        const el = cells[index].querySelector('input,select');
        return el ? (el.value ?? '') : '';
      };
      const prefill = {
        ordens: getVal(0),
        cen: getVal(4),
        np: pers,
        dataIn: getVal(7),
        inicio: getVal(8),
        dataFim: getVal(9),
        termino: getVal(10),
        texto: getVal(12)
      };
      newRows.push(prefill);
    });

    // add an empty spacer row after each group (except maybe last)
    if(idx !== personnels.length - 1) newRows.push(null); // marker for blank row
  });

  // Replace table body with the newRows (clear existing)
  tableBody.innerHTML = '';
  newRows.forEach(item => {
    if(item === null) {
      // empty spacer
      const tr = document.createElement('tr');
      tr.className = 'empty-row';
      const td = document.createElement('td');
      td.colSpan = 14;
      td.innerHTML = '&nbsp;';
      tr.appendChild(td);
      tableBody.appendChild(tr);
    } else {
      addEditableRow(item);
    }
  });

  // Inform user
  alert(`Foram geradas ${personnels.length} grupos (total de linhas: ${newRows.filter(r=>r!==null).length}).`);
}
