// Deploy as the script owner with access restricted to "Only myself".
function doGet() {
  centralAutorizar_();
  return HtmlService.createHtmlOutputFromFile('PortalCentral').setTitle('Central de Agentes').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
function centralAutorizar_() {
  const owner = Session.getEffectiveUser().getEmail();
  const visitor = Session.getActiveUser().getEmail();
  if (!owner || (visitor && visitor !== owner)) throw new Error('Este piloto é privado. Entre com a conta proprietária.');
}
function centralAba_() {
  const book = SpreadsheetApp.openById(PLANILHA_ID);
  let sheet = book.getSheetByName('CentralRegistros');
  if (!sheet) {
    sheet = book.insertSheet('CentralRegistros');
    sheet.appendRow(['ID', 'Registro JSON', 'Atualizado em']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}
function centralLer_(sheet) {
  if (sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow()-1, 3).getValues().map((row,i) => ({row:i+2, value:JSON.parse(row[1])}));
}
function centralSnapshot_(sheet) {
  const rows = centralLer_(sheet).map(x => x.value);
  let profile=rows.find(x=>x.id==='perfil');
  if(!profile){profile={id:'perfil',version:1,name:'Grace',assistant:'Minha assistente',start:Utilities.formatDate(new Date(),'America/Sao_Paulo','yyyy-MM-dd')};sheet.appendRow(['perfil',JSON.stringify(profile),new Date().toISOString()]);}
  return {records:rows.filter(x => x.id !== 'perfil'), profile:profile, updated:new Date().toISOString()};
}
function centralApi(request) {
  centralAutorizar_();
  if (!request || !['list','create','update','profile'].includes(request.action)) throw new Error('Comando inválido.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = centralAba_();
    if (request.action === 'list') return centralSnapshot_(sheet);
    const all = centralLer_(sheet);
    let record, existing;
    if (request.action === 'profile') {
      existing = all.find(x=>x.value.id==='perfil');
      record = Object.assign({id:'perfil',version:0},existing ? existing.value : {}, {
        name:centralTexto_(request.name,80), assistant:centralTexto_(request.assistant,80), start:centralData_(request.start)
      });
    } else if (request.action === 'create') {
      if (!['salao','pessoal','biblia'].includes(request.agent) || !['servico','agendamento','tarefa','leitura'].includes(request.type)) throw new Error('Tipo inválido.');
      if ((request.agent==='salao' && !['servico','agendamento'].includes(request.type)) || (request.agent==='pessoal' && request.type!=='tarefa') || (request.agent==='biblia' && request.type!=='leitura')) throw new Error('Agente incompatível.');
      record = {id:Utilities.getUuid(),agent:request.agent,type:request.type,title:centralTexto_(request.title,300),date:centralData_(request.date),time:centralHora_(request.time),weekly:Number(request.weekly)||0,status:'ativo',version:0,created:new Date().toISOString()};
      if (!record.title) throw new Error('Informe uma descrição.');
      if (!Number.isInteger(record.weekly)||record.weekly<0||record.weekly>7) throw new Error('Frequência inválida.');
    } else {
      existing = all.find(x=>x.value.id===request.id && x.value.id!=='perfil');
      if (!existing) throw new Error('Registro não encontrado. Atualize a lista.');
      if (Number(request.version)!==existing.value.version) throw new Error('Este registro mudou em outra janela. Atualize antes de editar.');
      record = Object.assign({},existing.value);
      if (request.title!==undefined) { record.title=centralTexto_(request.title,300); if(!record.title) throw new Error('Descrição obrigatória.'); }
      if (request.date!==undefined) record.date=centralData_(request.date);
      if (request.time!==undefined) record.time=centralHora_(request.time);
      if (request.weekly!==undefined) { const n=Number(request.weekly); if(!Number.isInteger(n)||n<0||n>7) throw new Error('Frequência inválida.'); record.weekly=n; }
      if (request.status!==undefined) {
        if (!['ativo','concluido','cancelado','excluido'].includes(request.status)) throw new Error('Situação inválida.');
        record.status=request.status;
      }
    }
    if(record.type==='agendamento' && record.status==='ativo') {
      if(!record.date || !record.time) throw new Error('Informe data e horário do agendamento.');
      if(all.some(x=>x.value.id!==record.id && x.value.type==='agendamento' && x.value.status==='ativo' && x.value.date===record.date && x.value.time===record.time)) throw new Error('Já existe um agendamento neste horário.');
    }
    record.version=(record.version||0)+1;
    record.updated=new Date().toISOString();
    const row=[record.id,JSON.stringify(record),record.updated];
    if(existing) sheet.getRange(existing.row,1,1,3).setValues([row]); else sheet.appendRow(row);
    SpreadsheetApp.flush();
    return centralSnapshot_(sheet);
  } finally { lock.releaseLock(); }
}
function centralTexto_(value,max) { return String(value||'').trim().slice(0,max); }
function centralData_(value) {
  if(!value) return '';
  const s=String(value); if(!/^\d{4}-\d{2}-\d{2}$/.test(s) || isNaN(Date.parse(s+'T12:00:00Z')) || new Date(s+'T12:00:00Z').toISOString().slice(0,10)!==s) throw new Error('Data inválida.'); return s;
}
function centralHora_(value) { if(!value) return ''; if(!/^([01]\d|2[0-3]):[0-5]\d$/.test(String(value))) throw new Error('Horário inválido.'); return String(value); }
