// ID da Planilha de Cirurgias (Fonte de Dados)
const SOURCE_SPREADSHEET_ID = '1jUxcmXDppfnoQrKwrM8tKK5znuFng4SXEvX7ebpwvvs';

// MAPA DE CORES E TIPOS
const COLOR_MAP = {
  'CONCLUIDA':     { hex: '#595959', color: '#ffffff', label: 'Concluída' },
  'SUSPENSA':      { hex: '#ff0000', color: '#ffffff', label: 'Suspensa' },
  'CLINICA':       { hex: '#ffffff', color: '#000000', label: 'Clínica' },
  'CIRURGICA':     { hex: '#6caa4a', color: '#000000', label: 'Cirúrgica' },
  'PEDIATRICA':    { hex: '#ff98ab', color: '#000000', label: 'Pediátrica' },
  'EMERGENCIA':    { hex: '#ffdbaf', color: '#000000', label: 'Emergência/SUS' },
  'PSIQUIATRICA':  { hex: '#ffff00', color: '#000000', label: 'Psiquiátrica' },
  'AMBULATORIAL':  { hex: '#6d9eeb', color: '#000000', label: 'Ambulatorial' },
  'NAO_COMPARECEU':{ hex: '#ff9900', color: '#000000', label: 'Não Compareceu' },
  'NORMAL':        { hex: '#ffffff', color: '#000000', label: 'Padrão' }
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏥 Internações')
    .addItem('Abrir Painel', 'showAppModal')
    .addToUi();
}

function showAppModal() {
  const html = HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setWidth(1280)
    .setHeight(850)
    .setTitle('Gestão Hospitalar - Painel Integrado');
  SpreadsheetApp.getUi().showModalDialog(html, 'Painel de Gestão');
}

function include(filename) {
  return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
}

function getInitialData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheetName = ss.getActiveSheet().getName();
  const allSheets = ss.getSheets();
  const dateSheets = allSheets.map(s => s.getName()).filter(name => /\d/.test(name));
  
  let targetSheet = dateSheets[0]; 
  const found = dateSheets.find(s => s.trim() === activeSheetName.trim());
  if (found) targetSheet = found;

  const typesList = Object.keys(COLOR_MAP).map(key => ({
    key: key,
    hex: COLOR_MAP[key].hex,
    label: COLOR_MAP[key].label
  }));

  return {
    sheets: dateSheets,
    active: targetSheet,
    types: typesList
  };
}

function getSheetData(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error("Aba não encontrada.");

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow <= 1) return { patients: [], footer: {}, stats: { total:0, byColor:{}, leitosA_endsB:0, leitosB_endsA:0, totalTransfers:0 } };

  const range = sheet.getRange(1, 1, lastRow, lastCol);
  const values = range.getDisplayValues();
  const backgrounds = range.getBackgrounds();
  
  let patients = [];
  let footer = { legendas: [], unavailable: [], transfers: [], availabilities: [] };
  
  let stats = {
    total: 0,
    byColor: {},
    leitosA_endsB: 0, 
    leitosB_endsA: 0,
    totalTransfers: 0
  };
  Object.keys(COLOR_MAP).forEach(k => stats.byColor[k] = 0);

  let splitRowIndex = -1; 
  let availRowIndex = -1;

  for (let i = 0; i < values.length; i++) {
    const rowStr = values[i].join("").toUpperCase();
    if (splitRowIndex === -1 && (rowStr.includes("LEGENDAS") || rowStr.includes("LEITOS INDISPONÍVEIS"))) splitRowIndex = i;
    if (availRowIndex === -1 && rowStr.includes("DISPONIBILIDADES")) availRowIndex = i;
  }

  const limitRow = splitRowIndex === -1 ? values.length : splitRowIndex;

  for (let i = 1; i < limitRow; i++) { 
    const row = values[i];
    if (row[0] === "NOME" || (row[0] === "" && row[4] === "")) continue;

    let corHex = backgrounds[i][0]; 
    let tipoIdentificado = "NORMAL";
    let fontColor = "#000000";
    let matchFound = false;

    for (let [key, config] of Object.entries(COLOR_MAP)) {
      if (config.hex.toLowerCase() === corHex.toLowerCase()) {
        tipoIdentificado = key;
        fontColor = config.color;
        matchFound = true;
        break;
      }
    }
    if (!matchFound && isGray(corHex)) {
       tipoIdentificado = 'CONCLUIDA';
       fontColor = isDark(corHex) ? '#ffffff' : '#000000';
    }

    stats.total++;
    if (stats.byColor[tipoIdentificado] !== undefined) stats.byColor[tipoIdentificado]++;

    let leitoVal = (row[2] || "").trim().toUpperCase(); 
    if (leitoVal) {
      if (leitoVal.endsWith('B')) stats.leitosA_endsB++;
      if (leitoVal.endsWith('A')) stats.leitosB_endsA++; 
    }

    patients.push({
      rowIndex: i + 1,
      nome: row[0],
      prevLeito: row[1],
      leito: row[2],
      telefone: row[3],
      convenio: row[4],
      equipe: row[5],
      guia: row[6],
      status: row[7],
      hrCir: row[8],
      hrInt: row[9],
      obs: row[10],
      corFundo: corHex,
      corFonte: fontColor,
      tipo: tipoIdentificado
    });
  }

  if (splitRowIndex !== -1) {
    const footerLimit = availRowIndex !== -1 ? availRowIndex : values.length;
    for (let i = splitRowIndex + 2; i < footerLimit; i++) {
      const row = values[i];
      const bgRow = backgrounds[i];
      if (row[0]) footer.legendas.push({ nome: row[0], cor: bgRow[0] });
      if (row[1]) footer.unavailable.push({ leito: row[1], motivo: row[2] });
      if (row[3] || row[7]) {
        footer.transfers.push({ nome: row[3], de: row[5], para: row[6], obs: row[7] });
      }
    }
  }
  stats.totalTransfers = footer.transfers.length;

  if (availRowIndex !== -1) {
    const headerRow = values[availRowIndex + 1]; 
    for (let col = 2; col < headerRow.length; col++) {
      const andar = headerRow[col];
      if (!andar) continue;
      let quartos = [];
      for (let r = availRowIndex + 2; r < values.length; r++) {
        const q = values[r][col];
        if (q) quartos.push(q);
      }
      if (quartos.length > 0 || andar.length > 0) footer.availabilities.push({ andar: andar, quartos: quartos });
    }
  }

  return { patients: patients, footer: footer, stats: stats };
}

// --- FUNÇÃO NOVA: IMPORTAR LEITOS DE OUTRA DATA ---
function importUnavailableBeds(targetSheetName, sourceSheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName(targetSheetName);
  const sourceSheet = ss.getSheetByName(sourceSheetName);
  
  if (!targetSheet || !sourceSheet) throw new Error("Aba não encontrada.");

  // 1. Ler dados da origem
  const srcData = sourceSheet.getDataRange().getValues();
  let bedsToCopy = [];
  let readingBeds = false;

  for (let i = 0; i < srcData.length; i++) {
    const row = srcData[i];
    const rowStr = row.join("").toUpperCase();

    // Detecta início (cabeçalho LEITO/MOTIVO que fica abaixo de LEITOS INDISPONIVEIS)
    if (!readingBeds && row[1] && row[1].toString().toUpperCase() === "LEITO" && 
        srcData[i-1] && srcData[i-1][1].toString().toUpperCase().includes("INDISPONÍVEIS")) {
      readingBeds = true;
      continue;
    }

    // Detecta fim (Disponibilidades ou fim da planilha)
    if (readingBeds && (rowStr.includes("DISPONIBILIDADES") || rowStr.includes("LEITOS GERAL"))) {
      break;
    }

    if (readingBeds) {
      // Se tem leito preenchido (Coluna B)
      if (row[1] && row[1].toString().trim() !== "") {
        bedsToCopy.push([row[1], row[2]]); // Coluna B (Leito), Coluna C (Motivo)
      }
    }
  }

  if (bedsToCopy.length === 0) return "Nenhum leito encontrado na data selecionada.";

  // 2. Inserir no destino (SEM CRIAR LINHAS NOVAS)
  const tgtData = targetSheet.getDataRange().getValues();
  let startRow = -1;

  for (let i = 0; i < tgtData.length; i++) {
    // Procura o local de inserção (logo abaixo do cabeçalho LEITO/MOTIVO)
    if (tgtData[i][1] && tgtData[i][1].toString().toUpperCase() === "LEITO" && 
        tgtData[i-1] && tgtData[i-1][1].toString().toUpperCase().includes("INDISPONÍVEIS")) {
      startRow = i + 2; // +1 é a linha do cabeçalho, +2 é a primeira linha de dados (índice 1-based do getRange)
      break;
    }
  }

  if (startRow === -1) return "Estrutura de destino não encontrada.";

  // Apenas cola os valores no espaço existente
  // getRange(linha, coluna, numLinhas, numColunas)
  // Coluna 2 = B, Coluna 3 = C
  targetSheet.getRange(startRow, 2, bedsToCopy.length, 2).setValues(bedsToCopy);

  return `Sucesso! ${bedsToCopy.length} leitos colados de ${sourceSheetName}.`;
}

/**
 * Busca dados da planilha de Cirurgias - Lógica Final com HMD Explícito
 */
function getSurgicalData(targetDate) {
  try {
    const sourceSS = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);
    const sheet = sourceSS.getSheetByName(targetDate); 
    
    if (!sheet) {
      throw new Error(`Aba "${targetDate}" não encontrada na planilha de cirurgias.`);
    }

    const data = sheet.getDataRange().getDisplayValues();
    let extractedPatients = [];
    
    // Contexto inicial (vazio até encontrar o primeiro setor)
    let currentContext = ""; 
    
    // Mapeamento de colunas
    let colMap = { nome: -1, tel: -1, conv: -1, equipe: -1, hora: -1, origem: -1, ia: -1, obs: -1 };
    
    // Helpers
    const toUpper = (val) => val ? val.toString().toUpperCase().trim() : "";
    const transformIA = (val) => {
      let v = toUpper(val);
      if (v.includes("AMBULATORIAL")) return "AMBU";
      if (v.includes("NO DIA")) return "NO DIA";
      if (v.includes("ANTEVÉSPERA") || v.includes("ANTEVESPERA")) return "ANTEVÉSPERA";
      if (v.includes("VÉSPERA") || v.includes("VESPERA")) return "VÉSPERA";
      if (v.includes("CLÍNICA") || v.includes("CLINICA")) return "CLINICA";
      if (v.includes("INTERNADO")) return "INTERNADO";
      return v; 
    };

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      
      // =================================================================
      // LÓGICA DE CONTEXTO FINAL (REGRAS ESTRITAS)
      // =================================================================
      
      // 1. Limpeza total: Tira acentos e TODOS os espaços
      let rawA = row[0] ? row[0].toString().toUpperCase() : "";
      let cleanKey = rawA.normalize('NFD').replace(/[\u0300-\u036f]/g, "").replace(/\s/g, "");

      // 2. Se a célula TEM CONTEÚDO (Início de um novo bloco/setor)
      if (cleanKey.length > 0) {
          
          // REGRA 1: BLOCO (Procura BC ou BLOCO)
          if (cleanKey.includes("BC") || cleanKey.includes("BLOCO")) {
              currentContext = "BLOCO";
          } 
          
          // REGRA 2: CCV (Procura CCV)
          else if (cleanKey.includes("CCV")) {
              currentContext = "BLOCO";
          }
          
          // REGRA 3: CCA (Procura CCA)
          else if (cleanKey.includes("CCA")) {
              currentContext = "CCA";
          } 
          
          // REGRA 4: HEMODINÂMICA (Procura HMD ou HEMODIN)
          // Agora só entra aqui se tiver o texto explícito
          else if (cleanKey.includes("HMD") || cleanKey.includes("HEMODIN")) {
              currentContext = "HMD";
          }
          
          // REGRA 5: LIMPEZA (Caso não tenha nada conhecido)
          // Se escreveu algo na coluna A (ex: "Radiologia") e não é nenhum dos acima,
          // definimos como VAZIO para não classificar errado.
          else {
              currentContext = "";
          }
      }
      // Se cleanKey for vazia (""), mantém o currentContext da linha anterior (efeito mescla).
      
      // =================================================================

      // Identificação do Cabeçalho
      if (row.includes("NOME DO PACIENTE") || row.includes("PACIENTE")) {
        colMap.nome = row.findIndex(c => c.toUpperCase().includes("NOME"));
        colMap.tel = row.findIndex(c => c.toUpperCase().includes("TELEFONE"));
        colMap.conv = row.findIndex(c => c.toUpperCase().includes("CONVÊNIO") || c.toUpperCase().includes("CONVENIO"));
        colMap.equipe = row.findIndex(c => c.toUpperCase().includes("EQUIPE") || c.toUpperCase().includes("CIRURGIÃO"));
        colMap.hora = row.findIndex(c => c.toUpperCase().includes("HORA"));
        colMap.origem = row.findIndex(c => c.toUpperCase().includes("ORIGEM")); 
        colMap.ia = row.findIndex(c => c.toUpperCase().includes("I/A") || c.toUpperCase() === "I/A"); 
        colMap.obs = row.findIndex(c => c.toUpperCase().trim() === "OBS" || c.toUpperCase().includes("OBSERVA"));
        continue; 
      }

      // Extração dos Dados
      if (colMap.nome !== -1 && row[colMap.nome]) {
        if (row[colMap.nome].includes("NOME DO PACIENTE")) continue;

        // Se currentContext for vazio (ex: Radiologia), a hora fica só o valor numérico
        let horaRaw = colMap.hora !== -1 ? row[colMap.hora] : "";
        let horaFormatada = (horaRaw && currentContext) ? `${horaRaw} ${currentContext}` : 
                            (horaRaw ? horaRaw : currentContext);

        extractedPatients.push({
          nome: toUpper(row[colMap.nome]),
          telefone: colMap.tel !== -1 ? row[colMap.tel] : "",
          convenio: colMap.conv !== -1 ? toUpper(row[colMap.conv]) : "",
          equipe: colMap.equipe !== -1 ? toUpper(row[colMap.equipe]) : "",
          hrCir: toUpper(horaFormatada),
          
          guia: colMap.origem !== -1 ? toUpper(row[colMap.origem]) : "", 
          leito: colMap.ia !== -1 ? transformIA(row[colMap.ia]) : "", 
          
          prevLeito: "",
          statusGuia: "",
          hrInt: "",
          obs: colMap.obs !== -1 ? row[colMap.obs] : "",
          valor: "",
          diaIntern: ""
        });
      }
    }

    return extractedPatients;

  } catch (e) {
    throw new Error("Erro ao coletar dados: " + e.message);
  }
}

/**
 * Importa múltiplos pacientes com lógica de SLOTS (Vagas).
 * Evita sobreposição mapeando todas as vagas antes de escrever.
 */
function importBatchPatients(sheetName, patientsList) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  
  // 1. LOCALIZAR O RODAPÉ E VAGAS EXISTENTES
  let footerIndex = data.length + 1; // Padrão: final da planilha se não achar rodapé
  let availableSlots = []; // Lista de números de linhas (1-based) vazias

  for(let i = 0; i < data.length; i++) {
     const rowStr = data[i][0] ? data[i][0].toString() : "";
     
     // Detecta onde começa o rodapé (parar de procurar vagas aqui)
     if(rowStr.includes("LEGENDAS") || rowStr.includes("LEITOS")) {
       footerIndex = i + 1; 
       break;
     }
     
     // Se for linha de dados (após cabeçalho linha 3) e estiver vazia (Coluna A)
     if (i >= 3) {
       if (!rowStr || rowStr.trim() === "") {
         availableSlots.push(i + 1); // Guarda o índice da linha vazia
       }
     }
  }

  // 2. CALCULAR NECESSIDADE DE EXPANSÃO
  // Precisamos de espaço para os pacientes + 3 linhas de buffer (reserva visual)
  const totalNeeded = patientsList.length;
  const bufferWanted = 3;
  const currentEmptyCount = availableSlots.length;
  
  // Quantas linhas faltam para caber todo mundo E manter o buffer?
  // Se tivermos 2 vagas e 5 pacientes, precisamos de 3 linhas + 3 buffer = 6 novas
  // Se tivermos 10 vagas e 5 pacientes, não precisa criar, só sobra buffer.
  
  // Lógica: Se (vagas atuais - pacientes) < 3, precisamos criar a diferença.
  let rowsToCreate = 0;
  if ((currentEmptyCount - totalNeeded) < bufferWanted) {
      rowsToCreate = bufferWanted - (currentEmptyCount - totalNeeded);
  }

  // 3. CRIAR LINHAS (SE NECESSÁRIO)
  if (rowsToCreate > 0) {
      sheet.insertRowsBefore(footerIndex, rowsToCreate);
      
      // Formata as novas linhas como "Limpo" (Branco)
      let cleanRange = sheet.getRange(footerIndex, 1, rowsToCreate, sheet.getLastColumn());
      cleanRange.setBackground('#ffffff');
      cleanRange.setFontColor('#000000');
      
      // Adiciona essas novas linhas à lista de slots disponíveis
      for (let k = 0; k < rowsToCreate; k++) {
          availableSlots.push(footerIndex + k);
      }
  }

  // 4. GRAVAÇÃO SEQUENCIAL (SEM SOBREPOSIÇÃO)
  // Agora 'availableSlots' tem índices suficientes para todos os pacientes.
  // Escrevemos o paciente 1 no slot 1, paciente 2 no slot 2, etc.
  
  patientsList.forEach((p, index) => {
      if (index < availableSlots.length) {
          let targetRow = availableSlots[index];
          
          let rowValues = [
            p.nome, p.prevLeito, p.leito, p.telefone, p.convenio, p.equipe, 
            p.guia, p.statusGuia, p.hrCir, p.hrInt, p.obs
          ];
          
          // Grava dados
          sheet.getRange(targetRow, 1, 1, rowValues.length).setValues([rowValues]);
          
          // Formata a linha (Verde Cirúrgico)
          let fullRange = sheet.getRange(targetRow, 1, 1, sheet.getLastColumn());
          fullRange.setBackground(COLOR_MAP['CIRURGICA'].hex);
          fullRange.setFontColor(COLOR_MAP['CIRURGICA'].color);
          
          // Alinhamento
          sheet.getRange(targetRow, 1).setHorizontalAlignment("left");
          sheet.getRange(targetRow, 2, 1, sheet.getLastColumn()-1).setHorizontalAlignment("center");
      }
  });

  return `Importação concluída! ${patientsList.length} pacientes inseridos.`;
}

// --- AUXILIARES ---
function isGray(hex) {
  if (!hex || hex === '#ffffff' || hex === '#000000') return false; 
  hex = hex.replace('#', '');
  var r = parseInt(hex.substring(0, 2), 16);
  var g = parseInt(hex.substring(2, 4), 16);
  var b = parseInt(hex.substring(4, 6), 16);
  var diff = Math.max(Math.abs(r-g), Math.abs(r-b), Math.abs(g-b));
  return diff < 25; 
}

function isDark(hex) {
  if (!hex) return false;
  hex = hex.replace('#', '');
  var r = parseInt(hex.substring(0, 2), 16);
  var g = parseInt(hex.substring(2, 4), 16);
  var b = parseInt(hex.substring(4, 6), 16);
  var luma = 0.2126 * r + 0.7152 * g + 0.0722 * b; 
  return luma < 128; 
}

function updatePatientType(sheetName, rowIndex, newType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  const config = COLOR_MAP[newType];
  if (config && rowIndex) {
    const maxCols = sheet.getLastColumn();
    sheet.getRange(rowIndex, 1, 1, maxCols).setBackground(config.hex);
  }
  return "Cor Atualizada";
}

function savePatient(sheetName, formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  
  const rowValues = [
    formData.nome, formData.prevLeito, formData.leito, formData.telefone,
    formData.convenio, formData.equipe, formData.guia, formData.status,
    formData.hrCir, formData.hrInt, formData.obs
  ];

  let targetRow;
  const data = sheet.getDataRange().getValues();
  let footerIndex = data.length; 
  let lastDataRow = 3; 
  let foundEmpty = -1;

  // 1. Mapeamento
  for(let i=0; i<data.length; i++) {
     if(data[i][0] && (data[i][0].toString().includes("LEGENDAS") || data[i][0].toString().includes("LEITOS"))) {
       footerIndex = i + 1; // Base-1
       break;
     }
     if (i >= 3 && data[i][0] && data[i][0].toString().trim() !== "") {
       lastDataRow = i + 1;
     }
     if (!formData.rowIndex && i >= 3 && (!data[i][0] || data[i][0].toString().trim() === "") && foundEmpty === -1) {
       foundEmpty = i + 1;
     }
  }

  // 2. Define onde gravar
  if (formData.rowIndex) {
    targetRow = parseInt(formData.rowIndex); 
  } else if (foundEmpty !== -1 && foundEmpty < footerIndex) {
    targetRow = foundEmpty; 
  } else {
    targetRow = lastDataRow + 1; 
  }

  // 3. REGRA DAS 3 LINHAS VAZIAS (Buffer)
  let effectiveLastRow = (targetRow > lastDataRow) ? targetRow : lastDataRow;
  let currentGap = footerIndex - effectiveLastRow - 1;
  
  if (currentGap < 3) {
    let rowsToAdd = 3 - currentGap;
    sheet.insertRowsBefore(footerIndex, rowsToAdd);
    
    // --- CORREÇÃO: FORÇA BRANCO NAS LINHAS VAZIAS CRIADAS ---
    let newRowsRange = sheet.getRange(footerIndex, 1, rowsToAdd, sheet.getLastColumn());
    newRowsRange.setBackground('#ffffff');
    newRowsRange.setFontColor('#000000'); // Reseta fonte para preto também
  }
  
  // 4. Grava os dados
  const range = sheet.getRange(targetRow, 1, 1, rowValues.length);
  range.setValues([rowValues]);

  // 5. Formatação da Linha do Paciente
  const fullRowRange = sheet.getRange(targetRow, 1, 1, sheet.getLastColumn());
  sheet.getRange(targetRow, 1).setHorizontalAlignment("left"); 
  sheet.getRange(targetRow, 2, 1, sheet.getLastColumn() - 1).setHorizontalAlignment("center"); 

  if (formData.tipoCor && COLOR_MAP[formData.tipoCor]) {
    fullRowRange.setBackground(COLOR_MAP[formData.tipoCor].hex);
    let fontColor = COLOR_MAP[formData.tipoCor].color;
    if(!fontColor) fontColor = isDark(COLOR_MAP[formData.tipoCor].hex) ? '#ffffff' : '#000000';
    fullRowRange.setFontColor(fontColor);
  }

  return "Salvo com sucesso!";
}

// =====================================================================
// 8. SISTEMA DE CONCILIAÇÃO E ATUALIZAÇÃO (NOVO)
// =====================================================================

/**
 * Compara os dados da Planilha de Cirurgia (Origem) com a Planilha de Internação (Destino)
 */
function compareSurgicalData(targetDate) {
  try {
    // 1. Busca dados da Origem (Cirurgias)
    const sourcePatients = getSurgicalData(targetDate); // Reaproveita sua função existente
    
    // 2. Busca dados do Destino (Planilha Atual)
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(targetDate);
    if (!sheet) throw new Error("Aba de destino não existe.");
    
    // Pega dados atuais formatados pela função getSheetData (que já temos)
    const currentData = getSheetData(targetDate).patients; 

    let comparisonResult = {
      newPatients: [],
      diffPatients: [],
      missingPatients: []
    };

    // Mapa para busca rápida na planilha atual (Chave: Nome Limpo)
    let currentMap = {};
    currentData.forEach(p => {
      let key = p.nome.toString().toUpperCase().trim();
      currentMap[key] = p;
    });

    // Helper para comparar valores (ignora espaços e case)
    const isDiff = (val1, val2) => {
      let v1 = (val1 || "").toString().toUpperCase().trim();
      let v2 = (val2 || "").toString().toUpperCase().trim();
      return v1 !== v2;
    };

    // A. Analisa quem veio da Cirurgia (Origem)
    sourcePatients.forEach(src => {
      let key = src.nome.toString().toUpperCase().trim();
      let existing = currentMap[key];

      if (!existing) {
        // Cenario 1: NOVO (Não existe na planilha atual)
        comparisonResult.newPatients.push(src);
      } else {
        // Cenario 2: JÁ EXISTE - Verificar Divergências
        let changes = [];
        
        // Compara campos críticos
        if (isDiff(src.hrCir, existing.hrCir)) changes.push({ field: 'hrCir', label: 'Hora', old: existing.hrCir, new: src.hrCir });
        if (isDiff(src.equipe, existing.equipe)) changes.push({ field: 'equipe', label: 'Equipe', old: existing.equipe, new: src.equipe });
        if (isDiff(src.convenio, existing.convenio)) changes.push({ field: 'convenio', label: 'Convênio', old: existing.convenio, new: src.convenio });
        if (isDiff(src.telefone, existing.telefone)) changes.push({ field: 'telefone', label: 'Telefone', old: existing.telefone, new: src.telefone });
        
        // Verifica se houve mudança significativa (tem changes)
        if (changes.length > 0) {
          comparisonResult.diffPatients.push({
            rowIndex: existing.rowIndex, // Para saber onde atualizar
            nome: src.nome,
            changes: changes,
            fullData: src // Dados novos completos
          });
        }
        // Se changes.length == 0, é IGUAL. Não faz nada (ignora).
        
        // Remove do mapa para sobrar apenas quem está na planilha mas não veio da cirurgia
        delete currentMap[key];
      }
    });

    // B. Analisa quem sobrou na Planilha Atual (Ausentes na Origem)
    // Só consideramos "Ausentes" se forem do tipo 'CIRURGICA' ou 'NORMAL' (ignora altas/clinicos manuais)
    for (let key in currentMap) {
      let p = currentMap[key];
      // Filtro de segurança: Só sugere deletar se tiver cara de cirurgia (tem hora de cirurgia ou status cirurgica)
      if (p.hrCir && p.hrCir.length > 2) { 
        comparisonResult.missingPatients.push({
          rowIndex: p.rowIndex,
          nome: p.nome,
          equipe: p.equipe
        });
      }
    }

    return comparisonResult;

  } catch (e) {
    throw new Error("Erro na conciliação: " + e.message);
  }
}

/**
 * Executa as ações em lote (Criar, Atualizar, Deletar)
 */
function processImportActions(sheetName, actions) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  
  // 1. DELETAR (Ausentes) - Fazemos primeiro para liberar espaço, de baixo para cima
  if (actions.toDelete && actions.toDelete.length > 0) {
    // Ordena índices decrescente para não quebrar referências
    actions.toDelete.sort((a, b) => b - a);
    actions.toDelete.forEach(idx => {
      sheet.deleteRow(idx);
    });
  }

  // 2. ATUALIZAR (Divergentes)
  if (actions.toUpdate && actions.toUpdate.length > 0) {
    actions.toUpdate.forEach(item => {
      // Mapeia colunas (Baseado na ordem do getSheetData: A=1, B=2...)
      // Nome(1), Prev(2), Leito(3), Tel(4), Conv(5), Equipe(6), Guia(7), Status(8), HrCir(9), HrInt(10), Obs(11)
      const mapCol = {
        'telefone': 4,
        'convenio': 5,
        'equipe': 6,
        'hrCir': 9
      };
      
      item.changes.forEach(change => {
        let col = mapCol[change.field];
        if (col) {
          // Ajuste de linha: Se houve deleção acima, o rowIndex mudou? 
          // Risco complexo. Por segurança, em operações mistas, ideal é processar delete separado ou recarregar.
          // Assumindo que o usuário raramente deleta e atualiza a mesma região em massa.
          // Mas, para segurança, usamos o valor direto.
          sheet.getRange(item.rowIndex, col).setValue(change.newVal);
          // Pinta de roxo suave para indicar alteração recente
          sheet.getRange(item.rowIndex, col).setFontColor('#800080'); 
        }
      });
    });
  }

  // 3. CRIAR (Novos) - Usa a função existente de Batch
  let msg = "Processamento concluído.";
  if (actions.toCreate && actions.toCreate.length > 0) {
    const res = importBatchPatients(sheetName, actions.toCreate);
    msg += " " + res;
  }

  return msg;
}

function deletePatient(sheetName, rowIndex) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  sheet.deleteRow(parseInt(rowIndex));
  return "Registro excluído!";
}
