/**
 * Rastreador de Vocabulário — backend
 * Cole este arquivo em Extensões > Apps Script (arquivo "Código.gs")
 *
 * Colunas da aba Vocabulário (cabeçalho na linha 5, dados a partir da 6):
 * B Data | C Palavra/Frase | D Pronúncia | E Significado | F Frase de Exemplo
 * G Status | H Biblioteca | I Tipo | J Glossário (JSON, uso interno)
 */

const ABA = 'Vocabulário';
const LINHA_CABECALHO = 5;
const LINHA_INICIAL = 6;
const COL_DATA = 2;
const COL_PALAVRA = 3;
const COL_PRONUNCIA = 4;
const COL_SIGNIFICADO = 5;
const COL_FRASE = 6;
const COL_STATUS = 7;
const COL_BIBLIOTECA = 8;
const COL_TIPO = 9;
const COL_GLOSSARIO = 10;
const TOTAL_COLS = 9; // B..J

const BASE_IA = 'https://generativelanguage.googleapis.com/v1beta';

// Ordem de preferência para conversar. Se um estiver ocupado, cai para o seguinte.
const MODELOS_PREFERIDOS = [
  'gemini-3.6-flash',
  'gemini-3.5-flash',
  'gemini-3.7-flash',
  'gemini-3.8-flash',
  'gemini-flash-latest',
  'gemini-3.5-flash-lite',
  'gemini-3.1-flash-lite'
];

// Modelos de voz, mesma família de vozes do Gemini Live.
const MODELOS_VOZ = [
  'gemini-3.8-flash-tts',
  'gemini-3.1-flash-tts-preview',
  'gemini-3.8-flash-lite-tts',
  'gemini-2.5-flash-preview-tts',
  'gemini-2.5-pro-preview-tts'
];

/**
 * Voz usada para ler as palavras. Opções conhecidas:
 * Zephyr (clara e animada), Puck (jovem e informal), Charon (grave),
 * Kore (firme), Aoede (leve), Leda (jovem), Orus (firme e maduro), Fenrir (intenso).
 */
const VOZ_IA = 'Zephyr';

/**
 * Mude esta marca sempre que mexer no tratamento do áudio.
 * Ela entra na chave do cache, então os áudios antigos são descartados
 * em vez de continuarem tocando por mais seis horas.
 */
const VERSAO_VOZ = 'v4';

/* ---------------------------------------------------------- app */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Meu Vocabulário')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Vocabulário')
    .addItem('Abrir app', 'abrirApp')
    .addItem('Preparar planilha (rodar uma vez)', 'prepararPlanilha')
    .addItem('Testar conexão com o Gemini', 'diagnosticoIA')
    .addToUi();
}

function abrirApp() {
  const html = HtmlService.createHtmlOutputFromFile('Index').setWidth(430).setHeight(680);
  SpreadsheetApp.getUi().showModalDialog(html, 'Meu Vocabulário');
}

/**
 * RODE UMA VEZ: renomeia a coluna D para Pronúncia, cria as colunas
 * Biblioteca / Tipo / Glossário e marca as linhas antigas como "Palavra".
 */
function prepararPlanilha() {
  const sh = aba_();
  sh.getRange(LINHA_CABECALHO, COL_PRONUNCIA).setValue('Pronúncia');
  sh.getRange(LINHA_CABECALHO, COL_BIBLIOTECA).setValue('Biblioteca');
  sh.getRange(LINHA_CABECALHO, COL_TIPO).setValue('Tipo');
  sh.getRange(LINHA_CABECALHO, COL_GLOSSARIO).setValue('Glossário');

  sh.getRange(LINHA_INICIAL, COL_PRONUNCIA, 995, 1).clearDataValidations();
  sh.setColumnWidth(COL_BIBLIOTECA, 150);
  sh.setColumnWidth(COL_TIPO, 80);
  sh.setColumnWidth(COL_GLOSSARIO, 60);

  const ultima = sh.getLastRow();
  if (ultima >= LINHA_INICIAL) {
    const n = ultima - LINHA_INICIAL + 1;
    const palavras = sh.getRange(LINHA_INICIAL, COL_PALAVRA, n, 1).getValues();
    const tipos = sh.getRange(LINHA_INICIAL, COL_TIPO, n, 1).getValues();
    for (let i = 0; i < n; i++) {
      if (String(palavras[i][0] || '').trim() && !String(tipos[i][0] || '').trim()) tipos[i][0] = 'Palavra';
    }
    sh.getRange(LINHA_INICIAL, COL_TIPO, n, 1).setValues(tipos);
  }
  SpreadsheetApp.flush();
}

/* ------------------------------------------------------- planilha */

function aba_() {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA);
  if (!s) throw new Error('Aba "' + ABA + '" não encontrada.');
  return s;
}

function fuso_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || 'America/Sao_Paulo';
}

function chaveDia_(valor) {
  if (valor instanceof Date) return Utilities.formatDate(valor, fuso_(), 'yyyy-MM-dd');
  const t = String(valor || '').trim();
  const br = t.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
  if (br) return br[3] + '-' + br[2] + '-' + br[1];
  const iso = t.match(/^\d{4}-\d{2}-\d{2}/);
  return iso ? iso[0] : '';
}

function lerLinhas_() {
  const sh = aba_();
  const ultima = sh.getLastRow();
  if (ultima < LINHA_INICIAL) return [];
  const valores = sh.getRange(LINHA_INICIAL, COL_DATA, ultima - LINHA_INICIAL + 1, TOTAL_COLS).getValues();
  const linhas = [];
  valores.forEach(function (v, i) {
    if (String(v[1] || '').trim() === '') return;
    let glossario = [];
    try { glossario = JSON.parse(String(v[8] || '[]')) || []; } catch (e) { glossario = []; }
    linhas.push({
      linha: LINHA_INICIAL + i,
      dia: chaveDia_(v[0]),
      dataBr: chaveDia_(v[0]) ? chaveDia_(v[0]).split('-').reverse().join('/') : '',
      palavra: String(v[1]).trim(),
      pronuncia: String(v[2] || '').trim(),
      significado: String(v[3] || '').trim(),
      frase: String(v[4] || '').trim(),
      status: String(v[5] || 'Novo').trim() || 'Novo',
      biblioteca: String(v[6] || '').trim(),
      tipo: String(v[7] || 'Palavra').trim() || 'Palavra',
      glossario: glossario
    });
  });
  return linhas;
}

/** Dados da tela inteira. */
function carregarDados() {
  const linhas = lerLinhas_();
  const hoje = Utilities.formatDate(new Date(), fuso_(), 'yyyy-MM-dd');

  const dias = {};
  linhas.forEach(function (l) { if (l.dia) dias[l.dia] = true; });

  let sequencia = 0;
  const d = new Date();
  if (!dias[hoje]) d.setDate(d.getDate() - 1);
  for (let i = 0; i < 400; i++) {
    const k = Utilities.formatDate(d, fuso_(), 'yyyy-MM-dd');
    if (!dias[k]) break;
    sequencia++;
    d.setDate(d.getDate() - 1);
  }

  const bibs = {};
  linhas.forEach(function (l) { if (l.biblioteca) bibs[l.biblioteca] = (bibs[l.biblioteca] || 0) + 1; });
  const bibliotecas = Object.keys(bibs).sort().map(function (n) { return { nome: n, qtd: bibs[n] }; });

  return {
    total: linhas.length,
    hoje: linhas.filter(function (l) { return l.dia === hoje; }).length,
    sequencia: sequencia,
    dominadas: linhas.filter(function (l) { return l.status === 'Dominado'; }).length,
    bibliotecas: bibliotecas,
    temIA: !!chaveIA_(),
    tema: PropertiesService.getUserProperties().getProperty('TEMA') || '',
    urlPratica: PropertiesService.getUserProperties().getProperty('URL_PRATICA') || '',
    sessoes: lerSessoes_(),
    itens: linhas.slice(-120).reverse()
  };
}

/** Lembra se você prefere o modo claro ou escuro. */
function salvarTema(tema) {
  PropertiesService.getUserProperties()
    .setProperty('TEMA', tema === 'escuro' ? 'escuro' : 'claro');
}

/** Grava uma palavra ou uma frase. */
function salvarItem(dados) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const tipo = dados.tipo === 'Frase' ? 'Frase' : 'Palavra';
    const texto = String(dados.palavra || '').trim();
    const significado = String(dados.significado || '').trim();
    if (!texto) throw new Error(tipo === 'Frase' ? 'Escreva a frase em inglês.' : 'Escreva a palavra em inglês.');
    if (!significado) throw new Error(tipo === 'Frase' ? 'Escreva a tradução da frase.' : 'Escreva o significado.');

    const sh = aba_();
    const linhas = lerLinhas_();

    const repetida = linhas.filter(function (l) {
      return l.palavra.toLowerCase() === texto.toLowerCase();
    })[0];
    if (repetida && !dados.confirmarRepetida) {
      return { repetida: true, mensagem: 'Você já salvou isso em ' + (repetida.dataBr || 'outro dia') + '.' };
    }

    const proxima = linhas.length ? linhas[linhas.length - 1].linha + 1 : LINHA_INICIAL;
    sh.getRange(proxima, COL_DATA, 1, TOTAL_COLS).setValues([[
      new Date(),
      texto,
      String(dados.pronuncia || '').trim(),
      significado,
      String(dados.frase || '').trim(),
      String(dados.status || 'Novo').trim(),
      String(dados.biblioteca || '').trim(),
      tipo,
      JSON.stringify(dados.glossario || [])
    ]]);
    sh.getRange(proxima, COL_DATA).setNumberFormat('dd/MM/yyyy');
    SpreadsheetApp.flush();

    const resumo = carregarDados();
    resumo.salvou = true;
    return resumo;
  } finally {
    lock.releaseLock();
  }
}

function atualizarStatus(linha, status) {
  aba_().getRange(Number(linha), COL_STATUS).setValue(status);
  SpreadsheetApp.flush();
  return carregarDados();
}

/** Move o item para uma biblioteca (pasta). String vazia tira da pasta. */
function moverBiblioteca(linha, biblioteca) {
  aba_().getRange(Number(linha), COL_BIBLIOTECA).setValue(String(biblioteca || '').trim());
  SpreadsheetApp.flush();
  return carregarDados();
}

/** Regrava uma linha já existente, mantendo a data e o status originais. */
function editarItem(linha, dados) {
  const l = Number(linha);
  if (!l || l < LINHA_INICIAL) throw new Error('Linha inválida.');

  const sh = aba_();
  const tipo = String(sh.getRange(l, COL_TIPO).getValue() || 'Palavra').trim() || 'Palavra';
  const texto = String(dados.palavra || '').trim();
  const significado = String(dados.significado || '').trim();
  if (!texto) throw new Error(tipo === 'Frase' ? 'Escreva a frase em inglês.' : 'Escreva a palavra em inglês.');
  if (!significado) throw new Error(tipo === 'Frase' ? 'Escreva a tradução da frase.' : 'Escreva o significado.');

  sh.getRange(l, COL_PALAVRA, 1, 4).setValues([[
    texto,
    String(dados.pronuncia || '').trim(),
    significado,
    String(dados.frase || '').trim()
  ]]);
  sh.getRange(l, COL_BIBLIOTECA).setValue(String(dados.biblioteca || '').trim());
  if (dados.glossario && dados.glossario.length) {
    sh.getRange(l, COL_GLOSSARIO).setValue(JSON.stringify(dados.glossario));
  }
  SpreadsheetApp.flush();

  const resumo = carregarDados();
  resumo.editou = true;
  return resumo;
}

function excluirItem(linha) {
  aba_().deleteRow(Number(linha));
  SpreadsheetApp.flush();
  return carregarDados();
}

/* ------------------------------------------------------ praticar */

const ABA_CONVERSAS = 'Conversas';

function abaConversas_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ABA_CONVERSAS);
  if (!sh) {
    sh = ss.insertSheet(ABA_CONVERSAS);
    sh.getRange(1, 1, 1, 5).setValues([['Data', 'Tema', 'Nível', 'Minutos', 'Transcrição']]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setColumnWidth(5, 600);
  }
  return sh;
}

function lerSessoes_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_CONVERSAS);
  if (!sh || sh.getLastRow() < 2) return [];
  const v = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
  return v.map(function (l, i) {
    return {
      linha: i + 2,
      dataBr: chaveDia_(l[0]) ? chaveDia_(l[0]).split('-').reverse().join('/') : '',
      tema: String(l[1] || ''),
      nivel: String(l[2] || ''),
      minutos: Number(l[3] || 0)
    };
  }).reverse().slice(0, 10);
}

/** Endereço da página externa que roda a conversa por voz. */
function salvarUrlPratica(url) {
  const u = String(url || '').trim();
  if (u.indexOf('https://') !== 0) throw new Error('O endereço precisa começar com https://');
  PropertiesService.getUserProperties().setProperty('URL_PRATICA', u);
  return carregarDados();
}

/** Pacote enviado à sala de voz: chave, voz e as palavras em estudo. */
function dadosDaPratica() {
  const chave = chaveIA_();
  if (!chave) throw new Error('Configure sua chave do Gemini primeiro.');
  const palavras = lerLinhas_()
    .filter(function (l) { return l.tipo === 'Palavra' && l.status !== 'Dominado'; })
    .slice(-30)
    .map(function (l) { return l.palavra; });
  return { chave: chave, voz: VOZ_IA, palavras: palavras };
}

function salvarSessao(dados) {
  const texto = String(dados.transcricao || '').trim();
  if (!texto) return carregarDados();
  abaConversas_().appendRow([
    new Date(),
    String(dados.tema || ''),
    String(dados.nivel || ''),
    Number(dados.minutos || 0),
    texto.slice(0, 45000)
  ]);
  SpreadsheetApp.flush();
  return carregarDados();
}

/* ------------------------------------------------------------ IA */

function chaveIA_() {
  return PropertiesService.getUserProperties().getProperty('CHAVE_GEMINI')
      || PropertiesService.getScriptProperties().getProperty('CHAVE_GEMINI')
      || '';
}

/** Guarda a chave do Google AI Studio na sua conta (não fica na planilha). */
function salvarChaveIA(chave) {
  const c = String(chave || '').trim();
  if (c.length < 20) throw new Error('Essa chave parece incompleta.');
  const props = PropertiesService.getUserProperties();
  props.setProperty('CHAVE_GEMINI', c);
  props.deleteProperty('MODELO_GEMINI');
  props.deleteProperty('LISTA_GEMINI');
  props.deleteProperty('MODELO_VOZ');
  return carregarDados();
}

function apagarChaveIA() {
  const props = PropertiesService.getUserProperties();
  props.deleteProperty('CHAVE_GEMINI');
  props.deleteProperty('MODELO_GEMINI');
  props.deleteProperty('LISTA_GEMINI');
  props.deleteProperty('MODELO_VOZ');
  return carregarDados();
}

/** Lista os modelos de texto que a sua chave realmente tem. */
function modelosDisponiveis_() {
  const r = UrlFetchApp.fetch(BASE_IA + '/models?pageSize=200', {
    method: 'get',
    headers: { 'x-goog-api-key': chaveIA_() },
    muteHttpExceptions: true
  });
  if (r.getResponseCode() !== 200) {
    throw new Error('Não consegui falar com o Google (erro ' + r.getResponseCode() + '). Confira se a chave está certa e se a API Generative Language está ativa no projeto dela.');
  }
  return (JSON.parse(r.getContentText()).models || [])
    .filter(function (m) {
      return (m.supportedGenerationMethods || []).indexOf('generateContent') > -1;
    })
    .map(function (m) { return String(m.name || '').replace('models/', ''); });
}

function serveParaTexto_(nome) {
  return ['image', 'tts', 'transcribe', 'embedding', 'audio', 'robotics', 'computer-use',
          'nano-banana', 'lyria', 'veo', 'deep-research', 'antigravity', 'gemma']
    .every(function (ruim) { return nome.indexOf(ruim) === -1; });
}

/**
 * Fila de modelos para tentar, do melhor para o reserva.
 * Fica guardada para não consultar a lista do Google a cada pergunta.
 */
function filaIA_(refazer) {
  const props = PropertiesService.getUserProperties();
  let fila = null;
  if (!refazer) {
    try { fila = JSON.parse(props.getProperty('LISTA_GEMINI') || 'null'); } catch (e) { fila = null; }
  }
  if (!fila || !fila.length) {
    const disponiveis = modelosDisponiveis_().filter(serveParaTexto_);
    if (!disponiveis.length) throw new Error('Nenhum modelo de texto disponível para essa chave.');
    fila = [];
    MODELOS_PREFERIDOS.forEach(function (n) {
      if (disponiveis.indexOf(n) > -1) fila.push(n);
    });
    disponiveis.forEach(function (n) {
      if (fila.length < 6 && fila.indexOf(n) === -1 && n.indexOf('flash') > -1) fila.push(n);
    });
    if (!fila.length) fila = disponiveis.slice(0, 3);
    props.setProperty('LISTA_GEMINI', JSON.stringify(fila));
  }
  const campeao = props.getProperty('MODELO_GEMINI');
  if (campeao && fila.indexOf(campeao) > -1) {
    fila = [campeao].concat(fila.filter(function (n) { return n !== campeao; }));
  }
  return fila;
}

/** Mostra qual modelo está em uso, a fila de reservas e o que a chave tem. */
function diagnosticoIA() {
  const ui = SpreadsheetApp.getUi();
  if (!chaveIA_()) { ui.alert('Nenhuma chave do Gemini configurada ainda.'); return; }
  try {
    const props = PropertiesService.getUserProperties();
    props.deleteProperty('MODELO_GEMINI');
    props.deleteProperty('MODELO_VOZ');
    const fila = filaIA_(true);
    const lista = modelosDisponiveis_();
    const vozes = lista.filter(function (n) { return n.indexOf('tts') > -1; });
    ui.alert('Conexão com o Gemini',
      'Conversa (tenta nesta ordem):\n' + fila.join('\n') +
      '\n\nVoz disponível:\n' + (vozes.join('\n') || 'nenhum modelo de voz nessa chave') +
      '\n\nTudo que a sua chave enxerga:\n' + lista.join('\n'),
      ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Deu problema', e.message, ui.ButtonSet.OK);
  }
}

function extrairTexto_(txt) {
  const dados = JSON.parse(txt);
  const cand = dados.candidates && dados.candidates[0];
  if (!cand) return '';
  return (cand.content && cand.content.parts || [])
    .filter(function (p) { return p.text && !p.thought; })
    .map(function (p) { return p.text; })
    .join('')
    .trim();
}

/**
 * Chama o Gemini. Se o modelo estiver ocupado (503) ou sumir (404),
 * passa na hora para o próximo da fila em vez de ficar insistindo.
 */
function chamarIA_(partes, sistema, comoJson) {
  const chave = chaveIA_();
  if (!chave) throw new Error('Configure sua chave do Gemini na aba Conversar.');

  const props = PropertiesService.getUserProperties();
  const fila = filaIA_(false);
  let ocupados = 0;

  for (let i = 0; i < fila.length; i++) {
    const modelo = fila[i];
    const corpo = {
      contents: partes,
      generationConfig: { temperature: 0.7, maxOutputTokens: 1400 }
    };
    if (modelo.indexOf('2.5') > -1) corpo.generationConfig.thinkingConfig = { thinkingBudget: 0 };
    if (sistema) corpo.systemInstruction = { parts: [{ text: sistema }] };
    if (comoJson) corpo.generationConfig.responseMimeType = 'application/json';

    const r = UrlFetchApp.fetch(BASE_IA + '/models/' + modelo + ':generateContent', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-goog-api-key': chave },
      payload: JSON.stringify(corpo),
      muteHttpExceptions: true
    });

    const codigo = r.getResponseCode();
    const txt = r.getContentText();

    if (codigo === 200) {
      const saida = extrairTexto_(txt);
      if (saida) {
        props.setProperty('MODELO_GEMINI', modelo);
        return saida;
      }
      continue;
    }

    if (codigo === 503 || codigo === 500 || codigo === 429) { ocupados++; continue; }
    if (codigo === 404) { props.deleteProperty('LISTA_GEMINI'); continue; }

    if (codigo === 400 && txt.indexOf('API key') > -1) throw new Error('Chave do Gemini inválida. Confira em aistudio.google.com/apikey.');
    if (codigo === 403) throw new Error('A chave existe mas não tem permissão. Ative a API Generative Language no projeto dela.');
    if (codigo === 400) throw new Error('O Gemini recusou o pedido. Tente reformular a pergunta.');
    throw new Error('O Gemini não respondeu (erro ' + codigo + ').');
  }

  if (ocupados) throw new Error('Os modelos do Gemini estão ocupados agora. Tente de novo em um minuto.');
  throw new Error('Nenhum modelo da sua chave conseguiu responder. Use o menu Vocabulário > Testar conexão com o Gemini.');
}

/* ------------------------------------------------------------ voz */

/** Lê uma amostra de 16 bits do áudio cru. */
function amostra_(b, i) {
  const baixo = b[i] & 255, alto = b[i + 1] & 255;
  let v = (alto << 8) | baixo;
  return v > 32767 ? v - 65536 : v;
}

/** Escreve uma amostra de 16 bits no áudio cru. */
function gravarAmostra_(b, i, v) {
  if (v > 32767) v = 32767;
  if (v < -32768) v = -32768;
  const u = v < 0 ? v + 65536 : v;
  const baixo = u & 255, alto = (u >> 8) & 255;
  b[i] = baixo > 127 ? baixo - 256 : baixo;
  b[i + 1] = alto > 127 ? alto - 256 : alto;
}

/**
 * Tira o silêncio das pontas (com folga no começo), aplica fade só na saída
 * — é ele que elimina o estalo — e deixa um respiro no final.
 */
function polirPcm_(bruto, taxa) {
  const b = Array.prototype.slice.call(bruto);
  if (b.length % 2 === 1) b.pop(); // byte solto desalinha a onda e vira ruído
  const total = Math.floor(b.length / 2);
  if (total < 200) return b;

  // limiar baixo e folga generosa: consoante fraca no começo não pode sumir
  const limiar = 120;
  let ini = 0;
  while (ini < total && Math.abs(amostra_(b, ini * 2)) < limiar) ini++;
  if (ini >= total) return b;
  ini = Math.max(0, ini - Math.floor(taxa * 0.09));

  let fim = total - 1;
  while (fim > ini && Math.abs(amostra_(b, fim * 2)) < limiar) fim--;
  fim = Math.min(total - 1, fim + Math.floor(taxa * 0.04));

  const corte = b.slice(ini * 2, (fim + 1) * 2);
  const n = Math.floor(corte.length / 2);

  // sem fade na entrada: ele comia o ataque da primeira sílaba
  const saida = Math.min(Math.floor(taxa * 0.06), Math.floor(n / 3));
  for (let i = 0; i < saida; i++) {
    const j = n - 1 - i;
    gravarAmostra_(corte, j * 2, Math.round(amostra_(corte, j * 2) * (i / saida)));
  }

  const respiro = Math.floor(taxa * 0.08) * 2;
  for (let k = 0; k < respiro; k++) corte.push(0);
  return corte;
}

/** Monta um arquivo WAV a partir do áudio cru (PCM 16 bits) que o Gemini devolve. */
function pcmParaWav_(bytes, taxa) {
  const canais = 1, bits = 16;
  const alinhamento = canais * bits / 8;
  const tam = bytes.length;
  const cab = [];
  function texto(s) { for (let i = 0; i < s.length; i++) cab.push(s.charCodeAt(i)); }
  function n32(n) { cab.push(n & 255, (n >> 8) & 255, (n >> 16) & 255, (n >> 24) & 255); }
  function n16(n) { cab.push(n & 255, (n >> 8) & 255); }

  texto('RIFF'); n32(36 + tam); texto('WAVE');
  texto('fmt '); n32(16); n16(1); n16(canais);
  n32(taxa); n32(taxa * alinhamento); n16(alinhamento); n16(bits);
  texto('data'); n32(tam);

  const assinados = cab.map(function (b) { return b > 127 ? b - 256 : b; });
  return Utilities.base64Encode(assinados.concat(bytes));
}

function filaVoz_() {
  const props = PropertiesService.getUserProperties();
  const salvo = props.getProperty('MODELO_VOZ');
  const disponiveis = modelosDisponiveis_().filter(function (n) { return n.indexOf('tts') > -1; });
  let fila = MODELOS_VOZ.filter(function (n) { return disponiveis.indexOf(n) > -1; });
  disponiveis.forEach(function (n) { if (fila.indexOf(n) === -1) fila.push(n); });
  if (salvo && fila.indexOf(salvo) > -1) fila = [salvo].concat(fila.filter(function (n) { return n !== salvo; }));
  return fila;
}

/**
 * Lê o texto em voz alta com uma das vozes do Gemini.
 * Devolve {audio: base64 de um WAV} para o navegador tocar.
 */
function falarIA(texto, voz) {
  const t = String(texto || '').trim();
  if (!t) throw new Error('Nada para falar.');
  if (!chaveIA_()) throw new Error('sem chave');

  const cache = CacheService.getUserCache();
  const marca = 'voz_' + VERSAO_VOZ + '_' + Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, t + '|' + (voz || VOZ_IA)));
  const guardado = cache.get(marca);
  if (guardado) return { audio: guardado };

  const props = PropertiesService.getUserProperties();
  const fila = filaVoz_();
  if (!fila.length) throw new Error('sem voz');

  // Só o texto: qualquer instrução extra acaba sendo lida em voz alta.
  for (let i = 0; i < fila.length; i++) {
    const modelo = fila[i];
    const r = UrlFetchApp.fetch(BASE_IA + '/models/' + modelo + ':generateContent', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-goog-api-key': chaveIA_() },
      payload: JSON.stringify({
        contents: [{ role: 'user', parts: [{ text: t }] }],
        generationConfig: {
          responseModalities: ['AUDIO'],
          speechConfig: {
            voiceConfig: { prebuiltVoiceConfig: { voiceName: voz || VOZ_IA } }
          }
        }
      }),
      muteHttpExceptions: true
    });

    if (r.getResponseCode() !== 200) continue;

    const dados = JSON.parse(r.getContentText());
    const partes = (dados.candidates && dados.candidates[0] &&
                    dados.candidates[0].content && dados.candidates[0].content.parts) || [];
    const pedacos = partes.filter(function (p) { return p.inlineData && p.inlineData.data; });
    if (!pedacos.length) continue;

    const mime = String(pedacos[0].inlineData.mimeType || '');
    const taxa = Number((mime.match(/rate=(\d+)/) || [])[1] || 24000);

    // o áudio pode vir fatiado: junta tudo antes de tratar
    let cru = [];
    pedacos.forEach(function (p) {
      cru = cru.concat(Array.prototype.slice.call(Utilities.base64Decode(p.inlineData.data)));
    });

    const wav = pcmParaWav_(polirPcm_(cru, taxa), taxa);

    props.setProperty('MODELO_VOZ', modelo);
    if (wav.length < 95000) cache.put(marca, wav, 21600); // 6 horas
    return { audio: wav };
  }
  throw new Error('voz ocupada');
}

/* ----------------------------------------------------------- chat */

const SISTEMA_CHAT =
  'Você é um professor de inglês conversando com um estudante brasileiro pelo celular. ' +
  'Responda SEMPRE em português, com os exemplos em inglês seguidos da tradução. ' +
  'Nunca devolva só a tradução seca: explique o sentido e mostre quando usar.\n' +
  'Formato, usando ## para títulos de seção, ** para negrito e - para listas:\n' +
  '## Tradução — uma linha com o sentido principal e a pronúncia simplificada entre parênteses.\n' +
  '## Quando usar — 2 ou 3 situações concretas (trabalho, conversa informal, texto escrito...).\n' +
  '## Exemplos — 2 ou 3 frases em inglês, cada uma com a tradução logo depois.\n' +
  '## Cuidado — registro, erro comum de brasileiro ou diferença entre palavras parecidas.\n' +
  'Se o estudante pedir várias opções (apelidos, sinônimos, formas de dizer algo), ' +
  'liste de 4 a 6 opções em inglês, cada uma com tradução e o tom que passa, e troque a seção Exemplos por essa lista.\n' +
  'Máximo de 230 palavras. Seja direto, sem enrolação e sem repetir a pergunta.\n' +
  'Se a mensagem for uma palavra ou expressão em inglês ou português, termine com uma última linha exatamente neste formato:\n' +
  'SALVAR: palavra em inglês | pronúncia simplificada | significado curto | frase de exemplo em inglês\n' +
  'Se a pergunta for sobre gramática ou outro assunto, responda normalmente e não inclua a linha SALVAR.';

/** Chat da aba Conversar. historico = [{autor:'eu'|'ia', texto:'...'}] */
function perguntarIA(pergunta, historico) {
  const p = String(pergunta || '').trim();
  if (!p) throw new Error('Escreva sua pergunta.');
  const partes = [];
  (historico || []).slice(-6).forEach(function (m) {
    partes.push({ role: m.autor === 'eu' ? 'user' : 'model', parts: [{ text: String(m.texto || '') }] });
  });
  partes.push({ role: 'user', parts: [{ text: p }] });
  return chamarIA_(partes, SISTEMA_CHAT, false);
}

/** Quebra uma frase em palavras com o significado de cada uma. */
function glossarioIA(frase) {
  const f = String(frase || '').trim();
  if (!f) return [];
  const pedido =
    'Frase em inglês: "' + f + '".\n' +
    'Devolva um array JSON. Um objeto por palavra relevante da frase, na ordem em que aparecem, ' +
    'ignorando artigos e pontuação. Cada objeto: {"p": a palavra exatamente como está na frase, ' +
    '"s": o significado dela NESSE contexto em português, no máximo 5 palavras}. Só o array, nada mais.';
  const bruto = chamarIA_([{ role: 'user', parts: [{ text: pedido }] }],
    'Você é um dicionário inglês-português. Responde apenas JSON válido.', true);
  try {
    const arr = JSON.parse(bruto);
    return (arr || []).filter(function (o) { return o && o.p; }).map(function (o) {
      return { p: String(o.p).trim(), s: String(o.s || '').trim() };
    });
  } catch (e) {
    return [];
  }
}

/** Significado de uma palavra solta dentro de uma frase (popup). */
function significadoRapido(palavra, contexto) {
  const p = String(palavra || '').trim();
  if (!p) return '';
  const pedido = 'O que significa "' + p + '"' +
    (contexto ? ' na frase "' + String(contexto).trim() + '"' : '') + '? ' +
    'Responda só com o significado em português, no máximo 8 palavras, sem ponto final.';
  return chamarIA_([{ role: 'user', parts: [{ text: pedido }] }],
    'Você é um dicionário inglês-português, direto e curto.', false);
}
