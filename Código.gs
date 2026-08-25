const CACHE_LOG_KEY = 'PAINEL_LOG_EXECUCAO';
const CACHE_INDICADORES_KEY = 'PAINEL_INDICADORES_CACHE';
const CACHE_INVENTARIO_KEY = 'PAINEL_INVENTARIO_CACHE';

const CONFIG = {
  PLANILHA_ID: '1qpSTiuKVqOnNJAbVlTWniE93l12mpzbcJT8ScUjImLQ',
  ABA_RESUMO: 'Resumo_HC',
  ABA_RELATORIO_DRIVE: 'Relatorio_Drive',
  ABA_INVENTARIO: 'Inventario',
  ABA_BASE_HC: 'Base_HC',
  PASTA_ID: '14ozLiqto91tnl2n-juJY7Xz4oylup2cY'
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Painel Yubico')
    .addItem('Abrir painel', 'abrirPainel')
    .addItem('Abrir tela TV', 'abrirTelaTV')
    .addItem('Atualizar indicadores', 'atualizarTudo')
    .addItem('Autorizar acesso ao Drive', 'autorizarAcessoDrive')
    .addToUi();
}

function obterUrlWebApp() {
  return ScriptApp.getService().getUrl();
}

function doGet(e) {
  const tela = e && e.parameter && e.parameter.tela;
  const arquivo = String(tela || '').toLowerCase() === 'tv' ? 'TV' : 'Index';
  return HtmlService.createHtmlOutputFromFile(arquivo).setTitle(arquivo === 'TV' ? 'Controle Yubico - TV' : 'Painel Yubico');
}

function autorizarAcessoDrive() {
  const pasta = DriveApp.getFolderById(CONFIG.PASTA_ID);
  const nome = pasta.getName();
  SpreadsheetApp.getUi().alert('Acesso ao Drive autorizado para a pasta: ' + nome + '. Agora tente o upload novamente.');
}

function abrirPainel() {
  const html = HtmlService.createHtmlOutputFromFile('Index')
    .setWidth(980)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Painel Yubico');
}

function abrirTelaUpload() {
  const html = HtmlService.createHtmlOutputFromFile('Upload')
    .setWidth(900)
    .setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(html, 'Enviar PDFs para o Drive');
}

function abrirTelaTV() {
  const html = HtmlService.createHtmlOutputFromFile('TV')
    .setWidth(1400)
    .setHeight(850);
  SpreadsheetApp.getUi().showModalDialog(html, 'Controle Yubico - TV');
}

function obterArvorePastas() {
  const raiz = DriveApp.getFolderById(CONFIG.PASTA_ID);
  return montarArvorePastas_(raiz, new Set());
}

function montarArvorePastas_(pasta, visitadas) {
  const id = pasta.getId();
  if (visitadas.has(id)) return null;
  visitadas.add(id);

  const filhos = [];
  const subpastas = pasta.getFolders();
  while (subpastas.hasNext()) {
    const arvoreFilha = montarArvorePastas_(subpastas.next(), visitadas);
    if (arvoreFilha) filhos.push(arvoreFilha);
  }
  filhos.sort(function(a, b) { return a.nome.localeCompare(b.nome, 'pt-BR'); });

  return { id: id, nome: pasta.getName(), filhos: filhos };
}

function uploadPdfParaPasta(pastaId, nomeArquivo, conteudoBase64) {
  if (!pastaId || !nomeArquivo || !conteudoBase64) throw new Error('Dados do arquivo incompletos.');
  if (!/\.pdf$/i.test(nomeArquivo)) throw new Error('Apenas arquivos PDF são permitidos.');

  const pasta = DriveApp.getFolderById(pastaId);
  const bytes = Utilities.base64Decode(conteudoBase64);
  const blob = Utilities.newBlob(bytes, MimeType.PDF, nomeArquivo);
  const arquivo = pasta.createFile(blob);

  return { id: arquivo.getId(), nome: arquivo.getName(), url: arquivo.getUrl() };
}

function obterDadosPainel() {
  let indicadores = null;
  const salvo = PropertiesService.getScriptProperties().getProperty(CACHE_INDICADORES_KEY);
  if (salvo) {
    try { indicadores = JSON.parse(salvo); } catch (erro) { indicadores = null; }
  }
  if (!indicadores) indicadores = calcularIndicadores_();
  return { indicadores: indicadores, processamento: obterUltimoProcessamento_(), log: obterLogExecucao() };
}

function obterDadosInventario() {
  const salvo = PropertiesService.getScriptProperties().getProperty(CACHE_INVENTARIO_KEY);
  if (salvo) {
    try { return JSON.parse(salvo); } catch (erro) {}
  }
  const planilha = SpreadsheetApp.openById(CONFIG.PLANILHA_ID);
  const aba = planilha.getSheetByName(CONFIG.ABA_INVENTARIO);
  if (!aba || aba.getLastRow() < 2) {
    const vazio = { atual: null, historico: [] };
    PropertiesService.getScriptProperties().setProperty(CACHE_INVENTARIO_KEY, JSON.stringify(vazio));
    return vazio;
  }
  const ultimaLinha = aba.getLastRow();
  const primeiraLinha = Math.max(2, ultimaLinha - 19);
  const valores = aba.getRange(primeiraLinha, 1, ultimaLinha - primeiraLinha + 1, 5).getValues();
  const registros = valores.filter(function(linha) { return linha[0] instanceof Date || linha[0]; }).map(function(linha) {
    return { timestamp: new Date(linha[0]).getTime(), data: formatarData_(new Date(linha[0])), yubico300: Number(linha[1]) || 0, yubico325: Number(linha[2]) || 0, yubico600: Number(linha[3]) || 0, registradoPor: String(linha[4] || '') };
  }).sort(function(a, b) { return b.timestamp - a.timestamp; }).slice(0, 20);
  registros.forEach(function(registro) { delete registro.timestamp; });
  const resultado = { atual: registros.length ? registros[0] : null, historico: registros };
  PropertiesService.getScriptProperties().setProperty(CACHE_INVENTARIO_KEY, JSON.stringify(resultado));
  return resultado;
}

function obterOpcoesResponsaveisInventario() {
  const planilha = SpreadsheetApp.openById(CONFIG.PLANILHA_ID);
  const aba = planilha.getSheetByName(CONFIG.ABA_BASE_HC);
  if (!aba || aba.getLastRow() < 2) throw new Error('A aba "' + CONFIG.ABA_BASE_HC + '" não foi encontrada ou está sem dados.');
  const valores = aba.getRange(2, 1, aba.getLastRow() - 1, 15).getDisplayValues();
  const grupos = {};
  valores.forEach(function(linha) {
    const nome = String(linha[0] || '').trim();
    const setor = String(linha[8] || '').trim();
    if (setor.toUpperCase() !== 'PESSOAS') return;
    const coordenador = String(linha[14] || '').trim();
    if (!grupos[setor]) grupos[setor] = { setor: setor, nomes: [], coordenadores: [] };
    if (nome && grupos[setor].nomes.indexOf(nome) === -1) grupos[setor].nomes.push(nome);
    if (coordenador && grupos[setor].coordenadores.indexOf(coordenador) === -1) grupos[setor].coordenadores.push(coordenador);
  });
  return Object.keys(grupos).sort(function(a, b) { return a.localeCompare(b, 'pt-BR'); }).map(function(setor) {
    const grupo = grupos[setor];
    return { setor: grupo.setor, nomes: grupo.nomes.sort(function(a, b) { return a.localeCompare(b, 'pt-BR'); }), coordenadores: grupo.coordenadores.sort(function(a, b) { return a.localeCompare(b, 'pt-BR'); }) };
  }).filter(function(grupo) { return grupo.nomes.length || grupo.coordenadores.length; });
}

function obterResponsaveisInventario_() {
  const grupos = obterOpcoesResponsaveisInventario();
  const nomes = [];
  grupos.forEach(function(grupo) {
    grupo.nomes.concat(grupo.coordenadores).forEach(function(nome) { if (nomes.indexOf(nome) === -1) nomes.push(nome); });
  });
  return nomes;
}

function salvarInventario(dados) {
  dados = dados || {};
  if (!dados.responsavel) throw new Error('Selecione a pessoa responsável pelo inventário.');
  if (obterResponsaveisInventario_().indexOf(String(dados.responsavel)) === -1) throw new Error('Responsável inválido para a aba Base_HC.');
  const campos = ['yubico300', 'yubico325', 'yubico600'];
  campos.forEach(function(campo) {
    const valor = Number(dados[campo]);
    if (!Number.isFinite(valor) || valor < 0 || Math.floor(valor) !== valor) throw new Error('Informe quantidades inteiras e não negativas para todos os modelos.');
  });
  const planilha = SpreadsheetApp.openById(CONFIG.PLANILHA_ID);
  let aba = planilha.getSheetByName(CONFIG.ABA_INVENTARIO);
  if (!aba) aba = planilha.insertSheet(CONFIG.ABA_INVENTARIO);
  if (aba.getLastRow() === 0) aba.getRange(1, 1, 1, 5).setValues([['Data do inventário', 'yubico_300', 'yubico_325', 'yubico_600', 'Registrado por']]);
  aba.appendRow([new Date(), Number(dados.yubico300), Number(dados.yubico325), Number(dados.yubico600), String(dados.responsavel)]);
  PropertiesService.getScriptProperties().deleteProperty(CACHE_INVENTARIO_KEY);
  return obterDadosInventario();
}

function obterPessoasAntigasSemYubicoTV_() {
  const planilha = SpreadsheetApp.openById(CONFIG.PLANILHA_ID);
  const aba = planilha.getSheetByName(CONFIG.ABA_BASE_HC);
  if (!aba || aba.getLastRow() < 2 || aba.getLastColumn() < 49) return [];
  const quantidadeColunas = Math.max(49, aba.getLastColumn());
  const cabecalhos = aba.getRange(1, 1, 1, quantidadeColunas).getDisplayValues()[0].map(normalizarTexto_);
  const valores = aba.getRange(2, 1, aba.getLastRow() - 1, quantidadeColunas).getValues();
  const exibidos = aba.getRange(2, 1, aba.getLastRow() - 1, quantidadeColunas).getDisplayValues();
  const colunaTurno = localizarColunaTV_(cabecalhos, ['TURNO', 'JORNADA', 'SHIFT'], -1);
  const colunaFuncao = localizarColunaTV_(cabecalhos, ['FUNCAO', 'CARGO', 'FUNCAO/CARGO', 'JOB'], -1);
  const colunaNivel = localizarColunaTV_(cabecalhos, ['NIVEL', 'NIVEL DO CARGO', 'GRADE', 'LEVEL'], -1);
  const candidatos = [];
  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0);
  for (let i = 0; i < valores.length; i++) {
    const nome = String(exibidos[i][0] || '').trim();
    const codigo = String(exibidos[i][48] || '').trim();
    const admissao = parseDataAdmissaoTV_(valores[i][3], exibidos[i][3]);
    if (!nome || codigo || !admissao) continue;
    const turno = colunaTurno >= 0 ? String(exibidos[i][colunaTurno] || '').trim() : 'Não informado';
    const funcao = colunaFuncao >= 0 ? String(exibidos[i][colunaFuncao] || '').trim() : 'Não informado';
    const nivel = colunaNivel >= 0 ? String(exibidos[i][colunaNivel] || '').trim() : '';
    const funcaoNivel = nivel ? (funcao ? funcao + ' (' + nivel + ')' : nivel) : (funcao || 'Não informado');
    const diasEmpresa = Math.max(0, Math.floor((hoje.getTime() - new Date(admissao.getFullYear(), admissao.getMonth(), admissao.getDate()).getTime()) / 86400000));
    candidatos.push({ nome: nome, turno: turno, funcao: funcaoNivel, admissao: formatarDataCurtaTV_(admissao), diasEmpresa: diasEmpresa, timestamp: admissao.getTime() });
  }
  candidatos.sort(function(a, b) { return a.timestamp - b.timestamp || a.nome.localeCompare(b.nome, 'pt-BR'); });
  return candidatos.slice(0, 20).map(function(item, indice) { return { posicao: indice + 1, nome: item.nome, turno: item.turno, funcao: item.funcao, admissao: item.admissao, diasEmpresa: item.diasEmpresa }; });
}

function parseDataAdmissaoTV_(valor, exibido) {
  if (Object.prototype.toString.call(valor) === '[object Date]' && !isNaN(valor.getTime())) return valor;
  const texto = String(exibido || valor || '').trim();
  if (!texto) return null;
  const partes = texto.split(/[\/\-.]/).map(Number);
  if (partes.length === 3 && partes[0] > 1000) return new Date(partes[0], partes[1] - 1, partes[2]);
  if (partes.length === 3 && partes[2] > 1000) return new Date(partes[2], partes[1] - 1, partes[0]);
  const data = new Date(texto);
  return isNaN(data.getTime()) ? null : data;
}

function formatarDataCurtaTV_(data) {
  return Utilities.formatDate(data, Session.getScriptTimeZone() || 'America/Sao_Paulo', 'dd/MM/yyyy');
}

function obterMapaIndicadoresResumoTV_() {
  const planilha = SpreadsheetApp.openById(CONFIG.PLANILHA_ID);
  const aba = planilha.getSheetByName(CONFIG.ABA_RESUMO);
  if (!aba || aba.getLastRow() < 2) return {};
  const valores = aba.getRange(1, 1, aba.getLastRow(), aba.getLastColumn()).getDisplayValues();
  const cabecalhos = valores[0].map(normalizarTexto_);
  const colunaNome = localizarColuna_(cabecalhos, 'NOME');
  const colunaCodigo = localizarColuna_(cabecalhos, 'CODIGO YUBICO');
  const colunaForms = localizarColuna_(cabecalhos, 'POSSUI FORMS');
  const mapa = {};
  for (let i = 1; i < valores.length; i++) {
    const nome = String(valores[i][colunaNome] || '').trim();
    if (!nome) continue;
    mapa[normalizarTexto_(nome)] = { codigo: colunaCodigo >= 0 ? String(valores[i][colunaCodigo] || '').trim() : '', forms: colunaForms >= 0 ? normalizarTexto_(valores[i][colunaForms]) : '' };
  }
  return mapa;
}

function obterDadosTVSlides() {
  const planilha = SpreadsheetApp.openById(CONFIG.PLANILHA_ID);
  const aba = planilha.getSheetByName(CONFIG.ABA_BASE_HC);
  if (!aba || aba.getLastRow() < 2) throw new Error('A aba "' + CONFIG.ABA_BASE_HC + '" não foi encontrada ou está sem dados.');
  const valores = aba.getRange(1, 1, aba.getLastRow(), aba.getLastColumn()).getDisplayValues();
  const cabecalhos = valores[0].map(normalizarTexto_);
  const mapaIndicadoresResumo = obterMapaIndicadoresResumoTV_();
  const colunaNome = localizarColunaTV_(cabecalhos, ['NOME', 'COLABORADOR', 'FUNCIONARIO'], 0);
  const colunaEmpregador = localizarColunaTV_(cabecalhos, ['EMPREGADOR', 'EMPRESA', 'AGENCIA'], -1);
  const colunaDepartamento = localizarColunaTV_(cabecalhos, ['DEPARTAMENTO', 'SETOR'], -1);
  const colunaTurno = localizarColunaTV_(cabecalhos, ['TURNO', 'JORNADA'], -1);
  const colunaFuncao = localizarColunaTV_(cabecalhos, ['FUNCAO', 'CARGO', 'FUNCAO/CARGO'], -1);
  const colunaCodigo = localizarColunaTV_(cabecalhos, ['CODIGO YUBICO', 'CODIGO', 'YUBICO'], -1);
  const colunaForms = localizarColunaTV_(cabecalhos, ['POSSUI FORMS', 'POSSUI FORMULARIO', 'FORMULARIO'], -1);
  const luft = { nome: 'Luft', departamento: [], turno: [], funcao: [] };
  const agencias = { nome: 'Agências', departamento: [], turno: [], funcao: [] };
  const agenciasDetalhes = {};
  for (let i = 1; i < valores.length; i++) {
    const linha = valores[i];
    if (!String(linha[colunaNome] || '').trim()) continue;
    const empregador = colunaEmpregador >= 0 ? String(linha[colunaEmpregador] || '').trim() : '';
    const ehLuft = normalizarTexto_(empregador).indexOf('LUFT') >= 0;
    const destino = ehLuft ? luft : agencias;
    const nomeAgencia = empregador || 'Agência não informada';
    const agenciaDetalhe = ehLuft ? null : (agenciasDetalhes[nomeAgencia] || (agenciasDetalhes[nomeAgencia] = { nome: nomeAgencia, departamento: [], turno: [], funcao: [] }));
    const pessoaResumo = mapaIndicadoresResumo[normalizarTexto_(linha[colunaNome])];
    const codigo = pessoaResumo ? pessoaResumo.codigo : (colunaCodigo >= 0 ? String(linha[colunaCodigo] || '').trim() : '');
    const forms = pessoaResumo ? pessoaResumo.forms : (colunaForms >= 0 ? normalizarTexto_(linha[colunaForms]) : '');
    const dados = {
      comYubico: codigo ? 1 : 0,
      comYubicoSemFormulario: codigo && forms === 'NAO' ? 1 : 0,
      semYubicoComFormulario: !codigo && forms === 'SIM' ? 1 : 0
    };
    adicionarGrupoTV_(destino.departamento, colunaDepartamento >= 0 ? linha[colunaDepartamento] : 'Sem departamento', dados);
    adicionarGrupoTV_(destino.turno, colunaTurno >= 0 ? linha[colunaTurno] : 'Sem turno', dados);
    adicionarGrupoTV_(destino.funcao, colunaFuncao >= 0 ? linha[colunaFuncao] : 'Sem função', dados);
    if (agenciaDetalhe) {
      adicionarGrupoTV_(agenciaDetalhe.departamento, colunaDepartamento >= 0 ? linha[colunaDepartamento] : 'Sem departamento', dados);
      adicionarGrupoTV_(agenciaDetalhe.turno, colunaTurno >= 0 ? linha[colunaTurno] : 'Sem turno', dados);
      adicionarGrupoTV_(agenciaDetalhe.funcao, colunaFuncao >= 0 ? linha[colunaFuncao] : 'Sem função', dados);
    }
  }
  const principais = calcularIndicadores_();
  const inventario = obterDadosInventario();
  const informativos = obterInformativos();
  const estoque = inventario.atual || { yubico300: 0, yubico325: 0, yubico600: 0 };
  const totalYubicoEstoque = (Number(estoque.yubico300) || 0) + (Number(estoque.yubico325) || 0) + (Number(estoque.yubico600) || 0);
  return {
    atualizadoEm: formatarData_(new Date()),
    principal: {
      quantidadePessoas: principais.quantidadePessoas || 0,
      comYubico: principais.comYubico || 0,
      semYubico: Math.max(0, (principais.quantidadePessoas || 0) - (principais.comYubico || 0)),
      comYubicoSemFormulario: principais.comYubicoSemFormulario || 0,
      semYubicoComFormulario: principais.semYubicoComFormulario || 0,
      totalYubicoEstoque: totalYubicoEstoque,
      totalYubico: (principais.comYubico || 0) + totalYubicoEstoque
    },
    inventario: { atual: inventario.atual, estoque: { yubico300: Number(estoque.yubico300) || 0, yubico325: Number(estoque.yubico325) || 0, yubico600: Number(estoque.yubico600) || 0, total: totalYubicoEstoque } },
    informativos: informativos,
    luft: finalizarGruposTV_(luft),
    agencias: finalizarGruposTV_(agencias),
    agenciasDetalhes: Object.keys(agenciasDetalhes).sort(function(a, b) { return a.localeCompare(b, 'pt-BR'); }).map(function(nome) { return finalizarGruposTV_(agenciasDetalhes[nome]); }),
    pessoasAntigasSemYubico: obterPessoasAntigasSemYubicoTV_()
  };
}

function localizarColunaTV_(cabecalhos, nomes, fallback) {
  for (let i = 0; i < nomes.length; i++) {
    const indice = cabecalhos.indexOf(normalizarTexto_(nomes[i]));
    if (indice >= 0) return indice;
  }
  return fallback;
}

function adicionarGrupoTV_(grupos, nome, dados) {
  const chave = String(nome || '').trim() || 'Não informado';
  let grupo = grupos.filter(function(item) { return item.nome === chave; })[0];
  if (!grupo) { grupo = { nome: chave, quantidadePessoas: 0, comYubico: 0, comYubicoSemFormulario: 0, semYubicoComFormulario: 0 }; grupos.push(grupo); }
  grupo.quantidadePessoas++;
  grupo.comYubico += dados.comYubico;
  grupo.comYubicoSemFormulario += dados.comYubicoSemFormulario;
  grupo.semYubicoComFormulario += dados.semYubicoComFormulario;
}

function finalizarGruposTV_(empresa) {
  return { nome: empresa.nome, departamento: ordenarGruposTV_(empresa.departamento), turno: ordenarGruposTV_(empresa.turno), funcao: ordenarGruposTV_(empresa.funcao) };
}

function ordenarGruposTV_(grupos) {
  return grupos.sort(function(a, b) { return b.quantidadePessoas - a.quantidadePessoas || a.nome.localeCompare(b.nome, 'pt-BR'); });
}

function atualizarTudo() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) throw new Error('Já existe uma atualização do Drive em andamento. Aguarde a conclusão antes de tentar novamente.');
  const inicio = new Date();
  limparLogExecucao_();
  registrarEvento_('Informação', 'Iniciando atualização do relatório do Drive.');

  const processamento = {
    status: 'Processando',
    inicio: formatarData_(inicio),
    fim: '',
    duracao: '',
    arquivos: 0,
    pastas: 0,
    mensagem: 'Atualizando relatório do Drive...'
  };
  salvarUltimoProcessamento_(processamento);

  try {
    const resultadoDrive = sincronizarRelatorioDrive();
    const indicadores = calcularIndicadores_();
    const fim = new Date();

    processamento.status = 'Concluído';
    processamento.fim = formatarData_(fim);
    processamento.duracao = ((fim - inicio) / 1000).toFixed(1) + ' s';
    processamento.arquivos = resultadoDrive.arquivos;
    processamento.pastas = resultadoDrive.pastas;
    processamento.mensagem = 'Sincronização concluída: ' + (resultadoDrive.novos || 0) + ' novo(s), ' + (resultadoDrive.alterados || 0) + ' alterado(s), ' + (resultadoDrive.removidos || 0) + ' removido(s).';
    salvarUltimoProcessamento_(processamento);
    registrarEvento_('Informação', processamento.mensagem);

    return { indicadores: indicadores, processamento: processamento, log: obterLogExecucao() };
  } catch (erro) {
    const fim = new Date();
    processamento.status = 'Erro';
    processamento.fim = formatarData_(fim);
    processamento.duracao = ((fim - inicio) / 1000).toFixed(1) + ' s';
    processamento.mensagem = erro && erro.message ? erro.message : String(erro);
    salvarUltimoProcessamento_(processamento);
    registrarEvento_('Erro', processamento.mensagem);
    throw new Error(processamento.mensagem);
  } finally {
    lock.releaseLock();
  }
}

function calcularIndicadores_() {
  const planilha = SpreadsheetApp.openById(CONFIG.PLANILHA_ID);
  const aba = planilha.getSheetByName(CONFIG.ABA_RESUMO);
  if (!aba) throw new Error('A aba "' + CONFIG.ABA_RESUMO + '" não foi encontrada.');

  const ultimaLinha = aba.getLastRow();
  const ultimaColuna = aba.getLastColumn();
  if (ultimaLinha < 1 || ultimaColuna < 1) throw new Error('A aba "' + CONFIG.ABA_RESUMO + '" está vazia.');

  const valores = aba.getRange(1, 1, ultimaLinha, ultimaColuna).getDisplayValues();
  const cabecalhos = valores[0].map(normalizarTexto_);
  const colunaNome = localizarColuna_(cabecalhos, 'NOME');
  const colunaCodigo = localizarColuna_(cabecalhos, 'CODIGO YUBICO');
  const colunaForms = localizarColuna_(cabecalhos, 'POSSUI FORMS');
  const colunaEmpregador = localizarColuna_(cabecalhos, 'EMPREGADOR');

  let quantidadePessoas = 0;
  let comYubico = 0;
  let comYubicoSemFormulario = 0;
  let semYubicoComFormulario = 0;
  const gruposEmpregador = {};

  for (let i = 1; i < valores.length; i++) {
    const nome = String(valores[i][colunaNome] || '').trim();
    if (nome === '') continue;

    quantidadePessoas++;
    const codigo = String(valores[i][colunaCodigo] || '').trim();
    const possuiForms = normalizarTexto_(valores[i][colunaForms]);
    const empregador = String(valores[i][colunaEmpregador] || '').trim() || 'Sem empregador informado';
    const temYubico = codigo !== '';

    if (!gruposEmpregador[empregador]) {
      gruposEmpregador[empregador] = { empregador: empregador, quantidadePessoas: 0, comYubico: 0, comYubicoSemFormulario: 0, semYubicoComFormulario: 0 };
    }
    const grupo = gruposEmpregador[empregador];
    grupo.quantidadePessoas++;
    if (temYubico) grupo.comYubico++;
    if (temYubico && possuiForms === 'NAO') grupo.comYubicoSemFormulario++;
    if (!temYubico && possuiForms === 'SIM') grupo.semYubicoComFormulario++;

    if (temYubico) comYubico++;
    if (temYubico && possuiForms === 'NAO') comYubicoSemFormulario++;
    if (!temYubico && possuiForms === 'SIM') semYubicoComFormulario++;
  }

  const porEmpregador = Object.keys(gruposEmpregador).map(function(nome) { return gruposEmpregador[nome]; }).sort(function(a, b) { return a.empregador.localeCompare(b.empregador, 'pt-BR'); });
  const resultado = { quantidadePessoas: quantidadePessoas, comYubico: comYubico, comYubicoSemFormulario: comYubicoSemFormulario, semYubicoComFormulario: semYubicoComFormulario, porEmpregador: porEmpregador, atualizadoEm: formatarData_(new Date()) };
  PropertiesService.getScriptProperties().setProperty(CACHE_INDICADORES_KEY, JSON.stringify(resultado));
  return resultado;
}

function localizarColuna_(cabecalhos, nome) {
  const indice = cabecalhos.indexOf(normalizarTexto_(nome));
  if (indice === -1) throw new Error('A coluna "' + nome + '" não foi encontrada na aba "' + CONFIG.ABA_RESUMO + '".');
  return indice;
}

function gerarRelatorioDrive() {
  const inicio = new Date();
  const planilha = SpreadsheetApp.openById(CONFIG.PLANILHA_ID);
  const aba = planilha.getSheetByName(CONFIG.ABA_RELATORIO_DRIVE) || planilha.insertSheet(CONFIG.ABA_RELATORIO_DRIVE);
  const pastaRaiz = DriveApp.getFolderById(CONFIG.PASTA_ID);
  const registros = [];
  const fila = [{ pasta: pastaRaiz, pastas: [pastaRaiz.getName()] }];
  const pastasProcessadas = new Set();
  let maiorQuantidadePastas = 1;

  while (fila.length > 0) {
    const item = fila.pop();
    const pasta = item.pasta;
    const nomesPastas = item.pastas;
    const pastaId = pasta.getId();
    if (pastasProcessadas.has(pastaId)) continue;
    pastasProcessadas.add(pastaId);
    maiorQuantidadePastas = Math.max(maiorQuantidadePastas, nomesPastas.length);

    const arquivos = pasta.getFiles();
    while (arquivos.hasNext()) {
      const arquivo = arquivos.next();
      registros.push({
        nome: removerExtensao_(arquivo.getName()),
        pastas: nomesPastas,
        id: arquivo.getId(),
        url: arquivo.getUrl(),
        mime: arquivo.getMimeType(),
        tamanho: arquivo.getSize(),
        atualizado: arquivo.getLastUpdated()
      });
    }

    const subpastas = pasta.getFolders();
    while (subpastas.hasNext()) {
      const subpasta = subpastas.next();
      if (!pastasProcessadas.has(subpasta.getId())) {
        fila.push({ pasta: subpasta, pastas: nomesPastas.concat([subpasta.getName()]) });
      }
    }
  }

  const cabecalho = ['Nome do arquivo sem extensão'];
  for (let i = 1; i <= maiorQuantidadePastas; i++) cabecalho.push('Pasta ' + i);
  cabecalho.push('ID do arquivo', 'URL do arquivo', 'Tipo MIME', 'Tamanho (bytes)', 'Data da última atualização');

  const linhas = registros.map(function(registro) {
    const linha = [registro.nome];
    for (let i = 0; i < maiorQuantidadePastas; i++) linha.push(registro.pastas[i] || '');
    linha.push(registro.id, registro.url, registro.mime, registro.tamanho, registro.atualizado);
    return linha;
  });
  const linhasAntigas = aba.getLastRow();
  const colunasAntigas = aba.getLastColumn();
  const totalLinhasNovas = linhas.length + 1;
  const matrizRelatorio = [cabecalho].concat(linhas);

  if (linhasAntigas > totalLinhasNovas && colunasAntigas > 0) {
    aba.getRange(totalLinhasNovas + 1, 1, linhasAntigas - totalLinhasNovas, colunasAntigas).clearContent();
  }
  if (colunasAntigas > cabecalho.length && linhasAntigas > 0) {
    aba.getRange(1, cabecalho.length + 1, Math.max(linhasAntigas, totalLinhasNovas), colunasAntigas - cabecalho.length).clearContent();
  }

  aba.getRange(1, 1, totalLinhasNovas, cabecalho.length).setValues(matrizRelatorio);
  const duracao = ((new Date() - inicio) / 1000).toFixed(1);
  console.log('Concluído: ' + registros.length + ' arquivo(s), ' + pastasProcessadas.size + ' pasta(s), em ' + duracao + ' segundo(s).');

  return { arquivos: registros.length, pastas: pastasProcessadas.size };
}

function sincronizarRelatorioDrive() {
  const inicio = new Date();
  const planilha = SpreadsheetApp.openById(CONFIG.PLANILHA_ID);
  const aba = planilha.getSheetByName(CONFIG.ABA_RELATORIO_DRIVE) || planilha.insertSheet(CONFIG.ABA_RELATORIO_DRIVE);
  const pastaRaiz = DriveApp.getFolderById(CONFIG.PASTA_ID);
  const registros = [];
  const fila = [{ pasta: pastaRaiz, pastas: [pastaRaiz.getName()] }];
  const pastasProcessadas = new Set();
  let maiorQuantidadePastas = 1;

  while (fila.length > 0) {
    const item = fila.pop();
    const pasta = item.pasta;
    const nomesPastas = item.pastas;
    const pastaId = pasta.getId();
    if (pastasProcessadas.has(pastaId)) continue;
    pastasProcessadas.add(pastaId);
    maiorQuantidadePastas = Math.max(maiorQuantidadePastas, nomesPastas.length);

    const arquivos = pasta.getFiles();
    while (arquivos.hasNext()) {
      const arquivo = arquivos.next();
      registros.push({ nome: removerExtensao_(arquivo.getName()), pastas: nomesPastas, id: arquivo.getId(), url: arquivo.getUrl(), mime: arquivo.getMimeType(), tamanho: arquivo.getSize(), atualizado: arquivo.getLastUpdated() });
    }

    const subpastas = pasta.getFolders();
    while (subpastas.hasNext()) {
      const subpasta = subpastas.next();
      if (!pastasProcessadas.has(subpasta.getId())) fila.push({ pasta: subpasta, pastas: nomesPastas.concat([subpasta.getName()]) });
    }
  }

  const cabecalho = ['Nome do arquivo sem extensão'];
  for (let i = 1; i <= maiorQuantidadePastas; i++) cabecalho.push('Pasta ' + i);
  cabecalho.push('ID do arquivo', 'URL do arquivo', 'Tipo MIME', 'Tamanho (bytes)', 'Data da última atualização');
  const quantidadeColunas = cabecalho.length;
  const ultimaLinha = aba.getLastRow();
  const dadosAtuais = ultimaLinha > 0 ? aba.getRange(1, 1, ultimaLinha, Math.max(aba.getLastColumn(), quantidadeColunas)).getValues() : [];
  const profundidadeAntiga = dadosAtuais.length > 0 ? Math.max(0, dadosAtuais[0].length - 6) : 0;
  if (dadosAtuais.length === 0 || profundidadeAntiga !== maiorQuantidadePastas) {
    const matriz = [cabecalho].concat(registros.map(function(registro) { return transformarRegistro_(registro, maiorQuantidadePastas); }));
    aba.getRange(1, 1, matriz.length, quantidadeColunas).setValues(matriz);
    return { arquivos: registros.length, pastas: pastasProcessadas.size, novos: registros.length, alterados: 0, removidos: 0, duracao: ((new Date() - inicio) / 1000).toFixed(1) + ' s' };
  }

  const colunaId = maiorQuantidadePastas + 1;
  const mapaAtual = new Map();
  for (let linha = 1; linha < dadosAtuais.length; linha++) {
    const id = String(dadosAtuais[linha][colunaId] || '').trim();
    if (id) mapaAtual.set(id, { numeroLinha: linha + 1, valores: dadosAtuais[linha] });
  }

  const idsEncontrados = new Set();
  const novasLinhas = [];
  let novos = 0;
  let alterados = 0;
  const linhasAlteradas = [];
  registros.forEach(function(registro) {
    idsEncontrados.add(registro.id);
    const linhaNova = transformarRegistro_(registro, maiorQuantidadePastas);
    const existente = mapaAtual.get(registro.id);
    if (!existente) {
      novasLinhas.push(linhaNova);
      novos++;
    } else if (!linhasIguais_(existente.valores, linhaNova)) {
      linhasAlteradas.push({ numeroLinha: existente.numeroLinha, valores: linhaNova });
      alterados++;
    }
  });

  escreverLinhasAgrupadas_(aba, linhasAlteradas, quantidadeColunas);
  if (novasLinhas.length > 0) aba.getRange(aba.getLastRow() + 1, 1, novasLinhas.length, quantidadeColunas).setValues(novasLinhas);

  const removidos = [];
  mapaAtual.forEach(function(existente, id) { if (!idsEncontrados.has(id)) removidos.push(existente.numeroLinha); });
  limparLinhasAgrupadas_(aba, removidos, quantidadeColunas);

  return { arquivos: registros.length, pastas: pastasProcessadas.size, novos: novos, alterados: alterados, removidos: removidos.length, duracao: ((new Date() - inicio) / 1000).toFixed(1) + ' s' };
}

function escreverLinhasAgrupadas_(aba, linhas, quantidadeColunas) {
  if (!linhas.length) return;
  linhas.sort(function(a, b) { return a.numeroLinha - b.numeroLinha; });
  let inicio = linhas[0].numeroLinha;
  let bloco = [linhas[0].valores];
  for (let i = 1; i < linhas.length; i++) {
    if (linhas[i].numeroLinha === inicio + bloco.length) {
      bloco.push(linhas[i].valores);
    } else {
      aba.getRange(inicio, 1, bloco.length, quantidadeColunas).setValues(bloco);
      inicio = linhas[i].numeroLinha;
      bloco = [linhas[i].valores];
    }
  }
  aba.getRange(inicio, 1, bloco.length, quantidadeColunas).setValues(bloco);
}

function limparLinhasAgrupadas_(aba, linhas, quantidadeColunas) {
  if (!linhas.length) return;
  linhas.sort(function(a, b) { return a - b; });
  let inicio = linhas[0];
  let quantidade = 1;
  for (let i = 1; i < linhas.length; i++) {
    if (linhas[i] === inicio + quantidade) {
      quantidade++;
    } else {
      aba.getRange(inicio, 1, quantidade, quantidadeColunas).clearContent();
      inicio = linhas[i];
      quantidade = 1;
    }
  }
  aba.getRange(inicio, 1, quantidade, quantidadeColunas).clearContent();
}

function transformarRegistro_(registro, quantidadePastas) {
  const linha = [registro.nome];
  for (let i = 0; i < quantidadePastas; i++) linha.push(registro.pastas[i] || '');
  linha.push(registro.id, registro.url, registro.mime, registro.tamanho, registro.atualizado);
  return linha;
}

function linhasIguais_(a, b) {
  if (!a || a.length < b.length) return false;
  for (let i = 0; i < b.length; i++) {
    const valorA = a[i] instanceof Date ? a[i].getTime() : String(a[i] == null ? '' : a[i]);
    const valorB = b[i] instanceof Date ? b[i].getTime() : String(b[i] == null ? '' : b[i]);
    if (valorA !== valorB) return false;
  }
  return true;
}

function removerExtensao_(nomeArquivo) {
  const ultimoPonto = nomeArquivo.lastIndexOf('.');
  return ultimoPonto <= 0 ? nomeArquivo : nomeArquivo.substring(0, ultimoPonto);
}

function normalizarTexto_(valor) {
  return String(valor || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toUpperCase();
}

function formatarData_(data) {
  return Utilities.formatDate(data, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
}

function salvarUltimoProcessamento_(dados) {
  PropertiesService.getDocumentProperties().setProperty('PAINEL_PROCESSAMENTO', JSON.stringify(dados));
}

function obterUltimoProcessamento_() {
  const salvo = PropertiesService.getDocumentProperties().getProperty('PAINEL_PROCESSAMENTO');
  return salvo ? JSON.parse(salvo) : { status: 'Aguardando', inicio: '', fim: '', duracao: '', arquivos: 0, pastas: 0, mensagem: 'Nenhuma atualização foi executada nesta sessão.' };
}

function limparLogExecucao_() {
  CacheService.getScriptCache().put(CACHE_LOG_KEY, JSON.stringify({ status: 'Processando', eventos: [] }), 21600);
}

function registrarEvento_(nivel, mensagem) {
  const cache = CacheService.getScriptCache();
  const salvo = cache.get(CACHE_LOG_KEY);
  const log = salvo ? JSON.parse(salvo) : { status: 'Processando', eventos: [] };
  log.status = nivel === 'Erro' ? 'Erro' : 'Processando';
  log.eventos.push({ horario: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm:ss'), nivel: nivel, mensagem: String(mensagem || '') });
  cache.put(CACHE_LOG_KEY, JSON.stringify(log), 21600);
}

function obterLogExecucao() {
  const salvo = CacheService.getScriptCache().get(CACHE_LOG_KEY);
  return salvo ? JSON.parse(salvo) : { status: 'Aguardando', eventos: [] };
}

function criarGatilhoDiario() {
  ScriptApp.getProjectTriggers()
    .filter(function(trigger) { return trigger.getHandlerFunction() === 'gerarRelatorioDrive'; })
    .forEach(function(trigger) { ScriptApp.deleteTrigger(trigger); });

  ScriptApp.newTrigger('gerarRelatorioDrive')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .create();
}

const EXPORTACOES_CONFIG = {
  ABA: 'Resumo_HC',
  DEPARTAMENTO: 1,
  TURNO: 7,
  EMPREGADOR: 9,
  YUBICO: 10,
  FORMULARIO: 11,
  CONFERE: 14
};

function obterOpcoesExportacaoResumoHC() {
  const dados = lerResumoHCExportacao_();
  return {
    totalRegistros: dados.linhas.length,
    empregadores: listarOpcoesExportacao_(dados.linhas, EXPORTACOES_CONFIG.EMPREGADOR),
    departamentos: listarOpcoesExportacao_(dados.linhas, EXPORTACOES_CONFIG.DEPARTAMENTO),
    turnos: listarOpcoesExportacao_(dados.linhas, EXPORTACOES_CONFIG.TURNO),
    yubicos: [{ valor: 'SIM', rotulo: 'Sim' }, { valor: 'NAO', rotulo: 'Não' }],
    formularios: [{ valor: 'SIM', rotulo: 'Sim' }, { valor: 'NAO', rotulo: 'Não' }],
    empregadoresConferem: [{ valor: 'SIM', rotulo: 'Sim' }, { valor: 'NAO', rotulo: 'Não' }]
  };
}

function contarResumoHCExportacao(filtros) {
  return filtrarResumoHCExportacao_(filtros, lerResumoHCExportacao_().linhas).length;
}

function exportarResumoHC(filtros, formato) {
  const tipo = String(formato || '').toLowerCase();
  if (tipo !== 'csv' && tipo !== 'xlsx') throw new Error('Formato de exportação inválido.');
  const dados = lerResumoHCExportacao_();
  const linhas = filtrarResumoHCExportacao_(filtros, dados.linhas);
  const matriz = [dados.cabecalhos].concat(linhas);
  const carimbo = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'America/Sao_Paulo', 'yyyyMMdd_HHmmss');
  const nome = 'Resumo_HC_' + carimbo + (tipo === 'csv' ? '.csv' : '.xlsx');
  const blob = tipo === 'csv' ? criarCsvResumoHC_(matriz, nome) : criarExcelResumoHC_(matriz, nome);
  return {
    nome: blob.getName(),
    mimeType: blob.getContentType(),
    base64: Utilities.base64Encode(blob.getBytes()),
    quantidadeRegistros: linhas.length
  };
}

function lerResumoHCExportacao_() {
  const planilha = SpreadsheetApp.openById(CONFIG.PLANILHA_ID);
  const aba = planilha.getSheetByName(EXPORTACOES_CONFIG.ABA);
  if (!aba) throw new Error('A aba Resumo_HC não foi encontrada.');
  const ultimaLinha = aba.getLastRow();
  const ultimaColuna = aba.getLastColumn();
  if (ultimaLinha < 1 || ultimaColuna < 1) throw new Error('A aba Resumo_HC está vazia.');
  const valores = aba.getRange(1, 1, ultimaLinha, ultimaColuna).getDisplayValues();
  const cabecalhos = valores[0].map(function(valor, indice) {
    return String(valor || '').trim() || ('Coluna ' + (indice + 1));
  });
  const linhas = valores.slice(1).filter(function(linha) {
    return linha.some(function(valor) { return String(valor || '').trim() !== ''; });
  });
  return { cabecalhos: cabecalhos, linhas: linhas };
}

function listarOpcoesExportacao_(linhas, indice) {
  const mapa = {};
  linhas.forEach(function(linha) {
    const valor = String(linha[indice] || '').trim() || 'Não informado';
    mapa[normalizarExportacao_(valor)] = valor;
  });
  return Object.keys(mapa).map(function(chave) {
    return { valor: mapa[chave], rotulo: mapa[chave] };
  }).sort(function(a, b) {
    return a.rotulo.localeCompare(b.rotulo, 'pt-BR');
  });
}

function filtrarResumoHCExportacao_(filtros, linhas) {
  filtros = filtros || {};
  const selecionados = {
    empregadores: normalizarSelecaoExportacao_(filtros.empregadores),
    departamentos: normalizarSelecaoExportacao_(filtros.departamentos),
    turnos: normalizarSelecaoExportacao_(filtros.turnos),
    yubicos: normalizarSelecaoExportacao_(filtros.yubicos),
    formularios: normalizarSelecaoExportacao_(filtros.formularios),
    empregadoresConferem: normalizarSelecaoExportacao_(filtros.empregadoresConferem)
  };
  return linhas.filter(function(linha) {
    return valorCorrespondeExportacao_(linha[EXPORTACOES_CONFIG.EMPREGADOR], selecionados.empregadores, '') &&
      valorCorrespondeExportacao_(linha[EXPORTACOES_CONFIG.DEPARTAMENTO], selecionados.departamentos, '') &&
      valorCorrespondeExportacao_(linha[EXPORTACOES_CONFIG.TURNO], selecionados.turnos, '') &&
      valorCorrespondeExportacao_(linha[EXPORTACOES_CONFIG.YUBICO], selecionados.yubicos, 'YUBICO') &&
      valorCorrespondeExportacao_(linha[EXPORTACOES_CONFIG.FORMULARIO], selecionados.formularios, 'SIMNAO') &&
      valorCorrespondeExportacao_(linha[EXPORTACOES_CONFIG.CONFERE], selecionados.empregadoresConferem, 'SIMNAO');
  });
}

function normalizarSelecaoExportacao_(lista) {
  if (!Array.isArray(lista)) return [];
  return lista.map(normalizarExportacao_).filter(function(valor, indice, valores) {
    return valor && valores.indexOf(valor) === indice;
  });
}

function valorCorrespondeExportacao_(valor, lista, tipo) {
  if (!lista.length) return true;
  let atual = '';
  if (tipo === 'YUBICO') atual = String(valor || '').trim() ? 'SIM' : 'NAO';
  else if (tipo === 'SIMNAO') atual = normalizarExportacao_(valor) === 'SIM' ? 'SIM' : 'NAO';
  else atual = normalizarExportacao_(String(valor || '').trim() || 'Não informado');
  return lista.indexOf(atual) !== -1;
}

function normalizarExportacao_(valor) {
  return String(valor || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toUpperCase();
}

function criarCsvResumoHC_(matriz, nome) {
  const conteudo = '\uFEFF' + matriz.map(function(linha) {
    return linha.map(function(valor) {
      return '"' + String(valor == null ? '' : valor).replace(/"/g, '""') + '"';
    }).join(';');
  }).join('\r\n');
  return Utilities.newBlob(conteudo, 'text/csv;charset=utf-8', nome);
}

function criarExcelResumoHC_(matriz, nome) {
  const totalLinhas = Math.max(1, matriz.length);
  const totalColunas = Math.max(1, matriz.reduce(function(maior, linha) { return Math.max(maior, linha.length); }, 0));
  const dimensao = 'A1:' + colunaExcel_(totalColunas) + totalLinhas;
  const linhas = matriz.map(function(linha, indiceLinha) {
    const celulas = [];
    for (let indiceColuna = 0; indiceColuna < totalColunas; indiceColuna++) {
      const referencia = colunaExcel_(indiceColuna + 1) + (indiceLinha + 1);
      const estilo = indiceLinha === 0 ? ' s="1"' : '';
      celulas.push('<c r="' + referencia + '" t="inlineStr"' + estilo + '><is><t xml:space="preserve">' + xmlExcel_(linha[indiceColuna]) + '</t></is></c>');
    }
    return '<row r="' + (indiceLinha + 1) + '">' + celulas.join('') + '</row>';
  }).join('');
  const larguras = [];
  for (let indiceColuna = 0; indiceColuna < totalColunas; indiceColuna++) {
    const tamanho = matriz.reduce(function(maior, linha) { return Math.max(maior, String(linha[indiceColuna] == null ? '' : linha[indiceColuna]).length); }, 8);
    larguras.push('<col min="' + (indiceColuna + 1) + '" max="' + (indiceColuna + 1) + '" width="' + Math.min(45, Math.max(10, tamanho + 2)) + '" customWidth="1"/>');
  }
  const arquivos = [
    Utilities.newBlob('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>', 'application/xml', '[Content_Types].xml'),
    Utilities.newBlob('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>', 'application/xml', '_rels/.rels'),
    Utilities.newBlob('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Resumo HC" sheetId="1" r:id="rId1"/></sheets></workbook>', 'application/xml', 'xl/workbook.xml'),
    Utilities.newBlob('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>', 'application/xml', 'xl/_rels/workbook.xml.rels'),
    Utilities.newBlob('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="10"/><name val="Arial"/></font><font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Arial"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF155EEF"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="1" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs></styleSheet>', 'application/xml', 'xl/styles.xml'),
    Utilities.newBlob('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="' + dimensao + '"/><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols>' + larguras.join('') + '</cols><sheetData>' + linhas + '</sheetData><autoFilter ref="' + dimensao + '"/></worksheet>', 'application/xml', 'xl/worksheets/sheet1.xml')
  ];
  return Utilities.zip(arquivos, nome).setContentType('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
}

function colunaExcel_(numero) {
  let resultado = '';
  while (numero > 0) {
    const resto = (numero - 1) % 26;
    resultado = String.fromCharCode(65 + resto) + resultado;
    numero = Math.floor((numero - 1) / 26);
  }
  return resultado;
}

function xmlExcel_(valor) {
  return String(valor == null ? '' : valor)
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

const EMPRESTIMOS_CONFIG = {
  ABA: 'Emprestimos',
  PASTA: 'Emprestimos'
};

function obterPessoasEmprestimos() {
  const planilha = SpreadsheetApp.openById(CONFIG.PLANILHA_ID);
  const aba = planilha.getSheetByName(CONFIG.ABA_RESUMO);
  if (!aba || aba.getLastRow() < 2) return [];
  const valores = aba.getRange(2, 1, aba.getLastRow() - 1, Math.max(10, aba.getLastColumn())).getDisplayValues();
  const mapa = {};
  valores.forEach(function(linha) {
    const cpf = String(linha[2] || '').trim();
    const chave = somenteDigitosEmprestimos_(cpf);
    if (!chave || mapa[chave]) return;
    mapa[chave] = {
      cpf: cpf,
      cpfChave: chave,
      nome: String(linha[0] || '').trim(),
      departamento: String(linha[1] || '').trim(),
      funcao: String(linha[4] || '').trim(),
      nivel: String(linha[5] || '').trim(),
      area: String(linha[6] || '').trim(),
      turno: String(linha[7] || '').trim(),
      setor: String(linha[8] || '').trim(),
      empregador: String(linha[9] || '').trim()
    };
  });
  return Object.keys(mapa).map(function(chave) { return mapa[chave]; }).sort(function(a, b) {
    return a.cpf.localeCompare(b.cpf, 'pt-BR');
  });
}

function obterPessoaEmprestimosPorCpf(cpf) {
  const chave = somenteDigitosEmprestimos_(cpf);
  if (!chave) return null;
  const planilha = SpreadsheetApp.openById(CONFIG.PLANILHA_ID);
  const aba = planilha.getSheetByName(CONFIG.ABA_RESUMO);
  if (!aba || aba.getLastRow() < 2) return null;
  const valores = aba.getRange(2, 1, aba.getLastRow() - 1, Math.max(10, aba.getLastColumn())).getDisplayValues();
  for (let indice = 0; indice < valores.length; indice++) {
    const linha = valores[indice];
    if (somenteDigitosEmprestimos_(linha[2]) !== chave) continue;
    return {
      cpf: String(linha[2] || '').trim(),
      cpfChave: chave,
      nome: String(linha[0] || '').trim(),
      departamento: String(linha[1] || '').trim(),
      funcao: String(linha[4] || '').trim(),
      nivel: String(linha[5] || '').trim(),
      area: String(linha[6] || '').trim(),
      turno: String(linha[7] || '').trim(),
      setor: String(linha[8] || '').trim(),
      empregador: String(linha[9] || '').trim()
    };
  }
  return null;
}

function registrarEmprestimo(dados) {
  dados = dados || {};
  const pessoa = obterPessoaEmprestimosPorCpf(dados.cpf);
  if (!pessoa) throw new Error('CPF não encontrado na aba Resumo_HC.');
  if (!dados.pastaId) throw new Error('Selecione a pasta onde o formulário será salvo.');
  if (!dados.nomeArquivo || !dados.arquivoBase64) throw new Error('Anexe o formulário de empréstimo.');
  const aparelho = String(dados.aparelho || '');
  if (['Yubico 300', 'Yubico 325', 'Yubico 600'].indexOf(aparelho) === -1) throw new Error('Selecione um modelo de Yubico válido.');
  const formulario = salvarFormularioEmprestimo_(dados.pastaId, dados.nomeArquivo, dados.tipoArquivo, dados.arquivoBase64, pessoa.cpf);
  const aba = obterAbaEmprestimos_();
  const id = Utilities.getUuid();
  aba.appendRow([
    id,
    pessoa.cpf,
    pessoa.nome,
    pessoa.departamento,
    pessoa.funcao,
    pessoa.nivel,
    pessoa.area,
    pessoa.turno,
    pessoa.setor,
    pessoa.empregador,
    aparelho,
    new Date(),
    formulario.nome,
    formulario.url,
    'PENDENTE',
    ''
  ]);
  return { id: id, status: 'PENDENTE', formularioUrl: formulario.url };
}

function obterEmprestimos() {
  const aba = obterAbaEmprestimos_();
  if (aba.getLastRow() < 2) return [];
  const intervalo = aba.getRange(2, 1, aba.getLastRow() - 1, 16);
  const valores = intervalo.getValues();
  const exibicao = intervalo.getDisplayValues();
  const hoje = new Date();
  const inicioHoje = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate()).getTime();
  return valores.map(function(linha, indice) {
    const texto = exibicao[indice];
    if (!String(texto[0] || '').trim()) return null;
    const status = String(texto[14] || '').trim();
    const dataEmprestimo = linha[11] instanceof Date ? linha[11] : null;
    let diasPendentes = null;
    if (status === 'PENDENTE' && dataEmprestimo && !isNaN(dataEmprestimo.getTime())) {
      const inicioEmprestimo = new Date(dataEmprestimo.getFullYear(), dataEmprestimo.getMonth(), dataEmprestimo.getDate()).getTime();
      diasPendentes = Math.max(0, Math.floor((inicioHoje - inicioEmprestimo) / 86400000));
    }
    return {
      id: String(texto[0] || ''),
      cpf: String(texto[1] || ''),
      nome: String(texto[2] || ''),
      departamento: String(texto[3] || ''),
      aparelho: String(texto[10] || ''),
      dataEmprestimo: String(texto[11] || ''),
      formularioNome: String(texto[12] || ''),
      formularioUrl: String(texto[13] || ''),
      status: status,
      dataDevolucao: String(texto[15] || ''),
      diasPendentes: diasPendentes,
      ordemData: dataEmprestimo ? dataEmprestimo.getTime() : 0
    };
  }).filter(function(registro) { return registro; }).sort(function(a, b) {
    const ordemA = a.status === 'PENDENTE' ? 0 : 1;
    const ordemB = b.status === 'PENDENTE' ? 0 : 1;
    return ordemA - ordemB || b.ordemData - a.ordemData;
  });
}

function registrarDevolucaoEmprestimo(id, excluirFormulario) {
  const chave = String(id || '').trim();
  if (!chave) throw new Error('Empréstimo inválido.');
  const aba = obterAbaEmprestimos_();
  if (aba.getLastRow() < 2) throw new Error('Nenhum empréstimo encontrado.');
  const ids = aba.getRange(2, 1, aba.getLastRow() - 1, 1).getDisplayValues();
  for (let indice = 0; indice < ids.length; indice++) {
    if (String(ids[indice][0] || '') !== chave) continue;
    const linha = indice + 2;
    let formularioExcluido = false;
    if (excluirFormulario) {
      const url = String(aba.getRange(linha, 14).getDisplayValue() || '').trim();
      if (url) {
        const arquivoId = extrairIdArquivoDrive_(url);
        if (!arquivoId) throw new Error('Não foi possível identificar o arquivo do formulário para exclusão.');
        DriveApp.getFileById(arquivoId).setTrashed(true);
        formularioExcluido = true;
        aba.getRange(linha, 13).setValue('Formulário excluído na devolução');
        aba.getRange(linha, 14).setValue('');
      }
    }
    aba.getRange(linha, 15).setValue('DEVOLVIDO');
    aba.getRange(linha, 16).setValue(new Date());
    return { id: chave, status: 'DEVOLVIDO', formularioExcluido: formularioExcluido };
  }
  throw new Error('Empréstimo não encontrado.');
}

function extrairIdArquivoDrive_(url) {
  const encontrado = String(url || '').match(/[a-zA-Z0-9_-]{20,}/);
  return encontrado ? encontrado[0] : '';
}

function obterAbaEmprestimos_() {
  const planilha = SpreadsheetApp.openById(CONFIG.PLANILHA_ID);
  let aba = planilha.getSheetByName(EMPRESTIMOS_CONFIG.ABA);
  if (!aba) {
    aba = planilha.insertSheet(EMPRESTIMOS_CONFIG.ABA);
    aba.getRange(1, 1, 1, 16).setValues([[
      'ID', 'CPF', 'Nome', 'Departamento', 'Função', 'Nível', 'Área', 'Turno', 'Setor', 'Empregador',
      'Aparelho / Patrimônio', 'Data do empréstimo', 'Formulário', 'URL do formulário', 'Status', 'Data da devolução'
    ]]);
    aba.setFrozenRows(1);
  }
  return aba;
}

function localizarPessoaEmprestimos_(cpf) {
  const chave = somenteDigitosEmprestimos_(cpf);
  return obterPessoasEmprestimos().filter(function(pessoa) { return pessoa.cpfChave === chave; })[0] || null;
}

function salvarFormularioEmprestimo_(pastaId, nomeArquivo, tipoArquivo, arquivoBase64, cpf) {
  const pasta = DriveApp.getFolderById(String(pastaId));
  const bytes = Utilities.base64Decode(String(arquivoBase64));
  const nome = String(nomeArquivo || 'formulario').replace(/[\/:*?"<>|]/g, '_');
  const tipo = String(tipoArquivo || 'application/octet-stream');
  const arquivo = pasta.createFile(Utilities.newBlob(bytes, tipo, somenteDigitosEmprestimos_(cpf) + '_' + new Date().getTime() + '_' + nome));
  return { nome: arquivo.getName(), url: arquivo.getUrl() };
}

function somenteDigitosEmprestimos_(valor) {
  return String(valor || '').replace(/\D/g, '');
}

const INFORMATIVOS_CONFIG = {
  ABA: 'Informativos',
  GERAIS: ['hcAtual', 'valorDisponivelCompra', 'estoqueAtual', 'aparelhosDanificados', 'aparelhosAmazon', 'aparelhosCnf5'],
  POR_EMPREGADOR: ['pendenteDevolucao', 'semSaldoDesconto', 'descontoParcial']
};

function obterInformativos() {
  const aba = obterAbaInformativos_();
  const resultado = {
    gerais: { hcAtual: '', valorDisponivelCompra: '', estoqueAtual: '', aparelhosDanificados: '', aparelhosAmazon: '', aparelhosCnf5: '' },
    porEmpregador: { pendenteDevolucao: {}, semSaldoDesconto: {}, descontoParcial: {} },
    empregadores: obterEmpregadoresInformativos_(),
    atualizadoEm: ''
  };
  if (aba.getLastRow() < 2) return resultado;
  const intervalo = aba.getRange(2, 1, aba.getLastRow() - 1, 5);
  const valores = intervalo.getDisplayValues();
  const brutos = intervalo.getValues();
  let ultimaAtualizacao = 0;
  valores.forEach(function(linha, indice) {
    const tipo = String(linha[0] || '');
    const chave = String(linha[1] || '');
    const empregador = String(linha[2] || '');
    const valor = String(linha[3] || '');
    const data = dataInformativo_(brutos[indice][4], linha[4]);
    if (tipo === 'GERAL' && INFORMATIVOS_CONFIG.GERAIS.indexOf(chave) !== -1) resultado.gerais[chave] = valor;
    if (tipo === 'EMPREGADOR' && INFORMATIVOS_CONFIG.POR_EMPREGADOR.indexOf(chave) !== -1 && empregador) resultado.porEmpregador[chave][empregador] = valor;
    if (data && data.getTime() > ultimaAtualizacao) ultimaAtualizacao = data.getTime();
  });
  if (ultimaAtualizacao) resultado.atualizadoEm = Utilities.formatDate(new Date(ultimaAtualizacao), Session.getScriptTimeZone() || 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm');
  const mapaEmpregadores = {};
  resultado.empregadores.forEach(function(empregador) { mapaEmpregadores[empregador] = true; });
  INFORMATIVOS_CONFIG.POR_EMPREGADOR.forEach(function(chave) {
    Object.keys(resultado.porEmpregador[chave]).forEach(function(empregador) { mapaEmpregadores[empregador] = true; });
  });
  resultado.empregadores = Object.keys(mapaEmpregadores).sort(function(a, b) { return a.localeCompare(b, 'pt-BR'); });
  return resultado;
}

function salvarInformativos(dados) {
  dados = dados || {};
  const gerais = dados.gerais || {};
  const porEmpregador = dados.porEmpregador || {};
  const agora = new Date();
  const linhas = [['Tipo', 'Chave', 'Empregador', 'Valor', 'Atualizado em']];
  INFORMATIVOS_CONFIG.GERAIS.forEach(function(chave) {
    linhas.push(['GERAL', chave, '', valorInformativo_(gerais[chave]), agora]);
  });
  const empregadores = Array.isArray(dados.empregadores) ? dados.empregadores.map(function(valor) { return String(valor || '').trim(); }).filter(Boolean) : [];
  INFORMATIVOS_CONFIG.POR_EMPREGADOR.forEach(function(chave) {
    const valores = porEmpregador[chave] || {};
    empregadores.forEach(function(empregador) {
      linhas.push(['EMPREGADOR', chave, empregador, valorInformativo_(valores[empregador]), agora]);
    });
  });
  const aba = obterAbaInformativos_();
  aba.clearContents();
  aba.getRange(1, 1, linhas.length, 5).setValues(linhas);
  aba.setFrozenRows(1);
  return obterInformativos();
}

function dataInformativo_(bruto, exibido) {
  if (Object.prototype.toString.call(bruto) === '[object Date]' && !isNaN(bruto.getTime())) return bruto;
  const texto = String(exibido || bruto || '').trim();
  const partes = texto.match(/^(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2})$/);
  if (partes) return new Date(Number(partes[3]), Number(partes[2]) - 1, Number(partes[1]), Number(partes[4]), Number(partes[5]));
  const data = new Date(texto);
  return isNaN(data.getTime()) ? null : data;
}

function obterAbaInformativos_() {
  const planilha = SpreadsheetApp.openById(CONFIG.PLANILHA_ID);
  let aba = planilha.getSheetByName(INFORMATIVOS_CONFIG.ABA);
  if (!aba) aba = planilha.insertSheet(INFORMATIVOS_CONFIG.ABA);
  if (aba.getLastRow() === 0) {
    aba.getRange(1, 1, 1, 5).setValues([['Tipo', 'Chave', 'Empregador', 'Valor', 'Atualizado em']]);
    aba.setFrozenRows(1);
  }
  return aba;
}

function obterEmpregadoresInformativos_() {
  const planilha = SpreadsheetApp.openById(CONFIG.PLANILHA_ID);
  const aba = planilha.getSheetByName(CONFIG.ABA_RESUMO);
  if (!aba || aba.getLastRow() < 2) return [];
  const valores = aba.getRange(2, 10, aba.getLastRow() - 1, 1).getDisplayValues();
  const mapa = {};
  valores.forEach(function(linha) {
    const empregador = String(linha[0] || '').trim();
    if (empregador) mapa[empregador] = true;
  });
  return Object.keys(mapa).sort(function(a, b) { return a.localeCompare(b, 'pt-BR'); });
}

function valorInformativo_(valor) {
  const texto = String(valor == null ? '' : valor).trim();
  if (!texto) return '0';
  const numero = Number(texto.replace(',', '.'));
  if (!Number.isFinite(numero) || numero < 0) throw new Error('Informe apenas valores numéricos não negativos nos informativos.');
  return texto;
}
