/**
 * Configuração central do projeto.
 * Concentra cor de marcação, limiar de similaridade de nomes, marcador de
 * evolução sem guia e nomes das colunas obrigatórias, de forma que adaptar a
 * automação a outra planilha exige alterar um único objeto.
 */
const CONFIG = {
  COR_PINTADA: '#d9ead3',
  SIMILARIDADE_MINIMA: 0.85,
  MARCADOR_SEM_GUIA: 'TERAPIA A MAIS',
  COLUNAS: {
    PACIENTE: 'PACIENTE',
    TERAPIA: 'TERAPIA',
    DATA: 'DATA DE ATENDIMENTO'
  }
};

// ======================= MENU =======================

/**
 * Trigger de abertura da planilha. Registra o menu "Automação de Datas" com
 * as ações principais (simulação, aplicação) e as ferramentas independentes
 * (reordenar, padronizar quebras de página, formatar documento).
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🔧 Automação de Datas')
    .addItem('🧪 Rodar Teste (Dry Run)', 'rodarTeste')
    .addItem('✅ Aplicar Automação', 'rodarAplicar')
    .addSeparator()
    .addItem('🔄 Reordenar evoluções do Doc', 'rodarSomenteReordenacao')
    .addItem('📄 Padronizar quebras de página', 'rodarPadronizarQuebrasDePagina')
    .addItem('📝 Formatar documento', 'rodarFormatarDocumento')
    .addToUi();
}

/**
 * Executa o fluxo principal em modo de simulação (dry run): calcula e exibe o
 * resumo do que seria alterado sem gravar nada no Doc nem na planilha.
 */
function rodarTeste() {
  executarAutomacao(true);
}

/**
 * Executa o fluxo principal em modo real, após confirmação explícita do
 * usuário, tornando a operação destrutiva sempre consciente.
 */
function rodarAplicar() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    'Aplicar automação',
    'As datas do Doc serão corrigidas e as linhas correspondentes, marcadas na planilha. Continuar?',
    ui.ButtonSet.YES_NO
  );
  if (resp === ui.Button.YES) {
    executarAutomacao(false);
  }
}

// ======================= FLUXO PRINCIPAL =======================

/**
 * Orquestra a conciliação entre as evoluções de um Google Doc e as guias da
 * planilha de pacientes.
 *
 * O paciente é identificado automaticamente pelo título do Doc (com tolerância
 * a erros de digitação). Para cada terapia, as evoluções são pareadas com as
 * guias em ordem cronológica por meio de cursores independentes, o que mantém
 * a conciliação em tempo linear. Linhas já processadas em execuções anteriores
 * são reconhecidas por um mapa de datas, tornando a rotina idempotente: rodar
 * duas vezes no mesmo Doc não reescreve evoluções que já estavam corretas,
 * mesmo quando uma guia nova entra no meio do histórico.
 *
 * Em dry run apenas exibe o resumo. Em modo real corrige as datas no Doc,
 * pinta as linhas consumidas na planilha em uma única chamada em lote e
 * encadeia a reordenação e a formatação do documento.
 */
function executarAutomacao(dryRun) {
  const ui = SpreadsheetApp.getUi();

  const inputResp = ui.prompt(
    dryRun ? 'Teste (Dry Run)' : 'Aplicar automação',
    'Link do Google Doc do paciente:',
    ui.ButtonSet.OK_CANCEL
  );
  if (inputResp.getSelectedButton() !== ui.Button.OK) return;

  const docUrl = inputResp.getResponseText().trim();
  if (!docUrl) {
    ui.alert('Nenhum link informado.');
    return;
  }

  let doc;
  try {
    doc = DocumentApp.openById(extrairIdDoDoc(docUrl));
  } catch (e) {
    ui.alert('Não foi possível abrir o Doc. Verifique o link.\n\n' + e.message);
    return;
  }

  const sheet = SpreadsheetApp.getActiveSheet();
  const headerMap = obterMapaDeColunas(sheet);
  if (!headerMap) {
    ui.alert('Colunas obrigatórias não encontradas: PACIENTE, TERAPIA, DATA DE ATENDIMENTO.');
    return;
  }

  const tituloDoc = doc.getName();
  const resultado = obterLinhasDoPaciente(sheet, headerMap, tituloDoc);

  if (!resultado.ok) {
    ui.alert('Paciente não identificado', resultado.motivo, ui.ButtonSet.OK);
    return;
  }

  const linhasDoPaciente = resultado.linhas;
  if (linhasDoPaciente.length === 0) {
    ui.alert('Nenhum registro encontrado para esse paciente.');
    return;
  }

  const porProcedimento = {};
  linhasDoPaciente.forEach(r => {
    if (!porProcedimento[r.terapia]) porProcedimento[r.terapia] = [];
    porProcedimento[r.terapia].push(r);
  });

  const cursorPorProcedimento = {};
  const jaCorretaPorProcedimento = {};
  Object.keys(porProcedimento).forEach(proc => {
    const candidatos = porProcedimento[proc];
    const mapa = new Map();
    for (let i = 0; i < candidatos.length; i++) {
      if (!candidatos[i].jaPintada) continue;
      const chave = formatarDataBr(candidatos[i].data);
      if (!mapa.has(chave)) mapa.set(chave, []);
      mapa.get(chave).push(candidatos[i]);
    }
    cursorPorProcedimento[proc] = 0;
    jaCorretaPorProcedimento[proc] = mapa;
  });

  const blocos = obterBlocosDoDoc(doc);
  if (blocos.length === 0) {
    ui.alert('Nenhuma evolução encontrada no formato esperado.');
    return;
  }

  const resumoPorProc = {};
  const usadasPorProcedimento = {};
  const acoes = [];

  blocos.forEach(bloco => {
    const proc = bloco.procedimento;
    if (!resumoPorProc[proc]) {
      resumoPorProc[proc] = { total: 0, encontrados: 0, semGuia: 0, jaCorreta: 0 };
      usadasPorProcedimento[proc] = 0;
    }
    resumoPorProc[proc].total++;

    const filaJaCorreta = jaCorretaPorProcedimento[proc];
    const fila = filaJaCorreta ? filaJaCorreta.get(bloco.dataStr) : null;
    if (fila && fila.length > 0) {
      fila.shift();
      resumoPorProc[proc].jaCorreta++;
      return;
    }

    const candidatos = porProcedimento[proc] || [];
    let cursor = cursorPorProcedimento[proc];
    if (cursor === undefined) cursor = 0;

    while (cursor < candidatos.length && candidatos[cursor].jaPintada) cursor++;

    const achou = cursor < candidatos.length ? candidatos[cursor] : null;

    if (achou) {
      cursorPorProcedimento[proc] = cursor + 1;
      usadasPorProcedimento[proc]++;
      resumoPorProc[proc].encontrados++;
      acoes.push({
        tipo: 'ATUALIZAR',
        elemento: bloco.elemento,
        dataAntiga: bloco.dataStr,
        dataNova: formatarDataBr(achou.data),
        linhaPlanilha: achou.linha
      });
    } else {
      cursorPorProcedimento[proc] = cursor;
      resumoPorProc[proc].semGuia++;
      acoes.push({
        tipo: 'SEM_GUIA',
        elemento: bloco.elemento,
        dataAntiga: bloco.dataStr
      });
    }
  });

  let sobrasPlanilha = 0;
  Object.keys(porProcedimento).forEach(proc => {
    const candidatos = porProcedimento[proc];
    const jaUsadasAntes = candidatos.filter(c => c.jaPintada).length;
    sobrasPlanilha += candidatos.length - jaUsadasAntes - (usadasPorProcedimento[proc] || 0);
  });

  let pintadasOcultas = 0;

  if (!dryRun) {
    aplicarAcoesNoDoc(acoes);
    const linhasPintadas = aplicarPinturaNaPlanilha(sheet, acoes);
    pintadasOcultas = contarLinhasOcultasPorFiltro(sheet, linhasPintadas);
  }

  mostrarResumo(dryRun, resumoPorProc, sobrasPlanilha, tituloDoc, pintadasOcultas);

  if (!dryRun) {
    perguntarEReordenar(doc, blocos, acoes);
  }
}

// ======================= AÇÕES DE MENU INDEPENDENTES =======================

/**
 * Ação de menu que reordena as evoluções de um Doc sem depender da planilha,
 * usando como chave a data já presente em cada cabeçalho. Informa quantas
 * evoluções sem data válida serão excluídas e exige confirmação antes de
 * qualquer remoção, lembrando que o histórico de versões do Doc permite
 * reverter a operação.
 */
function rodarSomenteReordenacao() {
  const ui = SpreadsheetApp.getUi();

  try {
    const resposta = ui.prompt(
      'Reordenar evoluções',
      'Link do Google Doc:',
      ui.ButtonSet.OK_CANCEL
    );
    if (resposta.getSelectedButton() !== ui.Button.OK) return;

    const url = resposta.getResponseText().trim();
    if (!url) {
      ui.alert('Nenhum link informado.');
      return;
    }

    const doc = DocumentApp.openById(extrairIdDoDoc(url));
    const blocos = obterBlocosDoDoc(doc);

    if (blocos.length === 0) {
      doc.saveAndClose();
      ui.alert(
        'Nenhuma evolução encontrada',
        'Formato esperado: "Evolução dia DD/MM/AAAA das..."',
        ui.ButtonSet.OK
      );
      return;
    }

    const acoes = blocos.map(bloco => {
      const ehMarcador = !parseDataBr(bloco.dataStr);
      return {
        tipo: ehMarcador ? 'SEM_GUIA' : 'ATUALIZAR',
        elemento: bloco.elemento,
        dataAntiga: bloco.dataStr,
        dataNova: bloco.dataStr
      };
    });

    const marcadores = acoes.filter(a => a.tipo === 'SEM_GUIA').length;
    if (marcadores > 0) {
      const confirma = ui.alert(
        'Evoluções sem guia',
        marcadores + ' evolução(ões) sem guia correspondente serão excluídas do Doc. ' +
        'Ação irreversível pelo script (use o histórico de versões do Doc para reverter). Continuar?',
        ui.ButtonSet.YES_NO
      );
      if (confirma !== ui.Button.YES) {
        doc.saveAndClose();
        ui.alert('Operação cancelada.');
        return;
      }
    }

    const resultado = reordenarEvolucoesPorData(doc, blocos, acoes);
    const quebras = normalizarQuebrasDePaginaDasEvolucoes(doc.getBody());

    doc.saveAndClose();

    let msg = quebras.totalEvolucoes + ' evolução(ões) processada(s).\n' +
      (resultado.reordenado ? 'Reordenadas por data.' : 'Já estavam em ordem.') +
      '\n' + quebras.configuradas + ' com quebra de página nova.';

    if (quebras.espacosLimpos > 0) {
      msg += '\n' + quebras.espacosLimpos + ' espaço(s) indevido(s) corrigido(s) no início de linha(s).';
    }

    if (resultado.excluidos > 0) {
      msg += '\n' + resultado.excluidos + ' excluída(s) por falta de data.';
    }

    ui.alert('Reordenação concluída', msg, ui.ButtonSet.OK);

  } catch (erro) {
    Logger.log('Erro na reordenação independente: ' + erro.stack);
    ui.alert('Erro ao reordenar', erro.message || String(erro), ui.ButtonSet.OK);
  }
}

/**
 * Ação de menu exclusivamente de formatação: garante que cada evolução (exceto
 * a primeira) comece em página nova e remove linhas em branco soltas antes dos
 * cabeçalhos. Não altera datas nem a ordem do conteúdo.
 */
function rodarPadronizarQuebrasDePagina() {
  const ui = SpreadsheetApp.getUi();

  try {
    const resposta = ui.prompt(
      'Padronizar quebras de página',
      'Link do Google Doc:',
      ui.ButtonSet.OK_CANCEL
    );
    if (resposta.getSelectedButton() !== ui.Button.OK) return;

    const url = resposta.getResponseText().trim();
    if (!url) {
      ui.alert('Nenhum link informado.');
      return;
    }

    const doc = DocumentApp.openById(extrairIdDoDoc(url));
    const resultado = normalizarQuebrasDePaginaDasEvolucoes(doc.getBody());

    if (resultado.totalEvolucoes === 0) {
      doc.saveAndClose();
      ui.alert(
        'Nenhuma evolução encontrada',
        'Formato esperado: "Evolução dia DD/MM/AAAA das..."',
        ui.ButtonSet.OK
      );
      return;
    }

    doc.saveAndClose();

    let msgPadronizacao = resultado.configuradas + ' de ' + resultado.totalEvolucoes + ' evolução(ões) com página nova.\n' +
      resultado.vaziosRemovidos + ' linha(s) em branco removida(s).';

    if (resultado.espacosLimpos > 0) {
      msgPadronizacao += '\n' + resultado.espacosLimpos + ' espaço(s) indevido(s) corrigido(s) no início de linha(s).';
    }

    ui.alert('Quebras padronizadas', msgPadronizacao, ui.ButtonSet.OK);
  } catch (erro) {
    Logger.log('Erro ao padronizar quebras de página: ' + erro.stack);
    ui.alert('Erro ao padronizar', erro.message || String(erro), ui.ButtonSet.OK);
  }
}

// ======================= LEITURA DO DOC =======================

/**
 * Constrói a expressão regular que reconhece o cabeçalho de uma evolução.
 * Aceita tanto uma data (DD/MM/AAAA) quanto o marcador de evolução sem guia,
 * o que permite reprocessar com segurança um Doc já tratado anteriormente.
 */
function regexCabecalho() {
  return new RegExp(
    'Evolu[çc][ãa]o\\s+dia\\s+(\\d{2}\\/\\d{2}\\/\\d{4}|' +
    escaparRegex(CONFIG.MARCADOR_SEM_GUIA) + ')',
    'i'
  );
}

const REGEX_PROCEDIMENTO = /Procedimentos?:\s*([^\n\r]+)/i;

/**
 * Lê o corpo do Doc em uma única varredura e devolve, para cada elemento de
 * primeiro nível, seu tipo, referência e texto. As etapas seguintes reutilizam
 * essa leitura em memória, evitando chamadas repetidas à API do Docs, que são
 * o principal gargalo de desempenho em documentos longos.
 */
function lerCorpo(body) {
  const total = body.getNumChildren();
  const itens = new Array(total);
  const TIPO_PARAGRAFO = DocumentApp.ElementType.PARAGRAPH;
  const TIPO_LISTA = DocumentApp.ElementType.LIST_ITEM;

  for (let i = 0; i < total; i++) {
    const filho = body.getChild(i);
    const tipo = filho.getType();
    let elemento = filho;
    let texto = null;

    if (tipo === TIPO_PARAGRAFO) {
      elemento = filho.asParagraph();
      texto = elemento.getText();
    } else if (tipo === TIPO_LISTA) {
      elemento = filho.asListItem();
      texto = elemento.getText();
    }

    itens[i] = { indice: i, tipo: tipo, elemento: elemento, texto: texto };
  }

  return itens;
}

/**
 * Identifica os blocos de evolução do Doc, cada um formado por um cabeçalho
 * ("Evolução dia ...") e pela linha "Procedimentos:" que o segue. Descarta
 * parágrafos vazios antes de aplicar qualquer regex e devolve a referência
 * viva do parágrafo, e não sua posição, de modo que o resultado continua
 * válido mesmo depois de outros parágrafos serem removidos.
 */
function obterBlocosDoDoc(doc) {
  const itens = lerCorpo(doc.getBody());
  const rxCabecalho = regexCabecalho();
  const blocos = [];
  let blocoAtual = null;

  for (let i = 0; i < itens.length; i++) {
    const item = itens[i];
    if (item.texto === null) continue;

    const texto = item.texto;
    if (texto.trim() === '') continue;

    const matchCabecalho = texto.match(rxCabecalho);
    if (matchCabecalho) {
      blocoAtual = {
        elemento: item.elemento,
        dataStr: matchCabecalho[1],
        procedimento: null
      };
      continue;
    }

    if (blocoAtual && blocoAtual.procedimento === null) {
      const matchProc = texto.match(REGEX_PROCEDIMENTO);
      if (matchProc) {
        blocoAtual.procedimento = normalizarProcedimento(matchProc[1]);
        blocos.push(blocoAtual);
        blocoAtual = null;
      }
    }
  }

  return blocos;
}

/**
 * Aplica ao Doc as ações calculadas: substitui a data do cabeçalho pela data
 * da guia correspondente ou, quando não há guia, pelo marcador de evolução
 * sem guia. Ignora ações que não mudariam nada.
 */
function aplicarAcoesNoDoc(acoes) {
  acoes.forEach(acao => {
    const elemento = acao.elemento;
    if (!elemento) return;

    const novoValor = acao.tipo === 'ATUALIZAR' ? acao.dataNova : CONFIG.MARCADOR_SEM_GUIA;
    if (acao.dataAntiga === novoValor) return;

    elemento.editAsText().replaceText(escaparRegex(acao.dataAntiga), novoValor);
  });
}

// ======================= LIMPEZA / PREPARAÇÃO =======================

/**
 * Prepara o corpo do Doc para a reordenação em uma única varredura. Remove as
 * linhas em branco que antecedem cada cabeçalho (reduzindo drasticamente a
 * quantidade de elementos a copiar), limpa elementos decorativos presos em
 * parágrafos vazios remanescentes e detecta elementos não suportados, que só
 * são removidos mediante confirmação do usuário. Nenhum cabeçalho nem texto
 * clínico é tocado.
 */
function prepararCorpoParaReordenacao(body) {
  const itens = lerCorpo(body);
  const rxCabecalho = regexCabecalho();

  const naoSuportados = [];
  const vaziosParaRemover = [];
  const vaziosMantidos = [];
  let pendentes = [];

  for (let i = 0; i < itens.length; i++) {
    const item = itens[i];

    if (item.tipo === DocumentApp.ElementType.UNSUPPORTED) {
      naoSuportados.push(item.elemento);
      pendentes = [];
      continue;
    }

    if (item.texto === null) {
      pendentes = [];
      continue;
    }

    const texto = item.texto.trim();

    if (texto === '') {
      if (item.tipo === DocumentApp.ElementType.PARAGRAPH) {
        pendentes.push(item.elemento);
      } else {
        pendentes = [];
      }
      continue;
    }

    if (rxCabecalho.test(texto)) {
      for (let j = 0; j < pendentes.length; j++) vaziosParaRemover.push(pendentes[j]);
    } else {
      for (let j = 0; j < pendentes.length; j++) vaziosMantidos.push(pendentes[j]);
    }

    pendentes = [];
  }

  for (let j = 0; j < pendentes.length; j++) vaziosMantidos.push(pendentes[j]);

  if (naoSuportados.length > 0) {
    confirmarRemocaoDeNaoSuportados(body, naoSuportados);
  }

  for (let i = vaziosParaRemover.length - 1; i >= 0; i--) {
    body.removeChild(vaziosParaRemover[i]);
  }

  for (let i = 0; i < vaziosMantidos.length; i++) {
    const paragrafo = vaziosMantidos[i];
    for (let j = paragrafo.getNumChildren() - 1; j >= 0; j--) {
      const interno = paragrafo.getChild(j);
      if (interno.getType() !== DocumentApp.ElementType.TEXT) {
        paragrafo.removeChild(interno);
      }
    }
  }

  return {
    vaziosRemovidos: vaziosParaRemover.length,
    naoSuportadosRemovidos: naoSuportados.length
  };
}

/**
 * Solicita confirmação do usuário antes de remover elementos não suportados
 * (resíduos invisíveis de conversão do Word). Se o usuário recusar, a
 * reordenação é abortada com um erro explícito, sem alterar o documento.
 */
function confirmarRemocaoDeNaoSuportados(body, elementos) {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    'Elementos não identificados',
    elementos.length + ' elemento(s) do documento não foram reconhecidos (resíduo da conversão ' +
    'do Word). Não fazem parte do texto de nenhuma evolução.\n\n' +
    'Remover para continuar a reordenação? Ação irreversível pelo script.',
    ui.ButtonSet.YES_NO
  );

  if (resp !== ui.Button.YES) {
    throw new Error(
      'Reordenação cancelada: ' + elementos.length + ' elemento(s) não suportado(s) não removido(s).'
    );
  }

  for (let i = elementos.length - 1; i >= 0; i--) {
    body.removeChild(elementos[i]);
  }
}

// ======================= QUEBRAS DE PÁGINA =======================

/**
 * Padroniza o layout das evoluções para que cada uma, exceto a primeira,
 * comece no topo de uma página. A quebra é inserida como primeiro filho do
 * próprio parágrafo do cabeçalho, de modo que acompanha o bloco quando ele é
 * movido. Também remove linhas em branco antes dos cabeçalhos e espaços
 * indevidos no início das linhas de cabeçalho e "Procedimentos:". Opera em uma
 * única varredura e só altera o que está fora do padrão, o que torna a rotina
 * idempotente.
 */
function normalizarQuebrasDePaginaDasEvolucoes(body) {
  const itens = lerCorpo(body);
  const rxCabecalho = regexCabecalho();

  const vaziosParaRemover = [];
  const quebrasParaRemover = [];
  const cabecalhosParaInserirQuebra = [];
  let pendentes = [];
  let ordemCabecalho = 0;
  let configuradas = 0;
  let espacosLimpos = 0;

  for (let i = 0; i < itens.length; i++) {
    const item = itens[i];

    if (item.texto === null) {
      pendentes = [];
      continue;
    }

    const textoOriginal = item.texto;
    const texto = textoOriginal.trim();

    if (texto === '') {
      pendentes.push(item.elemento);
      continue;
    }

    if (!rxCabecalho.test(texto)) {
      if (/^\s/.test(textoOriginal) && REGEX_PROCEDIMENTO.test(texto)) {
        item.elemento.editAsText().replaceText('^\\s+', '');
        espacosLimpos++;
      }
      pendentes = [];
      continue;
    }

    for (let j = 0; j < pendentes.length; j++) vaziosParaRemover.push(pendentes[j]);
    pendentes = [];

    const cabecalho = item.elemento;

    if (/^\s/.test(textoOriginal)) {
      cabecalho.editAsText().replaceText('^\\s+', '');
      espacosLimpos++;
    }

    const quebras = [];
    for (let j = cabecalho.getNumChildren() - 1; j >= 0; j--) {
      const interno = cabecalho.getChild(j);
      if (interno.getType() === DocumentApp.ElementType.PAGE_BREAK) {
        quebras.push({ elemento: interno, indice: j });
      }
    }

    if (ordemCabecalho === 0) {
      for (let j = 0; j < quebras.length; j++) {
        quebrasParaRemover.push({ pai: cabecalho, elemento: quebras[j].elemento });
      }
    } else {
      configuradas++;
      const jaEstaCerto = quebras.length === 1 && quebras[0].indice === 0;
      if (!jaEstaCerto) {
        for (let j = 0; j < quebras.length; j++) {
          quebrasParaRemover.push({ pai: cabecalho, elemento: quebras[j].elemento });
        }
        cabecalhosParaInserirQuebra.push(cabecalho);
      }
    }

    ordemCabecalho++;
  }

  for (let i = 0; i < quebrasParaRemover.length; i++) {
    quebrasParaRemover[i].pai.removeChild(quebrasParaRemover[i].elemento);
  }

  for (let i = vaziosParaRemover.length - 1; i >= 0; i--) {
    body.removeChild(vaziosParaRemover[i]);
  }

  for (let i = 0; i < cabecalhosParaInserirQuebra.length; i++) {
    cabecalhosParaInserirQuebra[i].insertPageBreak(0);
  }

  return {
    configuradas: configuradas,
    totalEvolucoes: ordemCabecalho,
    vaziosRemovidos: vaziosParaRemover.length,
    espacosLimpos: espacosLimpos
  };
}

// ======================= FORMATAÇÃO GERAL DO DOC =======================

/**
 * Aplica a formatação padrão ao corpo do documento (fonte Calibri, recuos,
 * espaçamentos e entrelinhas), preservando conteúdo, tamanho de fonte,
 * negrito, itálico, cores e alinhamento.
 */
function formatarDocumentoPadrao(doc) {
  formatarElementoDocumentoRecursivo(doc.getBody());
}

/**
 * Percorre recursivamente a árvore do documento (corpo, tabelas, linhas e
 * células) e aplica a formatação padrão a cada parágrafo e item de lista, de
 * modo que o texto dentro de tabelas também seja padronizado.
 */
function formatarElementoDocumentoRecursivo(elemento) {
  const tipo = elemento.getType();
  const ehParagrafo = tipo === DocumentApp.ElementType.PARAGRAPH;
  const ehItemLista = tipo === DocumentApp.ElementType.LIST_ITEM;

  if (ehParagrafo || ehItemLista) {
    const paragrafo = ehParagrafo ? elemento.asParagraph() : elemento.asListItem();

    paragrafo
      .setIndentStart(0)
      .setIndentFirstLine(0)
      .setIndentEnd(19)
      .setSpacingBefore(0)
      .setSpacingAfter(0)
      .setLineSpacing(1.08);

    const texto = paragrafo.editAsText();
    if (texto.getText().length > 0) {
      texto.setFontFamily('Calibri');
    }
    return;
  }

  if (typeof elemento.getNumChildren === 'function') {
    const quantidade = elemento.getNumChildren();
    for (let i = 0; i < quantidade; i++) {
      formatarElementoDocumentoRecursivo(elemento.getChild(i));
    }
  }
}

/**
 * Ação de menu independente que pede o link de um Doc e aplica a formatação
 * padrão. Não altera datas nem a planilha, e garante o fechamento do
 * documento mesmo em caso de falha.
 */
function rodarFormatarDocumento() {
  const ui = SpreadsheetApp.getUi();
  const resposta = ui.prompt(
    'Formatar documento',
    'Link do Google Doc que deseja formatar:',
    ui.ButtonSet.OK_CANCEL
  );

  if (resposta.getSelectedButton() !== ui.Button.OK) return;

  const url = resposta.getResponseText().trim();
  if (!url) {
    ui.alert('Nenhum link informado.');
    return;
  }

  let doc = null;
  try {
    doc = DocumentApp.openById(extrairIdDoDoc(url));
    formatarDocumentoPadrao(doc);
    doc.saveAndClose();

    ui.alert(
      'Formatação concluída',
      'O corpo do documento foi formatado com Calibri, recuo esquerdo de 0 pt, recuo direito de 19 pt, espaçamento antes/depois de 0 e entrelinhas 1,08. O conteúdo não foi reescrito.',
      ui.ButtonSet.OK
    );
  } catch (erro) {
    if (doc) {
      try {
        doc.saveAndClose();
      } catch (erroFechamento) {
        Logger.log('Erro ao fechar o documento após falha de formatação: ' + (erroFechamento.stack || erroFechamento));
      }
    }

    Logger.log('Erro ao formatar documento pelo menu: ' + (erro.stack || erro));
    ui.alert('Erro ao formatar documento', erro.message || String(erro), ui.ButtonSet.OK);
  }
}

// ======================= REORDENAÇÃO =======================

/**
 * Segunda etapa do fluxo principal: reordena as evoluções, padroniza as
 * quebras de página e formata o documento, sem novas perguntas ao usuário.
 * Como as datas e a planilha já foram gravadas na etapa anterior, qualquer
 * falha aqui é tratada isoladamente e não desfaz esse trabalho. A formatação
 * é opcional: se falhar, a reordenação continua válida e o usuário é avisado.
 */
function perguntarEReordenar(doc, blocos, acoes) {
  const ui = SpreadsheetApp.getUi();

  try {
    const resultado = reordenarEvolucoesPorData(doc, blocos, acoes);
    const quebras = normalizarQuebrasDePaginaDasEvolucoes(doc.getBody());

    let erroFormatacao = null;
    try {
      formatarDocumentoPadrao(doc);
    } catch (erro) {
      erroFormatacao = erro;
      Logger.log('Erro ao aplicar formatação padrão: ' + (erro.stack || erro));
    }

    doc.saveAndClose();

    let msg = resultado.reordenado
      ? 'Evoluções reordenadas por data (mais recentes primeiro).'
      : 'Evoluções já estavam em ordem.';

    msg += '\n' + quebras.configuradas + ' com quebra de página nova.';

    if (quebras.espacosLimpos > 0) {
      msg += '\n' + quebras.espacosLimpos + ' espaço(s) indevido(s) corrigido(s) no início de linha(s).';
    }

    if (resultado.excluidos > 0) {
      msg += '\n' + resultado.excluidos + ' excluída(s) por estarem marcadas como "' +
        CONFIG.MARCADOR_SEM_GUIA + '".';
    }

    if (erroFormatacao) {
      msg += '\nAtenção: a reordenação foi concluída, mas a formatação não foi aplicada: ' +
        (erroFormatacao.message || String(erroFormatacao));
    } else {
      msg += '\nFormatação geral aplicada ao documento.';
    }

    ui.alert('Reordenação concluída', msg, ui.ButtonSet.OK);
  } catch (e) {
    Logger.log('Erro ao reordenar: ' + (e.stack || e));
    doc.saveAndClose();
    ui.alert(
      'Erro ao reordenar',
      'As datas e a planilha já foram salvas normalmente na etapa anterior ' +
      'e não foram afetadas por este erro.\n\n' + e.message,
      ui.ButtonSet.OK
    );
  }
}

/**
 * Reordena os blocos de evolução por data, das mais recentes para as mais
 * antigas, e exclui os blocos sem guia.
 *
 * O algoritmo trabalha com referências de elementos e consulta a posição atual
 * apenas no momento de mover, eliminando a aritmética de deslocamento de
 * índices e a classe de bugs associada. Antes de modificar o documento, valida
 * a consistência dos blocos e os tipos de elemento suportados e faz cópias de
 * todo o conteúdo. As cópias são inseridas ao final e só então os originais
 * são removidos, de modo que qualquer falha anterior deixa o documento
 * intacto. A ordenação é estável: datas iguais preservam a ordem original.
 */
function reordenarEvolucoesPorData(doc, blocos, acoes) {
  if (!blocos || blocos.length === 0) {
    return { reordenado: false, excluidos: 0 };
  }

  if (acoes.length !== blocos.length) {
    throw new Error(
      'Inconsistência interna entre evoluções e ações (' + blocos.length + ' vs ' +
      acoes.length + '). Nada foi alterado.'
    );
  }

  const body = doc.getBody();

  prepararCorpoParaReordenacao(body);

  const inicios = new Array(blocos.length);
  for (let i = 0; i < blocos.length; i++) {
    let indice = -1;
    try {
      indice = body.getChildIndex(blocos[i].elemento);
    } catch (e) {
      indice = -1;
    }

    if (indice < 0) {
      throw new Error(
        'A evolução "' + blocos[i].dataStr + '" não está no corpo principal do documento ' +
        '(pode estar em tabela ou caixa de texto). Nada foi movido.'
      );
    }

    if (i > 0 && indice <= inicios[i - 1]) {
      throw new Error(
        'Evoluções fora de ordem no documento, perto de "' + blocos[i].dataStr + '". Nada foi movido.'
      );
    }

    inicios[i] = indice;
  }

  const totalFilhos = body.getNumChildren();
  const acaoPorBloco = acoes;
  const blocosCompletos = [];

  for (let i = 0; i < blocos.length; i++) {
    const inicioReal = inicios[i];
    const fimReal = (i + 1 < blocos.length) ? inicios[i + 1] - 1 : totalFilhos - 1;

    if (fimReal < inicioReal) {
      throw new Error(
        'Não foi possível determinar o fim da evolução "' + blocos[i].dataStr + '". Nada foi movido.'
      );
    }

    const acao = acaoPorBloco[i];
    const dataOrdenacao = (acao && acao.tipo === 'ATUALIZAR') ? parseDataBr(acao.dataNova) : null;
    const temGuia = !!dataOrdenacao;

    blocosCompletos.push({
      ordemOriginal: i,
      inicioReal: inicioReal,
      fimReal: fimReal,
      temGuia: temGuia,
      dataOrdenacao: dataOrdenacao
    });
  }

  const blocosParaManter = blocosCompletos.filter(b => b.temGuia);
  const blocosParaExcluir = blocosCompletos.filter(b => !b.temGuia);

  if (blocosParaManter.length === 0) {
    throw new Error('Nenhuma evolução com data válida para ordenar.');
  }

  const ordenados = blocosParaManter.slice().sort((a, b) => {
    const diferenca = b.dataOrdenacao - a.dataOrdenacao;
    return diferenca !== 0 ? diferenca : a.ordemOriginal - b.ordemOriginal;
  });

  const jaOrdenado =
    blocosParaExcluir.length === 0 &&
    ordenados.every((b, i) => b.ordemOriginal === i);

  if (jaOrdenado) {
    return { reordenado: false, excluidos: 0 };
  }

  const inicioRegiao = inicios[0];
  const fimRegiao = totalFilhos - 1;

  const elementos = new Array(fimRegiao + 1);
  for (let i = inicioRegiao; i <= fimRegiao; i++) {
    const filho = body.getChild(i);
    elementos[i] = { elemento: filho, tipo: filho.getType() };
  }

  const tiposSuportados = new Set([
    DocumentApp.ElementType.PARAGRAPH,
    DocumentApp.ElementType.LIST_ITEM,
    DocumentApp.ElementType.TABLE,
    DocumentApp.ElementType.INLINE_IMAGE,
    DocumentApp.ElementType.HORIZONTAL_RULE,
    DocumentApp.ElementType.PAGE_BREAK
  ]);

  const estranhos = [];
  for (let i = inicioRegiao; i <= fimRegiao; i++) {
    if (!tiposSuportados.has(elementos[i].tipo)) {
      estranhos.push('posição ' + i + ' — tipo "' + elementos[i].tipo + '"');
    }
  }

  if (estranhos.length > 0) {
    throw new Error(
      'Elemento(s) que não podem ser movidos com segurança:\n' +
      estranhos.slice(0, 5).join('\n') +
      (estranhos.length > 5 ? '\n... e mais ' + (estranhos.length - 5) + '.' : '') +
      '\n\nNada foi alterado nesta etapa. As datas continuam salvas normalmente.'
    );
  }

  const elementosParaRemover = [];
  const copiasPorBloco = ordenados.map(bloco => {
    const copias = [];
    for (let k = bloco.inicioReal; k <= bloco.fimReal; k++) {
      const registro = elementos[k];
      copias.push({ tipo: registro.tipo, elemento: registro.elemento.copy() });
      elementosParaRemover.push(registro.elemento);
    }
    return copias;
  });

  blocosParaExcluir.forEach(bloco => {
    for (let k = bloco.inicioReal; k <= bloco.fimReal; k++) {
      elementosParaRemover.push(elementos[k].elemento);
    }
  });

  let cursorInsercao = fimRegiao + 1;
  copiasPorBloco.forEach(copias => {
    copias.forEach(item => {
      inserirElementoNoCorpo(body, cursorInsercao, item.elemento);
      cursorInsercao++;
    });
  });

  for (let i = elementosParaRemover.length - 1; i >= 0; i--) {
    body.removeChild(elementosParaRemover[i]);
  }

  return { reordenado: true, excluidos: blocosParaExcluir.length };
}

/**
 * Insere no corpo do documento, na posição indicada, uma cópia de um elemento,
 * despachando para o método de inserção adequado a cada tipo (parágrafo, item
 * de lista, tabela, imagem, linha horizontal ou quebra de página). Falha de
 * forma explícita diante de um tipo desconhecido.
 */
function inserirElementoNoCorpo(body, indice, elemento) {
  const tipo = elemento.getType();

  switch (tipo) {
    case DocumentApp.ElementType.PARAGRAPH:
      return body.insertParagraph(indice, elemento.asParagraph());
    case DocumentApp.ElementType.LIST_ITEM:
      return body.insertListItem(indice, elemento.asListItem());
    case DocumentApp.ElementType.TABLE:
      return body.insertTable(indice, elemento.asTable());
    case DocumentApp.ElementType.INLINE_IMAGE:
      return body.insertImage(indice, elemento.asInlineImage());
    case DocumentApp.ElementType.HORIZONTAL_RULE:
      return body.insertHorizontalRule(indice);
    case DocumentApp.ElementType.PAGE_BREAK:
      return body.insertPageBreak(indice);
    default:
      throw new Error('Elemento não suportado: ' + tipo + '.');
  }
}

// ======================= PLANILHA =======================

/**
 * Localiza, pela linha de cabeçalho da planilha, as colunas PACIENTE, TERAPIA
 * e DATA DE ATENDIMENTO, ignorando acentos e caixa. Devolve null se alguma das
 * colunas obrigatórias não existir.
 */
function obterMapaDeColunas(sheet) {
  const cabecalho = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  cabecalho.forEach((nome, idx) => {
    const n = normalizarTexto(String(nome));
    if (n.indexOf('PACIENTE') !== -1 && map.paciente === undefined) map.paciente = idx;
    if (n.indexOf('TERAPIA') !== -1 && map.terapia === undefined) map.terapia = idx;
    if (n.indexOf('DATA') !== -1 && n.indexOf('ATENDIMENTO') !== -1 && map.data === undefined) map.data = idx;
  });
  if (map.paciente === undefined || map.terapia === undefined || map.data === undefined) return null;
  return map;
}

/**
 * Identifica o paciente a partir do título do Doc e devolve as guias dele,
 * ordenadas da mais recente para a mais antiga, indicando quais já foram
 * processadas (linhas pintadas). A planilha é lida em lote: uma chamada para
 * os valores e outra apenas para as cores da coluna PACIENTE, o que mantém o
 * desempenho aceitável mesmo com todos os pacientes na mesma aba.
 */
function obterLinhasDoPaciente(sheet, headerMap, tituloDoc) {
  const ultimaLinha = sheet.getLastRow();
  const ultimaColuna = sheet.getLastColumn();
  if (ultimaLinha < 2) {
    return { ok: false, motivo: 'A planilha está vazia.' };
  }

  const quantidadeLinhas = ultimaLinha - 1;

  const valores = sheet.getRange(2, 1, quantidadeLinhas, ultimaColuna).getValues();
  const backgroundsPaciente = sheet
    .getRange(2, headerMap.paciente + 1, quantidadeLinhas, 1)
    .getBackgrounds();

  const nomesDistintos = new Set();
  const pacientesNormalizados = new Array(quantidadeLinhas);

  for (let i = 0; i < quantidadeLinhas; i++) {
    const paciente = normalizarTexto(String(valores[i][headerMap.paciente] || ''));
    pacientesNormalizados[i] = paciente;
    if (paciente) nomesDistintos.add(paciente);
  }

  const checagem = conferirNomePaciente(tituloDoc, [...nomesDistintos]);
  if (!checagem.ok) return { ok: false, motivo: checagem.motivo };

  const nomeAlvo = checagem.nomeEscolhido;
  const linhas = [];
  const corPintada = CONFIG.COR_PINTADA.toLowerCase();

  for (let i = 0; i < quantidadeLinhas; i++) {
    if (pacientesNormalizados[i] !== nomeAlvo) continue;

    const terapia = normalizarProcedimento(String(valores[i][headerMap.terapia] || ''));
    const dataRaw = valores[i][headerMap.data];
    const data = (dataRaw instanceof Date) ? dataRaw : new Date(dataRaw);
    if (!terapia || isNaN(data.getTime())) continue;

    const corFundo = backgroundsPaciente[i][0];

    linhas.push({
      linha: i + 2,
      terapia: terapia,
      data: data,
      jaPintada: !!corFundo && corFundo.toLowerCase() === corPintada
    });
  }

  linhas.sort((a, b) => b.data - a.data);

  return { ok: true, linhas: linhas };
}

/**
 * Pinta na planilha as linhas cujas guias foram consumidas, usando uma única
 * chamada de RangeList em vez de uma chamada por linha. Devolve as linhas
 * pintadas, sem duplicatas.
 */
function aplicarPinturaNaPlanilha(sheet, acoes) {
  const linhasUsadas = [...new Set(
    acoes.filter(a => a.tipo === 'ATUALIZAR').map(a => a.linhaPlanilha)
  )];
  if (linhasUsadas.length === 0) return [];

  const ultimaColunaA1 = colunaParaLetra(sheet.getLastColumn());
  const enderecos = linhasUsadas.map(linha => 'A' + linha + ':' + ultimaColunaA1 + linha);

  sheet.getRangeList(enderecos).setBackground(CONFIG.COR_PINTADA);

  return linhasUsadas;
}

/**
 * Conta quantas das linhas recém-pintadas estão ocultas por um filtro ativo na
 * aba. Como o script processa a planilha inteira ignorando filtros, essa
 * verificação evita a falsa impressão de que nada mudou quando o filtro
 * esconde justamente as linhas alteradas.
 */
function contarLinhasOcultasPorFiltro(sheet, linhas) {
  if (!linhas || linhas.length === 0) return 0;
  if (!sheet.getFilter()) return 0;

  let ocultas = 0;
  for (let i = 0; i < linhas.length; i++) {
    if (sheet.isRowHiddenByFilter(linhas[i])) ocultas++;
  }
  return ocultas;
}

/**
 * Converte o número de uma coluna (1, 2, ..., 27) na letra correspondente da
 * notação A1 (A, B, ..., AA), tratando corretamente colunas de duas ou mais
 * letras.
 */
function colunaParaLetra(numero) {
  let resultado = '';
  while (numero > 0) {
    const resto = (numero - 1) % 26;
    resultado = String.fromCharCode(65 + resto) + resultado;
    numero = Math.floor((numero - 1) / 26);
  }
  return resultado;
}

// ======================= UTILITÁRIOS =======================

/**
 * Extrai o ID do documento a partir da URL do Google Docs, lançando um erro
 * claro quando o link é inválido.
 */
function extrairIdDoDoc(url) {
  const match = url.match(/[-\w]{25,}/);
  if (!match) throw new Error('Link inválido: ID do documento não encontrado.');
  return match[0];
}

/**
 * Normaliza um texto para comparação: remove acentos, converte para
 * maiúsculas e colapsa espaços repetidos.
 */
function normalizarTexto(txt) {
  return txt
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toUpperCase()
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Normaliza o nome de um procedimento nos dois lados da comparação (Doc e
 * planilha). Descarta sufixos que vêm colados do portal do convênio, como
 * "Código TUSS" ou o início do texto da evolução, para que o nome da terapia
 * case de forma confiável e nenhuma evolução com guia seja tratada como
 * excedente.
 */
function normalizarProcedimento(txt) {
  let t = normalizarTexto(String(txt || ''));
  t = t.replace(/\s*(CODIGO|TUSS|EVOLUCAO)\b[\s\S]*$/, '');
  t = t.replace(/[\s:;.,\-–]+$/, '');
  return t.trim();
}

/**
 * Escapa os caracteres especiais de expressão regular de um texto, permitindo
 * usá-lo como padrão literal.
 */
function escaparRegex(str) {
  return String(str).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Formata um Date como DD/MM/AAAA a partir de seus componentes locais, sem
 * depender do fuso horário do script.
 */
function formatarDataBr(data) {
  const dd = ('0' + data.getDate()).slice(-2);
  const mm = ('0' + (data.getMonth() + 1)).slice(-2);
  return dd + '/' + mm + '/' + data.getFullYear();
}

/**
 * Converte uma string DD/MM/AAAA em Date, validando o calendário (rejeita, por
 * exemplo, 31/02). Devolve null quando o texto não é uma data válida.
 */
function parseDataBr(str) {
  if (!str) return null;
  const match = String(str).match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (!match) return null;

  const dia = parseInt(match[1], 10);
  const mes = parseInt(match[2], 10) - 1;
  const ano = parseInt(match[3], 10);
  const data = new Date(ano, mes, dia);

  if (data.getFullYear() !== ano || data.getMonth() !== mes || data.getDate() !== dia) {
    return null;
  }
  return data;
}

/**
 * Calcula a distância de Levenshtein entre duas strings (número mínimo de
 * inserções, remoções e substituições). Usa apenas duas linhas da matriz de
 * programação dinâmica, reduzindo o uso de memória de O(m·n) para O(n).
 */
function levenshtein(a, b) {
  const m = a.length, n = b.length;
  if (m === 0) return n;
  if (n === 0) return m;

  let anterior = new Array(n + 1);
  let atual = new Array(n + 1);

  for (let j = 0; j <= n; j++) anterior[j] = j;

  for (let i = 1; i <= m; i++) {
    atual[0] = i;
    const ca = a[i - 1];
    for (let j = 1; j <= n; j++) {
      const custo = ca === b[j - 1] ? 0 : 1;
      atual[j] = Math.min(anterior[j] + 1, atual[j - 1] + 1, anterior[j - 1] + custo);
    }
    const troca = anterior;
    anterior = atual;
    atual = troca;
  }

  return anterior[n];
}

/**
 * Devolve a similaridade entre duas strings, de 0 a 1, como o complemento da
 * distância de Levenshtein normalizada pelo comprimento da maior string.
 */
function similaridade(a, b) {
  const maior = Math.max(a.length, b.length);
  if (maior === 0) return 1;
  return 1 - (levenshtein(a, b) / maior);
}

/**
 * Associa o título do Doc ao nome do paciente na planilha, com tolerância a
 * erros de digitação. Prefere a correspondência exata, mas só a aceita se não
 * existir outro nome quase idêntico. Na ausência de correspondência exata,
 * escolhe o melhor candidato acima do limiar de similaridade. Em qualquer
 * caso ambíguo, recusa-se a decidir e pede ao usuário que corrija os nomes,
 * evitando gravar dados no paciente errado.
 */
function conferirNomePaciente(tituloDoc, nomesPlanilha) {
  const tituloNorm = normalizarTexto(tituloDoc.replace(/[_\-]/g, ' '));

  if (nomesPlanilha.length === 0) {
    return { ok: false, motivo: 'Nenhum nome de paciente encontrado na planilha.' };
  }

  if (nomesPlanilha.indexOf(tituloNorm) !== -1) {
    let melhorAlternativo = null;

    for (let i = 0; i < nomesPlanilha.length; i++) {
      const nome = nomesPlanilha[i];
      if (nome === tituloNorm) continue;

      const maiorTamanho = Math.max(tituloNorm.length, nome.length);
      if (maiorTamanho === 0 ||
          Math.abs(tituloNorm.length - nome.length) / maiorTamanho >= 0.05) {
        continue;
      }

      const score = similaridade(tituloNorm, nome);
      if (score >= CONFIG.SIMILARIDADE_MINIMA &&
          (1 - score) < 0.05 &&
          (!melhorAlternativo || score > melhorAlternativo.score)) {
        melhorAlternativo = { nome: nome, score: score };
      }
    }

    if (melhorAlternativo) {
      return {
        ok: false,
        motivo: 'Ambiguidade entre "' + tituloNorm + '" e "' + melhorAlternativo.nome +
          '" para o título "' + tituloDoc + '". Renomeie o Doc ou ajuste os nomes na planilha.'
      };
    }

    return { ok: true, nomeEscolhido: tituloNorm };
  }

  let melhor = null;
  let segundo = null;

  for (let i = 0; i < nomesPlanilha.length; i++) {
    const candidato = { nome: nomesPlanilha[i], score: similaridade(tituloNorm, nomesPlanilha[i]) };
    if (!melhor || candidato.score > melhor.score) {
      segundo = melhor;
      melhor = candidato;
    } else if (!segundo || candidato.score > segundo.score) {
      segundo = candidato;
    }
  }

  if (melhor.score < CONFIG.SIMILARIDADE_MINIMA) {
    return {
      ok: false,
      motivo: 'Título do Doc: "' + tituloDoc + '"\nMelhor candidato: "' + melhor.nome +
        '" (' + (melhor.score * 100).toFixed(0) + '% de similaridade, mínimo ' +
        (CONFIG.SIMILARIDADE_MINIMA * 100).toFixed(0) + '%).'
    };
  }

  if (segundo && segundo.score >= CONFIG.SIMILARIDADE_MINIMA &&
      (melhor.score - segundo.score) < 0.05) {
    return {
      ok: false,
      motivo: 'Ambiguidade entre "' + melhor.nome + '" e "' + segundo.nome + '" para o título "' +
        tituloDoc + '". Renomeie o Doc ou ajuste os nomes na planilha.'
    };
  }

  return { ok: true, nomeEscolhido: melhor.nome };
}

/**
 * Exibe o resumo da execução em um diálogo: por terapia, quantas evoluções
 * foram atualizadas, já estavam corretas ou ficaram sem guia, além das linhas
 * da planilha não utilizadas e de um aviso quando há linhas marcadas ocultas
 * por filtro.
 */
function mostrarResumo(dryRun, resumoPorProc, sobrasPlanilha, tituloDoc, pintadasOcultas) {
  const ui = SpreadsheetApp.getUi();
  let msg = 'Paciente: ' + tituloDoc + '\n\n';

  Object.keys(resumoPorProc).forEach(proc => {
    const r = resumoPorProc[proc];
    msg += proc + ': ' + r.encontrados + ' atualizada(s)';
    if (r.jaCorreta > 0) msg += ', ' + r.jaCorreta + ' já corretas';
    if (r.semGuia > 0) msg += ', ' + r.semGuia + ' sem guia';
    msg += ' — ' + r.total + ' no total\n';
  });

  msg += '\nLinhas não utilizadas na planilha: ' + sobrasPlanilha;

  if (pintadasOcultas > 0) {
    msg += '\n\nAviso: ' + pintadasOcultas + ' linha(s) marcada(s) estão ocultas pelo filtro ativo.';
  }

  ui.alert(dryRun ? 'Simulação (nenhuma alteração aplicada)' : 'Automação aplicada', msg, ui.ButtonSet.OK);
}
