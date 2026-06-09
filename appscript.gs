/****************************************************
 * APP SCRIPT UNIFICADO
 * Módulos:
 * - LANÇAMENTOS
 * - BOLETOS PAGOS
 *
 * Regra principal:
 * - Deve existir apenas um onOpen() no projeto.
 * - Cada módulo tem seu próprio menu.
 * - Helpers foram prefixados para evitar colisões.
 ****************************************************/


/****************************************************
 * MENU CENTRAL
 ****************************************************/

const APP_FORMATOS = Object.freeze({
  DATA: 'dd/mm/yyyy',
  DATA_HORA: 'dd/mm/yyyy hh:mm',
  MOEDA_BRL: '[$R$-416] #,##0.00;[Red]([$R$-416] #,##0.00);-',
  INTEIRO: '0'
});


function APP_a1Range_(row, column, numRows, numColumns) {
  const height = numRows || 1;
  const width = numColumns || 1;
  const start = APP_columnToLetter_(column) + row;

  if (height === 1 && width === 1) {
    return start;
  }

  return start + ':' + APP_columnToLetter_(column + width - 1) + (row + height - 1);
}


function APP_columnToLetter_(column) {
  let remaining = column;
  let letter = '';

  while (remaining > 0) {
    const modulo = (remaining - 1) % 26;
    letter = String.fromCharCode(65 + modulo) + letter;
    remaining = Math.floor((remaining - modulo) / 26);
  }

  return letter;
}


function APP_setDataValidationForA1Ranges_(sheet, a1Ranges, rule) {
  a1Ranges.forEach(function (a1Range) {
    sheet.getRange(a1Range).setDataValidation(rule);
  });
}


function onOpen() {
  LA_criarMenuLancamentos_();
  BP_criarMenuBoletos_();
}


function LA_criarMenuLancamentos_() {
  SpreadsheetApp.getUi()
    .createMenu('Lançamentos')
    .addItem('Registrar lançamento', 'registrarLancamento')
    .addItem('Excluir lançamentos marcados', 'excluirLancamentosMarcados')
    .addItem('Desfazer última exclusão', 'desfazerUltimaExclusao')
    .addSeparator()
    .addItem('Adicionar item às listas', 'adicionarItemAsListas')
    .addItem('Limpar formulário', 'limparFormulario')
    .addSeparator()
    .addItem('Validar estrutura da planilha', 'validarEstruturaLancamentos')
    .addItem('Backup agora', 'fazerBackupDrive')
    .addToUi();
}


function BP_criarMenuBoletos_() {
  SpreadsheetApp.getUi()
    .createMenu('Boletos')
    .addItem('Registrar boleto', 'registrarBoleto')
    .addItem('Limpar formulário', 'limparFormularioBoletos')
    .addSeparator()
    .addItem('Configurar / revisar tudo', 'setupBoletos')
    .addItem('Configurar validações', 'configurarValidacoesBoletos')
    .addItem('Reaplicar fórmulas automáticas', 'reaplicarFormulasBoletos')
    .addSeparator()
    .addItem('Adicionar item às listas', 'adicionarItemAsListasBoletos')
    .addItem('Verificar estrutura da planilha', 'verificarEstruturaBoletos')
    .addToUi();
}


/****************************************************
 * ==================================================
 * MÓDULO 1 — LANÇAMENTOS
 * ==================================================
 ****************************************************/


/****************************************************
 * CONFIGURAÇÕES — LANÇAMENTOS
 ****************************************************/

const LA_ABA_LANCAMENTOS = 'LANÇAMENTOS';
const LA_ABA_CONFIG = 'CONFIG';

const LA_LINHA_CABECALHO = 19;
const LA_PRIMEIRA_LINHA_DADOS = 20;
const LA_ULTIMA_LINHA_DADOS = 5000;
const LA_TOTAL_LINHAS_DADOS = LA_ULTIMA_LINHA_DADOS - LA_PRIMEIRA_LINHA_DADOS + 1;
const LA_TOTAL_LINHAS_LISTAS = LA_ULTIMA_LINHA_DADOS - 1;

const LA_COLUNAS_LANCAMENTO = Object.freeze({
  ID: 1,
  DATA: 2,         // B
  MES: 3,          // C
  MES_NUM: 4,      // D — fórmula, não sobrescrever
  UNIDADE: 5,      // E
  TIPO: 6,         // F
  CATEGORIA: 7,    // G
  SUBCATEGORIA: 8, // H
  VALOR: 9,        // I
  FORMA_PGTO: 10,  // J
  STATUS: 11,      // K
  DATA_PGTO: 12,   // L
  OBSERVACOES: 13, // M
  CRIADO_EM: 14,   // N
  EXCLUIR: 15,     // O
  CHAVE_TECNICA: 16, // P
  CLASSE_TIPO: 90, // CL — fórmula, não sobrescrever
  LISTA_TIPOS: 18,       // R
  LISTA_CATEGORIAS: 19,  // S
  LISTA_STATUS: 20,      // T
  LISTA_FORMAS_PGTO: 21  // U
});

const LA_INTERVALOS_FORMULARIO = Object.freeze({
  AREA: 'frmAreaFormulario',
  DATA: 'frmDataLancamento',
  MES_COMPETENCIA: 'frmMesCompetencia',
  UNIDADE: 'frmUnidade',
  TIPO: 'frmTipoLancamento',
  CATEGORIA: 'frmCategoria',
  SUBCATEGORIA: 'frmSubcategoria',
  VALOR: 'frmValor',
  FORMA_PGTO: 'frmFormaPgto',
  STATUS: 'frmStatus',
  DATA_PGTO: 'frmDataPgto',
  OBSERVACOES: 'frmObservacoes'
});

const LA_AREA_FORMULARIO_A1 = 'A4:C14';

const LA_FORMULARIO_A1 = Object.freeze({
  frmDataLancamento: 'B4',
  frmMesCompetencia: 'B5',
  frmUnidade: 'B6',
  frmTipoLancamento: 'B7',
  frmCategoria: 'B8',
  frmSubcategoria: 'B9',
  frmValor: 'B10',
  frmFormaPgto: 'B11',
  frmStatus: 'B12',
  frmDataPgto: 'B13',
  frmObservacoes: 'B14'
});

const LA_INTERVALOS_OBRIGATORIOS_FORMULARIO = [
  LA_INTERVALOS_FORMULARIO.DATA,
  LA_INTERVALOS_FORMULARIO.UNIDADE,
  LA_INTERVALOS_FORMULARIO.TIPO,
  LA_INTERVALOS_FORMULARIO.CATEGORIA,
  LA_INTERVALOS_FORMULARIO.VALOR
];

const LA_INTERVALOS_DADOS_FORMULARIO = [
  LA_INTERVALOS_FORMULARIO.DATA,
  LA_INTERVALOS_FORMULARIO.MES_COMPETENCIA,
  LA_INTERVALOS_FORMULARIO.UNIDADE,
  LA_INTERVALOS_FORMULARIO.TIPO,
  LA_INTERVALOS_FORMULARIO.CATEGORIA,
  LA_INTERVALOS_FORMULARIO.SUBCATEGORIA,
  LA_INTERVALOS_FORMULARIO.VALOR,
  LA_INTERVALOS_FORMULARIO.FORMA_PGTO,
  LA_INTERVALOS_FORMULARIO.STATUS,
  LA_INTERVALOS_FORMULARIO.DATA_PGTO,
  LA_INTERVALOS_FORMULARIO.OBSERVACOES
];

const LA_CLASSES_RELATORIO = [
  'Faturamento',
  'Despesa',
  'Ignorar'
];

const LA_ABAS_LOJAS_DRE = [
  'CENTRO',
  'C. DO MAR',
  'OUTLET',
  'TNC.'
];

const LA_CLASSES_DRE_AUTOMATICAS = [
  'Faturamento',
  'Despesa'
];


/****************************************************
 * REGISTRAR LANÇAMENTO
 ****************************************************/

function registrarLancamento() {
  LA_executarComErroAmigavel_('Registrar lançamento', function () {
    const ss = SpreadsheetApp.getActive();
    const ui = SpreadsheetApp.getUi();
    const sheet = LA_obterAbaObrigatoria_(ss, LA_ABA_LANCAMENTOS);
    const namedRangeMap = LA_mapearIntervalosNomeados_(
      ss,
      LA_INTERVALOS_OBRIGATORIOS_FORMULARIO.concat(LA_INTERVALOS_DADOS_FORMULARIO)
    );

    const intervalosAusentes = LA_listarIntervalosNomeadosAusentes_(
      ss,
      LA_INTERVALOS_OBRIGATORIOS_FORMULARIO,
      namedRangeMap
    );

    if (intervalosAusentes.length > 0) {
      ui.alert(
        'Configuração incompleta',
        'Não encontrei estes intervalos nomeados obrigatórios:\n\n' +
          intervalosAusentes.join('\n') +
          '\n\nCorrija os nomes no Google Sheets antes de registrar.',
        ui.ButtonSet.OK
      );
      return;
    }

    const lancamento = LA_lerLancamentoDoFormulario_(ss, namedRangeMap);
    const configIndex = LA_criarIndiceConfigLancamentos_(LA_getConfigValues_(ss));
    const erroValidacao = LA_validarLancamento_(lancamento, configIndex);

    if (erroValidacao) {
      ui.alert('Revise o lançamento', erroValidacao, ui.ButtonSet.OK);
      return;
    }

    lancamento.mes = lancamento.mesManual || LA_mesPorData_(lancamento.data);

    const nextRow = LA_executarComLockDocumento_(function () {
      const row = LA_proximaLinha_(sheet);

      if (!row) {
        ui.alert(
          'Tabela cheia',
          'A tabela de lançamentos está cheia. Expanda o intervalo A20:Q5000.',
          ui.ButtonSet.OK
        );
        return null;
      }

      const duplicateKey = LA_makeKey_(
        lancamento.data,
        lancamento.unidade,
        lancamento.tipo,
        lancamento.categoria,
        lancamento.subcategoria,
        lancamento.valor
      );

      if (LA_existeChaveTecnicaLancamento_(sheet, duplicateKey)) {
        const response = ui.alert(
          'Possível duplicidade',
          'Já existe lançamento com a mesma data, unidade, tipo, categoria, subcategoria e valor. Deseja registrar mesmo assim?',
          ui.ButtonSet.YES_NO
        );

        if (response !== ui.Button.YES) return null;
      }

      lancamento.chaveTecnica = duplicateKey;

      LA_escreverLancamento_(sheet, row, lancamento);
      SpreadsheetApp.flush();

      return row;
    });

    if (!nextRow) return;

    LA_limparFormulario_();

    ss.toast(
      'Lançamento registrado com sucesso na linha ' + nextRow + '.',
      'LANÇAMENTOS',
      5
    );
  });
}


function LA_lerLancamentoDoFormulario_(spreadsheet, namedRangeMap) {
  const values = LA_lerValoresIntervalosNomeados_(
    spreadsheet || SpreadsheetApp.getActive(),
    LA_INTERVALOS_DADOS_FORMULARIO,
    namedRangeMap
  );
  const dataRaw = values[LA_INTERVALOS_FORMULARIO.DATA];
  const dataPgtoRaw = values[LA_INTERVALOS_FORMULARIO.DATA_PGTO];

  return {
    data: LA_parseData_(dataRaw),
    dataRaw: dataRaw,
    mesManual: LA_normalizarTexto_(values[LA_INTERVALOS_FORMULARIO.MES_COMPETENCIA]),
    unidade: LA_normalizarTexto_(values[LA_INTERVALOS_FORMULARIO.UNIDADE]),
    tipo: LA_normalizarTexto_(values[LA_INTERVALOS_FORMULARIO.TIPO]),
    categoria: LA_normalizarTexto_(values[LA_INTERVALOS_FORMULARIO.CATEGORIA]),
    subcategoria: LA_normalizarTexto_(values[LA_INTERVALOS_FORMULARIO.SUBCATEGORIA]),
    valor: LA_parseValor_(values[LA_INTERVALOS_FORMULARIO.VALOR]),
    formaPgto: LA_normalizarTexto_(values[LA_INTERVALOS_FORMULARIO.FORMA_PGTO]),
    status: LA_normalizarTexto_(values[LA_INTERVALOS_FORMULARIO.STATUS]) || 'Pendente',
    dataPgto: LA_parseData_(dataPgtoRaw),
    dataPgtoRaw: dataPgtoRaw,
    obs: LA_normalizarTexto_(values[LA_INTERVALOS_FORMULARIO.OBSERVACOES]),
    criadoEm: new Date()
  };
}


function LA_validarLancamento_(lancamento, configIndex) {
  const erros = [];

  if (!lancamento.data) {
    erros.push(
      LA_campoPreenchido_(lancamento.dataRaw)
        ? 'A Data precisa estar em um formato reconhecido, como 31/05/2026.'
        : 'Informe uma Data válida.'
    );
  }

  if (!lancamento.unidade) {
    erros.push('Informe a Unidade.');
  }

  if (!lancamento.tipo) {
    erros.push('Informe o Tipo.');
  }

  if (!lancamento.categoria) {
    erros.push('Informe a Categoria.');
  }

  if (!Number.isFinite(lancamento.valor)) {
    erros.push('Informe um Valor numérico válido.');
  }

  if (lancamento.formaPgto && LA_equalsTexto_(lancamento.formaPgto, 'Taxas Getnet')) {
    erros.push(
      '"Taxas Getnet" não é uma Forma de Pagamento válida.\n' +
      'Registre como: Despesa > Despesas Financeiras > Taxas Getnet.'
    );
  }

  if (LA_campoPreenchido_(lancamento.dataPgtoRaw) && !lancamento.dataPgto) {
    erros.push('Informe uma Data de Pagamento válida ou deixe o campo em branco.');
  }

  if (
    lancamento.tipo &&
    lancamento.categoria &&
    !LA_configPossuiTipoCategoria_(configIndex, lancamento.tipo, lancamento.categoria)
  ) {
    erros.push('Categoria inválida para o Tipo selecionado.');
  }

  if (
    lancamento.tipo &&
    lancamento.categoria &&
    lancamento.subcategoria &&
    !LA_configPossuiTipoCategoriaSubcategoria_(
      configIndex,
      lancamento.tipo,
      lancamento.categoria,
      lancamento.subcategoria
    )
  ) {
    erros.push('Subcategoria inválida para a combinação Tipo + Categoria selecionada.');
  }

  return erros.join('\n');
}


function LA_escreverLancamento_(sheet, row, lancamento) {
  const chaveTecnica = lancamento.chaveTecnica || LA_makeKey_(
    lancamento.data,
    lancamento.unidade,
    lancamento.tipo,
    lancamento.categoria,
    lancamento.subcategoria,
    lancamento.valor
  );

  // B:C — data e mês (A e D são fórmulas, não sobrescrever)
  sheet
    .getRange(row, LA_COLUNAS_LANCAMENTO.DATA, 1, 2)
    .setValues([[lancamento.data, lancamento.mes]]);

  // E:P — dados + chave em uma única chamada (col 5 a 16)
  sheet
    .getRange(
      row,
      LA_COLUNAS_LANCAMENTO.UNIDADE,
      1,
      LA_COLUNAS_LANCAMENTO.CHAVE_TECNICA - LA_COLUNAS_LANCAMENTO.UNIDADE + 1
    )
    .setValues([[
      lancamento.unidade,
      lancamento.tipo,
      lancamento.categoria,
      lancamento.subcategoria,
      lancamento.valor,
      lancamento.formaPgto,
      lancamento.status,
      lancamento.dataPgto || '',
      lancamento.obs,
      lancamento.criadoEm,
      false,
      chaveTecnica
    ]]);

  LA_aplicarFormatosLinhaLancamento_(sheet, row);
}


function LA_aplicarFormatosLinhaLancamento_(sheet, row) {
  sheet
    .getRangeList([
      APP_a1Range_(row, LA_COLUNAS_LANCAMENTO.DATA),
      APP_a1Range_(row, LA_COLUNAS_LANCAMENTO.DATA_PGTO)
    ])
    .setNumberFormat(APP_FORMATOS.DATA);

  sheet
    .getRange(row, LA_COLUNAS_LANCAMENTO.VALOR)
    .setNumberFormat(APP_FORMATOS.MOEDA_BRL);

  sheet
    .getRange(row, LA_COLUNAS_LANCAMENTO.CRIADO_EM)
    .setNumberFormat(APP_FORMATOS.DATA_HORA);
}


/****************************************************
 * EXCLUIR LANÇAMENTOS MARCADOS
 ****************************************************/

function excluirLancamentosMarcados() {
  LA_executarComErroAmigavel_('Excluir lançamentos marcados', function () {
    const ss = SpreadsheetApp.getActive();
    const ui = SpreadsheetApp.getUi();
    const sheet = LA_obterAbaObrigatoria_(ss, LA_ABA_LANCAMENTOS);

    const checks = sheet
      .getRange(
        LA_PRIMEIRA_LINHA_DADOS,
        LA_COLUNAS_LANCAMENTO.EXCLUIR,
        LA_TOTAL_LINHAS_DADOS,
        1
      )
      .getValues();

    const rowsToClear = [];

    checks.forEach(function (row, index) {
      if (row[0] === true) {
        rowsToClear.push(LA_PRIMEIRA_LINHA_DADOS + index);
      }
    });

    if (rowsToClear.length === 0) {
      ui.alert('Nenhum lançamento marcado para exclusão.');
      return;
    }

    const response = ui.alert(
      'Confirmar exclusão',
      'Deseja excluir ' + rowsToClear.length + ' lançamento(s) marcado(s)?',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) return;

    LA_executarComLockDocumento_(function () {
      // Salva backup das linhas antes de apagar (para desfazer)
      LA_salvarBackupExclusao_(sheet, rowsToClear);

      const rangesParaLimpar = [];
      const rangesParaDesmarcar = [];

      LA_agruparLinhasConsecutivas_(rowsToClear).forEach(function (grupo) {
        rangesParaLimpar.push(
          APP_a1Range_(grupo.inicio, LA_COLUNAS_LANCAMENTO.DATA, grupo.quantidade, 2),
          APP_a1Range_(grupo.inicio, LA_COLUNAS_LANCAMENTO.UNIDADE, grupo.quantidade, LA_COLUNAS_LANCAMENTO.CRIADO_EM - LA_COLUNAS_LANCAMENTO.UNIDADE + 1),
          APP_a1Range_(grupo.inicio, LA_COLUNAS_LANCAMENTO.CHAVE_TECNICA, grupo.quantidade, 1)
        );

        rangesParaDesmarcar.push(
          APP_a1Range_(grupo.inicio, LA_COLUNAS_LANCAMENTO.EXCLUIR, grupo.quantidade, 1)
        );
      });

      sheet.getRangeList(rangesParaLimpar).clearContent();
      sheet.getRangeList(rangesParaDesmarcar).setValue(false);

      SpreadsheetApp.flush();
    });

    ss.toast(
      rowsToClear.length + ' lançamento(s) excluído(s).',
      'LANÇAMENTOS',
      5
    );
  });
}


/****************************************************
 * LIMPAR FORMULÁRIO — LANÇAMENTOS
 ****************************************************/

function limparFormulario() {
  LA_executarComErroAmigavel_('Limpar formulário', LA_limparFormulario_);
}


function LA_limparFormulario_() {
  const ss = SpreadsheetApp.getActive();
  LA_obterAbaObrigatoria_(ss, LA_ABA_LANCAMENTOS)
    .getRange(LA_AREA_FORMULARIO_A1)
    .clearContent();
}


/****************************************************
 * ADICIONAR ITENS ÀS LISTAS — LANÇAMENTOS
 ****************************************************/

function adicionarItemAsListas() {
  LA_executarComErroAmigavel_('Adicionar item às listas', function () {
    const ss = SpreadsheetApp.getActive();
    const ui = SpreadsheetApp.getUi();
    const cfg = LA_obterAbaObrigatoria_(ss, LA_ABA_CONFIG);

    LA_ensureConfigClassHeader_();

    const data = LA_getConfigRowsFromSheet_(cfg);

    const tiposExistentes = LA_uniquePreserveOrder_(
      data.map(function (row) { return row[0]; }).filter(Boolean)
    );

    const tipoDigitado = LA_promptTextComExemplos_(
      'Adicionar item às listas',
      'Digite o Tipo exatamente como deve aparecer.',
      [
        'Exemplos existentes:',
        tiposExistentes.slice(0, 10).join(', '),
        '',
        'Se for um Tipo novo, depois você escolherá como ele entra nos relatórios.'
      ]
    );

    if (!tipoDigitado) return;

    const tipo = LA_matchExistingText_(tipoDigitado, tiposExistentes);

    const categoriasDoTipo = LA_uniquePreserveOrder_(
      data
        .filter(function (row) { return LA_equalsTexto_(row[0], tipo); })
        .map(function (row) { return row[1]; })
        .filter(Boolean)
    );

    const todasCategorias = LA_uniquePreserveOrder_(
      data.map(function (row) { return row[1]; }).filter(Boolean)
    );

    const categoriaDigitada = LA_promptTextComExemplos_(
      'Adicionar item às listas',
      'Digite a Categoria.',
      categoriasDoTipo.length
        ? [
            'Exemplos para "' + tipo + '":',
            categoriasDoTipo.slice(0, 12).join(', ')
          ]
        : [
            'Esse Tipo ainda não possui categorias cadastradas.',
            'Exemplos gerais:',
            todasCategorias.slice(0, 12).join(', '),
            '',
            'Você pode criar uma Categoria nova se necessário.'
          ]
    );

    if (!categoriaDigitada) return;

    const categoria = LA_matchExistingText_(
      categoriaDigitada,
      categoriasDoTipo.length ? categoriasDoTipo : todasCategorias
    );

    const subcategoriasDaCategoria = LA_uniquePreserveOrder_(
      data
        .filter(function (row) {
          return LA_equalsTexto_(row[0], tipo) && LA_equalsTexto_(row[1], categoria);
        })
        .map(function (row) { return row[2]; })
        .filter(Boolean)
    );

    const subcategoriaDigitada = LA_promptTextComExemplos_(
      'Adicionar item às listas',
      'Digite a Subcategoria / Fornecedor, se houver.',
      subcategoriasDaCategoria.length
        ? [
            'Exemplos para "' + categoria + '":',
            subcategoriasDaCategoria.slice(0, 12).join(', '),
            '',
            'Se não houver subcategoria, deixe em branco.'
          ]
        : [
            'Ainda não há subcategorias cadastradas para essa combinação.',
            'Exemplos comuns:',
            'Zeiss, Pix, Internet, Tráfego pago, Energia, Fornecedor específico',
            '',
            'Se não houver subcategoria, deixe em branco.'
          ]
    );

    const subcategoria = subcategoriaDigitada
      ? LA_matchExistingText_(subcategoriaDigitada, subcategoriasDaCategoria)
      : '';

    const linhaDuplicada = data.find(function (row) {
      return (
        LA_equalsTexto_(row[0], tipo) &&
        LA_equalsTexto_(row[1], categoria) &&
        LA_equalsTexto_(row[2], subcategoria)
      );
    });

    const classeExistente = data
      .filter(function (row) { return LA_equalsTexto_(row[0], tipo); })
      .map(function (row) { return LA_matchClasseRelatorio_(row[5]); })
      .find(Boolean);

    if (linhaDuplicada) {
      const classeDuplicada = LA_matchClasseRelatorio_(linhaDuplicada[5]) || classeExistente;

      if (!classeDuplicada) {
        ui.alert('Esse item já existe na CONFIG, mas está sem ClasseRelatorio válida.');
        return;
      }

      const refletirExistente = ui.alert(
        'Item já existe na CONFIG',
        'Esse Tipo + Categoria + Subcategoria já existe na CONFIG.\n\n' +
          'Deseja apenas refletir esse item nas lojas/DRE agora?',
        ui.ButtonSet.YES_NO
      );

      if (refletirExistente !== ui.Button.YES) return;

      LA_refletirItemConfigNasLojas_(ss, ui, {
        tipo: tipo,
        categoria: categoria,
        subcategoria: subcategoria,
        classeRelatorio: classeDuplicada,
        origem: 'CONFIG existente',
        criadoEm: new Date()
      });
      return;
    }

    const classe = classeExistente || LA_promptClasseRelatorio_();

    if (!classe) return;

    const confirmar = ui.alert(
      'Confirmar novo item',
      'Adicionar este item?\n\n' +
        'Tipo: ' + tipo + '\n' +
        'Categoria: ' + categoria + '\n' +
        'Subcategoria: ' + (subcategoria || '(sem subcategoria)') + '\n' +
        'Classe no relatório: ' + classe,
      ui.ButtonSet.YES_NO
    );

    if (confirmar !== ui.Button.YES) return;

    const item = {
      tipo: tipo,
      categoria: categoria,
      subcategoria: subcategoria,
      classeRelatorio: classe,
      origem: 'Menu Apps Script',
      criadoEm: new Date()
    };

    LA_executarComLockDocumento_(function () {
      const nextRow = LA_firstBlankRow_(cfg, 1, 2);

      cfg.getRange(nextRow, 1, 1, 6).setValues([[
        item.tipo,
        item.categoria,
        item.subcategoria,
        item.origem,
        'Dropdown dependente',
        item.classeRelatorio
      ]]);

      LA_rebuildListasVisiveis_();
      SpreadsheetApp.flush();
    });

    LA_refletirItemConfigNasLojas_(ss, ui, item);

    ss.toast(
      'Item adicionado: ' +
        item.tipo +
        ' > ' +
        item.categoria +
        (item.subcategoria ? ' > ' + item.subcategoria : '') +
        ' | Classe: ' +
        item.classeRelatorio,
      'Listas atualizadas',
      5
    );
  });
}


function LA_refletirItemConfigNasLojas_(spreadsheet, ui, item) {
  const lojasSelecionadas = LA_perguntarLojasParaNovoItem_(item);

  if (lojasSelecionadas === null) {
    LA_mostrarResultadoInsercaoLojas_(ui, {
      inseridas: [],
      ignoradas: ['Seleção de lojas cancelada; item criado apenas na CONFIG.'],
      pendentes: []
    });
    return;
  }

  const resultadoLojas = LA_executarComLockDocumento_(function () {
    const resultado = LA_inserirItemNasLojasSelecionadas_(spreadsheet, item, lojasSelecionadas);
    SpreadsheetApp.flush();
    return resultado;
  });

  LA_mostrarResultadoInsercaoLojas_(ui, resultadoLojas);
}


function LA_perguntarLojasParaNovoItem_(item) {
  const ui = SpreadsheetApp.getUi();
  const abasLojas = LA_obterAbasLojasDre_();

  if (!LA_CLASSES_DRE_AUTOMATICAS.some(function (classe) {
    return LA_equalsTexto_(classe, item.classeRelatorio);
  })) {
    ui.alert(
      'Reflexo nas lojas',
      'A classe "' + item.classeRelatorio + '" foi salva apenas na CONFIG.\n\n' +
        'O script só insere automaticamente nas lojas itens de Faturamento ou Despesa.',
      ui.ButtonSet.OK
    );
    return [];
  }

  const opcoes = [
    'Digite "todas" para inserir em todas as lojas.',
    'Digite "nenhuma" para salvar apenas na CONFIG.',
    'Ou digite os números separados por vírgula:',
  ].concat(
    abasLojas.map(function (loja, index) {
      return (index + 1) + ' - ' + loja;
    })
  );

  const response = ui.prompt(
    'Refletir item nas lojas',
    'Em quais lojas este item deve aparecer no Resumo Financeiro/DRE?\n\n' +
      'Item: ' + item.categoria +
      (item.subcategoria ? ' > ' + item.subcategoria : '') +
      '\nClasse: ' + item.classeRelatorio +
      '\n\n' +
      opcoes.join('\n'),
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return null;

  const raw = LA_normalizarTexto_(response.getResponseText());
  const normalized = LA_chaveTexto_(raw).replace(/[^a-z0-9]/g, '');

  if (!raw || normalized === 'nenhuma' || normalized === 'nao') {
    return [];
  }

  if (normalized === 'todas' || normalized === 'todos') {
    return abasLojas.slice();
  }

  const selected = {};

  LA_tokenizarSelecaoLojasDre_(raw).forEach(function (part) {
    const token = LA_normalizarTexto_(part);
    const index = Number(token);

    if (Number.isInteger(index) && index >= 1 && index <= abasLojas.length) {
      selected[abasLojas[index - 1]] = true;
      return;
    }

    abasLojas.forEach(function (loja) {
      if (LA_equalsTexto_(loja, token)) {
        selected[loja] = true;
      }
    });
  });

  const lojas = Object.keys(selected);

  if (!lojas.length) {
    ui.alert(
      'Lojas não reconhecidas',
      'Não consegui identificar as lojas informadas. O item será salvo apenas na CONFIG.',
      ui.ButtonSet.OK
    );
  }

  return lojas;
}


function LA_tokenizarSelecaoLojasDre_(raw) {
  return LA_normalizarTexto_(raw)
    .replace(/\r?\n/g, ',')
    .replace(/\s+e\s+/gi, ',')
    .replace(/\s+and\s+/gi, ',')
    .split(/[;,]/)
    .map(LA_normalizarTexto_)
    .filter(Boolean);
}


function LA_obterAbasLojasDre_() {
  return LA_ABAS_LOJAS_DRE.slice();
}


function LA_inserirItemNasLojasSelecionadas_(spreadsheet, item, lojas) {
  const resultado = {
    inseridas: [],
    ignoradas: [],
    pendentes: []
  };

  if (!lojas.length) {
    resultado.ignoradas.push('Nenhuma loja selecionada; item salvo apenas na CONFIG.');
    return resultado;
  }

  lojas.forEach(function (nomeLoja) {
    const sheet = spreadsheet.getSheetByName(nomeLoja);

    if (!sheet) {
      resultado.pendentes.push(nomeLoja + ': aba não encontrada.');
      console.warn(JSON.stringify({
        module: 'LANÇAMENTOS_DRE',
        action: 'sheet_not_found',
        loja: nomeLoja,
        item: item,
        timestamp: new Date().toISOString()
      }));
      return;
    }

    const status = LA_inserirItemNaLoja_(sheet, item);

    if (status.status === 'inserted') {
      resultado.inseridas.push(nomeLoja + ': linha ' + status.row + '.');
      LA_logResultadoInsercaoLoja_(nomeLoja, item, 'INSERIDO', status.row, '');
      return;
    }

    if (status.status === 'exists') {
      resultado.ignoradas.push(nomeLoja + ': item já existia.');
      LA_logResultadoInsercaoLoja_(nomeLoja, item, 'DUPLICADO', null, 'Item já existe na loja.');
      return;
    }

    resultado.pendentes.push(nomeLoja + ': ' + status.message);
    LA_logResultadoInsercaoLoja_(nomeLoja, item, 'NAO_INSERIDO', null, status.message);
  });

  console.info(JSON.stringify({
    module: 'LANÇAMENTOS_DRE',
    action: 'insert_item_result',
    item: item,
    resultado: resultado,
    timestamp: new Date().toISOString()
  }));

  return resultado;
}


function LA_inserirItemNaLoja_(sheet, item) {
  const context = LA_lerContextoDreLoja_(sheet);

  if (LA_itemJaExisteNaLoja_(context, item)) {
    return { status: 'exists' };
  }

  const location = LA_localizarSecaoDre_(context, item);

  if (!location) {
    return {
      status: 'manual',
      message: 'não encontrei linha-modelo compatível por Categoria ou ClasseRelatorio.'
    };
  }

  const fillColumns = LA_resolverColunasPreenchimentoDre_(context, location.sourceRow, item);

  if (!fillColumns.label && !fillColumns.categoria && !fillColumns.subcategoria) {
    return {
      status: 'manual',
      message: 'encontrei linha-modelo, mas não identifiquei onde escrever Categoria/Subcategoria.'
    };
  }

  const targetRow = LA_copiarLinhaModeloDre_(sheet, location.sourceRow, context.maxColumns, fillColumns);

  LA_preencherLinhaDre_(sheet, targetRow, fillColumns, item);

  console.info(JSON.stringify({
    module: 'LANÇAMENTOS_DRE',
    action: 'insert_item_sheet',
    loja: sheet.getName(),
    sourceRow: location.sourceRow,
    targetRow: targetRow,
    reason: location.reason,
    item: item,
    timestamp: new Date().toISOString()
  }));

  return {
    status: 'inserted',
    row: targetRow
  };
}


function LA_lerContextoDreLoja_(sheet) {
  const range = sheet.getDataRange();
  const values = range.getDisplayValues();

  return {
    sheet: sheet,
    values: values,
    maxColumns: range.getNumColumns(),
    columns: LA_inferirColunasDre_(values)
  };
}


function LA_inferirColunasDre_(values) {
  const columns = {
    tipo: 0,
    categoria: 0,
    subcategoria: 0,
    classeRelatorio: 0
  };
  const rowsToScan = Math.min(values.length, 30);

  for (let row = 0; row < rowsToScan; row++) {
    values[row].forEach(function (cell, index) {
      const key = LA_chaveTexto_(cell).replace(/[^a-z0-9]/g, '');

      if (!columns.tipo && key === 'tipo') {
        columns.tipo = index + 1;
      }

      if (!columns.categoria && key === 'categoria') {
        columns.categoria = index + 1;
      }

      if (!columns.subcategoria && (
        key === 'subcategoria' ||
        key === 'fornecedor' ||
        key === 'subcategoriafornecedor'
      )) {
        columns.subcategoria = index + 1;
      }

      if (!columns.classeRelatorio && (
        key === 'classe' ||
        key === 'classerelatorio' ||
        key === 'classedorelatorios'
      )) {
        columns.classeRelatorio = index + 1;
      }
    });
  }

  return columns;
}


function LA_itemJaExisteNaLoja_(context, item) {
  const columns = context.columns;

  return context.values.some(function (row) {
    if (LA_linhaDreEhTotalOuSecao_(row)) return false;

    if (!columns.categoria && !columns.subcategoria) {
      const itemPrincipalExiste = item.subcategoria
        ? LA_linhaContemTextoDre_(row, item.subcategoria)
        : LA_linhaContemTextoDre_(row, item.categoria);
      const classeOkSemColuna = !columns.classeRelatorio ||
        LA_equalsTexto_(row[columns.classeRelatorio - 1], item.classeRelatorio);

      return itemPrincipalExiste && classeOkSemColuna;
    }

    const categoriaOk = LA_colunaOuLinhaContemDre_(row, columns.categoria, item.categoria);
    const subcategoriaOk = !item.subcategoria ||
      LA_colunaOuLinhaContemDre_(row, columns.subcategoria, item.subcategoria);
    const classeOk = !columns.classeRelatorio ||
      LA_equalsTexto_(row[columns.classeRelatorio - 1], item.classeRelatorio);

    return categoriaOk && subcategoriaOk && classeOk;
  });
}


function LA_localizarSecaoDre_(context, item) {
  const columns = context.columns;
  const sameCategoryRows = [];
  const sameClassRows = [];

  context.values.forEach(function (row, index) {
    if (LA_linhaDreEhTotalOuSecao_(row)) return;

    const hasText = row.some(function (cell) {
      return LA_normalizarTexto_(cell) !== '';
    });

    if (!hasText) return;

    if (LA_colunaOuLinhaContemDre_(row, columns.categoria, item.categoria)) {
      sameCategoryRows.push(index + 1);
      return;
    }

    if (
      LA_colunaOuLinhaContemDre_(row, columns.classeRelatorio, item.classeRelatorio) ||
      LA_linhaContemTextoDre_(row, item.classeRelatorio) ||
      LA_linhaContemTextoDre_(row, item.tipo)
    ) {
      sameClassRows.push(index + 1);
    }
  });

  if (sameCategoryRows.length) {
    return {
      sourceRow: sameCategoryRows[sameCategoryRows.length - 1],
      reason: 'same_category'
    };
  }

  if (sameClassRows.length) {
    return {
      sourceRow: sameClassRows[sameClassRows.length - 1],
      reason: 'same_class_or_type'
    };
  }

  return null;
}


function LA_resolverColunasPreenchimentoDre_(context, sourceRow, item) {
  const row = context.values[sourceRow - 1] || [];
  const columns = {
    tipo: context.columns.tipo,
    categoria: context.columns.categoria,
    subcategoria: context.columns.subcategoria,
    classeRelatorio: context.columns.classeRelatorio,
    label: 0
  };

  const categoriaMatch = LA_encontrarColunaTextoNaLinhaDre_(row, item.categoria);
  const subcategoriaMatch = LA_encontrarColunaTextoNaLinhaDre_(row, item.subcategoria);

  if (!columns.categoria && !columns.subcategoria) {
    columns.label = subcategoriaMatch || categoriaMatch || LA_encontrarPrimeiraColunaTextoDre_(row);
    return columns;
  }

  // Mesmo quando existem colunas técnicas (Categoria/Subcategoria/Classe),
  // a DRE costuma ter uma coluna descritiva visível (ex.: "Conta") que herdou
  // o texto da linha-modelo (ex.: "Outros juros") e precisa ser sobrescrita
  // com o nome do item criado.
  columns.label = LA_encontrarColunaRotuloVisivelDre_(row, columns);

  return columns;
}


function LA_encontrarColunaRotuloVisivelDre_(row, columns) {
  const tecnicas = {};

  [columns.tipo, columns.categoria, columns.subcategoria, columns.classeRelatorio]
    .forEach(function (column) {
      if (column) tecnicas[column] = true;
    });

  for (let index = 0; index < row.length; index++) {
    const column = index + 1;

    if (tecnicas[column]) continue;

    const value = LA_normalizarTexto_(row[index]);

    if (value && !LA_valorTextoPareceNumeroDre_(value)) {
      return column;
    }
  }

  return 0;
}


function LA_copiarLinhaModeloDre_(sheet, sourceRow, maxColumns, columns) {
  sheet.insertRowsAfter(sourceRow, 1);

  const targetRow = sourceRow + 1;
  const sourceRange = sheet.getRange(sourceRow, 1, 1, maxColumns);
  const targetRange = sheet.getRange(targetRow, 1, 1, maxColumns);

  sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  sheet.setRowHeight(targetRow, sheet.getRowHeight(sourceRow));

  LA_limparValoresNaoFormulaDaLinhaDre_(sheet, targetRow, maxColumns, columns);

  return targetRow;
}


function LA_limparValoresNaoFormulaDaLinhaDre_(sheet, row, maxColumns, columns) {
  const formulas = sheet.getRange(row, 1, 1, maxColumns).getFormulas()[0];
  const values = sheet.getRange(row, 1, 1, maxColumns).getValues()[0];
  const protectedColumns = LA_colunasTextoDre_(columns);
  const rangesToClear = [];

  values.forEach(function (value, index) {
    const column = index + 1;

    if (protectedColumns[column] || formulas[index]) return;

    if (typeof value === 'number' || value instanceof Date) {
      rangesToClear.push(APP_a1Range_(row, column));
    }
  });

  if (rangesToClear.length) {
    sheet.getRangeList(rangesToClear).clearContent();
  }
}


function LA_preencherLinhaDre_(sheet, row, columns, item) {
  // Regra do rótulo visível: usa a Subcategoria; se não houver, usa a Categoria.
  const rotuloVisivel = item.subcategoria || item.categoria;

  // Sempre sobrescreve a coluna descritiva visível com o nome do item criado,
  // para a nova linha não continuar exibindo o texto da linha-modelo copiada.
  if (columns.label) {
    sheet.getRange(row, columns.label).setValue(rotuloVisivel);
  }

  if (columns.tipo) {
    sheet.getRange(row, columns.tipo).setValue(item.tipo);
  }

  if (columns.categoria) {
    sheet.getRange(row, columns.categoria).setValue(item.categoria);
  }

  if (columns.subcategoria) {
    sheet.getRange(row, columns.subcategoria).setValue(item.subcategoria || '');
  }

  if (columns.classeRelatorio) {
    sheet.getRange(row, columns.classeRelatorio).setValue(item.classeRelatorio);
  }
}


function LA_colunasTextoDre_(columns) {
  const protectedColumns = {};

  [
    columns.tipo,
    columns.categoria,
    columns.subcategoria,
    columns.classeRelatorio,
    columns.label
  ].forEach(function (column) {
    if (column) {
      protectedColumns[column] = true;
    }
  });

  return protectedColumns;
}


function LA_encontrarColunaTextoNaLinhaDre_(row, expected) {
  const expectedKey = LA_chaveTexto_(expected);

  if (!expectedKey) return 0;

  for (let index = 0; index < row.length; index++) {
    const cellKey = LA_chaveTexto_(row[index]);

    if (cellKey && (cellKey === expectedKey || cellKey.indexOf(expectedKey) !== -1)) {
      return index + 1;
    }
  }

  return 0;
}


function LA_encontrarPrimeiraColunaTextoDre_(row) {
  for (let index = 0; index < row.length; index++) {
    const value = LA_normalizarTexto_(row[index]);

    if (value && !LA_valorTextoPareceNumeroDre_(value)) {
      return index + 1;
    }
  }

  return 0;
}


function LA_valorTextoPareceNumeroDre_(value) {
  return /^-?R?\$?\s*[\d.,%]+$/.test(LA_normalizarTexto_(value));
}


function LA_colunaOuLinhaContemDre_(row, column, expected) {
  if (!expected) return true;

  if (column && LA_equalsTexto_(row[column - 1], expected)) {
    return true;
  }

  return LA_linhaContemTextoDre_(row, expected);
}


function LA_logResultadoInsercaoLoja_(loja, item, status, row, motivo) {
  const payload = {
    action: 'INSERIR_ITEM_LOJA',
    loja: loja,
    categoria: item.categoria,
    subcategoria: item.subcategoria,
    classeRelatorio: item.classeRelatorio,
    status: status
  };

  if (row) {
    payload.row = row;
  }

  if (motivo) {
    payload.motivo = motivo;
  }

  if (status === 'INSERIDO') {
    console.info(JSON.stringify(payload));
    return;
  }

  console.warn(JSON.stringify(payload));
}


function LA_linhaContemTextoDre_(row, expected) {
  const expectedKey = LA_chaveTexto_(expected);

  if (!expectedKey) return true;

  return row.some(function (cell) {
    return LA_chaveTexto_(cell).indexOf(expectedKey) !== -1;
  });
}


function LA_linhaDreEhTotalOuSecao_(row) {
  const texto = LA_chaveTexto_(row.join(' '));

  return (
    texto.indexOf('total') !== -1 ||
    texto.indexOf('subtotal') !== -1 ||
    texto.indexOf('saldo') !== -1 ||
    texto.indexOf('resultado') !== -1
  );
}


function LA_mostrarResultadoInsercaoLojas_(ui, resultado) {
  if (!resultado) return;

  const linhas = [];

  if (resultado.inseridas.length) {
    linhas.push('Inseridas:\n' + resultado.inseridas.join('\n'));
  }

  if (resultado.ignoradas.length) {
    linhas.push('Ignoradas:\n' + resultado.ignoradas.join('\n'));
  }

  if (resultado.pendentes.length) {
    linhas.push('Revisão manual necessária:\n' + resultado.pendentes.join('\n'));
  }

  if (!linhas.length) return;

  ui.alert('Reflexo nas lojas/DRE', linhas.join('\n\n'), ui.ButtonSet.OK);
}


/****************************************************
 * CONFIGURAÇÃO DAS LISTAS — LANÇAMENTOS
 ****************************************************/

function LA_ensureConfigClassHeader_() {
  const ss = SpreadsheetApp.getActive();
  const cfg = LA_obterAbaObrigatoria_(ss, LA_ABA_CONFIG);

  if (cfg.getMaxColumns() < 6) {
    cfg.insertColumnsAfter(cfg.getMaxColumns(), 6 - cfg.getMaxColumns());
  }

  cfg.getRange('F1').setValue('ClasseRelatorio');

  const lastRow = Math.max(cfg.getLastRow(), 2);
  const range = cfg.getRange(2, 1, lastRow - 1, 6);
  const values = range.getValues();

  let mudou = false;

  values.forEach(function (row) {
    const tipo = LA_normalizarTexto_(row[0]);

    if (!tipo) return;

    const classeAtual = LA_matchClasseRelatorio_(row[5]);
    const classePeloTipo = LA_matchClasseRelatorio_(tipo);
    const novaClasse = classeAtual || classePeloTipo || 'Ignorar';

    if (row[5] !== novaClasse) {
      row[5] = novaClasse;
      mudou = true;
    }
  });

  if (mudou) {
    range.setValues(values);
  }

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(LA_CLASSES_RELATORIO, true)
    .setAllowInvalid(false)
    .build();

  cfg.getRange('F2:F800').setDataValidation(rule);

  LA_setNamedRangesSafe_([
    { name: 'rngConfigListas', range: cfg.getRange('A1:F800') },
    { name: 'cfgTipoClasseRelatorio', range: cfg.getRange('A1:F800') },
    { name: 'cfgTipoCategoriaSubcategoria', range: cfg.getRange('A1:F800') }
  ]);
}


function LA_rebuildListasVisiveis_() {
  const ss = SpreadsheetApp.getActive();
  const cfg = LA_obterAbaObrigatoria_(ss, LA_ABA_CONFIG);
  const lanc = LA_obterAbaObrigatoria_(ss, LA_ABA_LANCAMENTOS);
  const data = LA_getConfigRowsFromSheet_(cfg);

  const tipos = LA_uniquePreserveOrder_(
    data.map(function (row) { return row[0]; }).filter(Boolean)
  );

  const categorias = LA_uniquePreserveOrder_(
    data.map(function (row) { return row[1]; }).filter(Boolean)
  );

  // Formas de pagamento e Status são ranges fixos na planilha (U e T); não gerenciados aqui.

  lanc
    .getRangeList([
      APP_a1Range_(2, LA_COLUNAS_LANCAMENTO.LISTA_TIPOS, LA_TOTAL_LINHAS_LISTAS, 1),
      APP_a1Range_(2, LA_COLUNAS_LANCAMENTO.LISTA_CATEGORIAS, LA_TOTAL_LINHAS_LISTAS, 1)
    ])
    .clearContent();

  LA_escreverListaVertical_(lanc, LA_COLUNAS_LANCAMENTO.LISTA_TIPOS, tipos);
  LA_escreverListaVertical_(lanc, LA_COLUNAS_LANCAMENTO.LISTA_CATEGORIAS, categorias);

  LA_setNamedRangesSafe_([
    {
      name: 'rngTiposLancamento',
      range: lanc.getRange(2, LA_COLUNAS_LANCAMENTO.LISTA_TIPOS, LA_TOTAL_LINHAS_LISTAS, 1)
    },
    {
      name: 'rngCategoriasLancamento',
      range: lanc.getRange(2, LA_COLUNAS_LANCAMENTO.LISTA_CATEGORIAS, LA_TOTAL_LINHAS_LISTAS, 1)
    }
  ]);
}


function LA_promptClasseRelatorio_() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.prompt(
    'Classe do relatório',
    'Como esse novo Tipo deve entrar nos relatórios?\n\n' +
      'Opções:\n' +
      LA_CLASSES_RELATORIO.join('\n') +
      '\n\n' +
      'Dica:\n' +
      '- Use "Faturamento" para receitas.\n' +
      '- Use "Despesa" para custos/despesas.\n' +
      '- Use "Ignorar" se esse Tipo não deve afetar os relatórios.',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return '';

  const classe = LA_matchClasseRelatorio_(response.getResponseText());

  if (!classe) {
    ui.alert(
      'Classe inválida. Use exatamente uma destas opções:\n\n' +
      LA_CLASSES_RELATORIO.join('\n')
    );
    return '';
  }

  return classe;
}


/****************************************************
 * VALIDAÇÕES — LANÇAMENTOS
 ****************************************************/

function LA_isValidTipoCategoria_(tipo, categoria) {
  const index = LA_criarIndiceConfigLancamentos_(LA_getConfigValues_());

  return LA_configPossuiTipoCategoria_(index, tipo, categoria);
}


function LA_isValidTipoCategoriaSubcategoria_(tipo, categoria, subcategoria) {
  const index = LA_criarIndiceConfigLancamentos_(LA_getConfigValues_());

  return LA_configPossuiTipoCategoriaSubcategoria_(index, tipo, categoria, subcategoria);
}


function LA_criarIndiceConfigLancamentos_(values) {
  const index = {
    tipoCategoria: Object.create(null),
    tipoCategoriaSubcategoria: Object.create(null)
  };

  values.forEach(function (row) {
    const tipo = LA_normalizarTexto_(row[0]);
    const categoria = LA_normalizarTexto_(row[1]);
    const subcategoria = LA_normalizarTexto_(row[2]);

    if (!tipo || !categoria) return;

    index.tipoCategoria[LA_chaveCompostaTexto_([tipo, categoria])] = true;
    index.tipoCategoriaSubcategoria[LA_chaveCompostaTexto_([tipo, categoria, subcategoria])] = true;
  });

  return index;
}


function LA_configPossuiTipoCategoria_(index, tipo, categoria) {
  return Boolean(index.tipoCategoria[LA_chaveCompostaTexto_([tipo, categoria])]);
}


function LA_configPossuiTipoCategoriaSubcategoria_(index, tipo, categoria, subcategoria) {
  return Boolean(
    index.tipoCategoriaSubcategoria[LA_chaveCompostaTexto_([tipo, categoria, subcategoria])]
  );
}


function LA_getConfigValues_(spreadsheet) {
  const ss = spreadsheet || SpreadsheetApp.getActive();

  const range =
    ss.getRangeByName('rngConfigListas') ||
    ss.getRangeByName('cfgTipoCategoriaSubcategoria');

  const values = range
    ? range.getValues()
    : LA_getConfigRowsFromSheet_(LA_obterAbaObrigatoria_(ss, LA_ABA_CONFIG));

  if (values.length > 0) {
    const primeira = LA_chaveTexto_(values[0][0]);
    // Pula linha de cabeçalho (qualquer variação de "Tipo" ou célula vazia na primeira col)
    if (primeira === 'tipo' || primeira === 'tipos' || primeira === '') {
      return values.slice(1);
    }
  }

  return values;
}


/****************************************************
 * HELPERS — LANÇAMENTOS
 ****************************************************/

function LA_executarComErroAmigavel_(contexto, callback) {
  try {
    callback();
  } catch (error) {
    const ui = SpreadsheetApp.getUi();

    console.error(JSON.stringify({
      module: 'LANÇAMENTOS',
      contexto: contexto,
      message: LA_getMensagemErro_(error),
      stack: error && error.stack ? error.stack : '',
      timestamp: new Date().toISOString()
    }));

    ui.alert(
      contexto,
      'Não foi possível concluir a operação.\n\n' + LA_getMensagemErro_(error),
      ui.ButtonSet.OK
    );
  }
}


function LA_executarComLockDocumento_(callback) {
  const lock = LockService.getDocumentLock();
  let lockObtido = false;

  try {
    try {
      lock.waitLock(10000);
    } catch (e) {
      throw new Error(
        'A planilha está sendo usada por outra pessoa ou aba.\n' +
        'Aguarde alguns segundos e tente novamente.'
      );
    }
    lockObtido = true;
    return callback();
  } finally {
    if (lockObtido) {
      lock.releaseLock();
    }
  }
}


function LA_obterAbaObrigatoria_(spreadsheet, nomeAba) {
  const sheet = spreadsheet.getSheetByName(nomeAba);

  if (!sheet) {
    throw new Error('Aba "' + nomeAba + '" não encontrada.');
  }

  return sheet;
}


function LA_getValue_(namedRange) {
  const range = SpreadsheetApp.getActive().getRangeByName(namedRange);
  return range ? range.getValue() : '';
}


function LA_lerValoresIntervalosNomeados_(spreadsheet, names, namedRangeMap) {
  const ranges = namedRangeMap || LA_mapearIntervalosNomeados_(spreadsheet, names);
  const lancamentos = spreadsheet.getSheetByName(LA_ABA_LANCAMENTOS);
  const values = {};

  names.forEach(function (name) {
    if (lancamentos && LA_FORMULARIO_A1[name]) {
      values[name] = lancamentos.getRange(LA_FORMULARIO_A1[name]).getValue();
      return;
    }

    values[name] = ranges[name] ? ranges[name].getRange().getValue() : '';
  });

  return values;
}


function LA_mapearIntervalosNomeados_(spreadsheet, names) {
  const wanted = names.reduce(function (acc, name) {
    acc[name] = true;
    return acc;
  }, {});
  const map = {};

  spreadsheet.getNamedRanges().forEach(function (namedRange) {
    const name = namedRange.getName();

    if (wanted[name]) {
      map[name] = namedRange;
    }
  });

  return map;
}


function LA_listarIntervalosNomeadosAusentes_(spreadsheet, names, namedRangeMap) {
  const ranges = namedRangeMap || LA_mapearIntervalosNomeados_(spreadsheet, names);
  const lancamentos = spreadsheet.getSheetByName(LA_ABA_LANCAMENTOS);

  return names.filter(function (name) {
    if (lancamentos && LA_FORMULARIO_A1[name]) return false;
    return !ranges[name];
  });
}


function LA_parseValor_(value) {
  if (typeof value === 'number') return value;
  if (value instanceof Date) return NaN;

  const raw = String(value || '').trim();

  if (!raw) return NaN;

  const isNegative = raw.includes('-') || /^\(.*\)$/.test(raw);

  let clean = raw
    .replace(/[^\d,.-]/g, '')
    .replace(/[()]/g, '')
    .replace(/-/g, '');

  if (!clean) return NaN;

  const lastComma = clean.lastIndexOf(',');
  const lastDot = clean.lastIndexOf('.');

  if (lastComma > -1 && lastDot > -1) {
    clean = lastComma > lastDot
      ? clean.replace(/\./g, '').replace(',', '.')
      : clean.replace(/,/g, '');
  } else if (lastComma > -1) {
    clean = LA_normalizarSeparadorUnico_(clean, ',');
  } else if (lastDot > -1) {
    clean = LA_normalizarSeparadorUnico_(clean, '.');
  }

  const number = Number(clean);

  if (!Number.isFinite(number)) return NaN;

  return isNegative ? -number : number;
}


function LA_normalizarSeparadorUnico_(value, separator) {
  const escapedSeparator = separator === '.' ? '\\.' : separator;
  const regex = new RegExp(escapedSeparator, 'g');
  const parts = value.split(separator);
  const decimalCandidate = parts[parts.length - 1];

  if (parts.length > 2) {
    const todosGruposDeMilhar = parts.slice(1).every(function (part) {
      return part.length === 3;
    });

    if (todosGruposDeMilhar) {
      return value.replace(regex, '');
    }

    return parts.slice(0, -1).join('') + '.' + decimalCandidate;
  }

  if (decimalCandidate.length === 3 && parts[0].length <= 3) {
    return value.replace(regex, '');
  }

  return separator === ','
    ? value.replace(/\./g, '').replace(',', '.')
    : value;
}


function LA_parseData_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return LA_dateOnly_(value);
  }

  const raw = LA_normalizarTexto_(value);

  if (!raw) return null;

  const brDate = raw.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/);

  if (brDate) {
    return LA_criarDataValidada_(
      Number(brDate[3]),
      Number(brDate[2]),
      Number(brDate[1])
    );
  }

  const isoDate = raw.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);

  if (isoDate) {
    return LA_criarDataValidada_(
      Number(isoDate[1]),
      Number(isoDate[2]),
      Number(isoDate[3])
    );
  }

  const parsed = new Date(raw);

  if (isNaN(parsed.getTime())) return null;

  return LA_dateOnly_(parsed);
}


function LA_criarDataValidada_(year, month, day) {
  const fullYear = year < 100 ? 2000 + year : year;
  const date = new Date(fullYear, month - 1, day);

  if (
    date.getFullYear() !== fullYear ||
    date.getMonth() !== month - 1 ||
    date.getDate() !== day
  ) {
    return null;
  }

  return LA_dateOnly_(date);
}


function LA_dateOnly_(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}


function LA_mesPorData_(dateValue) {
  const data = LA_parseData_(dateValue);

  if (!data) return '';

  const meses = [
    'JANEIRO',
    'FEVEREIRO',
    'MARÇO',
    'ABRIL',
    'MAIO',
    'JUNHO',
    'JULHO',
    'AGOSTO',
    'SETEMBRO',
    'OUTUBRO',
    'NOVEMBRO',
    'DEZEMBRO'
  ];

  return meses[data.getMonth()];
}


function LA_proximaLinha_(sheet) {
  const lastRow = sheet.getLastRow();

  if (lastRow < LA_PRIMEIRA_LINHA_DADOS) return LA_PRIMEIRA_LINHA_DADOS;

  // Lê apenas até a última linha usada + 1 (não lê 4980 linhas desnecessariamente)
  const numRows = Math.min(lastRow - LA_PRIMEIRA_LINHA_DADOS + 2, LA_TOTAL_LINHAS_DADOS);

  const values = sheet
    .getRange(LA_PRIMEIRA_LINHA_DADOS, LA_COLUNAS_LANCAMENTO.DATA, numRows, 1)
    .getValues();

  const idx = values.findIndex(function (row) { return row[0] === ''; });

  if (idx !== -1) return LA_PRIMEIRA_LINHA_DADOS + idx;

  return lastRow < LA_ULTIMA_LINHA_DADOS ? lastRow + 1 : null;
}


function LA_existeChaveTecnicaLancamento_(sheet, duplicateKey) {
  if (!duplicateKey) return false;

  const lastRow = sheet.getLastRow();
  if (lastRow < LA_PRIMEIRA_LINHA_DADOS) return false;

  const numRows = Math.min(lastRow, LA_ULTIMA_LINHA_DADOS) - LA_PRIMEIRA_LINHA_DADOS + 1;

  const values = sheet
    .getRange(LA_PRIMEIRA_LINHA_DADOS, LA_COLUNAS_LANCAMENTO.CHAVE_TECNICA, numRows, 1)
    .getValues();

  return values.some(function (row) { return row[0] === duplicateKey; });
}


function LA_makeKey_(data, unidade, tipo, categoria, subcategoria, valor) {
  const tz = Session.getScriptTimeZone();
  // Evita reconstruir Date a partir de Date (risco de drift de timezone)
  const dateObj = data instanceof Date ? data : new Date(data);
  const dataKey = Utilities.formatDate(dateObj, tz, 'yyyyMMdd');

  return [
    dataKey,
    String(unidade || '').trim(),
    String(tipo || '').trim(),
    String(categoria || '').trim(),
    String(subcategoria || '').trim(),
    Number(valor).toFixed(2)
  ].join('|');
}


function LA_formatDateKey_(data) {
  const parsed = LA_parseData_(data);
  if (!parsed) return '';
  const tz = Session.getScriptTimeZone();
  return Utilities.formatDate(parsed, tz, 'yyyyMMdd');
}


function LA_promptTextComExemplos_(title, message, examples) {
  const ui = SpreadsheetApp.getUi();

  const textoExemplos = Array.isArray(examples) && examples.length
    ? '\n\n' + examples.join('\n')
    : '';

  const response = ui.prompt(
    title,
    message + textoExemplos,
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return '';

  return LA_normalizarTexto_(response.getResponseText());
}


function LA_normalizarTexto_(value) {
  return String(value || '').trim();
}


function LA_chaveTexto_(value) {
  return LA_normalizarTexto_(value)
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}


function LA_chaveCompostaTexto_(values) {
  return values.map(LA_chaveTexto_).join('|');
}


function LA_equalsTexto_(a, b) {
  return LA_chaveTexto_(a) === LA_chaveTexto_(b);
}


function LA_matchExistingText_(input, existingList) {
  const texto = LA_normalizarTexto_(input);
  const chave = LA_chaveTexto_(texto);

  const found = existingList.find(function (item) {
    return LA_chaveTexto_(item) === chave;
  });

  return found ? LA_normalizarTexto_(found) : texto;
}


function LA_matchClasseRelatorio_(value) {
  const texto = LA_normalizarTexto_(value);

  return LA_CLASSES_RELATORIO.find(function (classe) {
    return LA_equalsTexto_(classe, texto);
  }) || '';
}


function LA_uniquePreserveOrder_(arr) {
  const seen = {};

  return arr
    .map(function (value) { return LA_normalizarTexto_(value); })
    .filter(function (value) {
      const key = LA_chaveTexto_(value);

      if (!value || seen[key]) return false;

      seen[key] = true;
      return true;
    });
}


function LA_firstBlankRow_(sheet, col, startRow) {
  const values = sheet
    .getRange(startRow, col, sheet.getMaxRows() - startRow + 1, 1)
    .getValues();

  const idx = values.findIndex(function (row) {
    return row[0] === '';
  });

  return idx === -1 ? sheet.getLastRow() + 1 : startRow + idx;
}


function LA_getConfigRowsFromSheet_(sheet) {
  const lastRow = Math.max(sheet.getLastRow(), 2);
  return sheet.getRange(2, 1, lastRow - 1, 6).getValues();
}


function LA_limparListaVertical_(sheet, column) {
  sheet.getRange(2, column, LA_TOTAL_LINHAS_LISTAS, 1).clearContent();
}


function LA_escreverListaVertical_(sheet, column, values) {
  if (!values.length) return;

  if (values.length > LA_TOTAL_LINHAS_LISTAS) {
    throw new Error(
      'A lista da coluna ' + column + ' possui mais itens do que o intervalo disponível.'
    );
  }

  sheet
    .getRange(2, column, values.length, 1)
    .setValues(values.map(function (value) { return [value]; }));
}


function LA_agruparLinhasConsecutivas_(rows) {
  if (!rows.length) return [];

  const grupos = [];
  let inicio = rows[0];
  let anterior = rows[0];

  rows.slice(1).forEach(function (row) {
    if (row === anterior + 1) {
      anterior = row;
      return;
    }

    grupos.push({
      inicio: inicio,
      quantidade: anterior - inicio + 1
    });

    inicio = row;
    anterior = row;
  });

  grupos.push({
    inicio: inicio,
    quantidade: anterior - inicio + 1
  });

  return grupos;
}


function LA_campoPreenchido_(value) {
  if (value instanceof Date) return true;
  return LA_normalizarTexto_(value) !== '';
}


function LA_setNamedRangeSafe_(name, range) {
  LA_setNamedRangesSafe_([{ name: name, range: range }]);
}


function LA_setNamedRangesSafe_(items) {
  const ss = SpreadsheetApp.getActive();
  const names = items.reduce(function (acc, item) {
    acc[item.name] = true;
    return acc;
  }, {});

  ss.getNamedRanges().forEach(function (namedRange) {
    if (names[namedRange.getName()]) {
      namedRange.remove();
    }
  });

  items.forEach(function (item) {
    ss.setNamedRange(item.name, item.range);
  });
}


function LA_getMensagemErro_(error) {
  if (error && error.message) return error.message;
  return String(error || 'Erro desconhecido.');
}


/****************************************************
 * DESFAZER ÚLTIMA EXCLUSÃO — LANÇAMENTOS
 ****************************************************/

function LA_salvarBackupExclusao_(sheet, rows) {
  try {
    const numCols = LA_COLUNAS_LANCAMENTO.CHAVE_TECNICA;
    const dados = rows.map(function (row) {
      return {
        row: row,
        values: sheet.getRange(row, 1, 1, numCols).getValues()[0]
      };
    });

    const payload = JSON.stringify({
      timestamp: new Date().toISOString(),
      dados: dados
    });

    if (payload.length > 8000) return; // Não armazena se muito grande

    PropertiesService.getDocumentProperties()
      .setProperty('LA_ULTIMO_BACKUP_EXCLUSAO', payload);
  } catch (e) {
    // Falha silenciosa — não impede a exclusão
  }
}


function desfazerUltimaExclusao() {
  LA_executarComErroAmigavel_('Desfazer última exclusão', function () {
    const ss = SpreadsheetApp.getActive();
    const ui = SpreadsheetApp.getUi();
    const sheet = LA_obterAbaObrigatoria_(ss, LA_ABA_LANCAMENTOS);

    const raw = PropertiesService.getDocumentProperties()
      .getProperty('LA_ULTIMO_BACKUP_EXCLUSAO');

    if (!raw) {
      ui.alert('Nenhuma exclusão recente encontrada para desfazer.');
      return;
    }

    let backup;
    try {
      backup = JSON.parse(raw);
    } catch (e) {
      ui.alert('O backup da última exclusão está corrompido e não pode ser restaurado.');
      return;
    }

    const quando = new Date(backup.timestamp);
    const formatado = Utilities.formatDate(quando, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');

    const response = ui.alert(
      'Desfazer exclusão',
      'Restaurar ' + backup.dados.length + ' lançamento(s) excluído(s) em ' + formatado + '?\n\n' +
      'Os dados serão gravados de volta nas linhas originais.',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) return;

    LA_executarComLockDocumento_(function () {
      backup.dados.forEach(function (item) {
        sheet.getRange(item.row, 1, 1, item.values.length).setValues([item.values]);
      });

      SpreadsheetApp.flush();
    });

    PropertiesService.getDocumentProperties()
      .deleteProperty('LA_ULTIMO_BACKUP_EXCLUSAO');

    ss.toast(
      backup.dados.length + ' lançamento(s) restaurado(s) com sucesso.',
      'LANÇAMENTOS',
      5
    );
  });
}


/****************************************************
 * VALIDAR ESTRUTURA — LANÇAMENTOS
 ****************************************************/

function validarEstruturaLancamentos() {
  LA_executarComErroAmigavel_('Validar estrutura da planilha', function () {
    const ss = SpreadsheetApp.getActive();
    const ui = SpreadsheetApp.getUi();

    const abaLanc = ss.getSheetByName(LA_ABA_LANCAMENTOS);
    const abaCfg = ss.getSheetByName(LA_ABA_CONFIG);

    const intervalosObrigatorios = LA_INTERVALOS_OBRIGATORIOS_FORMULARIO;
    const intervalosAusentes = LA_listarIntervalosNomeadosAusentes_(ss, intervalosObrigatorios);

    const checks = [
      ['Aba LANÇAMENTOS existe', Boolean(abaLanc)],
      ['Aba CONFIG existe', Boolean(abaCfg)],
      ['Cabeçalho LANÇAMENTOS está na linha 19', abaLanc ? abaLanc.getRange(LA_LINHA_CABECALHO, 1, 1, LA_COLUNAS_LANCAMENTO.EXCLUIR).getNumColumns() === LA_COLUNAS_LANCAMENTO.EXCLUIR : false],
      ['Aba LANÇAMENTOS alcança linha 5000', abaLanc ? abaLanc.getMaxRows() >= LA_ULTIMA_LINHA_DADOS : false],
      ['CONFIG tem coluna F (ClasseRelatorio)', abaCfg ? abaCfg.getMaxColumns() >= 6 : false],
      ['Intervalos nomeados obrigatórios existem', intervalosAusentes.length === 0],
      ['Formulário atual existe em LANÇAMENTOS!' + LA_AREA_FORMULARIO_A1, abaLanc ? abaLanc.getRange(LA_AREA_FORMULARIO_A1).getNumRows() === 11 : false],
      ['CONFIG tem dados cadastrados', abaCfg ? abaCfg.getLastRow() > 1 : false],
      ['Backup de exclusão disponível', Boolean(
        PropertiesService.getDocumentProperties().getProperty('LA_ULTIMO_BACKUP_EXCLUSAO')
      )]
    ];

    const problemas = checks
      .filter(function (c) { return !c[1]; })
      .map(function (c) { return '✗ ' + c[0]; });

    const ok = checks
      .filter(function (c) { return c[1]; })
      .map(function (c) { return '✓ ' + c[0]; });

    if (problemas.length === 0) {
      ui.alert(
        'Estrutura OK',
        'Todos os ' + ok.length + ' itens verificados estão corretos.\n\n' + ok.join('\n'),
        ui.ButtonSet.OK
      );
      return;
    }

    const detalheAusentes = intervalosAusentes.length
      ? '\n\nIntervalos nomeados ausentes:\n' + intervalosAusentes.join('\n')
      : '';

    ui.alert(
      'Problemas encontrados',
      'Itens com problema:\n' + problemas.join('\n') +
      '\n\nItens OK:\n' + ok.join('\n') +
      detalheAusentes,
      ui.ButtonSet.OK
    );
  });
}


/****************************************************
 * BACKUP PARA DRIVE — COMPARTILHADO
 ****************************************************/

function fazerBackupDrive() {
  LA_executarComErroAmigavel_('Backup agora', function () {
    const ss = SpreadsheetApp.getActive();
    const ui = SpreadsheetApp.getUi();
    const tz = Session.getScriptTimeZone();

    const response = ui.alert(
      'Backup agora',
      'Isso criará uma cópia desta planilha no Google Drive com a data e hora atuais.\n\nDeseja continuar?',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) return;

    const timestamp = Utilities.formatDate(new Date(), tz, 'dd-MM-yyyy HH:mm');
    const nomeBackup = 'BACKUP ' + timestamp + ' — ' + ss.getName();

    const file = DriveApp.getFileById(ss.getId());
    const backup = file.makeCopy(nomeBackup);

    ss.toast(
      'Backup criado: "' + nomeBackup + '"',
      'Backup',
      8
    );

    console.info(JSON.stringify({
      module: 'BACKUP',
      arquivo: nomeBackup,
      fileId: backup.getId(),
      timestamp: new Date().toISOString()
    }));
  });
}


/****************************************************
 * ==================================================
 * MÓDULO 2 — BOLETOS PAGOS
 * ==================================================
 ****************************************************/


/****************************************************
 * CONFIGURAÇÕES — BOLETOS
 ****************************************************/

const BP_CONFIG = Object.freeze({
  sheets: Object.freeze({
    boletos: 'BOLETOS PAGOS',
    confi: 'CONFI',
    relatorio: 'RELATÓRIO DE PAGAMENTOS',
  }),

  rows: Object.freeze({
    header: 14,
    firstData: 15,
    lastData: 5000,
    confiFirstData: 2,
  }),

  form: Object.freeze({
    range: 'B4:B12',
    dataVencimentoIndex: 0,
    lojaIndex: 1,
    empresaIndex: 2,
    descricaoIndex: 3,
    valorPrevistoIndex: 4,
    valorPagoIndex: 5,
    formaPagamentoIndex: 6,
    dataPagamentoIndex: 7,
    observacoesIndex: 8,
  }),

  confiColumns: Object.freeze({
    loja: 1,
    empresa: 2,
    categoria: 3,
    formaPagamento: 4,
    status: 5,
  }),

  dataColumns: Object.freeze({
    dataVencimento: 1,
    diaSemana: 2,
    semana: 3,
    loja: 4,
    empresa: 5,
    descricao: 6,
    formaPagamento: 7,
    status: 8,
    valorPrevisto: 9,
    valorPago: 10,
    diferenca: 11,
    dataPagamento: 12,
    observacoes: 13,
    mes: 14,
    numeroMes: 15,
    ano: 16,
    semanaRelatorio: 17,
    statusAutomatico: 18,
    diferencaAutomatica: 19,
    valorEmAberto: 20,
    valorEmAtraso: 21,
    quantidade: 22,
  }),

  cache: Object.freeze({
    confiKey: 'BP_CONFI_LISTS_V1',
    ttlSeconds: 300,
  }),

  menuName: 'Boletos',
});


/****************************************************
 * SETUP — BOLETOS
 ****************************************************/

function setupBoletos() {
  const startedAt = Date.now();

  try {
    const ctx = BP_getContext_({
      requireRelatorio: true,
    });

    BP_garantirLinhasBase_(ctx.boletos);
    BP_garantirCabecalhosBoletos_(ctx.boletos);
    configurarValidacoesBoletos(false);
    BP_aplicarFormulasAutomaticas_(ctx.boletos);
    BP_formatarEstruturaBoletos_(ctx.boletos);

    SpreadsheetApp.flush();

    ctx.ss.toast(
      'Boletos configurado: validações, fórmulas e layout revisados.',
      BP_CONFIG.menuName,
      6
    );

    BP_logInfo_('setupBoletos concluído', {
      durationMs: Date.now() - startedAt,
    });
  } catch (error) {
    BP_showError_('Erro ao configurar boletos', error);
  }
}


/****************************************************
 * REGISTRAR BOLETO
 ****************************************************/

function registrarBoleto() {
  const startedAt = Date.now();
  const operationId = BP_criarOperationId_();
  const lock = LockService.getDocumentLock();

  try {
    lock.waitLock(10000);

    const ctx = BP_getContext_();
    const form = BP_lerFormularioBoletos_(ctx.boletos);
    const confiIndex = BP_getConfiIndex_(ctx.confi);

    const validation = BP_validarFormularioBoleto_(form, confiIndex);

    if (!validation.ok) {
      ctx.ui.alert(validation.message);
      return;
    }

    BP_garantirLinhasBase_(ctx.boletos);

    const baseInfo = BP_analisarBaseBoletos_(ctx.boletos, form);
    const nextRow = baseInfo.nextRow;

    if (!nextRow) {
      ctx.ui.alert(
        'A tabela de boletos está cheia.\n\n' +
        'Expanda a base operacional para além da linha ' + BP_CONFIG.rows.lastData + '.'
      );
      return;
    }

    if (baseInfo.duplicado) {
      const resposta = ctx.ui.alert(
        'Possível duplicidade',
        'Já existe um boleto com a mesma Data de vencimento, Loja, Empresa, Descrição e Valor previsto.\n\nDeseja registrar mesmo assim?',
        ctx.ui.ButtonSet.YES_NO
      );

      if (resposta !== ctx.ui.Button.YES) return;
    }

    BP_escreverBoleto_(ctx.boletos, nextRow, form);
    BP_formatarLinhaBoleto_(ctx.boletos, nextRow);

    SpreadsheetApp.flush();

    limparFormularioBoletos(false);

    ctx.ss.toast(
      'Boleto registrado com sucesso na linha ' + nextRow + '.',
      BP_CONFIG.sheets.boletos,
      5
    );

    BP_logInfo_('registrarBoleto concluído', {
      operationId,
      row: nextRow,
      loja: form.loja,
      empresa: form.empresa,
      valorPrevisto: form.valorPrevisto,
      durationMs: Date.now() - startedAt,
    });
  } catch (error) {
    BP_showError_('Erro ao registrar boleto', error);
  } finally {
    try {
      lock.releaseLock();
    } catch (error) {
      BP_logWarn_('Falha ao liberar lock', {
        operationId,
        error: BP_getErrorMessage_(error),
      });
    }
  }
}


/****************************************************
 * FORMULÁRIO — BOLETOS
 ****************************************************/

function BP_lerFormularioBoletos_(sheet) {
  const values = sheet
    .getRange(BP_CONFIG.form.range)
    .getValues()
    .flat();

  return {
    dataVencimento: BP_parseData_(values[BP_CONFIG.form.dataVencimentoIndex]),
    loja: BP_texto_(values[BP_CONFIG.form.lojaIndex]),
    empresa: BP_texto_(values[BP_CONFIG.form.empresaIndex]),
    descricao: BP_texto_(values[BP_CONFIG.form.descricaoIndex]),
    valorPrevisto: BP_parseValor_(values[BP_CONFIG.form.valorPrevistoIndex]),
    valorPago: BP_parseValorOpcional_(values[BP_CONFIG.form.valorPagoIndex]),
    formaPagamento: BP_texto_(values[BP_CONFIG.form.formaPagamentoIndex]),
    dataPagamento: BP_parseDataOpcional_(values[BP_CONFIG.form.dataPagamentoIndex]),
    observacoes: BP_texto_(values[BP_CONFIG.form.observacoesIndex]),
  };
}


function limparFormularioBoletos(showToast) {
  try {
    const ss = SpreadsheetApp.getActive();
    const boletos = BP_getRequiredSheet_(ss, BP_CONFIG.sheets.boletos);

    boletos.getRange(BP_CONFIG.form.range).clearContent();

    if (showToast !== false) {
      ss.toast('Formulário de boletos limpo.', BP_CONFIG.menuName, 4);
    }
  } catch (error) {
    BP_showError_('Erro ao limpar formulário', error);
  }
}


/****************************************************
 * VALIDAÇÃO — BOLETOS
 ****************************************************/

function BP_validarFormularioBoleto_(form, confiIndex) {
  if (!BP_ehDataValida_(form.dataVencimento)) {
    return BP_fail_('Informe uma Data de vencimento válida.');
  }

  if (!form.loja) {
    return BP_fail_('Informe a Loja.');
  }

  if (!form.empresa) {
    return BP_fail_('Informe a Empresa / Fornecedor.');
  }

  if (!Number.isFinite(form.valorPrevisto) || form.valorPrevisto <= 0) {
    return BP_fail_('Informe um Valor a pagar válido e maior que zero.');
  }

  if (form.valorPago !== '' && (!Number.isFinite(form.valorPago) || form.valorPago < 0)) {
    return BP_fail_('Informe um Valor pago válido ou deixe o campo em branco.');
  }

  if (form.dataPagamento !== '' && !BP_ehDataValida_(form.dataPagamento)) {
    return BP_fail_('Informe uma Data de pagamento válida ou deixe o campo em branco.');
  }

  if (form.valorPago !== '' && form.valorPago > 0 && !BP_ehDataValida_(form.dataPagamento)) {
    return BP_fail_('Se houver Valor pago, informe também a Data de pagamento.');
  }

  if (BP_ehDataValida_(form.dataPagamento) && (form.valorPago === '' || form.valorPago <= 0)) {
    return BP_fail_('Se houver Data de pagamento, informe também o Valor pago.');
  }

  if (!confiIndex.lojas.has(BP_normalizarComparacao_(form.loja))) {
    return BP_fail_(
      'A loja "' + form.loja + '" não existe na aba CONFI.\n\n' +
      'Adicione a loja antes de registrar.'
    );
  }

  if (!confiIndex.empresas.has(BP_normalizarComparacao_(form.empresa))) {
    return BP_fail_(
      'A empresa/fornecedor "' + form.empresa + '" não existe na aba CONFI.\n\n' +
      'Adicione antes de registrar.'
    );
  }

  if (
    form.formaPagamento &&
    !confiIndex.formasPagamento.has(BP_normalizarComparacao_(form.formaPagamento))
  ) {
    return BP_fail_(
      'A forma de pagamento "' + form.formaPagamento + '" não existe na aba CONFI.\n\n' +
      'Adicione antes de registrar.'
    );
  }

  return {
    ok: true,
    message: '',
  };
}


function BP_fail_(message) {
  return {
    ok: false,
    message,
  };
}


/****************************************************
 * ESCRITA — BOLETOS
 ****************************************************/

function BP_escreverBoleto_(sheet, row, form) {
  const col = BP_CONFIG.dataColumns;

  sheet
    .getRange(row, col.dataVencimento, 1, 1)
    .setValue(form.dataVencimento);

  sheet
    .getRange(row, col.loja, 1, 4)
    .setValues([[
      form.loja,
      form.empresa,
      form.descricao,
      form.formaPagamento,
    ]]);

  sheet
    .getRange(row, col.valorPrevisto, 1, 2)
    .setValues([[
      form.valorPrevisto,
      form.valorPago === '' ? '' : form.valorPago,
    ]]);

  sheet
    .getRange(row, col.dataPagamento, 1, 2)
    .setValues([[
      BP_ehDataValida_(form.dataPagamento) ? form.dataPagamento : '',
      form.observacoes,
    ]]);
}


/****************************************************
 * DUPLICIDADE — BOLETOS
 ****************************************************/

function BP_analisarBaseBoletos_(sheet, form) {
  const first = BP_CONFIG.rows.firstData;
  const totalRows = BP_CONFIG.rows.lastData - first + 1;
  const col = BP_CONFIG.dataColumns;
  const targetKey = BP_chaveBoleto_(
    form.dataVencimento,
    form.loja,
    form.empresa,
    form.descricao,
    form.valorPrevisto
  );

  const values = sheet
    .getRange(first, 1, totalRows, col.valorPrevisto)
    .getValues();

  let nextRow = null;
  let duplicated = false;
  let lastManualRow = first - 1;

  values.forEach(function (row, index) {
    const dataVencimento = row[col.dataVencimento - 1];

    if (dataVencimento === '' && nextRow === null) {
      nextRow = first + index;
      return;
    }

    if (dataVencimento === '') return;

    lastManualRow = first + index;

    if (duplicated) return;

    const loja = row[col.loja - 1];
    const empresa = row[col.empresa - 1];
    const descricao = row[col.descricao - 1];
    const valorPrevisto = row[col.valorPrevisto - 1];

    if (!loja || !empresa || !valorPrevisto) return;

    duplicated = BP_chaveBoleto_(
      dataVencimento,
      loja,
      empresa,
      descricao,
      valorPrevisto
    ) === targetKey;
  });

  return {
    nextRow,
    duplicado: duplicated,
    ultimaLinhaManual: lastManualRow,
  };
}


function BP_boletoDuplicado_(sheet, form) {
  return BP_analisarBaseBoletos_(sheet, form).duplicado;
}


/****************************************************
 * VALIDAÇÕES — BOLETOS
 ****************************************************/

function configurarValidacoesBoletos(showToast) {
  try {
    const ctx = BP_getContext_();

    const maxConfiRows = Math.max(ctx.confi.getMaxRows() - 1, 1);
    const col = BP_CONFIG.confiColumns;

    const regraLojas = BP_criarRegraLista_(
      ctx.confi.getRange(2, col.loja, maxConfiRows, 1)
    );

    const regraEmpresas = BP_criarRegraLista_(
      ctx.confi.getRange(2, col.empresa, maxConfiRows, 1)
    );

    const regraFormasPagamento = BP_criarRegraLista_(
      ctx.confi.getRange(2, col.formaPagamento, maxConfiRows, 1)
    );

    const regraData = SpreadsheetApp.newDataValidation()
      .requireDate()
      .setAllowInvalid(false)
      .build();

    const regraValorPositivo = SpreadsheetApp.newDataValidation()
      .requireNumberGreaterThan(0)
      .setAllowInvalid(false)
      .build();

    const regraValorNaoNegativo = SpreadsheetApp.newDataValidation()
      .requireNumberGreaterThanOrEqualTo(0)
      .setAllowInvalid(false)
      .build();

    BP_aplicarValidacoesFormulario_(
      ctx.boletos,
      regraData,
      regraLojas,
      regraEmpresas,
      regraFormasPagamento,
      regraValorPositivo,
      regraValorNaoNegativo
    );

    BP_aplicarValidacoesTabela_(
      ctx.boletos,
      regraData,
      regraLojas,
      regraEmpresas,
      regraFormasPagamento,
      regraValorPositivo,
      regraValorNaoNegativo
    );

    BP_limparValidacoesColunasAutomaticas_(ctx.boletos);

    SpreadsheetApp.flush();

    if (showToast !== false) {
      ctx.ss.toast(
        'Validações configuradas com base na aba CONFI.',
        BP_CONFIG.menuName,
        5
      );
    }
  } catch (error) {
    BP_showError_('Erro ao configurar validações', error);
  }
}


function BP_aplicarValidacoesFormulario_(
  sheet,
  regraData,
  regraLojas,
  regraEmpresas,
  regraFormasPagamento,
  regraValorPositivo,
  regraValorNaoNegativo
) {
  APP_setDataValidationForA1Ranges_(sheet, ['B4', 'B11'], regraData);
  sheet.getRange('B5').setDataValidation(regraLojas);
  sheet.getRange('B6').setDataValidation(regraEmpresas);
  sheet.getRange('B8').setDataValidation(regraValorPositivo);
  sheet.getRange('B9').setDataValidation(regraValorNaoNegativo);
  sheet.getRange('B10').setDataValidation(regraFormasPagamento);
}


function BP_aplicarValidacoesTabela_(
  sheet,
  regraData,
  regraLojas,
  regraEmpresas,
  regraFormasPagamento,
  regraValorPositivo,
  regraValorNaoNegativo
) {
  const first = BP_CONFIG.rows.firstData;
  const last = BP_CONFIG.rows.lastData;
  const totalRows = last - first + 1;
  const col = BP_CONFIG.dataColumns;

  APP_setDataValidationForA1Ranges_(
    sheet,
    [
      APP_a1Range_(first, col.dataVencimento, totalRows, 1),
      APP_a1Range_(first, col.dataPagamento, totalRows, 1)
    ],
    regraData
  );

  sheet.getRange(first, col.loja, totalRows, 1).setDataValidation(regraLojas);
  sheet.getRange(first, col.empresa, totalRows, 1).setDataValidation(regraEmpresas);
  sheet.getRange(first, col.formaPagamento, totalRows, 1).setDataValidation(regraFormasPagamento);
  sheet.getRange(first, col.valorPrevisto, totalRows, 1).setDataValidation(regraValorPositivo);
  sheet.getRange(first, col.valorPago, totalRows, 1).setDataValidation(regraValorNaoNegativo);
}


function BP_criarRegraLista_(range) {
  return SpreadsheetApp.newDataValidation()
    .requireValueInRange(range, true)
    .setAllowInvalid(false)
    .build();
}


/****************************************************
 * FÓRMULAS AUTOMÁTICAS — BOLETOS
 *
 * Padrão confirmado no seu Google Sheets:
 * - setFormula()
 * - Funções em inglês
 * - Separador ;
 ****************************************************/

function reaplicarFormulasBoletos() {
  try {
    const ctx = BP_getContext_();

    const resposta = ctx.ui.alert(
      'Reaplicar fórmulas automáticas',
      'Isso vai limpar e reaplicar as fórmulas em B:C, H, K e N:V a partir da linha 15.\n\nAs colunas manuais não serão apagadas.\n\nDeseja continuar?',
      ctx.ui.ButtonSet.YES_NO
    );

    if (resposta !== ctx.ui.Button.YES) return;

    BP_aplicarFormulasAutomaticas_(ctx.boletos);

    SpreadsheetApp.flush();

    ctx.ss.toast(
      'Fórmulas automáticas reaplicadas com sucesso.',
      BP_CONFIG.menuName,
      5
    );
  } catch (error) {
    BP_showError_('Erro ao reaplicar fórmulas', error);
  }
}


function BP_aplicarFormulasAutomaticas_(sheet) {
  BP_limparConteudoColunasAutomaticas_(sheet);
  BP_limparValidacoesColunasAutomaticas_(sheet);

  const first = BP_CONFIG.rows.firstData;
  const dataRange = 'A' + first + ':A';
  const valorPrevistoRange = 'I' + first + ':I';
  const valorPagoRange = 'J' + first + ':J';
  const statusAutomaticoRange = 'R' + first + ':R';

  const formulaSemana =
    '=ARRAYFORMULA(IF(' + dataRange + '="";"";"Semana "&IF(INT((DAY(' +
    dataRange + ')-1)/7)+1>5;5;INT((DAY(' + dataRange + ')-1)/7)+1)))';
  const formulaDiferenca =
    '=ARRAYFORMULA(IF(' + valorPrevistoRange + '="";"";N(' +
    valorPrevistoRange + ')-N(' + valorPagoRange + ')))';

  sheet
    .getRange('B' + first)
    .setFormula('=ARRAYFORMULA(IF(' + dataRange + '="";"";TEXT(' + dataRange + ';"dddd")))');

  sheet.getRange('C' + first).setFormula(formulaSemana);

  sheet
    .getRange('H' + first)
    .setFormula('=ARRAYFORMULA(IF(' + statusAutomaticoRange + '="";"";' + statusAutomaticoRange + '))');

  sheet.getRange('K' + first).setFormula(formulaDiferenca);

  sheet
    .getRange('N' + first)
    .setFormula('=ARRAYFORMULA(IF(' + dataRange + '="";"";UPPER(TEXT(' + dataRange + ';"mmmm"))))');

  sheet
    .getRange('O' + first)
    .setFormula('=ARRAYFORMULA(IF(' + dataRange + '="";"";MONTH(' + dataRange + ')))');

  sheet
    .getRange('P' + first)
    .setFormula('=ARRAYFORMULA(IF(' + dataRange + '="";"";YEAR(' + dataRange + ')))');

  sheet.getRange('Q' + first).setFormula(formulaSemana);

  sheet
    .getRange('R' + first)
    .setFormula(
      '=ARRAYFORMULA(IF(' + dataRange + '="";"";IF(N(' + valorPagoRange + ')>=N(' +
        valorPrevistoRange + ');"Pago";IF(N(' + valorPagoRange + ')>0;"Pago parcial";IF(' +
        dataRange + '<TODAY();"Em atraso";"Em aberto")))))'
    );

  sheet.getRange('S' + first).setFormula(formulaDiferenca);

  sheet
    .getRange('T' + first)
    .setFormula(
      '=ARRAYFORMULA(IF(' + dataRange + '="";"";IF((N(' + valorPagoRange + ')<N(' +
        valorPrevistoRange + '))*(' + dataRange + '>=TODAY());N(' + valorPrevistoRange +
        ')-N(' + valorPagoRange + ');0)))'
    );

  sheet
    .getRange('U' + first)
    .setFormula(
      '=ARRAYFORMULA(IF(' + dataRange + '="";"";IF((N(' + valorPagoRange + ')<N(' +
        valorPrevistoRange + '))*(' + dataRange + '<TODAY());N(' + valorPrevistoRange +
        ')-N(' + valorPagoRange + ');0)))'
    );

  sheet.getRange('V' + first).setFormula('=ARRAYFORMULA(IF(' + dataRange + '="";"";1))');
}


function BP_limparConteudoColunasAutomaticas_(sheet) {
  const first = BP_CONFIG.rows.firstData;
  const last = BP_CONFIG.rows.lastData;

  sheet
    .getRangeList([
      'B' + first + ':C' + last,
      'H' + first + ':H' + last,
      'K' + first + ':K' + last,
      'N' + first + ':V' + last,
    ])
    .clearContent();
}


function BP_limparValidacoesColunasAutomaticas_(sheet) {
  const first = BP_CONFIG.rows.firstData;
  const last = BP_CONFIG.rows.lastData;

  sheet
    .getRangeList([
      'B' + first + ':C' + last,
      'H' + first + ':H' + last,
      'K' + first + ':K' + last,
      'N' + first + ':V' + last,
    ])
    .clearDataValidations();
}


/****************************************************
 * ADICIONAR ITEM ÀS LISTAS — BOLETOS
 ****************************************************/

function adicionarItemAsListasBoletos() {
  try {
    const ctx = BP_getContext_();

    const opcoesTexto = [
      'Loja',
      'Empresa',
      'Categoria',
      'Forma de Pagamento',
    ];

    const tipoResp = ctx.ui.prompt(
      'Adicionar item às listas',
      'Qual lista deseja atualizar?\n\nOpções:\n' + opcoesTexto.join('\n'),
      ctx.ui.ButtonSet.OK_CANCEL
    );

    if (tipoResp.getSelectedButton() !== ctx.ui.Button.OK) return;

    const tipoDigitado = BP_texto_(tipoResp.getResponseText());
    const tipoNormalizado = BP_normalizarComparacao_(tipoDigitado);
    const coluna = BP_colunaListaConfi_(tipoNormalizado);

    if (!coluna) {
      ctx.ui.alert(
        'Opção inválida.\n\nUse uma destas opções:\n' +
        opcoesTexto.join('\n')
      );
      return;
    }

    const itemResp = ctx.ui.prompt(
      'Adicionar item às listas',
      'Digite o novo item para a lista "' + tipoDigitado + '":',
      ctx.ui.ButtonSet.OK_CANCEL
    );

    if (itemResp.getSelectedButton() !== ctx.ui.Button.OK) return;

    const item = BP_texto_(itemResp.getResponseText());

    if (!item) {
      ctx.ui.alert('Nenhum item informado.');
      return;
    }

    if (BP_existeNaLista_(ctx.confi, coluna, item)) {
      ctx.ui.alert('Esse item já existe nessa lista.');
      return;
    }

    const nextRow = BP_primeiraLinhaVaziaNaColuna_(
      ctx.confi,
      coluna,
      BP_CONFIG.rows.confiFirstData
    );

    ctx.confi.getRange(nextRow, coluna).setValue(item);

    BP_invalidateConfiCache_();
    configurarValidacoesBoletos(false);

    ctx.ss.toast(
      'Item adicionado em CONFI: ' + item,
      BP_CONFIG.sheets.confi,
      5
    );
  } catch (error) {
    BP_showError_('Erro ao adicionar item', error);
  }
}


function BP_colunaListaConfi_(tipoNormalizado) {
  const mapa = {
    LOJA: BP_CONFIG.confiColumns.loja,
    LOJAS: BP_CONFIG.confiColumns.loja,
    EMPRESA: BP_CONFIG.confiColumns.empresa,
    EMPRESAS: BP_CONFIG.confiColumns.empresa,
    FORNECEDOR: BP_CONFIG.confiColumns.empresa,
    FORNECEDORES: BP_CONFIG.confiColumns.empresa,
    CATEGORIA: BP_CONFIG.confiColumns.categoria,
    CATEGORIAS: BP_CONFIG.confiColumns.categoria,
    FORMADEPAGAMENTO: BP_CONFIG.confiColumns.formaPagamento,
    FORMASDEPAGAMENTO: BP_CONFIG.confiColumns.formaPagamento,
    PAGAMENTO: BP_CONFIG.confiColumns.formaPagamento,
  };

  return mapa[tipoNormalizado];
}


/****************************************************
 * VERIFICAÇÃO — BOLETOS
 ****************************************************/

function verificarEstruturaBoletos() {
  try {
    const ctx = BP_getContext_({
      requireRelatorio: true,
    });
    const totaisConfi = BP_contarItensListas_(ctx.confi, [
      BP_CONFIG.confiColumns.loja,
      BP_CONFIG.confiColumns.empresa,
      BP_CONFIG.confiColumns.formaPagamento,
    ]);

    const checks = [
      ['Aba BOLETOS PAGOS existe', Boolean(ctx.boletos)],
      ['Aba CONFI existe', Boolean(ctx.confi)],
      ['Aba RELATÓRIO DE PAGAMENTOS existe', Boolean(ctx.relatorio)],
      ['Formulário B4:B12 existe', ctx.boletos.getRange(BP_CONFIG.form.range).getNumRows() === 9],
      ['Cabeçalho A14:V14 tem 22 colunas', ctx.boletos.getRange(BP_CONFIG.rows.header, 1, 1, 22).getNumColumns() === 22],
      ['Base alcança linha 5000', ctx.boletos.getMaxRows() >= BP_CONFIG.rows.lastData],
      ['Fórmula em B15 existe', ctx.boletos.getRange('B15').getFormula() !== ''],
      ['Fórmula em H15 existe', ctx.boletos.getRange('H15').getFormula() !== ''],
      ['Fórmula em R15 existe', ctx.boletos.getRange('R15').getFormula() !== ''],
      ['Lista de lojas em CONFI tem itens', totaisConfi[BP_CONFIG.confiColumns.loja] > 0],
      ['Lista de empresas em CONFI tem itens', totaisConfi[BP_CONFIG.confiColumns.empresa] > 0],
      ['Lista de formas de pagamento em CONFI tem itens', totaisConfi[BP_CONFIG.confiColumns.formaPagamento] > 0],
    ];

    const problemas = checks
      .filter(function (item) { return !item[1]; })
      .map(function (item) { return '- ' + item[0]; });

    if (problemas.length) {
      ctx.ui.alert('Foram encontrados problemas:\n\n' + problemas.join('\n'));
      return;
    }

    ctx.ui.alert(
      'Estrutura validada com sucesso.\n\n' +
      'BOLETOS PAGOS, CONFI e RELATÓRIO DE PAGAMENTOS estão compatíveis com o script.'
    );
  } catch (error) {
    BP_showError_('Erro na verificação', error);
  }
}


/****************************************************
 * CABEÇALHOS E FORMATAÇÃO — BOLETOS
 ****************************************************/

function BP_garantirCabecalhosBoletos_(sheet) {
  const headers = [[
    'Data de Vencimento',
    'Dia da Semana',
    'Semana',
    'Unidade/Grupo',
    'Empresa',
    'Descrição/Observação',
    'Forma de Pagamento',
    'Status',
    'Valor Previsto',
    'Valor Pago',
    'Diferença',
    'Data de Pagamento',
    'Observações',
    'Mês',
    'Nº Mês',
    'Ano',
    'Semana do Mês',
    'Status Automático',
    'Diferença Automática',
    'Valor em Aberto',
    'Valor em Atraso',
    'Qtd',
  ]];

  sheet
    .getRange(BP_CONFIG.rows.header, 1, 1, 22)
    .setValues(headers);
}


function BP_formatarEstruturaBoletos_(sheet) {
  const first = BP_CONFIG.rows.firstData;
  const last = BP_CONFIG.rows.lastData;
  const totalRows = last - first + 1;
  const col = BP_CONFIG.dataColumns;

  sheet.setFrozenRows(BP_CONFIG.rows.header);

  sheet
    .getRange(BP_CONFIG.rows.header, 1, 1, 22)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#1D2330')
    .setFontColor('#FFFFFF');

  sheet
    .getRangeList([
      APP_a1Range_(first, col.dataVencimento, totalRows, 1),
      APP_a1Range_(first, col.dataPagamento, totalRows, 1)
    ])
    .setNumberFormat(APP_FORMATOS.DATA);

  sheet
    .getRangeList([
      APP_a1Range_(first, col.valorPrevisto, totalRows, 3),
      APP_a1Range_(first, col.valorEmAberto, totalRows, 2)
    ])
    .setNumberFormat(APP_FORMATOS.MOEDA_BRL);

  sheet.getRange(first, col.quantidade, totalRows, 1).setNumberFormat(APP_FORMATOS.INTEIRO);
}


function BP_formatarLinhaBoleto_(sheet, row) {
  const col = BP_CONFIG.dataColumns;

  sheet
    .getRangeList([
      APP_a1Range_(row, col.dataVencimento),
      APP_a1Range_(row, col.dataPagamento)
    ])
    .setNumberFormat(APP_FORMATOS.DATA);

  sheet
    .getRangeList([
      APP_a1Range_(row, col.valorPrevisto, 1, 3),
      APP_a1Range_(row, col.valorEmAberto, 1, 2)
    ])
    .setNumberFormat(APP_FORMATOS.MOEDA_BRL);
}


/****************************************************
 * CONTEXTO / SHEETS — BOLETOS
 ****************************************************/

function BP_getContext_(options) {
  const opts = options || {};
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  const boletos = BP_getRequiredSheet_(ss, BP_CONFIG.sheets.boletos);
  const confi = BP_getRequiredSheet_(ss, BP_CONFIG.sheets.confi);
  const relatorio = opts.requireRelatorio
    ? BP_getRequiredSheet_(ss, BP_CONFIG.sheets.relatorio)
    : ss.getSheetByName(BP_CONFIG.sheets.relatorio);

  return {
    ss,
    ui,
    boletos,
    confi,
    relatorio,
  };
}


function BP_getRequiredSheet_(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error('Aba "' + sheetName + '" não encontrada.');
  }

  return sheet;
}


function BP_garantirLinhasBase_(sheet) {
  const required = BP_CONFIG.rows.lastData;
  const current = sheet.getMaxRows();

  if (current < required) {
    sheet.insertRowsAfter(current, required - current);
  }
}


function BP_proximaLinhaBoleto_(sheet) {
  const totalRows = BP_CONFIG.rows.lastData - BP_CONFIG.rows.firstData + 1;

  const values = sheet
    .getRange(
      BP_CONFIG.rows.firstData,
      BP_CONFIG.dataColumns.dataVencimento,
      totalRows,
      1
    )
    .getValues();

  const index = values.findIndex(function (row) {
    return row[0] === '';
  });

  return index === -1 ? null : BP_CONFIG.rows.firstData + index;
}


function BP_ultimaLinhaManual_(sheet) {
  const totalRows = BP_CONFIG.rows.lastData - BP_CONFIG.rows.firstData + 1;

  const values = sheet
    .getRange(
      BP_CONFIG.rows.firstData,
      BP_CONFIG.dataColumns.dataVencimento,
      totalRows,
      1
    )
    .getValues();

  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] !== '') {
      return BP_CONFIG.rows.firstData + i;
    }
  }

  return BP_CONFIG.rows.firstData - 1;
}


/****************************************************
 * CONFI / CACHE — BOLETOS
 ****************************************************/

function BP_getConfiIndex_(sheet) {
  const lists = BP_getConfiLists_(sheet);

  return {
    lojas: BP_toNormalizedSet_(lists.lojas),
    empresas: BP_toNormalizedSet_(lists.empresas),
    categorias: BP_toNormalizedSet_(lists.categorias),
    formasPagamento: BP_toNormalizedSet_(lists.formasPagamento),
    status: BP_toNormalizedSet_(lists.status),
  };
}


function BP_getConfiLists_(sheet) {
  const cache = CacheService.getDocumentCache();
  const cached = cache.get(BP_CONFIG.cache.confiKey);

  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (error) {
      cache.remove(BP_CONFIG.cache.confiKey);
      BP_logWarn_('Cache CONFI inválido descartado', {
        error: BP_getErrorMessage_(error),
      });
    }
  }

  const lastRow = Math.max(sheet.getLastRow(), BP_CONFIG.rows.confiFirstData);
  const numRows = Math.max(lastRow - BP_CONFIG.rows.confiFirstData + 1, 1);
  const maxCol = Math.max.apply(null, Object.values(BP_CONFIG.confiColumns));

  const values = sheet
    .getRange(BP_CONFIG.rows.confiFirstData, 1, numRows, maxCol)
    .getValues();

  const lists = {
    lojas: [],
    empresas: [],
    categorias: [],
    formasPagamento: [],
    status: [],
  };

  values.forEach(function (row) {
    BP_pushIfNotBlank_(lists.lojas, row[BP_CONFIG.confiColumns.loja - 1]);
    BP_pushIfNotBlank_(lists.empresas, row[BP_CONFIG.confiColumns.empresa - 1]);
    BP_pushIfNotBlank_(lists.categorias, row[BP_CONFIG.confiColumns.categoria - 1]);
    BP_pushIfNotBlank_(lists.formasPagamento, row[BP_CONFIG.confiColumns.formaPagamento - 1]);
    BP_pushIfNotBlank_(lists.status, row[BP_CONFIG.confiColumns.status - 1]);
  });

  const serialized = JSON.stringify(lists);

  // Cache tem limite de 100KB por valor; só armazena se couber
  if (serialized.length < 90000) {
    try {
      cache.put(BP_CONFIG.cache.confiKey, serialized, BP_CONFIG.cache.ttlSeconds);
    } catch (error) {
      BP_logWarn_('Cache CONFI falhou ao gravar', {
        error: BP_getErrorMessage_(error),
      });
    }
  }

  return lists;
}


function BP_invalidateConfiCache_() {
  CacheService.getDocumentCache().remove(BP_CONFIG.cache.confiKey);
}


function BP_existeNaLista_(sheet, col, value) {
  const alvo = BP_normalizarComparacao_(value);

  if (!alvo) return false;

  const lastRow = Math.max(sheet.getLastRow(), BP_CONFIG.rows.confiFirstData);
  const numRows = Math.max(lastRow - BP_CONFIG.rows.confiFirstData + 1, 1);

  const values = sheet
    .getRange(BP_CONFIG.rows.confiFirstData, col, numRows, 1)
    .getValues()
    .flat();

  return values.some(function (item) {
    return BP_normalizarComparacao_(item) === alvo;
  });
}


function BP_contarItensLista_(sheet, col) {
  const lastRow = Math.max(sheet.getLastRow(), BP_CONFIG.rows.confiFirstData);
  const numRows = Math.max(lastRow - BP_CONFIG.rows.confiFirstData + 1, 1);

  return sheet
    .getRange(BP_CONFIG.rows.confiFirstData, col, numRows, 1)
    .getValues()
    .flat()
    .map(BP_texto_)
    .filter(Boolean)
    .length;
}


function BP_contarItensListas_(sheet, columns) {
  const lastRow = Math.max(sheet.getLastRow(), BP_CONFIG.rows.confiFirstData);
  const numRows = Math.max(lastRow - BP_CONFIG.rows.confiFirstData + 1, 1);
  const maxColumn = Math.max.apply(null, columns);
  const totals = {};

  columns.forEach(function (column) {
    totals[column] = 0;
  });

  sheet
    .getRange(BP_CONFIG.rows.confiFirstData, 1, numRows, maxColumn)
    .getValues()
    .forEach(function (row) {
      columns.forEach(function (column) {
        if (BP_texto_(row[column - 1])) {
          totals[column] += 1;
        }
      });
    });

  return totals;
}


function BP_primeiraLinhaVaziaNaColuna_(sheet, col, startRow) {
  const maxRows = sheet.getMaxRows();

  const values = sheet
    .getRange(startRow, col, maxRows - startRow + 1, 1)
    .getValues();

  const index = values.findIndex(function (row) {
    return row[0] === '';
  });

  if (index === -1) {
    sheet.insertRowsAfter(maxRows, 100);
    return maxRows + 1;
  }

  return startRow + index;
}


function BP_pushIfNotBlank_(target, value) {
  const text = BP_texto_(value);

  if (text) {
    target.push(text);
  }
}


function BP_toNormalizedSet_(values) {
  return new Set(
    values
      .map(BP_normalizarComparacao_)
      .filter(Boolean)
  );
}


/****************************************************
 * CHAVE DE DUPLICIDADE — BOLETOS
 ****************************************************/

function BP_chaveBoleto_(data, loja, empresa, descricao, valor) {
  const tz = Session.getScriptTimeZone();
  const dataKey = Utilities.formatDate(new Date(data), tz, 'yyyyMMdd');

  return [
    dataKey,
    BP_normalizarComparacao_(loja),
    BP_normalizarComparacao_(empresa),
    BP_normalizarComparacao_(descricao),
    Number(valor).toFixed(2),
  ].join('|');
}


/****************************************************
 * PARSING / NORMALIZAÇÃO — BOLETOS
 ****************************************************/

function BP_texto_(value) {
  return String(value ?? '').trim();
}


function BP_normalizarComparacao_(value) {
  return BP_texto_(value)
    .toUpperCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^A-Z0-9]/g, '');
}


function BP_parseData_(value) {
  if (BP_ehDataValida_(value)) return value;

  const raw = BP_texto_(value);

  if (!raw) return value;

  const match = raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);

  if (!match) return value;

  const dia = Number(match[1]);
  const mes = Number(match[2]);
  const ano = Number(match[3]);

  const date = new Date(ano, mes - 1, dia);

  if (
    date.getFullYear() === ano &&
    date.getMonth() === mes - 1 &&
    date.getDate() === dia
  ) {
    return date;
  }

  return value;
}


function BP_parseDataOpcional_(value) {
  if (value === '' || value === null || value === undefined) return '';
  return BP_parseData_(value);
}


function BP_ehDataValida_(value) {
  return value instanceof Date && !isNaN(value.getTime());
}


function BP_parseValor_(value) {
  if (typeof value === 'number') return value;

  const raw = BP_texto_(value);

  if (!raw) return NaN;

  const isNegative = raw.includes('-') || /^\(.*\)$/.test(raw);

  const clean = raw
    .replace(/[R$\s]/g, '')
    .replace(/[()]/g, '')
    .replace(/-/g, '')
    .replace(/\./g, '')
    .replace(',', '.');

  const number = Number(clean);

  if (!Number.isFinite(number)) return NaN;

  return isNegative ? -number : number;
}


function BP_parseValorOpcional_(value) {
  if (value === '' || value === null || value === undefined) return '';

  const parsed = BP_parseValor_(value);

  return Number.isFinite(parsed) ? parsed : NaN;
}


/****************************************************
 * LOG / ERROS — BOLETOS
 ****************************************************/

function BP_criarOperationId_() {
  return Utilities.getUuid().slice(0, 8);
}


function BP_logInfo_(message, payload) {
  console.info(JSON.stringify({
    module: 'BOLETOS',
    level: 'INFO',
    message: message,
    payload: payload || {},
    timestamp: new Date().toISOString(),
  }));
}


function BP_logWarn_(message, payload) {
  console.warn(JSON.stringify({
    module: 'BOLETOS',
    level: 'WARN',
    message: message,
    payload: payload || {},
    timestamp: new Date().toISOString(),
  }));
}


function BP_showError_(title, error) {
  const message = BP_getErrorMessage_(error);

  console.error(JSON.stringify({
    module: 'BOLETOS',
    level: 'ERROR',
    title: title,
    message: message,
    stack: error && error.stack ? error.stack : '',
    timestamp: new Date().toISOString(),
  }));

  SpreadsheetApp.getUi().alert(title + ':\n\n' + message);
}


function BP_getErrorMessage_(error) {
  if (!error) return 'Erro desconhecido.';
  if (error.message) return error.message;
  return String(error);
}
