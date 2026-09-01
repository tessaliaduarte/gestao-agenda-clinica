const CONFIG_FERIADOS_AUTOMATICOS = {
  ABAS: {
    FERIADOS: 'Calendário da Prefeitura',
    BLOQUEIOS: 'Bloqueios'
  },

  FERIADOS: {
    DATA: 1,
    DESCRICAO: 2,
    ATENDIMENTO: 4
  },

  BLOQUEIOS: {
    DATA: 1,
    HORARIO: 2,
    FISIOTERAPEUTA: 3,
    ABRANGENCIA: 4,
    MOTIVO: 5,
    ACAO: 6,
    STATUS: 7
  },

  PRIMEIRA_LINHA_DADOS: 2,
  QUANTIDADE_COLUNAS_BLOQUEIOS: 7
};


/**
 * Instala o gatilho que transforma um feriado cadastrado
 * em bloqueio de dia inteiro.
 *
 * Execute manualmente somente uma vez.
 */
function instalarGatilhoFeriadosAutomaticos() {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  try {
    validarDependenciasFeriadosAutomaticos_(
      ss
    );

    removerGatilhosFeriadosAutomaticos_();

    ScriptApp
      .newTrigger(
        'processarEdicaoFeriadosAutomaticos'
      )
      .forSpreadsheet(ss)
      .onEdit()
      .create();

    ss.toast(
      'O processamento automático de feriados foi ativado.',
      'Feriados automáticos',
      6
    );
  } catch (erro) {
    ss.toast(
      erro && erro.message
        ? erro.message
        : String(erro),
      'Erro ao ativar feriados',
      10
    );

    throw erro;
  }
}


/**
 * Remove apenas gatilhos anteriores deste mesmo módulo.
 */
function removerGatilhosFeriadosAutomaticos_() {
  ScriptApp
    .getProjectTriggers()
    .forEach(function (gatilho) {
      if (
        gatilho.getHandlerFunction() ===
          'processarEdicaoFeriadosAutomaticos'
      ) {
        ScriptApp.deleteTrigger(
          gatilho
        );
      }
    });
}


/**
 * Gatilho executado ao editar o Calendário da Prefeitura.
 *
 * O processamento acontece quando o campo Atendimento
 * da linha passa a ser "Não".
 */
function processarEdicaoFeriadosAutomaticos(e) {
  if (!e || !e.range) {
    return;
  }

  const aba = e.range.getSheet();
  const linhaInicial = e.range.getRow();
  const linhaFinal =
    linhaInicial +
    e.range.getNumRows() -
    1;
  const colunaInicial =
    e.range.getColumn();
  const colunaFinal =
    colunaInicial +
    e.range.getNumColumns() -
    1;

  if (
    aba.getName() !==
      CONFIG_FERIADOS_AUTOMATICOS
        .ABAS.FERIADOS ||

    linhaFinal <
      CONFIG_FERIADOS_AUTOMATICOS
        .PRIMEIRA_LINHA_DADOS ||

    colunaInicial >
      CONFIG_FERIADOS_AUTOMATICOS
        .FERIADOS.ATENDIMENTO ||

    colunaFinal <
      CONFIG_FERIADOS_AUTOMATICOS
        .FERIADOS.ATENDIMENTO
  ) {
    return;
  }

  const ss =
    e.source ||
    SpreadsheetApp.getActiveSpreadsheet();

  for (
    let linha = Math.max(
      linhaInicial,
      CONFIG_FERIADOS_AUTOMATICOS
        .PRIMEIRA_LINHA_DADOS
    );
    linha <= linhaFinal;
    linha++
  ) {
    const atendimento =
      normalizarTextoFeriadoAutomatico_(
        aba
          .getRange(
            linha,
            CONFIG_FERIADOS_AUTOMATICOS
              .FERIADOS.ATENDIMENTO
          )
          .getDisplayValue()
      );

    if (atendimento !== 'nao') {
      continue;
    }

    try {
      const resultado =
        processarLinhaFeriadoAutomatico_(
          ss,
          linha
        );

      ss.toast(
        montarMensagemResultadoFeriado_(
          resultado
        ),
        'Feriado processado',
        8
      );
    } catch (erro) {
      ss.toast(
        erro && erro.message
          ? erro.message
          : String(erro),
        'Erro ao processar feriado',
        10
      );

      console.error(erro);
    }
  }
}


/**
 * Processa manualmente a linha selecionada no
 * Calendário da Prefeitura.
 *
 * Use esta função para feriados que já estavam cadastrados
 * antes da instalação do gatilho.
 */
function processarFeriadoSelecionadoAgora() {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const aba =
    ss.getActiveSheet();

  const intervalo =
    aba.getActiveRange();

  try {
    if (
      aba.getName() !==
        CONFIG_FERIADOS_AUTOMATICOS
          .ABAS.FERIADOS
    ) {
      throw new Error(
        'Abra a aba "Calendário da Prefeitura" e selecione a linha do feriado.'
      );
    }

    if (
      !intervalo ||
      intervalo.getRow() <
        CONFIG_FERIADOS_AUTOMATICOS
          .PRIMEIRA_LINHA_DADOS
    ) {
      throw new Error(
        'Selecione uma célula na linha do feriado que será processado.'
      );
    }

    const resultado =
      processarLinhaFeriadoAutomatico_(
        ss,
        intervalo.getRow()
      );

    ss.toast(
      montarMensagemResultadoFeriado_(
        resultado
      ),
      'Feriado processado',
      10
    );
  } catch (erro) {
    ss.toast(
      erro && erro.message
        ? erro.message
        : String(erro),
      'Erro ao processar feriado',
      10
    );

    throw erro;
  }
}
/**
 * Valida o feriado, cria ou atualiza seu bloqueio
 * e executa o fluxo já existente de bloqueios.
 */
function processarLinhaFeriadoAutomatico_(
  ss,
  linhaFeriado
) {
  validarDependenciasFeriadosAutomaticos_(
    ss
  );

  const abaFeriados =
    ss.getSheetByName(
      CONFIG_FERIADOS_AUTOMATICOS
        .ABAS.FERIADOS
    );

  const feriado =
    lerFeriadoAutomatico_(
      abaFeriados,
      linhaFeriado
    );

  validarFeriadoAutomatico_(
    feriado
  );

  const linhaBloqueio =
    criarOuAtualizarBloqueioFeriado_(
      ss,
      feriado
    );

  const resultado =
    processarLinhaBloqueioComLock_(
      ss,
      linhaBloqueio
    );

  atualizarModulosAposBloqueio_(
    ss,
    resultado.datasAfetadas
  );

  registrarProcessamentoFeriado_(
    abaFeriados,
    feriado,
    linhaBloqueio,
    resultado
  );

  return {
    data: feriado.data,
    descricao: feriado.descricao,
    linhaBloqueio: linhaBloqueio,
    total: resultado.total || 0,
    sessoes: resultado.sessoes || 0,
    avaliacoes: resultado.avaliacoes || 0
  };
}


function lerFeriadoAutomatico_(
  abaFeriados,
  linha
) {
  const colunas =
    CONFIG_FERIADOS_AUTOMATICOS
      .FERIADOS;

  const valores =
    abaFeriados
      .getRange(
        linha,
        1,
        1,
        Math.max(
          colunas.ATENDIMENTO,
          abaFeriados.getLastColumn()
        )
      )
      .getValues()[0];

  return {
    linha: linha,
    data:
      valores[colunas.DATA - 1],

    descricao: String(
      valores[
        colunas.DESCRICAO - 1
      ] || 'Feriado'
    ).trim(),

    atendimento: String(
      valores[
        colunas.ATENDIMENTO - 1
      ] || ''
    ).trim()
  };
}


function validarFeriadoAutomatico_(
  feriado
) {
  if (!(feriado.data instanceof Date)) {
    throw new Error(
      'A linha selecionada não possui uma data válida.'
    );
  }

  if (
    normalizarTextoFeriadoAutomatico_(
      feriado.atendimento
    ) !== 'nao'
  ) {
    throw new Error(
      'O feriado somente será processado quando Atendimento estiver como "Não".'
    );
  }
}


/**
 * Cria um bloqueio de dia inteiro ou reaproveita o bloqueio
 * do mesmo feriado, evitando linhas duplicadas.
 */
function criarOuAtualizarBloqueioFeriado_(
  ss,
  feriado
) {
  const abaBloqueios =
    ss.getSheetByName(
      CONFIG_FERIADOS_AUTOMATICOS
        .ABAS.BLOQUEIOS
    );

  const linhaExistente =
    localizarBloqueioFeriadoExistente_(
      abaBloqueios,
      feriado
    );

  const linha =
    linhaExistente ||
    Math.max(
      abaBloqueios.getLastRow() + 1,
      CONFIG_FERIADOS_AUTOMATICOS
        .PRIMEIRA_LINHA_DADOS
    );

  const motivo =
    'Feriado — ' +
    feriado.descricao;

  abaBloqueios
    .getRange(
      linha,
      1,
      1,
      CONFIG_FERIADOS_AUTOMATICOS
        .QUANTIDADE_COLUNAS_BLOQUEIOS
    )
    .setValues([[
      new Date(feriado.data),
      '',
      '',
      'Dia inteiro',
      motivo,
      'Sinalizar pacientes já agendados',
      'Ativo'
    ]]);

  abaBloqueios
    .getRange(
      linha,
      CONFIG_FERIADOS_AUTOMATICOS
        .BLOQUEIOS.DATA
    )
    .setNumberFormat(
      'dd/MM/yyyy'
    );

  abaBloqueios
    .getRange(
      linha,
      CONFIG_FERIADOS_AUTOMATICOS
        .BLOQUEIOS.MOTIVO
    )
    .setNote(
      'Bloqueio criado automaticamente a partir da linha ' +
        feriado.linha +
        ' da aba Calendário da Prefeitura.'
    );

  SpreadsheetApp.flush();

  return linha;
}


function localizarBloqueioFeriadoExistente_(
  abaBloqueios,
  feriado
) {
  const ultimaLinha =
    abaBloqueios.getLastRow();

  if (
    ultimaLinha <
      CONFIG_FERIADOS_AUTOMATICOS
        .PRIMEIRA_LINHA_DADOS
  ) {
    return 0;
  }

  const dados =
    abaBloqueios
      .getRange(
        CONFIG_FERIADOS_AUTOMATICOS
          .PRIMEIRA_LINHA_DADOS,
        1,
        ultimaLinha -
          CONFIG_FERIADOS_AUTOMATICOS
            .PRIMEIRA_LINHA_DADOS +
          1,
        CONFIG_FERIADOS_AUTOMATICOS
          .QUANTIDADE_COLUNAS_BLOQUEIOS
      )
      .getValues();

  const chaveData =
    chaveDataFeriadoAutomatico_(
      feriado.data
    );

  for (
    let indice = 0;
    indice < dados.length;
    indice++
  ) {
    const linha = dados[indice];

    const data =
      linha[
        CONFIG_FERIADOS_AUTOMATICOS
          .BLOQUEIOS.DATA - 1
      ];

    const abrangencia =
      normalizarTextoFeriadoAutomatico_(
        linha[
          CONFIG_FERIADOS_AUTOMATICOS
            .BLOQUEIOS.ABRANGENCIA - 1
        ]
      );

    const motivo =
      normalizarTextoFeriadoAutomatico_(
        linha[
          CONFIG_FERIADOS_AUTOMATICOS
            .BLOQUEIOS.MOTIVO - 1
        ]
      );

    if (
      data instanceof Date &&

      chaveDataFeriadoAutomatico_(
        data
      ) === chaveData &&

      abrangencia ===
        'dia inteiro' &&

      motivo.indexOf(
        'feriado'
      ) === 0
    ) {
      return (
        indice +
        CONFIG_FERIADOS_AUTOMATICOS
          .PRIMEIRA_LINHA_DADOS
      );
    }
  }

  return 0;
}
function registrarProcessamentoFeriado_(
  abaFeriados,
  feriado,
  linhaBloqueio,
  resultado
) {
  abaFeriados
    .getRange(
      feriado.linha,
      CONFIG_FERIADOS_AUTOMATICOS
        .FERIADOS.ATENDIMENTO
    )
    .setNote(
      'Feriado processado automaticamente em ' +
        Utilities.formatDate(
          new Date(),
          Session.getScriptTimeZone(),
          'dd/MM/yyyy HH:mm'
        ) +
        '. Linha do bloqueio: ' +
        linhaBloqueio +
        '. Sessões afetadas: ' +
        (resultado.sessoes || 0) +
        '. Avaliações afetadas: ' +
        (resultado.avaliacoes || 0) +
        '.'
    );
}


function validarDependenciasFeriadosAutomaticos_(
  ss
) {
  const abaFeriados =
    ss.getSheetByName(
      CONFIG_FERIADOS_AUTOMATICOS
        .ABAS.FERIADOS
    );

  const abaBloqueios =
    ss.getSheetByName(
      CONFIG_FERIADOS_AUTOMATICOS
        .ABAS.BLOQUEIOS
    );

  if (!abaFeriados) {
    throw new Error(
      'A aba "Calendário da Prefeitura" não foi encontrada.'
    );
  }

  if (!abaBloqueios) {
    throw new Error(
      'A aba "Bloqueios" não foi encontrada.'
    );
  }

  if (
    typeof
      processarLinhaBloqueioComLock_ !==
      'function'
  ) {
    throw new Error(
      'A função de processamento de bloqueios não foi encontrada. Confira o arquivo BloqueiosAutomaticos.gs.'
    );
  }

  if (
    typeof
      atualizarModulosAposBloqueio_ !==
      'function'
  ) {
    throw new Error(
      'A função de atualização após bloqueio não foi encontrada. Confira o arquivo BloqueiosAutomaticos.gs.'
    );
  }
}


function montarMensagemResultadoFeriado_(
  resultado
) {
  return (
    'Data: ' +
    Utilities.formatDate(
      resultado.data,
      Session.getScriptTimeZone(),
      'dd/MM/yyyy'
    ) +
    '. Sessões afetadas: ' +
    resultado.sessoes +
    '. Avaliações afetadas: ' +
    resultado.avaliacoes +
    '.'
  );
}


function chaveDataFeriadoAutomatico_(
  data
) {
  return Utilities.formatDate(
    new Date(data),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );
}


function normalizarTextoFeriadoAutomatico_(
  valor
) {
  return String(valor || '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(
      /[\u0300-\u036f]/g,
      ''
    )
    .replace(
      /[–—−]/g,
      '-'
    )
    .replace(
      /\s+/g,
      ' '
    );
}
