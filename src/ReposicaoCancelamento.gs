const CONFIG_REPOSICAO_CANCELAMENTO = {
  ABAS: {
    AGENDA: 'Agenda',
    AGENDAMENTOS: 'Agendamentos',
    CADASTRO: 'Cadastro de Pacientes',
    FERIADOS: 'Calendário da Prefeitura',
    BLOQUEIOS: 'Bloqueios'
  },

  AGENDA: {
    PRIMEIRA_LINHA: 5,
    COLUNA_STATUS: 4,
    COLUNA_ID_AGENDAMENTO: 6
  },

  AGENDAMENTOS: {
    ID_AGENDAMENTO: 1,
    ID_PACIENTE: 2,
    PRONTUARIO: 3,
    NOME_PACIENTE: 4,
    ID_CICLO: 5,
    CICLO_NUMERO: 6,
    DATA: 7,
    DIA: 8,
    HORARIO: 9,
    FISIOTERAPEUTA: 10,
    TIPO_GRUPO: 11,
    EVENTO: 12,
    NUMERO_SESSAO: 13,
    TOTAL_PRESCRITO: 14,
    LIMITE_GRUPO: 15,
    STATUS: 16,
    MOTIVO: 17,
    CONTA_COMO_SESSAO: 18,
    AVISAR_PACIENTE: 19,
    CRIADO_EM: 20,
    ATUALIZADO_EM: 21,
    FATURAVEL: 22
  },

  CADASTRO: {
    ID_PACIENTE: 1,
    DATA_PREVISTA_TERMINO: 20
  },

  QUANTIDADE_COLUNAS_AGENDAMENTOS: 22,
  PRAZO_MAXIMO_BUSCA_DIAS: 365,

  STATUS_QUE_OCUPAM_VAGA: [
    'agendado',
    'compareceu',
    'falta justificada',
    'falta nao justificada'
  ]
};


/**
 * Instala o gatilho necessário para gerar reposições
 * automaticamente quando houver cancelamento pela clínica.
 *
 * Esta função deve ser executada manualmente apenas uma vez.
 */
function instalarGatilhoReposicaoCancelamento() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    removerGatilhosDuplicadosReposicaoCancelamento_();

    ScriptApp.newTrigger(
      'processarEdicaoReposicaoCancelamento'
    )
      .forSpreadsheet(ss)
      .onEdit()
      .create();

    ui.alert(
      'Gatilho instalado',
      'A reposição automática por cancelamento da clínica foi ativada com sucesso.',
      ui.ButtonSet.OK
    );
  } catch (erro) {
    ui.alert(
      'Erro ao instalar o gatilho',
      erro && erro.message
        ? erro.message
        : String(erro),
      ui.ButtonSet.OK
    );
  }
}


/**
 * Remove gatilhos anteriores desta mesma função,
 * evitando que uma reposição seja criada mais de uma vez.
 */
function removerGatilhosDuplicadosReposicaoCancelamento_() {
  const gatilhos = ScriptApp.getProjectTriggers();

  gatilhos.forEach(function(gatilho) {
    if (
      gatilho.getHandlerFunction() ===
      'processarEdicaoReposicaoCancelamento'
    ) {
      ScriptApp.deleteTrigger(gatilho);
    }
  });
}


/**
 * Função executada pelo gatilho instalado.
 */
function processarEdicaoReposicaoCancelamento(e) {
  if (!e || !e.range) {
    return;
  }

  const abaEditada = e.range.getSheet();

  if (
    abaEditada.getName() !==
    CONFIG_REPOSICAO_CANCELAMENTO.ABAS.AGENDA
  ) {
    return;
  }

  if (
    e.range.getRow() <
      CONFIG_REPOSICAO_CANCELAMENTO
        .AGENDA
        .PRIMEIRA_LINHA ||
    e.range.getColumn() !==
      CONFIG_REPOSICAO_CANCELAMENTO
        .AGENDA
        .COLUNA_STATUS
  ) {
    return;
  }

  const novoStatus = normalizarTextoReposicao_(
    e.value || e.range.getValue()
  );

  if (novoStatus !== 'cancelado pela clinica') {
    return;
  }

  const lock = LockService.getDocumentLock();
  let bloqueioObtido = false;

  try {
    lock.waitLock(30000);
    bloqueioObtido = true;

    const ss = e.source ||
      SpreadsheetApp.getActiveSpreadsheet();

    const idAgendamento = String(
      abaEditada.getRange(
        e.range.getRow(),
        CONFIG_REPOSICAO_CANCELAMENTO
          .AGENDA
          .COLUNA_ID_AGENDAMENTO
      ).getValue() || ''
    ).trim();

    if (!idAgendamento) {
      throw new Error(
        'O ID do agendamento não foi encontrado na Agenda.'
      );
    }

    const resultadoReposicao =
      gerarReposicaoCancelamentoClinica_(
        ss,
        idAgendamento
      );

    SpreadsheetApp.flush();

    if (
      resultadoReposicao &&
      resultadoReposicao.semReposicao
    ) {
      ss.toast(
        'A sessão foi cancelada pela clínica, mas não foi encontrada vaga para reposição automática. O caso ficará pendente para regularização.',
        'Reposição pendente',
        9
      );
    } else {
      ss.toast(
        'A sessão cancelada foi reposta automaticamente no final do ciclo.',
        'Reposição gerada',
        6
      );
    }
  } catch (erro) {
    const ss = e.source ||
      SpreadsheetApp.getActiveSpreadsheet();

    ss.toast(
      erro && erro.message
        ? erro.message
        : String(erro),
      'Erro ao gerar reposição',
      8
    );

    console.error(erro);
  } finally {
    if (bloqueioObtido) {
      try {
        lock.releaseLock();
      } catch (erroLock) {
        // O bloqueio será liberado automaticamente.
      }
    }
  }
}


/**
 * Gera a reposição de uma sessão cancelada pela clínica.
 */function gerarReposicaoCancelamentoClinica_(
  ss,
  idAgendamentoOriginal
) {
  const abas =
    obterAbasReposicaoCancelamento_(
      ss
    );

  const original =
    localizarAgendamentoReposicao_(
      abas.agendamentos,
      idAgendamentoOriginal
    );

  if (!original) {
    throw new Error(
      'O agendamento cancelado não foi encontrado.'
    );
  }

  if (
    normalizarTextoReposicao_(
      original.evento
    ) !== 'sessao'
  ) {
    throw new Error(
      'A reposição automática é permitida somente para sessões.'
    );
  }

  /*
   * Impede a criação de uma segunda reposição
   * para o mesmo cancelamento.
   */
  const reposicaoExistente =
    localizarReposicaoExistente_(
      abas.agendamentos,
      original.idAgendamento
    );

  if (reposicaoExistente) {
    atualizarSessaoOriginalComoCancelada_(
      abas.agendamentos,
      original
    );

    atualizarMotivoSessaoOriginal_(
      abas.agendamentos,
      original.linha,
      reposicaoExistente.idAgendamento
    );

    return reposicaoExistente;
  }

  const dadosCiclo =
    obterDadosCicloReposicao_(
      abas.agendamentos,
      original
    );

  if (
    dadosCiclo.diasSemana.length === 0
  ) {
    throw new Error(
      'Não foi possível identificar os dias semanais do ciclo.'
    );
  }

  const feriados =
    lerFeriadosReposicao_(
      abas.feriados
    );

  const bloqueios =
    lerBloqueiosReposicao_(
      abas.bloqueios
    );

  const todosAgendamentos =
    lerTodosAgendamentosReposicao_(
      abas.agendamentos
    );

  /*
   * Primeiro procura a nova data.
   *
   * Dessa forma, o módulo sabe exatamente se conseguiu
   * repor antes de finalizar todo o fluxo.
   */
  const novaData =
    localizarProximaDataReposicao_(
      original,
      dadosCiclo,
      feriados,
      bloqueios,
      todosAgendamentos
    );

  /*
   * O cancelamento da clínica continua válido mesmo
   * quando nenhuma vaga de reposição é encontrada.
   *
   * Nesse cenário, a sessão original não conta como
   * realizada nem como faturável e recebe uma indicação
   * explícita de que a reposição precisa ser regularizada.
   */
  atualizarSessaoOriginalComoCancelada_(
    abas.agendamentos,
    original
  );

  if (!novaData) {
    registrarReposicaoPendenteSemVaga_(
      abas.agendamentos,
      original,
      CONFIG_REPOSICAO_CANCELAMENTO
        .PRAZO_MAXIMO_BUSCA_DIAS
    );

    return {
      idAgendamento:
        original.idAgendamento,
      data: null,
      linha: original.linha,
      semReposicao: true,
      motivo:
        'Não foi encontrada vaga automática para reposição nos próximos ' +
        CONFIG_REPOSICAO_CANCELAMENTO
          .PRAZO_MAXIMO_BUSCA_DIAS +
        ' dias.'
    };
  }

  const novoIdAgendamento =
    gerarProximoIdAgendamentoReposicao_(
      abas.agendamentos
    );

  const agora = new Date();

  const novaLinha =
    montarLinhaReposicao_(
      original,
      novoIdAgendamento,
      novaData,
      agora
    );

  const novaLinhaPlanilha =
    Math.max(
      abas.agendamentos.getLastRow() + 1,
      2
    );

  abas.agendamentos.getRange(
    novaLinhaPlanilha,
    1,
    1,
    CONFIG_REPOSICAO_CANCELAMENTO
      .QUANTIDADE_COLUNAS_AGENDAMENTOS
  ).setValues([novaLinha]);

  abas.agendamentos.getRange(
    novaLinhaPlanilha,
    CONFIG_REPOSICAO_CANCELAMENTO
      .AGENDAMENTOS
      .DATA
  )
    .setNumberFormat('dd/MM/yyyy');

  abas.agendamentos.getRange(
    novaLinhaPlanilha,
    CONFIG_REPOSICAO_CANCELAMENTO
      .AGENDAMENTOS
      .HORARIO
  )
    .setNumberFormat('HH:mm');

  abas.agendamentos.getRange(
    novaLinhaPlanilha,
    CONFIG_REPOSICAO_CANCELAMENTO
      .AGENDAMENTOS
      .CRIADO_EM,
    1,
    2
  )
    .setNumberFormat(
      'dd/MM/yyyy HH:mm'
    );

  atualizarMotivoSessaoOriginal_(
    abas.agendamentos,
    original.linha,
    novoIdAgendamento
  );

  atualizarDataTerminoPacienteReposicao_(
    abas.cadastro,
    original.idPaciente,
    novaData
  );

  return {
    idAgendamento:
      novoIdAgendamento,
    data:
      novaData,
    linha:
      novaLinhaPlanilha,
    semReposicao:
      false
  };
}


/**
 * Obtém e valida as abas necessárias.
 */
function obterAbasReposicaoCancelamento_(ss) {
  const nomes =
    CONFIG_REPOSICAO_CANCELAMENTO.ABAS;

  const abas = {
    agenda: ss.getSheetByName(
      nomes.AGENDA
    ),

    agendamentos: ss.getSheetByName(
      nomes.AGENDAMENTOS
    ),

    cadastro: ss.getSheetByName(
      nomes.CADASTRO
    ),

    feriados: ss.getSheetByName(
      nomes.FERIADOS
    ),

    bloqueios: ss.getSheetByName(
      nomes.BLOQUEIOS
    )
  };

  Object.keys(abas).forEach(function(chave) {
    if (!abas[chave]) {
      throw new Error(
        'A aba necessária "' +
          nomes[
            {
              agenda: 'AGENDA',
              agendamentos: 'AGENDAMENTOS',
              cadastro: 'CADASTRO',
              feriados: 'FERIADOS',
              bloqueios: 'BLOQUEIOS'
            }[chave]
          ] +
          '" não foi encontrada.'
      );
    }
  });

  return abas;
}


/**
 * Localiza um agendamento pelo ID.
 */function localizarAgendamentoReposicao_(
  abaAgendamentos,
  idAgendamento
) {
  const ultimaLinha =
    abaAgendamentos.getLastRow();

  if (ultimaLinha < 2) {
    return null;
  }

  const dados = abaAgendamentos
    .getRange(
      2,
      1,
      ultimaLinha - 1,
      CONFIG_REPOSICAO_CANCELAMENTO
        .QUANTIDADE_COLUNAS_AGENDAMENTOS
    )
    .getValues();

  const idProcurado =
    normalizarTextoReposicao_(
      idAgendamento
    );

  for (
    let indice = 0;
    indice < dados.length;
    indice++
  ) {
    const linha = dados[indice];

    const idAtual =
      normalizarTextoReposicao_(
        linha[
          CONFIG_REPOSICAO_CANCELAMENTO
            .AGENDAMENTOS
            .ID_AGENDAMENTO - 1
        ]
      );

    if (idAtual === idProcurado) {
      return montarObjetoAgendamentoReposicao_(
        linha,
        indice + 2
      );
    }
  }

  return null;
}


/**
 * Converte uma linha da aba Agendamentos em objeto.
 */
function montarObjetoAgendamentoReposicao_(
  linha,
  numeroLinha
) {
  const colunas =
    CONFIG_REPOSICAO_CANCELAMENTO
      .AGENDAMENTOS;

  return {
    linha: numeroLinha,

    idAgendamento: String(
      linha[colunas.ID_AGENDAMENTO - 1] || ''
    ).trim(),

    idPaciente: String(
      linha[colunas.ID_PACIENTE - 1] || ''
    ).trim(),

    prontuario: String(
      linha[colunas.PRONTUARIO - 1] || ''
    ).trim(),

    nomePaciente: String(
      linha[colunas.NOME_PACIENTE - 1] || ''
    ).trim(),

    idCiclo: String(
      linha[colunas.ID_CICLO - 1] || ''
    ).trim(),

    cicloNumero:
      Number(
        linha[colunas.CICLO_NUMERO - 1]
      ) || 0,

    data:
      linha[colunas.DATA - 1],

    dia: String(
      linha[colunas.DIA - 1] || ''
    ).trim(),

    horario:
      linha[colunas.HORARIO - 1],

    fisioterapeuta: String(
      linha[colunas.FISIOTERAPEUTA - 1] || ''
    ).trim(),

    tipoGrupo: String(
      linha[colunas.TIPO_GRUPO - 1] || ''
    ).trim(),

    evento: String(
      linha[colunas.EVENTO - 1] || ''
    ).trim(),

    numeroSessao:
      Number(
        linha[colunas.NUMERO_SESSAO - 1]
      ) || 0,

    totalPrescrito:
      Number(
        linha[colunas.TOTAL_PRESCRITO - 1]
      ) || 0,

    limiteGrupo:
      Number(
        linha[colunas.LIMITE_GRUPO - 1]
      ) || 0,

    status: String(
      linha[colunas.STATUS - 1] || ''
    ).trim(),

    motivo: String(
      linha[colunas.MOTIVO - 1] || ''
    ).trim()
  };
}


/**
 * Atualiza o registro original cancelado.
 */
function atualizarSessaoOriginalComoCancelada_(
  abaAgendamentos,
  original
) {
  const agora = new Date();
  const colunas =
    CONFIG_REPOSICAO_CANCELAMENTO
      .AGENDAMENTOS;

  abaAgendamentos.getRange(
    original.linha,
    colunas.STATUS
  ).setValue(
    'Cancelado pela Clínica'
  );

  abaAgendamentos.getRange(
    original.linha,
    colunas.CONTA_COMO_SESSAO
  ).setValue(
    'Não'
  );

  abaAgendamentos.getRange(
    original.linha,
    colunas.FATURAVEL
  ).setValue(
    'Não'
  );

  abaAgendamentos.getRange(
    original.linha,
    colunas.ATUALIZADO_EM
  )
    .setValue(agora)
    .setNumberFormat('dd/MM/yyyy HH:mm');
}


/**
 * Registra de forma explícita que a clínica cancelou
 * a sessão, porém não foi encontrada uma vaga automática
 * para reposição.
 *
 * A sessão permanece sem contar e sem faturamento.
 * O texto no Motivo permite que Pendências e relatórios
 * identifiquem facilmente o caso.
 */
function registrarReposicaoPendenteSemVaga_(
  abaAgendamentos,
  original,
  prazoDias
) {
  const colunas =
    CONFIG_REPOSICAO_CANCELAMENTO
      .AGENDAMENTOS;

  const motivoAtual =
    String(
      abaAgendamentos.getRange(
        original.linha,
        colunas.MOTIVO
      ).getValue() || ''
    ).trim();

  const identificacao =
    'Reposição pendente — sem vaga automática nos próximos ' +
    prazoDias +
    ' dias';

  if (
    normalizarTextoReposicao_(
      motivoAtual
    ).indexOf(
      normalizarTextoReposicao_(
        identificacao
      )
    ) !== -1
  ) {
    return;
  }

  const novoMotivo =
    motivoAtual
      ? motivoAtual +
        ' | ' +
        identificacao
      : 'Cancelado pela clínica | ' +
        identificacao;

  abaAgendamentos.getRange(
    original.linha,
    colunas.MOTIVO
  ).setValue(
    novoMotivo
  );

  abaAgendamentos.getRange(
    original.linha,
    colunas.AVISAR_PACIENTE
  ).setValue(
    'Sim'
  );

  abaAgendamentos.getRange(
    original.linha,
    colunas.ATUALIZADO_EM
  )
    .setValue(
      new Date()
    )
    .setNumberFormat(
      'dd/MM/yyyy HH:mm'
    );
}


/**
 * Procura uma reposição já criada para o agendamento.
 */
function localizarReposicaoExistente_(
  abaAgendamentos,
  idAgendamentoOriginal
) {
  const ultimaLinha =
    abaAgendamentos.getLastRow();

  if (ultimaLinha < 2) {
    return null;
  }

  const dados = abaAgendamentos
    .getRange(
      2,
      1,
      ultimaLinha - 1,
      CONFIG_REPOSICAO_CANCELAMENTO
        .QUANTIDADE_COLUNAS_AGENDAMENTOS
    )
    .getValues();

  const textoProcurado =
    normalizarTextoReposicao_(
      'Reposição automática de ' +
        idAgendamentoOriginal
    );

  for (
    let indice = 0;
    indice < dados.length;
    indice++
  ) {
    const motivo =
      normalizarTextoReposicao_(
        dados[indice][
          CONFIG_REPOSICAO_CANCELAMENTO
            .AGENDAMENTOS
            .MOTIVO - 1
        ]
      );

    if (motivo.indexOf(textoProcurado) !== -1) {
      return {
        linha: indice + 2,

        idAgendamento: String(
          dados[indice][
            CONFIG_REPOSICAO_CANCELAMENTO
              .AGENDAMENTOS
              .ID_AGENDAMENTO - 1
          ] || ''
        ).trim(),

        data:
          dados[indice][
            CONFIG_REPOSICAO_CANCELAMENTO
              .AGENDAMENTOS
              .DATA - 1
          ]
      };
    }
  }

  return null;
}


/**
 * Lê todos os registros do mesmo ciclo e identifica:
 * - dias da semana;
 * - última data agendada.
 */function obterDadosCicloReposicao_(
  abaAgendamentos,
  original
) {
  const ultimaLinha =
    abaAgendamentos.getLastRow();

  const dados = abaAgendamentos
    .getRange(
      2,
      1,
      ultimaLinha - 1,
      CONFIG_REPOSICAO_CANCELAMENTO
        .QUANTIDADE_COLUNAS_AGENDAMENTOS
    )
    .getValues();

  const diasSemana = [];
  let ultimaData = new Date(original.data);

  dados.forEach(function(linha) {
    const idCiclo = String(
      linha[
        CONFIG_REPOSICAO_CANCELAMENTO
          .AGENDAMENTOS
          .ID_CICLO - 1
      ] || ''
    ).trim();

    const evento =
      normalizarTextoReposicao_(
        linha[
          CONFIG_REPOSICAO_CANCELAMENTO
            .AGENDAMENTOS
            .EVENTO - 1
        ]
      );

    const horario =
      chaveHorarioReposicao_(
        linha[
          CONFIG_REPOSICAO_CANCELAMENTO
            .AGENDAMENTOS
            .HORARIO - 1
        ]
      );

    const fisioterapeuta =
      normalizarTextoReposicao_(
        linha[
          CONFIG_REPOSICAO_CANCELAMENTO
            .AGENDAMENTOS
            .FISIOTERAPEUTA - 1
        ]
      );

    const data =
      linha[
        CONFIG_REPOSICAO_CANCELAMENTO
          .AGENDAMENTOS
          .DATA - 1
      ];

    if (
      idCiclo !== original.idCiclo ||
      evento !== 'sessao' ||
      horario !==
        chaveHorarioReposicao_(
          original.horario
        ) ||
      fisioterapeuta !==
        normalizarTextoReposicao_(
          original.fisioterapeuta
        ) ||
      !(data instanceof Date)
    ) {
      return;
    }

    const diaSemana = data.getDay();

    if (
      diaSemana >= 1 &&
      diaSemana <= 5 &&
      diasSemana.indexOf(diaSemana) === -1
    ) {
      diasSemana.push(diaSemana);
    }

    if (data.getTime() > ultimaData.getTime()) {
      ultimaData = new Date(data);
    }
  });

  diasSemana.sort();

  return {
    diasSemana: diasSemana,
    ultimaData: ultimaData
  };
}


/**
 * Lê todos os agendamentos.
 */
function lerTodosAgendamentosReposicao_(
  abaAgendamentos
) {
  const ultimaLinha =
    abaAgendamentos.getLastRow();

  if (ultimaLinha < 2) {
    return [];
  }

  return abaAgendamentos
    .getRange(
      2,
      1,
      ultimaLinha - 1,
      CONFIG_REPOSICAO_CANCELAMENTO
        .QUANTIDADE_COLUNAS_AGENDAMENTOS
    )
    .getValues();
}


/**
 * Lê os feriados e dias sem atendimento.
 */
function lerFeriadosReposicao_(
  abaFeriados
) {
  const feriados = {};
  const ultimaLinha =
    abaFeriados.getLastRow();

  if (ultimaLinha < 2) {
    return feriados;
  }

  const quantidadeColunas = Math.max(
    abaFeriados.getLastColumn(),
    4
  );

  const dados = abaFeriados
    .getRange(
      2,
      1,
      ultimaLinha - 1,
      quantidadeColunas
    )
    .getValues();

  dados.forEach(function(linha) {
    const data = linha[0];

    if (!(data instanceof Date)) {
      return;
    }

    /*
     * A coluna D informa se há atendimento.
     * Quando estiver "Não", o dia é bloqueado.
     */
    const atendimento =
      normalizarTextoReposicao_(linha[3]);

    if (
      atendimento === 'nao' ||
      atendimento === 'não'
    ) {
      feriados[chaveDataReposicao_(data)] = true;
    }
  });

  return feriados;
}


/**
 * Lê os bloqueios cadastrados.
 */
function lerBloqueiosReposicao_(
  abaBloqueios
) {
  const bloqueios = [];
  const ultimaLinha =
    abaBloqueios.getLastRow();

  if (ultimaLinha < 2) {
    return bloqueios;
  }

  const quantidadeColunas = Math.max(
    abaBloqueios.getLastColumn(),
    7
  );

  const dados = abaBloqueios
    .getRange(
      2,
      1,
      ultimaLinha - 1,
      quantidadeColunas
    )
    .getValues();

  dados.forEach(function(linha) {
    const data = linha[0];

    if (!(data instanceof Date)) {
      return;
    }

    const status =
      normalizarTextoReposicao_(linha[6]);

    if (
      status &&
      status !== 'ativo' &&
      status !== 'bloqueado'
    ) {
      return;
    }

    bloqueios.push({
      data: chaveDataReposicao_(data),

      horario: linha[1]
        ? chaveHorarioReposicao_(linha[1])
        : '',

      fisioterapeuta:
        normalizarTextoReposicao_(linha[2]),

      abrangencia:
        normalizarTextoReposicao_(linha[3])
    });
  });

  return bloqueios;
}


/**
 * Procura a primeira data válida após o fim atual do ciclo.
 */function localizarProximaDataReposicao_(
  original,
  dadosCiclo,
  feriados,
  bloqueios,
  todosAgendamentos
) {
  const data = new Date(
    dadosCiclo.ultimaData
  );

  data.setHours(0, 0, 0, 0);
  data.setDate(data.getDate() + 1);

  let diasAnalisados = 0;

  while (
    diasAnalisados <
    CONFIG_REPOSICAO_CANCELAMENTO
      .PRAZO_MAXIMO_BUSCA_DIAS
  ) {
    diasAnalisados++;

    const diaSemana = data.getDay();

    if (
      dadosCiclo.diasSemana.indexOf(
        diaSemana
      ) !== -1 &&
      !feriados[
        chaveDataReposicao_(data)
      ] &&
      !estaBloqueadoReposicao_(
        bloqueios,
        data,
        original.horario,
        original.fisioterapeuta
      )
    ) {
      const disponibilidade =
        verificarDisponibilidadeReposicao_(
          todosAgendamentos,
          data,
          original
        );

      if (disponibilidade) {
        return new Date(data);
      }
    }

    data.setDate(data.getDate() + 1);
  }

  return null;
}


/**
 * Verifica bloqueios por data, horário e profissional.
 */
function estaBloqueadoReposicao_(
  bloqueios,
  data,
  horario,
  fisioterapeuta
) {
  const dataProcurada =
    chaveDataReposicao_(data);

  const horarioProcurado =
    chaveHorarioReposicao_(horario);

  const profissionalProcurado =
    normalizarTextoReposicao_(
      fisioterapeuta
    );

  return bloqueios.some(function(bloqueio) {
    if (bloqueio.data !== dataProcurada) {
      return false;
    }

    const bloqueiaHorario =
      !bloqueio.horario ||
      bloqueio.horario ===
        horarioProcurado ||
      bloqueio.abrangencia ===
        'dia inteiro' ||
      bloqueio.abrangencia ===
        'turno inteiro';

    const bloqueiaProfissional =
      !bloqueio.fisioterapeuta ||
      bloqueio.fisioterapeuta ===
        profissionalProcurado ||
      bloqueio.fisioterapeuta ===
        'todos';

    return (
      bloqueiaHorario &&
      bloqueiaProfissional
    );
  });
}


/**
 * Confere capacidade e compatibilidade do grupo.
 */
function verificarDisponibilidadeReposicao_(
  todosAgendamentos,
  data,
  original
) {
  const chaveData =
    chaveDataReposicao_(data);

  const chaveHorario =
    chaveHorarioReposicao_(
      original.horario
    );

  const fisioterapeuta =
    normalizarTextoReposicao_(
      original.fisioterapeuta
    );

  const tipoGrupo =
    normalizarTextoReposicao_(
      original.tipoGrupo
    );

  const capacidade =
    obterCapacidadeTipoReposicao_(
      original.tipoGrupo
    );

  let ocupacao = 0;
  const tiposExistentes = {};

  todosAgendamentos.forEach(function(linha) {
    const dataAtual =
      linha[
        CONFIG_REPOSICAO_CANCELAMENTO
          .AGENDAMENTOS
          .DATA - 1
      ];

    if (!(dataAtual instanceof Date)) {
      return;
    }

    const evento =
      normalizarTextoReposicao_(
        linha[
          CONFIG_REPOSICAO_CANCELAMENTO
            .AGENDAMENTOS
            .EVENTO - 1
        ]
      );

    const status =
      normalizarTextoReposicao_(
        linha[
          CONFIG_REPOSICAO_CANCELAMENTO
            .AGENDAMENTOS
            .STATUS - 1
        ]
      );

    if (
      chaveDataReposicao_(dataAtual) !==
        chaveData ||
      chaveHorarioReposicao_(
        linha[
          CONFIG_REPOSICAO_CANCELAMENTO
            .AGENDAMENTOS
            .HORARIO - 1
        ]
      ) !== chaveHorario ||
      normalizarTextoReposicao_(
        linha[
          CONFIG_REPOSICAO_CANCELAMENTO
            .AGENDAMENTOS
            .FISIOTERAPEUTA - 1
        ]
      ) !== fisioterapeuta ||
      evento !== 'sessao' ||
      CONFIG_REPOSICAO_CANCELAMENTO
        .STATUS_QUE_OCUPAM_VAGA
        .indexOf(status) === -1
    ) {
      return;
    }

    ocupacao++;

    const tipoExistente =
      normalizarTextoReposicao_(
        linha[
          CONFIG_REPOSICAO_CANCELAMENTO
            .AGENDAMENTOS
            .TIPO_GRUPO - 1
        ]
      );

    if (tipoExistente) {
      tiposExistentes[tipoExistente] = true;
    }
  });

  const tipos = Object.keys(
    tiposExistentes
  );

  /*
   * Se o horário já estiver ocupado, deverá existir
   * somente um tipo de grupo e ele deverá ser igual
   * ao tipo da reposição.
   */
  if (
    tipos.length > 0 &&
    (
      tipos.length !== 1 ||
      tipos[0] !== tipoGrupo
    )
  ) {
    return false;
  }

  return ocupacao < capacidade;
}


/**
 * Retorna a capacidade oficial de cada tipo de
 * atendimento e rejeita tipos antigos ou desconhecidos.
 */
function obterCapacidadeTipoReposicao_(
  tipoAtendimento
) {
  const tipo =
    normalizarTextoReposicao_(
      tipoAtendimento
    );

  if (
    tipo ===
      'atendimento com maior supervisao'
  ) {
    return 2;
  }

  const gruposComCapacidadeSeis = [
    'grupo de mmss',
    'grupo de mmii',
    'grupo de coluna'
  ];

  if (
    gruposComCapacidadeSeis
      .indexOf(tipo) !== -1
  ) {
    return 6;
  }

  throw new Error(
    'O tipo de atendimento "' +
      String(tipoAtendimento || '') +
      '" não é reconhecido. Atualize o cadastro do paciente antes de gerar a reposição.'
  );
}


/**
 * Monta a nova linha da reposição.
 */function montarLinhaReposicao_(
  original,
  novoId,
  novaData,
  agora
) {
  return [
    novoId,
    original.idPaciente,
    original.prontuario,
    original.nomePaciente,
    original.idCiclo,
    original.cicloNumero,
    new Date(novaData),
    obterNomeDiaReposicao_(novaData),
    original.horario,
    original.fisioterapeuta,
    original.tipoGrupo,
    'Sessão',
    original.numeroSessao,
    original.totalPrescrito,

    /*
     * Grava a capacidade oficial do tipo atual,
     * corrigindo possíveis valores antigos.
     */
    obterCapacidadeTipoReposicao_(
      original.tipoGrupo
    ),

    'Agendado',
    'Reposição automática de ' +
      original.idAgendamento,
    'Não',
    'Sim',
    agora,
    agora,
    'Não'
  ];
}


/**
 * Gera o próximo ID de agendamento.
 */
function gerarProximoIdAgendamentoReposicao_(
  abaAgendamentos
) {
  const ultimaLinha =
    abaAgendamentos.getLastRow();

  if (ultimaLinha < 2) {
    return 'AG-000001';
  }

  const ids = abaAgendamentos
    .getRange(
      2,
      1,
      ultimaLinha - 1,
      1
    )
    .getDisplayValues();

  let maiorNumero = 0;

  ids.forEach(function(linha) {
    const texto = String(
      linha[0] || ''
    ).trim();

    const resultado =
      texto.match(/(\d+)$/);

    if (!resultado) {
      return;
    }

    const numero =
      Number(resultado[1]);

    if (
      Number.isFinite(numero) &&
      numero > maiorNumero
    ) {
      maiorNumero = numero;
    }
  });

  return (
    'AG-' +
    String(maiorNumero + 1)
      .padStart(6, '0')
  );
}


/**
 * Registra no agendamento original qual reposição foi criada.
 */
function atualizarMotivoSessaoOriginal_(
  abaAgendamentos,
  linhaOriginal,
  idReposicao
) {
  const colunaMotivo =
    CONFIG_REPOSICAO_CANCELAMENTO
      .AGENDAMENTOS
      .MOTIVO;

  const motivoAtual = String(
    abaAgendamentos.getRange(
      linhaOriginal,
      colunaMotivo
    ).getValue() || ''
  ).trim();

  const identificacao =
    'Reposição gerada: ' +
    idReposicao;

  if (
    normalizarTextoReposicao_(motivoAtual)
      .indexOf(
        normalizarTextoReposicao_(
          identificacao
        )
      ) !== -1
  ) {
    return;
  }

  const novoMotivo = motivoAtual
    ? motivoAtual + ' | ' + identificacao
    : 'Cancelado pela clínica | ' +
      identificacao;

  abaAgendamentos.getRange(
    linhaOriginal,
    colunaMotivo
  ).setValue(novoMotivo);
}


/**
 * Atualiza a data prevista de término no cadastro.
 */
function atualizarDataTerminoPacienteReposicao_(
  abaCadastro,
  idPaciente,
  novaData
) {
  const ultimaLinha =
    abaCadastro.getLastRow();

  if (ultimaLinha < 2) {
    return;
  }

  const ids = abaCadastro
    .getRange(
      2,
      CONFIG_REPOSICAO_CANCELAMENTO
        .CADASTRO
        .ID_PACIENTE,
      ultimaLinha - 1,
      1
    )
    .getDisplayValues();

  const idProcurado =
    normalizarTextoReposicao_(
      idPaciente
    );

  for (
    let indice = 0;
    indice < ids.length;
    indice++
  ) {
    const idAtual =
      normalizarTextoReposicao_(
        ids[indice][0]
      );

    if (idAtual !== idProcurado) {
      continue;
    }

    const linhaCadastro =
      indice + 2;

    const celulaDataTermino =
      abaCadastro.getRange(
        linhaCadastro,
        CONFIG_REPOSICAO_CANCELAMENTO
          .CADASTRO
          .DATA_PREVISTA_TERMINO
      );

    const dataAtual =
      celulaDataTermino.getValue();

    if (
      !(dataAtual instanceof Date) ||
      novaData.getTime() >
        dataAtual.getTime()
    ) {
      celulaDataTermino
        .setValue(new Date(novaData))
        .setNumberFormat('dd/MM/yyyy');
    }

    return;
  }
}


/**
 * Retorna o nome do dia da semana.
 */
function obterNomeDiaReposicao_(data) {
  const nomes = [
    'Domingo',
    'Segunda-feira',
    'Terça-feira',
    'Quarta-feira',
    'Quinta-feira',
    'Sexta-feira',
    'Sábado'
  ];

  return nomes[
    new Date(data).getDay()
  ];
}


/**
 * Padroniza a data.
 */
function chaveDataReposicao_(data) {
  return Utilities.formatDate(
    new Date(data),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );
}


/**
 * Padroniza o horário.
 */
function chaveHorarioReposicao_(valor) {
  if (valor instanceof Date) {
    return Utilities.formatDate(
      valor,
      Session.getScriptTimeZone(),
      'HH:mm'
    );
  }

  const texto = String(
    valor || ''
  ).trim();

  const resultado = texto.match(
    /(\d{1,2}):(\d{2})/
  );

  if (!resultado) {
    return texto;
  }

  return (
    String(
      Number(resultado[1])
    ).padStart(2, '0') +
    ':' +
    resultado[2]
  );
}


/**
 * Padroniza textos para comparação.
 */
function normalizarTextoReposicao_(valor) {
  return String(valor || '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(
      /[\u0300-\u036f]/g,
      ''
    );
}
