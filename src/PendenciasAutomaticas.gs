const CONFIG_PENDENCIAS_AUTOMATICAS = {
  ABAS: {
    CADASTRO: 'Cadastro de Pacientes',
    AGENDAMENTOS: 'Agendamentos',
    PENDENCIAS: 'Pendências',
    HISTORICO: 'Histórico de Pendências',
    AGENDA: 'Agenda',
    FERIADOS: 'Calendário da Prefeitura',
    BLOQUEIOS: 'Bloqueios'
  },

  CADASTRO: {
    ID_PACIENTE: 1,
    PRONTUARIO: 2,
    NOME: 3,
    TELEFONE: 5,
    HORARIO: 6,
    TIPO_ATENDIMENTO: 7,
    SEGUNDA: 9,
    TERCA: 10,
    QUARTA: 11,
    QUINTA: 12,
    SEXTA: 13,
    SESSOES_PRESCRITAS: 14,
    SESSOES_REALIZADAS: 15,
    SESSOES_RESTANTES: 16,
    DATA_AVALIACAO: 17,
    DATA_INICIO: 19,
    DATA_TERMINO: 20,
    STATUS: 21,
    FISIOTERAPEUTA: 22,
    DESFECHO: 24
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

  PENDENCIAS: {
    ID: 1,
    PRIORIDADE: 2,
    PENDENCIA: 3,
    PACIENTE: 4,
    PRONTUARIO: 5,
    TELEFONE: 6,
    RESPONSAVEL: 7,
    DATA_REFERENCIA: 8,
    PRAZO: 9,
    SITUACAO: 10,
    DATA_CONCLUSAO: 11
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

  QUANTIDADE_COLUNAS_CADASTRO: 24,
  QUANTIDADE_COLUNAS_AGENDAMENTOS: 22,
  QUANTIDADE_COLUNAS_PENDENCIAS: 11,
  DIAS_ANTECEDENCIA_DOCUMENTACAO: 14,
  DIAS_ANTECEDENCIA_FERIADO: 7
};


function atualizarPendenciasAutomaticas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getDocumentLock();
  let bloqueioObtido = false;

  try {
    lock.waitLock(30000);
    bloqueioObtido = true;

    const abas = obterAbasPendencias_(ss);

    moverPendenciasConcluidasParaHistorico_(
      abas.pendencias,
      abas.historico
    );

    const idsConcluidos =
      obterTodosIdsPendencias_(abas.historico);

    const pendenciasExistentes =
      lerPendenciasAtivas_(
        abas.pendencias,
        idsConcluidos
      );

    const contexto = {
      hoje: removerHorarioPendencias_(new Date()),
      pacientes:
        lerPacientesPendencias_(abas.cadastro),
      agendamentos:
        lerAgendamentosPendencias_(
          abas.agendamentos
        ),
      feriados:
        lerFeriadosPendencias_(abas.feriados),
      bloqueios:
        lerBloqueiosPendencias_(abas.bloqueios),
      idsConcluidos: idsConcluidos,
      pendencias: []
    };

    contexto.pacientesPorId = {};

    contexto.pacientes.forEach(function (paciente) {
      contexto.pacientesPorId[paciente.id] =
        paciente;
    });

    contexto.agendamentosPorPaciente = {};

    contexto.agendamentos.forEach(
      function (agendamento) {
        const idPaciente =
          agendamento.idPaciente;

        if (
          !contexto
            .agendamentosPorPaciente[idPaciente]
        ) {
          contexto
            .agendamentosPorPaciente[idPaciente] =
              [];
        }

        contexto
          .agendamentosPorPaciente[idPaciente]
          .push(agendamento);
      }
    );

    gerarPendenciaMensalRegulacao_(contexto);

    contexto.pacientes.forEach(
      function (paciente) {
        gerarPendenciasPaciente_(
          paciente,
          contexto
        );
      }
    );

    const pendenciasExistentesValidas =
      removerDesfechosGenericosPorAbandono_(
        pendenciasExistentes,
        contexto
      );

    const pendenciasFinais =
      combinarPendencias_(
        pendenciasExistentesValidas,
        contexto.pendencias,
        idsConcluidos
      );

    ordenarPendencias_(pendenciasFinais);

    gravarPendenciasAtivas_(
      abas.pendencias,
      pendenciasFinais
    );

    SpreadsheetApp.flush();

    return pendenciasFinais.length;
  } finally {
    if (bloqueioObtido) {
      try {
        lock.releaseLock();
      } catch (erroLock) {
        // Liberação automática.
      }
    }
  }
}


function instalarGatilhosPendenciasAutomaticas() {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const ui = SpreadsheetApp.getUi();

  try {
    removerGatilhosPendenciasDuplicados_();

    ScriptApp
      .newTrigger(
        'processarEdicaoPendenciasAutomaticas'
      )
      .forSpreadsheet(ss)
      .onEdit()
      .create();

    ScriptApp
      .newTrigger(
        'atualizarPendenciasAoAbrir'
      )
      .forSpreadsheet(ss)
      .onOpen()
      .create();

    ScriptApp
      .newTrigger(
        'atualizarPendenciasDiariamente'
      )
      .timeBased()
      .everyDays(1)
      .atHour(6)
      .create();

    const quantidade =
      atualizarPendenciasAutomaticas();

    ui.alert(
      'Pendências automáticas ativadas',
      'Os gatilhos foram instalados com sucesso.\n\n' +
        'Pendências ativas encontradas: ' +
        quantidade,
      ui.ButtonSet.OK
    );
  } catch (erro) {
    ui.alert(
      'Erro ao ativar as pendências',
      erro && erro.message
        ? erro.message
        : String(erro),
      ui.ButtonSet.OK
    );

    throw erro;
  }
}


function atualizarPendenciasAoAbrir(e) {
  try {
    atualizarPendenciasAutomaticas();
  } catch (erro) {
    console.error(erro);
  }
}


function atualizarPendenciasDiariamente() {
  try {
    atualizarPendenciasAutomaticas();
  } catch (erro) {
    console.error(erro);
  }
}


function processarEdicaoPendenciasAutomaticas(e) {
  if (!e || !e.range) {
    return;
  }

  const aba = e.range.getSheet();
  const nomeAba = aba.getName();
  const linha = e.range.getRow();
  const coluna = e.range.getColumn();

  const nomes =
    CONFIG_PENDENCIAS_AUTOMATICAS.ABAS;

  const colunasPendencias =
    CONFIG_PENDENCIAS_AUTOMATICAS.PENDENCIAS;

  if (
    nomeAba === nomes.PENDENCIAS &&
    linha >= 2 &&
    (
      coluna ===
        colunasPendencias.SITUACAO ||
      coluna ===
        colunasPendencias.DATA_CONCLUSAO
    )
  ) {
    processarConclusaoPendenciaEditada_(
      e,
      aba,
      linha
    );

    return;
  }

  const alteracaoRelevante =
    nomeAba === nomes.CADASTRO ||
    nomeAba === nomes.AGENDAMENTOS ||
    nomeAba === nomes.BLOQUEIOS ||
    nomeAba === nomes.FERIADOS ||
    (
      nomeAba === nomes.AGENDA &&
      coluna === 4 &&
      linha >= 5
    );

  if (!alteracaoRelevante) {
    return;
  }

  if (nomeAba === nomes.AGENDA) {
    Utilities.sleep(1500);
  }

  try {
    atualizarPendenciasAutomaticas();
  } catch (erro) {
    console.error(erro);
  }
}


function processarConclusaoPendenciaEditada_(
  e,
  aba,
  linha
) {
  const colunas =
    CONFIG_PENDENCIAS_AUTOMATICAS.PENDENCIAS;

  const situacao =
    normalizarTextoPendencias_(
      aba
        .getRange(
          linha,
          colunas.SITUACAO
        )
        .getValue()
    );

  const dataConclusao =
    aba
      .getRange(
        linha,
        colunas.DATA_CONCLUSAO
      )
      .getValue();

  const concluida =
    situacao === 'concluida' ||
    situacao === 'concluido';

  if (
    concluida &&
    !dataValidaPendencias_(dataConclusao)
  ) {
    e.source.toast(
      'Preencha também a Data da conclusão. A pendência continuará aberta até os dois campos estarem preenchidos.',
      'Data da conclusão obrigatória',
      7
    );

    return;
  }

  if (
    !concluida &&
    dataValidaPendencias_(dataConclusao)
  ) {
    e.source.toast(
      'Para finalizar, altere também a Situação para Concluída.',
      'Situação ainda pendente',
      6
    );

    return;
  }

  if (!concluida) {
    return;
  }

  try {
    atualizarPendenciasAutomaticas();

    e.source.toast(
      'A pendência concluída foi enviada para o histórico.',
      'Pendência concluída',
      5
    );
  } catch (erro) {
    console.error(erro);

    e.source.toast(
      erro && erro.message
        ? erro.message
        : String(erro),
      'Erro ao concluir pendência',
      8
    );
  }
}


function removerGatilhosPendenciasDuplicados_() {
  const funcoes = [
    'processarEdicaoPendenciasAutomaticas',
    'atualizarPendenciasAoAbrir',
    'atualizarPendenciasDiariamente'
  ];

  ScriptApp
    .getProjectTriggers()
    .forEach(function (gatilho) {
      if (
        funcoes.indexOf(
          gatilho.getHandlerFunction()
        ) !== -1
      ) {
        ScriptApp.deleteTrigger(gatilho);
      }
    });
}


function obterAbasPendencias_(ss) {
  const nomes =
    CONFIG_PENDENCIAS_AUTOMATICAS.ABAS;

  const abas = {
    cadastro:
      ss.getSheetByName(nomes.CADASTRO),

    agendamentos:
      ss.getSheetByName(nomes.AGENDAMENTOS),

    pendencias:
      ss.getSheetByName(nomes.PENDENCIAS),

    historico:
      ss.getSheetByName(nomes.HISTORICO),

    feriados:
      ss.getSheetByName(nomes.FERIADOS),

    bloqueios:
      ss.getSheetByName(nomes.BLOQUEIOS)
  };

  Object.keys(abas).forEach(
    function (chave) {
      if (!abas[chave]) {
        const nomesPorChave = {
          cadastro: nomes.CADASTRO,
          agendamentos: nomes.AGENDAMENTOS,
          pendencias: nomes.PENDENCIAS,
          historico: nomes.HISTORICO,
          feriados: nomes.FERIADOS,
          bloqueios: nomes.BLOQUEIOS
        };

        throw new Error(
          'A aba necessária "' +
            nomesPorChave[chave] +
            '" não foi encontrada.'
        );
      }
    }
  );

  return abas;
}
function moverPendenciasConcluidasParaHistorico_(
  abaPendencias,
  abaHistorico
) {
  const ultimaLinha =
    abaPendencias.getLastRow();

  if (ultimaLinha < 2) {
    return;
  }

  const dados =
    abaPendencias
      .getRange(
        2,
        1,
        ultimaLinha - 1,
        CONFIG_PENDENCIAS_AUTOMATICAS
          .QUANTIDADE_COLUNAS_PENDENCIAS
      )
      .getValues();

  const idsHistorico =
    obterTodosIdsPendencias_(abaHistorico);

  const linhasHistorico = [];

  dados.forEach(function (linha) {
    const id = String(
      linha[
        CONFIG_PENDENCIAS_AUTOMATICAS
          .PENDENCIAS.ID - 1
      ] || ''
    ).trim();

    const situacao =
      normalizarTextoPendencias_(
        linha[
          CONFIG_PENDENCIAS_AUTOMATICAS
            .PENDENCIAS.SITUACAO - 1
        ]
      );

    const dataConclusao =
      linha[
        CONFIG_PENDENCIAS_AUTOMATICAS
          .PENDENCIAS.DATA_CONCLUSAO - 1
      ];

    const concluida =
      situacao === 'concluida' ||
      situacao === 'concluido';

    if (
      !id ||
      !concluida ||
      !dataValidaPendencias_(
        dataConclusao
      ) ||
      idsHistorico[id]
    ) {
      return;
    }

    const novaLinha = linha.slice();

    novaLinha[
      CONFIG_PENDENCIAS_AUTOMATICAS
        .PENDENCIAS.SITUACAO - 1
    ] = 'Concluída';

    novaLinha[
      CONFIG_PENDENCIAS_AUTOMATICAS
        .PENDENCIAS.DATA_CONCLUSAO - 1
    ] = removerHorarioPendencias_(
      dataConclusao
    );

    linhasHistorico.push(novaLinha);
    idsHistorico[id] = true;
  });

  if (linhasHistorico.length === 0) {
    return;
  }

  const primeiraLinha = Math.max(
    abaHistorico.getLastRow() + 1,
    2
  );

  abaHistorico
    .getRange(
      primeiraLinha,
      1,
      linhasHistorico.length,
      CONFIG_PENDENCIAS_AUTOMATICAS
        .QUANTIDADE_COLUNAS_PENDENCIAS
    )
    .setValues(linhasHistorico);

  abaHistorico
    .getRange(
      primeiraLinha,
      CONFIG_PENDENCIAS_AUTOMATICAS
        .PENDENCIAS.DATA_REFERENCIA,
      linhasHistorico.length,
      2
    )
    .setNumberFormat('dd/MM/yyyy');

  abaHistorico
    .getRange(
      primeiraLinha,
      CONFIG_PENDENCIAS_AUTOMATICAS
        .PENDENCIAS.DATA_CONCLUSAO,
      linhasHistorico.length,
      1
    )
    .setNumberFormat('dd/MM/yyyy');
}


function obterTodosIdsPendencias_(aba) {
  const resultado = {};
  const ultimaLinha = aba.getLastRow();

  if (ultimaLinha < 2) {
    return resultado;
  }

  const ids =
    aba
      .getRange(
        2,
        CONFIG_PENDENCIAS_AUTOMATICAS
          .PENDENCIAS.ID,
        ultimaLinha - 1,
        1
      )
      .getDisplayValues();

  ids.forEach(function (linha) {
    const id = String(
      linha[0] || ''
    ).trim();

    if (id) {
      resultado[id] = true;
    }
  });

  return resultado;
}


function lerPendenciasAtivas_(
  aba,
  idsConcluidos
) {
  const ultimaLinha = aba.getLastRow();

  if (ultimaLinha < 2) {
    return [];
  }

  const dados =
    aba
      .getRange(
        2,
        1,
        ultimaLinha - 1,
        CONFIG_PENDENCIAS_AUTOMATICAS
          .QUANTIDADE_COLUNAS_PENDENCIAS
      )
      .getValues();

  return dados.filter(function (linha) {
    const id = String(
      linha[0] || ''
    ).trim();

    if (!id || idsConcluidos[id]) {
      return false;
    }

    const situacao =
      normalizarTextoPendencias_(
        linha[
          CONFIG_PENDENCIAS_AUTOMATICAS
            .PENDENCIAS.SITUACAO - 1
        ]
      );

    const dataConclusao =
      linha[
        CONFIG_PENDENCIAS_AUTOMATICAS
          .PENDENCIAS.DATA_CONCLUSAO - 1
      ];

    const concluida =
      situacao === 'concluida' ||
      situacao === 'concluido';

    return !(
      concluida &&
      dataValidaPendencias_(dataConclusao)
    );
  });
}


function lerPacientesPendencias_(
  abaCadastro
) {
  const ultimaLinha =
    abaCadastro.getLastRow();

  if (ultimaLinha < 2) {
    return [];
  }

  const dados =
    abaCadastro
      .getRange(
        2,
        1,
        ultimaLinha - 1,
        CONFIG_PENDENCIAS_AUTOMATICAS
          .QUANTIDADE_COLUNAS_CADASTRO
      )
      .getValues();

  return dados
    .map(function (linha) {
      return montarPacientePendencias_(
        linha
      );
    })
    .filter(function (paciente) {
      return Boolean(paciente.id);
    });
}


function montarPacientePendencias_(
  linha
) {
  const c =
    CONFIG_PENDENCIAS_AUTOMATICAS
      .CADASTRO;

  return {
    id: String(
      linha[c.ID_PACIENTE - 1] || ''
    ).trim(),

    prontuario: String(
      linha[c.PRONTUARIO - 1] || ''
    ).trim(),

    nome: String(
      linha[c.NOME - 1] || ''
    ).trim(),

    telefone: String(
      linha[c.TELEFONE - 1] || ''
    ).trim(),

    horario:
      linha[c.HORARIO - 1],

    tipoAtendimento: String(
      linha[
        c.TIPO_ATENDIMENTO - 1
      ] || ''
    ).trim(),

    diasSemana: {
      1: valorAtivoPendencias_(
        linha[c.SEGUNDA - 1]
      ),

      2: valorAtivoPendencias_(
        linha[c.TERCA - 1]
      ),

      3: valorAtivoPendencias_(
        linha[c.QUARTA - 1]
      ),

      4: valorAtivoPendencias_(
        linha[c.QUINTA - 1]
      ),

      5: valorAtivoPendencias_(
        linha[c.SEXTA - 1]
      )
    },

    sessoesPrescritas:
      Number(
        linha[
          c.SESSOES_PRESCRITAS - 1
        ]
      ) || 0,

    sessoesRealizadas:
      Number(
        linha[
          c.SESSOES_REALIZADAS - 1
        ]
      ) || 0,

    sessoesRestantes:
      Number(
        linha[
          c.SESSOES_RESTANTES - 1
        ]
      ) || 0,

    dataAvaliacao:
      linha[c.DATA_AVALIACAO - 1],

    dataInicio:
      linha[c.DATA_INICIO - 1],

    dataTermino:
      linha[c.DATA_TERMINO - 1],

    status: String(
      linha[c.STATUS - 1] || ''
    ).trim(),

    fisioterapeuta: String(
      linha[
        c.FISIOTERAPEUTA - 1
      ] || ''
    ).trim(),

    desfecho: String(
      linha[c.DESFECHO - 1] || ''
    ).trim()
  };
}


function lerAgendamentosPendencias_(
  abaAgendamentos
) {
  const ultimaLinha =
    abaAgendamentos.getLastRow();

  if (ultimaLinha < 2) {
    return [];
  }

  const dados =
    abaAgendamentos
      .getRange(
        2,
        1,
        ultimaLinha - 1,
        CONFIG_PENDENCIAS_AUTOMATICAS
          .QUANTIDADE_COLUNAS_AGENDAMENTOS
      )
      .getValues();

  return dados
    .map(function (linha, indice) {
      const c =
        CONFIG_PENDENCIAS_AUTOMATICAS
          .AGENDAMENTOS;

      return {
        linha: indice + 2,

        idAgendamento: String(
          linha[
            c.ID_AGENDAMENTO - 1
          ] || ''
        ).trim(),

        idPaciente: String(
          linha[
            c.ID_PACIENTE - 1
          ] || ''
        ).trim(),

        prontuario: String(
          linha[c.PRONTUARIO - 1] || ''
        ).trim(),

        nomePaciente: String(
          linha[
            c.NOME_PACIENTE - 1
          ] || ''
        ).trim(),

        idCiclo: String(
          linha[c.ID_CICLO - 1] || ''
        ).trim(),

        cicloNumero:
          Number(
            linha[
              c.CICLO_NUMERO - 1
            ]
          ) || 0,

        data:
          linha[c.DATA - 1],

        horario:
          linha[c.HORARIO - 1],

        fisioterapeuta: String(
          linha[
            c.FISIOTERAPEUTA - 1
          ] || ''
        ).trim(),

        tipoGrupo: String(
          linha[
            c.TIPO_GRUPO - 1
          ] || ''
        ).trim(),

        evento: String(
          linha[c.EVENTO - 1] || ''
        ).trim(),

        numeroSessao:
          Number(
            linha[
              c.NUMERO_SESSAO - 1
            ]
          ) || 0,

        totalPrescrito:
          Number(
            linha[
              c.TOTAL_PRESCRITO - 1
            ]
          ) || 0,

        limiteGrupo:
          Number(
            linha[
              c.LIMITE_GRUPO - 1
            ]
          ) || 0,

        status: String(
          linha[c.STATUS - 1] || ''
        ).trim(),

        motivo: String(
          linha[c.MOTIVO - 1] || ''
        ).trim(),

        avisarPaciente: String(
          linha[
            c.AVISAR_PACIENTE - 1
          ] || ''
        ).trim()
      };
    })
    .filter(function (agendamento) {
      return Boolean(
        agendamento.idAgendamento
      );
    });
}
function lerFeriadosPendencias_(
  abaFeriados
) {
  const ultimaLinha =
    abaFeriados.getLastRow();

  if (ultimaLinha < 2) {
    return [];
  }

  const dados =
    abaFeriados
      .getRange(
        2,
        1,
        ultimaLinha - 1,
        Math.max(
          abaFeriados.getLastColumn(),
          4
        )
      )
      .getValues();

  return dados
    .map(function (linha) {
      const c =
        CONFIG_PENDENCIAS_AUTOMATICAS
          .FERIADOS;

      return {
        data:
          linha[c.DATA - 1],

        descricao: String(
          linha[
            c.DESCRICAO - 1
          ] || 'Feriado'
        ).trim(),

        atendimento:
          normalizarTextoPendencias_(
            linha[
              c.ATENDIMENTO - 1
            ]
          )
      };
    })
    .filter(function (feriado) {
      return (
        dataValidaPendencias_(
          feriado.data
        ) &&
        feriado.atendimento === 'nao'
      );
    });
}


function lerBloqueiosPendencias_(
  abaBloqueios
) {
  const ultimaLinha =
    abaBloqueios.getLastRow();

  if (ultimaLinha < 2) {
    return [];
  }

  const dados =
    abaBloqueios
      .getRange(
        2,
        1,
        ultimaLinha - 1,
        Math.max(
          abaBloqueios.getLastColumn(),
          7
        )
      )
      .getValues();

  return dados
    .map(function (linha, indice) {
      const c =
        CONFIG_PENDENCIAS_AUTOMATICAS
          .BLOQUEIOS;

      return {
        linha: indice + 2,

        data:
          linha[c.DATA - 1],

        horario:
          linha[c.HORARIO - 1],

        fisioterapeuta: String(
          linha[
            c.FISIOTERAPEUTA - 1
          ] || ''
        ).trim(),

        abrangencia: String(
          linha[
            c.ABRANGENCIA - 1
          ] || ''
        ).trim(),

        motivo: String(
          linha[c.MOTIVO - 1] || ''
        ).trim(),

        acao: String(
          linha[c.ACAO - 1] || ''
        ).trim(),

        status: String(
          linha[c.STATUS - 1] || ''
        ).trim()
      };
    })
    .filter(function (bloqueio) {
      const status =
        normalizarTextoPendencias_(
          bloqueio.status
        );

      return (
        dataValidaPendencias_(
          bloqueio.data
        ) &&
        (
          !status ||
          status === 'ativo' ||
          status === 'bloqueado'
        )
      );
    });
}


function gerarPendenciasPaciente_(
  paciente,
  contexto
) {
  const agendamentos =
    contexto
      .agendamentosPorPaciente[
        paciente.id
      ] || [];

  gerarPendenciasAtendimentosSemResultado_(
    paciente,
    agendamentos,
    contexto
  );

  gerarPendenciasAvaliacao_(
    paciente,
    agendamentos,
    contexto
  );

  gerarPendenciasReagendamentoAvaliacao_(
    paciente,
    agendamentos,
    contexto
  );

  gerarPendenciasPlanejamento_(
    paciente,
    agendamentos,
    contexto
  );

  gerarPendenciaDocumentacaoPreventiva_(
    paciente,
    agendamentos,
    contexto
  );

  gerarPendenciasFaltas_(
    paciente,
    agendamentos,
    contexto
  );

  gerarPendenciasCancelamentoClinica_(
    paciente,
    agendamentos,
    contexto
  );

  gerarPendenciasFeriados_(
    paciente,
    agendamentos,
    contexto
  );

  gerarPendenciaSessoesSemAgenda_(
    paciente,
    agendamentos,
    contexto
  );

  gerarPendenciasDesfechoRenovacao_(
    paciente,
    agendamentos,
    contexto
  );

  gerarPendenciaFaltaUltimaSessao_(
    paciente,
    agendamentos,
    contexto
  );
}


function gerarPendenciasAtendimentosSemResultado_(
  paciente,
  agendamentos,
  contexto
) {
  agendamentos.forEach(
    function (agendamento) {
      if (
        !dataValidaPendencias_(
          agendamento.data
        ) ||
        normalizarTextoPendencias_(
          agendamento.status
        ) !== 'agendado' ||
        removerHorarioPendencias_(
          agendamento.data
        ).getTime() >=
          contexto.hoje.getTime()
      ) {
        return;
      }

      adicionarPendencia_(
        contexto,

        criarIdPendencia_(
          paciente.id,
          'RESULTADO-' +
            agendamento.idAgendamento
        ),

        'Urgente',
        'Registrar resultado do atendimento',
        paciente,
        'Recepção',
        agendamento.data,
        contexto.hoje
      );
    }
  );
}


function gerarPendenciasAvaliacao_(
  paciente,
  agendamentos,
  contexto
) {
  const avaliacoes =
    agendamentos.filter(
      function (agendamento) {
        return (
          normalizarTextoPendencias_(
            agendamento.evento
          ) === 'avaliacao'
        );
      }
    );

  avaliacoes.forEach(
    function (avaliacao) {
      const status =
        normalizarTextoPendencias_(
          avaliacao.status
        );

      if (
        status ===
          'falta justificada' ||
        status ===
          'falta nao justificada'
      ) {
        adicionarPendencia_(
          contexto,

          criarIdPendencia_(
            paciente.id,
            'FALTA-AVALIACAO-' +
              avaliacao.idAgendamento
          ),

          'Alta',
          'Informar falta na avaliação à Regulação',
          paciente,
          'Recepção',
          avaliacao.data,

          proximoDiaUtilPendencias_(
            avaliacao.data
          )
        );
      }
    }
  );

  const possuiAvaliacaoFutura =
    avaliacoes.some(
      function (avaliacao) {
        return (
          normalizarTextoPendencias_(
            avaliacao.status
          ) === 'agendado' &&

          dataValidaPendencias_(
            avaliacao.data
          ) &&

          removerHorarioPendencias_(
            avaliacao.data
          ).getTime() >=
            contexto.hoje.getTime()
        );
      }
    );

  const possuiAvaliacaoComFalta =
    avaliacoes.some(
      function (avaliacao) {
        const status =
          normalizarTextoPendencias_(
            avaliacao.status
          );

        return (
          status ===
            'falta justificada' ||
          status ===
            'falta nao justificada'
        );
      }
    );

  const possuiAvaliacaoCanceladaClinica =
    avaliacoes.some(
      function (avaliacao) {
        return (
          normalizarTextoPendencias_(
            avaliacao.status
          ) ===
            'cancelado pela clinica'
        );
      }
    );

  const statusPaciente =
    normalizarTextoPendencias_(
      paciente.status
    );

  if (
    (
      statusPaciente ===
        'avaliacao agendada' ||
      statusPaciente ===
        'aguardando avaliacao'
    ) &&
    !possuiAvaliacaoFutura &&
    !possuiAvaliacaoComFalta &&
    !possuiAvaliacaoCanceladaClinica
  ) {
    adicionarPendencia_(
      contexto,

      criarIdPendencia_(
        paciente.id,
        'AVALIACAO'
      ),

      'Alta',
      'Agendar avaliação',
      paciente,
      'Recepção',
      paciente.dataAvaliacao,
      contexto.hoje
    );
  }
}


function gerarPendenciasPlanejamento_(
  paciente,
  agendamentos,
  contexto
) {
  const statusPaciente =
    normalizarTextoPendencias_(
      paciente.status
    );

  if (
    statusPaciente !==
      'avaliado - aguardando agendamento'
  ) {
    return;
  }

  const ultimaAvaliacaoComparecida =
    obterUltimoAgendamento_(
      agendamentos.filter(
        function (agendamento) {
          return (
            normalizarTextoPendencias_(
              agendamento.evento
            ) === 'avaliacao' &&

            normalizarTextoPendencias_(
              agendamento.status
            ) === 'compareceu'
          );
        }
      )
    );

  if (!ultimaAvaliacaoComparecida) {
    return;
  }

  const idBase =
    ultimaAvaliacaoComparecida
      .idAgendamento;

  if (
    !planoTratamentoCompletoPendencias_(
      paciente
    )
  ) {
    adicionarPendencia_(
      contexto,

      criarIdPendencia_(
        paciente.id,
        'PLANO-' + idBase
      ),

      'Alta',
      'Definir plano de tratamento',
      paciente,

      paciente.fisioterapeuta ||
        'Fisioterapeuta',

      ultimaAvaliacaoComparecida.data,

      proximoDiaUtilPendencias_(
        ultimaAvaliacaoComparecida.data
      )
    );

    return;
  }

  const possuiSessoes =
    agendamentos.some(
      function (agendamento) {
        return (
          normalizarTextoPendencias_(
            agendamento.evento
          ) === 'sessao'
        );
      }
    );

  if (!possuiSessoes) {
    adicionarPendencia_(
      contexto,

      criarIdPendencia_(
        paciente.id,
        'AGENDAMENTO-' + idBase
      ),

      'Alta',
      'Agendar ciclo de sessões',
      paciente,
      'Recepção',
      ultimaAvaliacaoComparecida.data,
      contexto.hoje
    );
  }
}
function gerarPendenciaDocumentacaoPreventiva_(
  paciente,
  agendamentos,
  contexto
) {
  if (
    normalizarTextoPendencias_(
      paciente.status
    ) !== 'em tratamento' ||

    !dataValidaPendencias_(
      paciente.dataTermino
    )
  ) {
    return;
  }

  const termino =
    removerHorarioPendencias_(
      paciente.dataTermino
    );

  const diasAteTermino =
    diferencaDiasPendencias_(
      contexto.hoje,
      termino
    );

  if (
    diasAteTermino >
      CONFIG_PENDENCIAS_AUTOMATICAS
        .DIAS_ANTECEDENCIA_DOCUMENTACAO
  ) {
    return;
  }

  const ciclo =
    obterCicloAtualPendencias_(
      agendamentos
    );

  let prioridade = 'Média';

  if (diasAteTermino <= 3) {
    prioridade = 'Urgente';
  } else if (diasAteTermino <= 7) {
    prioridade = 'Alta';
  }

  adicionarPendencia_(
    contexto,

    criarIdPendencia_(
      paciente.id,
      'DOCUMENTACAO-' + ciclo.id
    ),

    prioridade,

    'Preparar documentação de encerramento do tratamento',

    paciente,

    paciente.fisioterapeuta ||
      'Fisioterapeuta',

    termino,
    termino
  );
}


function gerarPendenciasFaltas_(
  paciente,
  agendamentos,
  contexto
) {
  const ciclo =
    obterCicloAtualPendencias_(
      agendamentos
    );

  if (!ciclo.id) {
    return;
  }

  const faltas =
    ciclo.sessoes
      .filter(function (sessao) {
        return (
          normalizarTextoPendencias_(
            sessao.status
          ) ===
            'falta nao justificada'
        );
      })
      .sort(
        compararAgendamentosPorData_
      );

  if (faltas.length === 2) {
    adicionarPendencia_(
      contexto,

      criarIdPendencia_(
        paciente.id,
        'DUAS-FALTAS-' + ciclo.id
      ),

      'Alta',

      'Contatar paciente — 2 faltas não justificadas',

      paciente,
      'Recepção',
      faltas[1].data,

      proximoDiaUtilPendencias_(
        faltas[1].data
      )
    );
  }

  if (faltas.length >= 3) {
    adicionarPendencia_(
      contexto,

      criarIdPendencia_(
        paciente.id,
        'ABANDONO-' + ciclo.id
      ),

      'Urgente',
      'Registrar alta por abandono',
      paciente,

      paciente.fisioterapeuta ||
        'Fisioterapeuta',

      faltas[2].data,

      proximoDiaUtilPendencias_(
        faltas[2].data
      )
    );
  }
}


function gerarPendenciasCancelamentoClinica_(
  paciente,
  agendamentos,
  contexto
) {
  agendamentos.forEach(
    function (agendamento) {
      if (
        normalizarTextoPendencias_(
          agendamento.evento
        ) !== 'sessao' ||

        normalizarTextoPendencias_(
          agendamento.status
        ) !==
          'cancelado pela clinica'
      ) {
        return;
      }

      let descricao =
        'Contatar paciente sobre cancelamento e reposição';

      if (agendamento.motivo) {
        descricao +=
          ' — ' +
          agendamento.motivo;
      }

      adicionarPendencia_(
        contexto,

        criarIdPendencia_(
          paciente.id,
          'CANCELAMENTO-' +
            agendamento.idAgendamento
        ),

        'Alta',
        descricao,
        paciente,
        'Recepção',
        agendamento.data,

        proximoDiaUtilPendencias_(
          contexto.hoje
        )
      );
    }
  );
}


function gerarPendenciasFeriados_(
  paciente,
  agendamentos,
  contexto
) {
  if (
    normalizarTextoPendencias_(
      paciente.status
    ) !== 'em tratamento' ||

    !dataValidaPendencias_(
      paciente.dataInicio
    ) ||

    !dataValidaPendencias_(
      paciente.dataTermino
    )
  ) {
    return;
  }

  const inicio =
    removerHorarioPendencias_(
      paciente.dataInicio
    );

  const termino =
    removerHorarioPendencias_(
      paciente.dataTermino
    );

  contexto.feriados.forEach(
    function (feriado) {
      const dataFeriado =
        removerHorarioPendencias_(
          feriado.data
        );

      const chaveDataFeriado =
        chaveDataPendencias_(
          dataFeriado
        );

      /*
       * Verifica se o feriado já foi convertido
       * em bloqueio de dia inteiro.
       */
      const possuiBloqueioDoFeriado =
        contexto.bloqueios.some(
          function (bloqueio) {
            if (
              !dataValidaPendencias_(
                bloqueio.data
              )
            ) {
              return false;
            }

            const mesmaData =
              chaveDataPendencias_(
                bloqueio.data
              ) === chaveDataFeriado;

            const diaInteiro =
              normalizarTextoPendencias_(
                bloqueio.abrangencia
              ) === 'dia inteiro';

            const motivoFeriado =
              normalizarTextoPendencias_(
                bloqueio.motivo
              ).indexOf(
                'feriado'
              ) === 0;

            const statusBloqueio =
              normalizarTextoPendencias_(
                bloqueio.status
              );

            const bloqueioAtivo =
              !statusBloqueio ||
              statusBloqueio === 'ativo' ||
              statusBloqueio === 'bloqueado';

            return (
              mesmaData &&
              diaInteiro &&
              motivoFeriado &&
              bloqueioAtivo
            );
          }
        );

      /*
       * Verifica se algum atendimento desse paciente
       * já foi cancelado e tratado pelo bloqueio.
       */
      const atendimentoJaTratado =
        agendamentos.some(
          function (agendamento) {
            if (
              !dataValidaPendencias_(
                agendamento.data
              )
            ) {
              return false;
            }

            const mesmaData =
              chaveDataPendencias_(
                agendamento.data
              ) === chaveDataFeriado;

            const statusCancelado =
              normalizarTextoPendencias_(
                agendamento.status
              ) ===
                'cancelado pela clinica';

            const evento =
              normalizarTextoPendencias_(
                agendamento.evento
              );

            return (
              mesmaData &&
              statusCancelado &&
              (
                evento === 'sessao' ||
                evento === 'avaliacao'
              )
            );
          }
        );

      /*
       * Quando o atendimento já foi cancelado ou
       * reagendado pelo bloqueio do feriado, a pendência
       * de contato já foi criada pelo fluxo de bloqueios.
       * Portanto, não cria outra pendência preventiva.
       */
      if (
        possuiBloqueioDoFeriado &&
        atendimentoJaTratado
      ) {
        return;
      }

      const diasAteFeriado =
        diferencaDiasPendencias_(
          contexto.hoje,
          dataFeriado
        );

      if (
        diasAteFeriado < 0 ||

        diasAteFeriado >
          CONFIG_PENDENCIAS_AUTOMATICAS
            .DIAS_ANTECEDENCIA_FERIADO ||

        dataFeriado.getTime() <
          inicio.getTime() ||

        dataFeriado.getTime() >
          termino.getTime() ||

        !pacienteTemAtendimentoNoDiaSemana_(
          paciente,
          agendamentos,
          dataFeriado.getDay()
        )
      ) {
        return;
      }

      let prazo =
        adicionarDiasPendencias_(
          dataFeriado,
          -1
        );

      if (
        prazo.getTime() <
          contexto.hoje.getTime()
      ) {
        prazo = contexto.hoje;
      }

      adicionarPendencia_(
        contexto,

        criarIdPendencia_(
          paciente.id,
          'FERIADO-' +
            chaveDataPendencias_(
              dataFeriado
            )
        ),

        diasAteFeriado <= 2
          ? 'Urgente'
          : 'Alta',

        'Avisar paciente sobre feriado e alteração da data — ' +
          feriado.descricao,

        paciente,
        'Recepção',
        dataFeriado,
        prazo
      );
    }
  );
}


function gerarPendenciaSessoesSemAgenda_(
  paciente,
  agendamentos,
  contexto
) {
  if (
    normalizarTextoPendencias_(
      paciente.status
    ) !== 'em tratamento' ||

    paciente.sessoesRestantes <= 0
  ) {
    return;
  }

  const possuiSessaoFutura =
    agendamentos.some(
      function (agendamento) {
        return (
          normalizarTextoPendencias_(
            agendamento.evento
          ) === 'sessao' &&

          normalizarTextoPendencias_(
            agendamento.status
          ) === 'agendado' &&

          dataValidaPendencias_(
            agendamento.data
          ) &&

          removerHorarioPendencias_(
            agendamento.data
          ).getTime() >=
            contexto.hoje.getTime()
        );
      }
    );

  if (possuiSessaoFutura) {
    return;
  }

  const ciclo =
    obterCicloAtualPendencias_(
      agendamentos
    );

  adicionarPendencia_(
    contexto,

    criarIdPendencia_(
      paciente.id,
      'SEM-AGENDA-' + ciclo.id
    ),

    'Urgente',

    'Regularizar agendamento das sessões restantes',

    paciente,
    'Recepção',
    contexto.hoje,
    contexto.hoje
  );
}


function gerarPendenciasDesfechoRenovacao_(
  paciente,
  agendamentos,
  contexto
) {
  const status =
    normalizarTextoPendencias_(
      paciente.status
    );

  const desfecho =
    normalizarTextoPendencias_(
      paciente.desfecho
    );

  const ciclo =
    obterCicloAtualPendencias_(
      agendamentos
    );

  const altaPorAbandono =
    possuiTresFaltasNaoJustificadas_(
      ciclo
    );

  if (
    status === 'ciclo concluido' &&
    !desfecho &&
    !altaPorAbandono
  ) {
    adicionarPendencia_(
      contexto,

      criarIdPendencia_(
        paciente.id,
        'DESFECHO-' + ciclo.id
      ),

      'Urgente',

      'Registrar desfecho: Alta, APS ou Renovação',

      paciente,

      paciente.fisioterapeuta ||
        'Fisioterapeuta',

      paciente.dataTermino,
      contexto.hoje
    );
  }

  if (
    desfecho.indexOf('abandono') !== -1
  ) {
    adicionarPendencia_(
      contexto,

      criarIdPendencia_(
        paciente.id,
        'COMUNICAR-ABANDONO-' +
          ciclo.id
      ),

      'Urgente',

      'Comunicar alta por abandono ao paciente',

      paciente,
      'Recepção',
      paciente.dataTermino,

      proximoDiaUtilPendencias_(
        contexto.hoje
      )
    );
  }

  if (
    desfecho.indexOf('renovacao') === -1
  ) {
    return;
  }

  const idPlanejamento =
    criarIdPendencia_(
      paciente.id,
      'PLANEJAR-RENOVACAO-' +
        ciclo.id
    );

  adicionarPendencia_(
    contexto,
    idPlanejamento,
    'Alta',
    'Definir planejamento da renovação',
    paciente,

    paciente.fisioterapeuta ||
      'Fisioterapeuta',

    paciente.dataTermino,

    proximoDiaUtilPendencias_(
      contexto.hoje
    )
  );

  const planejamentoSalvo =
    status ===
      'avaliado - aguardando agendamento' &&
    paciente.sessoesPrescritas > 0 &&
    paciente.sessoesRestantes > 0 &&
    Boolean(
      chaveHorarioPendencias_(
        paciente.horario
      )
    ) &&
    Boolean(
      String(
        paciente.tipoAtendimento || ''
      ).trim()
    ) &&
    Boolean(
      String(
        paciente.fisioterapeuta || ''
      ).trim()
    ) &&
    Object.keys(
      paciente.diasSemana || {}
    ).some(function(dia) {
      return paciente.diasSemana[dia];
    });

  if (
    contexto.idsConcluidos[
      idPlanejamento
    ] &&
    planejamentoSalvo
  ) {
    adicionarPendencia_(
      contexto,

      criarIdPendencia_(
        paciente.id,
        'AGENDAR-RENOVACAO-' +
          ciclo.id
      ),

      'Alta',
      'Agendar ciclo de renovação',
      paciente,
      'Recepção',
      contexto.hoje,
      contexto.hoje
    );
  }
}
function gerarPendenciaFaltaUltimaSessao_(
  paciente,
  agendamentos,
  contexto
) {
  const ultimaSessao =
    localizarUltimaSessaoPendencias_(
      agendamentos
    );

  if (!ultimaSessao) {
    return;
  }

  const status =
    normalizarTextoPendencias_(
      ultimaSessao.status
    );

  if (
    status !== 'falta justificada' &&
    status !== 'falta nao justificada'
  ) {
    return;
  }

  adicionarPendencia_(
    contexto,

    criarIdPendencia_(
      paciente.id,
      'FALTA-FINAL-' +
        ultimaSessao.idAgendamento
    ),

    'Urgente',
    'Entregar documentação ao paciente',
    paciente,
    'Recepção',
    ultimaSessao.data,
    ''
  );
}


function gerarPendenciaMensalRegulacao_(
  contexto
) {
  if (contexto.hoje.getDate() < 15) {
    return;
  }

  const referencia = new Date(
    contexto.hoje.getFullYear(),
    contexto.hoje.getMonth(),
    15
  );

  const competencia =
    Utilities.formatDate(
      referencia,
      Session.getScriptTimeZone(),
      'yyyy-MM'
    );

  adicionarPendencia_(
    contexto,
    'PEND-VAGAS-REGULACAO-' +
      competencia,

    contexto.hoje.getDate() === 15
      ? 'Alta'
      : 'Urgente',

    'Enviar vagas mensais à Regulação',

    pacienteVazioPendencias_(),
    'Recepção',
    referencia,
    referencia
  );
}


function gerarPendenciasBloqueios_(
  contexto
) {
  contexto.bloqueios.forEach(
    function (bloqueio) {
      contexto.agendamentos.forEach(
        function (agendamento) {
          if (
            normalizarTextoPendencias_(
              agendamento.status
            ) !== 'agendado' ||

            !dataValidaPendencias_(
              agendamento.data
            ) ||

            removerHorarioPendencias_(
              agendamento.data
            ).getTime() <
              contexto.hoje.getTime() ||

            !bloqueioAfetaAgendamento_(
              bloqueio,
              agendamento
            )
          ) {
            return;
          }

          const paciente =
            contexto.pacientesPorId[
              agendamento.idPaciente
            ] || {
              id:
                agendamento.idPaciente,

              nome:
                agendamento.nomePaciente,

              prontuario:
                agendamento.prontuario,

              telefone: ''
            };

          adicionarPendencia_(
            contexto,

            criarIdPendencia_(
              paciente.id,

              'BLOQUEIO-' +
                agendamento.idAgendamento
            ),

            'Alta',

            'Avisar paciente sobre alteração do agendamento' +
              (
                bloqueio.motivo
                  ? ' — ' +
                    bloqueio.motivo
                  : ''
              ),

            paciente,
            'Recepção',
            agendamento.data,
            contexto.hoje
          );
        }
      );
    }
  );
}


function bloqueioAfetaAgendamento_(
  bloqueio,
  agendamento
) {
  if (
    chaveDataPendencias_(
      bloqueio.data
    ) !==
    chaveDataPendencias_(
      agendamento.data
    )
  ) {
    return false;
  }

  const profissionalBloqueio =
    normalizarTextoPendencias_(
      bloqueio.fisioterapeuta
    );

  const profissionalAgendamento =
    normalizarTextoPendencias_(
      agendamento.fisioterapeuta
    );

  if (
    profissionalBloqueio &&
    profissionalBloqueio !== 'todos' &&
    profissionalBloqueio !==
      profissionalAgendamento
  ) {
    return false;
  }

  const abrangencia =
    normalizarTextoPendencias_(
      bloqueio.abrangencia
    );

  if (abrangencia === 'dia inteiro') {
    return true;
  }

  if (
    abrangencia === 'turno inteiro'
  ) {
    const horarioBloqueio =
      chaveHorarioPendencias_(
        bloqueio.horario
      );

    if (!horarioBloqueio) {
      return true;
    }

    return (
      obterTurnoPendencias_(
        horarioBloqueio
      ) ===
      obterTurnoPendencias_(
        agendamento.horario
      )
    );
  }

  const horarioBloqueio =
    chaveHorarioPendencias_(
      bloqueio.horario
    );

  return (
    !horarioBloqueio ||
    horarioBloqueio ===
      chaveHorarioPendencias_(
        agendamento.horario
      )
  );
}


function combinarPendencias_(
  existentes,
  geradas,
  idsConcluidos
) {
  const geradasPorId = {};
  const resultado = [];
  const idsUsados = {};

  geradas.forEach(function (linha) {
    geradasPorId[
      String(linha[0] || '').trim()
    ] = linha;
  });

  existentes.forEach(
    function (linhaExistente) {
      let id = String(
        linhaExistente[0] || ''
      ).trim();

      /*
       * Remove do fluxo ativo as antigas pendências
       * genéricas criadas diretamente pelo bloqueio.
       *
       * O novo fluxo utiliza somente a pendência
       * específica do cancelamento e da reposição.
       */
      if (
        normalizarTextoPendencias_(id)
          .indexOf(
            'pend-bloqueio-'
          ) === 0
      ) {
        return;
      }

      if (
        !id ||
        idsConcluidos[id] ||
        idsUsados[id]
      ) {
        return;
      }

      if (!geradasPorId[id]) {
        const idEquivalente =
          localizarIdEquivalentePendencias_(
            linhaExistente,
            geradasPorId
          );

        if (idEquivalente) {
          id = idEquivalente;
        }
      }

      if (idsUsados[id]) {
        return;
      }

      if (
        !geradasPorId[id] &&
        pendenciaAutomaticaTransitoria_(
          linhaExistente[2]
        )
      ) {
        return;
      }

      if (geradasPorId[id]) {
        const atualizada =
          geradasPorId[id].slice();

        atualizada[
          CONFIG_PENDENCIAS_AUTOMATICAS
            .PENDENCIAS.SITUACAO - 1
        ] =
          linhaExistente[
            CONFIG_PENDENCIAS_AUTOMATICAS
              .PENDENCIAS.SITUACAO - 1
          ] || 'Pendente';

        atualizada[
          CONFIG_PENDENCIAS_AUTOMATICAS
            .PENDENCIAS
            .DATA_CONCLUSAO - 1
        ] =
          linhaExistente[
            CONFIG_PENDENCIAS_AUTOMATICAS
              .PENDENCIAS
              .DATA_CONCLUSAO - 1
          ] || '';

        resultado.push(atualizada);

        delete geradasPorId[id];
      } else {
        resultado.push(
          linhaExistente.slice()
        );
      }

      idsUsados[id] = true;
    }
  );

  Object.keys(geradasPorId).forEach(
    function (id) {
      if (
        !idsConcluidos[id] &&
        !idsUsados[id]
      ) {
        resultado.push(
          geradasPorId[id]
        );

        idsUsados[id] = true;
      }
    }
  );

  return resultado;
}


/**
 * Identifica pendências automáticas que representam
 * um estado atual do sistema.
 *
 * Se a condição que gerou uma dessas pendências deixar
 * de existir, ela deve sair da aba Pendências em vez de
 * permanecer aberta indefinidamente.
 *
 * Pendências criadas diretamente por outros módulos,
 * como o encerramento por desistência, não entram aqui.
 */
function pendenciaAutomaticaTransitoria_(
  descricao
) {
  const texto =
    normalizarTextoPendencias_(
      descricao
    );

  const descricoes = [
    'registrar resultado do atendimento',
    'informar falta na avaliacao a regulacao',
    'agendar avaliacao',
    'definir plano de tratamento',
    'agendar ciclo de sessoes',
    'preparar documentacao de encerramento do tratamento',
    'contatar paciente - 2 faltas nao justificadas',
    'registrar alta por abandono',
    'contatar paciente sobre cancelamento e reposicao',
    'regularizar agendamento das sessoes restantes',
    'registrar desfecho: alta, aps ou renovacao',
    'comunicar alta por abandono ao paciente',
    'definir planejamento da renovacao',
    'agendar ciclo de renovacao',
    'entregar documentacao ao paciente',
    'enviar vagas mensais a regulacao'
  ];

  if (
    descricoes.indexOf(
      texto
    ) !== -1
  ) {
    return true;
  }

  return (
    texto.indexOf(
      'avisar paciente sobre feriado e alteracao da data'
    ) === 0 ||
    texto.indexOf(
      'avisar paciente sobre alteracao do agendamento'
    ) === 0
  );
}


function localizarIdEquivalentePendencias_(
  linhaExistente,
  geradasPorId
) {
  const categoriaExistente =
    categoriaPendencia_(
      linhaExistente[2]
    );

  if (!categoriaExistente) {
    return '';
  }

  const prontuario =
    normalizarTextoPendencias_(
      linhaExistente[4]
    );

  const paciente =
    normalizarTextoPendencias_(
      linhaExistente[3]
    );

  const ids =
    Object.keys(geradasPorId);

  for (
    let indice = 0;
    indice < ids.length;
    indice++
  ) {
    const id = ids[indice];

    const linhaGerada =
      geradasPorId[id];

    if (
      categoriaPendencia_(
        linhaGerada[2]
      ) !== categoriaExistente
    ) {
      continue;
    }

    const mesmoProntuario =
      prontuario &&
      prontuario ===
        normalizarTextoPendencias_(
          linhaGerada[4]
        );

    const mesmoPaciente =
      paciente &&
      paciente ===
        normalizarTextoPendencias_(
          linhaGerada[3]
        );

    if (
      mesmoProntuario ||
      mesmoPaciente
    ) {
      return id;
    }
  }

  return '';
}


function categoriaPendencia_(
  descricao
) {
  const texto =
    normalizarTextoPendencias_(
      descricao
    );

  if (
    texto.indexOf(
      'agendar ciclo de sessoes'
    ) !== -1
  ) {
    return 'AGENDAR-CICLO';
  }

  if (
    texto.indexOf(
      'registrar desfecho'
    ) !== -1
  ) {
    return 'DESFECHO';
  }

  if (
    texto.indexOf(
      'preparar relatorio/sumario de alta'
    ) !== -1 ||

    texto.indexOf(
      'preparar documentacao de encerramento'
    ) !== -1
  ) {
    return 'DOCUMENTACAO';
  }

  if (
    texto.indexOf(
      'paciente faltou a ultima sessao'
    ) !== -1 ||

    texto.indexOf(
      'entregar documentacao ao paciente'
    ) !== -1
  ) {
    return 'ENTREGA-DOCUMENTACAO';
  }

  return '';
}
function adicionarPendencia_(
  contexto,
  id,
  prioridade,
  descricao,
  paciente,
  responsavel,
  dataReferencia,
  prazo
) {
  if (
    !id ||
    contexto.idsConcluidos[id]
  ) {
    return;
  }

  const jaExiste =
    contexto.pendencias.some(
      function (linha) {
        return linha[0] === id;
      }
    );

  if (jaExiste) {
    return;
  }

  paciente =
    paciente ||
    pacienteVazioPendencias_();

  contexto.pendencias.push([
    id,
    prioridade,
    descricao,

    paciente.nome || '',
    paciente.prontuario || '',
    paciente.telefone || '',

    responsavel,

    dataValidaPendencias_(
      dataReferencia
    )
      ? removerHorarioPendencias_(
          dataReferencia
        )
      : '',

    dataValidaPendencias_(prazo)
      ? removerHorarioPendencias_(
          prazo
        )
      : '',

    'Pendente',
    ''
  ]);
}


function gravarPendenciasAtivas_(
  abaPendencias,
  pendencias
) {
  const ultimaLinha =
    abaPendencias.getLastRow();

  if (ultimaLinha >= 2) {
    abaPendencias
      .getRange(
        2,
        1,
        ultimaLinha - 1,
        CONFIG_PENDENCIAS_AUTOMATICAS
          .QUANTIDADE_COLUNAS_PENDENCIAS
      )
      .clearContent();
  }

  if (pendencias.length === 0) {
    return;
  }

  abaPendencias
    .getRange(
      2,
      1,
      pendencias.length,
      CONFIG_PENDENCIAS_AUTOMATICAS
        .QUANTIDADE_COLUNAS_PENDENCIAS
    )
    .setValues(pendencias);

  abaPendencias
    .getRange(
      2,
      CONFIG_PENDENCIAS_AUTOMATICAS
        .PENDENCIAS.DATA_REFERENCIA,
      pendencias.length,
      2
    )
    .setNumberFormat('dd/MM/yyyy');

  abaPendencias
    .getRange(
      2,
      CONFIG_PENDENCIAS_AUTOMATICAS
        .PENDENCIAS.DATA_CONCLUSAO,
      pendencias.length,
      1
    )
    .setNumberFormat('dd/MM/yyyy');

  const validacaoSituacao =
    SpreadsheetApp
      .newDataValidation()
      .requireValueInList(
        [
          'Pendente',
          'Concluída'
        ],
        true
      )
      .setAllowInvalid(false)
      .setHelpText(
        'Para concluir, selecione Concluída e informe a Data da conclusão.'
      )
      .build();

  const validacaoData =
    SpreadsheetApp
      .newDataValidation()
      .requireDate()
      .setAllowInvalid(false)
      .setHelpText(
        'Informe a data real em que a ação foi concluída.'
      )
      .build();

  abaPendencias
    .getRange(
      2,
      CONFIG_PENDENCIAS_AUTOMATICAS
        .PENDENCIAS.SITUACAO,
      pendencias.length,
      1
    )
    .setDataValidation(
      validacaoSituacao
    );

  abaPendencias
    .getRange(
      2,
      CONFIG_PENDENCIAS_AUTOMATICAS
        .PENDENCIAS.DATA_CONCLUSAO,
      pendencias.length,
      1
    )
    .setDataValidation(
      validacaoData
    );
}


function ordenarPendencias_(
  pendencias
) {
  const pesos = {
    Urgente: 1,
    Alta: 2,
    Média: 3,
    Baixa: 4
  };

  pendencias.sort(
    function (a, b) {
      const pesoA =
        pesos[a[1]] || 99;

      const pesoB =
        pesos[b[1]] || 99;

      if (pesoA !== pesoB) {
        return pesoA - pesoB;
      }

      const dataA =
        dataValidaPendencias_(a[8])
          ? removerHorarioPendencias_(
              a[8]
            ).getTime()
          : Number.MAX_SAFE_INTEGER;

      const dataB =
        dataValidaPendencias_(b[8])
          ? removerHorarioPendencias_(
              b[8]
            ).getTime()
          : Number.MAX_SAFE_INTEGER;

      return dataA - dataB;
    }
  );
}


function obterCicloAtualPendencias_(
  agendamentos
) {
  const sessoes =
    agendamentos.filter(
      function (agendamento) {
        const evento =
          normalizarTextoPendencias_(
            agendamento.evento
          );

        const status =
          normalizarTextoPendencias_(
            agendamento.status
          );

        /*
         * Linhas canceladas por remanejamento são
         * histórico do ciclo, não programação vigente.
         * Ignorá-las aqui evita feriados, última sessão
         * e outras pendências baseadas no padrão antigo.
         */
        return (
          evento === 'sessao' &&
          status !==
            'cancelado por remanejamento'
        );
      }
    );

  if (sessoes.length === 0) {
    return {
      id: 'SEM-CICLO',
      numero: 0,
      sessoes: []
    };
  }

  let maiorNumero = 0;

  sessoes.forEach(
    function (sessao) {
      if (
        sessao.cicloNumero >
          maiorNumero
      ) {
        maiorNumero =
          sessao.cicloNumero;
      }
    }
  );

  const sessoesCiclo =
    sessoes.filter(
      function (sessao) {
        return (
          sessao.cicloNumero ===
          maiorNumero
        );
      }
    );

  const id = String(
    (
      sessoesCiclo[0] &&
      sessoesCiclo[0].idCiclo
    ) ||
      'CICLO-' + maiorNumero
  ).trim();

  return {
    id:
      id ||
      'CICLO-' + maiorNumero,

    numero: maiorNumero,
    sessoes: sessoesCiclo
  };
}


function localizarUltimaSessaoPendencias_(
  agendamentos
) {
  const ciclo =
    obterCicloAtualPendencias_(
      agendamentos
    );

  let resultado = null;

  ciclo.sessoes.forEach(
    function (sessao) {
      if (
        sessao.numeroSessao !==
          sessao.totalPrescrito ||

        sessao.totalPrescrito <= 0
      ) {
        return;
      }

      if (
        !resultado ||
        (
          dataValidaPendencias_(
            sessao.data
          ) &&
          (
            !dataValidaPendencias_(
              resultado.data
            ) ||

            sessao.data.getTime() >
              resultado.data.getTime()
          )
        )
      ) {
        resultado = sessao;
      }
    }
  );

  return resultado;
}


function obterUltimoAgendamento_(
  agendamentos
) {
  if (
    !agendamentos ||
    agendamentos.length === 0
  ) {
    return null;
  }

  return agendamentos
    .slice()
    .sort(
      compararAgendamentosPorData_
    )
    .pop();
}


function compararAgendamentosPorData_(
  a,
  b
) {
  const tempoA =
    dataValidaPendencias_(a.data)
      ? a.data.getTime()
      : 0;

  const tempoB =
    dataValidaPendencias_(b.data)
      ? b.data.getTime()
      : 0;

  return tempoA - tempoB;
}


function planoTratamentoCompletoPendencias_(
  paciente
) {
  const possuiDia =
    Object.keys(
      paciente.diasSemana
    ).some(function (dia) {
      return paciente.diasSemana[dia];
    });

  return Boolean(
    paciente.sessoesPrescritas > 0 &&
    chaveHorarioPendencias_(
      paciente.horario
    ) &&
    paciente.tipoAtendimento &&
    paciente.fisioterapeuta &&
    possuiDia
  );
}


function pacienteTemAtendimentoNoDiaSemana_(
  paciente,
  agendamentos,
  diaSemana
) {
  if (
    paciente.diasSemana[diaSemana]
  ) {
    return true;
  }

  const ciclo =
    obterCicloAtualPendencias_(
      agendamentos
    );

  return ciclo.sessoes.some(
    function (sessao) {
      return (
        dataValidaPendencias_(
          sessao.data
        ) &&
        sessao.data.getDay() ===
          diaSemana
      );
    }
  );
}


function criarIdPendencia_(
  idPaciente,
  tipo
) {
  return (
    'PEND-' +
    limparIdPendencias_(tipo) +
    '-' +
    limparIdPendencias_(
      idPaciente
    )
  );
}


function limparIdPendencias_(
  valor
) {
  return normalizarTextoPendencias_(
    valor
  )
    .toUpperCase()
    .replace(
      /[^A-Z0-9]+/g,
      '-'
    )
    .replace(
      /^-+|-+$/g,
      ''
    );
}
function pacienteVazioPendencias_() {
  return {
    id: '',
    nome: '',
    prontuario: '',
    telefone: ''
  };
}


function valorAtivoPendencias_(
  valor
) {
  if (
    valor === true ||
    Number(valor) === 1
  ) {
    return true;
  }

  const texto =
    normalizarTextoPendencias_(
      valor
    );

  return (
    texto === 'sim' ||
    texto === 'x' ||
    texto === 'verdadeiro'
  );
}


function dataValidaPendencias_(
  valor
) {
  return (
    valor instanceof Date &&
    !isNaN(valor.getTime())
  );
}


function removerHorarioPendencias_(
  data
) {
  const resultado =
    new Date(data);

  resultado.setHours(
    0,
    0,
    0,
    0
  );

  return resultado;
}


function adicionarDiasPendencias_(
  data,
  quantidade
) {
  const resultado =
    removerHorarioPendencias_(data);

  resultado.setDate(
    resultado.getDate() +
      quantidade
  );

  return resultado;
}


function proximoDiaUtilPendencias_(
  data
) {
  const resultado =
    adicionarDiasPendencias_(
      dataValidaPendencias_(data)
        ? data
        : new Date(),
      1
    );

  while (
    resultado.getDay() === 0 ||
    resultado.getDay() === 6
  ) {
    resultado.setDate(
      resultado.getDate() + 1
    );
  }

  return resultado;
}


function diferencaDiasPendencias_(
  dataInicial,
  dataFinal
) {
  const inicio =
    removerHorarioPendencias_(
      dataInicial
    );

  const fim =
    removerHorarioPendencias_(
      dataFinal
    );

  return Math.round(
    (
      fim.getTime() -
      inicio.getTime()
    ) /
      86400000
  );
}


function chaveDataPendencias_(
  data
) {
  return Utilities.formatDate(
    removerHorarioPendencias_(data),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );
}


function chaveHorarioPendencias_(
  valor
) {
  if (
    valor instanceof Date &&
    !isNaN(valor.getTime())
  ) {
    return Utilities.formatDate(
      valor,
      Session.getScriptTimeZone(),
      'HH:mm'
    );
  }

  const texto = String(
    valor || ''
  ).trim();

  const resultado =
    texto.match(
      /(\d{1,2}):(\d{2})/
    );

  if (!resultado) {
    return '';
  }

  return (
    String(
      Number(resultado[1])
    ).padStart(2, '0') +
    ':' +
    resultado[2]
  );
}


function obterTurnoPendencias_(
  horario
) {
  const chave =
    chaveHorarioPendencias_(
      horario
    );

  if (!chave) {
    return '';
  }

  return (
    Number(
      chave.split(':')[0]
    ) < 13
      ? 'manha'
      : 'tarde'
  );
}


function normalizarTextoPendencias_(
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


/**
 * Gera a ação de comunicação ou de reagendamento quando
 * uma avaliação foi cancelada por um bloqueio da clínica.
 */
function gerarPendenciasReagendamentoAvaliacao_(
  paciente,
  agendamentos,
  contexto
) {
  agendamentos.forEach(
    function (avaliacao) {
      if (
        normalizarTextoPendencias_(
          avaliacao.evento
        ) !== 'avaliacao' ||

        normalizarTextoPendencias_(
          avaliacao.status
        ) !==
          'cancelado pela clinica'
      ) {
        return;
      }

      const foiReagendada =
        normalizarTextoPendencias_(
          avaliacao.motivo
        ).indexOf(
          'reagendada automaticamente'
        ) !== -1;

      adicionarPendencia_(
        contexto,

        criarIdPendencia_(
          paciente.id,
          'REAGENDAMENTO-AVALIACAO-' +
            avaliacao.idAgendamento
        ),

        foiReagendada
          ? 'Alta'
          : 'Urgente',

        foiReagendada
          ? 'Contatar paciente sobre reagendamento da avaliação'
          : 'Reagendar avaliação cancelada pela clínica',

        paciente,
        'Recepção',
        avaliacao.data,

        proximoDiaUtilPendencias_(
          contexto.hoje
        )
      );
    }
  );
}


/**
 * Remove a pendência genérica de desfecho quando o ciclo
 * já atingiu três faltas não justificadas.
 *
 * Nesse caso, permanece somente a pendência específica:
 * "Registrar alta por abandono".
 */
function removerDesfechosGenericosPorAbandono_(
  pendenciasExistentes,
  contexto
) {
  const ciclosComAbandono = [];

  contexto.pacientes.forEach(
    function (paciente) {
      const agendamentos =
        contexto
          .agendamentosPorPaciente[
            paciente.id
          ] || [];

      const ciclo =
        obterCicloAtualPendencias_(
          agendamentos
        );

      if (
        !ciclo.id ||
        !possuiTresFaltasNaoJustificadas_(
          ciclo
        )
      ) {
        return;
      }

      ciclosComAbandono.push({
        idPaciente:
          normalizarTextoPendencias_(
            paciente.id
          ),

        idCiclo:
          normalizarTextoPendencias_(
            ciclo.id
          )
      });
    }
  );

  if (ciclosComAbandono.length === 0) {
    return pendenciasExistentes;
  }

  return pendenciasExistentes.filter(
    function (linha) {
      const descricao =
        normalizarTextoPendencias_(
          linha[
            CONFIG_PENDENCIAS_AUTOMATICAS
              .PENDENCIAS.PENDENCIA - 1
          ]
        );

      if (
        descricao !==
        'registrar desfecho: alta, aps ou renovacao'
      ) {
        return true;
      }

      const idPendencia =
        normalizarTextoPendencias_(
          linha[
            CONFIG_PENDENCIAS_AUTOMATICAS
              .PENDENCIAS.ID - 1
          ]
        );

      const deveRemover =
        ciclosComAbandono.some(
          function (registro) {
            return (
              idPendencia.indexOf(
                registro.idPaciente
              ) !== -1 &&

              idPendencia.indexOf(
                registro.idCiclo
              ) !== -1
            );
          }
        );

      return !deveRemover;
    }
  );
}


/**
 * Informa se o ciclo atual atingiu três faltas
 * não justificadas.
 */
function possuiTresFaltasNaoJustificadas_(
  ciclo
) {
  if (
    !ciclo ||
    !Array.isArray(ciclo.sessoes)
  ) {
    return false;
  }

  let quantidadeFaltas = 0;

  ciclo.sessoes.forEach(
    function (sessao) {
      if (
        normalizarTextoPendencias_(
          sessao.status
        ) ===
          'falta nao justificada'
      ) {
        quantidadeFaltas++;
      }
    }
  );

  return quantidadeFaltas >= 3;
}
