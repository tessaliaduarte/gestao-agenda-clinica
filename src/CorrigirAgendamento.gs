const CONFIG_CORRIGIR_AGENDAMENTO = {
  ABAS: {
    CADASTRO: 'Cadastro de Pacientes',
    AGENDAMENTOS: 'Agendamentos',
    PENDENCIAS: 'Pendências',
    HISTORICO_AJUSTES: 'Histórico de Ajustes'
  },

  CADASTRO: {
    ID_PACIENTE: 1,
    PRONTUARIO: 2,
    NOME: 3,
    CPF: 4,
    TELEFONE: 5,
    HORARIO: 6,
    TIPO_ATENDIMENTO: 7,
    SESSOES_PRESCRITAS: 14,
    SESSOES_REALIZADAS: 15,
    SESSOES_RESTANTES: 16,
    DATA_INICIO: 19,
    DATA_TERMINO: 20,
    STATUS: 21,
    FISIOTERAPEUTA: 22
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

  QUANTIDADE_COLUNAS_CADASTRO: 24,
  QUANTIDADE_COLUNAS_AGENDAMENTOS: 22,
  QUANTIDADE_COLUNAS_PENDENCIAS: 11,

  CABECALHO_HISTORICO: [
    'ID Ajuste',
    'Data e Hora do Ajuste',
    'ID Paciente',
    'Prontuário',
    'Paciente',
    'ID Ciclo',
    'Ciclo Nº',
    'ID Agendamento',
    'Data Original',
    'Dia',
    'Horário Original',
    'Fisioterapeuta Original',
    'Tipo do Grupo',
    'Nº Sessão',
    'Total Prescrito',
    'Status Original',
    'Motivo do Ajuste'
  ]
};


/**
 * Abre a tela para localizar o paciente e corrigir
 * um ciclo agendado incorretamente.
 */
function abrirFormularioCorrigirAgendamento() {
  const html = HtmlService
    .createHtmlOutputFromFile(
      'FormularioCorrigirAgendamento'
    )
    .setWidth(900)
    .setHeight(680);

  SpreadsheetApp.getUi().showModalDialog(
    html,
    'Corrigir agendamento'
  );
}


/**
 * Pesquisa pacientes por nome, CPF, prontuário ou ID.
 */
function pesquisarPacientesCorrecaoAgendamento(
  termoPesquisa
) {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const abaCadastro =
    ss.getSheetByName(
      CONFIG_CORRIGIR_AGENDAMENTO
        .ABAS.CADASTRO
    );

  if (!abaCadastro) {
    throw new Error(
      'A aba "Cadastro de Pacientes" não foi encontrada.'
    );
  }

  const termo = String(
    termoPesquisa || ''
  ).trim();

  if (!termo) {
    throw new Error(
      'Informe o nome, CPF, prontuário ou ID do paciente.'
    );
  }

  const ultimaLinha =
    abaCadastro.getLastRow();

  if (ultimaLinha < 2) {
    return [];
  }

  const dados = abaCadastro
    .getRange(
      2,
      1,
      ultimaLinha - 1,
      CONFIG_CORRIGIR_AGENDAMENTO
        .QUANTIDADE_COLUNAS_CADASTRO
    )
    .getValues();

  const termoNormalizado =
    normalizarTextoCorrecaoAgendamento_(
      termo
    );

  const termoNumerico =
    somenteNumerosCorrecaoAgendamento_(
      termo
    );

  const resultados = [];

  dados.forEach(function(linha) {
    const id = String(
      linha[
        CONFIG_CORRIGIR_AGENDAMENTO
          .CADASTRO.ID_PACIENTE - 1
      ] || ''
    ).trim();

    if (!id) {
      return;
    }

    const prontuario = String(
      linha[
        CONFIG_CORRIGIR_AGENDAMENTO
          .CADASTRO.PRONTUARIO - 1
      ] || ''
    ).trim();

    const nome = String(
      linha[
        CONFIG_CORRIGIR_AGENDAMENTO
          .CADASTRO.NOME - 1
      ] || ''
    ).trim();

    const cpf = String(
      linha[
        CONFIG_CORRIGIR_AGENDAMENTO
          .CADASTRO.CPF - 1
      ] || ''
    ).trim();

    const corresponde =
      normalizarTextoCorrecaoAgendamento_(
        id
      ).indexOf(termoNormalizado) !== -1 ||

      normalizarTextoCorrecaoAgendamento_(
        prontuario
      ).indexOf(termoNormalizado) !== -1 ||

      normalizarTextoCorrecaoAgendamento_(
        nome
      ).indexOf(termoNormalizado) !== -1 ||

      (
        termoNumerico &&
        somenteNumerosCorrecaoAgendamento_(
          cpf
        ).indexOf(termoNumerico) !== -1
      );

    if (!corresponde) {
      return;
    }

    resultados.push({
      id: id,
      prontuario: prontuario,
      nome: nome,
      cpf: cpf,
      status: String(
        linha[
          CONFIG_CORRIGIR_AGENDAMENTO
            .CADASTRO.STATUS - 1
        ] || ''
      ).trim(),
      sessoesPrescritas:
        Number(
          linha[
            CONFIG_CORRIGIR_AGENDAMENTO
              .CADASTRO
              .SESSOES_PRESCRITAS - 1
          ]
        ) || 0,
      sessoesRealizadas:
        Number(
          linha[
            CONFIG_CORRIGIR_AGENDAMENTO
              .CADASTRO
              .SESSOES_REALIZADAS - 1
          ]
        ) || 0
    });
  });

  resultados.sort(function(a, b) {
    return a.nome.localeCompare(
      b.nome,
      'pt-BR'
    );
  });

  return resultados.slice(0, 20);
}


/**
 * Retorna o ciclo mais recente e informa se ele
 * pode ser desfeito com segurança.
 */
function obterDadosCorrecaoAgendamento(
  idPaciente
) {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const contexto =
    obterContextoCorrecaoAgendamento_(
      ss,
      idPaciente
    );

  return montarResumoCorrecaoAgendamento_(
    contexto
  );
}


/**
 * Arquiva e remove automaticamente as sessões
 * ainda não iniciadas de um ciclo incorreto.
 */
function corrigirAgendamentoDoCiclo(
  idPaciente,
  idCicloInformado
) {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const lock =
    LockService.getDocumentLock();

  let bloqueioObtido = false;
  let resultado = null;

  try {
    lock.waitLock(30000);
    bloqueioObtido = true;

    const contexto =
      obterContextoCorrecaoAgendamento_(
        ss,
        idPaciente
      );

    if (
      String(
        contexto.idCiclo || ''
      ).trim() !==
      String(
        idCicloInformado || ''
      ).trim()
    ) {
      throw new Error(
        'O ciclo foi alterado depois que a tela foi aberta. Pesquise o paciente novamente.'
      );
    }

    validarCicloParaCorrecaoAgendamento_(
      contexto
    );

    const abaHistorico =
      obterOuCriarHistoricoAjustes_(
        ss
      );

    const agora = new Date();

    const idAjuste =
      gerarIdAjusteCorrecaoAgendamento_(
        agora
      );

    registrarHistoricoCorrecaoAgendamento_(
      abaHistorico,
      contexto,
      idAjuste,
      agora
    );

    limparSessoesCorrecaoAgendamento_(
      contexto.abaAgendamentos,
      contexto.sessoes
    );

    prepararPacienteParaReagendamento_(
      contexto
    );

    removerPendenciasVinculadasCorrecao_(
      ss,
      contexto
    );

    SpreadsheetApp.flush();

    resultado = {
      sucesso: true,
      idAjuste: idAjuste,
      paciente: contexto.paciente.nome,
      prontuario:
        contexto.paciente.prontuario,
      sessoesArquivadas:
        contexto.sessoes.length,
      mensagem:
        'O agendamento incorreto foi desfeito com segurança. O paciente está pronto para uma nova consulta de disponibilidade.'
    };
  } finally {
    if (bloqueioObtido) {
      try {
        lock.releaseLock();
      } catch (erroLock) {
        // Liberação automática.
      }
    }
  }

  return resultado;
}
/**
 * Localiza o paciente e o ciclo mais recente.
 */
function obterContextoCorrecaoAgendamento_(
  ss,
  idPaciente
) {
  const abaCadastro =
    ss.getSheetByName(
      CONFIG_CORRIGIR_AGENDAMENTO
        .ABAS.CADASTRO
    );

  const abaAgendamentos =
    ss.getSheetByName(
      CONFIG_CORRIGIR_AGENDAMENTO
        .ABAS.AGENDAMENTOS
    );

  if (!abaCadastro) {
    throw new Error(
      'A aba "Cadastro de Pacientes" não foi encontrada.'
    );
  }

  if (!abaAgendamentos) {
    throw new Error(
      'A aba "Agendamentos" não foi encontrada.'
    );
  }

  const paciente =
    localizarPacienteCorrecaoAgendamento_(
      abaCadastro,
      idPaciente
    );

  if (!paciente) {
    throw new Error(
      'O paciente não foi encontrado.'
    );
  }

  const sessoesPaciente =
    lerSessoesPacienteCorrecaoAgendamento_(
      abaAgendamentos,
      paciente.id
    );

  if (sessoesPaciente.length === 0) {
    throw new Error(
      'O paciente não possui um ciclo de sessões agendado para corrigir.'
    );
  }

  const cicloAtual =
    localizarCicloMaisRecenteCorrecao_(
      sessoesPaciente
    );

  const sessoesCiclo =
    sessoesPaciente.filter(
      function(sessao) {
        return (
          sessao.idCiclo ===
            cicloAtual.idCiclo &&
          sessao.cicloNumero ===
            cicloAtual.cicloNumero
        );
      }
    );

  sessoesCiclo.sort(function(a, b) {
    const diferencaData =
      a.data.getTime() -
      b.data.getTime();

    if (diferencaData !== 0) {
      return diferencaData;
    }

    return (
      a.numeroSessao -
      b.numeroSessao
    );
  });

  return {
    ss: ss,
    abaCadastro: abaCadastro,
    abaAgendamentos:
      abaAgendamentos,
    paciente: paciente,
    idCiclo: cicloAtual.idCiclo,
    cicloNumero:
      cicloAtual.cicloNumero,
    sessoes: sessoesCiclo
  };
}


/**
 * Localiza o cadastro pelo ID do paciente.
 */
function localizarPacienteCorrecaoAgendamento_(
  abaCadastro,
  idPaciente
) {
  const ultimaLinha =
    abaCadastro.getLastRow();

  if (ultimaLinha < 2) {
    return null;
  }

  const dados = abaCadastro
    .getRange(
      2,
      1,
      ultimaLinha - 1,
      CONFIG_CORRIGIR_AGENDAMENTO
        .QUANTIDADE_COLUNAS_CADASTRO
    )
    .getValues();

  const idProcurado =
    normalizarTextoCorrecaoAgendamento_(
      idPaciente
    );

  for (
    let indice = 0;
    indice < dados.length;
    indice++
  ) {
    const linha = dados[indice];

    const idAtual =
      normalizarTextoCorrecaoAgendamento_(
        linha[
          CONFIG_CORRIGIR_AGENDAMENTO
            .CADASTRO.ID_PACIENTE - 1
        ]
      );

    if (idAtual !== idProcurado) {
      continue;
    }

    return {
      linha: indice + 2,
      valores: linha,
      id: String(
        linha[
          CONFIG_CORRIGIR_AGENDAMENTO
            .CADASTRO.ID_PACIENTE - 1
        ] || ''
      ).trim(),
      prontuario: String(
        linha[
          CONFIG_CORRIGIR_AGENDAMENTO
            .CADASTRO.PRONTUARIO - 1
        ] || ''
      ).trim(),
      nome: String(
        linha[
          CONFIG_CORRIGIR_AGENDAMENTO
            .CADASTRO.NOME - 1
        ] || ''
      ).trim(),
      cpf: String(
        linha[
          CONFIG_CORRIGIR_AGENDAMENTO
            .CADASTRO.CPF - 1
        ] || ''
      ).trim(),
      telefone: String(
        linha[
          CONFIG_CORRIGIR_AGENDAMENTO
            .CADASTRO.TELEFONE - 1
        ] || ''
      ).trim(),
      sessoesPrescritas:
        Number(
          linha[
            CONFIG_CORRIGIR_AGENDAMENTO
              .CADASTRO
              .SESSOES_PRESCRITAS - 1
          ]
        ) || 0,
      sessoesRealizadas:
        Number(
          linha[
            CONFIG_CORRIGIR_AGENDAMENTO
              .CADASTRO
              .SESSOES_REALIZADAS - 1
          ]
        ) || 0,
      status: String(
        linha[
          CONFIG_CORRIGIR_AGENDAMENTO
            .CADASTRO.STATUS - 1
        ] || ''
      ).trim()
    };
  }

  return null;
}


/**
 * Lê somente registros de sessão do paciente.
 */
function lerSessoesPacienteCorrecaoAgendamento_(
  abaAgendamentos,
  idPaciente
) {
  const ultimaLinha =
    abaAgendamentos.getLastRow();

  if (ultimaLinha < 2) {
    return [];
  }

  const dados = abaAgendamentos
    .getRange(
      2,
      1,
      ultimaLinha - 1,
      CONFIG_CORRIGIR_AGENDAMENTO
        .QUANTIDADE_COLUNAS_AGENDAMENTOS
    )
    .getValues();

  const idProcurado =
    normalizarTextoCorrecaoAgendamento_(
      idPaciente
    );

  const resultado = [];

  dados.forEach(function(linha, indice) {
    const idAtual =
      normalizarTextoCorrecaoAgendamento_(
        linha[
          CONFIG_CORRIGIR_AGENDAMENTO
            .AGENDAMENTOS.ID_PACIENTE -
            1
        ]
      );

    const evento =
      normalizarTextoCorrecaoAgendamento_(
        linha[
          CONFIG_CORRIGIR_AGENDAMENTO
            .AGENDAMENTOS.EVENTO - 1
        ]
      );

    const data =
      linha[
        CONFIG_CORRIGIR_AGENDAMENTO
          .AGENDAMENTOS.DATA - 1
      ];

    const status =
      normalizarTextoCorrecaoAgendamento_(
        linha[
          CONFIG_CORRIGIR_AGENDAMENTO
            .AGENDAMENTOS.STATUS - 1
        ]
      );

    /*
     * Sessões antigas preservadas por um remanejamento
     * não fazem parte da agenda ativa que pode ser
     * corrigida/desfeita. Elas permanecem no histórico.
     */
    if (
      idAtual !== idProcurado ||
      evento !== 'sessao' ||
      status ===
        'cancelado por remanejamento' ||
      !(data instanceof Date)
    ) {
      return;
    }

    resultado.push({
      linhaPlanilha: indice + 2,
      valores: linha,
      idAgendamento: String(
        linha[
          CONFIG_CORRIGIR_AGENDAMENTO
            .AGENDAMENTOS.ID_AGENDAMENTO -
            1
        ] || ''
      ).trim(),
      idCiclo: String(
        linha[
          CONFIG_CORRIGIR_AGENDAMENTO
            .AGENDAMENTOS.ID_CICLO - 1
        ] || ''
      ).trim(),
      cicloNumero:
        Number(
          linha[
            CONFIG_CORRIGIR_AGENDAMENTO
              .AGENDAMENTOS.CICLO_NUMERO -
              1
          ]
        ) || 0,
      data: new Date(data),
      dia: String(
        linha[
          CONFIG_CORRIGIR_AGENDAMENTO
            .AGENDAMENTOS.DIA - 1
        ] || ''
      ).trim(),
      horario:
        linha[
          CONFIG_CORRIGIR_AGENDAMENTO
            .AGENDAMENTOS.HORARIO - 1
        ],
      fisioterapeuta: String(
        linha[
          CONFIG_CORRIGIR_AGENDAMENTO
            .AGENDAMENTOS
            .FISIOTERAPEUTA - 1
        ] || ''
      ).trim(),
      tipoGrupo: String(
        linha[
          CONFIG_CORRIGIR_AGENDAMENTO
            .AGENDAMENTOS.TIPO_GRUPO - 1
        ] || ''
      ).trim(),
      numeroSessao:
        Number(
          linha[
            CONFIG_CORRIGIR_AGENDAMENTO
              .AGENDAMENTOS
              .NUMERO_SESSAO - 1
          ]
        ) || 0,
      totalPrescrito:
        Number(
          linha[
            CONFIG_CORRIGIR_AGENDAMENTO
              .AGENDAMENTOS
              .TOTAL_PRESCRITO - 1
          ]
        ) || 0,
      status: String(
        linha[
          CONFIG_CORRIGIR_AGENDAMENTO
            .AGENDAMENTOS.STATUS - 1
        ] || ''
      ).trim()
    });
  });

  return resultado;
}


/**
 * Define qual é o ciclo mais recente.
 */
function localizarCicloMaisRecenteCorrecao_(
  sessoes
) {
  let cicloSelecionado = null;

  sessoes.forEach(function(sessao) {
    if (!cicloSelecionado) {
      cicloSelecionado = sessao;
      return;
    }

    if (
      sessao.cicloNumero >
        cicloSelecionado.cicloNumero
    ) {
      cicloSelecionado = sessao;
      return;
    }

    if (
      sessao.cicloNumero ===
        cicloSelecionado.cicloNumero &&
      sessao.data.getTime() >
        cicloSelecionado.data.getTime()
    ) {
      cicloSelecionado = sessao;
    }
  });

  return {
    idCiclo:
      cicloSelecionado.idCiclo,
    cicloNumero:
      cicloSelecionado.cicloNumero
  };
}
/**
 * Monta os dados apresentados na tela.
 */
function montarResumoCorrecaoAgendamento_(
  contexto
) {
  const motivoBloqueio =
    obterMotivoBloqueioCorrecao_(
      contexto
    );

  const primeiraSessao =
    contexto.sessoes[0];

  const ultimaSessao =
    contexto.sessoes[
      contexto.sessoes.length - 1
    ];

  return {
    paciente: {
      id: contexto.paciente.id,
      prontuario:
        contexto.paciente.prontuario,
      nome: contexto.paciente.nome,
      cpf: contexto.paciente.cpf,
      status:
        contexto.paciente.status,
      sessoesPrescritas:
        contexto.paciente
          .sessoesPrescritas,
      sessoesRealizadas:
        contexto.paciente
          .sessoesRealizadas
    },

    ciclo: {
      id: contexto.idCiclo,
      numero: contexto.cicloNumero,
      quantidadeSessoes:
        contexto.sessoes.length,
      primeiraData:
        formatarDataCorrecaoAgendamento_(
          primeiraSessao.data
        ),
      ultimaData:
        formatarDataCorrecaoAgendamento_(
          ultimaSessao.data
        ),
      horario:
        formatarHorarioCorrecaoAgendamento_(
          primeiraSessao.horario
        ),
      fisioterapeuta:
        primeiraSessao.fisioterapeuta,
      tipoGrupo:
        primeiraSessao.tipoGrupo
    },

    sessoes: contexto.sessoes.map(
      function(sessao) {
        return {
          idAgendamento:
            sessao.idAgendamento,
          numero:
            sessao.numeroSessao,
          data:
            formatarDataCorrecaoAgendamento_(
              sessao.data
            ),
          dia: sessao.dia,
          horario:
            formatarHorarioCorrecaoAgendamento_(
              sessao.horario
            ),
          status: sessao.status
        };
      }
    ),

    podeCorrigir: !motivoBloqueio,
    motivoBloqueio:
      motivoBloqueio
  };
}


/**
 * Impede a remoção de ciclos que já começaram
 * ou possuem algum resultado registrado.
 */
function obterMotivoBloqueioCorrecao_(
  contexto
) {
  if (
    contexto.paciente
      .sessoesRealizadas > 0
  ) {
    return (
      'O ciclo já possui sessões contabilizadas. ' +
      'Use futuramente a função de ajuste de ciclo em andamento.'
    );
  }

  const sessaoNaoAgendada =
    contexto.sessoes.find(
      function(sessao) {
        return (
          normalizarTextoCorrecaoAgendamento_(
            sessao.status
          ) !== 'agendado'
        );
      }
    );

  if (sessaoNaoAgendada) {
    return (
      'A sessão ' +
      sessaoNaoAgendada.idAgendamento +
      ' possui o status "' +
      sessaoNaoAgendada.status +
      '". Somente ciclos inteiramente agendados podem ser desfeitos por esta função.'
    );
  }

  return '';
}


function validarCicloParaCorrecaoAgendamento_(
  contexto
) {
  const motivo =
    obterMotivoBloqueioCorrecao_(
      contexto
    );

  if (motivo) {
    throw new Error(motivo);
  }

  if (contexto.sessoes.length === 0) {
    throw new Error(
      'Nenhuma sessão foi encontrada para corrigir.'
    );
  }
}


/**
 * Cria a aba técnica de auditoria quando necessário.
 */
function obterOuCriarHistoricoAjustes_(
  ss
) {
  const nomeAba =
    CONFIG_CORRIGIR_AGENDAMENTO
      .ABAS.HISTORICO_AJUSTES;

  let aba =
    ss.getSheetByName(nomeAba);

  if (!aba) {
    aba = ss.insertSheet(nomeAba);
  }

  const cabecalho =
    CONFIG_CORRIGIR_AGENDAMENTO
      .CABECALHO_HISTORICO;

  aba.getRange(
    1,
    1,
    1,
    cabecalho.length
  )
    .setValues([cabecalho])
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true)
    .setBackground('#d9ead3');

  aba.setFrozenRows(1);
  aba.setRowHeight(1, 42);

  const possuiProtecao =
    aba.getProtections(
      SpreadsheetApp.ProtectionType.SHEET
    ).some(function(protecao) {
      return (
        protecao.getDescription() ===
        'SIGAF - Histórico de Ajustes'
      );
    });

  if (!possuiProtecao) {
    aba.protect()
      .setDescription(
        'SIGAF - Histórico de Ajustes'
      )
      .setWarningOnly(true);
  }

  return aba;
}


/**
 * Arquiva uma cópia das sessões antes da remoção.
 */
function registrarHistoricoCorrecaoAgendamento_(
  abaHistorico,
  contexto,
  idAjuste,
  agora
) {
  const linhas =
    contexto.sessoes.map(
      function(sessao) {
        return [
          idAjuste,
          agora,
          contexto.paciente.id,
          contexto.paciente.prontuario,
          contexto.paciente.nome,
          contexto.idCiclo,
          contexto.cicloNumero,
          sessao.idAgendamento,
          new Date(sessao.data),
          sessao.dia,
          sessao.horario,
          sessao.fisioterapeuta,
          sessao.tipoGrupo,
          sessao.numeroSessao,
          sessao.totalPrescrito,
          sessao.status,
          'Correção de agendamento antes do início do ciclo'
        ];
      }
    );

  const linhaDestino = Math.max(
    abaHistorico.getLastRow() + 1,
    2
  );

  abaHistorico.getRange(
    linhaDestino,
    1,
    linhas.length,
    CONFIG_CORRIGIR_AGENDAMENTO
      .CABECALHO_HISTORICO.length
  ).setValues(linhas);

  abaHistorico.getRange(
    linhaDestino,
    2,
    linhas.length,
    1
  ).setNumberFormat(
    'dd/MM/yyyy HH:mm'
  );

  abaHistorico.getRange(
    linhaDestino,
    9,
    linhas.length,
    1
  ).setNumberFormat(
    'dd/MM/yyyy'
  );

  abaHistorico.getRange(
    linhaDestino,
    11,
    linhas.length,
    1
  ).setNumberFormat(
    'HH:mm'
  );
}


/**
 * Limpa somente o conteúdo das linhas de sessão.
 * A formatação e as validações permanecem.
 */
function limparSessoesCorrecaoAgendamento_(
  abaAgendamentos,
  sessoes
) {
  const intervalos =
    sessoes.map(function(sessao) {
      return (
        'A' +
        sessao.linhaPlanilha +
        ':V' +
        sessao.linhaPlanilha
      );
    });

  abaAgendamentos
    .getRangeList(intervalos)
    .clearContent();
}


/**
 * Retorna o cadastro ao estado anterior ao
 * agendamento, preservando o plano terapêutico.
 */
function prepararPacienteParaReagendamento_(
  contexto
) {
  const aba =
    contexto.abaCadastro;

  const linha =
    contexto.paciente.linha;

  const colunas =
    CONFIG_CORRIGIR_AGENDAMENTO
      .CADASTRO;

  aba.getRange(
    linha,
    colunas.DATA_INICIO,
    1,
    2
  ).clearContent();

  aba.getRange(
    linha,
    colunas.SESSOES_REALIZADAS
  ).setValue(0);

  aba.getRange(
    linha,
    colunas.SESSOES_RESTANTES
  ).setValue(
    contexto.paciente
      .sessoesPrescritas
  );

  aba.getRange(
    linha,
    colunas.STATUS
  ).setValue(
    'Avaliado – aguardando agendamento'
  );
}


/**
 * Remove somente as pendências automáticas ligadas
 * às sessões e ao ciclo que acabaram de ser arquivados.
 */
function removerPendenciasVinculadasCorrecao_(
  ss,
  contexto
) {
  const abaPendencias =
    ss.getSheetByName(
      CONFIG_CORRIGIR_AGENDAMENTO
        .ABAS.PENDENCIAS
    );

  if (
    !abaPendencias ||
    abaPendencias.getLastRow() < 2
  ) {
    return;
  }

  const ultimaLinha =
    abaPendencias.getLastRow();

  const quantidadeColunas =
    CONFIG_CORRIGIR_AGENDAMENTO
      .QUANTIDADE_COLUNAS_PENDENCIAS;

  const dados = abaPendencias
    .getRange(
      2,
      1,
      ultimaLinha - 1,
      quantidadeColunas
    )
    .getValues();

  const idCiclo =
    normalizarTextoCorrecaoAgendamento_(
      contexto.idCiclo
    );

  const idsAgendamentos =
    contexto.sessoes.map(
      function(sessao) {
        return normalizarTextoCorrecaoAgendamento_(
          sessao.idAgendamento
        );
      }
    );

  const preservadas =
    dados.filter(function(linha) {
      const idPendencia =
        normalizarTextoCorrecaoAgendamento_(
          linha[0]
        );

      if (!idPendencia) {
        return false;
      }

      if (
        idCiclo &&
        idPendencia.indexOf(
          idCiclo
        ) !== -1
      ) {
        return false;
      }

      const vinculadaAgendamento =
        idsAgendamentos.some(
          function(idAgendamento) {
            return (
              idAgendamento &&
              idPendencia.indexOf(
                idAgendamento
              ) !== -1
            );
          }
        );

      return !vinculadaAgendamento;
    });

  regravarPendenciasCorrecao_(
    abaPendencias,
    preservadas
  );
}
/**
 * Limpa pendências transitórias antigas que não
 * correspondem mais ao estado atual da agenda.
 *
 * Execute manualmente uma vez para corrigir resíduos
 * de ciclos desfeitos antes desta atualização.
 */
function limparPendenciasObsoletasAposCorrecao() {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const ui =
    SpreadsheetApp.getUi();

  const lock =
    LockService.getDocumentLock();

  let bloqueioObtido = false;

  try {
    lock.waitLock(30000);
    bloqueioObtido = true;

    const abaPendencias =
      ss.getSheetByName(
        CONFIG_CORRIGIR_AGENDAMENTO
          .ABAS.PENDENCIAS
      );

    const abaCadastro =
      ss.getSheetByName(
        CONFIG_CORRIGIR_AGENDAMENTO
          .ABAS.CADASTRO
      );

    const abaAgendamentos =
      ss.getSheetByName(
        CONFIG_CORRIGIR_AGENDAMENTO
          .ABAS.AGENDAMENTOS
      );

    if (
      !abaPendencias ||
      !abaCadastro ||
      !abaAgendamentos
    ) {
      throw new Error(
        'As abas necessárias para a limpeza não foram encontradas.'
      );
    }

    if (abaPendencias.getLastRow() < 2) {
      ui.alert(
        'Nenhuma pendência precisa ser limpa.',
        ui.ButtonSet.OK
      );
      return;
    }

    const pacientes =
      mapearPacientesLimpezaCorrecao_(
        abaCadastro
      );

    const agendamentos =
      mapearAgendamentosLimpezaCorrecao_(
        abaAgendamentos
      );

    const ultimaLinha =
      abaPendencias.getLastRow();

    const dados = abaPendencias
      .getRange(
        2,
        1,
        ultimaLinha - 1,
        CONFIG_CORRIGIR_AGENDAMENTO
          .QUANTIDADE_COLUNAS_PENDENCIAS
      )
      .getValues();

    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);

    let removidas = 0;

    const preservadas =
      dados.filter(function(linha) {
        const obsoleta =
          pendenciaTransitoriaObsoletaCorrecao_(
            linha,
            pacientes,
            agendamentos,
            hoje
          );

        if (obsoleta) {
          removidas++;
          return false;
        }

        return true;
      });

    regravarPendenciasCorrecao_(
      abaPendencias,
      preservadas
    );

    SpreadsheetApp.flush();

    ui.alert(
      'Limpeza concluída',
      removidas > 0
        ? (
          'Pendências antigas removidas: ' +
          removidas +
          '. As pendências válidas foram preservadas.'
        )
        : 'Nenhuma pendência obsoleta foi encontrada.',
      ui.ButtonSet.OK
    );
  } catch (erro) {
    ui.alert(
      'Erro ao limpar pendências',
      erro && erro.message
        ? erro.message
        : String(erro),
      ui.ButtonSet.OK
    );

    throw erro;
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


function mapearPacientesLimpezaCorrecao_(
  abaCadastro
) {
  const resultado = {};
  const ultimaLinha =
    abaCadastro.getLastRow();

  if (ultimaLinha < 2) {
    return resultado;
  }

  const dados = abaCadastro
    .getRange(
      2,
      1,
      ultimaLinha - 1,
      CONFIG_CORRIGIR_AGENDAMENTO
        .QUANTIDADE_COLUNAS_CADASTRO
    )
    .getValues();

  dados.forEach(function(linha) {
    const idPaciente = String(
      linha[
        CONFIG_CORRIGIR_AGENDAMENTO
          .CADASTRO.ID_PACIENTE - 1
      ] || ''
    ).trim();

    const prontuario = String(
      linha[
        CONFIG_CORRIGIR_AGENDAMENTO
          .CADASTRO.PRONTUARIO - 1
      ] || ''
    ).trim();

    if (!idPaciente) {
      return;
    }

    const paciente = {
      id: idPaciente,
      prontuario: prontuario,
      status: String(
        linha[
          CONFIG_CORRIGIR_AGENDAMENTO
            .CADASTRO.STATUS - 1
        ] || ''
      ).trim(),
      sessoesRestantes:
        Number(
          linha[
            CONFIG_CORRIGIR_AGENDAMENTO
              .CADASTRO
              .SESSOES_RESTANTES - 1
          ]
        ) || 0,
      dataTermino:
        linha[
          CONFIG_CORRIGIR_AGENDAMENTO
            .CADASTRO.DATA_TERMINO - 1
        ]
    };

    resultado[
      'ID:' +
      normalizarTextoCorrecaoAgendamento_(
        idPaciente
      )
    ] = paciente;

    if (prontuario) {
      resultado[
        'PRONTUARIO:' +
        normalizarTextoCorrecaoAgendamento_(
          prontuario
        )
      ] = paciente;
    }
  });

  return resultado;
}


function mapearAgendamentosLimpezaCorrecao_(
  abaAgendamentos
) {
  const porId = {};
  const porPaciente = {};

  const ultimaLinha =
    abaAgendamentos.getLastRow();

  if (ultimaLinha < 2) {
    return {
      porId: porId,
      porPaciente: porPaciente
    };
  }

  const dados = abaAgendamentos
    .getRange(
      2,
      1,
      ultimaLinha - 1,
      CONFIG_CORRIGIR_AGENDAMENTO
        .QUANTIDADE_COLUNAS_AGENDAMENTOS
    )
    .getValues();

  dados.forEach(function(linha) {
    const idAgendamento = String(
      linha[
        CONFIG_CORRIGIR_AGENDAMENTO
          .AGENDAMENTOS.ID_AGENDAMENTO -
          1
      ] || ''
    ).trim();

    const idPaciente = String(
      linha[
        CONFIG_CORRIGIR_AGENDAMENTO
          .AGENDAMENTOS.ID_PACIENTE - 1
      ] || ''
    ).trim();

    if (!idAgendamento || !idPaciente) {
      return;
    }

    const registro = {
      idAgendamento: idAgendamento,
      idPaciente: idPaciente,
      evento: String(
        linha[
          CONFIG_CORRIGIR_AGENDAMENTO
            .AGENDAMENTOS.EVENTO - 1
        ] || ''
      ).trim(),
      data:
        linha[
          CONFIG_CORRIGIR_AGENDAMENTO
            .AGENDAMENTOS.DATA - 1
        ],
      status: String(
        linha[
          CONFIG_CORRIGIR_AGENDAMENTO
            .AGENDAMENTOS.STATUS - 1
        ] || ''
      ).trim()
    };

    porId[
      normalizarTextoCorrecaoAgendamento_(
        idAgendamento
      )
    ] = registro;

    const chavePaciente =
      normalizarTextoCorrecaoAgendamento_(
        idPaciente
      );

    if (!porPaciente[chavePaciente]) {
      porPaciente[chavePaciente] = [];
    }

    porPaciente[chavePaciente].push(
      registro
    );
  });

  return {
    porId: porId,
    porPaciente: porPaciente
  };
}


function pendenciaTransitoriaObsoletaCorrecao_(
  linha,
  pacientes,
  agendamentos,
  hoje
) {
  const idPendencia =
    normalizarTextoCorrecaoAgendamento_(
      linha[0]
    );

  const prontuario =
    normalizarTextoCorrecaoAgendamento_(
      linha[4]
    );

  if (!idPendencia) {
    return true;
  }

  const paciente =
    pacientes[
      'PRONTUARIO:' + prontuario
    ] || null;

  if (
    idPendencia.indexOf(
      'pend-resultado-'
    ) === 0
  ) {
    const ids =
      Object.keys(
        agendamentos.porId
      );

    const idEncontrado =
      ids.find(function(id) {
        return (
          idPendencia.indexOf(id) !== -1
        );
      });

    if (!idEncontrado) {
      return true;
    }

    const atendimento =
      agendamentos.porId[
        idEncontrado
      ];

    if (
      !(
        atendimento.data instanceof Date
      )
    ) {
      return true;
    }

    const dataAtendimento =
      new Date(atendimento.data);

    dataAtendimento.setHours(
      0,
      0,
      0,
      0
    );

    return !(
      normalizarTextoCorrecaoAgendamento_(
        atendimento.status
      ) === 'agendado' &&
      dataAtendimento.getTime() <
        hoje.getTime()
    );
  }

  if (
    idPendencia.indexOf(
      'pend-sem-agenda-'
    ) === 0
  ) {
    if (
      !paciente ||
      paciente.sessoesRestantes <= 0
    ) {
      return true;
    }

    const registros =
      agendamentos.porPaciente[
        normalizarTextoCorrecaoAgendamento_(
          paciente.id
        )
      ] || [];

    const possuiSessaoFutura =
      registros.some(function(registro) {
        if (
          normalizarTextoCorrecaoAgendamento_(
            registro.evento
          ) !== 'sessao' ||
          !(
            registro.data instanceof Date
          )
        ) {
          return false;
        }

        const data =
          new Date(registro.data);

        data.setHours(0, 0, 0, 0);

        const status =
          normalizarTextoCorrecaoAgendamento_(
            registro.status
          );

        return (
          data.getTime() >=
            hoje.getTime() &&
          [
            'agendado',
            'compareceu',
            'falta justificada',
            'falta nao justificada'
          ].indexOf(status) !== -1
        );
      });

    return possuiSessaoFutura;
  }

  if (
    idPendencia.indexOf(
      'pend-documentacao-'
    ) === 0
  ) {
    if (
      !paciente ||
      normalizarTextoCorrecaoAgendamento_(
        paciente.status
      ) !== 'em tratamento' ||
      !(
        paciente.dataTermino
        instanceof Date
      )
    ) {
      return true;
    }

    const termino =
      new Date(
        paciente.dataTermino
      );

    termino.setHours(0, 0, 0, 0);

    const diferencaDias =
      Math.round(
        (
          termino.getTime() -
          hoje.getTime()
        ) /
        86400000
      );

    return diferencaDias > 14;
  }

  return false;
}


function regravarPendenciasCorrecao_(
  abaPendencias,
  linhas
) {
  const ultimaLinha =
    abaPendencias.getLastRow();

  const quantidadeColunas =
    CONFIG_CORRIGIR_AGENDAMENTO
      .QUANTIDADE_COLUNAS_PENDENCIAS;

  if (ultimaLinha >= 2) {
    abaPendencias.getRange(
      2,
      1,
      ultimaLinha - 1,
      quantidadeColunas
    ).clearContent();
  }

  if (linhas.length === 0) {
    return;
  }

  abaPendencias.getRange(
    2,
    1,
    linhas.length,
    quantidadeColunas
  ).setValues(linhas);

  abaPendencias.getRange(
    2,
    8,
    linhas.length,
    2
  ).setNumberFormat(
    'dd/MM/yyyy'
  );

  abaPendencias.getRange(
    2,
    11,
    linhas.length,
    1
  ).setNumberFormat(
    'dd/MM/yyyy'
  );
}


function gerarIdAjusteCorrecaoAgendamento_(
  data
) {
  const base =
    Utilities.formatDate(
      data,
      Session.getScriptTimeZone(),
      'yyyyMMdd-HHmmss'
    );

  const aleatorio =
    String(
      Math.floor(
        1000 + Math.random() * 9000
      )
    );

  return (
    'AJ-' +
    base +
    '-' +
    aleatorio
  );
}


function formatarDataCorrecaoAgendamento_(
  data
) {
  if (!(data instanceof Date)) {
    return '';
  }

  return Utilities.formatDate(
    data,
    Session.getScriptTimeZone(),
    'dd/MM/yyyy'
  );
}


function formatarHorarioCorrecaoAgendamento_(
  valor
) {
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

  const resultado =
    texto.match(
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


function somenteNumerosCorrecaoAgendamento_(
  valor
) {
  return String(
    valor || ''
  ).replace(/\D/g, '');
}


function normalizarTextoCorrecaoAgendamento_(
  valor
) {
  return String(valor || '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(
      /[\u0300-\u036f]/g,
      ''
    );
}
