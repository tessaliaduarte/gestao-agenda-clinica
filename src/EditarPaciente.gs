const CONFIG_AJUSTE_CICLO_EDICAO = {
  ABA_AGENDAMENTOS: 'Agendamentos',
  QUANTIDADE_COLUNAS: 22,

  COLUNAS: {
    ID_AGENDAMENTO: 1,
    ID_PACIENTE: 2,
    ID_CICLO: 5,
    CICLO_NUMERO: 6,
    DATA: 7,
    EVENTO: 12,
    NUMERO_SESSAO: 13,
    TOTAL_PRESCRITO: 14,
    TIPO_GRUPO: 11,
    LIMITE_GRUPO: 15,
    STATUS: 16,
    MOTIVO: 17,
    CONTA_COMO_SESSAO: 18,
    AVISAR_PACIENTE: 19,
    ATUALIZADO_EM: 21,
    FATURAVEL: 22
  },

  TIPOS_ATENDIMENTO: {
    'atendimento com maior supervisao':
      'Atendimento com maior supervisão',

    'grupo de mmss':
      'Grupo de MMSS',

    'grupo de mmii':
      'Grupo de MMII',

    'grupo de coluna':
      'Grupo de Coluna'
  },

  EVENTO_ENCERRADO:
    'Sessão encerrada por ajuste',

  STATUS_ENCERRADO:
    'Encerrado por ajuste'
};


const CONFIG_REMANEJAMENTO_EDICAO = {
  ABA_STATUS_SESSAO:
    'Status da Sessão',

  STATUS_REMANEJADO:
    'Cancelado por remanejamento',

  COLUNAS_AGENDAMENTOS: {
    ID_AGENDAMENTO: 1,
    ID_PACIENTE: 2,
    PRONTUARIO: 3,
    PACIENTE: 4,
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
  }
};


/**
 * Abre o formulário para localizar e editar um paciente.
 */
function abrirFormularioEditarPaciente() {
  const html =
    HtmlService
      .createHtmlOutputFromFile(
        'FormularioEditarPaciente'
      )
      .setWidth(850)
      .setHeight(720);

  SpreadsheetApp
    .getUi()
    .showModalDialog(
      html,
      'SIGAF — Editar paciente'
    );
}


/**
 * Busca pacientes por ID, prontuário, CPF ou nome.
 */
function buscarPacientesParaEdicao(
  termoBusca
) {
  const termoOriginal =
    String(
      termoBusca || ''
    ).trim();

  if (!termoOriginal) {
    throw new Error(
      'Informe o ID, prontuário, CPF ou nome do paciente.'
    );
  }

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const aba =
    ss.getSheetByName(
      CONFIG_MODULO_PACIENTES
        .ABA_CADASTRO
    );

  if (!aba) {
    throw new Error(
      'A aba "Cadastro de Pacientes" não foi encontrada.'
    );
  }

  const ultimaLinha =
    aba.getLastRow();

  if (ultimaLinha < 2) {
    return [];
  }

  const dados =
    aba
      .getRange(
        2,
        1,
        ultimaLinha - 1,
        24
      )
      .getValues();

  const termoNormalizado =
    normalizarTextoPaciente_(
      termoOriginal
    );

  const termoNumerico =
    termoOriginal.replace(
      /\D/g,
      ''
    );

  const resultados = [];

  dados.forEach(
    function(linha, indice) {
      const id =
        String(
          linha[
            CONFIG_MODULO_PACIENTES
              .COLUNAS.ID - 1
          ] || ''
        ).trim();

      if (!id) {
        return;
      }

      const prontuario =
        String(
          linha[
            CONFIG_MODULO_PACIENTES
              .COLUNAS.PRONTUARIO - 1
          ] || ''
        ).trim();

      const nome =
        String(
          linha[
            CONFIG_MODULO_PACIENTES
              .COLUNAS.NOME - 1
          ] || ''
        ).trim();

      const cpf =
        String(
          linha[
            CONFIG_MODULO_PACIENTES
              .COLUNAS.CPF - 1
          ] || ''
        ).trim();

      const cpfNumerico =
        cpf.replace(
          /\D/g,
          ''
        );

      const corresponde =
        normalizarTextoPaciente_(
          id
        ) === termoNormalizado ||

        prontuario ===
          termoOriginal ||

        (
          termoNumerico &&
          prontuario ===
            termoNumerico
        ) ||

        (
          termoNumerico &&
          cpfNumerico ===
            termoNumerico
        ) ||

        normalizarTextoPaciente_(
          nome
        ).indexOf(
          termoNormalizado
        ) !== -1;

      if (!corresponde) {
        return;
      }

      resultados.push(
        montarResumoPacienteEdicao_(
          linha,
          indice + 2
        )
      );
    }
  );

  resultados.sort(
    function(a, b) {
      return a.nome.localeCompare(
        b.nome,
        'pt-BR'
      );
    }
  );

  return resultados.slice(
    0,
    30
  );
}


/**
 * Carrega os dados completos de um paciente pelo ID.
 */
function obterPacienteParaEdicao(
  idPaciente
) {
  const idProcurado =
    String(
      idPaciente || ''
    ).trim();

  if (!idProcurado) {
    throw new Error(
      'O ID do paciente não foi informado.'
    );
  }

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const aba =
    ss.getSheetByName(
      CONFIG_MODULO_PACIENTES
        .ABA_CADASTRO
    );

  if (!aba) {
    throw new Error(
      'A aba "Cadastro de Pacientes" não foi encontrada.'
    );
  }

  const localizacao =
    localizarLinhaPacientePorId_(
      aba,
      idProcurado
    );

  if (!localizacao) {
    throw new Error(
      'Paciente não encontrado.'
    );
  }

  const linha =
    aba
      .getRange(
        localizacao.linha,
        1,
        1,
        24
      )
      .getValues()[0];

  const paciente =
    montarDadosCompletosPacienteEdicao_(
      linha
    );

  paciente.remanejamento =
    obterResumoRemanejamentoEdicao_(
      ss,
      paciente
    );

  return paciente;
}
/**
 * Atualiza o cadastro e, quando necessário,
 * reduz o ciclo já existente.
 */
function atualizarPaciente(
  dados
) {
  const lock =
    LockService.getDocumentLock();

  let bloqueioObtido = false;

  try {
    lock.waitLock(30000);
    bloqueioObtido = true;

    if (!dados) {
      throw new Error(
        'Nenhuma informação foi recebida.'
      );
    }

    const idPaciente =
      String(
        dados.id || ''
      ).trim();

    if (!idPaciente) {
      throw new Error(
        'O ID do paciente não foi informado.'
      );
    }

    const ss =
      SpreadsheetApp
        .getActiveSpreadsheet();

    const aba =
      ss.getSheetByName(
        CONFIG_MODULO_PACIENTES
          .ABA_CADASTRO
      );

    if (!aba) {
      throw new Error(
        'A aba "Cadastro de Pacientes" não foi encontrada.'
      );
    }

    const localizacao =
      localizarLinhaPacientePorId_(
        aba,
        idPaciente
      );

    if (!localizacao) {
      throw new Error(
        'Paciente não encontrado.'
      );
    }

    const pacienteValidado =
      validarDadosEdicaoPaciente_(
        dados
      );

    verificarCpfDuplicadoNaEdicao_(
      aba,
      pacienteValidado.cpf,
      idPaciente
    );

    const linhaAtual =
      aba
        .getRange(
          localizacao.linha,
          1,
          1,
          24
        )
        .getValues()[0];

    const prontuarioAtual =
      linhaAtual[
        CONFIG_MODULO_PACIENTES
          .COLUNAS.PRONTUARIO - 1
      ];

    const sessoesRealizadas =
      converterNumeroInteiroEdicao_(
        linhaAtual[
          CONFIG_MODULO_PACIENTES
            .COLUNAS.SESSOES_REALIZADAS - 1
        ],
        0
      );

    const sessoesPrescritasAnteriores =
      converterNumeroInteiroEdicao_(
        linhaAtual[
          CONFIG_MODULO_PACIENTES
            .COLUNAS.SESSOES_PRESCRITAS - 1
        ],
        0
      );

    const mudancaAgenda =
      detectarMudancaAgendaEdicao_(
        linhaAtual,
        pacienteValidado
      );

    const statusAntesEdicao =
      normalizarTextoPaciente_(
        linhaAtual[
          CONFIG_MODULO_PACIENTES
            .COLUNAS.STATUS - 1
        ]
      );

    let planoRemanejamento =
      null;

    if (
      statusAntesEdicao ===
        'em tratamento' &&
      mudancaAgenda.alterouAgenda
    ) {
      const quantidadeFuturaDesejada =
        Math.max(
          pacienteValidado
            .sessoesPrescritas -
            sessoesRealizadas,
          0
        );

      planoRemanejamento =
        prepararPlanoRemanejamentoEdicao_(
          ss,
          idPaciente,
          pacienteValidado,
          quantidadeFuturaDesejada
        );
    }

    if (
      pacienteValidado
        .sessoesPrescritas <
      sessoesRealizadas
    ) {
      throw new Error(
        'As sessões prescritas não podem ser menores que as sessões já realizadas.'
      );
    }

    const sessoesRestantes =
      Math.max(
        pacienteValidado
          .sessoesPrescritas -
          sessoesRealizadas,
        0
      );

    const limiteGrupo =
      obterLimiteGrupoPaciente_(
        pacienteValidado
          .tipoAtendimento
      );

    /*
     * Ajusta a quantidade do ciclo,
     * somente quando ela foi reduzida.
     */
    const ajusteCiclo =
      ajustarQuantidadeSessoesCicloEdicao_(
        ss,
        idPaciente,
        sessoesPrescritasAnteriores,
        pacienteValidado
          .sessoesPrescritas,
        sessoesRealizadas
      );

    let quantidadeSessoesReclassificadas =
      0;

    let resultadoRemanejamento =
      null;

    /*
     * Durante um ciclo ativo, mudanças de dias, horário,
     * fisioterapeuta ou tipo de atendimento passam pelo
     * remanejamento seguro. Sessões realizadas permanecem
     * intactas; somente futuras Agendadas são substituídas.
     */
    if (planoRemanejamento) {
      resultadoRemanejamento =
        executarRemanejamentoEdicao_(
          ss,
          planoRemanejamento,
          pacienteValidado,
          limiteGrupo
        );
    } else {
      /*
       * Fora do fluxo de remanejamento, mantém a lógica
       * já existente de reclassificação do ciclo.
       */
      quantidadeSessoesReclassificadas =
        atualizarTipoAtendimentoCicloEdicao_(
          ss,
          idPaciente,
          pacienteValidado
            .tipoAtendimento,
          limiteGrupo
        );
    }

    const statusAtual =
      String(
        linhaAtual[
          CONFIG_MODULO_PACIENTES
            .COLUNAS.STATUS - 1
        ] || ''
      ).trim();

    const statusAtualizado =
      definirStatusAposEdicaoPaciente_(
        statusAtual,
        pacienteValidado
          .sessoesPrescritas,
        sessoesRealizadas,
        ajusteCiclo
          .possuiCicloAtivo
      );

    const registroAtualizado = [[
      idPaciente,
      prontuarioAtual,
      pacienteValidado.nome,
      pacienteValidado.cpfFormatado,
      pacienteValidado.telefone,
      pacienteValidado
        .horarioSessao || '',
      pacienteValidado
        .tipoAtendimento || '',
      limiteGrupo,

      pacienteValidado
        .dias.segunda,
      pacienteValidado
        .dias.terca,
      pacienteValidado
        .dias.quarta,
      pacienteValidado
        .dias.quinta,
      pacienteValidado
        .dias.sexta,

      pacienteValidado
        .sessoesPrescritas,
      sessoesRealizadas,
      sessoesRestantes,

      pacienteValidado
        .dataAvaliacao,
      pacienteValidado
        .horarioAvaliacao,

      linhaAtual[
        CONFIG_MODULO_PACIENTES
          .COLUNAS.DATA_INICIO - 1
      ],

      resultadoRemanejamento
        ? resultadoRemanejamento
            .dataTermino
        : (
            ajusteCiclo.quantidadeAlterada
              ? ajusteCiclo.dataTermino
              : linhaAtual[
                  CONFIG_MODULO_PACIENTES
                    .COLUNAS.DATA_TERMINO - 1
                ]
          ),

      statusAtualizado,
      pacienteValidado
        .fisioterapeuta || '',
      pacienteValidado
        .observacao || '',

      linhaAtual[
        CONFIG_MODULO_PACIENTES
          .COLUNAS.DESFECHO - 1
      ]
    ]];

    aba
      .getRange(
        localizacao.linha,
        1,
        1,
        24
      )
      .setValues(
        registroAtualizado
      );

    formatarLinhaNovoPaciente_(
      aba,
      localizacao.linha
    );

    aba
      .getRange(
        localizacao.linha,
        CONFIG_MODULO_PACIENTES
          .COLUNAS.DATA_TERMINO
      )
      .setNumberFormat(
        'dd/MM/yyyy'
      );

    SpreadsheetApp.flush();

    if (
      (
        ajusteCiclo
          .quantidadeAlterada ||
        quantidadeSessoesReclassificadas > 0 ||
        Boolean(
          resultadoRemanejamento
        )
      ) &&
      typeof carregarAgendaDiaria ===
        'function'
    ) {
      try {
        carregarAgendaDiaria();
      } catch (erroAgenda) {
        console.error(
          erroAgenda
        );
      }
    }

    const mensagemAjuste =
      ajusteCiclo.quantidadeAlterada
        ? (
            '\n\nCiclo ajustado de ' +
            ajusteCiclo
              .quantidadeAnteriorCiclo +
            ' para ' +
            pacienteValidado
              .sessoesPrescritas +
            ' sessões.' +
            '\nSessões futuras encerradas: ' +
            ajusteCiclo
              .sessoesEncerradas +
            '.' +
            '\nNova previsão de término: ' +
            formatarDataMensagemEdicao_(
              ajusteCiclo.dataTermino
            ) +
            '.'
          )
        : '';

    const mensagemGrupo =
      quantidadeSessoesReclassificadas > 0
        ? (
            '\n\nTipo de atendimento atualizado no ciclo: ' +
            pacienteValidado
              .tipoAtendimento +
            '.' +
            '\nSessões atualizadas: ' +
            quantidadeSessoesReclassificadas +
            '.'
          )
        : '';

    const mensagemRemanejamento =
      resultadoRemanejamento
        ? (
            '\n\nRemanejamento realizado com segurança.' +
            '\nSessões futuras remanejadas: ' +
            resultadoRemanejamento
              .quantidadeRemanejada +
            '.' +
            '\nNovo término previsto: ' +
            formatarDataMensagemEdicao_(
              resultadoRemanejamento
                .dataTermino
            ) +
            '.'
          )
        : '';

    return {
      sucesso: true,
      id: idPaciente,
      prontuario:
        prontuarioAtual,
      nome:
        pacienteValidado.nome,
      status:
        statusAtualizado,

      mensagem:
        'Cadastro atualizado com sucesso.' +
        '\n\nID: ' +
        idPaciente +
        '\nProntuário: ' +
        prontuarioAtual +
        '\nPaciente: ' +
        pacienteValidado.nome +
        '\nStatus: ' +
        statusAtualizado +
        mensagemAjuste +
        mensagemGrupo +
        mensagemRemanejamento
    };
  } catch (erro) {
    return {
      sucesso: false,

      mensagem:
        erro && erro.message
          ? erro.message
          : String(erro)
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
}

/**
 * Verifica a disponibilidade de um remanejamento sem
 * alterar qualquer dado.
 *
 * Chamado pelo FormularioEditarPaciente quando um paciente
 * Em tratamento muda dias, horário, fisioterapeuta ou tipo.
 */
function verificarRemanejamentoEdicaoPaciente(
  dados
) {
  try {
    if (!dados) {
      throw new Error(
        'Nenhuma informação foi recebida.'
      );
    }

    const idPaciente =
      String(
        dados.id || ''
      ).trim();

    if (!idPaciente) {
      throw new Error(
        'O ID do paciente não foi informado.'
      );
    }

    const pacienteValidado =
      validarDadosEdicaoPaciente_(
        dados
      );

    const ss =
      SpreadsheetApp
        .getActiveSpreadsheet();

    const abaCadastro =
      ss.getSheetByName(
        CONFIG_MODULO_PACIENTES
          .ABA_CADASTRO
      );

    if (!abaCadastro) {
      throw new Error(
        'A aba "Cadastro de Pacientes" não foi encontrada.'
      );
    }

    const localizacao =
      localizarLinhaPacientePorId_(
        abaCadastro,
        idPaciente
      );

    if (!localizacao) {
      throw new Error(
        'Paciente não encontrado.'
      );
    }

    const linhaAtual =
      abaCadastro
        .getRange(
          localizacao.linha,
          1,
          1,
          24
        )
        .getValues()[0];

    const statusAtual =
      normalizarTextoPaciente_(
        linhaAtual[
          CONFIG_MODULO_PACIENTES
            .COLUNAS.STATUS - 1
        ]
      );

    if (
      statusAtual !==
        'em tratamento'
    ) {
      return {
        sucesso: true,
        requerRemanejamento: false,
        mensagem:
          'Este paciente ainda não exige remanejamento do ciclo ativo.'
      };
    }

    const mudanca =
      detectarMudancaAgendaEdicao_(
        linhaAtual,
        pacienteValidado
      );

    if (!mudanca.alterouAgenda) {
      return {
        sucesso: true,
        requerRemanejamento: false,
        mensagem:
          'Dias, horário, fisioterapeuta e tipo de atendimento não foram alterados.'
      };
    }

    const sessoesRealizadas =
      converterNumeroInteiroEdicao_(
        linhaAtual[
          CONFIG_MODULO_PACIENTES
            .COLUNAS.SESSOES_REALIZADAS - 1
        ],
        0
      );

    if (
      pacienteValidado
        .sessoesPrescritas <
      sessoesRealizadas
    ) {
      throw new Error(
        'As sessões prescritas não podem ser menores que as sessões já realizadas.'
      );
    }

    const quantidadeFutura =
      Math.max(
        pacienteValidado
          .sessoesPrescritas -
          sessoesRealizadas,
        0
      );

    const plano =
      prepararPlanoRemanejamentoEdicao_(
        ss,
        idPaciente,
        pacienteValidado,
        quantidadeFutura
      );

    return {
      sucesso: true,
      requerRemanejamento: true,
      quantidade:
        plano.sessoesSelecionadas
          .length,
      dataInicio:
        formatarDataMensagemEdicao_(
          plano.novasDatas[0]
        ),
      dataTermino:
        formatarDataMensagemEdicao_(
          plano.novasDatas[
            plano.novasDatas.length - 1
          ]
        ),
      dias:
        plano.dias
          .map(
            nomeDiaRemanejamentoEdicao_
          )
          .join(' / '),
      horario:
        plano.horario.exibicao,
      fisioterapeuta:
        plano.fisioterapeuta,
      tipoAtendimento:
        pacienteValidado
          .tipoAtendimento,
      mensagem:
        'Disponibilidade confirmada para todas as sessões futuras.'
    };
  } catch (erro) {
    return {
      sucesso: false,
      requerRemanejamento: true,
      mensagem:
        erro && erro.message
          ? erro.message
          : String(erro)
    };
  }
}


/**
 * Resume a situação do ciclo para o formulário.
 */
function obterResumoRemanejamentoEdicao_(
  ss,
  paciente
) {
  const status =
    normalizarTextoPaciente_(
      paciente.status
    );

  if (
    status !==
      'em tratamento'
  ) {
    return {
      emTratamento: false,
      sessoesFuturasAgendadas: 0
    };
  }

  const abaAgendamentos =
    ss.getSheetByName(
      CONFIG_AJUSTE_CICLO_EDICAO
        .ABA_AGENDAMENTOS
    );

  if (!abaAgendamentos) {
    return {
      emTratamento: true,
      sessoesFuturasAgendadas: 0
    };
  }

  const ciclo =
    obterCicloAtualRemanejamentoEdicao_(
      abaAgendamentos,
      paciente.id
    );

  if (!ciclo.id) {
    return {
      emTratamento: true,
      sessoesFuturasAgendadas: 0
    };
  }

  const futuras =
    obterSessoesFuturasRemanejamentoEdicao_(
      abaAgendamentos,
      paciente.id,
      ciclo.id
    );

  return {
    emTratamento: true,
    idCiclo: ciclo.id,
    cicloNumero: ciclo.numero,
    sessoesFuturasAgendadas:
      futuras.length
  };
}


/**
 * Identifica se houve mudança na programação do ciclo.
 */
function detectarMudancaAgendaEdicao_(
  linhaAtual,
  pacienteValidado
) {
  const c =
    CONFIG_MODULO_PACIENTES
      .COLUNAS;

  const horarioMudou =
    chaveHorarioRemanejamentoEdicao_(
      linhaAtual[
        c.HORARIO - 1
      ]
    ) !==
    chaveHorarioRemanejamentoEdicao_(
      pacienteValidado
        .horarioSessao
    );

  const fisioterapeutaMudou =
    normalizarTextoPaciente_(
      linhaAtual[
        c.FISIOTERAPEUTA - 1
      ]
    ) !==
    normalizarTextoPaciente_(
      pacienteValidado
        .fisioterapeuta
    );

  const tipoMudou =
    normalizarTextoPaciente_(
      linhaAtual[
        c.TIPO_ATENDIMENTO - 1
      ]
    ) !==
    normalizarTextoPaciente_(
      pacienteValidado
        .tipoAtendimento
    );

  const diasAtuais = [
    converterCheckboxEdicao_(
      linhaAtual[c.SEGUNDA - 1]
    ),
    converterCheckboxEdicao_(
      linhaAtual[c.TERCA - 1]
    ),
    converterCheckboxEdicao_(
      linhaAtual[c.QUARTA - 1]
    ),
    converterCheckboxEdicao_(
      linhaAtual[c.QUINTA - 1]
    ),
    converterCheckboxEdicao_(
      linhaAtual[c.SEXTA - 1]
    )
  ];

  const novosDias = [
    Boolean(
      pacienteValidado
        .dias.segunda
    ),
    Boolean(
      pacienteValidado
        .dias.terca
    ),
    Boolean(
      pacienteValidado
        .dias.quarta
    ),
    Boolean(
      pacienteValidado
        .dias.quinta
    ),
    Boolean(
      pacienteValidado
        .dias.sexta
    )
  ];

  const diasMudaram =
    diasAtuais.some(
      function(valor, indice) {
        return (
          Boolean(valor) !==
          Boolean(
            novosDias[indice]
          )
        );
      }
    );

  return {
    alterouAgenda:
      horarioMudou ||
      fisioterapeutaMudou ||
      tipoMudou ||
      diasMudaram,
    horarioMudou:
      horarioMudou,
    fisioterapeutaMudou:
      fisioterapeutaMudou,
    tipoMudou:
      tipoMudou,
    diasMudaram:
      diasMudaram
  };
}


/**
 * Prepara um remanejamento completo sem gravar nada.
 */
function prepararPlanoRemanejamentoEdicao_(
  ss,
  idPaciente,
  pacienteValidado,
  quantidadeFuturaDesejada
) {
  validarDependenciasRemanejamentoEdicao_();

  if (
    quantidadeFuturaDesejada < 1
  ) {
    throw new Error(
      'Não existem sessões futuras a remanejar.'
    );
  }

  const abaAgendamentos =
    ss.getSheetByName(
      CONFIG_AJUSTE_CICLO_EDICAO
        .ABA_AGENDAMENTOS
    );

  const abaHorarios =
    ss.getSheetByName(
      CONFIG_MODULO_PACIENTES
        .ABA_HORARIOS
    );

  const abaFisioterapeutas =
    ss.getSheetByName(
      CONFIG_MODULO_PACIENTES
        .ABA_FISIOTERAPEUTAS
    );

  const abaFeriados =
    ss.getSheetByName(
      'Calendário da Prefeitura'
    );

  const abaBloqueios =
    ss.getSheetByName(
      'Bloqueios'
    );

  if (
    !abaAgendamentos ||
    !abaHorarios ||
    !abaFisioterapeutas ||
    !abaFeriados ||
    !abaBloqueios
  ) {
    throw new Error(
      'Uma das abas necessárias ao remanejamento não foi encontrada.'
    );
  }

  const ciclo =
    obterCicloAtualRemanejamentoEdicao_(
      abaAgendamentos,
      idPaciente
    );

  if (!ciclo.id) {
    throw new Error(
      'Não foi encontrado um ciclo ativo para o paciente.'
    );
  }

  const futuras =
    obterSessoesFuturasRemanejamentoEdicao_(
      abaAgendamentos,
      idPaciente,
      ciclo.id
    );

  if (
    futuras.length <
      quantidadeFuturaDesejada
  ) {
    throw new Error(
      'O ciclo possui apenas ' +
      futuras.length +
      ' sessão(ões) futura(s) Agendada(s), mas o Cadastro indica ' +
      quantidadeFuturaDesejada +
      ' sessão(ões) ainda necessária(s). Regularize o ciclo antes de alterar a agenda.'
    );
  }

  const sessoesSelecionadas =
    futuras
      .filter(function(sessao) {
        return (
          sessao.numeroSessao <=
          pacienteValidado
            .sessoesPrescritas
        );
      })
      .slice(
        0,
        quantidadeFuturaDesejada
      );

  if (
    sessoesSelecionadas.length !==
      quantidadeFuturaDesejada
  ) {
    throw new Error(
      'Não foi possível identificar todas as sessões futuras que precisam ser remanejadas.'
    );
  }

  const dias =
    obterDiasNumericosEdicao_(
      pacienteValidado.dias
    );

  if (dias.length === 0) {
    throw new Error(
      'Selecione pelo menos um dia da semana.'
    );
  }

  const limiteGrupo =
    obterLimiteGrupoPaciente_(
      pacienteValidado
        .tipoAtendimento
    );

  const horarios =
    lerHorariosPermitidosVagasSessoes_(
      abaHorarios
    );

  const chaveHorario =
    chaveHorarioVagasSessoes_(
      pacienteValidado
        .horarioSessao
    );

  const horario =
    horarios.find(
      function(item) {
        return (
          item.chave ===
            chaveHorario
        );
      }
    );

  if (!horario) {
    throw new Error(
      'O horário selecionado não está disponível para sessões.'
    );
  }

  const fisioterapeutasPorTurno =
    lerFisioterapeutasPorTurnoVagasSessoes_(
      abaFisioterapeutas
    );

  const turno =
    normalizarTextoVagasSessoes_(
      horario.turno
    );

  const profissionais =
    fisioterapeutasPorTurno[
      turno
    ] || [];

  const fisioterapeuta =
    String(
      pacienteValidado
        .fisioterapeuta || ''
    ).trim();

  const profissionalValido =
    profissionais.some(
      function(nome) {
        return (
          normalizarTextoPaciente_(
            nome
          ) ===
          normalizarTextoPaciente_(
            fisioterapeuta
          )
        );
      }
    );

  if (!profissionalValido) {
    throw new Error(
      'O fisioterapeuta selecionado não atende no turno correspondente ao horário escolhido.'
    );
  }

  const contexto = {
    paciente: {
      tipoGrupo:
        pacienteValidado
          .tipoAtendimento,
      capacidade:
        limiteGrupo,
      quantidadeParaAgendar:
        quantidadeFuturaDesejada
    },

    dataInicial:
      new Date(
        sessoesSelecionadas[0]
          .data
      ),

    horarios:
      horarios,

    fisioterapeutasPorTurno:
      fisioterapeutasPorTurno,

    feriados:
      lerFeriadosVagasSessoes_(
        abaFeriados
      ),

    bloqueios:
      lerBloqueiosVagasSessoes_(
        abaBloqueios
      ),

    ocupacoes:
      lerOcupacoesRemanejamentoEdicao_(
        abaAgendamentos,
        futuras.map(
          function(sessao) {
            return sessao
              .idAgendamento;
          }
        )
      )
  };

  const sequencia =
    simularSequenciaVagasSessoes_(
      contexto,
      dias,
      horario,
      fisioterapeuta
    );

  if (!sequencia.valida) {
    throw new Error(
      'Não existe uma sequência completa de vagas para todas as sessões futuras nesse padrão. Nenhuma alteração foi realizada.'
    );
  }

  validarDatasDuplicadasPacienteEdicao_(
    abaAgendamentos,
    idPaciente,
    futuras,
    sequencia.datas
  );

  return {
    ciclo: ciclo,
    sessoesSelecionadas:
      sessoesSelecionadas,
    todasFuturas:
      futuras,
    novasDatas:
      sequencia.datas,
    dias: dias,
    horario: horario,
    fisioterapeuta:
      fisioterapeuta,
    limiteGrupo:
      limiteGrupo
  };
}


/**
 * Executa o remanejamento depois de toda a validação.
 */
function executarRemanejamentoEdicao_(
  ss,
  plano,
  pacienteValidado,
  limiteGrupo
) {
  const abaAgendamentos =
    ss.getSheetByName(
      CONFIG_AJUSTE_CICLO_EDICAO
        .ABA_AGENDAMENTOS
    );

  if (!abaAgendamentos) {
    throw new Error(
      'A aba "Agendamentos" não foi encontrada.'
    );
  }

  garantirStatusRemanejamentoEdicao_(
    ss
  );

  /*
   * Após eventual redução de quantidade, relê apenas as
   * sessões que ainda permanecem Agendadas no ciclo.
   */
  const atuais =
    obterSessoesFuturasRemanejamentoEdicao_(
      abaAgendamentos,
      plano.sessoesSelecionadas[0]
        .idPaciente,
      plano.ciclo.id
    )
      .filter(function(sessao) {
        return (
          sessao.numeroSessao <=
            pacienteValidado
              .sessoesPrescritas
        );
      })
      .slice(
        0,
        plano.novasDatas.length
      );

  if (
    atuais.length !==
      plano.novasDatas.length
  ) {
    throw new Error(
      'O ciclo foi alterado enquanto o remanejamento era processado. Tente novamente.'
    );
  }

  const novosIds =
    gerarIdsAgendamentoRemanejamentoEdicao_(
      abaAgendamentos,
      atuais.length
    );

  const agora =
    new Date();

  const a =
    CONFIG_REMANEJAMENTO_EDICAO
      .COLUNAS_AGENDAMENTOS;

  const novasLinhas =
    atuais.map(
      function(sessao, indice) {
        const linha =
          sessao.valores.slice();

        linha[
          a.ID_AGENDAMENTO - 1
        ] =
          novosIds[indice];

        linha[
          a.DATA - 1
        ] =
          new Date(
            plano.novasDatas[
              indice
            ]
          );

        linha[
          a.DIA - 1
        ] =
          nomeDiaRemanejamentoEdicao_(
            plano.novasDatas[
              indice
            ].getDay()
          );

        linha[
          a.HORARIO - 1
        ] =
          plano.horario.valor;

        linha[
          a.FISIOTERAPEUTA - 1
        ] =
          plano.fisioterapeuta;

        linha[
          a.TIPO_GRUPO - 1
        ] =
          pacienteValidado
            .tipoAtendimento;

        linha[
          a.LIMITE_GRUPO - 1
        ] =
          limiteGrupo;

        linha[
          a.STATUS - 1
        ] =
          'Agendado';

        linha[
          a.MOTIVO - 1
        ] =
          'Remanejamento de ' +
          sessao.idAgendamento;

        linha[
          a.CONTA_COMO_SESSAO - 1
        ] =
          'Não';

        linha[
          a.AVISAR_PACIENTE - 1
        ] =
          'Não';

        linha[
          a.CRIADO_EM - 1
        ] =
          agora;

        linha[
          a.ATUALIZADO_EM - 1
        ] =
          agora;

        linha[
          a.FATURAVEL - 1
        ] =
          'Não';

        return linha;
      }
    );

  const primeiraNovaLinha =
    Math.max(
      abaAgendamentos.getLastRow() + 1,
      2
    );

  abaAgendamentos
    .getRange(
      primeiraNovaLinha,
      1,
      novasLinhas.length,
      22
    )
    .setValues(
      novasLinhas
    );

  try {
    abaAgendamentos
      .getRange(
        primeiraNovaLinha,
        a.DATA,
        novasLinhas.length,
        1
      )
      .setNumberFormat(
        'dd/MM/yyyy'
      );

    abaAgendamentos
      .getRange(
        primeiraNovaLinha,
        a.HORARIO,
        novasLinhas.length,
        1
      )
      .setNumberFormat(
        'HH:mm'
      );

    abaAgendamentos
      .getRange(
        primeiraNovaLinha,
        a.CRIADO_EM,
        novasLinhas.length,
        2
      )
      .setNumberFormat(
        'dd/MM/yyyy HH:mm'
      );

    cancelarSessoesAntigasRemanejamentoEdicao_(
      abaAgendamentos,
      atuais,
      novosIds,
      plano.novasDatas,
      agora
    );
  } catch (erro) {
    try {
      abaAgendamentos
        .getRange(
          primeiraNovaLinha,
          1,
          novasLinhas.length,
          22
        )
        .clearContent();
    } catch (erroLimpeza) {
      console.error(
        erroLimpeza
      );
    }

    throw erro;
  }

  return {
    quantidadeRemanejada:
      atuais.length,
    dataTermino:
      new Date(
        plano.novasDatas[
          plano.novasDatas.length - 1
        ]
      )
  };
}


/**
 * Cancela as sessões futuras antigas sem apagá-las.
 */
function cancelarSessoesAntigasRemanejamentoEdicao_(
  abaAgendamentos,
  sessoes,
  novosIds,
  novasDatas,
  agora
) {
  const a =
    CONFIG_REMANEJAMENTO_EDICAO
      .COLUNAS_AGENDAMENTOS;

  sessoes.forEach(
    function(sessao, indice) {
      definirValorTecnicoRemanejamentoEdicao_(
        abaAgendamentos,
        sessao.linha,
        a.STATUS,
        CONFIG_REMANEJAMENTO_EDICAO
          .STATUS_REMANEJADO
      );

      abaAgendamentos
        .getRange(
          sessao.linha,
          a.MOTIVO
        )
        .setValue(
          'Remanejado para ' +
          novosIds[indice] +
          ' em ' +
          formatarDataMensagemEdicao_(
            novasDatas[indice]
          )
        );

      abaAgendamentos
        .getRange(
          sessao.linha,
          a.CONTA_COMO_SESSAO
        )
        .setValue(
          'Não'
        );

      abaAgendamentos
        .getRange(
          sessao.linha,
          a.AVISAR_PACIENTE
        )
        .setValue(
          'Não'
        );

      abaAgendamentos
        .getRange(
          sessao.linha,
          a.ATUALIZADO_EM
        )
        .setValue(
          agora
        )
        .setNumberFormat(
          'dd/MM/yyyy HH:mm'
        );

      abaAgendamentos
        .getRange(
          sessao.linha,
          a.FATURAVEL
        )
        .setValue(
          'Não'
        );
    }
  );
}


/**
 * Identifica o ciclo mais recente do paciente.
 */
function obterCicloAtualRemanejamentoEdicao_(
  abaAgendamentos,
  idPaciente
) {
  const ultimaLinha =
    abaAgendamentos.getLastRow();

  if (ultimaLinha < 2) {
    return {
      id: '',
      numero: 0
    };
  }

  const dados =
    abaAgendamentos
      .getRange(
        2,
        1,
        ultimaLinha - 1,
        22
      )
      .getValues();

  const a =
    CONFIG_REMANEJAMENTO_EDICAO
      .COLUNAS_AGENDAMENTOS;

  const procurado =
    normalizarTextoPaciente_(
      idPaciente
    );

  let numero = 0;
  let id = '';

  dados.forEach(function(linha) {
    if (
      normalizarTextoPaciente_(
        linha[a.ID_PACIENTE - 1]
      ) !== procurado
    ) {
      return;
    }

    const numeroAtual =
      Number(
        linha[a.CICLO_NUMERO - 1]
      ) || 0;

    const idAtual =
      String(
        linha[a.ID_CICLO - 1] || ''
      ).trim();

    if (
      numeroAtual > numero ||
      (
        numeroAtual === numero &&
        idAtual &&
        !id
      )
    ) {
      numero =
        numeroAtual;
      id =
        idAtual;
    }
  });

  return {
    id: id,
    numero: numero
  };
}


/**
 * Lista somente sessões futuras ainda Agendadas.
 */
function obterSessoesFuturasRemanejamentoEdicao_(
  abaAgendamentos,
  idPaciente,
  idCiclo
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
        22
      )
      .getValues();

  const a =
    CONFIG_REMANEJAMENTO_EDICAO
      .COLUNAS_AGENDAMENTOS;

  const pacienteProcurado =
    normalizarTextoPaciente_(
      idPaciente
    );

  const cicloProcurado =
    normalizarTextoPaciente_(
      idCiclo
    );

  const hoje =
    removerHorarioRemanejamentoEdicao_(
      new Date()
    ).getTime();

  const resultado = [];

  dados.forEach(
    function(linha, indice) {
      const data =
        linha[a.DATA - 1];

      if (
        !(data instanceof Date)
      ) {
        return;
      }

      if (
        normalizarTextoPaciente_(
          linha[a.ID_PACIENTE - 1]
        ) !== pacienteProcurado ||
        normalizarTextoPaciente_(
          linha[a.ID_CICLO - 1]
        ) !== cicloProcurado ||
        normalizarTextoPaciente_(
          linha[a.EVENTO - 1]
        ) !== 'sessao' ||
        normalizarTextoPaciente_(
          linha[a.STATUS - 1]
        ) !== 'agendado' ||
        removerHorarioRemanejamentoEdicao_(
          data
        ).getTime() < hoje
      ) {
        return;
      }

      resultado.push({
        linha:
          indice + 2,
        idAgendamento:
          String(
            linha[
              a.ID_AGENDAMENTO - 1
            ] || ''
          ).trim(),
        idPaciente:
          String(
            linha[
              a.ID_PACIENTE - 1
            ] || ''
          ).trim(),
        numeroSessao:
          Number(
            linha[
              a.NUMERO_SESSAO - 1
            ]
          ) || 0,
        data:
          new Date(data),
        valores:
          linha.slice()
      });
    }
  );

  resultado.sort(
    function(a, b) {
      if (
        a.numeroSessao !==
          b.numeroSessao
      ) {
        return (
          a.numeroSessao -
          b.numeroSessao
        );
      }

      return (
        a.data.getTime() -
        b.data.getTime()
      );
    }
  );

  return resultado;
}


/**
 * Lê ocupações ignorando todas as sessões futuras atuais
 * do paciente que serão substituídas.
 */
function lerOcupacoesRemanejamentoEdicao_(
  abaAgendamentos,
  idsIgnorados
) {
  const ignorados = {};

  idsIgnorados.forEach(
    function(id) {
      ignorados[
        normalizarTextoPaciente_(
          id
        )
      ] = true;
    }
  );

  const ultimaLinha =
    abaAgendamentos.getLastRow();

  const ocupacoes = {};

  if (ultimaLinha < 2) {
    return ocupacoes;
  }

  const dados =
    abaAgendamentos
      .getRange(
        2,
        1,
        ultimaLinha - 1,
        22
      )
      .getValues();

  const a =
    CONFIG_REMANEJAMENTO_EDICAO
      .COLUNAS_AGENDAMENTOS;

  const statusAceitos =
    CONFIG_VAGAS_SESSOES
      .STATUS_QUE_OCUPAM_VAGA
      .map(
        normalizarTextoVagasSessoes_
      );

  dados.forEach(function(linha) {
    const id =
      normalizarTextoPaciente_(
        linha[
          a.ID_AGENDAMENTO - 1
        ]
      );

    if (ignorados[id]) {
      return;
    }

    const data =
      linha[a.DATA - 1];

    const horario =
      linha[a.HORARIO - 1];

    const fisioterapeuta =
      String(
        linha[
          a.FISIOTERAPEUTA - 1
        ] || ''
      ).trim();

    const tipoGrupo =
      String(
        linha[
          a.TIPO_GRUPO - 1
        ] || ''
      ).trim();

    const evento =
      normalizarTextoVagasSessoes_(
        linha[a.EVENTO - 1]
      );

    const status =
      normalizarTextoVagasSessoes_(
        linha[a.STATUS - 1]
      );

    if (
      !(data instanceof Date) ||
      !horario ||
      !fisioterapeuta ||
      evento !== 'sessao' ||
      statusAceitos.indexOf(
        status
      ) === -1
    ) {
      return;
    }

    const chave =
      montarChaveOcupacaoVagasSessoes_(
        data,
        horario,
        fisioterapeuta
      );

    if (!ocupacoes[chave]) {
      ocupacoes[chave] = {
        quantidade: 0,
        tipos: {}
      };
    }

    ocupacoes[chave]
      .quantidade++;

    const tipoNormalizado =
      normalizarTextoVagasSessoes_(
        tipoGrupo
      );

    if (tipoNormalizado) {
      ocupacoes[chave]
        .tipos[
          tipoNormalizado
        ] = true;
    }
  });

  return ocupacoes;
}


/**
 * Impede duas sessões ativas do mesmo paciente na mesma data.
 */
function validarDatasDuplicadasPacienteEdicao_(
  abaAgendamentos,
  idPaciente,
  sessoesIgnoradas,
  novasDatas
) {
  const idsIgnorados = {};

  sessoesIgnoradas.forEach(
    function(sessao) {
      idsIgnorados[
        normalizarTextoPaciente_(
          sessao.idAgendamento
        )
      ] = true;
    }
  );

  const datasNovas = {};

  novasDatas.forEach(
    function(data) {
      datasNovas[
        chaveDataVagasSessoes_(
          data
        )
      ] = true;
    }
  );

  const ultimaLinha =
    abaAgendamentos.getLastRow();

  if (ultimaLinha < 2) {
    return;
  }

  const dados =
    abaAgendamentos
      .getRange(
        2,
        1,
        ultimaLinha - 1,
        22
      )
      .getValues();

  const a =
    CONFIG_REMANEJAMENTO_EDICAO
      .COLUNAS_AGENDAMENTOS;

  const procurado =
    normalizarTextoPaciente_(
      idPaciente
    );

  const statusAtivos =
    CONFIG_VAGAS_SESSOES
      .STATUS_QUE_OCUPAM_VAGA
      .map(
        normalizarTextoVagasSessoes_
      );

  dados.forEach(function(linha) {
    const id =
      normalizarTextoPaciente_(
        linha[
          a.ID_AGENDAMENTO - 1
        ]
      );

    if (idsIgnorados[id]) {
      return;
    }

    const data =
      linha[a.DATA - 1];

    if (
      !(data instanceof Date) ||
      !datasNovas[
        chaveDataVagasSessoes_(
          data
        )
      ]
    ) {
      return;
    }

    if (
      normalizarTextoPaciente_(
        linha[a.ID_PACIENTE - 1]
      ) !== procurado ||
      normalizarTextoPaciente_(
        linha[a.EVENTO - 1]
      ) !== 'sessao' ||
      statusAtivos.indexOf(
        normalizarTextoPaciente_(
          linha[a.STATUS - 1]
        )
      ) === -1
    ) {
      return;
    }

    throw new Error(
      'O paciente já possui outra sessão ativa em ' +
      formatarDataMensagemEdicao_(
        data
      ) +
      '.'
    );
  });
}


/**
 * Gera IDs para as novas sessões.
 */
function gerarIdsAgendamentoRemanejamentoEdicao_(
  abaAgendamentos,
  quantidade
) {
  const ultimaLinha =
    abaAgendamentos.getLastRow();

  let maiorNumero = 0;

  if (ultimaLinha >= 2) {
    abaAgendamentos
      .getRange(
        2,
        1,
        ultimaLinha - 1,
        1
      )
      .getDisplayValues()
      .forEach(function(linha) {
        const texto =
          String(
            linha[0] || ''
          ).trim();

        const resultado =
          texto.match(
            /(\d+)$/
          );

        if (!resultado) {
          return;
        }

        const numero =
          Number(
            resultado[1]
          );

        if (
          Number.isFinite(numero) &&
          numero > maiorNumero
        ) {
          maiorNumero =
            numero;
        }
      });
  }

  const ids = [];

  for (
    let indice = 1;
    indice <= quantidade;
    indice++
  ) {
    ids.push(
      'AG-' +
      String(
        maiorNumero + indice
      ).padStart(
        6,
        '0'
      )
    );
  }

  return ids;
}


/**
 * Adiciona o status técnico à lista oficial.
 */
function garantirStatusRemanejamentoEdicao_(
  ss
) {
  const aba =
    ss.getSheetByName(
      CONFIG_REMANEJAMENTO_EDICAO
        .ABA_STATUS_SESSAO
    );

  if (!aba) {
    throw new Error(
      'A aba "Status da Sessão" não foi encontrada.'
    );
  }

  const ultimaLinha =
    aba.getLastRow();

  const valores =
    ultimaLinha >= 2
      ? aba
          .getRange(
            2,
            1,
            ultimaLinha - 1,
            1
          )
          .getDisplayValues()
      : [];

  const procurado =
    normalizarTextoPaciente_(
      CONFIG_REMANEJAMENTO_EDICAO
        .STATUS_REMANEJADO
    );

  const existe =
    valores.some(
      function(linha) {
        return (
          normalizarTextoPaciente_(
            linha[0]
          ) === procurado
        );
      }
    );

  if (!existe) {
    aba
      .getRange(
        Math.max(
          ultimaLinha + 1,
          2
        ),
        1
      )
      .setValue(
        CONFIG_REMANEJAMENTO_EDICAO
          .STATUS_REMANEJADO
      );
  }
}


/**
 * Define status técnico sem esbarrar em validação de dados.
 */
function definirValorTecnicoRemanejamentoEdicao_(
  aba,
  linha,
  coluna,
  valor
) {
  const celula =
    aba.getRange(
      linha,
      coluna
    );

  const validacao =
    celula.getDataValidation();

  if (validacao) {
    celula.clearDataValidations();
  }

  celula.setValue(
    valor
  );

  if (validacao) {
    celula.setDataValidation(
      validacao
    );
  }
}


/**
 * Dias marcados no formulário em números do JavaScript.
 */
function obterDiasNumericosEdicao_(
  dias
) {
  const resultado = [];

  if (dias.segunda) {
    resultado.push(1);
  }

  if (dias.terca) {
    resultado.push(2);
  }

  if (dias.quarta) {
    resultado.push(3);
  }

  if (dias.quinta) {
    resultado.push(4);
  }

  if (dias.sexta) {
    resultado.push(5);
  }

  return resultado;
}


function nomeDiaRemanejamentoEdicao_(
  numero
) {
  const nomes = {
    1: 'Segunda-feira',
    2: 'Terça-feira',
    3: 'Quarta-feira',
    4: 'Quinta-feira',
    5: 'Sexta-feira'
  };

  return (
    nomes[
      Number(numero)
    ] || ''
  );
}


function chaveHorarioRemanejamentoEdicao_(
  valor
) {
  if (
    valor instanceof Date &&
    !isNaN(
      valor.getTime()
    )
  ) {
    return Utilities.formatDate(
      valor,
      Session.getScriptTimeZone(),
      'HH:mm'
    );
  }

  const texto =
    String(
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
      Number(
        resultado[1]
      )
    ).padStart(
      2,
      '0'
    ) +
    ':' +
    resultado[2]
  );
}


function removerHorarioRemanejamentoEdicao_(
  data
) {
  const copia =
    new Date(data);

  copia.setHours(
    0,
    0,
    0,
    0
  );

  return copia;
}


/**
 * Confirma que o motor de vagas necessário está disponível.
 */
function validarDependenciasRemanejamentoEdicao_() {
  const dependencias = [
    {
      nome:
        'lerHorariosPermitidosVagasSessoes_',
      tipo:
        typeof lerHorariosPermitidosVagasSessoes_
    },
    {
      nome:
        'lerFisioterapeutasPorTurnoVagasSessoes_',
      tipo:
        typeof lerFisioterapeutasPorTurnoVagasSessoes_
    },
    {
      nome:
        'lerFeriadosVagasSessoes_',
      tipo:
        typeof lerFeriadosVagasSessoes_
    },
    {
      nome:
        'lerBloqueiosVagasSessoes_',
      tipo:
        typeof lerBloqueiosVagasSessoes_
    },
    {
      nome:
        'simularSequenciaVagasSessoes_',
      tipo:
        typeof simularSequenciaVagasSessoes_
    },
    {
      nome:
        'montarChaveOcupacaoVagasSessoes_',
      tipo:
        typeof montarChaveOcupacaoVagasSessoes_
    },
    {
      nome:
        'chaveDataVagasSessoes_',
      tipo:
        typeof chaveDataVagasSessoes_
    }
  ];

  const ausente =
    dependencias.find(
      function(item) {
        return (
          item.tipo !==
            'function'
        );
      }
    );

  if (ausente) {
    throw new Error(
      'O arquivo VagasSessoes.gs precisa estar atualizado. Função ausente: ' +
      ausente.nome
    );
  }
}

/**
 * Localiza o paciente pelo ID.
 */
function localizarLinhaPacientePorId_(
  aba,
  idPaciente
) {
  const ultimaLinha =
    aba.getLastRow();

  if (ultimaLinha < 2) {
    return null;
  }

  const ids =
    aba
      .getRange(
        2,
        CONFIG_MODULO_PACIENTES
          .COLUNAS.ID,
        ultimaLinha - 1,
        1
      )
      .getDisplayValues();

  const idProcurado =
    normalizarTextoPaciente_(
      idPaciente
    );

  for (
    let indice = 0;
    indice < ids.length;
    indice++
  ) {
    const idAtual =
      normalizarTextoPaciente_(
        ids[indice][0]
      );

    if (
      idAtual === idProcurado
    ) {
      return {
        linha: indice + 2
      };
    }
  }

  return null;
}


/**
 * Monta o resumo apresentado na busca.
 */
function montarResumoPacienteEdicao_(
  linha,
  numeroLinha
) {
  return {
    linha: numeroLinha,

    id: String(
      linha[
        CONFIG_MODULO_PACIENTES
          .COLUNAS.ID - 1
      ] || ''
    ).trim(),

    prontuario: String(
      linha[
        CONFIG_MODULO_PACIENTES
          .COLUNAS.PRONTUARIO - 1
      ] || ''
    ).trim(),

    nome: String(
      linha[
        CONFIG_MODULO_PACIENTES
          .COLUNAS.NOME - 1
      ] || ''
    ).trim(),

    cpf: String(
      linha[
        CONFIG_MODULO_PACIENTES
          .COLUNAS.CPF - 1
      ] || ''
    ).trim(),

    telefone: String(
      linha[
        CONFIG_MODULO_PACIENTES
          .COLUNAS.TELEFONE - 1
      ] || ''
    ).trim(),

    status: String(
      linha[
        CONFIG_MODULO_PACIENTES
          .COLUNAS.STATUS - 1
      ] || ''
    ).trim()
  };
}


/**
 * Monta os dados completos utilizados pelo formulário.
 */
function montarDadosCompletosPacienteEdicao_(
  linha
) {
  return {
    id: String(
      linha[
        CONFIG_MODULO_PACIENTES
          .COLUNAS.ID - 1
      ] || ''
    ).trim(),

    prontuario: String(
      linha[
        CONFIG_MODULO_PACIENTES
          .COLUNAS.PRONTUARIO - 1
      ] || ''
    ).trim(),

    nome: String(
      linha[
        CONFIG_MODULO_PACIENTES
          .COLUNAS.NOME - 1
      ] || ''
    ).trim(),

    cpf: String(
      linha[
        CONFIG_MODULO_PACIENTES
          .COLUNAS.CPF - 1
      ] || ''
    ).trim(),

    telefone: String(
      linha[
        CONFIG_MODULO_PACIENTES
          .COLUNAS.TELEFONE - 1
      ] || ''
    ).trim(),

    horarioSessao:
      formatarHorarioParaEdicao_(
        linha[
          CONFIG_MODULO_PACIENTES
            .COLUNAS.HORARIO - 1
        ]
      ),

    tipoAtendimento:
      String(
        linha[
          CONFIG_MODULO_PACIENTES
            .COLUNAS.TIPO_ATENDIMENTO - 1
        ] || ''
      ).trim(),

    segunda:
      converterCheckboxEdicao_(
        linha[
          CONFIG_MODULO_PACIENTES
            .COLUNAS.SEGUNDA - 1
        ]
      ),

    terca:
      converterCheckboxEdicao_(
        linha[
          CONFIG_MODULO_PACIENTES
            .COLUNAS.TERCA - 1
        ]
      ),

    quarta:
      converterCheckboxEdicao_(
        linha[
          CONFIG_MODULO_PACIENTES
            .COLUNAS.QUARTA - 1
        ]
      ),

    quinta:
      converterCheckboxEdicao_(
        linha[
          CONFIG_MODULO_PACIENTES
            .COLUNAS.QUINTA - 1
        ]
      ),

    sexta:
      converterCheckboxEdicao_(
        linha[
          CONFIG_MODULO_PACIENTES
            .COLUNAS.SEXTA - 1
        ]
      ),

    sessoesPrescritas:
      converterNumeroInteiroEdicao_(
        linha[
          CONFIG_MODULO_PACIENTES
            .COLUNAS.SESSOES_PRESCRITAS - 1
        ],
        0
      ),

    sessoesRealizadas:
      converterNumeroInteiroEdicao_(
        linha[
          CONFIG_MODULO_PACIENTES
            .COLUNAS.SESSOES_REALIZADAS - 1
        ],
        0
      ),

    sessoesRestantes:
      converterNumeroInteiroEdicao_(
        linha[
          CONFIG_MODULO_PACIENTES
            .COLUNAS.SESSOES_RESTANTES - 1
        ],
        0
      ),

    dataAvaliacao:
      formatarDataParaCampoHtml_(
        linha[
          CONFIG_MODULO_PACIENTES
            .COLUNAS.DATA_AVALIACAO - 1
        ]
      ),

    horarioAvaliacao:
      formatarHorarioParaEdicao_(
        linha[
          CONFIG_MODULO_PACIENTES
            .COLUNAS.HORARIO_AVALIACAO - 1
        ]
      ),

    status: String(
      linha[
        CONFIG_MODULO_PACIENTES
          .COLUNAS.STATUS - 1
      ] || ''
    ).trim(),

    fisioterapeuta: String(
      linha[
        CONFIG_MODULO_PACIENTES
          .COLUNAS.FISIOTERAPEUTA - 1
      ] || ''
    ).trim(),

    observacao: String(
      linha[
        CONFIG_MODULO_PACIENTES
          .COLUNAS.OBSERVACAO - 1
      ] || ''
    ).trim()
  };
}
/**
 * Valida os dados enviados pelo formulário.
 */
function validarDadosEdicaoPaciente_(
  dados
) {
  const nome =
    String(
      dados.nome || ''
    ).trim();

  const telefone =
    String(
      dados.telefone || ''
    ).trim();

  const cpfNumeros =
    String(
      dados.cpf || ''
    ).replace(
      /\D/g,
      ''
    );

  const dataAvaliacao =
    converterDataFormulario_(
      dados.dataAvaliacao
    );

  const horarioAvaliacao =
    converterHorarioFormulario_(
      dados.horarioAvaliacao
    );

  const horarioSessao =
    dados.horarioSessao
      ? converterHorarioFormulario_(
          dados.horarioSessao
        )
      : '';

  const tipoAtendimentoInformado =
    String(
      dados.tipoAtendimento || ''
    ).trim();

  const tipoAtendimento =
    padronizarTipoAtendimentoEdicao_(
      tipoAtendimentoInformado
    );

  const fisioterapeuta =
    String(
      dados.fisioterapeuta || ''
    ).trim();

  const observacao =
    String(
      dados.observacao || ''
    ).trim();

  const sessoesPrescritas =
    Number(
      dados.sessoesPrescritas || 0
    );

  if (!nome) {
    throw new Error(
      'Informe o nome completo do paciente.'
    );
  }

  if (!telefone) {
    throw new Error(
      'Informe o telefone do paciente.'
    );
  }

  if (!dataAvaliacao) {
    throw new Error(
      'Informe uma data de avaliação válida.'
    );
  }

  if (!horarioAvaliacao) {
    throw new Error(
      'Informe o horário da avaliação.'
    );
  }

  if (
    cpfNumeros &&
    (
      cpfNumeros.length !== 11 ||
      !validarCpfNovoPaciente_(
        cpfNumeros
      )
    )
  ) {
    throw new Error(
      'O CPF informado é inválido.'
    );
  }

  if (
    !Number.isInteger(
      sessoesPrescritas
    ) ||
    sessoesPrescritas < 0 ||
    sessoesPrescritas > 20
  ) {
    throw new Error(
      'As sessões prescritas devem ser um número inteiro entre 0 e 20.'
    );
  }

  const dias = {
    segunda:
      Boolean(dados.segunda),

    terca:
      Boolean(dados.terca),

    quarta:
      Boolean(dados.quarta),

    quinta:
      Boolean(dados.quinta),

    sexta:
      Boolean(dados.sexta)
  };

  const possuiDiaSelecionado =
    Object
      .keys(dias)
      .some(
        function(chave) {
          return dias[chave];
        }
      );

  if (
    sessoesPrescritas > 0 &&
    !tipoAtendimentoInformado
  ) {
    throw new Error(
      'Selecione o tipo de atendimento.'
    );
  }

  if (
    tipoAtendimentoInformado &&
    !tipoAtendimento
  ) {
    throw new Error(
      'O tipo de atendimento informado não é válido. Escolha Maior supervisão, Grupo de MMSS, Grupo de MMII ou Grupo de Coluna.'
    );
  }

  if (
    sessoesPrescritas > 0 &&
    !horarioSessao
  ) {
    throw new Error(
      'Selecione o horário preferencial das sessões.'
    );
  }

  if (
    sessoesPrescritas > 0 &&
    !possuiDiaSelecionado
  ) {
    throw new Error(
      'Marque pelo menos um dia de preferência.'
    );
  }

  return {
    nome:
      normalizarNomePaciente_(
        nome
      ),

    cpf:
      cpfNumeros,

    cpfFormatado:
      cpfNumeros
        ? formatarCpfPaciente_(
            cpfNumeros
          )
        : '',

    telefone:
      telefone,

    dataAvaliacao:
      dataAvaliacao,

    horarioAvaliacao:
      horarioAvaliacao,

    horarioSessao:
      horarioSessao,

    tipoAtendimento:
      tipoAtendimento,

    fisioterapeuta:
      fisioterapeuta,

    sessoesPrescritas:
      sessoesPrescritas,

    dias:
      dias,

    observacao:
      observacao
  };
}


/**
 * Verifica se o CPF pertence a outro paciente.
 */
function verificarCpfDuplicadoNaEdicao_(
  aba,
  cpf,
  idPacienteAtual
) {
  const cpfProcurado =
    String(
      cpf || ''
    ).replace(
      /\D/g,
      ''
    );

  if (!cpfProcurado) {
    return;
  }

  const ultimaLinha =
    aba.getLastRow();

  if (ultimaLinha < 2) {
    return;
  }

  const dados =
    aba
      .getRange(
        2,
        1,
        ultimaLinha - 1,
        4
      )
      .getDisplayValues();

  dados.forEach(
    function(linha) {
      const idExistente =
        String(
          linha[0] || ''
        ).trim();

      const cpfExistente =
        String(
          linha[3] || ''
        ).replace(
          /\D/g,
          ''
        );

      if (
        idExistente &&
        idExistente !==
          idPacienteAtual &&
        cpfExistente ===
          cpfProcurado
      ) {
        throw new Error(
          'Este CPF já pertence a outro paciente.'
        );
      }
    }
  );
}


/**
 * Padroniza o tipo de atendimento recebido.
 */
function padronizarTipoAtendimentoEdicao_(
  valor
) {
  const chave =
    normalizarTextoPaciente_(
      valor
    );

  if (!chave) {
    return '';
  }

  return (
    CONFIG_AJUSTE_CICLO_EDICAO
      .TIPOS_ATENDIMENTO[chave] ||
    ''
  );
}


/**
 * Define o status depois da edição.
 */
function definirStatusAposEdicaoPaciente_(
  statusAtual,
  sessoesPrescritas,
  sessoesRealizadas,
  possuiCicloAtivo
) {
  const statusNormalizado =
    normalizarTextoPaciente_(
      statusAtual
    );

  if (
    statusNormalizado ===
      'ciclo concluido' ||
    statusNormalizado ===
      'inativo'
  ) {
    return statusAtual;
  }

  if (
    sessoesPrescritas > 0 &&
    sessoesRealizadas >=
      sessoesPrescritas
  ) {
    return 'Ciclo concluído';
  }

  if (
    sessoesPrescritas > 0 &&
    (
      sessoesRealizadas > 0 ||
      possuiCicloAtivo
    )
  ) {
    return 'Em tratamento';
  }

  if (
    sessoesPrescritas > 0
  ) {
    return (
      'Avaliado – aguardando agendamento'
    );
  }

  return 'Avaliação agendada';
}


/**
 * Formata uma data para o campo HTML.
 */
function formatarDataParaCampoHtml_(
  valor
) {
  if (!valor) {
    return '';
  }

  if (
    valor instanceof Date
  ) {
    return Utilities.formatDate(
      valor,
      Session.getScriptTimeZone(),
      'yyyy-MM-dd'
    );
  }

  const texto =
    String(
      valor || ''
    ).trim();

  const brasileiro =
    texto.match(
      /^(\d{2})\/(\d{2})\/(\d{4})$/
    );

  if (brasileiro) {
    return (
      brasileiro[3] +
      '-' +
      brasileiro[2] +
      '-' +
      brasileiro[1]
    );
  }

  const internacional =
    texto.match(
      /^(\d{4})-(\d{2})-(\d{2})$/
    );

  if (internacional) {
    return texto;
  }

  return '';
}


/**
 * Formata o horário utilizado pelo formulário.
 */
function formatarHorarioParaEdicao_(
  valor
) {
  if (!valor) {
    return '';
  }

  return formatarHorarioFormulario_(
    valor
  );
}


/**
 * Converte valores das caixas de seleção.
 */
function converterCheckboxEdicao_(
  valor
) {
  if (valor === true) {
    return true;
  }

  const texto =
    normalizarTextoPaciente_(
      valor
    );

  return (
    texto === 'true' ||
    texto === 'verdadeiro' ||
    texto === 'sim' ||
    texto === 'x'
  );
}


/**
 * Converte um valor numérico em inteiro.
 */
function converterNumeroInteiroEdicao_(
  valor,
  valorPadrao
) {
  const numero =
    Number(
      valor
    );

  if (
    Number.isFinite(
      numero
    ) &&
    Number.isInteger(
      numero
    )
  ) {
    return numero;
  }

  return valorPadrao;
}
/**
 * Ajusta a quantidade de sessões do ciclo.
 *
 * Atualiza somente as linhas do ciclo selecionado,
 * sem regravar toda a aba Agendamentos.
 */
function ajustarQuantidadeSessoesCicloEdicao_(
  ss,
  idPaciente,
  quantidadeAnterior,
  novaQuantidade,
  sessoesRealizadas
) {
  const resultadoPadrao = {
    quantidadeAlterada: false,
    possuiCicloAtivo: false,
    sessoesEncerradas: 0,
    dataTermino: '',
    quantidadeAnteriorCiclo:
      quantidadeAnterior
  };

  const abaAgendamentos =
    ss.getSheetByName(
      CONFIG_AJUSTE_CICLO_EDICAO
        .ABA_AGENDAMENTOS
    );

  if (!abaAgendamentos) {
    throw new Error(
      'A aba "Agendamentos" não foi encontrada.'
    );
  }

  const ultimaLinha =
    abaAgendamentos.getLastRow();

  if (ultimaLinha < 2) {
    return resultadoPadrao;
  }

  const dados =
    abaAgendamentos
      .getRange(
        2,
        1,
        ultimaLinha - 1,
        CONFIG_AJUSTE_CICLO_EDICAO
          .QUANTIDADE_COLUNAS
      )
      .getValues();

  const cicloAtual =
    localizarCicloAtualEdicao_(
      dados,
      idPaciente
    );

  if (!cicloAtual) {
    return resultadoPadrao;
  }

  resultadoPadrao
    .possuiCicloAtivo =
      existeSessaoAtivaCicloEdicao_(
        dados,
        cicloAtual.idCiclo
      );

  resultadoPadrao
    .dataTermino =
      obterDataTerminoCicloEdicao_(
        dados,
        cicloAtual.idCiclo
      );

  const quantidadeCiclo =
    obterQuantidadePrescritaCicloEdicao_(
      dados,
      cicloAtual.idCiclo
    ) || quantidadeAnterior;

  resultadoPadrao
    .quantidadeAnteriorCiclo =
      quantidadeCiclo;

  if (
    novaQuantidade ===
      quantidadeCiclo
  ) {
    return resultadoPadrao;
  }

  if (
    novaQuantidade >
      quantidadeCiclo
  ) {
    throw new Error(
      'Este paciente já possui um ciclo em andamento com ' +
        quantidadeCiclo +
        ' sessões. Para aumentar a quantidade, conclua o ciclo atual e registre uma renovação. A edição direta só pode reduzir o ciclo existente.'
    );
  }

  if (
    novaQuantidade <
      sessoesRealizadas
  ) {
    throw new Error(
      'A nova quantidade não pode ser menor que as sessões já realizadas.'
    );
  }

  const colunas =
    CONFIG_AJUSTE_CICLO_EDICAO
      .COLUNAS;

  const agora =
    new Date();

  let sessoesEncerradas = 0;

  const linhasAlteradas = [];

  dados.forEach(
    function(linha, indice) {
      const idCiclo =
        String(
          linha[
            colunas.ID_CICLO - 1
          ] || ''
        ).trim();

      if (
        idCiclo !==
          cicloAtual.idCiclo
      ) {
        return;
      }

      const evento =
        normalizarTextoPaciente_(
          linha[
            colunas.EVENTO - 1
          ]
        );

      if (
        evento !== 'sessao' &&
        evento !==
          'sessao encerrada por ajuste'
      ) {
        return;
      }

      const numeroSessao =
        Number(
          linha[
            colunas.NUMERO_SESSAO - 1
          ]
        ) || 0;

      if (
        numeroSessao > 0 &&
        numeroSessao <=
          novaQuantidade &&
        evento === 'sessao'
      ) {
        linha[
          colunas.TOTAL_PRESCRITO - 1
        ] = novaQuantidade;

        linha[
          colunas.ATUALIZADO_EM - 1
        ] = agora;

        linhasAlteradas.push({
          linhaPlanilha:
            indice + 2,

          dados:
            linha,

          encerrada:
            false
        });

        return;
      }

      if (
        numeroSessao <=
          novaQuantidade
      ) {
        return;
      }

      if (
        evento === 'sessao'
      ) {
        sessoesEncerradas++;
      }

      linha[
        colunas.EVENTO - 1
      ] =
        CONFIG_AJUSTE_CICLO_EDICAO
          .EVENTO_ENCERRADO;

      linha[
        colunas.STATUS - 1
      ] =
        CONFIG_AJUSTE_CICLO_EDICAO
          .STATUS_ENCERRADO;

      linha[
        colunas.MOTIVO - 1
      ] =
        adicionarMotivoAjusteCicloEdicao_(
          linha[
            colunas.MOTIVO - 1
          ],
          quantidadeCiclo,
          novaQuantidade
        );

      linha[
        colunas.CONTA_COMO_SESSAO - 1
      ] = 'Não';

      linha[
        colunas.AVISAR_PACIENTE - 1
      ] = 'Não';

      linha[
        colunas.ATUALIZADO_EM - 1
      ] = agora;

      linha[
        colunas.FATURAVEL - 1
      ] = 'Não';

      linhasAlteradas.push({
        linhaPlanilha:
          indice + 2,

        dados:
          linha,

        encerrada:
          true
      });
    }
  );

  /*
   * Grava somente as linhas pertencentes ao ciclo
   * que foi efetivamente alterado.
   */
  linhasAlteradas.forEach(
    function(item) {
      const numeroLinha =
        item.linhaPlanilha;

      const linha =
        item.dados;

      abaAgendamentos
        .getRange(
          numeroLinha,
          colunas.TOTAL_PRESCRITO
        )
        .setValue(
          linha[
            colunas.TOTAL_PRESCRITO - 1
          ]
        );

      if (item.encerrada) {
        abaAgendamentos
          .getRange(
            numeroLinha,
            colunas.EVENTO
          )
          .setValue(
            linha[
              colunas.EVENTO - 1
            ]
          );

        definirValorComValidacaoPreservadaEdicao_(
          abaAgendamentos,
          numeroLinha,
          colunas.STATUS,
          linha[
            colunas.STATUS - 1
          ]
        );

        abaAgendamentos
          .getRange(
            numeroLinha,
            colunas.MOTIVO,
            1,
            3
          )
          .setValues([[
            linha[
              colunas.MOTIVO - 1
            ],

            linha[
              colunas.CONTA_COMO_SESSAO - 1
            ],

            linha[
              colunas.AVISAR_PACIENTE - 1
            ]
          ]]);

        abaAgendamentos
          .getRange(
            numeroLinha,
            colunas.FATURAVEL
          )
          .setValue(
            linha[
              colunas.FATURAVEL - 1
            ]
          );
      }

      abaAgendamentos
        .getRange(
          numeroLinha,
          colunas.ATUALIZADO_EM
        )
        .setValue(
          linha[
            colunas.ATUALIZADO_EM - 1
          ]
        )
        .setNumberFormat(
          'dd/MM/yyyy HH:mm'
        );
    }
  );

  const dataTermino =
    obterDataTerminoCicloEdicao_(
      dados,
      cicloAtual.idCiclo
    );

  return {
    quantidadeAlterada: true,

    possuiCicloAtivo:
      novaQuantidade >
        sessoesRealizadas,

    sessoesEncerradas:
      sessoesEncerradas,

    dataTermino:
      dataTermino,

    quantidadeAnteriorCiclo:
      quantidadeCiclo
  };
}
/**
 * Atualiza o tipo de atendimento somente nas sessões
 * pertencentes ao ciclo mais recente do paciente.
 *
 * Não regrava agendamentos de outros pacientes.
 */
function atualizarTipoAtendimentoCicloEdicao_(
  ss,
  idPaciente,
  tipoAtendimento,
  limiteGrupo
) {
  if (!tipoAtendimento) {
    return 0;
  }

  const abaAgendamentos =
    ss.getSheetByName(
      CONFIG_AJUSTE_CICLO_EDICAO
        .ABA_AGENDAMENTOS
    );

  if (!abaAgendamentos) {
    throw new Error(
      'A aba "Agendamentos" não foi encontrada.'
    );
  }

  const ultimaLinha =
    abaAgendamentos.getLastRow();

  if (ultimaLinha < 2) {
    return 0;
  }

  const dados =
    abaAgendamentos
      .getRange(
        2,
        1,
        ultimaLinha - 1,
        CONFIG_AJUSTE_CICLO_EDICAO
          .QUANTIDADE_COLUNAS
      )
      .getValues();

  const cicloAtual =
    localizarCicloAtualEdicao_(
      dados,
      idPaciente
    );

  if (!cicloAtual) {
    return 0;
  }

  const colunas =
    CONFIG_AJUSTE_CICLO_EDICAO
      .COLUNAS;

  const agora =
    new Date();

  let quantidadeAtualizada = 0;

  const linhasAtualizadas = [];

  dados.forEach(
    function(linha, indice) {
      const idCiclo =
        String(
          linha[
            colunas.ID_CICLO - 1
          ] || ''
        ).trim();

      if (
        idCiclo !==
          cicloAtual.idCiclo
      ) {
        return;
      }

      const evento =
        normalizarTextoPaciente_(
          linha[
            colunas.EVENTO - 1
          ]
        );

      if (
        evento !== 'sessao' &&
        evento !==
          'sessao encerrada por ajuste'
      ) {
        return;
      }

      const tipoAtual =
        String(
          linha[
            colunas.TIPO_GRUPO - 1
          ] || ''
        ).trim();

      const limiteAtual =
        Number(
          linha[
            colunas.LIMITE_GRUPO - 1
          ]
        ) || 0;

      if (
        normalizarTextoPaciente_(
          tipoAtual
        ) ===
          normalizarTextoPaciente_(
            tipoAtendimento
          ) &&
        limiteAtual ===
          limiteGrupo
      ) {
        return;
      }

      linhasAtualizadas.push({
        linhaPlanilha:
          indice + 2,

        tipoGrupo:
          tipoAtendimento,

        limiteGrupo:
          limiteGrupo,

        atualizadoEm:
          agora
      });

      quantidadeAtualizada++;
    }
  );

  if (
    quantidadeAtualizada === 0
  ) {
    return 0;
  }

  /*
   * Grava somente as três células necessárias
   * de cada sessão pertencente ao ciclo.
   */
  linhasAtualizadas.forEach(
    function(item) {
      abaAgendamentos
        .getRange(
          item.linhaPlanilha,
          colunas.TIPO_GRUPO
        )
        .setValue(
          item.tipoGrupo
        );

      abaAgendamentos
        .getRange(
          item.linhaPlanilha,
          colunas.LIMITE_GRUPO
        )
        .setValue(
          item.limiteGrupo
        );

      abaAgendamentos
        .getRange(
          item.linhaPlanilha,
          colunas.ATUALIZADO_EM
        )
        .setValue(
          item.atualizadoEm
        )
        .setNumberFormat(
          'dd/MM/yyyy HH:mm'
        );
    }
  );

  return quantidadeAtualizada;
}


/**
 * Grava um valor técnico sem perder a regra de
 * validação existente na célula.
 */
function definirValorComValidacaoPreservadaEdicao_(
  aba,
  linha,
  coluna,
  valor
) {
  const celula =
    aba.getRange(
      linha,
      coluna
    );

  const validacao =
    celula.getDataValidation();

  if (validacao) {
    celula.clearDataValidations();
  }

  celula.setValue(valor);

  if (validacao) {
    celula.setDataValidation(
      validacao
    );
  }
}


/**
 * Obtém o total prescrito registrado no ciclo.
 */
function obterQuantidadePrescritaCicloEdicao_(
  dados,
  idCiclo
) {
  const colunas =
    CONFIG_AJUSTE_CICLO_EDICAO
      .COLUNAS;

  let quantidade = 0;

  dados.forEach(
    function(linha) {
      const cicloAtual =
        String(
          linha[
            colunas.ID_CICLO - 1
          ] || ''
        ).trim();

      const evento =
        normalizarTextoPaciente_(
          linha[
            colunas.EVENTO - 1
          ]
        );

      const total =
        Number(
          linha[
            colunas.TOTAL_PRESCRITO - 1
          ]
        ) || 0;

      if (
        cicloAtual === idCiclo &&
        evento === 'sessao' &&
        total > quantidade
      ) {
        quantidade = total;
      }
    }
  );

  return quantidade;
}


/**
 * Localiza o ciclo mais recente do paciente.
 */
function localizarCicloAtualEdicao_(
  dados,
  idPaciente
) {
  const colunas =
    CONFIG_AJUSTE_CICLO_EDICAO
      .COLUNAS;

  const idProcurado =
    normalizarTextoPaciente_(
      idPaciente
    );

  let resultado = null;

  dados.forEach(
    function(linha) {
      const idAtual =
        normalizarTextoPaciente_(
          linha[
            colunas.ID_PACIENTE - 1
          ]
        );

      const evento =
        normalizarTextoPaciente_(
          linha[
            colunas.EVENTO - 1
          ]
        );

      if (
        idAtual !== idProcurado ||
        (
          evento !== 'sessao' &&
          evento !==
            'sessao encerrada por ajuste'
        )
      ) {
        return;
      }

      const idCiclo =
        String(
          linha[
            colunas.ID_CICLO - 1
          ] || ''
        ).trim();

      const numeroCiclo =
        Number(
          linha[
            colunas.CICLO_NUMERO - 1
          ]
        ) || 0;

      if (
        idCiclo &&
        (
          !resultado ||
          numeroCiclo >
            resultado.numeroCiclo
        )
      ) {
        resultado = {
          idCiclo:
            idCiclo,

          numeroCiclo:
            numeroCiclo
        };
      }
    }
  );

  return resultado;
}


/**
 * Verifica se existe sessão ativa no ciclo.
 */
function existeSessaoAtivaCicloEdicao_(
  dados,
  idCiclo
) {
  const colunas =
    CONFIG_AJUSTE_CICLO_EDICAO
      .COLUNAS;

  return dados.some(
    function(linha) {
      const cicloAtual =
        String(
          linha[
            colunas.ID_CICLO - 1
          ] || ''
        ).trim();

      const evento =
        normalizarTextoPaciente_(
          linha[
            colunas.EVENTO - 1
          ]
        );

      return (
        cicloAtual === idCiclo &&
        evento === 'sessao'
      );
    }
  );
}


/**
 * Calcula a última data válida do ciclo.
 */
function obterDataTerminoCicloEdicao_(
  dados,
  idCiclo
) {
  const colunas =
    CONFIG_AJUSTE_CICLO_EDICAO
      .COLUNAS;

  let ultimaData = null;

  dados.forEach(
    function(linha) {
      const cicloAtual =
        String(
          linha[
            colunas.ID_CICLO - 1
          ] || ''
        ).trim();

      const evento =
        normalizarTextoPaciente_(
          linha[
            colunas.EVENTO - 1
          ]
        );

      const status =
        normalizarTextoPaciente_(
          linha[
            colunas.STATUS - 1
          ]
        );

      const data =
        linha[
          colunas.DATA - 1
        ];

      if (
        cicloAtual !== idCiclo ||
        evento !== 'sessao' ||
        status ===
          'encerrado por ajuste' ||
        status ===
          'cancelado pela clinica' ||
        !(
          data instanceof Date
        )
      ) {
        return;
      }

      if (
        !ultimaData ||
        data.getTime() >
          ultimaData.getTime()
      ) {
        ultimaData =
          new Date(data);
      }
    }
  );

  return ultimaData || '';
}


/**
 * Acrescenta a justificativa da redução do ciclo.
 */
function adicionarMotivoAjusteCicloEdicao_(
  motivoAtual,
  quantidadeAnterior,
  novaQuantidade
) {
  const motivo =
    String(
      motivoAtual || ''
    ).trim();

  const identificacao =
    'Ciclo reduzido de ' +
    quantidadeAnterior +
    ' para ' +
    novaQuantidade +
    ' sessões';

  if (
    normalizarTextoPaciente_(
      motivo
    ).indexOf(
      normalizarTextoPaciente_(
        identificacao
      )
    ) !== -1
  ) {
    return motivo;
  }

  return motivo
    ? (
        motivo +
        ' | ' +
        identificacao
      )
    : identificacao;
}


/**
 * Formata a data apresentada na confirmação.
 */
function formatarDataMensagemEdicao_(
  data
) {
  if (
    !(data instanceof Date)
  ) {
    return 'não definida';
  }

  return Utilities.formatDate(
    data,
    Session.getScriptTimeZone(),
    'dd/MM/yyyy'
  );
}
