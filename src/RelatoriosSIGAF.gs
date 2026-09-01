const CONFIG_RELATORIOS_SIGAF = {
  ABAS: {
    RELATORIO: 'Relatório Mensal SIGAF',
    AGENDAMENTOS: 'Agendamentos',
    HISTORICO_DESFECHOS:
      'Histórico de Desfechos',
    CADASTRO: 'Cadastro de Pacientes'
  },

  AGENDAMENTOS: {
    ID_AGENDAMENTO: 1,
    ID_PACIENTE: 2,
    PRONTUARIO: 3,
    PACIENTE: 4,
    ID_CICLO: 5,
    CICLO_NUMERO: 6,
    DATA: 7,
    HORARIO: 9,
    FISIOTERAPEUTA: 10,
    EVENTO: 12,
    NUMERO_SESSAO: 13,
    TOTAL_PRESCRITO: 14,
    STATUS: 16,
    MOTIVO: 17,
    FATURAVEL: 22
  },

  HISTORICO: {
    ID_REGISTRO: 1,
    ID_PACIENTE: 2,
    PRONTUARIO: 3,
    PACIENTE: 4,
    ID_CICLO: 5,
    CICLO_NUMERO: 6,
    DESFECHO: 7,
    DATA_DESFECHO: 8,
    SESSOES_PRESCRITAS: 9,
    SESSOES_REALIZADAS: 10,
    SESSOES_RESTANTES: 11,
    FUTURAS_CANCELADAS: 12,
    MOTIVO: 13
  },

  CADASTRO: {
    ID_PACIENTE: 1,
    TELEFONE: 5
  },

  QUANTIDADE_COLUNAS_AGENDAMENTOS: 22,
  QUANTIDADE_COLUNAS_HISTORICO: 14,
  QUANTIDADE_COLUNAS_CADASTRO: 24,
  QUANTIDADE_COLUNAS_RELATORIO: 14,

  DESFECHOS: [
    'Alta',
    'Encaminhamento para APS',
    'Renovação',
    'Alta por abandono',
    'Desistência do tratamento'
  ]
};


/**
 * Abre diretamente o relatório do mês atual.
 */
function abrirRelatorioMesAtualSIGAF() {
  const competencia =
    interpretarCompetenciaRelatorio_(
      ''
    );

  executarAberturaRelatorioMensalSIGAF_(
    competencia
  );
}


/**
 * Mantém compatibilidade com o nome antigo.
 */
function abrirRelatorioMensalSIGAF() {
  abrirRelatorioOutroMesSIGAF();
}


/**
 * Solicita outra competência e abre o relatório.
 */
function abrirRelatorioOutroMesSIGAF() {
  const ui =
    SpreadsheetApp.getUi();

  const resposta = ui.prompt(
    'Escolher mês do relatório',
    'Informe a competência no formato MM/AAAA.\n\n' +
      'Exemplo: 08/2026.',
    ui.ButtonSet.OK_CANCEL
  );

  if (
    resposta.getSelectedButton() !==
      ui.Button.OK
  ) {
    return;
  }

  const competencia =
    interpretarCompetenciaRelatorio_(
      resposta.getResponseText()
    );

  executarAberturaRelatorioMensalSIGAF_(
    competencia
  );
}


/**
 * Motor único para abrir a competência solicitada.
 */
function executarAberturaRelatorioMensalSIGAF_(
  competencia
) {
  const ui =
    SpreadsheetApp.getUi();

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const lock =
    LockService.getDocumentLock();

  let bloqueioObtido = false;

  try {
    lock.waitLock(30000);
    bloqueioObtido = true;

    const resultado =
      gerarRelatorioMensalSIGAF_(
        ss,
        competencia
      );

    SpreadsheetApp.flush();

    resultado.aba.activate();

    ui.alert(
      'Relatório atualizado',
      'Competência: ' +
        formatarCompetenciaRelatorio_(
          competencia
        ) +
        '\n\nAvaliações marcadas: ' +
        resultado.resumo
          .avaliacoesMarcadas +
        '\nSessões marcadas: ' +
        resultado.resumo
          .sessoesMarcadas +
        '\nSessões remanejadas: ' +
        (
          resultado.resumo
            .sessoesRemanejadas ||
          0
        ) +
        '\nAtendimentos realizados: ' +
        resultado.resumo
          .totalRealizados +
        '\nDesfechos registrados: ' +
        resultado.resumo
          .totalDesfechos,
      ui.ButtonSet.OK
    );
  } catch (erro) {
    ui.alert(
      'Erro ao gerar relatório',
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


/**
 * Gera todos os dados e monta a aba do relatório.
 */
function gerarRelatorioMensalSIGAF_(
  ss,
  competencia
) {
  const abas =
    obterAbasRelatorioSIGAF_(
      ss
    );

  const cadastroPorId =
    lerCadastroRelatorio_(
      abas.cadastro
    );

  const agendamentos =
    lerAgendamentosRelatorio_(
      abas.agendamentos
    )
      .filter(function (registro) {
        return dataPertenceCompetenciaRelatorio_(
          registro.data,
          competencia
        );
      })
      .sort(
        compararRegistrosRelatorio_
      );

  const historico =
    removerDesfechosDuplicadosRelatorio_(
      lerHistoricoDesfechosRelatorio_(
        abas.historico
      )
        .filter(function (registro) {
          return dataPertenceCompetenciaRelatorio_(
            registro.data,
            competencia
          );
        })
    )
      .sort(
        compararRegistrosRelatorio_
      );

  const resumo =
    calcularResumoRelatorio_(
      agendamentos,
      historico
    );

  const producao =
    calcularProducaoFisioterapeutasRelatorio_(
      agendamentos
    );

  montarAbaRelatorioMensal_(
    abas.relatorio,
    competencia,
    resumo,
    producao,
    agendamentos,
    historico,
    cadastroPorId
  );

  return {
    aba: abas.relatorio,
    resumo: resumo
  };
}


/**
 * Localiza as abas utilizadas pelo relatório.
 */
function obterAbasRelatorioSIGAF_(
  ss
) {
  const nomes =
    CONFIG_RELATORIOS_SIGAF.ABAS;

  let abaRelatorio =
    ss.getSheetByName(
      nomes.RELATORIO
    );

  if (!abaRelatorio) {
    abaRelatorio =
      ss.insertSheet(
        nomes.RELATORIO
      );
  }

  const abas = {
    relatorio: abaRelatorio,

    agendamentos:
      ss.getSheetByName(
        nomes.AGENDAMENTOS
      ),

    historico:
      ss.getSheetByName(
        nomes.HISTORICO_DESFECHOS
      ),

    cadastro:
      ss.getSheetByName(
        nomes.CADASTRO
      )
  };

  [
    'agendamentos',
    'historico',
    'cadastro'
  ].forEach(function (chave) {
    if (!abas[chave]) {
      throw new Error(
        'A aba necessária "' +
          nomes[
            {
              agendamentos:
                'AGENDAMENTOS',
              historico:
                'HISTORICO_DESFECHOS',
              cadastro:
                'CADASTRO'
            }[chave]
          ] +
          '" não foi encontrada.'
      );
    }
  });

  return abas;
}


/**
 * Lê os telefones existentes no Cadastro de Pacientes.
 */
function lerCadastroRelatorio_(
  aba
) {
  const resultado = {};

  const ultimaLinha =
    aba.getLastRow();

  if (ultimaLinha < 2) {
    return resultado;
  }

  const dados =
    aba
      .getRange(
        2,
        1,
        ultimaLinha - 1,
        CONFIG_RELATORIOS_SIGAF
          .QUANTIDADE_COLUNAS_CADASTRO
      )
      .getValues();

  const c =
    CONFIG_RELATORIOS_SIGAF
      .CADASTRO;

  dados.forEach(function (linha) {
    const id = String(
      linha[
        c.ID_PACIENTE - 1
      ] || ''
    ).trim();

    if (!id) {
      return;
    }

    resultado[
      normalizarTextoRelatorio_(id)
    ] = {
      telefone: String(
        linha[
          c.TELEFONE - 1
        ] || ''
      ).trim()
    };
  });

  return resultado;
}
/**
 * Lê os atendimentos registrados.
 */
function lerAgendamentosRelatorio_(
  aba
) {
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
        CONFIG_RELATORIOS_SIGAF
          .QUANTIDADE_COLUNAS_AGENDAMENTOS
      )
      .getValues();

  const c =
    CONFIG_RELATORIOS_SIGAF
      .AGENDAMENTOS;

  return dados
    .map(function (linha) {
      return {
        id: String(
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
          linha[
            c.PRONTUARIO - 1
          ] || ''
        ).trim(),

        paciente: String(
          linha[
            c.PACIENTE - 1
          ] || ''
        ).trim(),

        idCiclo: String(
          linha[
            c.ID_CICLO - 1
          ] || ''
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

        evento: String(
          linha[
            c.EVENTO - 1
          ] || ''
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

        status: String(
          linha[
            c.STATUS - 1
          ] || ''
        ).trim(),

        motivo: String(
          linha[
            c.MOTIVO - 1
          ] || ''
        ).trim(),

        faturavel: String(
          linha[
            c.FATURAVEL - 1
          ] || ''
        ).trim()
      };
    })
    .filter(function (registro) {
      const evento =
        normalizarTextoRelatorio_(
          registro.evento
        );

      return (
        registro.id &&
        registro.data instanceof Date &&
        (
          evento === 'avaliacao' ||
          evento === 'sessao'
        )
      );
    });
}


/**
 * Lê os desfechos registrados no histórico.
 */
function lerHistoricoDesfechosRelatorio_(
  aba
) {
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
        CONFIG_RELATORIOS_SIGAF
          .QUANTIDADE_COLUNAS_HISTORICO
      )
      .getValues();

  const c =
    CONFIG_RELATORIOS_SIGAF
      .HISTORICO;

  return dados
    .map(function (linha) {
      return {
        id: String(
          linha[
            c.ID_REGISTRO - 1
          ] || ''
        ).trim(),

        idPaciente: String(
          linha[
            c.ID_PACIENTE - 1
          ] || ''
        ).trim(),

        prontuario: String(
          linha[
            c.PRONTUARIO - 1
          ] || ''
        ).trim(),

        paciente: String(
          linha[
            c.PACIENTE - 1
          ] || ''
        ).trim(),

        idCiclo: String(
          linha[
            c.ID_CICLO - 1
          ] || ''
        ).trim(),

        cicloNumero:
          Number(
            linha[
              c.CICLO_NUMERO - 1
            ]
          ) || 0,

        desfecho: String(
          linha[
            c.DESFECHO - 1
          ] || ''
        ).trim(),

        data:
          linha[
            c.DATA_DESFECHO - 1
          ],

        prescritas:
          Number(
            linha[
              c.SESSOES_PRESCRITAS - 1
            ]
          ) || 0,

        realizadas:
          Number(
            linha[
              c.SESSOES_REALIZADAS - 1
            ]
          ) || 0,

        restantes:
          Number(
            linha[
              c.SESSOES_RESTANTES - 1
            ]
          ) || 0,

        futurasCanceladas:
          Number(
            linha[
              c.FUTURAS_CANCELADAS - 1
            ]
          ) || 0,

        motivo: String(
          linha[
            c.MOTIVO - 1
          ] || ''
        ).trim()
      };
    })
    .filter(function (registro) {
      return (
        registro.id &&
        registro.idPaciente &&
        registro.data instanceof Date &&
        Boolean(registro.desfecho)
      );
    });
}


/**
 * Evita que o mesmo paciente apareça repetido
 * para o mesmo desfecho na mesma competência.
 */
function removerDesfechosDuplicadosRelatorio_(
  registros
) {
  const porChave = {};

  registros.forEach(function (registro) {
    /*
     * Um mesmo paciente pode ter mais de um ciclo e,
     * consequentemente, mais de um desfecho legítimo na
     * mesma competência.
     *
     * Por isso, a chave precisa considerar o ciclo.
     * Quando não houver ID de ciclo, usamos o próprio
     * ID do registro histórico para não eliminar um
     * desfecho válido.
     */
    const identificadorCiclo =
      registro.idCiclo
        ? normalizarTextoRelatorio_(
            registro.idCiclo
          )
        : (
            'sem-ciclo|' +
            normalizarTextoRelatorio_(
              registro.id
            )
          );

    const chave =
      normalizarTextoRelatorio_(
        registro.idPaciente
      ) +
      '|' +
      identificadorCiclo +
      '|' +
      normalizarTextoRelatorio_(
        registro.desfecho
      );

    if (
      !porChave[chave] ||
      registro.data.getTime() >
        porChave[chave].data.getTime()
    ) {
      porChave[chave] =
        registro;
    }
  });

  return Object.keys(
    porChave
  ).map(function (chave) {
    return porChave[chave];
  });
}


/**
 * Calcula os indicadores gerais da competência.
 */
function calcularResumoRelatorio_(
  agendamentos,
  historico
) {
  const avaliacoes =
    agendamentos.filter(
      function (registro) {
        return (
          normalizarTextoRelatorio_(
            registro.evento
          ) === 'avaliacao'
        );
      }
    );

  const sessoes =
    agendamentos.filter(
      function (registro) {
        return (
          normalizarTextoRelatorio_(
            registro.evento
          ) === 'sessao'
        );
      }
    );

  /*
   * Remanejamento é uma alteração administrativa da
   * programação, não uma sessão assistencial adicional.
   * Mantemos essas linhas no histórico, mas elas não
   * entram em "Sessões marcadas" nem em cancelamentos.
   */
  const sessoesRemanejadas =
    sessoes.filter(
      function (registro) {
        return (
          normalizarTextoRelatorio_(
            registro.status
          ) ===
            'cancelado por remanejamento'
        );
      }
    );

  const sessoesOperacionais =
    sessoes.filter(
      function (registro) {
        return (
          normalizarTextoRelatorio_(
            registro.status
          ) !==
            'cancelado por remanejamento'
        );
      }
    );

  const resumo = {
    avaliacoesMarcadas:
      avaliacoes.length,

    avaliacoesComparecidas:
      contarStatusRelatorio_(
        avaliacoes,
        'compareceu'
      ),

    avaliacoesFaltaJustificada:
      contarStatusRelatorio_(
        avaliacoes,
        'falta justificada'
      ),

    avaliacoesFaltaNaoJustificada:
      contarStatusRelatorio_(
        avaliacoes,
        'falta nao justificada'
      ),

    avaliacoesCanceladasClinica:
      contarStatusRelatorio_(
        avaliacoes,
        'cancelado pela clinica'
      ),

    avaliacoesSemResultado:
      contarSemResultadoRelatorio_(
        avaliacoes
      ),

    avaliacoesFuturas:
      contarFuturosRelatorio_(
        avaliacoes
      ),

    sessoesMarcadas:
      sessoesOperacionais.length,

    sessoesRemanejadas:
      sessoesRemanejadas.length,

    sessoesComparecidas:
      contarStatusRelatorio_(
        sessoesOperacionais,
        'compareceu'
      ),

    sessoesFaltaJustificada:
      contarStatusRelatorio_(
        sessoesOperacionais,
        'falta justificada'
      ),

    sessoesFaltaNaoJustificada:
      contarStatusRelatorio_(
        sessoesOperacionais,
        'falta nao justificada'
      ),

    sessoesCanceladasClinica:
      contarStatusRelatorio_(
        sessoesOperacionais,
        'cancelado pela clinica'
      ),

    sessoesOutrosCancelamentos:
      contarOutrosCancelamentosRelatorio_(
        sessoesOperacionais
      ),

    sessoesSemResultado:
      contarSemResultadoRelatorio_(
        sessoesOperacionais
      ),

    sessoesFuturas:
      contarFuturosRelatorio_(
        sessoesOperacionais
      ),

    totalDesfechos:
      historico.length
  };

  resumo.totalMarcados =
    resumo.avaliacoesMarcadas +
    resumo.sessoesMarcadas;

  resumo.totalRealizados =
    resumo.avaliacoesComparecidas +
    resumo.sessoesComparecidas;

  return resumo;
}


/**
 * Conta registros com determinado status.
 */
function contarStatusRelatorio_(
  registros,
  status
) {
  return registros.filter(
    function (registro) {
      return (
        normalizarTextoRelatorio_(
          registro.status
        ) === status
      );
    }
  ).length;
}


/**
 * Conta atendimentos passados que continuam como Agendado.
 */
function contarSemResultadoRelatorio_(
  registros
) {
  const hoje =
    removerHorarioRelatorio_(
      new Date()
    );

  return registros.filter(
    function (registro) {
      return (
        normalizarTextoRelatorio_(
          registro.status
        ) === 'agendado' &&

        removerHorarioRelatorio_(
          registro.data
        ).getTime() <
          hoje.getTime()
      );
    }
  ).length;
}


/**
 * Conta atendimentos futuros que permanecem agendados.
 */
function contarFuturosRelatorio_(
  registros
) {
  const hoje =
    removerHorarioRelatorio_(
      new Date()
    );

  return registros.filter(
    function (registro) {
      return (
        normalizarTextoRelatorio_(
          registro.status
        ) === 'agendado' &&

        removerHorarioRelatorio_(
          registro.data
        ).getTime() >=
          hoje.getTime()
      );
    }
  ).length;
}


/**
 * Conta outros cancelamentos que não sejam
 * cancelamentos realizados pela clínica.
 */
function contarOutrosCancelamentosRelatorio_(
  registros
) {
  return registros.filter(
    function (registro) {
      const status =
        normalizarTextoRelatorio_(
          registro.status
        );

      return (
        status.indexOf(
          'cancelado'
        ) === 0 &&
        status !==
          'cancelado pela clinica' &&
        status !==
          'cancelado por remanejamento'
      );
    }
  ).length;
}
/**
 * Calcula os atendimentos realizados
 * por fisioterapeuta.
 */
function calcularProducaoFisioterapeutasRelatorio_(
  agendamentos
) {
  const porProfissional = {};

  agendamentos.forEach(
    function (registro) {
      if (
        normalizarTextoRelatorio_(
          registro.status
        ) !== 'compareceu'
      ) {
        return;
      }

      const nome =
        registro.fisioterapeuta ||
        'Não informado';

      const chave =
        normalizarTextoRelatorio_(
          nome
        );

      if (!porProfissional[chave]) {
        porProfissional[chave] = {
          fisioterapeuta: nome,
          avaliacoes: 0,
          sessoes: 0
        };
      }

      const evento =
        normalizarTextoRelatorio_(
          registro.evento
        );

      if (evento === 'avaliacao') {
        porProfissional[chave]
          .avaliacoes++;
      }

      if (evento === 'sessao') {
        porProfissional[chave]
          .sessoes++;
      }
    }
  );

  return Object.keys(
    porProfissional
  )
    .map(function (chave) {
      const item =
        porProfissional[chave];

      item.total =
        item.avaliacoes +
        item.sessoes;

      return item;
    })
    .sort(function (a, b) {
      return a.fisioterapeuta
        .localeCompare(
          b.fisioterapeuta,
          'pt-BR'
        );
    });
}


/**
 * Monta visualmente a aba do relatório mensal.
 */
function montarAbaRelatorioMensal_(
  aba,
  competencia,
  resumo,
  producao,
  agendamentos,
  historico,
  cadastroPorId
) {
  prepararAbaRelatorio_(
    aba
  );

  const totalColunas =
    CONFIG_RELATORIOS_SIGAF
      .QUANTIDADE_COLUNAS_RELATORIO;

  const nomeMes =
    obterNomeMesRelatorio_(
      competencia.mes
    );

  aba
    .getRange(
      1,
      1,
      1,
      totalColunas
    )
    .merge()
    .setValue(
      'RELATÓRIO MENSAL SIGAF — ' +
        nomeMes.toUpperCase() +
        ' DE ' +
        competencia.ano
    )
    .setBackground('#4f81bd')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(16)
    .setHorizontalAlignment(
      'center'
    )
    .setVerticalAlignment(
      'middle'
    );

  aba.setRowHeight(1, 36);

  aba
    .getRange(
      2,
      1,
      1,
      totalColunas
    )
    .merge()
    .setValue(
      'Competência: ' +
        formatarCompetenciaRelatorio_(
          competencia
        ) +
        ' | Atualizado em: ' +
        Utilities.formatDate(
          new Date(),
          Session.getScriptTimeZone(),
          'dd/MM/yyyy HH:mm'
        )
    )
    .setFontColor('#666666')
    .setHorizontalAlignment(
      'center'
    );

  montarQuadrosResumoRelatorio_(
    aba,
    resumo,
    historico
  );

  let linha = 14;

  linha =
    montarTabelaProducaoRelatorio_(
      aba,
      linha,
      producao
    );

  const avaliacoes =
    agendamentos.filter(
      function (registro) {
        return (
          normalizarTextoRelatorio_(
            registro.evento
          ) === 'avaliacao'
        );
      }
    );

  const sessoes =
    agendamentos.filter(
      function (registro) {
        return (
          normalizarTextoRelatorio_(
            registro.evento
          ) === 'sessao'
        );
      }
    );

  linha =
    montarTabelaAvaliacoesRelatorio_(
      aba,
      linha,
      avaliacoes
    );

  linha =
    montarTabelaSessoesRelatorio_(
      aba,
      linha,
      sessoes
    );

  linha =
    montarTabelaFaltasAvaliacaoRelatorio_(
      aba,
      linha,
      avaliacoes,
      cadastroPorId
    );

  CONFIG_RELATORIOS_SIGAF
    .DESFECHOS
    .forEach(function (desfecho) {
      linha =
        montarTabelaDesfechoRelatorio_(
          aba,
          linha,
          historico,
          desfecho,
          cadastroPorId
        );
    });

  const conhecidos =
    CONFIG_RELATORIOS_SIGAF
      .DESFECHOS
      .map(
        normalizarTextoRelatorio_
      );

  const outros =
    historico.filter(
      function (registro) {
        return (
          conhecidos.indexOf(
            normalizarTextoRelatorio_(
              registro.desfecho
            )
          ) === -1
        );
      }
    );

  if (outros.length > 0) {
    linha =
      montarTabelaDesfechosOutrosRelatorio_(
        aba,
        linha,
        outros,
        cadastroPorId
      );
  }

  finalizarFormatacaoRelatorio_(
    aba,
    linha
  );
}


/**
 * Limpa somente a aba de saída do relatório.
 */
function prepararAbaRelatorio_(
  aba
) {
  const filtro =
    aba.getFilter();

  if (filtro) {
    filtro.remove();
  }

  const ultimaLinha =
    Math.max(
      aba.getLastRow(),
      1
    );

  /*
   * Abrange todas as 14 colunas utilizadas pelo relatório,
   * garantindo que nenhum intervalo mesclado fique parcial.
   */
  aba
    .getRange(
      1,
      1,
      ultimaLinha,
      CONFIG_RELATORIOS_SIGAF
        .QUANTIDADE_COLUNAS_RELATORIO
    )
    .breakApart();

  aba.clear();

  aba.setFrozenRows(2);
  aba.setHiddenGridlines(true);
}


/**
 * Monta os quatro quadros de indicadores.
 */
function montarQuadrosResumoRelatorio_(
  aba,
  resumo,
  historico
) {
  escreverQuadroRelatorio_(
    aba,
    4,
    1,
    'RESUMO GERAL',
    [
      [
        'Avaliações marcadas',
        resumo.avaliacoesMarcadas
      ],
      [
        'Avaliações comparecidas',
        resumo.avaliacoesComparecidas
      ],
      [
        'Sessões marcadas',
        resumo.sessoesMarcadas
      ],
      [
        'Sessões comparecidas',
        resumo.sessoesComparecidas
      ],
      [
        'Total marcado',
        resumo.totalMarcados
      ],
      [
        'Total realizado',
        resumo.totalRealizados
      ]
    ]
  );

  escreverQuadroRelatorio_(
    aba,
    4,
    4,
    'AVALIAÇÕES',
    [
      [
        'Compareceu',
        resumo.avaliacoesComparecidas
      ],
      [
        'Falta Justificada',
        resumo
          .avaliacoesFaltaJustificada
      ],
      [
        'Falta Não Justificada',
        resumo
          .avaliacoesFaltaNaoJustificada
      ],
      [
        'Cancelada pela clínica',
        resumo
          .avaliacoesCanceladasClinica
      ],
      [
        'Sem resultado — passado',
        resumo
          .avaliacoesSemResultado
      ],
      [
        'Agendada — futura',
        resumo.avaliacoesFuturas
      ]
    ]
  );

  escreverQuadroRelatorio_(
    aba,
    4,
    7,
    'SESSÕES',
    [
      [
        'Compareceu',
        resumo.sessoesComparecidas
      ],
      [
        'Falta Justificada',
        resumo
          .sessoesFaltaJustificada
      ],
      [
        'Falta Não Justificada',
        resumo
          .sessoesFaltaNaoJustificada
      ],
      [
        'Cancelada pela clínica',
        resumo
          .sessoesCanceladasClinica
      ],
      [
        'Remanejadas',
        resumo
          .sessoesRemanejadas
      ],
      [
        'Outros cancelamentos',
        resumo
          .sessoesOutrosCancelamentos
      ],
      [
        'Sem resultado — passado',
        resumo.sessoesSemResultado
      ],
      [
        'Agendada — futura',
        resumo.sessoesFuturas
      ]
    ]
  );

  const totaisDesfechos =
    calcularTotaisDesfechosRelatorio_(
      historico
    );

  escreverQuadroRelatorio_(
    aba,
    4,
    10,
    'DESFECHOS',
    [
      [
        'Alta',
        totaisDesfechos.alta || 0
      ],
      [
        'Encaminhamento para APS',
        totaisDesfechos.aps || 0
      ],
      [
        'Renovação',
        totaisDesfechos.renovacao || 0
      ],
      [
        'Alta por abandono',
        totaisDesfechos.abandono || 0
      ],
      [
        'Desistência',
        totaisDesfechos.desistencia || 0
      ],
      [
        'Total de desfechos',
        resumo.totalDesfechos
      ]
    ]
  );
}


/**
 * Formata um quadro de indicadores.
 */
function escreverQuadroRelatorio_(
  aba,
  linha,
  coluna,
  titulo,
  dados
) {
  aba
    .getRange(
      linha,
      coluna,
      1,
      2
    )
    .merge()
    .setValue(titulo)
    .setBackground('#d9ead3')
    .setFontWeight('bold')
    .setHorizontalAlignment(
      'center'
    );

  aba
    .getRange(
      linha + 1,
      coluna,
      dados.length,
      2
    )
    .setValues(dados)
    .setBorder(
      true,
      true,
      true,
      true,
      true,
      true,
      '#c9d3dd',
      SpreadsheetApp.BorderStyle.SOLID
    );

  aba
    .getRange(
      linha + 1,
      coluna + 1,
      dados.length,
      1
    )
    .setNumberFormat('0')
    .setHorizontalAlignment(
      'center'
    );
}


/**
 * Conta cada categoria de desfecho.
 */
function calcularTotaisDesfechosRelatorio_(
  historico
) {
  const totais = {
    alta: 0,
    aps: 0,
    renovacao: 0,
    abandono: 0,
    desistencia: 0
  };

  historico.forEach(
    function (registro) {
      const texto =
        normalizarTextoRelatorio_(
          registro.desfecho
        );

      if (texto === 'alta') {
        totais.alta++;
      }

      if (
        texto ===
          'encaminhamento para aps'
      ) {
        totais.aps++;
      }

      if (texto === 'renovacao') {
        totais.renovacao++;
      }

      if (
        texto ===
          'alta por abandono'
      ) {
        totais.abandono++;
      }

      if (
        texto ===
          'desistencia do tratamento'
      ) {
        totais.desistencia++;
      }
    }
  );

  return totais;
}
/**
 * Monta a tabela de produção por fisioterapeuta.
 */
function montarTabelaProducaoRelatorio_(
  aba,
  linha,
  producao
) {
  linha =
    escreverTituloSecaoRelatorio_(
      aba,
      linha,
      'PRODUÇÃO POR FISIOTERAPEUTA'
    );

  const cabecalhos = [[
    'Fisioterapeuta',
    'Avaliações realizadas',
    'Sessões realizadas',
    'Total realizado'
  ]];

  const valores =
    producao.map(function (item) {
      return [
        item.fisioterapeuta,
        item.avaliacoes,
        item.sessoes,
        item.total
      ];
    });

  return escreverTabelaRelatorio_(
    aba,
    linha,
    cabecalhos,
    valores,
    4
  );
}


/**
 * Monta a tabela detalhada de avaliações.
 */
function montarTabelaAvaliacoesRelatorio_(
  aba,
  linha,
  avaliacoes
) {
  linha =
    escreverTituloSecaoRelatorio_(
      aba,
      linha,
      'AVALIAÇÕES DO MÊS'
    );

  const cabecalhos = [[
    'Data',
    'Horário',
    'ID Agendamento',
    'Prontuário',
    'Paciente',
    'Resultado',
    'Fisioterapeuta',
    'Motivo'
  ]];

  const valores =
    avaliacoes.map(
      function (registro) {
        return [
          registro.data,
          registro.horario,
          registro.id,
          registro.prontuario,
          registro.paciente,
          registro.status,
          registro.fisioterapeuta,
          registro.motivo
        ];
      }
    );

  return escreverTabelaRelatorio_(
    aba,
    linha,
    cabecalhos,
    valores,
    8
  );
}


/**
 * Monta a tabela detalhada de sessões.
 */
function montarTabelaSessoesRelatorio_(
  aba,
  linha,
  sessoes
) {
  linha =
    escreverTituloSecaoRelatorio_(
      aba,
      linha,
      'SESSÕES DO MÊS'
    );

  const cabecalhos = [[
    'Data',
    'Horário',
    'ID Agendamento',
    'Prontuário',
    'Paciente',
    'Sessão',
    'Resultado',
    'Fisioterapeuta',
    'Faturável',
    'Motivo'
  ]];

  const valores =
    sessoes.map(
      function (registro) {
        return [
          registro.data,
          registro.horario,
          registro.id,
          registro.prontuario,
          registro.paciente,

          registro.numeroSessao
            ? (
                registro.numeroSessao +
                '/' +
                registro.totalPrescrito
              )
            : '',

          registro.status,
          registro.fisioterapeuta,
          registro.faturavel,
          registro.motivo
        ];
      }
    );

  return escreverTabelaRelatorio_(
    aba,
    linha,
    cabecalhos,
    valores,
    10
  );
}


/**
 * Monta uma tabela separada com faltas
 * ocorridas em avaliações.
 */
function montarTabelaFaltasAvaliacaoRelatorio_(
  aba,
  linha,
  avaliacoes,
  cadastroPorId
) {
  const faltas =
    avaliacoes.filter(
      function (registro) {
        const status =
          normalizarTextoRelatorio_(
            registro.status
          );

        return (
          status ===
            'falta justificada' ||
          status ===
            'falta nao justificada'
        );
      }
    );

  linha =
    escreverTituloSecaoRelatorio_(
      aba,
      linha,
      'FALTAS EM AVALIAÇÕES'
    );

  const cabecalhos = [[
    'Data',
    'Prontuário',
    'Paciente',
    'Telefone',
    'Tipo da falta',
    'Fisioterapeuta'
  ]];

  const valores =
    faltas.map(function (registro) {
      const cadastro =
        cadastroPorId[
          normalizarTextoRelatorio_(
            registro.idPaciente
          )
        ] || {};

      return [
        registro.data,
        registro.prontuario,
        registro.paciente,
        cadastro.telefone || '',
        registro.status,
        registro.fisioterapeuta
      ];
    });

  return escreverTabelaRelatorio_(
    aba,
    linha,
    cabecalhos,
    valores,
    6
  );
}


/**
 * Monta uma tabela para cada desfecho.
 */
function montarTabelaDesfechoRelatorio_(
  aba,
  linha,
  historico,
  desfecho,
  cadastroPorId
) {
  const registros =
    historico.filter(
      function (registro) {
        return (
          normalizarTextoRelatorio_(
            registro.desfecho
          ) ===
          normalizarTextoRelatorio_(
            desfecho
          )
        );
      }
    );

  linha =
    escreverTituloSecaoRelatorio_(
      aba,
      linha,
      desfecho.toUpperCase()
    );

  const cabecalhos = [[
    'Data',
    'Prontuário',
    'Paciente',
    'Telefone',
    'ID Ciclo',
    'Ciclo Nº',
    'Prescritas',
    'Realizadas',
    'Restantes',
    'Futuras canceladas',
    'Motivo'
  ]];

  const valores =
    registros.map(
      function (registro) {
        const cadastro =
          cadastroPorId[
            normalizarTextoRelatorio_(
              registro.idPaciente
            )
          ] || {};

        return [
          registro.data,
          registro.prontuario,
          registro.paciente,
          cadastro.telefone || '',
          registro.idCiclo,
          registro.cicloNumero || '',
          registro.prescritas,
          registro.realizadas,
          registro.restantes,
          registro.futurasCanceladas,
          registro.motivo
        ];
      }
    );

  return escreverTabelaRelatorio_(
    aba,
    linha,
    cabecalhos,
    valores,
    11
  );
}


/**
 * Apresenta desfechos que não pertençam
 * às categorias previamente definidas.
 */
function montarTabelaDesfechosOutrosRelatorio_(
  aba,
  linha,
  registros,
  cadastroPorId
) {
  linha =
    escreverTituloSecaoRelatorio_(
      aba,
      linha,
      'OUTROS DESFECHOS'
    );

  const cabecalhos = [[
    'Data',
    'Desfecho',
    'Prontuário',
    'Paciente',
    'Telefone',
    'ID Ciclo',
    'Ciclo Nº',
    'Prescritas',
    'Realizadas',
    'Restantes',
    'Motivo'
  ]];

  const valores =
    registros.map(
      function (registro) {
        const cadastro =
          cadastroPorId[
            normalizarTextoRelatorio_(
              registro.idPaciente
            )
          ] || {};

        return [
          registro.data,
          registro.desfecho,
          registro.prontuario,
          registro.paciente,
          cadastro.telefone || '',
          registro.idCiclo,
          registro.cicloNumero || '',
          registro.prescritas,
          registro.realizadas,
          registro.restantes,
          registro.motivo
        ];
      }
    );

  return escreverTabelaRelatorio_(
    aba,
    linha,
    cabecalhos,
    valores,
    11
  );
}


/**
 * Cria o título azul de cada seção do relatório.
 */
function escreverTituloSecaoRelatorio_(
  aba,
  linha,
  titulo
) {
  aba
    .getRange(
      linha,
      1,
      1,
      CONFIG_RELATORIOS_SIGAF
        .QUANTIDADE_COLUNAS_RELATORIO
    )
    .merge()
    .setValue(titulo)
    .setBackground('#4f81bd')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment(
      'left'
    );

  return linha + 1;
}
/**
 * Escreve e formata uma tabela do relatório.
 */
function escreverTabelaRelatorio_(
  aba,
  linha,
  cabecalhos,
  valores,
  quantidadeColunas
) {
  aba
    .getRange(
      linha,
      1,
      1,
      quantidadeColunas
    )
    .setValues(cabecalhos)
    .setBackground('#d9ead3')
    .setFontWeight('bold')
    .setHorizontalAlignment(
      'center'
    )
    .setVerticalAlignment(
      'middle'
    )
    .setWrap(true);

  const linhaDados =
    linha + 1;

  if (valores.length === 0) {
    aba
      .getRange(
        linhaDados,
        1,
        1,
        quantidadeColunas
      )
      .merge()
      .setValue(
        'Nenhum registro encontrado nesta competência.'
      )
      .setFontColor('#666666')
      .setFontStyle('italic');

    return linhaDados + 3;
  }

  aba
    .getRange(
      linhaDados,
      1,
      valores.length,
      quantidadeColunas
    )
    .setValues(valores)
    .setVerticalAlignment(
      'middle'
    )
    .setWrap(true);

  aba
    .getRange(
      linha,
      1,
      valores.length + 1,
      quantidadeColunas
    )
    .setBorder(
      true,
      true,
      true,
      true,
      true,
      true,
      '#c9d3dd',
      SpreadsheetApp.BorderStyle.SOLID
    );

  /*
   * A primeira coluna das tabelas detalhadas
   * contém a data.
   */
  if (
    valores[0][0] instanceof Date
  ) {
    aba
      .getRange(
        linhaDados,
        1,
        valores.length,
        1
      )
      .setNumberFormat(
        'dd/MM/yyyy'
      );
  }

  /*
   * A segunda coluna das tabelas de atendimentos
   * contém o horário.
   */
  if (
    quantidadeColunas >= 8 &&
    valores[0][1] instanceof Date
  ) {
    aba
      .getRange(
        linhaDados,
        2,
        valores.length,
        1
      )
      .setNumberFormat(
        'HH:mm'
      );
  }

  return (
    linhaDados +
    valores.length +
    2
  );
}


/**
 * Aplica a formatação final da aba.
 */
function finalizarFormatacaoRelatorio_(
  aba,
  ultimaLinha
) {
  const colunas = [
    105,
    95,
    135,
    135,
    210,
    110,
    145,
    180,
    105,
    120,
    120,
    120,
    140,
    180
  ];

  colunas.forEach(
    function (largura, indice) {
      aba.setColumnWidth(
        indice + 1,
        largura
      );
    }
  );

  aba
    .getRange(
      1,
      1,
      Math.max(
        ultimaLinha,
        2
      ),
      CONFIG_RELATORIOS_SIGAF
        .QUANTIDADE_COLUNAS_RELATORIO
    )
    .setFontFamily('Arial')
    .setVerticalAlignment(
      'middle'
    );
}


/**
 * Interpreta o mês informado pelo usuário.
 */
function interpretarCompetenciaRelatorio_(
  texto
) {
  const valor =
    String(texto || '')
      .trim();

  if (!valor) {
    const hoje = new Date();

    return {
      mes: hoje.getMonth(),
      ano: hoje.getFullYear()
    };
  }

  const resultado =
    valor.match(
      /^(\d{1,2})\/(\d{4})$/
    );

  if (!resultado) {
    throw new Error(
      'Informe a competência no formato MM/AAAA. Exemplo: 07/2026.'
    );
  }

  const mes =
    Number(resultado[1]);

  const ano =
    Number(resultado[2]);

  if (
    mes < 1 ||
    mes > 12 ||
    ano < 2000 ||
    ano > 2100
  ) {
    throw new Error(
      'A competência informada é inválida.'
    );
  }

  return {
    mes: mes - 1,
    ano: ano
  };
}


/**
 * Formata a competência como MM/AAAA.
 */
function formatarCompetenciaRelatorio_(
  competencia
) {
  return (
    String(
      competencia.mes + 1
    ).padStart(2, '0') +
    '/' +
    competencia.ano
  );
}


/**
 * Verifica se uma data pertence à competência.
 */
function dataPertenceCompetenciaRelatorio_(
  data,
  competencia
) {
  return (
    data instanceof Date &&
    data.getFullYear() ===
      competencia.ano &&
    data.getMonth() ===
      competencia.mes
  );
}


/**
 * Ordena os registros por data, horário e paciente.
 */
function compararRegistrosRelatorio_(
  a,
  b
) {
  const dataA =
    a.data instanceof Date
      ? a.data.getTime()
      : 0;

  const dataB =
    b.data instanceof Date
      ? b.data.getTime()
      : 0;

  if (dataA !== dataB) {
    return dataA - dataB;
  }

  const horarioA =
    chaveHorarioRelatorio_(
      a.horario
    );

  const horarioB =
    chaveHorarioRelatorio_(
      b.horario
    );

  if (horarioA !== horarioB) {
    return horarioA.localeCompare(
      horarioB
    );
  }

  return String(
    a.paciente || ''
  ).localeCompare(
    String(
      b.paciente || ''
    ),
    'pt-BR'
  );
}


/**
 * Padroniza horários para ordenação.
 */
function chaveHorarioRelatorio_(
  valor
) {
  if (valor instanceof Date) {
    return Utilities.formatDate(
      valor,
      Session.getScriptTimeZone(),
      'HH:mm'
    );
  }

  if (
    typeof valor === 'number' &&
    Number.isFinite(valor)
  ) {
    const totalMinutos =
      Math.round(
        valor * 24 * 60
      );

    const horas =
      Math.floor(
        totalMinutos / 60
      ) % 24;

    const minutos =
      totalMinutos % 60;

    return (
      String(horas)
        .padStart(2, '0') +
      ':' +
      String(minutos)
        .padStart(2, '0')
    );
  }

  const texto =
    String(valor || '')
      .trim();

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


/**
 * Remove o horário de uma data.
 */
function removerHorarioRelatorio_(
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
 * Retorna o nome do mês.
 */
function obterNomeMesRelatorio_(
  indiceMes
) {
  const meses = [
    'Janeiro',
    'Fevereiro',
    'Março',
    'Abril',
    'Maio',
    'Junho',
    'Julho',
    'Agosto',
    'Setembro',
    'Outubro',
    'Novembro',
    'Dezembro'
  ];

  return meses[indiceMes];
}


/**
 * Padroniza textos para comparação.
 */
function normalizarTextoRelatorio_(
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
