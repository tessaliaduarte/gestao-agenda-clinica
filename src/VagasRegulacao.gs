const CONFIG_VAGAS_REGULACAO = {
  ABAS: {
    VAGAS: 'Vagas para Regulação',
    HORARIOS: 'Horários',
    FISIOTERAPEUTAS: 'Fisioterapeutas',
    FERIADOS: 'Calendário da Prefeitura',
    BLOQUEIOS: 'Bloqueios',
    AGENDAMENTOS: 'Agendamentos'
  },

  VAGAS: {
    DATA: 1,
    DIA: 2,
    HORARIO: 3,
    FISIOTERAPEUTA: 4,
    TURNO: 5,
    SITUACAO: 6
  },

  HORARIOS: {
    HORARIO: 1,
    TURNO: 2,
    PERMITE_AVALIACAO: 3
  },

  FISIOTERAPEUTAS: {
    NOME: 2,
    TURNO: 3
  },

  FERIADOS: {
    DATA: 1,
    ATENDIMENTO: 4
  },

  BLOQUEIOS: {
    DATA: 1,
    HORARIO: 2,
    FISIOTERAPEUTA: 3,
    ABRANGENCIA: 4,
    STATUS: 7
  },

  AGENDAMENTOS: {
    DATA: 7,
    HORARIO: 9,
    FISIOTERAPEUTA: 10,
    STATUS: 16
  },

  QUANTIDADE_COLUNAS_VAGAS: 6,
  QUANTIDADE_COLUNAS_AGENDAMENTOS: 22,

  STATUS_QUE_OCUPAM_HORARIO: [
    'agendado',
    'compareceu',
    'falta justificada',
    'falta nao justificada'
  ]
};


/**
 * Gera ou atualiza as vagas mensais de avaliação
 * que serão enviadas à Regulação.
 *
 * Informe a competência como MM/AAAA.
 * Se deixar o campo vazio, será usado o próximo mês.
 */
function gerarVagasMensaisRegulacao() {
  const ui = SpreadsheetApp.getUi();

  const resposta = ui.prompt(
    'Gerar vagas para a Regulação',
    'Informe a competência no formato MM/AAAA.\n\n' +
      'Se deixar o campo vazio, o sistema utilizará o próximo mês.',
    ui.ButtonSet.OK_CANCEL
  );

  if (
    resposta.getSelectedButton() !==
    ui.Button.OK
  ) {
    return;
  }

  const competencia =
    interpretarCompetenciaVagasRegulacao_(
      resposta.getResponseText()
    );

  const lock = LockService.getDocumentLock();
  let bloqueioObtido = false;

  try {
    lock.waitLock(30000);
    bloqueioObtido = true;

    const ss =
      SpreadsheetApp.getActiveSpreadsheet();

    const resultado =
      gerarVagasCompetenciaRegulacao_(
        ss,
        competencia
      );

    SpreadsheetApp.flush();

    resultado.aba.activate();

    ui.alert(
      'Vagas para a Regulação',
      'Competência: ' +
        formatarCompetenciaVagasRegulacao_(
          competencia
        ) +
        '\n\n' +
        'Vagas disponíveis: ' +
        resultado.quantidadeDisponiveis +
        '\n' +
        'Vagas já agendadas preservadas: ' +
        resultado.quantidadeAgendadas +
        '\n' +
        'Datas sem atendimento ignoradas: ' +
        resultado.quantidadeFeriados +
        '\n' +
        'Horários bloqueados ignorados: ' +
        resultado.quantidadeBloqueios +
        '\n' +
        'Horários ocupados ignorados: ' +
        resultado.quantidadeOcupados +
        '\n\n' +
        'A aba "Vagas para Regulação" foi atualizada.',
      ui.ButtonSet.OK
    );
  } catch (erro) {
    ui.alert(
      'Erro ao gerar vagas',
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
        // O bloqueio será liberado automaticamente.
      }
    }
  }
}


/**
 * Gera as vagas para a competência selecionada.
 */
function gerarVagasCompetenciaRegulacao_(
  ss,
  competencia
) {
  const abas =
    obterAbasVagasRegulacao_(ss);

  const horarios =
    lerHorariosVagasRegulacao_(
      abas.horarios
    );

  const fisioterapeutas =
    lerFisioterapeutasVagasRegulacao_(
      abas.fisioterapeutas
    );

  if (horarios.length === 0) {
    throw new Error(
      'Não existem horários que permitam avaliação na aba "Horários".'
    );
  }

  if (fisioterapeutas.length === 0) {
    throw new Error(
      'Não existem fisioterapeutas cadastrados.'
    );
  }

  const feriados =
    lerFeriadosVagasRegulacao_(
      abas.feriados
    );

  const bloqueios =
    lerBloqueiosVagasRegulacao_(
      abas.bloqueios
    );

  const ocupacoes =
    lerOcupacoesVagasRegulacao_(
      abas.agendamentos
    );

  const registrosExistentes =
    lerRegistrosVagasRegulacao_(
      abas.vagas
    );

  /*
   * Recalcula somente as vagas disponíveis da
   * competência escolhida.
   *
   * Vagas agendadas e outras competências
   * permanecem preservadas.
   */
  const registrosPreservados =
    registrosExistentes.filter(
      function(registro) {
        if (
          !dataPertenceCompetenciaVagasRegulacao_(
            registro.data,
            competencia
          )
        ) {
          return true;
        }

        return (
          normalizarTextoVagasRegulacao_(
            registro.situacao
          ) !== 'disponivel'
        );
      }
    );

  const chavesPreservadas = {};

  registrosPreservados.forEach(
    function(registro) {
      const chave =
        criarChaveVagaRegulacao_(
          registro.data,
          registro.horario,
          registro.fisioterapeuta
        );

      if (chave) {
        chavesPreservadas[chave] = true;
      }
    }
  );

  const primeiraData = new Date(
    competencia.ano,
    competencia.mes,
    1
  );

  const ultimaData = new Date(
    competencia.ano,
    competencia.mes + 1,
    0
  );

  primeiraData.setHours(0, 0, 0, 0);
  ultimaData.setHours(0, 0, 0, 0);

  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0);

  const novasVagas = [];

  const feriadosContados = {};
  const bloqueiosContados = {};
  const ocupacoesContadas = {};

  const dataAtual =
    new Date(primeiraData);

  while (
    dataAtual.getTime() <=
    ultimaData.getTime()
  ) {
    const diaSemana =
      dataAtual.getDay();

    const dataPassada =
      dataAtual.getTime() <
      hoje.getTime();

    if (
      !dataPassada &&
      diaSemana !== 0 &&
      diaSemana !== 6
    ) {
      const chaveData =
        chaveDataVagasRegulacao_(
          dataAtual
        );

      if (feriados[chaveData]) {
        feriadosContados[chaveData] =
          true;
      } else {
        fisioterapeutas.forEach(
          function(fisioterapeuta) {
            horarios.forEach(
              function(horario) {
                if (
                  !turnosCompativeisVagasRegulacao_(
                    fisioterapeuta.turno,
                    horario.turno
                  )
                ) {
                  return;
                }

                const chave =
                  criarChaveVagaRegulacao_(
                    dataAtual,
                    horario.horario,
                    fisioterapeuta.nome
                  );

                if (
                  !chave ||
                  chavesPreservadas[chave]
                ) {
                  return;
                }

                if (
                  estaBloqueadoVagasRegulacao_(
                    bloqueios,
                    dataAtual,
                    horario.horario,
                    fisioterapeuta.nome,
                    horario.turno
                  )
                ) {
                  bloqueiosContados[chave] =
                    true;

                  return;
                }

                if (ocupacoes[chave]) {
                  ocupacoesContadas[chave] =
                    true;

                  return;
                }

                novasVagas.push({
                  data: new Date(dataAtual),

                  dia:
                    obterNomeDiaVagasRegulacao_(
                      dataAtual
                    ),

                  horario:
                    horario.horario,

                  fisioterapeuta:
                    fisioterapeuta.nome,

                  turno:
                    horario.turno ||
                    fisioterapeuta.turno,

                  situacao:
                    'Disponível'
                });

                chavesPreservadas[chave] =
                  true;
              }
            );
          }
        );
      }
    }

    dataAtual.setDate(
      dataAtual.getDate() + 1
    );
  }

  const registrosFinais =
    registrosPreservados.concat(
      novasVagas
    );

  registrosFinais.sort(
    compararVagasRegulacao_
  );

  gravarVagasRegulacao_(
    abas.vagas,
    registrosFinais
  );

  const quantidadeAgendadas =
    registrosPreservados.filter(
      function(registro) {
        return (
          dataPertenceCompetenciaVagasRegulacao_(
            registro.data,
            competencia
          ) &&
          normalizarTextoVagasRegulacao_(
            registro.situacao
          ) === 'agendada'
        );
      }
    ).length;

  return {
    aba: abas.vagas,

    quantidadeDisponiveis:
      novasVagas.length,

    quantidadeAgendadas:
      quantidadeAgendadas,

    quantidadeFeriados:
      Object.keys(
        feriadosContados
      ).length,

    quantidadeBloqueios:
      Object.keys(
        bloqueiosContados
      ).length,

    quantidadeOcupados:
      Object.keys(
        ocupacoesContadas
      ).length
  };
}


/**
 * Obtém e valida as abas necessárias.
 */
function obterAbasVagasRegulacao_(ss) {
  const nomes =
    CONFIG_VAGAS_REGULACAO.ABAS;

  const abas = {
    vagas:
      ss.getSheetByName(
        nomes.VAGAS
      ),

    horarios:
      ss.getSheetByName(
        nomes.HORARIOS
      ),

    fisioterapeutas:
      ss.getSheetByName(
        nomes.FISIOTERAPEUTAS
      ),

    feriados:
      ss.getSheetByName(
        nomes.FERIADOS
      ),

    bloqueios:
      ss.getSheetByName(
        nomes.BLOQUEIOS
      ),

    agendamentos:
      ss.getSheetByName(
        nomes.AGENDAMENTOS
      )
  };

  Object.keys(abas).forEach(
    function(chave) {
      if (!abas[chave]) {
        const mapa = {
          vagas: nomes.VAGAS,
          horarios: nomes.HORARIOS,
          fisioterapeutas:
            nomes.FISIOTERAPEUTAS,
          feriados: nomes.FERIADOS,
          bloqueios: nomes.BLOQUEIOS,
          agendamentos:
            nomes.AGENDAMENTOS
        };

        throw new Error(
          'A aba necessária "' +
            mapa[chave] +
            '" não foi encontrada.'
        );
      }
    }
  );

  return abas;
}


/**
 * Lê os horários que permitem avaliação.
 */
function lerHorariosVagasRegulacao_(
  abaHorarios
) {
  const ultimaLinha =
    abaHorarios.getLastRow();

  if (ultimaLinha < 2) {
    return [];
  }

  const dados =
    abaHorarios.getRange(
      2,
      1,
      ultimaLinha - 1,
      3
    ).getValues();

  return dados
    .map(function(linha) {
      return {
        horario: linha[
          CONFIG_VAGAS_REGULACAO
            .HORARIOS
            .HORARIO - 1
        ],

        turno: String(
          linha[
            CONFIG_VAGAS_REGULACAO
              .HORARIOS
              .TURNO - 1
          ] || ''
        ).trim(),

        permiteAvaliacao:
          normalizarTextoVagasRegulacao_(
            linha[
              CONFIG_VAGAS_REGULACAO
                .HORARIOS
                .PERMITE_AVALIACAO - 1
            ]
          )
      };
    })
    .filter(function(horario) {
      return (
        Boolean(
          chaveHorarioVagasRegulacao_(
            horario.horario
          )
        ) &&
        horario.permiteAvaliacao ===
          'sim'
      );
    });
}


/**
 * Lê os fisioterapeutas e seus turnos.
 */
function lerFisioterapeutasVagasRegulacao_(
  abaFisioterapeutas
) {
  const ultimaLinha =
    abaFisioterapeutas.getLastRow();

  if (ultimaLinha < 2) {
    return [];
  }

  const dados =
    abaFisioterapeutas.getRange(
      2,
      1,
      ultimaLinha - 1,
      3
    ).getValues();

  return dados
    .map(function(linha) {
      return {
        nome: String(
          linha[
            CONFIG_VAGAS_REGULACAO
              .FISIOTERAPEUTAS
              .NOME - 1
          ] || ''
        ).trim(),

        turno: String(
          linha[
            CONFIG_VAGAS_REGULACAO
              .FISIOTERAPEUTAS
              .TURNO - 1
          ] || ''
        ).trim()
      };
    })
    .filter(function(fisioterapeuta) {
      return Boolean(
        fisioterapeuta.nome
      );
    });
}


/**
 * Lê os feriados e dias sem atendimento.
 */
function lerFeriadosVagasRegulacao_(
  abaFeriados
) {
  const resultado = {};

  const ultimaLinha =
    abaFeriados.getLastRow();

  if (ultimaLinha < 2) {
    return resultado;
  }

  const quantidadeColunas =
    Math.max(
      abaFeriados.getLastColumn(),
      4
    );

  const dados =
    abaFeriados.getRange(
      2,
      1,
      ultimaLinha - 1,
      quantidadeColunas
    ).getValues();

  dados.forEach(function(linha) {
    const data = linha[
      CONFIG_VAGAS_REGULACAO
        .FERIADOS
        .DATA - 1
    ];

    const atendimento =
      normalizarTextoVagasRegulacao_(
        linha[
          CONFIG_VAGAS_REGULACAO
            .FERIADOS
            .ATENDIMENTO - 1
        ]
      );

    if (
      data instanceof Date &&
      atendimento === 'nao'
    ) {
      resultado[
        chaveDataVagasRegulacao_(
          data
        )
      ] = true;
    }
  });

  return resultado;
}


/**
 * Lê os bloqueios cadastrados.
 */
function lerBloqueiosVagasRegulacao_(
  abaBloqueios
) {
  const ultimaLinha =
    abaBloqueios.getLastRow();

  if (ultimaLinha < 2) {
    return [];
  }

  const quantidadeColunas =
    Math.max(
      abaBloqueios.getLastColumn(),
      7
    );

  const dados =
    abaBloqueios.getRange(
      2,
      1,
      ultimaLinha - 1,
      quantidadeColunas
    ).getValues();

  return dados
    .map(function(linha) {
      return {
        data: linha[
          CONFIG_VAGAS_REGULACAO
            .BLOQUEIOS
            .DATA - 1
        ],

        horario: linha[
          CONFIG_VAGAS_REGULACAO
            .BLOQUEIOS
            .HORARIO - 1
        ],

        fisioterapeuta: String(
          linha[
            CONFIG_VAGAS_REGULACAO
              .BLOQUEIOS
              .FISIOTERAPEUTA - 1
          ] || ''
        ).trim(),

        abrangencia: String(
          linha[
            CONFIG_VAGAS_REGULACAO
              .BLOQUEIOS
              .ABRANGENCIA - 1
          ] || ''
        ).trim(),

        status: String(
          linha[
            CONFIG_VAGAS_REGULACAO
              .BLOQUEIOS
              .STATUS - 1
          ] || ''
        ).trim()
      };
    })
    .filter(function(bloqueio) {
      if (
        !(
          bloqueio.data
          instanceof Date
        )
      ) {
        return false;
      }

      const status =
        normalizarTextoVagasRegulacao_(
          bloqueio.status
        );

      return (
        !status ||
        status === 'ativo' ||
        status === 'bloqueado'
      );
    });
}


/**
 * Lê os horários já ocupados.
 */
function lerOcupacoesVagasRegulacao_(
  abaAgendamentos
) {
  const resultado = {};

  const ultimaLinha =
    abaAgendamentos.getLastRow();

  if (ultimaLinha < 2) {
    return resultado;
  }

  const dados =
    abaAgendamentos.getRange(
      2,
      1,
      ultimaLinha - 1,
      CONFIG_VAGAS_REGULACAO
        .QUANTIDADE_COLUNAS_AGENDAMENTOS
    ).getValues();

  dados.forEach(function(linha) {
    const data = linha[
      CONFIG_VAGAS_REGULACAO
        .AGENDAMENTOS
        .DATA - 1
    ];

    const horario = linha[
      CONFIG_VAGAS_REGULACAO
        .AGENDAMENTOS
        .HORARIO - 1
    ];

    const fisioterapeuta = linha[
      CONFIG_VAGAS_REGULACAO
        .AGENDAMENTOS
        .FISIOTERAPEUTA - 1
    ];

    const status =
      normalizarTextoVagasRegulacao_(
        linha[
          CONFIG_VAGAS_REGULACAO
            .AGENDAMENTOS
            .STATUS - 1
        ]
      );

    if (
      !(
        data instanceof Date
      ) ||
      CONFIG_VAGAS_REGULACAO
        .STATUS_QUE_OCUPAM_HORARIO
        .indexOf(status) === -1
    ) {
      return;
    }

    const chave =
      criarChaveVagaRegulacao_(
        data,
        horario,
        fisioterapeuta
      );

    if (chave) {
      resultado[chave] = true;
    }
  });

  return resultado;
}


/**
 * Lê as vagas existentes.
 */
function lerRegistrosVagasRegulacao_(
  abaVagas
) {
  const ultimaLinha =
    abaVagas.getLastRow();

  if (ultimaLinha < 2) {
    return [];
  }

  const dados =
    abaVagas.getRange(
      2,
      1,
      ultimaLinha - 1,
      CONFIG_VAGAS_REGULACAO
        .QUANTIDADE_COLUNAS_VAGAS
    ).getValues();

  return dados
    .map(function(linha) {
      return {
        data: linha[
          CONFIG_VAGAS_REGULACAO
            .VAGAS
            .DATA - 1
        ],

        dia: String(
          linha[
            CONFIG_VAGAS_REGULACAO
              .VAGAS
              .DIA - 1
          ] || ''
        ).trim(),

        horario: linha[
          CONFIG_VAGAS_REGULACAO
            .VAGAS
            .HORARIO - 1
        ],

        fisioterapeuta: String(
          linha[
            CONFIG_VAGAS_REGULACAO
              .VAGAS
              .FISIOTERAPEUTA - 1
          ] || ''
        ).trim(),

        turno: String(
          linha[
            CONFIG_VAGAS_REGULACAO
              .VAGAS
              .TURNO - 1
          ] || ''
        ).trim(),

        situacao: String(
          linha[
            CONFIG_VAGAS_REGULACAO
              .VAGAS
              .SITUACAO - 1
          ] || ''
        ).trim()
      };
    })
    .filter(function(registro) {
      return (
        registro.data
          instanceof Date &&
        Boolean(
          chaveHorarioVagasRegulacao_(
            registro.horario
          )
        ) &&
        Boolean(
          registro.fisioterapeuta
        )
      );
    });
}


/**
 * Verifica se o horário está bloqueado.
 */
function estaBloqueadoVagasRegulacao_(
  bloqueios,
  data,
  horario,
  fisioterapeuta,
  turno
) {
  const chaveData =
    chaveDataVagasRegulacao_(data);

  const chaveHorario =
    chaveHorarioVagasRegulacao_(
      horario
    );

  const profissional =
    normalizarTextoVagasRegulacao_(
      fisioterapeuta
    );

  const turnoProcurado =
    normalizarTextoVagasRegulacao_(
      turno
    );

  return bloqueios.some(
    function(bloqueio) {
      if (
        chaveDataVagasRegulacao_(
          bloqueio.data
        ) !== chaveData
      ) {
        return false;
      }

      const profissionalBloqueio =
        normalizarTextoVagasRegulacao_(
          bloqueio.fisioterapeuta
        );

      const bloqueiaProfissional =
        !profissionalBloqueio ||
        profissionalBloqueio ===
          profissional ||
        profissionalBloqueio ===
          'todos';

      if (!bloqueiaProfissional) {
        return false;
      }

      const abrangencia =
        normalizarTextoVagasRegulacao_(
          bloqueio.abrangencia
        );

      if (
        abrangencia ===
          'dia inteiro' ||
        abrangencia === 'dia'
      ) {
        return true;
      }

      if (
        abrangencia ===
          'turno inteiro' ||
        abrangencia === 'turno'
      ) {
        const turnoBloqueio =
          obterTurnoHorarioVagasRegulacao_(
            bloqueio.horario
          );

        return (
          !turnoBloqueio ||
          turnoBloqueio ===
            turnoProcurado
        );
      }

      const horarioBloqueio =
        chaveHorarioVagasRegulacao_(
          bloqueio.horario
        );

      return (
        !horarioBloqueio ||
        horarioBloqueio ===
          chaveHorario
      );
    }
  );
}


/**
 * Grava as vagas preservando a formatação.
 */
function gravarVagasRegulacao_(
  abaVagas,
  registros
) {
  const ultimaLinhaAtual =
    Math.max(
      abaVagas.getLastRow(),
      2
    );

  abaVagas.getRange(
    2,
    1,
    ultimaLinhaAtual - 1,
    CONFIG_VAGAS_REGULACAO
      .QUANTIDADE_COLUNAS_VAGAS
  ).clearContent();

  if (registros.length === 0) {
    return;
  }

  const linhaFinalNecessaria =
    registros.length + 1;

  if (
    abaVagas.getMaxRows() <
    linhaFinalNecessaria
  ) {
    abaVagas.insertRowsAfter(
      abaVagas.getMaxRows(),
      linhaFinalNecessaria -
        abaVagas.getMaxRows()
    );
  }

  const valores =
    registros.map(function(registro) {
      return [
        new Date(registro.data),

        registro.dia ||
          obterNomeDiaVagasRegulacao_(
            registro.data
          ),

        registro.horario,
        registro.fisioterapeuta,
        registro.turno,

        registro.situacao ||
          'Disponível'
      ];
    });

  abaVagas.getRange(
    2,
    1,
    valores.length,
    CONFIG_VAGAS_REGULACAO
      .QUANTIDADE_COLUNAS_VAGAS
  ).setValues(valores);

  abaVagas.getRange(
    2,
    CONFIG_VAGAS_REGULACAO
      .VAGAS
      .DATA,
    valores.length,
    1
  ).setNumberFormat('dd/MM/yyyy');

  abaVagas.getRange(
    2,
    CONFIG_VAGAS_REGULACAO
      .VAGAS
      .HORARIO,
    valores.length,
    1
  ).setNumberFormat('HH:mm');
}


/**
 * Interpreta a competência informada.
 */
function interpretarCompetenciaVagasRegulacao_(
  texto
) {
  const valor = String(
    texto || ''
  ).trim();

  if (!valor) {
    const proximoMes = new Date();

    proximoMes.setDate(1);

    proximoMes.setMonth(
      proximoMes.getMonth() + 1
    );

    return {
      mes: proximoMes.getMonth(),
      ano: proximoMes.getFullYear()
    };
  }

  const resultado = valor.match(
    /^(\d{1,2})\/(\d{4})$/
  );

  if (!resultado) {
    throw new Error(
      'Informe a competência no formato MM/AAAA. Exemplo: 08/2026.'
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


function formatarCompetenciaVagasRegulacao_(
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


function dataPertenceCompetenciaVagasRegulacao_(
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


function turnosCompativeisVagasRegulacao_(
  turnoFisioterapeuta,
  turnoHorario
) {
  const profissional =
    normalizarTextoVagasRegulacao_(
      turnoFisioterapeuta
    );

  const horario =
    normalizarTextoVagasRegulacao_(
      turnoHorario
    );

  return (
    !profissional ||
    !horario ||
    profissional === horario
  );
}


function criarChaveVagaRegulacao_(
  data,
  horario,
  fisioterapeuta
) {
  if (
    !(
      data instanceof Date
    )
  ) {
    return '';
  }

  const chaveHorario =
    chaveHorarioVagasRegulacao_(
      horario
    );

  const profissional =
    normalizarTextoVagasRegulacao_(
      fisioterapeuta
    );

  if (
    !chaveHorario ||
    !profissional
  ) {
    return '';
  }

  return (
    chaveDataVagasRegulacao_(
      data
    ) +
    '|' +
    chaveHorario +
    '|' +
    profissional
  );
}


function compararVagasRegulacao_(
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
    chaveHorarioVagasRegulacao_(
      a.horario
    );

  const horarioB =
    chaveHorarioVagasRegulacao_(
      b.horario
    );

  if (horarioA !== horarioB) {
    return horarioA.localeCompare(
      horarioB
    );
  }

  return String(
    a.fisioterapeuta || ''
  ).localeCompare(
    String(
      b.fisioterapeuta || ''
    ),
    'pt-BR'
  );
}


function obterNomeDiaVagasRegulacao_(
  data
) {
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


function chaveDataVagasRegulacao_(
  data
) {
  return Utilities.formatDate(
    new Date(data),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );
}


function chaveHorarioVagasRegulacao_(
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

  const texto = String(
    valor || ''
  ).trim();

  const resultado = texto.match(
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


function obterTurnoHorarioVagasRegulacao_(
  horario
) {
  const chave =
    chaveHorarioVagasRegulacao_(
      horario
    );

  if (!chave) {
    return '';
  }

  const hora =
    Number(
      chave.split(':')[0]
    );

  return hora < 13
    ? 'manha'
    : 'tarde';
}


function normalizarTextoVagasRegulacao_(
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
