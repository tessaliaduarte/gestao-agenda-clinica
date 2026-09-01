const CONFIG_BLOQUEIOS_AUTOMATICOS = {
  ABAS: {
    BLOQUEIOS: 'Bloqueios',
    AGENDAMENTOS: 'Agendamentos',
    CADASTRO: 'Cadastro de Pacientes',
    VAGAS_REGULACAO: 'Vagas para Regulação',
    STATUS_SESSAO: 'Status da Sessão',
    AGENDA: 'Agenda'
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
    DATA_AVALIACAO: 17,
    HORARIO_AVALIACAO: 18,
    STATUS: 21,
    FISIOTERAPEUTA: 22
  },

  VAGAS_REGULACAO: {
    DATA: 1,
    DIA: 2,
    HORARIO: 3,
    FISIOTERAPEUTA: 4,
    TURNO: 5,
    SITUACAO: 6
  },

  QUANTIDADE_COLUNAS_AGENDAMENTOS: 22
};


/**
 * Ativa o processamento automático dos bloqueios.
 *
 * Execute manualmente somente uma vez após colar o arquivo.
 */
function instalarGatilhoBloqueiosAutomaticos() {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const ui = SpreadsheetApp.getUi();

  try {
    removerGatilhosBloqueiosAutomaticos_();

    ScriptApp
      .newTrigger(
        'processarEdicaoBloqueiosAutomaticos'
      )
      .forSpreadsheet(ss)
      .onEdit()
      .create();

    removerCancelamentoClinicaDaListaStatus_(
      ss
    );

    const resultado =
      processarTodosBloqueiosAtivosInterno_(
        ss
      );

    atualizarModulosAposBloqueio_(
      ss,
      resultado.datasAfetadas
    );

    ui.alert(
      'Bloqueios automáticos ativados',
      'O fluxo foi configurado com sucesso.\n\n' +
        'Atendimentos processados agora: ' +
        resultado.total,
      ui.ButtonSet.OK
    );
  } catch (erro) {
    ui.alert(
      'Erro ao ativar bloqueios automáticos',
      erro && erro.message
        ? erro.message
        : String(erro),
      ui.ButtonSet.OK
    );

    throw erro;
  }
}


/**
 * Permite processar novamente todos os bloqueios ativos.
 * É seguro executar mais de uma vez, pois somente registros
 * ainda Agendados são processados.
 */
function processarBloqueiosAtivosAgora() {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const ui = SpreadsheetApp.getUi();

  try {
    const resultado =
      processarTodosBloqueiosAtivosInterno_(
        ss
      );

    atualizarModulosAposBloqueio_(
      ss,
      resultado.datasAfetadas
    );

    ui.alert(
      'Bloqueios processados',
      'Atendimentos processados: ' +
        resultado.total,
      ui.ButtonSet.OK
    );
  } catch (erro) {
    ui.alert(
      'Erro ao processar bloqueios',
      erro && erro.message
        ? erro.message
        : String(erro),
      ui.ButtonSet.OK
    );

    throw erro;
  }
}


/**
 * Gatilho executado quando o Status do bloqueio é alterado.
 */
function processarEdicaoBloqueiosAutomaticos(e) {
  if (!e || !e.range) {
    return;
  }

  const aba = e.range.getSheet();

  if (
    aba.getName() !==
      CONFIG_BLOQUEIOS_AUTOMATICOS
        .ABAS.BLOQUEIOS ||
    e.range.getRow() < 2 ||
    e.range.getColumn() !==
      CONFIG_BLOQUEIOS_AUTOMATICOS
        .BLOQUEIOS.STATUS
  ) {
    return;
  }

  const status =
    normalizarTextoBloqueioAutomatico_(
      e.value ||
      e.range.getValue()
    );

  if (
    status !== 'ativo' &&
    status !== 'bloqueado'
  ) {
    return;
  }

  const ss =
    e.source ||
    SpreadsheetApp.getActiveSpreadsheet();

  try {
    const resultado =
      processarLinhaBloqueioComLock_(
        ss,
        e.range.getRow()
      );

    atualizarModulosAposBloqueio_(
      ss,
      resultado.datasAfetadas
    );

    ss.toast(
      resultado.total +
        ' atendimento(s) processado(s).',
      'Bloqueio aplicado',
      7
    );
  } catch (erro) {
    console.error(erro);

    ss.toast(
      erro && erro.message
        ? erro.message
        : String(erro),
      'Erro ao aplicar bloqueio',
      9
    );
  }
}


function removerGatilhosBloqueiosAutomaticos_() {
  ScriptApp
    .getProjectTriggers()
    .forEach(function (gatilho) {
      if (
        gatilho.getHandlerFunction() ===
          'processarEdicaoBloqueiosAutomaticos'
      ) {
        ScriptApp.deleteTrigger(gatilho);
      }
    });
}


function processarTodosBloqueiosAtivosInterno_(
  ss
) {
  const lock =
    LockService.getDocumentLock();

  let bloqueioObtido = false;

  try {
    lock.waitLock(30000);
    bloqueioObtido = true;

    const abas =
      obterAbasBloqueiosAutomaticos_(ss);

    const ultimaLinha =
      abas.bloqueios.getLastRow();

    let total = 0;
    const datasAfetadas = {};

    if (ultimaLinha < 2) {
      return {
        total: 0,
        datasAfetadas: []
      };
    }

    const status = abas.bloqueios
      .getRange(
        2,
        CONFIG_BLOQUEIOS_AUTOMATICOS
          .BLOQUEIOS.STATUS,
        ultimaLinha - 1,
        1
      )
      .getDisplayValues();

    status.forEach(
      function (linha, indice) {
        const valor =
          normalizarTextoBloqueioAutomatico_(
            linha[0]
          );

        if (
          valor !== 'ativo' &&
          valor !== 'bloqueado'
        ) {
          return;
        }

        const resultado =
          processarLinhaBloqueioInterno_(
            ss,
            abas,
            indice + 2
          );

        total += resultado.total;

        resultado.datasAfetadas.forEach(
          function (data) {
            datasAfetadas[
              chaveDataBloqueioAutomatico_(
                data
              )
            ] = data;
          }
        );
      }
    );

    SpreadsheetApp.flush();

    return {
      total: total,
      datasAfetadas:
        Object.keys(datasAfetadas)
          .map(function (chave) {
            return datasAfetadas[chave];
          })
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


function processarLinhaBloqueioComLock_(
  ss,
  linhaBloqueio
) {
  const lock =
    LockService.getDocumentLock();

  let bloqueioObtido = false;

  try {
    lock.waitLock(30000);
    bloqueioObtido = true;

    const abas =
      obterAbasBloqueiosAutomaticos_(ss);

    const resultado =
      processarLinhaBloqueioInterno_(
        ss,
        abas,
        linhaBloqueio
      );

    SpreadsheetApp.flush();

    return resultado;
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


function obterAbasBloqueiosAutomaticos_(ss) {
  const nomes =
    CONFIG_BLOQUEIOS_AUTOMATICOS.ABAS;

  const abas = {
    bloqueios:
      ss.getSheetByName(
        nomes.BLOQUEIOS
      ),

    agendamentos:
      ss.getSheetByName(
        nomes.AGENDAMENTOS
      ),

    cadastro:
      ss.getSheetByName(
        nomes.CADASTRO
      ),

    vagasRegulacao:
      ss.getSheetByName(
        nomes.VAGAS_REGULACAO
      ),

    statusSessao:
      ss.getSheetByName(
        nomes.STATUS_SESSAO
      ),

    agenda:
      ss.getSheetByName(
        nomes.AGENDA
      )
  };

  Object.keys(abas).forEach(
    function (chave) {
      if (!abas[chave]) {
        throw new Error(
          'Uma aba necessária ao fluxo de bloqueios não foi encontrada: ' +
            chave
        );
      }
    }
  );

  return abas;
}
function processarLinhaBloqueioInterno_(
  ss,
  abas,
  linhaBloqueio
) {
  const bloqueio =
    lerBloqueioAutomatico_(
      abas.bloqueios,
      linhaBloqueio
    );

  validarBloqueioAutomatico_(
    bloqueio
  );

  const afetados =
    localizarAgendamentosAfetadosBloqueio_(
      abas.agendamentos,
      bloqueio
    );

  const datasAfetadas = [
    bloqueio.data
  ];

  let sessoes = 0;
  let avaliacoes = 0;
  let sessoesSemReposicao = 0;
  let avaliacoesSemVaga = 0;
  const erros = [];

  afetados.forEach(
    function (agendamento) {
      const evento =
        normalizarTextoBloqueioAutomatico_(
          agendamento.evento
        );

      try {
        if (evento === 'sessao') {
          if (
            typeof
              gerarReposicaoCancelamentoClinica_ !==
              'function'
          ) {
            throw new Error(
              'A função de reposição não foi encontrada. Confira o arquivo ReposicaoCancelamento.gs.'
            );
          }

          const reposicao =
            gerarReposicaoCancelamentoClinica_(
              ss,
              agendamento.idAgendamento
            );

          sessoes++;

          if (
            reposicao &&
            reposicao.semReposicao
          ) {
            sessoesSemReposicao++;
          }

          if (
            reposicao &&
            reposicao.data instanceof Date
          ) {
            datasAfetadas.push(
              reposicao.data
            );
          }

          return;
        }

        if (evento === 'avaliacao') {
          const reagendamento =
            reagendarAvaliacaoBloqueada_(
              ss,
              abas,
              agendamento,
              bloqueio
            );

          avaliacoes++;

          if (!reagendamento) {
            avaliacoesSemVaga++;
          }

          if (
            reagendamento &&
            reagendamento.data instanceof Date
          ) {
            datasAfetadas.push(
              reagendamento.data
            );
          }
        }
      } catch (erroAtendimento) {
        erros.push(
          agendamento.idAgendamento +
          ': ' +
          (
            erroAtendimento &&
            erroAtendimento.message
              ? erroAtendimento.message
              : String(erroAtendimento)
          )
        );

        console.error(
          erroAtendimento
        );
      }
    }
  );

  let nota =
    'Processamento automático concluído em ' +
    Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'dd/MM/yyyy HH:mm'
    ) +
    '. Sessões processadas: ' +
    sessoes +
    '. Avaliações processadas: ' +
    avaliacoes +
    '. Sessões sem reposição automática: ' +
    sessoesSemReposicao +
    '. Avaliações sem nova vaga: ' +
    avaliacoesSemVaga +
    '. Erros: ' +
    erros.length +
    '.';

  if (erros.length > 0) {
    nota +=
      '\n\nDetalhes dos erros:\n' +
      erros.join('\n');
  }

  abas.bloqueios
    .getRange(
      linhaBloqueio,
      CONFIG_BLOQUEIOS_AUTOMATICOS
        .BLOQUEIOS.ACAO
    )
    .setNote(
      nota
    );

  return {
    total:
      sessoes + avaliacoes,
    sessoes: sessoes,
    avaliacoes: avaliacoes,
    sessoesSemReposicao:
      sessoesSemReposicao,
    avaliacoesSemVaga:
      avaliacoesSemVaga,
    erros: erros.length,
    datasAfetadas: datasAfetadas
  };
}


function lerBloqueioAutomatico_(
  abaBloqueios,
  linha
) {
  const valores =
    abaBloqueios
      .getRange(
        linha,
        1,
        1,
        7
      )
      .getValues()[0];

  const colunas =
    CONFIG_BLOQUEIOS_AUTOMATICOS
      .BLOQUEIOS;

  return {
    linha: linha,
    data:
      valores[colunas.DATA - 1],
    horario:
      valores[colunas.HORARIO - 1],
    fisioterapeuta: String(
      valores[
        colunas.FISIOTERAPEUTA - 1
      ] || ''
    ).trim(),
    abrangencia: String(
      valores[
        colunas.ABRANGENCIA - 1
      ] || ''
    ).trim(),
    motivo: String(
      valores[colunas.MOTIVO - 1] ||
        'Bloqueio da clínica'
    ).trim(),
    acao: String(
      valores[colunas.ACAO - 1] || ''
    ).trim(),
    status: String(
      valores[colunas.STATUS - 1] || ''
    ).trim()
  };
}


function validarBloqueioAutomatico_(
  bloqueio
) {
  if (!(bloqueio.data instanceof Date)) {
    throw new Error(
      'Informe uma data válida no bloqueio.'
    );
  }

  const status =
    normalizarTextoBloqueioAutomatico_(
      bloqueio.status
    );

  if (
    status !== 'ativo' &&
    status !== 'bloqueado'
  ) {
    throw new Error(
      'O Status do bloqueio deve ser Ativo.'
    );
  }

  const abrangencia =
    normalizarTextoBloqueioAutomatico_(
      bloqueio.abrangencia
    );

  if (
    abrangencia !== 'dia inteiro' &&
    !chaveHorarioBloqueioAutomatico_(
      bloqueio.horario
    )
  ) {
    throw new Error(
      'Informe o horário de referência do bloqueio.'
    );
  }
}


function localizarAgendamentosAfetadosBloqueio_(
  abaAgendamentos,
  bloqueio
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
      CONFIG_BLOQUEIOS_AUTOMATICOS
        .QUANTIDADE_COLUNAS_AGENDAMENTOS
    )
    .getValues();

  const colunas =
    CONFIG_BLOQUEIOS_AUTOMATICOS
      .AGENDAMENTOS;

  return dados
    .map(function (linha, indice) {
      return {
        linhaPlanilha:
          indice + 2,
        valores: linha,
        idAgendamento: String(
          linha[
            colunas.ID_AGENDAMENTO - 1
          ] || ''
        ).trim(),
        idPaciente: String(
          linha[
            colunas.ID_PACIENTE - 1
          ] || ''
        ).trim(),
        data:
          linha[colunas.DATA - 1],
        horario:
          linha[colunas.HORARIO - 1],
        fisioterapeuta: String(
          linha[
            colunas.FISIOTERAPEUTA - 1
          ] || ''
        ).trim(),
        evento: String(
          linha[colunas.EVENTO - 1] || ''
        ).trim(),
        status: String(
          linha[colunas.STATUS - 1] || ''
        ).trim()
      };
    })
    .filter(function (agendamento) {
      const evento =
        normalizarTextoBloqueioAutomatico_(
          agendamento.evento
        );

      return (
        agendamento.idAgendamento &&
        (
          evento === 'sessao' ||
          evento === 'avaliacao'
        ) &&
        normalizarTextoBloqueioAutomatico_(
          agendamento.status
        ) === 'agendado' &&
        bloqueioAfetaAtendimentoAutomatico_(
          bloqueio,
          agendamento
        )
      );
    });
}


function bloqueioAfetaAtendimentoAutomatico_(
  bloqueio,
  agendamento
) {
  if (
    !(agendamento.data instanceof Date) ||
    chaveDataBloqueioAutomatico_(
      bloqueio.data
    ) !==
      chaveDataBloqueioAutomatico_(
        agendamento.data
      )
  ) {
    return false;
  }

  const profissionalBloqueio =
    normalizarTextoBloqueioAutomatico_(
      bloqueio.fisioterapeuta
    );

  const profissionalAgendamento =
    normalizarTextoBloqueioAutomatico_(
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
    normalizarTextoBloqueioAutomatico_(
      bloqueio.abrangencia
    );

  if (abrangencia === 'dia inteiro') {
    return true;
  }

  if (abrangencia === 'turno inteiro') {
    return (
      obterTurnoBloqueioAutomatico_(
        bloqueio.horario
      ) ===
      obterTurnoBloqueioAutomatico_(
        agendamento.horario
      )
    );
  }

  return (
    chaveHorarioBloqueioAutomatico_(
      bloqueio.horario
    ) ===
    chaveHorarioBloqueioAutomatico_(
      agendamento.horario
    )
  );
}
function reagendarAvaliacaoBloqueada_(
  ss,
  abas,
  agendamento,
  bloqueio
) {
  const vaga =
    localizarProximaVagaAvaliacao_(
      abas,
      agendamento
    );

  const agora = new Date();

  marcarVagaOriginalAvaliacaoBloqueada_(
    abas.vagasRegulacao,
    agendamento
  );

  if (!vaga) {
    atualizarAvaliacaoOriginalBloqueada_(
      abas.agendamentos,
      agendamento,
      '',
      bloqueio,
      agora
    );

    atualizarCadastroAvaliacaoSemVaga_(
      abas.cadastro,
      agendamento.idPaciente,
      agendamento,
      bloqueio
    );

    registrarPendenciaReagendamentoAvaliacao_(
      ss,
      abas.cadastro,
      agendamento,
      bloqueio
    );

    return null;
  }

  const novoId =
    gerarProximoIdAgendamentoBloqueio_(
      abas.agendamentos
    );

  atualizarAvaliacaoOriginalBloqueada_(
    abas.agendamentos,
    agendamento,
    novoId,
    bloqueio,
    agora
  );

  const novaLinha =
    agendamento.valores.slice();

  const colunas =
    CONFIG_BLOQUEIOS_AUTOMATICOS
      .AGENDAMENTOS;

  novaLinha[
    colunas.ID_AGENDAMENTO - 1
  ] = novoId;

  novaLinha[
    colunas.DATA - 1
  ] = new Date(vaga.data);

  novaLinha[
    colunas.DIA - 1
  ] =
    vaga.dia ||
    obterNomeDiaBloqueioAutomatico_(
      vaga.data
    );

  novaLinha[
    colunas.HORARIO - 1
  ] = vaga.horario;

  novaLinha[
    colunas.FISIOTERAPEUTA - 1
  ] = vaga.fisioterapeuta;

  novaLinha[
    colunas.STATUS - 1
  ] = 'Agendado';

  novaLinha[
    colunas.MOTIVO - 1
  ] =
    'Reagendamento automático de ' +
    agendamento.idAgendamento +
    ' por bloqueio da clínica: ' +
    bloqueio.motivo;

  novaLinha[
    colunas.CONTA_COMO_SESSAO - 1
  ] = 'Não';

  novaLinha[
    colunas.AVISAR_PACIENTE - 1
  ] = 'Sim';

  novaLinha[
    colunas.CRIADO_EM - 1
  ] = agora;

  novaLinha[
    colunas.ATUALIZADO_EM - 1
  ] = agora;

  novaLinha[
    colunas.FATURAVEL - 1
  ] = 'Não';

  const novaLinhaPlanilha =
    Math.max(
      abas.agendamentos.getLastRow() + 1,
      2
    );

  abas.agendamentos
    .getRange(
      novaLinhaPlanilha,
      1,
      1,
      CONFIG_BLOQUEIOS_AUTOMATICOS
        .QUANTIDADE_COLUNAS_AGENDAMENTOS
    )
    .setValues([novaLinha]);

  formatarLinhaAvaliacaoReagendada_(
    abas.agendamentos,
    novaLinhaPlanilha
  );

  abas.vagasRegulacao
    .getRange(
      vaga.linha,
      CONFIG_BLOQUEIOS_AUTOMATICOS
        .VAGAS_REGULACAO.SITUACAO
    )
    .setValue('Agendada');

  atualizarCadastroAvaliacaoReagendada_(
    abas.cadastro,
    agendamento.idPaciente,
    vaga
  );

  return {
    idAgendamento: novoId,
    data: new Date(vaga.data),
    horario: vaga.horario
  };
}


function marcarVagaOriginalAvaliacaoBloqueada_(
  abaVagas,
  agendamento
) {
  const ultimaLinha =
    abaVagas.getLastRow();

  if (ultimaLinha < 2) {
    return;
  }

  const dados = abaVagas
    .getRange(
      2,
      1,
      ultimaLinha - 1,
      6
    )
    .getValues();

  const colunas =
    CONFIG_BLOQUEIOS_AUTOMATICOS
      .VAGAS_REGULACAO;

  for (
    let indice = 0;
    indice < dados.length;
    indice++
  ) {
    const linha = dados[indice];
    const data =
      linha[colunas.DATA - 1];

    if (
      !(data instanceof Date) ||
      chaveDataBloqueioAutomatico_(
        data
      ) !==
        chaveDataBloqueioAutomatico_(
          agendamento.data
        ) ||
      chaveHorarioBloqueioAutomatico_(
        linha[colunas.HORARIO - 1]
      ) !==
        chaveHorarioBloqueioAutomatico_(
          agendamento.horario
        ) ||
      normalizarTextoBloqueioAutomatico_(
        linha[
          colunas.FISIOTERAPEUTA - 1
        ]
      ) !==
        normalizarTextoBloqueioAutomatico_(
          agendamento.fisioterapeuta
        )
    ) {
      continue;
    }

    abaVagas
      .getRange(
        indice + 2,
        colunas.SITUACAO
      )
      .setValue('Bloqueada');

    return;
  }
}


function localizarProximaVagaAvaliacao_(
  abas,
  agendamento
) {
  const ultimaLinha =
    abas.vagasRegulacao.getLastRow();

  if (ultimaLinha < 2) {
    return null;
  }

  const dados = abas.vagasRegulacao
    .getRange(
      2,
      1,
      ultimaLinha - 1,
      6
    )
    .getValues();

  const colunas =
    CONFIG_BLOQUEIOS_AUTOMATICOS
      .VAGAS_REGULACAO;

  const turnoOriginal =
    obterTurnoBloqueioAutomatico_(
      agendamento.horario
    );

  const profissionalOriginal =
    normalizarTextoBloqueioAutomatico_(
      agendamento.fisioterapeuta
    );

  const momentoOriginal =
    combinarDataHorarioBloqueioAutomatico_(
      agendamento.data,
      agendamento.horario
    );

  const vagas = dados
    .map(function (linha, indice) {
      return {
        linha: indice + 2,
        data:
          linha[colunas.DATA - 1],
        dia: String(
          linha[colunas.DIA - 1] || ''
        ).trim(),
        horario:
          linha[colunas.HORARIO - 1],
        fisioterapeuta: String(
          linha[
            colunas.FISIOTERAPEUTA - 1
          ] || ''
        ).trim(),
        turno: String(
          linha[colunas.TURNO - 1] || ''
        ).trim(),
        situacao: String(
          linha[colunas.SITUACAO - 1] || ''
        ).trim()
      };
    })
    .filter(function (vaga) {
      if (
        !(vaga.data instanceof Date) ||
        normalizarTextoBloqueioAutomatico_(
          vaga.situacao
        ) !== 'disponivel' ||
        normalizarTextoBloqueioAutomatico_(
          vaga.fisioterapeuta
        ) !== profissionalOriginal
      ) {
        return false;
      }

      const turnoVaga =
        normalizarTextoBloqueioAutomatico_(
          vaga.turno
        ) ||
        obterTurnoBloqueioAutomatico_(
          vaga.horario
        );

      if (turnoVaga !== turnoOriginal) {
        return false;
      }

      const momentoVaga =
        combinarDataHorarioBloqueioAutomatico_(
          vaga.data,
          vaga.horario
        );

      if (
        !momentoVaga ||
        !momentoOriginal ||
        momentoVaga.getTime() <=
          momentoOriginal.getTime()
      ) {
        return false;
      }

      const pseudoAgendamento = {
        data: vaga.data,
        horario: vaga.horario,
        fisioterapeuta:
          vaga.fisioterapeuta
      };

      if (
        estaAtendimentoBloqueadoPorQualquerBloqueio_(
          abas.bloqueios,
          pseudoAgendamento
        )
      ) {
        return false;
      }

      return !vagaAvaliacaoJaOcupada_(
        abas.agendamentos,
        vaga
      );
    });

  vagas.sort(function (a, b) {
    return (
      combinarDataHorarioBloqueioAutomatico_(
        a.data,
        a.horario
      ).getTime() -
      combinarDataHorarioBloqueioAutomatico_(
        b.data,
        b.horario
      ).getTime()
    );
  });

  return vagas[0] || null;
}


function estaAtendimentoBloqueadoPorQualquerBloqueio_(
  abaBloqueios,
  agendamento
) {
  const ultimaLinha =
    abaBloqueios.getLastRow();

  if (ultimaLinha < 2) {
    return false;
  }

  const dados = abaBloqueios
    .getRange(
      2,
      1,
      ultimaLinha - 1,
      7
    )
    .getValues();

  return dados.some(
    function (linha, indice) {
      const bloqueio = {
        linha: indice + 2,
        data: linha[0],
        horario: linha[1],
        fisioterapeuta:
          String(linha[2] || '').trim(),
        abrangencia:
          String(linha[3] || '').trim(),
        motivo:
          String(linha[4] || '').trim(),
        acao:
          String(linha[5] || '').trim(),
        status:
          String(linha[6] || '').trim()
      };

      const status =
        normalizarTextoBloqueioAutomatico_(
          bloqueio.status
        );

      return (
        (
          status === 'ativo' ||
          status === 'bloqueado'
        ) &&
        bloqueioAfetaAtendimentoAutomatico_(
          bloqueio,
          agendamento
        )
      );
    }
  );
}


function vagaAvaliacaoJaOcupada_(
  abaAgendamentos,
  vaga
) {
  const ultimaLinha =
    abaAgendamentos.getLastRow();

  if (ultimaLinha < 2) {
    return false;
  }

  const dados = abaAgendamentos
    .getRange(
      2,
      1,
      ultimaLinha - 1,
      CONFIG_BLOQUEIOS_AUTOMATICOS
        .QUANTIDADE_COLUNAS_AGENDAMENTOS
    )
    .getValues();

  const colunas =
    CONFIG_BLOQUEIOS_AUTOMATICOS
      .AGENDAMENTOS;

  return dados.some(function (linha) {
    const data =
      linha[colunas.DATA - 1];

    return (
      data instanceof Date &&
      chaveDataBloqueioAutomatico_(
        data
      ) ===
        chaveDataBloqueioAutomatico_(
          vaga.data
        ) &&
      chaveHorarioBloqueioAutomatico_(
        linha[colunas.HORARIO - 1]
      ) ===
        chaveHorarioBloqueioAutomatico_(
          vaga.horario
        ) &&
      normalizarTextoBloqueioAutomatico_(
        linha[
          colunas.FISIOTERAPEUTA - 1
        ]
      ) ===
        normalizarTextoBloqueioAutomatico_(
          vaga.fisioterapeuta
        ) &&
      normalizarTextoBloqueioAutomatico_(
        linha[colunas.EVENTO - 1]
      ) === 'avaliacao' &&
      normalizarTextoBloqueioAutomatico_(
        linha[colunas.STATUS - 1]
      ) === 'agendado'
    );
  });
}
function atualizarAvaliacaoOriginalBloqueada_(
  abaAgendamentos,
  agendamento,
  novoId,
  bloqueio,
  agora
) {
  const colunas =
    CONFIG_BLOQUEIOS_AUTOMATICOS
      .AGENDAMENTOS;

  let motivo =
    'Cancelada por bloqueio da clínica: ' +
    bloqueio.motivo;

  if (novoId) {
    motivo +=
      ' | Reagendada automaticamente: ' +
      novoId;
  } else {
    motivo +=
      ' | Sem vaga futura compatível';
  }

  abaAgendamentos
    .getRange(
      agendamento.linhaPlanilha,
      colunas.STATUS
    )
    .setValue(
      'Cancelado pela Clínica'
    );

  abaAgendamentos
    .getRange(
      agendamento.linhaPlanilha,
      colunas.MOTIVO
    )
    .setValue(motivo);

  abaAgendamentos
    .getRange(
      agendamento.linhaPlanilha,
      colunas.CONTA_COMO_SESSAO
    )
    .setValue('Não');

  abaAgendamentos
    .getRange(
      agendamento.linhaPlanilha,
      colunas.AVISAR_PACIENTE
    )
    .setValue('Sim');

  abaAgendamentos
    .getRange(
      agendamento.linhaPlanilha,
      colunas.ATUALIZADO_EM
    )
    .setValue(agora)
    .setNumberFormat(
      'dd/MM/yyyy HH:mm'
    );

  abaAgendamentos
    .getRange(
      agendamento.linhaPlanilha,
      colunas.FATURAVEL
    )
    .setValue('Não');
}


/**
 * Quando uma avaliação é cancelada pela clínica e não
 * existe nova vaga automática, remove do Cadastro os
 * dados da vaga antiga.
 *
 * O status oficial "Avaliação agendada" é preservado
 * porque ele pertence à lista válida do Cadastro, mas a
 * célula recebe uma observação informando que o paciente
 * está aguardando reagendamento.
 */
function atualizarCadastroAvaliacaoSemVaga_(
  abaCadastro,
  idPaciente,
  agendamento,
  bloqueio
) {
  const ultimaLinha =
    abaCadastro.getLastRow();

  if (ultimaLinha < 2) {
    return;
  }

  const ids =
    abaCadastro
      .getRange(
        2,
        CONFIG_BLOQUEIOS_AUTOMATICOS
          .CADASTRO.ID_PACIENTE,
        ultimaLinha - 1,
        1
      )
      .getDisplayValues();

  const procurado =
    normalizarTextoBloqueioAutomatico_(
      idPaciente
    );

  for (
    let indice = 0;
    indice < ids.length;
    indice++
  ) {
    if (
      normalizarTextoBloqueioAutomatico_(
        ids[indice][0]
      ) !== procurado
    ) {
      continue;
    }

    const linha =
      indice + 2;

    abaCadastro
      .getRange(
        linha,
        CONFIG_BLOQUEIOS_AUTOMATICOS
          .CADASTRO.DATA_AVALIACAO
      )
      .clearContent();

    abaCadastro
      .getRange(
        linha,
        CONFIG_BLOQUEIOS_AUTOMATICOS
          .CADASTRO.HORARIO_AVALIACAO
      )
      .clearContent();

    const celulaStatus =
      abaCadastro.getRange(
        linha,
        CONFIG_BLOQUEIOS_AUTOMATICOS
          .CADASTRO.STATUS
      );

    celulaStatus
      .setValue(
        'Avaliação agendada'
      )
      .setNote(
        'A avaliação anterior foi cancelada pela clínica e não houve vaga automática para reagendamento. ' +
        'Agendamento original: ' +
        agendamento.idAgendamento +
        '. Motivo do bloqueio: ' +
        bloqueio.motivo +
        '.'
      );

    return;
  }
}


/**
 * Cria uma pendência urgente quando a avaliação foi
 * cancelada por bloqueio e não existe nova vaga automática.
 *
 * A pendência é criada apenas se ainda não existir uma
 * pendência com o mesmo ID.
 */
function registrarPendenciaReagendamentoAvaliacao_(
  ss,
  abaCadastro,
  agendamento,
  bloqueio
) {
  const abaPendencias =
    ss.getSheetByName(
      'Pendências'
    );

  if (!abaPendencias) {
    return;
  }

  const idPendencia =
    'PEND-BLOQ-AVAL-' +
    String(
      agendamento.idAgendamento || ''
    )
      .trim()
      .toUpperCase()
      .replace(
        /[^A-Z0-9]+/g,
        '-'
      )
      .replace(
        /^-+|-+$/g,
        ''
      );

  const ultimaLinhaPendencias =
    abaPendencias.getLastRow();

  if (
    ultimaLinhaPendencias >= 2
  ) {
    const ids =
      abaPendencias
        .getRange(
          2,
          1,
          ultimaLinhaPendencias - 1,
          1
        )
        .getDisplayValues();

    const existe =
      ids.some(function(linha) {
        return (
          normalizarTextoBloqueioAutomatico_(
            linha[0]
          ) ===
          normalizarTextoBloqueioAutomatico_(
            idPendencia
          )
        );
      });

    if (existe) {
      return;
    }
  }

  const paciente =
    obterDadosPacientePendenciaBloqueio_(
      abaCadastro,
      agendamento.idPaciente
    );

  const hoje =
    removerHorarioBloqueioAutomatico_(
      new Date()
    );

  const novaLinha =
    Math.max(
      abaPendencias.getLastRow() + 1,
      2
    );

  abaPendencias
    .getRange(
      novaLinha,
      1,
      1,
      11
    )
    .setValues([[
      idPendencia,
      'Urgente',
      'Reagendar avaliação cancelada por bloqueio — sem vaga automática',
      paciente.nome,
      paciente.prontuario,
      paciente.telefone,
      'Recepção',
      new Date(bloqueio.data),
      hoje,
      '',
      ''
    ]]);

  abaPendencias
    .getRange(
      novaLinha,
      8,
      1,
      2
    )
    .setNumberFormat(
      'dd/MM/yyyy'
    );
}


/**
 * Obtém os dados mínimos do paciente necessários para
 * registrar a pendência de reagendamento da avaliação.
 */
function obterDadosPacientePendenciaBloqueio_(
  abaCadastro,
  idPaciente
) {
  const resultado = {
    prontuario: '',
    nome: '',
    telefone: ''
  };

  const ultimaLinha =
    abaCadastro.getLastRow();

  if (ultimaLinha < 2) {
    return resultado;
  }

  const dados =
    abaCadastro
      .getRange(
        2,
        1,
        ultimaLinha - 1,
        Math.max(
          CONFIG_BLOQUEIOS_AUTOMATICOS
            .CADASTRO.STATUS,
          5
        )
      )
      .getValues();

  const procurado =
    normalizarTextoBloqueioAutomatico_(
      idPaciente
    );

  for (
    let indice = 0;
    indice < dados.length;
    indice++
  ) {
    const linha =
      dados[indice];

    if (
      normalizarTextoBloqueioAutomatico_(
        linha[
          CONFIG_BLOQUEIOS_AUTOMATICOS
            .CADASTRO.ID_PACIENTE - 1
        ]
      ) !== procurado
    ) {
      continue;
    }

    resultado.prontuario =
      String(
        linha[1] || ''
      ).trim();

    resultado.nome =
      String(
        linha[2] || ''
      ).trim();

    resultado.telefone =
      String(
        linha[4] || ''
      ).trim();

    return resultado;
  }

  return resultado;
}


/**
 * Remove o horário de uma data.
 */
function removerHorarioBloqueioAutomatico_(
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


function atualizarCadastroAvaliacaoReagendada_(
  abaCadastro,
  idPaciente,
  vaga
) {
  const ultimaLinha =
    abaCadastro.getLastRow();

  if (ultimaLinha < 2) {
    return;
  }

  const ids = abaCadastro
    .getRange(
      2,
      CONFIG_BLOQUEIOS_AUTOMATICOS
        .CADASTRO.ID_PACIENTE,
      ultimaLinha - 1,
      1
    )
    .getDisplayValues();

  const idProcurado =
    normalizarTextoBloqueioAutomatico_(
      idPaciente
    );

  for (
    let indice = 0;
    indice < ids.length;
    indice++
  ) {
    if (
      normalizarTextoBloqueioAutomatico_(
        ids[indice][0]
      ) !== idProcurado
    ) {
      continue;
    }

    const linha = indice + 2;

    abaCadastro
      .getRange(
        linha,
        CONFIG_BLOQUEIOS_AUTOMATICOS
          .CADASTRO.DATA_AVALIACAO
      )
      .setValue(
        new Date(vaga.data)
      )
      .setNumberFormat(
        'dd/MM/yyyy'
      );

    abaCadastro
      .getRange(
        linha,
        CONFIG_BLOQUEIOS_AUTOMATICOS
          .CADASTRO.HORARIO_AVALIACAO
      )
      .setValue(vaga.horario)
      .setNumberFormat('HH:mm');

    abaCadastro
      .getRange(
        linha,
        CONFIG_BLOQUEIOS_AUTOMATICOS
          .CADASTRO.STATUS
      )
      .setValue(
        'Avaliação agendada'
      );

    abaCadastro
      .getRange(
        linha,
        CONFIG_BLOQUEIOS_AUTOMATICOS
          .CADASTRO.FISIOTERAPEUTA
      )
      .setValue(
        vaga.fisioterapeuta
      );

    return;
  }
}


function formatarLinhaAvaliacaoReagendada_(
  abaAgendamentos,
  linha
) {
  const colunas =
    CONFIG_BLOQUEIOS_AUTOMATICOS
      .AGENDAMENTOS;

  abaAgendamentos
    .getRange(
      linha,
      colunas.DATA
    )
    .setNumberFormat('dd/MM/yyyy');

  abaAgendamentos
    .getRange(
      linha,
      colunas.HORARIO
    )
    .setNumberFormat('HH:mm');

  abaAgendamentos
    .getRange(
      linha,
      colunas.CRIADO_EM,
      1,
      2
    )
    .setNumberFormat(
      'dd/MM/yyyy HH:mm'
    );
}


function gerarProximoIdAgendamentoBloqueio_(
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

  ids.forEach(function (linha) {
    const resultado =
      String(linha[0] || '')
        .match(/(\d+)$/);

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


function removerCancelamentoClinicaDaListaStatus_(
  ss
) {
  const aba = ss.getSheetByName(
    CONFIG_BLOQUEIOS_AUTOMATICOS
      .ABAS.STATUS_SESSAO
  );

  if (!aba || aba.getLastRow() < 2) {
    return;
  }

  const quantidade =
    aba.getLastRow() - 1;

  const valores = aba
    .getRange(
      2,
      1,
      quantidade,
      1
    )
    .getDisplayValues()
    .map(function (linha) {
      return String(
        linha[0] || ''
      ).trim();
    })
    .filter(function (status) {
      return (
        status &&
        normalizarTextoBloqueioAutomatico_(
          status
        ) !==
          'cancelado pela clinica'
      );
    });

  aba.getRange(
    2,
    1,
    quantidade,
    1
  ).clearContent();

  if (valores.length > 0) {
    aba.getRange(
      2,
      1,
      valores.length,
      1
    ).setValues(
      valores.map(function (status) {
        return [status];
      })
    );
  }
}


function atualizarModulosAposBloqueio_(
  ss,
  datasAfetadas
) {
  try {
    if (
      typeof
        atualizarPendenciasAutomaticas ===
        'function'
    ) {
      atualizarPendenciasAutomaticas();
    }
  } catch (erroPendencias) {
    console.error(
      erroPendencias
    );
  }

  try {
    const abaAgenda = ss.getSheetByName(
      CONFIG_BLOQUEIOS_AUTOMATICOS
        .ABAS.AGENDA
    );

    const dataAgenda =
      abaAgenda &&
      abaAgenda.getRange('B2').getValue();

    const agendaAfetada =
      dataAgenda instanceof Date &&
      (datasAfetadas || [])
        .some(function (data) {
          return (
            data instanceof Date &&
            chaveDataBloqueioAutomatico_(
              data
            ) ===
              chaveDataBloqueioAutomatico_(
                dataAgenda
              )
          );
        });

    if (
      agendaAfetada &&
      typeof carregarAgendaDiaria ===
        'function'
    ) {
      carregarAgendaDiaria();
    }
  } catch (erroAgenda) {
    console.error(erroAgenda);
  }
}


function combinarDataHorarioBloqueioAutomatico_(
  data,
  horario
) {
  if (!(data instanceof Date)) {
    return null;
  }

  const chave =
    chaveHorarioBloqueioAutomatico_(
      horario
    );

  if (!chave) {
    return null;
  }

  const partes =
    chave.split(':');

  const resultado =
    new Date(data);

  resultado.setHours(
    Number(partes[0]),
    Number(partes[1]),
    0,
    0
  );

  return resultado;
}


function obterNomeDiaBloqueioAutomatico_(
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


function obterTurnoBloqueioAutomatico_(
  horario
) {
  const chave =
    chaveHorarioBloqueioAutomatico_(
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


function chaveDataBloqueioAutomatico_(
  data
) {
  return Utilities.formatDate(
    new Date(data),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );
}


function chaveHorarioBloqueioAutomatico_(
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

  if (typeof valor === 'number') {
    const minutos =
      Math.round(
        valor * 24 * 60
      );

    return (
      String(
        Math.floor(minutos / 60) % 24
      ).padStart(2, '0') +
      ':' +
      String(
        minutos % 60
      ).padStart(2, '0')
    );
  }

  const texto =
    String(valor || '').trim();

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


function normalizarTextoBloqueioAutomatico_(
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
