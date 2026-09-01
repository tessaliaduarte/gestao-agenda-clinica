const CONFIG_RESULTADO_AVALIACAO = {
  ABAS: {
    AGENDA: 'Agenda',
    AGENDAMENTOS: 'Agendamentos',
    CADASTRO: 'Cadastro de Pacientes'
  },

  AGENDA: {
    PRIMEIRA_LINHA: 5,
    COLUNA_STATUS: 4,
    COLUNA_ID_AGENDAMENTO: 6
  },

  AGENDAMENTOS: {
    ID_AGENDAMENTO: 1,
    ID_PACIENTE: 2,
    EVENTO: 12,
    STATUS: 16,
    CONTA_COMO_SESSAO: 18,
    ATUALIZADO_EM: 21,
    FATURAVEL: 22
  },

  CADASTRO: {
    ID_PACIENTE: 1,
    STATUS: 21
  },

  QUANTIDADE_COLUNAS_AGENDAMENTOS: 22
};


/**
 * Instala o gatilho que processa o resultado
 * das avaliações.
 *
 * Execute esta função manualmente apenas uma vez.
 */
function instalarGatilhoResultadoAvaliacao() {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const ui =
    SpreadsheetApp.getUi();

  try {
    removerGatilhosResultadoAvaliacao_();

    ScriptApp.newTrigger(
      'processarResultadoAvaliacao'
    )
      .forSpreadsheet(ss)
      .onEdit()
      .create();

    ui.alert(
      'Resultado da avaliação ativado',
      'O sistema atualizará automaticamente o status do paciente após o resultado da avaliação.',
      ui.ButtonSet.OK
    );
  } catch (erro) {
    ui.alert(
      'Erro ao ativar resultado da avaliação',
      erro && erro.message
        ? erro.message
        : String(erro),
      ui.ButtonSet.OK
    );
  }
}


/**
 * Remove gatilhos duplicados deste módulo.
 */
function removerGatilhosResultadoAvaliacao_() {
  ScriptApp
    .getProjectTriggers()
    .forEach(function(gatilho) {
      if (
        gatilho.getHandlerFunction() ===
        'processarResultadoAvaliacao'
      ) {
        ScriptApp.deleteTrigger(
          gatilho
        );
      }
    });
}


/**
 * Processa alterações feitas na coluna Status
 * da Agenda.
 */
function processarResultadoAvaliacao(e) {
  if (!e || !e.range) {
    return;
  }

  const abaAgenda =
    e.range.getSheet();

  if (
    abaAgenda.getName() !==
    CONFIG_RESULTADO_AVALIACAO
      .ABAS.AGENDA
  ) {
    return;
  }

  const linha =
    e.range.getRow();

  const coluna =
    e.range.getColumn();

  if (
    linha <
      CONFIG_RESULTADO_AVALIACAO
        .AGENDA.PRIMEIRA_LINHA ||
    coluna !==
      CONFIG_RESULTADO_AVALIACAO
        .AGENDA.COLUNA_STATUS
  ) {
    return;
  }

  const idAgendamento = String(
    abaAgenda.getRange(
      linha,
      CONFIG_RESULTADO_AVALIACAO
        .AGENDA.COLUNA_ID_AGENDAMENTO
    ).getValue() || ''
  ).trim();

  if (!idAgendamento) {
    return;
  }

  const novoStatus = String(
    e.value ||
    e.range.getValue() ||
    ''
  ).trim();

  if (!novoStatus) {
    return;
  }

  const lock =
    LockService.getDocumentLock();

  let bloqueioObtido = false;

  try {
    lock.waitLock(30000);
    bloqueioObtido = true;

    const ss =
      e.source ||
      SpreadsheetApp
        .getActiveSpreadsheet();

    const abaAgendamentos =
      ss.getSheetByName(
        CONFIG_RESULTADO_AVALIACAO
          .ABAS.AGENDAMENTOS
      );

    const abaCadastro =
      ss.getSheetByName(
        CONFIG_RESULTADO_AVALIACAO
          .ABAS.CADASTRO
      );

    if (!abaAgendamentos) {
      throw new Error(
        'A aba "Agendamentos" não foi encontrada.'
      );
    }

    if (!abaCadastro) {
      throw new Error(
        'A aba "Cadastro de Pacientes" não foi encontrada.'
      );
    }

    const agendamento =
      localizarAgendamentoAvaliacao_(
        abaAgendamentos,
        idAgendamento
      );

    if (!agendamento) {
      throw new Error(
        'O agendamento não foi encontrado.'
      );
    }

    const evento =
      normalizarTextoResultadoAvaliacao_(
        agendamento.evento
      );

    if (evento !== 'avaliacao') {
      return;
    }

    /*
     * Garante que a avaliação fique com os
     * campos corretos em Agendamentos.
     */
    atualizarRegistroAvaliacao_(
      abaAgendamentos,
      agendamento.linha,
      novoStatus
    );

    /*
     * Atualiza o fluxo do paciente no Cadastro.
     */
    atualizarStatusPacienteAvaliacao_(
      abaCadastro,
      agendamento.idPaciente,
      novoStatus
    );

    SpreadsheetApp.flush();

    const statusNormalizado =
      normalizarTextoResultadoAvaliacao_(
        novoStatus
      );

    if (
      statusNormalizado ===
      'compareceu'
    ) {
      ss.toast(
        'Avaliação concluída. O paciente está aguardando o agendamento das sessões.',
        'Avaliação registrada',
        6
      );
    }
  } catch (erro) {
    const ss =
      e.source ||
      SpreadsheetApp
        .getActiveSpreadsheet();

    ss.toast(
      erro && erro.message
        ? erro.message
        : String(erro),
      'Erro no resultado da avaliação',
      7
    );

    console.error(erro);
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
 * Localiza o registro na aba Agendamentos.
 */
function localizarAgendamentoAvaliacao_(
  abaAgendamentos,
  idAgendamento
) {
  const ultimaLinha =
    abaAgendamentos.getLastRow();

  if (ultimaLinha < 2) {
    return null;
  }

  const dados =
    abaAgendamentos.getRange(
      2,
      1,
      ultimaLinha - 1,
      CONFIG_RESULTADO_AVALIACAO
        .QUANTIDADE_COLUNAS_AGENDAMENTOS
    ).getValues();

  const idProcurado =
    normalizarTextoResultadoAvaliacao_(
      idAgendamento
    );

  for (
    let indice = 0;
    indice < dados.length;
    indice++
  ) {
    const linha =
      dados[indice];

    const idAtual =
      normalizarTextoResultadoAvaliacao_(
        linha[
          CONFIG_RESULTADO_AVALIACAO
            .AGENDAMENTOS
            .ID_AGENDAMENTO - 1
        ]
      );

    if (idAtual === idProcurado) {
      return {
        linha: indice + 2,

        idPaciente: String(
          linha[
            CONFIG_RESULTADO_AVALIACAO
              .AGENDAMENTOS
              .ID_PACIENTE - 1
          ] || ''
        ).trim(),

        evento: String(
          linha[
            CONFIG_RESULTADO_AVALIACAO
              .AGENDAMENTOS
              .EVENTO - 1
          ] || ''
        ).trim()
      };
    }
  }

  return null;
}


/**
 * Atualiza o registro da avaliação.
 */
function atualizarRegistroAvaliacao_(
  abaAgendamentos,
  linha,
  status
) {
  const colunas =
    CONFIG_RESULTADO_AVALIACAO
      .AGENDAMENTOS;

  const statusNormalizado =
    normalizarTextoResultadoAvaliacao_(
      status
    );

  let faturavel = 'Não';

  if (
    statusNormalizado ===
    'compareceu'
  ) {
    faturavel = 'Sim';
  }

  abaAgendamentos.getRange(
    linha,
    colunas.STATUS
  ).setValue(status);

  /*
   * Avaliação nunca conta como sessão.
   */
  abaAgendamentos.getRange(
    linha,
    colunas.CONTA_COMO_SESSAO
  ).setValue('Não');

  abaAgendamentos.getRange(
    linha,
    colunas.FATURAVEL
  ).setValue(faturavel);

  abaAgendamentos.getRange(
    linha,
    colunas.ATUALIZADO_EM
  )
    .setValue(new Date())
    .setNumberFormat(
      'dd/MM/yyyy HH:mm'
    );
}


/**
 * Atualiza o status do paciente conforme
 * o resultado da avaliação.
 */
function atualizarStatusPacienteAvaliacao_(
  abaCadastro,
  idPaciente,
  statusAvaliacao
) {
  const ultimaLinha =
    abaCadastro.getLastRow();

  if (ultimaLinha < 2) {
    throw new Error(
      'O Cadastro de Pacientes está vazio.'
    );
  }

  const ids =
    abaCadastro.getRange(
      2,
      CONFIG_RESULTADO_AVALIACAO
        .CADASTRO.ID_PACIENTE,
      ultimaLinha - 1,
      1
    ).getDisplayValues();

  const idProcurado =
    normalizarTextoResultadoAvaliacao_(
      idPaciente
    );

  let linhaPaciente = null;

  for (
    let indice = 0;
    indice < ids.length;
    indice++
  ) {
    const idAtual =
      normalizarTextoResultadoAvaliacao_(
        ids[indice][0]
      );

    if (idAtual === idProcurado) {
      linhaPaciente =
        indice + 2;

      break;
    }
  }

  if (!linhaPaciente) {
    throw new Error(
      'O paciente da avaliação não foi encontrado no Cadastro.'
    );
  }

  const statusNormalizado =
    normalizarTextoResultadoAvaliacao_(
      statusAvaliacao
    );

  let novoStatus =
    'Avaliação agendada';

  if (
    statusNormalizado ===
    'compareceu'
  ) {
    novoStatus =
      'Avaliado – aguardando agendamento';
  }

  abaCadastro.getRange(
    linhaPaciente,
    CONFIG_RESULTADO_AVALIACAO
      .CADASTRO.STATUS
  ).setValue(novoStatus);
}


/**
 * Padroniza textos para comparação.
 */
function normalizarTextoResultadoAvaliacao_(
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
