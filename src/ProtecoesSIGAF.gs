const CONFIG_PROTECOES_SIGAF = {
  PREFIXO_DESCRICAO:
    'SIGAF — Proteção de aviso — '
};


/**
 * Aplica proteções de aviso às áreas automáticas.
 *
 * As proteções não bloqueiam edições e não interferem
 * nos formulários, menus ou gatilhos.
 */
function aplicarProtecoesAvisoSIGAF() {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const ui =
    SpreadsheetApp.getUi();

  try {
    removerProtecoesAvisoSIGAF_(
      ss
    );

    /*
     * Abas inteiramente automáticas ou históricas.
     */
    protegerPlanilhaAvisoSIGAF_(
      ss,
      'Cadastro de Pacientes',
      'Cadastro de Pacientes'
    );

    protegerPlanilhaAvisoSIGAF_(
      ss,
      'Agendamentos',
      'Agendamentos'
    );

    protegerPlanilhaAvisoSIGAF_(
      ss,
      'Histórico de Pendências',
      'Histórico de Pendências'
    );

    protegerPlanilhaAvisoSIGAF_(
      ss,
      'Histórico de Desfechos',
      'Histórico de Desfechos'
    );

    protegerPlanilhaAvisoSIGAF_(
      ss,
      'Vagas para Regulação',
      'Vagas para Regulação'
    );

    /*
     * Abas de configuração.
     */
    protegerPlanilhaAvisoSIGAF_(
      ss,
      'Fisioterapeutas',
      'Fisioterapeutas'
    );

    protegerPlanilhaAvisoSIGAF_(
      ss,
      'Horários',
      'Horários'
    );

    protegerPlanilhaAvisoSIGAF_(
      ss,
      'Tipos de Evento',
      'Tipos de Evento'
    );

    protegerPlanilhaAvisoSIGAF_(
      ss,
      'Tipos de Grupo',
      'Tipos de Grupo'
    );

    protegerPlanilhaAvisoSIGAF_(
      ss,
      'Status da Sessão',
      'Status da Sessão'
    );

    protegerPlanilhaAvisoSIGAF_(
      ss,
      'Calendário da Prefeitura',
      'Calendário da Prefeitura'
    );

    /*
     * Proteções parciais da Agenda.
     *
     * Permanecem livres:
     * B2 — data;
     * B3 — fisioterapeuta;
     * D5:D — status dos atendimentos.
     */
    protegerCabecalhoAgendaAvisoSIGAF_(
      ss
    );

    protegerDadosAgendaAvisoSIGAF_(
      ss
    );

    /*
     * Na aba Pendências, J e K permanecem livres:
     * J — Situação;
     * K — Data da conclusão.
     */
    protegerPendenciasAvisoSIGAF_(
      ss
    );

    /*
     * Na aba Vagas para Sessões, a coluna A
     * permanece livre para selecionar a opção.
     */
    protegerVagasSessoesAvisoSIGAF_(
      ss
    );

    SpreadsheetApp.flush();

    ui.alert(
      'Proteções aplicadas',
      'As áreas automáticas receberam proteção de aviso.\n\n' +
        'Os scripts continuarão funcionando normalmente.',
      ui.ButtonSet.OK
    );
  } catch (erro) {
    ui.alert(
      'Erro ao aplicar proteções',
      erro && erro.message
        ? erro.message
        : String(erro),
      ui.ButtonSet.OK
    );

    throw erro;
  }
}


/**
 * Protege uma planilha inteira apenas com aviso.
 */
function protegerPlanilhaAvisoSIGAF_(
  ss,
  nomeAba,
  identificacao
) {
  const aba =
    ss.getSheetByName(nomeAba);

  if (!aba) {
    throw new Error(
      'A aba "' +
        nomeAba +
        '" não foi encontrada.'
    );
  }

  const protecao =
    aba.protect();

  protecao.setDescription(
    CONFIG_PROTECOES_SIGAF
      .PREFIXO_DESCRICAO +
      identificacao
  );

  protecao.setWarningOnly(true);
}


/**
 * Protege um intervalo apenas com aviso.
 */
function protegerIntervaloAvisoSIGAF_(
  aba,
  intervaloA1,
  identificacao
) {
  const protecao =
    aba.getRange(
      intervaloA1
    ).protect();

  protecao.setDescription(
    CONFIG_PROTECOES_SIGAF
      .PREFIXO_DESCRICAO +
      identificacao
  );

  protecao.setWarningOnly(true);
}


/**
 * Protege os cabeçalhos e elementos fixos da Agenda.
 */
function protegerCabecalhoAgendaAvisoSIGAF_(
  ss
) {
  const aba =
    ss.getSheetByName('Agenda');

  if (!aba) {
    throw new Error(
      'A aba "Agenda" não foi encontrada.'
    );
  }

  protegerIntervaloAvisoSIGAF_(
    aba,
    'A1:F1',
    'Agenda — título'
  );

  protegerIntervaloAvisoSIGAF_(
    aba,
    'A2:A4',
    'Agenda — rótulos'
  );

  protegerIntervaloAvisoSIGAF_(
    aba,
    'B4:F4',
    'Agenda — cabeçalhos'
  );

  protegerIntervaloAvisoSIGAF_(
    aba,
    'C2:F3',
    'Agenda — estrutura superior'
  );
}


/**
 * Protege os dados automáticos da Agenda.
 */
function protegerDadosAgendaAvisoSIGAF_(
  ss
) {
  const aba =
    ss.getSheetByName('Agenda');

  if (!aba) {
    throw new Error(
      'A aba "Agenda" não foi encontrada.'
    );
  }

  const ultimaLinha =
    Math.max(
      aba.getMaxRows(),
      5
    );

  protegerIntervaloAvisoSIGAF_(
    aba,
    'A5:C' + ultimaLinha,
    'Agenda — dados automáticos A:C'
  );

  protegerIntervaloAvisoSIGAF_(
    aba,
    'E5:F' + ultimaLinha,
    'Agenda — dados automáticos E:F'
  );
}


/**
 * Protege os campos automáticos de Pendências.
 */
function protegerPendenciasAvisoSIGAF_(
  ss
) {
  const aba =
    ss.getSheetByName('Pendências');

  if (!aba) {
    throw new Error(
      'A aba "Pendências" não foi encontrada.'
    );
  }

  const ultimaLinha =
    Math.max(
      aba.getMaxRows(),
      2
    );

  protegerIntervaloAvisoSIGAF_(
    aba,
    'A1:K1',
    'Pendências — cabeçalho'
  );

  protegerIntervaloAvisoSIGAF_(
    aba,
    'A2:I' + ultimaLinha,
    'Pendências — dados automáticos'
  );
}


/**
 * Protege os resultados da consulta de vagas.
 */
function protegerVagasSessoesAvisoSIGAF_(
  ss
) {
  const aba =
    ss.getSheetByName(
      'Vagas para Sessões'
    );

  if (!aba) {
    throw new Error(
      'A aba "Vagas para Sessões" não foi encontrada.'
    );
  }

  const ultimaLinha =
    Math.max(
      aba.getMaxRows(),
      2
    );

  protegerIntervaloAvisoSIGAF_(
    aba,
    'A1:N1',
    'Vagas para Sessões — cabeçalho'
  );

  protegerIntervaloAvisoSIGAF_(
    aba,
    'B2:N' + ultimaLinha,
    'Vagas para Sessões — resultados'
  );
}


/**
 * Remove somente proteções criadas por este módulo,
 * evitando duplicidade quando ele for executado novamente.
 */
function removerProtecoesAvisoSIGAF_(
  ss
) {
  const prefixo =
    CONFIG_PROTECOES_SIGAF
      .PREFIXO_DESCRICAO;

  ss.getSheets().forEach(
    function(aba) {
      const protecoesPlanilha =
        aba.getProtections(
          SpreadsheetApp
            .ProtectionType
            .SHEET
        );

      const protecoesIntervalo =
        aba.getProtections(
          SpreadsheetApp
            .ProtectionType
            .RANGE
        );

      protecoesPlanilha
        .concat(
          protecoesIntervalo
        )
        .forEach(
          function(protecao) {
            const descricao =
              String(
                protecao.getDescription() ||
                ''
              );

            if (
              descricao.indexOf(
                prefixo
              ) === 0
            ) {
              protecao.remove();
            }
          }
        );
    }
  );
}
