const CONFIG_LIMPEZA_TESTES_SIGAF = {
  ABAS_DADOS: [
    {
      nome: 'Cadastro de Pacientes',
      primeiraLinha: 2,
      primeiraColuna: 1,
      quantidadeColunas: 24
    },
    {
      nome: 'Agendamentos',
      primeiraLinha: 2,
      primeiraColuna: 1,
      quantidadeColunas: 22
    },
    {
      nome: 'Pendências',
      primeiraLinha: 2,
      primeiraColuna: 1,
      quantidadeColunas: 11
    },
    {
      nome: 'Histórico de Pendências',
      primeiraLinha: 2,
      primeiraColuna: 1,
      quantidadeColunas: 11
    },
    {
      nome: 'Histórico de Desfechos',
      primeiraLinha: 2,
      primeiraColuna: 1,
      quantidadeColunas: 14
    },
    {
      nome: 'Histórico de Ajustes',
      primeiraLinha: 2,
      primeiraColuna: 1,
      quantidadeColunas: 17,
      opcional: true
    },
    {
      nome: 'Bloqueios',
      primeiraLinha: 2,
      primeiraColuna: 1,
      quantidadeColunas: 7,
      opcional: true
    }
  ],

  ABAS_GERADAS: [
    'Vagas para Sessões',
    'Vagas para Regulação',
    'Relatório Mensal SIGAF'
  ]
};


/**
 * Mostra o que será apagado, sem alterar nenhum dado.
 *
 * Execute esta função antes da limpeza definitiva.
 */
function revisarLimpezaDadosTesteSIGAF() {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const ui =
    SpreadsheetApp.getUi();

  const linhas = [];

  CONFIG_LIMPEZA_TESTES_SIGAF
    .ABAS_DADOS
    .forEach(function(config) {
      const aba =
        ss.getSheetByName(
          config.nome
        );

      if (!aba) {
        if (!config.opcional) {
          linhas.push(
            config.nome +
            ': ABA NÃO ENCONTRADA'
          );
        }

        return;
      }

      const quantidade =
        Math.max(
          aba.getLastRow() -
            config.primeiraLinha +
            1,
          0
        );

      linhas.push(
        config.nome +
        ': ' +
        quantidade +
        ' linha(s) de dados'
      );
    });

  CONFIG_LIMPEZA_TESTES_SIGAF
    .ABAS_GERADAS
    .forEach(function(nome) {
      const aba =
        ss.getSheetByName(nome);

      if (!aba) {
        return;
      }

      linhas.push(
        nome +
        ': conteúdo operacional será regenerado'
      );
    });

  const abaAgenda =
    ss.getSheetByName(
      'Agenda'
    );

  if (abaAgenda) {
    linhas.push(
      'Agenda: registros exibidos a partir da linha 5 serão limpos'
    );
  }

  ui.alert(
    'Revisão da limpeza dos dados de teste',
    'NENHUM dado foi apagado nesta etapa.\n\n' +
      linhas.join('\n') +
      '\n\nAs abas de configuração serão PRESERVADAS:\n' +
      'Fisioterapeutas, Horários, Tipos de Evento, Tipos de Grupo, Status da Sessão e Calendário da Prefeitura.\n\n' +
      'Faça o backup antes de executar a limpeza definitiva.',
    ui.ButtonSet.OK
  );
}


/**
 * Limpa todos os dados de teste do SIGAF.
 *
 * IMPORTANTE:
 * Execute somente DEPOIS de criar uma cópia de segurança
 * da planilha.
 *
 * Esta função preserva cabeçalhos, fórmulas estruturais e
 * abas de configuração.
 */
function limparDadosTesteSIGAF() {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const ui =
    SpreadsheetApp.getUi();

  const primeiraConfirmacao =
    ui.alert(
      'Limpar todos os dados de teste?',
      'Esta ação apagará os pacientes, agendamentos, pendências, históricos e dados operacionais atuais.\n\n' +
        'As configurações do SIGAF serão preservadas.\n\n' +
        'Continue somente se o BACKUP já tiver sido realizado.',
      ui.ButtonSet.YES_NO
    );

  if (
    primeiraConfirmacao !==
      ui.Button.YES
  ) {
    return;
  }

  const resposta =
    ui.prompt(
      'Confirmação final',
      'Para confirmar a limpeza definitiva, digite exatamente:\n\nLIMPAR TESTES',
      ui.ButtonSet.OK_CANCEL
    );

  if (
    resposta.getSelectedButton() !==
      ui.Button.OK
  ) {
    return;
  }

  if (
    String(
      resposta.getResponseText() ||
      ''
    ).trim() !==
      'LIMPAR TESTES'
  ) {
    ui.alert(
      'Limpeza cancelada',
      'O texto de confirmação não corresponde ao solicitado. Nenhum dado foi apagado.',
      ui.ButtonSet.OK
    );

    return;
  }

  const lock =
    LockService.getDocumentLock();

  let bloqueioObtido = false;

  try {
    lock.waitLock(30000);
    bloqueioObtido = true;

    CONFIG_LIMPEZA_TESTES_SIGAF
      .ABAS_DADOS
      .forEach(function(config) {
        limparDadosAbaTesteSIGAF_(
          ss,
          config
        );
      });

    limparAgendaTesteSIGAF_(
      ss
    );

    limparAbaGeradaTesteSIGAF_(
      ss,
      'Vagas para Sessões'
    );

    limparAbaGeradaTesteSIGAF_(
      ss,
      'Vagas para Regulação'
    );

    limparRelatorioTesteSIGAF_(
      ss
    );

    /*
     * Remove seleções/valores residuais sem destruir
     * estrutura e configuração.
     */
    const abaCadastro =
      ss.getSheetByName(
        'Cadastro de Pacientes'
      );

    if (
      abaCadastro &&
      abaCadastro.getLastRow() < 2
    ) {
      // Nenhuma ação adicional necessária.
    }

    SpreadsheetApp.flush();

    ui.alert(
      'Limpeza concluída',
      'Todos os dados operacionais de teste foram removidos.\n\n' +
        'As abas de configuração foram preservadas.\n\n' +
        'O SIGAF está pronto para a validação definitiva após a etapa de backup/checagem.',
      ui.ButtonSet.OK
    );
  } catch (erro) {
    ui.alert(
      'Erro na limpeza',
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
 * Limpa linhas de dados mantendo o cabeçalho e a estrutura.
 */
function limparDadosAbaTesteSIGAF_(
  ss,
  config
) {
  const aba =
    ss.getSheetByName(
      config.nome
    );

  if (!aba) {
    if (config.opcional) {
      return;
    }

    throw new Error(
      'A aba obrigatória "' +
        config.nome +
        '" não foi encontrada.'
    );
  }

  const ultimaLinha =
    aba.getLastRow();

  if (
    ultimaLinha <
      config.primeiraLinha
  ) {
    return;
  }

  const quantidadeLinhas =
    ultimaLinha -
    config.primeiraLinha +
    1;

  aba
    .getRange(
      config.primeiraLinha,
      config.primeiraColuna,
      quantidadeLinhas,
      config.quantidadeColunas
    )
    .clearContent()
    .clearNote();
}


/**
 * Limpa os registros visíveis da Agenda,
 * preservando título, filtros e cabeçalhos.
 */
function limparAgendaTesteSIGAF_(
  ss
) {
  const aba =
    ss.getSheetByName(
      'Agenda'
    );

  if (!aba) {
    return;
  }

  const primeiraLinha =
    5;

  const ultimaLinha =
    Math.max(
      aba.getLastRow(),
      primeiraLinha
    );

  if (
    ultimaLinha >=
      primeiraLinha
  ) {
    aba
      .getRange(
        primeiraLinha,
        1,
        ultimaLinha -
          primeiraLinha +
          1,
        6
      )
      .clearContent()
      .clearNote()
      .clearDataValidations()
      .setBackground(null)
      .setFontWeight('normal')
      .setFontColor(null);
  }
}


/**
 * Limpa uma aba de resultado gerado,
 * mantendo a primeira linha/cabeçalho.
 */
function limparAbaGeradaTesteSIGAF_(
  ss,
  nomeAba
) {
  const aba =
    ss.getSheetByName(
      nomeAba
    );

  if (!aba) {
    return;
  }

  const ultimaLinha =
    aba.getLastRow();

  if (ultimaLinha < 2) {
    return;
  }

  aba
    .getRange(
      2,
      1,
      ultimaLinha - 1,
      aba.getMaxColumns()
    )
    .clearContent()
    .clearNote();
}


/**
 * Limpa o relatório mensal gerado.
 *
 * A aba será reconstruída normalmente na próxima geração.
 */
function limparRelatorioTesteSIGAF_(
  ss
) {
  const aba =
    ss.getSheetByName(
      'Relatório Mensal SIGAF'
    );

  if (!aba) {
    return;
  }

  aba.clearContents();
}


/**
 * Checagem rápida depois da limpeza.
 */
function verificarSistemaLimpoSIGAF() {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const ui =
    SpreadsheetApp.getUi();

  const abasQueDevemEstarVazias = [
    'Cadastro de Pacientes',
    'Agendamentos',
    'Pendências',
    'Histórico de Pendências',
    'Histórico de Desfechos',
    'Histórico de Ajustes',
    'Bloqueios'
  ];

  const problemas = [];

  abasQueDevemEstarVazias.forEach(
    function(nome) {
      const aba =
        ss.getSheetByName(
          nome
        );

      if (!aba) {
        return;
      }

      const ultimaLinha =
        aba.getLastRow();

      if (ultimaLinha >= 2) {
        const valores =
          aba
            .getRange(
              2,
              1,
              ultimaLinha - 1,
              aba.getMaxColumns()
            )
            .getDisplayValues();

        const possuiDados =
          valores.some(
            function(linha) {
              return linha.some(
                function(valor) {
                  return (
                    String(
                      valor || ''
                    ).trim() !==
                    ''
                  );
                }
              );
            }
          );

        if (possuiDados) {
          problemas.push(
            nome
          );
        }
      }
    }
  );

  if (problemas.length === 0) {
    ui.alert(
      'Verificação concluída',
      'Os dados operacionais de teste foram removidos com sucesso.',
      ui.ButtonSet.OK
    );

    return;
  }

  ui.alert(
    'Ainda existem dados',
    'Foram encontrados registros nas seguintes abas:\n\n' +
      problemas.join('\n'),
    ui.ButtonSet.OK
  );
}
