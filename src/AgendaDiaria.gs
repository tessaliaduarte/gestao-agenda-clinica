const CONFIG_AGENDA_DIARIA = {
  ABAS: {
    AGENDA: 'Agenda',
    AGENDAMENTOS: 'Agendamentos',
    CADASTRO: 'Cadastro de Pacientes',
    FISIOTERAPEUTAS: 'Fisioterapeutas',
    STATUS_SESSAO: 'Status da Sessão',
    BLOQUEIOS: 'Bloqueios',
    HORARIOS: 'Horários'
  },

  LINHAS: {
    TITULO: 1,
    DATA: 2,
    FISIOTERAPEUTA: 3,
    CABECALHO: 4,
    PRIMEIRO_REGISTRO: 5
  },

  COLUNAS_AGENDA: {
    HORARIO: 1,
    PACIENTE: 2,
    EVENTO_SESSAO: 3,
    STATUS: 4,
    ALERTA: 5,
    ID_AGENDAMENTO: 6
  },

  COLUNAS_AGENDAMENTOS: {
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

  COLUNAS_CADASTRO: {
    ID_PACIENTE: 1,
    SESSOES_PRESCRITAS: 14,
    SESSOES_REALIZADAS: 15,
    SESSOES_RESTANTES: 16,
    STATUS: 21
  },

  COLUNAS_BLOQUEIOS: {
    DATA: 1,
    HORARIO: 2,
    FISIOTERAPEUTA: 3,
    ABRANGENCIA: 4,
    MOTIVO: 5,
    ACAO: 6,
    STATUS: 7
  },

  QUANTIDADE_COLUNAS_AGENDAMENTOS: 22,

  STATUS_PADRAO: [
    'Agendado',
    'Compareceu',
    'Falta Justificada',
    'Falta Não Justificada',
    'Cancelado por desistência'
  ]
};


/**
 * Abre a agenda do dia atual.
 */
function abrirAgendaHoje() {
  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0);

  configurarEstruturaAgendaDiaria_();
  definirDataAgendaDiaria_(hoje);
  carregarAgendaDiaria();
}


/**
 * Solicita uma data e abre a agenda daquele dia.
 */
function selecionarDataAgenda() {
  const ui = SpreadsheetApp.getUi();
  const hoje = new Date();

  const sugestao = Utilities.formatDate(
    hoje,
    Session.getScriptTimeZone(),
    'dd/MM/yyyy'
  );

  const resposta = ui.prompt(
    'Selecionar data da agenda',
    'Informe a data no formato DD/MM/AAAA.\n\n' +
      'Sugestão: ' +
      sugestao,
    ui.ButtonSet.OK_CANCEL
  );

  if (resposta.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const data = converterTextoParaDataAgendaDiaria_(
    resposta.getResponseText()
  );

  if (!data) {
    ui.alert(
      'Data inválida',
      'Informe a data no formato DD/MM/AAAA.',
      ui.ButtonSet.OK
    );
    return;
  }

  configurarEstruturaAgendaDiaria_();
  definirDataAgendaDiaria_(data);
  carregarAgendaDiaria();
}


/**
 * Exibe a agenda do dia anterior.
 */
function abrirDiaAnteriorAgenda() {
  configurarEstruturaAgendaDiaria_();

  const aba = obterAbaAgendaDiaria_();
  let data = aba.getRange('B2').getValue();

  if (!(data instanceof Date)) {
    data = new Date();
  }

  data = new Date(data);
  data.setDate(data.getDate() - 1);
  data.setHours(0, 0, 0, 0);

  definirDataAgendaDiaria_(data);
  carregarAgendaDiaria();
}


/**
 * Exibe a agenda do próximo dia.
 */
function abrirProximoDiaAgenda() {
  configurarEstruturaAgendaDiaria_();

  const aba = obterAbaAgendaDiaria_();
  let data = aba.getRange('B2').getValue();

  if (!(data instanceof Date)) {
    data = new Date();
  }

  data = new Date(data);
  data.setDate(data.getDate() + 1);
  data.setHours(0, 0, 0, 0);

  definirDataAgendaDiaria_(data);
  carregarAgendaDiaria();
}


/**
 * Atualiza a agenda usando a data e o fisioterapeuta
 * selecionados na própria aba.
 */
function carregarAgendaDiaria() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    configurarEstruturaAgendaDiaria_();

    const abas = obterAbasAgendaDiaria_(ss);
    const dataSelecionada = abas.agenda.getRange('B2').getValue();

    if (!(dataSelecionada instanceof Date)) {
      throw new Error(
        'Selecione uma data válida no campo Data.'
      );
    }

    let fisioterapeuta = String(
      abas.agenda.getRange('B3').getValue() || ''
    ).trim();

    if (!fisioterapeuta) {
      fisioterapeuta = 'Todos';
      abas.agenda.getRange('B3').setValue('Todos');
    }

    const registros = buscarRegistrosAgendaDiaria_(
      abas.agendamentos,
      abas.bloqueios,
      abas.horarios,
      dataSelecionada,
      fisioterapeuta
    );

    limparRegistrosAgendaDiaria_(abas.agenda);

    if (registros.length === 0) {
      registrarAgendaVazia_(
        abas.agenda,
        dataSelecionada,
        fisioterapeuta
      );
    } else {
      gravarRegistrosAgendaDiaria_(
        abas.agenda,
        abas.statusSessao,
        registros
      );
    }

    ss.setActiveSheet(abas.agenda);
    abas.agenda.getRange('A5').activate();

    ss.toast(
      registros.length +
        ' agendamento(s) encontrado(s).',
      'Agenda atualizada',
      4
    );
  } catch (erro) {
    ui.alert(
      'Erro ao carregar agenda',
      erro && erro.message
        ? erro.message
        : String(erro),
      ui.ButtonSet.OK
    );
  }
}


/**
 * Recria o layout operacional da aba Agenda.
 *
 * Esta função remove totalmente o modelo antigo,
 * inclusive dados existentes após a coluna F.
 */
function configurarEstruturaAgendaDiaria_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const aba = ss.getSheetByName(
    CONFIG_AGENDA_DIARIA.ABAS.AGENDA
  );

  const abaFisioterapeutas = ss.getSheetByName(
    CONFIG_AGENDA_DIARIA.ABAS.FISIOTERAPEUTAS
  );

  if (!aba) {
    throw new Error(
      'A aba "Agenda" não foi encontrada.'
    );
  }

  if (!abaFisioterapeutas) {
    throw new Error(
      'A aba "Fisioterapeutas" não foi encontrada.'
    );
  }

  let dataPreservada = aba.getRange('B2').getValue();
  let fisioterapeutaPreservado = String(
    aba.getRange('B3').getValue() || ''
  ).trim();

  if (!(dataPreservada instanceof Date)) {
    dataPreservada = new Date();
    dataPreservada.setHours(0, 0, 0, 0);
  }

  const fisioterapeutas =
    obterListaFisioterapeutasAgendaDiaria_(
      abaFisioterapeutas
    );

  const fisioterapeutaValido =
    fisioterapeutas.some(function(nome) {
      return (
        normalizarTextoAgendaDiaria_(nome) ===
        normalizarTextoAgendaDiaria_(
          fisioterapeutaPreservado
        )
      );
    });

  if (
    !fisioterapeutaPreservado ||
    !fisioterapeutaValido
  ) {
    fisioterapeutaPreservado = 'Todos';
  }

  try {
    aba.getDataRange().breakApart();
  } catch (erro) {
    // Continua caso não existam intervalos mesclados.
  }

  aba.clear();
  aba.setConditionalFormatRules([]);

  aba.getRange('A1:F1').merge();

  aba.getRange('A1')
    .setValue('AGENDA DA FISIOTERAPIA')
    .setFontSize(20)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground('#4f81bd')
    .setFontColor('#ffffff');

  aba.setRowHeight(1, 40);

  aba.getRange('A2')
    .setValue('Data')
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground('#4f81bd');

  aba.getRange('A3')
    .setValue('Fisioterapeuta')
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground('#4f81bd');

  aba.getRange('B2:F2')
    .setBackground('#4f81bd')
    .setFontColor('#ffffff');

  aba.getRange('B3:F3')
    .setBackground('#4f81bd')
    .setFontColor('#ffffff');

  aba.getRange('B2')
    .setValue(dataPreservada)
    .setNumberFormat('dd/MM/yyyy')
    .setHorizontalAlignment('center')
    .setFontSize(12);

  aba.getRange('B3')
    .setValue(fisioterapeutaPreservado)
    .setHorizontalAlignment('left')
    .setFontSize(12);

  aba.getRange('A4:F4')
    .setValues([[
      'Horário',
      'Paciente',
      'Evento/Sessão',
      'Status',
      'Alerta',
      'ID Agendamento'
    ]])
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground('#4f81bd')
    .setFontColor('#ffffff')
    .setWrap(true);

  aba.setRowHeight(4, 32);
  aba.setFrozenRows(4);

  aba.setColumnWidth(1, 100);
  aba.setColumnWidth(2, 240);
  aba.setColumnWidth(3, 150);
  aba.setColumnWidth(4, 190);
  aba.setColumnWidth(5, 210);
  aba.setColumnWidth(6, 135);

  if (aba.getMaxColumns() > 6) {
    aba.getRange(
      1,
      7,
      aba.getMaxRows(),
      aba.getMaxColumns() - 6
    )
      .clearContent()
      .clearFormat()
      .clearDataValidations();
  }

  configurarListaFisioterapeutasAgendaDiaria_(
    ss,
    aba
  );
}


/**
 * Obtém a lista de fisioterapeutas cadastrados.
 */function obterListaFisioterapeutasAgendaDiaria_(
  abaFisioterapeutas
) {
  const nomes = ['Todos'];
  const ultimaLinha = abaFisioterapeutas.getLastRow();

  if (ultimaLinha < 2) {
    return nomes;
  }

  const dados = abaFisioterapeutas
    .getRange(
      2,
      2,
      ultimaLinha - 1,
      1
    )
    .getDisplayValues();

  dados.forEach(function(linha) {
    const nome = String(linha[0] || '').trim();

    if (nome && nomes.indexOf(nome) === -1) {
      nomes.push(nome);
    }
  });

  return nomes;
}


/**
 * Cria a lista suspensa de fisioterapeutas.
 */
function configurarListaFisioterapeutasAgendaDiaria_(
  ss,
  abaAgenda
) {
  const abaFisioterapeutas = ss.getSheetByName(
    CONFIG_AGENDA_DIARIA.ABAS.FISIOTERAPEUTAS
  );

  if (!abaFisioterapeutas) {
    throw new Error(
      'A aba "Fisioterapeutas" não foi encontrada.'
    );
  }

  const nomes =
    obterListaFisioterapeutasAgendaDiaria_(
      abaFisioterapeutas
    );

  const validacao = SpreadsheetApp
    .newDataValidation()
    .requireValueInList(nomes, true)
    .setAllowInvalid(false)
    .setHelpText(
      'Selecione um fisioterapeuta ou a opção Todos.'
    )
    .build();

  abaAgenda.getRange('B3').setDataValidation(validacao);
}


/**
 * Define a data exibida pela agenda.
 */
function definirDataAgendaDiaria_(data) {
  const aba = obterAbaAgendaDiaria_();

  aba.getRange('B2')
    .setValue(new Date(data))
    .setNumberFormat('dd/MM/yyyy');
}


/**
 * Retorna as abas utilizadas pelo módulo.
 */
function obterAbasAgendaDiaria_(ss) {
  const abas = {
    agenda: ss.getSheetByName(
      CONFIG_AGENDA_DIARIA.ABAS.AGENDA
    ),

    agendamentos: ss.getSheetByName(
      CONFIG_AGENDA_DIARIA.ABAS.AGENDAMENTOS
    ),

    cadastro: ss.getSheetByName(
      CONFIG_AGENDA_DIARIA.ABAS.CADASTRO
    ),

    fisioterapeutas: ss.getSheetByName(
      CONFIG_AGENDA_DIARIA.ABAS.FISIOTERAPEUTAS
    ),

    statusSessao: ss.getSheetByName(
      CONFIG_AGENDA_DIARIA.ABAS.STATUS_SESSAO
    ),

    bloqueios: ss.getSheetByName(
      CONFIG_AGENDA_DIARIA.ABAS.BLOQUEIOS
    ),

    horarios: ss.getSheetByName(
      CONFIG_AGENDA_DIARIA.ABAS.HORARIOS
    )
  };

  Object.keys(abas).forEach(function(chave) {
    if (!abas[chave]) {
      throw new Error(
        'Uma aba necessária ao funcionamento da Agenda não foi encontrada.'
      );
    }
  });

  return abas;
}


/**
 * Retorna somente a aba Agenda.
 */
function obterAbaAgendaDiaria_() {
  const aba = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(
      CONFIG_AGENDA_DIARIA.ABAS.AGENDA
    );

  if (!aba) {
    throw new Error(
      'A aba "Agenda" não foi encontrada.'
    );
  }

  return aba;
}


/**
 * Busca os registros da data e do profissional selecionados.
 */
function buscarRegistrosAgendaDiaria_(
  abaAgendamentos,
  abaBloqueios,
  abaHorarios,
  dataSelecionada,
  fisioterapeutaSelecionado
) {
  const ultimaLinha = abaAgendamentos.getLastRow();

  const dados = ultimaLinha >= 2
    ? abaAgendamentos
        .getRange(
          2,
          1,
          ultimaLinha - 1,
          CONFIG_AGENDA_DIARIA
            .QUANTIDADE_COLUNAS_AGENDAMENTOS
        )
        .getValues()
    : [];

  const chaveData = chaveDataAgendaDiaria_(
    dataSelecionada
  );

  const filtroProfissional =
    normalizarTextoAgendaDiaria_(
      fisioterapeutaSelecionado
    );

  const todosProfissionais =
    !filtroProfissional ||
    filtroProfissional === 'todos';

  const registros = [];

  dados.forEach(function(linha) {
    const data = linha[
      CONFIG_AGENDA_DIARIA
        .COLUNAS_AGENDAMENTOS
        .DATA - 1
    ];

    if (!(data instanceof Date)) {
      return;
    }

    if (chaveDataAgendaDiaria_(data) !== chaveData) {
      return;
    }

    /*
     * Sessões remanejadas permanecem na aba Agendamentos
     * apenas como histórico. Elas não representam mais um
     * atendimento ativo e, portanto, não devem aparecer
     * na Agenda diária.
     */
    const statusRegistro =
      normalizarTextoAgendaDiaria_(
        linha[
          CONFIG_AGENDA_DIARIA
            .COLUNAS_AGENDAMENTOS
            .STATUS - 1
        ]
      );

    if (
      statusRegistro ===
        'cancelado por remanejamento'
    ) {
      return;
    }

    const fisioterapeuta = String(
      linha[
        CONFIG_AGENDA_DIARIA
          .COLUNAS_AGENDAMENTOS
          .FISIOTERAPEUTA - 1
      ] || ''
    ).trim();

    if (
      !todosProfissionais &&
      normalizarTextoAgendaDiaria_(fisioterapeuta) !==
        filtroProfissional
    ) {
      return;
    }

    const evento = String(
      linha[
        CONFIG_AGENDA_DIARIA
          .COLUNAS_AGENDAMENTOS
          .EVENTO - 1
      ] || ''
    ).trim();

    const numeroSessao = Number(
      linha[
        CONFIG_AGENDA_DIARIA
          .COLUNAS_AGENDAMENTOS
          .NUMERO_SESSAO - 1
      ]
    ) || 0;

    const totalPrescrito = Number(
      linha[
        CONFIG_AGENDA_DIARIA
          .COLUNAS_AGENDAMENTOS
          .TOTAL_PRESCRITO - 1
      ]
    ) || 0;

    const idCiclo = String(
      linha[
        CONFIG_AGENDA_DIARIA
          .COLUNAS_AGENDAMENTOS
          .ID_CICLO - 1
      ] || ''
    ).trim();

    const alerta = gerarAlertaAgendaDiaria_(
      dados,
      linha,
      idCiclo,
      numeroSessao,
      totalPrescrito
    );

    registros.push({
      horario: linha[
        CONFIG_AGENDA_DIARIA
          .COLUNAS_AGENDAMENTOS
          .HORARIO - 1
      ],

      paciente: String(
        linha[
          CONFIG_AGENDA_DIARIA
            .COLUNAS_AGENDAMENTOS
            .NOME_PACIENTE - 1
        ] || ''
      ).trim(),

      eventoSessao:
        formatarEventoSessaoAgendaDiaria_(
          evento,
          numeroSessao,
          totalPrescrito
        ),

      status: String(
        linha[
          CONFIG_AGENDA_DIARIA
            .COLUNAS_AGENDAMENTOS
            .STATUS - 1
        ] || 'Agendado'
      ).trim(),

      alerta: alerta,

      idAgendamento: String(
        linha[
          CONFIG_AGENDA_DIARIA
            .COLUNAS_AGENDAMENTOS
            .ID_AGENDAMENTO - 1
        ] || ''
      ).trim(),

      fisioterapeuta: fisioterapeuta
    });
  });

  adicionarBloqueiosAgendaDiaria_(
    registros,
    abaBloqueios,
    abaHorarios,
    dataSelecionada,
    fisioterapeutaSelecionado
  );

  registros.sort(function(a, b) {
    const horarioA =
      chaveHorarioAgendaDiaria_(a.horario);

    const horarioB =
      chaveHorarioAgendaDiaria_(b.horario);

    if (horarioA < horarioB) {
      return -1;
    }

    if (horarioA > horarioB) {
      return 1;
    }

    const profissional =
      a.fisioterapeuta.localeCompare(
        b.fisioterapeuta
      );

    if (profissional !== 0) {
      return profissional;
    }

    return a.paciente.localeCompare(b.paciente);
  });

  return registros;
}


/**
 * Formata o conteúdo da coluna Evento/Sessão.
 */
function formatarEventoSessaoAgendaDiaria_(
  evento,
  numeroSessao,
  totalPrescrito
) {
  const eventoNormalizado =
    normalizarTextoAgendaDiaria_(evento);

  if (
    eventoNormalizado === 'sessao' &&
    numeroSessao > 0 &&
    totalPrescrito > 0
  ) {
    return (
      'Sessão ' +
      numeroSessao +
      '/' +
      totalPrescrito
    );
  }

  return evento || 'Atendimento';
}


/**
 * Gera alertas para a agenda.
 */
function gerarAlertaAgendaDiaria_(
  todosAgendamentos,
  linhaAtual,
  idCiclo,
  numeroSessao,
  totalPrescrito
) {
  const evento = normalizarTextoAgendaDiaria_(
    linhaAtual[
      CONFIG_AGENDA_DIARIA
        .COLUNAS_AGENDAMENTOS
        .EVENTO - 1
    ]
  );

  if (evento !== 'sessao') {
    return '';
  }

  if (
    numeroSessao > 0 &&
    totalPrescrito > 0 &&
    numeroSessao === totalPrescrito
  ) {
    return 'Última sessão';
  }

  if (
    numeroSessao > 0 &&
    totalPrescrito > 0 &&
    totalPrescrito - numeroSessao === 1
  ) {
    return '1 sessão restante';
  }

  if (idCiclo) {
    let faltasNaoJustificadas = 0;

    todosAgendamentos.forEach(function(linha) {
      const cicloAtual = String(
        linha[
          CONFIG_AGENDA_DIARIA
            .COLUNAS_AGENDAMENTOS
            .ID_CICLO - 1
        ] || ''
      ).trim();

      const status =
        normalizarTextoAgendaDiaria_(
          linha[
            CONFIG_AGENDA_DIARIA
              .COLUNAS_AGENDAMENTOS
              .STATUS - 1
          ]
        );

      if (
        cicloAtual === idCiclo &&
        status === 'falta nao justificada'
      ) {
        faltasNaoJustificadas++;
      }
    });

    if (faltasNaoJustificadas >= 3) {
      return '3 faltas — verificar abandono';
    }

    if (faltasNaoJustificadas === 2) {
      return 'Atenção: 2 faltas';
    }
  }

  return '';
}


/**
 * Limpa os registros abaixo do cabeçalho.
function limparRegistrosAgendaDiaria_(aba) {
  const primeiraLinha =
    CONFIG_AGENDA_DIARIA
      .LINHAS
      .PRIMEIRO_REGISTRO;

  const quantidadeLinhas =
    Math.max(
      aba.getMaxRows() -
        primeiraLinha +
        1,
      1
    );

  const intervalo = aba.getRange(
    primeiraLinha,
    1,
    quantidadeLinhas,
    6
  );

  intervalo.breakApart();
  intervalo.clearContent();
  intervalo.clearFormat();
  intervalo.clearDataValidations();

  aba.setConditionalFormatRules([]);
}


/**
 * Grava os atendimentos na agenda.
 */
function gravarRegistrosAgendaDiaria_(
  abaAgenda,
  abaStatusSessao,
  registros
) {
  const linhas = registros.map(function(registro) {
    return [
      registro.horario,
      registro.paciente,
      registro.eventoSessao,
      registro.status,
      registro.alerta,
      registro.idAgendamento
    ];
  });

  const primeiraLinha =
    CONFIG_AGENDA_DIARIA.LINHAS.PRIMEIRO_REGISTRO;

  abaAgenda.getRange(
    primeiraLinha,
    1,
    linhas.length,
    6
  )
    .setValues(linhas)
    .setVerticalAlignment('middle')
    .setWrap(true);

  abaAgenda.getRange(
    primeiraLinha,
    1,
    linhas.length,
    1
  )
    .setNumberFormat('HH:mm')
    .setHorizontalAlignment('center');

  abaAgenda.getRange(
    primeiraLinha,
    3,
    linhas.length,
    1
  ).setHorizontalAlignment('center');

  abaAgenda.getRange(
    primeiraLinha,
    4,
    linhas.length,
    1
  ).setHorizontalAlignment('center');

  abaAgenda.getRange(
    primeiraLinha,
    6,
    linhas.length,
    1
  ).setHorizontalAlignment('center');

  abaAgenda.getRange(
    primeiraLinha,
    1,
    linhas.length,
    6
  ).setBorder(
    true,
    true,
    true,
    true,
    true,
    true,
    '#4f81bd',
    SpreadsheetApp.BorderStyle.SOLID
  );

  for (let i = 0; i < linhas.length; i++) {
    const numeroLinha = primeiraLinha + i;

    abaAgenda.setRowHeight(numeroLinha, 30);

    if (i % 2 === 0) {
      abaAgenda
        .getRange(numeroLinha, 1, 1, 6)
        .setBackground('#eeeeee');
    } else {
      abaAgenda
        .getRange(numeroLinha, 1, 1, 6)
        .setBackground('#ffffff');
    }
  }

  aplicarValidacaoStatusAgendaDiaria_(
    abaAgenda,
    abaStatusSessao,
    registros
  );

  aplicarCoresStatusAgendaDiaria_(
    abaAgenda,
    linhas.length
  );

  registros.forEach(
    function (registro, indice) {
      if (!registro.ehBloqueio) {
        return;
      }

      const numeroLinha =
        primeiraLinha + indice;

      abaAgenda
        .getRange(
          numeroLinha,
          CONFIG_AGENDA_DIARIA
            .COLUNAS_AGENDA.STATUS
        )
        .clearDataValidations();

      abaAgenda
        .getRange(
          numeroLinha,
          1,
          1,
          6
        )
        .setBackground('#d9d9d9')
        .setFontWeight('bold')
        .setFontColor('#444444');
    }
  );
}


/**
 * Cria a lista suspensa da coluna Status.
 *
 * A lista é aplicada somente aos atendimentos editáveis.
 * Bloqueios e cancelamentos internos permanecem visíveis,
 * mas sem lista suspensa.
 */
function aplicarValidacaoStatusAgendaDiaria_(
  abaAgenda,
  abaStatusSessao,
  registros
) {
  let statusPermitidos = [];

  const ultimaLinhaStatus =
    abaStatusSessao.getLastRow();

  if (ultimaLinhaStatus >= 2) {
    statusPermitidos = abaStatusSessao
      .getRange(
        2,
        1,
        ultimaLinhaStatus - 1,
        1
      )
      .getDisplayValues()
      .map(function(linha) {
        return String(linha[0] || '').trim();
      })
      .filter(function(status) {
        const normalizado =
          normalizarTextoAgendaDiaria_(
            status
          );

        return (
          status &&
          normalizado !== 'bloqueado' &&
          normalizado.indexOf(
            'cancelado'
          ) !== 0
        );
      });
  }

  if (statusPermitidos.length === 0) {
    statusPermitidos = [
      'Agendado',
      'Compareceu',
      'Falta Justificada',
      'Falta Não Justificada'
    ];
  }

  const validacao = SpreadsheetApp
    .newDataValidation()
    .requireValueInList(
      statusPermitidos,
      true
    )
    .setAllowInvalid(false)
    .setHelpText(
      'Selecione o resultado do atendimento.'
    )
    .build();

  registros.forEach(
    function(registro, indice) {
      const celula =
        abaAgenda.getRange(
          CONFIG_AGENDA_DIARIA
            .LINHAS
            .PRIMEIRO_REGISTRO +
            indice,
          CONFIG_AGENDA_DIARIA
            .COLUNAS_AGENDA
            .STATUS
        );

      const statusNormalizado =
        normalizarTextoAgendaDiaria_(
          registro.status
        );

      const statusInterno =
        statusNormalizado ===
          'bloqueado' ||
        statusNormalizado.indexOf(
          'cancelado'
        ) === 0;

      if (
        registro.ehBloqueio ||
        statusInterno
      ) {
        celula.clearDataValidations();
        return;
      }

      celula.setDataValidation(
        validacao
      );
    }
  );
}


/**
 * Aplica cores na coluna Status.
 */
function aplicarCoresStatusAgendaDiaria_(
  abaAgenda,
  quantidadeLinhas
) {
  const intervaloStatus = abaAgenda.getRange(
    CONFIG_AGENDA_DIARIA.LINHAS.PRIMEIRO_REGISTRO,
    CONFIG_AGENDA_DIARIA
      .COLUNAS_AGENDA
      .STATUS,
    quantidadeLinhas,
    1
  );

  const regras = [
    SpreadsheetApp
      .newConditionalFormatRule()
      .whenTextEqualTo('Agendado')
      .setBackground('#cfe2f3')
      .setRanges([intervaloStatus])
      .build(),

    SpreadsheetApp
      .newConditionalFormatRule()
      .whenTextEqualTo('Compareceu')
      .setBackground('#d9ead3')
      .setRanges([intervaloStatus])
      .build(),

    SpreadsheetApp
      .newConditionalFormatRule()
      .whenTextEqualTo('Falta Justificada')
      .setBackground('#fff2cc')
      .setRanges([intervaloStatus])
      .build(),

    SpreadsheetApp
      .newConditionalFormatRule()
      .whenTextEqualTo('Falta Não Justificada')
      .setBackground('#f4cccc')
      .setRanges([intervaloStatus])
      .build(),

    SpreadsheetApp
      .newConditionalFormatRule()
      .whenTextEqualTo('Cancelado pela Clínica')
      .setBackground('#d9d9d9')
      .setRanges([intervaloStatus])
      .build(),

    SpreadsheetApp
      .newConditionalFormatRule()
      .whenTextEqualTo('Bloqueado')
      .setBackground('#b7b7b7')
      .setRanges([intervaloStatus])
      .build()
  ];

  abaAgenda.setConditionalFormatRules(regras);
}


/**
 * Exibe mensagem quando não há atendimentos.
 */
function registrarAgendaVazia_(
  aba,
  data,
  fisioterapeuta
) {
  aba.getRange('A5:F5').merge();

  let texto =
    'Nenhum agendamento encontrado para ' +
    Utilities.formatDate(
      data,
      Session.getScriptTimeZone(),
      'dd/MM/yyyy'
    );

  if (
    fisioterapeuta &&
    normalizarTextoAgendaDiaria_(fisioterapeuta) !==
      'todos'
  ) {
    texto +=
      ' — Fisioterapeuta: ' +
      fisioterapeuta;
  }

  aba.getRange('A5')
    .setValue(texto)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setFontWeight('bold')
    .setBackground('#fff2cc');

  aba.setRowHeight(5, 35);
}


/**
 * Executado automaticamente quando uma célula é editada.
 */
function onEdit(e) {
  if (!e || !e.range) {
    return;
  }

  const aba = e.range.getSheet();

  if (
    aba.getName() !==
    CONFIG_AGENDA_DIARIA.ABAS.AGENDA
  ) {
    return;
  }

  const linha = e.range.getRow();
  const coluna = e.range.getColumn();

  if (
    (linha === 2 && coluna === 2) ||
    (linha === 3 && coluna === 2)
  ) {
    carregarAgendaDiaria();
    return;
  }

  if (
    linha >=
      CONFIG_AGENDA_DIARIA.LINHAS.PRIMEIRO_REGISTRO &&
    coluna ===
      CONFIG_AGENDA_DIARIA
        .COLUNAS_AGENDA
        .STATUS
  ) {
    sincronizarStatusAgendaDiaria_(e);
  }
}


/**
 * Sincroniza uma alteração feita na Agenda com
 * a aba Agendamentos.
 */function sincronizarStatusAgendaDiaria_(e) {
  const lock =
    LockService.getDocumentLock();

  let bloqueioObtido = false;

  const abaAgenda =
    e.range.getSheet();

  const linhaAgenda =
    e.range.getRow();

  const statusAnteriorAgenda =
    String(
      e.oldValue || ''
    ).trim();

  try {
    lock.waitLock(30000);
    bloqueioObtido = true;

    const novoStatus = String(
      e.range.getValue() || ''
    ).trim();

    const idAgendamento = String(
      abaAgenda.getRange(
        linhaAgenda,
        CONFIG_AGENDA_DIARIA
          .COLUNAS_AGENDA
          .ID_AGENDAMENTO
      ).getValue() || ''
    ).trim();

    if (!idAgendamento) {
      return;
    }

    if (!novoStatus) {
      throw new Error(
        'O status do atendimento não pode ficar vazio.'
      );
    }

    const ss =
      SpreadsheetApp.getActiveSpreadsheet();

    const abas =
      obterAbasAgendaDiaria_(ss);

    const registro =
      localizarAgendamentoPorIdAgendaDiaria_(
        abas.agendamentos,
        idAgendamento
      );

    if (!registro) {
      throw new Error(
        'O ID do agendamento não foi encontrado.'
      );
    }

    const statusOriginal =
      normalizarTextoAgendaDiaria_(
        registro.status
      );

    if (
      statusOriginal === 'bloqueado' ||
      statusOriginal.indexOf(
        'cancelado'
      ) === 0
    ) {
      throw new Error(
        'Este atendimento foi cancelado ou bloqueado pelo sistema e não pode ter o status alterado pela Agenda.'
      );
    }

    const linhaCadastro =
      localizarLinhaPacienteCadastroAgendaDiaria_(
        abas.cadastro,
        registro.idPaciente
      );

    if (!linhaCadastro) {
      throw new Error(
        'O paciente do atendimento não foi encontrado no Cadastro.'
      );
    }

    const statusPaciente = String(
      abas.cadastro.getRange(
        linhaCadastro,
        CONFIG_AGENDA_DIARIA
          .COLUNAS_CADASTRO
          .STATUS
      ).getValue() || ''
    ).trim();

    if (
      normalizarTextoAgendaDiaria_(
        statusPaciente
      ) === 'inativo'
    ) {
      throw new Error(
        'O paciente está Inativo. O resultado deste atendimento não pode ser alterado pela Agenda.'
      );
    }

    const novoStatusNormalizado =
      normalizarTextoAgendaDiaria_(
        novoStatus
      );

    const statusDeResultado = [
      'compareceu',
      'falta justificada',
      'falta nao justificada'
    ];

    if (
      statusDeResultado.indexOf(
        novoStatusNormalizado
      ) !== -1
    ) {
      if (
        !(
          registro.data
          instanceof Date
        )
      ) {
        throw new Error(
          'A data do atendimento não é válida.'
        );
      }

      const dataAtendimento =
        removerHorarioAgendaDiaria_(
          registro.data
        );

      const hoje =
        removerHorarioAgendaDiaria_(
          new Date()
        );

      if (
        dataAtendimento.getTime() >
          hoje.getTime()
      ) {
        throw new Error(
          'Não é permitido registrar presença ou falta em um atendimento futuro.'
        );
      }
    }

    const regras =
      obterRegrasStatusAgendaDiaria_(
        novoStatus,
        registro.evento
      );

    const agora =
      new Date();

    abas.agendamentos.getRange(
      registro.linha,
      CONFIG_AGENDA_DIARIA
        .COLUNAS_AGENDAMENTOS
        .STATUS
    ).setValue(novoStatus);

    abas.agendamentos.getRange(
      registro.linha,
      CONFIG_AGENDA_DIARIA
        .COLUNAS_AGENDAMENTOS
        .CONTA_COMO_SESSAO
    ).setValue(
      regras.contaComoSessao
    );

    abas.agendamentos.getRange(
      registro.linha,
      CONFIG_AGENDA_DIARIA
        .COLUNAS_AGENDAMENTOS
        .ATUALIZADO_EM
    )
      .setValue(agora)
      .setNumberFormat(
        'dd/MM/yyyy HH:mm'
      );

    abas.agendamentos.getRange(
      registro.linha,
      CONFIG_AGENDA_DIARIA
        .COLUNAS_AGENDAMENTOS
        .FATURAVEL
    ).setValue(
      regras.faturavel
    );

    recalcularCicloPacienteAgendaDiaria_(
      abas.agendamentos,
      abas.cadastro,
      registro.idPaciente,
      registro.idCiclo,
      registro.totalPrescrito
    );

    SpreadsheetApp.flush();

    carregarAgendaDiaria();

    ss.toast(
      'Status atualizado para "' +
        novoStatus +
        '".',
      'Agenda',
      4
    );
  } catch (erro) {
    try {
      abaAgenda
        .getRange(
          linhaAgenda,
          CONFIG_AGENDA_DIARIA
            .COLUNAS_AGENDA
            .STATUS
        )
        .setValue(
          statusAnteriorAgenda
        );
    } catch (erroRestauracao) {
      console.error(
        erroRestauracao
      );
    }

    SpreadsheetApp
      .getActiveSpreadsheet()
      .toast(
        erro && erro.message
          ? erro.message
          : String(erro),
        'Erro ao atualizar status',
        7
      );
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
 * Localiza uma linha da aba Agendamentos pelo ID.
 */
function localizarAgendamentoPorIdAgendaDiaria_(
  abaAgendamentos,
  idAgendamento
) {
  const ultimaLinha = abaAgendamentos.getLastRow();

  if (ultimaLinha < 2) {
    return null;
  }

  const dados = abaAgendamentos
    .getRange(
      2,
      1,
      ultimaLinha - 1,
      CONFIG_AGENDA_DIARIA
        .QUANTIDADE_COLUNAS_AGENDAMENTOS
    )
    .getValues();

  const idProcurado =
    normalizarTextoAgendaDiaria_(
      idAgendamento
    );

  for (let indice = 0; indice < dados.length; indice++) {
    const linha = dados[indice];

    const idAtual =
      normalizarTextoAgendaDiaria_(
        linha[
          CONFIG_AGENDA_DIARIA
            .COLUNAS_AGENDAMENTOS
            .ID_AGENDAMENTO - 1
        ]
      );

    if (idAtual === idProcurado) {
      return {
        linha: indice + 2,

        idPaciente: String(
          linha[
            CONFIG_AGENDA_DIARIA
              .COLUNAS_AGENDAMENTOS
              .ID_PACIENTE - 1
          ] || ''
        ).trim(),

        idCiclo: String(
          linha[
            CONFIG_AGENDA_DIARIA
              .COLUNAS_AGENDAMENTOS
              .ID_CICLO - 1
          ] || ''
        ).trim(),

        data:
          linha[
            CONFIG_AGENDA_DIARIA
              .COLUNAS_AGENDAMENTOS
              .DATA - 1
          ],

        evento: String(
          linha[
            CONFIG_AGENDA_DIARIA
              .COLUNAS_AGENDAMENTOS
              .EVENTO - 1
          ] || ''
        ).trim(),

        status: String(
          linha[
            CONFIG_AGENDA_DIARIA
              .COLUNAS_AGENDAMENTOS
              .STATUS - 1
          ] || ''
        ).trim(),

        totalPrescrito: Number(
          linha[
            CONFIG_AGENDA_DIARIA
              .COLUNAS_AGENDAMENTOS
              .TOTAL_PRESCRITO - 1
          ]
        ) || 0
      };
    }
  }

  return null;
}


/**
 * Define as regras de cada status.
 */
function obterRegrasStatusAgendaDiaria_(
  status,
  evento
) {
  const statusNormalizado =
    normalizarTextoAgendaDiaria_(status);

  const eventoNormalizado =
    normalizarTextoAgendaDiaria_(evento);

  const eSessao =
    eventoNormalizado === 'sessao';

  if (statusNormalizado === 'compareceu') {
    return {
      contaComoSessao: eSessao ? 'Sim' : 'Não',
      faturavel: 'Sim'
    };
  }

  if (
    statusNormalizado === 'falta justificada' ||
    statusNormalizado === 'falta nao justificada'
  ) {
    return {
      contaComoSessao: eSessao ? 'Sim' : 'Não',
      faturavel: 'Não'
    };
  }

  if (
    statusNormalizado === 'cancelado pela clinica'
  ) {
    return {
      contaComoSessao: 'Não',
      faturavel: 'Não'
    };
  }

  return {
    contaComoSessao: 'Não',
    faturavel: 'Não'
  };
}


/**
 * Recalcula sessões realizadas e restantes.
 */
function recalcularCicloPacienteAgendaDiaria_(
  abaAgendamentos,
  abaCadastro,
  idPaciente,
  idCiclo,
  totalPrescrito
) {
  if (!idPaciente || !idCiclo) {
    return;
  }

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
        CONFIG_AGENDA_DIARIA
          .QUANTIDADE_COLUNAS_AGENDAMENTOS
      )
      .getValues();

  let realizadas = 0;

  dados.forEach(function(linha) {
    const cicloAtual = String(
      linha[
        CONFIG_AGENDA_DIARIA
          .COLUNAS_AGENDAMENTOS
          .ID_CICLO - 1
      ] || ''
    ).trim();

    const evento =
      normalizarTextoAgendaDiaria_(
        linha[
          CONFIG_AGENDA_DIARIA
            .COLUNAS_AGENDAMENTOS
            .EVENTO - 1
        ]
      );

    const contaComoSessao =
      normalizarTextoAgendaDiaria_(
        linha[
          CONFIG_AGENDA_DIARIA
            .COLUNAS_AGENDAMENTOS
            .CONTA_COMO_SESSAO - 1
        ]
      );

    if (
      cicloAtual === idCiclo &&
      evento === 'sessao' &&
      contaComoSessao === 'sim'
    ) {
      realizadas++;
    }
  });

  const prescritas =
    Number(totalPrescrito) || 0;

  const restantes =
    Math.max(
      prescritas - realizadas,
      0
    );

  const linhaCadastro =
    localizarLinhaPacienteCadastroAgendaDiaria_(
      abaCadastro,
      idPaciente
    );

  if (!linhaCadastro) {
    return;
  }

  abaCadastro.getRange(
    linhaCadastro,
    CONFIG_AGENDA_DIARIA
      .COLUNAS_CADASTRO
      .SESSOES_REALIZADAS
  ).setValue(realizadas);

  abaCadastro.getRange(
    linhaCadastro,
    CONFIG_AGENDA_DIARIA
      .COLUNAS_CADASTRO
      .SESSOES_RESTANTES
  ).setValue(restantes);

  const statusAtual =
    normalizarTextoAgendaDiaria_(
      abaCadastro.getRange(
        linhaCadastro,
        CONFIG_AGENDA_DIARIA
          .COLUNAS_CADASTRO
          .STATUS
      ).getValue()
    );

  /*
   * Um paciente Inativo nunca pode ser reativado
   * por uma alteração feita na Agenda.
   */
  if (statusAtual === 'inativo') {
    return;
  }

  if (
    prescritas > 0 &&
    restantes === 0
  ) {
    abaCadastro.getRange(
      linhaCadastro,
      CONFIG_AGENDA_DIARIA
        .COLUNAS_CADASTRO
        .STATUS
    ).setValue(
      'Ciclo concluído'
    );
  } else {
    abaCadastro.getRange(
      linhaCadastro,
      CONFIG_AGENDA_DIARIA
        .COLUNAS_CADASTRO
        .STATUS
    ).setValue(
      'Em tratamento'
    );
  }
}


/**
 * Localiza o paciente no Cadastro pelo ID.
 */
function localizarLinhaPacienteCadastroAgendaDiaria_(
  abaCadastro,
  idPaciente
) {
  const ultimaLinha = abaCadastro.getLastRow();

  if (ultimaLinha < 2) {
    return null;
  }

  const ids = abaCadastro
    .getRange(
      2,
      CONFIG_AGENDA_DIARIA
        .COLUNAS_CADASTRO
        .ID_PACIENTE,
      ultimaLinha - 1,
      1
    )
    .getDisplayValues();

  const idProcurado =
    normalizarTextoAgendaDiaria_(
      idPaciente
    );

  for (let indice = 0; indice < ids.length; indice++) {
    const idAtual =
      normalizarTextoAgendaDiaria_(
        ids[indice][0]
      );

    if (idAtual === idProcurado) {
      return indice + 2;
    }
  }

  return null;
}


/**
 * Padroniza datas para comparação.
 */function chaveDataAgendaDiaria_(data) {
  return Utilities.formatDate(
    new Date(data),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );
}


/**
 * Padroniza horários para ordenação.
 */
function chaveHorarioAgendaDiaria_(valor) {
  if (valor instanceof Date) {
    return Utilities.formatDate(
      valor,
      Session.getScriptTimeZone(),
      'HH:mm'
    );
  }

  const texto = String(valor || '').trim();

  const resultado = texto.match(
    /(\d{1,2}):(\d{2})/
  );

  if (!resultado) {
    return texto;
  }

  return (
    String(Number(resultado[1])).padStart(2, '0') +
    ':' +
    resultado[2]
  );
}


/**
 * Converte DD/MM/AAAA para Date.
 */
function converterTextoParaDataAgendaDiaria_(texto) {
  const resultado = String(texto || '')
    .trim()
    .match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);

  if (!resultado) {
    return null;
  }

  const dia = Number(resultado[1]);
  const mes = Number(resultado[2]) - 1;
  const ano = Number(resultado[3]);

  const data = new Date(ano, mes, dia);

  if (
    data.getFullYear() !== ano ||
    data.getMonth() !== mes ||
    data.getDate() !== dia
  ) {
    return null;
  }

  data.setHours(0, 0, 0, 0);

  return data;
}


/**
 * Remove o componente de horário de uma data.
 *
 * Usado para comparar somente o dia do atendimento
 * com a data atual, sem considerar horas/minutos.
 */
function removerHorarioAgendaDiaria_(data) {
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
 * Padroniza textos para comparação.
 */
function normalizarTextoAgendaDiaria_(valor) {
  return String(valor || '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}


/**
 * Acrescenta à visualização da Agenda os horários
 * bloqueados na data selecionada.
 */
function adicionarBloqueiosAgendaDiaria_(
  registros,
  abaBloqueios,
  abaHorarios,
  dataSelecionada,
  fisioterapeutaSelecionado
) {
  if (
    !abaBloqueios ||
    abaBloqueios.getLastRow() < 2
  ) {
    return;
  }

  const horarios =
    lerHorariosParaBloqueiosAgendaDiaria_(
      abaHorarios
    );

  const dados = abaBloqueios
    .getRange(
      2,
      1,
      abaBloqueios.getLastRow() - 1,
      7
    )
    .getValues();

  const chaveData =
    chaveDataAgendaDiaria_(
      dataSelecionada
    );

  const filtroProfissional =
    normalizarTextoAgendaDiaria_(
      fisioterapeutaSelecionado
    );

  const todosProfissionais =
    !filtroProfissional ||
    filtroProfissional === 'todos';

  const chavesUsadas = {};

  dados.forEach(function (linha) {
    const colunas =
      CONFIG_AGENDA_DIARIA
        .COLUNAS_BLOQUEIOS;

    const data =
      linha[colunas.DATA - 1];

    const status =
      normalizarTextoAgendaDiaria_(
        linha[colunas.STATUS - 1]
      );

    if (
      !(data instanceof Date) ||
      chaveDataAgendaDiaria_(data) !==
        chaveData ||
      (
        status !== 'ativo' &&
        status !== 'bloqueado'
      )
    ) {
      return;
    }

    const profissional = String(
      linha[
        colunas.FISIOTERAPEUTA - 1
      ] || 'Todos'
    ).trim();

    const profissionalNormalizado =
      normalizarTextoAgendaDiaria_(
        profissional
      );

    if (
      !todosProfissionais &&
      profissionalNormalizado &&
      profissionalNormalizado !== 'todos' &&
      profissionalNormalizado !==
        filtroProfissional
    ) {
      return;
    }

    const abrangencia = String(
      linha[
        colunas.ABRANGENCIA - 1
      ] || 'Horário específico'
    ).trim();

    const motivo = String(
      linha[
        colunas.MOTIVO - 1
      ] || 'Bloqueio da clínica'
    ).trim();

    const horariosAfetados =
      obterHorariosAfetadosBloqueioAgendaDiaria_(
        linha[colunas.HORARIO - 1],
        abrangencia,
        horarios
      );

    horariosAfetados.forEach(
      function (horario) {
        const profissionalExibido =
          todosProfissionais
            ? (
                profissional ||
                'Todos'
              )
            : fisioterapeutaSelecionado;

        const chave =
          chaveHorarioAgendaDiaria_(
            horario
          ) +
          '|' +
          normalizarTextoAgendaDiaria_(
            profissionalExibido
          );

        if (chavesUsadas[chave]) {
          return;
        }

        chavesUsadas[chave] = true;

        let alerta =
          abrangencia +
          ' — ' +
          motivo;

        if (profissionalExibido) {
          alerta +=
            ' — ' +
            profissionalExibido;
        }

        registros.push({
          horario: horario,
          paciente:
            'ATENDIMENTO BLOQUEADO',
          eventoSessao: 'Bloqueio',
          status: 'Bloqueado',
          alerta: alerta,
          idAgendamento: '',
          fisioterapeuta:
            profissionalExibido || '',
          ehBloqueio: true
        });
      }
    );
  });
}


function lerHorariosParaBloqueiosAgendaDiaria_(
  abaHorarios
) {
  if (
    !abaHorarios ||
    abaHorarios.getLastRow() < 2
  ) {
    return [];
  }

  return abaHorarios
    .getRange(
      2,
      1,
      abaHorarios.getLastRow() - 1,
      2
    )
    .getValues()
    .map(function (linha) {
      return {
        valor: linha[0],
        chave:
          chaveHorarioAgendaDiaria_(
            linha[0]
          ),
        turno:
          normalizarTextoAgendaDiaria_(
            linha[1]
          )
      };
    })
    .filter(function (horario) {
      return Boolean(horario.chave);
    });
}


function obterHorariosAfetadosBloqueioAgendaDiaria_(
  horarioBloqueio,
  abrangencia,
  horarios
) {
  const abrangenciaNormalizada =
    normalizarTextoAgendaDiaria_(
      abrangencia
    );

  if (
    abrangenciaNormalizada ===
      'dia inteiro'
  ) {
    return horarios.map(
      function (horario) {
        return horario.valor;
      }
    );
  }

  if (
    abrangenciaNormalizada ===
      'turno inteiro'
  ) {
    const chaveBloqueio =
      chaveHorarioAgendaDiaria_(
        horarioBloqueio
      );

    const horarioReferencia =
      horarios.find(
        function (horario) {
          return (
            horario.chave ===
            chaveBloqueio
          );
        }
      );

    if (!horarioReferencia) {
      return horarioBloqueio
        ? [horarioBloqueio]
        : [];
    }

    return horarios
      .filter(function (horario) {
        return (
          horario.turno ===
          horarioReferencia.turno
        );
      })
      .map(function (horario) {
        return horario.valor;
      });
  }

  return horarioBloqueio
    ? [horarioBloqueio]
    : [];
}
function limparRegistrosAgendaDiaria_(aba) {
  const primeiraLinha =
    CONFIG_AGENDA_DIARIA
      .LINHAS
      .PRIMEIRO_REGISTRO;

  const quantidadeLinhas =
    Math.max(
      aba.getMaxRows() -
        primeiraLinha +
        1,
      1
    );

  const intervalo = aba.getRange(
    primeiraLinha,
    1,
    quantidadeLinhas,
    6
  );

  intervalo.breakApart();
  intervalo.clearContent();
  intervalo.clearFormat();
  intervalo.clearDataValidations();

  aba.setConditionalFormatRules([]);
}
