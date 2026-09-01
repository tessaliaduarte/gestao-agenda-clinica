const CONFIG_DESFECHO_TRATAMENTO = {
  ABAS: {
    CADASTRO: 'Cadastro de Pacientes',
    AGENDAMENTOS: 'Agendamentos',
    HISTORICO_DESFECHOS:
      'Histórico de Desfechos',
    STATUS_SESSAO: 'Status da Sessão'
  },

  CADASTRO: {
    ID_PACIENTE: 1,
    PRONTUARIO: 2,
    NOME: 3,
    CPF: 4,
    TELEFONE: 5,
    SESSOES_PRESCRITAS: 14,
    SESSOES_REALIZADAS: 15,
    SESSOES_RESTANTES: 16,
    STATUS: 21,
    FISIOTERAPEUTA: 22,
    DESFECHO: 24
  },

  AGENDAMENTOS: {
    ID_AGENDAMENTO: 1,
    ID_PACIENTE: 2,
    ID_CICLO: 5,
    CICLO_NUMERO: 6,
    DATA: 7,
    EVENTO: 12,
    STATUS: 16,
    MOTIVO: 17,
    CONTA_COMO_SESSAO: 18,
    AVISAR_PACIENTE: 19,
    ATUALIZADO_EM: 21,
    FATURAVEL: 22
  },

  QUANTIDADE_COLUNAS_CADASTRO: 24,
  QUANTIDADE_COLUNAS_AGENDAMENTOS: 22,
  QUANTIDADE_COLUNAS_HISTORICO: 14,
  LIMITE_RESULTADOS_PESQUISA: 30,

  STATUS_AGENDAMENTO_ORIGINAL:
    'Agendado',

  STATUS_AGENDAMENTO_ENCERRAMENTO:
    'Cancelado por encerramento',

  OPCOES: {
    '1': {
      desfecho: 'Alta',
      status: 'Inativo'
    },

    '2': {
      desfecho: 'Encaminhamento para APS',
      status: 'Inativo'
    },

    '3': {
      desfecho: 'Renovação',
      status: 'Ciclo concluído'
    },

    '4': {
      desfecho: 'Alta por abandono',
      status: 'Inativo'
    }
  }
};


/**
 * Mantém compatibilidade com o item atual do Menu.gs.
 * Agora a função abre uma página de pesquisa.
 */
function registrarDesfechoPacienteSelecionado() {
  abrirFormularioDesfechoTratamento();
}


/**
 * Abre a página de pesquisa e registro do desfecho.
 */
function abrirFormularioDesfechoTratamento() {
  const html =
    HtmlService
      .createHtmlOutputFromFile(
        'FormularioDesfecho'
      )
      .setWidth(780)
      .setHeight(720);

  SpreadsheetApp
    .getUi()
    .showModalDialog(
      html,
      'SIGAF — Registrar desfecho'
    );
}


/**
 * Pesquisa por nome, parte do nome, ID, prontuário ou CPF.
 */
function buscarPacientesParaDesfecho(
  termoBusca
) {
  const termoOriginal =
    String(
      termoBusca || ''
    ).trim();

  if (!termoOriginal) {
    throw new Error(
      'Informe o nome, ID, prontuário ou CPF do paciente.'
    );
  }

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const abaCadastro =
    obterAbaDesfecho_(
      ss,
      CONFIG_DESFECHO_TRATAMENTO
        .ABAS.CADASTRO
    );

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
        CONFIG_DESFECHO_TRATAMENTO
          .QUANTIDADE_COLUNAS_CADASTRO
      )
      .getValues();

  const colunas =
    CONFIG_DESFECHO_TRATAMENTO
      .CADASTRO;

  const termoNormalizado =
    normalizarTextoDesfecho_(
      termoOriginal
    );

  const termoNumerico =
    termoOriginal.replace(
      /\D/g,
      ''
    );

  const resultados = [];

  dados.forEach(function (linha) {
    const id = String(
      linha[
        colunas.ID_PACIENTE - 1
      ] || ''
    ).trim();

    if (!id) {
      return;
    }

    const prontuario = String(
      linha[
        colunas.PRONTUARIO - 1
      ] || ''
    ).trim();

    const nome = String(
      linha[
        colunas.NOME - 1
      ] || ''
    ).trim();

    const cpf = String(
      linha[
        colunas.CPF - 1
      ] || ''
    ).trim();

    const cpfNumerico =
      cpf.replace(
        /\D/g,
        ''
      );

    const corresponde =
      normalizarTextoDesfecho_(id) ===
        termoNormalizado ||

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

      normalizarTextoDesfecho_(
        nome
      ).indexOf(
        termoNormalizado
      ) !== -1;

    if (!corresponde) {
      return;
    }

    resultados.push({
      id: id,
      prontuario: prontuario,
      nome: nome,
      cpf: cpf,

      telefone: String(
        linha[
          colunas.TELEFONE - 1
        ] || ''
      ).trim(),

      status: String(
        linha[
          colunas.STATUS - 1
        ] || ''
      ).trim(),

      desfecho: String(
        linha[
          colunas.DESFECHO - 1
        ] || ''
      ).trim()
    });
  });

  resultados.sort(function (a, b) {
    return a.nome.localeCompare(
      b.nome,
      'pt-BR'
    );
  });

  return resultados.slice(
    0,
    CONFIG_DESFECHO_TRATAMENTO
      .LIMITE_RESULTADOS_PESQUISA
  );
}
/**
 * Carrega os dados completos do paciente escolhido.
 */
function obterPacienteParaDesfecho(
  idPaciente
) {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const abaCadastro =
    obterAbaDesfecho_(
      ss,
      CONFIG_DESFECHO_TRATAMENTO
        .ABAS.CADASTRO
    );

  const paciente =
    localizarPacienteDesfecho_(
      abaCadastro,
      idPaciente
    );

  if (!paciente) {
    throw new Error(
      'O paciente selecionado não foi encontrado.'
    );
  }

  const abaAgendamentos =
    obterAbaDesfecho_(
      ss,
      CONFIG_DESFECHO_TRATAMENTO
        .ABAS.AGENDAMENTOS
    );

  const ciclo =
    obterCicloAtualDesfecho_(
      abaAgendamentos,
      paciente.id
    );

  return {
    id: paciente.id,
    prontuario: paciente.prontuario,
    nome: paciente.nome,
    cpf: paciente.cpf,
    telefone: paciente.telefone,
    status: paciente.status,

    fisioterapeuta:
      paciente.fisioterapeuta,

    desfechoAtual:
      paciente.desfecho,

    sessoesPrescritas:
      paciente.sessoesPrescritas,

    sessoesRealizadas:
      paciente.sessoesRealizadas,

    sessoesRestantes:
      paciente.sessoesRestantes,

    idCiclo: ciclo.id,
    cicloNumero: ciclo.numero
  };
}


/**
 * Recebe a confirmação do formulário e registra o desfecho.
 */
function salvarDesfechoPaciente(
  entrada
) {
  const dados =
    entrada || {};

  const idPaciente =
    String(
      dados.idPaciente || ''
    ).trim();

  const opcao =
    interpretarOpcaoDesfecho_(
      dados.opcao
    );

  if (!idPaciente) {
    throw new Error(
      'Selecione um paciente.'
    );
  }

  if (!opcao) {
    throw new Error(
      'Selecione um desfecho válido.'
    );
  }

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const lock =
    LockService.getDocumentLock();

  let bloqueioObtido = false;

  try {
    lock.waitLock(30000);
    bloqueioObtido = true;

    const abaCadastro =
      obterAbaDesfecho_(
        ss,
        CONFIG_DESFECHO_TRATAMENTO
          .ABAS.CADASTRO
      );

    const paciente =
      localizarPacienteDesfecho_(
        abaCadastro,
        idPaciente
      );

    if (!paciente) {
      throw new Error(
        'O paciente selecionado não foi encontrado.'
      );
    }

    const abaAgendamentos =
      obterAbaDesfecho_(
        ss,
        CONFIG_DESFECHO_TRATAMENTO
          .ABAS.AGENDAMENTOS
      );

    const ciclo =
      obterCicloAtualDesfecho_(
        abaAgendamentos,
        paciente.id
      );

    const agora = new Date();

    const dataDesfecho =
      removerHorarioDesfecho_(
        agora
      );

    const abaHistorico =
      obterOuCriarHistoricoDesfecho_(
        ss
      );

    validarAusenciaDesfechoNoCiclo_(
      abaHistorico,
      paciente.id,
      ciclo
    );

    let quantidadeCancelada = 0;

    /*
     * Renovação preserva o ciclo encerrado.
     *
     * Alta, APS e alta por abandono cancelam
     * somente as sessões futuras ainda Agendadas.
     */
    if (
      normalizarTextoDesfecho_(
        opcao.desfecho
      ) !== 'renovacao'
    ) {
      garantirStatusEncerramentoDesfecho_(
        ss
      );

      quantidadeCancelada =
        cancelarSessoesFuturasPorDesfecho_(
          abaAgendamentos,
          paciente,
          ciclo,
          opcao,
          dataDesfecho,
          agora
        );
    }

    registrarHistoricoDesfecho_(
      abaHistorico,
      paciente,
      ciclo,
      opcao,
      dataDesfecho,
      agora,
      quantidadeCancelada
    );

    abaCadastro
      .getRange(
        paciente.linha,
        CONFIG_DESFECHO_TRATAMENTO
          .CADASTRO.DESFECHO
      )
      .setValue(
        opcao.desfecho
      );

    abaCadastro
      .getRange(
        paciente.linha,
        CONFIG_DESFECHO_TRATAMENTO
          .CADASTRO.STATUS
      )
      .setValue(
        opcao.status
      );

    SpreadsheetApp.flush();

    /*
     * A atualização das pendências é complementar.
     * Um erro nela não deve desfazer um desfecho
     * já registrado corretamente.
     */
    try {
      if (
        typeof atualizarPendenciasAutomaticas ===
          'function'
      ) {
        atualizarPendenciasAutomaticas();
      }
    } catch (erroPendencias) {
      console.warn(
        'O desfecho foi registrado, mas a atualização ' +
        'das pendências apresentou erro: ' +
        (
          erroPendencias &&
          erroPendencias.message
            ? erroPendencias.message
            : String(erroPendencias)
        )
      );
    }

    return {
      sucesso: true,
      paciente: paciente.nome,
      desfecho: opcao.desfecho,
      status: opcao.status,
      sessoesFuturasCanceladas:
        quantidadeCancelada,

      mensagem:
        'O desfecho foi registrado com sucesso.\n\n' +
        'Paciente: ' +
        paciente.nome +
        '\nDesfecho: ' +
        opcao.desfecho +
        '\nStatus: ' +
        opcao.status +
        (
          normalizarTextoDesfecho_(
            opcao.desfecho
          ) === 'renovacao'
            ? ''
            : (
                '\nSessões futuras canceladas: ' +
                quantidadeCancelada
              )
        ) +
        '\n\nO registro também foi incluído no Histórico de Desfechos.'

    };
  } catch (erro) {
    throw new Error(
      erro && erro.message
        ? erro.message
        : String(erro)
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


function localizarPacienteDesfecho_(
  abaCadastro,
  idPaciente
) {
  const ultimaLinha =
    abaCadastro.getLastRow();

  if (ultimaLinha < 2) {
    return null;
  }

  const dados =
    abaCadastro
      .getRange(
        2,
        1,
        ultimaLinha - 1,
        CONFIG_DESFECHO_TRATAMENTO
          .QUANTIDADE_COLUNAS_CADASTRO
      )
      .getValues();

  const procurado =
    normalizarTextoDesfecho_(
      idPaciente
    );

  for (
    let indice = 0;
    indice < dados.length;
    indice++
  ) {
    const linha = dados[indice];

    const id =
      normalizarTextoDesfecho_(
        linha[
          CONFIG_DESFECHO_TRATAMENTO
            .CADASTRO.ID_PACIENTE - 1
        ]
      );

    if (id !== procurado) {
      continue;
    }

    return montarPacienteDesfecho_(
      linha,
      indice + 2
    );
  }

  return null;
}


function montarPacienteDesfecho_(
  dados,
  linha
) {
  const colunas =
    CONFIG_DESFECHO_TRATAMENTO
      .CADASTRO;

  return {
    linha: linha,

    id: String(
      dados[
        colunas.ID_PACIENTE - 1
      ] || ''
    ).trim(),

    prontuario: String(
      dados[
        colunas.PRONTUARIO - 1
      ] || ''
    ).trim(),

    nome: String(
      dados[
        colunas.NOME - 1
      ] || ''
    ).trim(),

    cpf: String(
      dados[
        colunas.CPF - 1
      ] || ''
    ).trim(),

    telefone: String(
      dados[
        colunas.TELEFONE - 1
      ] || ''
    ).trim(),

    sessoesPrescritas:
      Number(
        dados[
          colunas.SESSOES_PRESCRITAS - 1
        ]
      ) || 0,

    sessoesRealizadas:
      Number(
        dados[
          colunas.SESSOES_REALIZADAS - 1
        ]
      ) || 0,

    sessoesRestantes:
      Number(
        dados[
          colunas.SESSOES_RESTANTES - 1
        ]
      ) || 0,

    status: String(
      dados[
        colunas.STATUS - 1
      ] || ''
    ).trim(),

    fisioterapeuta: String(
      dados[
        colunas.FISIOTERAPEUTA - 1
      ] || ''
    ).trim(),

    desfecho: String(
      dados[
        colunas.DESFECHO - 1
      ] || ''
    ).trim()
  };
}


function obterCicloAtualDesfecho_(
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
        CONFIG_DESFECHO_TRATAMENTO
          .QUANTIDADE_COLUNAS_AGENDAMENTOS
      )
      .getValues();

  const colunas =
    CONFIG_DESFECHO_TRATAMENTO
      .AGENDAMENTOS;

  const idProcurado =
    normalizarTextoDesfecho_(
      idPaciente
    );

  let numeroCiclo = 0;
  let idCiclo = '';

  dados.forEach(function (linha) {
    if (
      normalizarTextoDesfecho_(
        linha[
          colunas.ID_PACIENTE - 1
        ]
      ) !== idProcurado
    ) {
      return;
    }

    const numero =
      Number(
        linha[
          colunas.CICLO_NUMERO - 1
        ]
      ) || 0;

    const id =
      String(
        linha[
          colunas.ID_CICLO - 1
        ] || ''
      ).trim();

    if (
      numero > numeroCiclo ||
      (
        numero === numeroCiclo &&
        id &&
        !idCiclo
      )
    ) {
      numeroCiclo = numero;
      idCiclo = id;
    }
  });

  return {
    id: idCiclo,
    numero: numeroCiclo
  };
}
function registrarHistoricoDesfecho_(
  aba,
  paciente,
  ciclo,
  opcao,
  dataDesfecho,
  agora,
  quantidadeCancelada
) {
  const idRegistro =
    criarIdRegistroDesfecho_(
      paciente.id,
      ciclo,
      opcao.desfecho,
      dataDesfecho
    );

  if (
    existeIdHistoricoDesfecho_(
      aba,
      idRegistro
    )
  ) {
    throw new Error(
      'Este desfecho já foi registrado para o ciclo atual.'
    );
  }

  const novaLinha =
    Math.max(
      aba.getLastRow() + 1,
      2
    );

  aba
    .getRange(
      novaLinha,
      1,
      1,
      CONFIG_DESFECHO_TRATAMENTO
        .QUANTIDADE_COLUNAS_HISTORICO
    )
    .setValues([[
      idRegistro,
      paciente.id,
      paciente.prontuario,
      paciente.nome,
      ciclo.id,
      ciclo.numero || '',
      opcao.desfecho,
      dataDesfecho,
      paciente.sessoesPrescritas,
      paciente.sessoesRealizadas,
      paciente.sessoesRestantes,
      Number(quantidadeCancelada) || 0,
      'Registrado pelo menu SIGAF',
      agora
    ]]);

  aba
    .getRange(
      novaLinha,
      8
    )
    .setNumberFormat(
      'dd/MM/yyyy'
    );

  aba
    .getRange(
      novaLinha,
      9,
      1,
      4
    )
    .setNumberFormat('0');

  aba
    .getRange(
      novaLinha,
      14
    )
    .setNumberFormat(
      'dd/MM/yyyy HH:mm'
    );
}


/**
 * Impede dois desfechos finais diferentes
 * para o mesmo ciclo.
 */
function validarAusenciaDesfechoNoCiclo_(
  abaHistorico,
  idPaciente,
  ciclo
) {
  if (!ciclo || !ciclo.id) {
    return;
  }

  const ultimaLinha =
    abaHistorico.getLastRow();

  if (ultimaLinha < 2) {
    return;
  }

  const dados = abaHistorico
    .getRange(
      2,
      1,
      ultimaLinha - 1,
      CONFIG_DESFECHO_TRATAMENTO
        .QUANTIDADE_COLUNAS_HISTORICO
    )
    .getDisplayValues();

  const idPacienteProcurado =
    normalizarTextoDesfecho_(
      idPaciente
    );

  const idCicloProcurado =
    normalizarTextoDesfecho_(
      ciclo.id
    );

  const registroExistente =
    dados.find(function(linha) {
      return (
        normalizarTextoDesfecho_(
          linha[1]
        ) === idPacienteProcurado &&
        normalizarTextoDesfecho_(
          linha[4]
        ) === idCicloProcurado &&
        Boolean(
          String(linha[6] || '').trim()
        )
      );
    });

  if (registroExistente) {
    throw new Error(
      'O ciclo ' +
      ciclo.numero +
      ' já possui o desfecho "' +
      registroExistente[6] +
      '". Para alterar um desfecho já registrado, será necessário usar uma função própria de correção.'
    );
  }
}


/**
 * Cancela somente sessões futuras ainda Agendadas
 * quando o tratamento é encerrado.
 *
 * Altera apenas as colunas necessárias e não
 * regrava a linha inteira de Agendamentos.
 */
function cancelarSessoesFuturasPorDesfecho_(
  abaAgendamentos,
  paciente,
  ciclo,
  opcao,
  dataDesfecho,
  agora
) {
  if (!ciclo || !ciclo.id) {
    return 0;
  }

  const ultimaLinha =
    abaAgendamentos.getLastRow();

  if (ultimaLinha < 2) {
    return 0;
  }

  const quantidadeLinhas =
    ultimaLinha - 1;

  const dados = abaAgendamentos
    .getRange(
      2,
      1,
      quantidadeLinhas,
      CONFIG_DESFECHO_TRATAMENTO
        .QUANTIDADE_COLUNAS_AGENDAMENTOS
    )
    .getValues();

  const a =
    CONFIG_DESFECHO_TRATAMENTO
      .AGENDAMENTOS;

  const idPacienteProcurado =
    normalizarTextoDesfecho_(
      paciente.id
    );

  const idCicloProcurado =
    normalizarTextoDesfecho_(
      ciclo.id
    );

  const dataLimite =
    removerHorarioDesfecho_(
      dataDesfecho
    ).getTime();

  const motivo =
    opcao.desfecho +
    ' registrada em ' +
    Utilities.formatDate(
      dataDesfecho,
      Session.getScriptTimeZone(),
      'dd/MM/yyyy'
    );

  const linhasParaCancelar = [];

  dados.forEach(function(linha, indice) {
    const data =
      linha[a.DATA - 1];

    const corresponde =
      normalizarTextoDesfecho_(
        linha[a.ID_PACIENTE - 1]
      ) === idPacienteProcurado &&

      normalizarTextoDesfecho_(
        linha[a.ID_CICLO - 1]
      ) === idCicloProcurado &&

      Number(
        linha[a.CICLO_NUMERO - 1]
      ) === Number(ciclo.numero) &&

      normalizarTextoDesfecho_(
        linha[a.EVENTO - 1]
      ) === 'sessao' &&

      normalizarTextoDesfecho_(
        linha[a.STATUS - 1]
      ) ===
        normalizarTextoDesfecho_(
          CONFIG_DESFECHO_TRATAMENTO
            .STATUS_AGENDAMENTO_ORIGINAL
        ) &&

      data instanceof Date &&

      removerHorarioDesfecho_(
        data
      ).getTime() >= dataLimite;

    if (corresponde) {
      linhasParaCancelar.push(
        indice + 2
      );
    }
  });

  linhasParaCancelar.forEach(
    function(numeroLinha) {
      abaAgendamentos
        .getRange(
          numeroLinha,
          a.STATUS
        )
        .setValue(
          CONFIG_DESFECHO_TRATAMENTO
            .STATUS_AGENDAMENTO_ENCERRAMENTO
        );

      abaAgendamentos
        .getRange(
          numeroLinha,
          a.MOTIVO
        )
        .setValue(
          motivo
        );

      abaAgendamentos
        .getRange(
          numeroLinha,
          a.CONTA_COMO_SESSAO
        )
        .setValue('Não');

      abaAgendamentos
        .getRange(
          numeroLinha,
          a.AVISAR_PACIENTE
        )
        .setValue('Não');

      abaAgendamentos
        .getRange(
          numeroLinha,
          a.ATUALIZADO_EM
        )
        .setValue(agora)
        .setNumberFormat(
          'dd/MM/yyyy HH:mm'
        );

      abaAgendamentos
        .getRange(
          numeroLinha,
          a.FATURAVEL
        )
        .setValue('Não');
    }
  );

  return linhasParaCancelar.length;
}


/**
 * Garante que o novo status exista na lista
 * oficial da aba Status da Sessão.
 */
function garantirStatusEncerramentoDesfecho_(
  ss
) {
  const abaStatus =
    obterAbaDesfecho_(
      ss,
      CONFIG_DESFECHO_TRATAMENTO
        .ABAS.STATUS_SESSAO
    );

  const ultimaLinha =
    abaStatus.getLastRow();

  const valores =
    ultimaLinha >= 2
      ? abaStatus
          .getRange(
            2,
            1,
            ultimaLinha - 1,
            1
          )
          .getDisplayValues()
      : [];

  const procurado =
    normalizarTextoDesfecho_(
      CONFIG_DESFECHO_TRATAMENTO
        .STATUS_AGENDAMENTO_ENCERRAMENTO
    );

  const existe =
    valores.some(function(linha) {
      return (
        normalizarTextoDesfecho_(
          linha[0]
        ) === procurado
      );
    });

  if (!existe) {
    abaStatus
      .getRange(
        Math.max(
          ultimaLinha + 1,
          2
        ),
        1
      )
      .setValue(
        CONFIG_DESFECHO_TRATAMENTO
          .STATUS_AGENDAMENTO_ENCERRAMENTO
      );
  }
}


function obterOuCriarHistoricoDesfecho_(
  ss
) {
  const nomeAba =
    CONFIG_DESFECHO_TRATAMENTO
      .ABAS.HISTORICO_DESFECHOS;

  let aba =
    ss.getSheetByName(
      nomeAba
    );

  if (!aba) {
    aba = ss.insertSheet(
      nomeAba
    );
  }

  const cabecalhos = [
    'ID do Registro',
    'ID Paciente',
    'Prontuário',
    'Paciente',
    'ID Ciclo',
    'Ciclo Nº',
    'Desfecho',
    'Data do Desfecho',
    'Sessões Prescritas',
    'Sessões Realizadas',
    'Sessões Restantes',
    'Sessões Futuras Canceladas',
    'Motivo Informado',
    'Data e Hora do Registro'
  ];

  aba
    .getRange(
      1,
      1,
      1,
      cabecalhos.length
    )
    .setValues([cabecalhos])
    .setBackground('#4f81bd')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);

  aba.setFrozenRows(1);
  aba.setRowHeight(1, 42);

  return aba;
}


function existeIdHistoricoDesfecho_(
  aba,
  idRegistro
) {
  const ultimaLinha =
    aba.getLastRow();

  if (ultimaLinha < 2) {
    return false;
  }

  const procurado =
    normalizarTextoDesfecho_(
      idRegistro
    );

  return aba
    .getRange(
      2,
      1,
      ultimaLinha - 1,
      1
    )
    .getDisplayValues()
    .some(function (linha) {
      return (
        normalizarTextoDesfecho_(
          linha[0]
        ) === procurado
      );
    });
}


function criarIdRegistroDesfecho_(
  idPaciente,
  ciclo,
  desfecho,
  dataDesfecho
) {
  const cicloChave =
    ciclo.id ||
    (
      'SEM-CICLO-' +
      chaveDataDesfecho_(
        dataDesfecho
      )
    );

  return (
    'DESFECHO-' +
    limparIdDesfecho_(
      desfecho
    ) +
    '-' +
    limparIdDesfecho_(
      idPaciente
    ) +
    '-' +
    limparIdDesfecho_(
      cicloChave
    )
  );
}
function limparIdDesfecho_(
  valor
) {
  return normalizarTextoDesfecho_(
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


function interpretarOpcaoDesfecho_(
  valor
) {
  const texto =
    normalizarTextoDesfecho_(
      valor
    );

  if (
    CONFIG_DESFECHO_TRATAMENTO
      .OPCOES[texto]
  ) {
    return CONFIG_DESFECHO_TRATAMENTO
      .OPCOES[texto];
  }

  const equivalencias = {
    alta: '1',
    aps: '2',
    'encaminhamento para aps': '2',
    renovacao: '3',
    abandono: '4',
    'alta por abandono': '4'
  };

  const numero =
    equivalencias[texto];

  if (!numero) {
    return null;
  }

  return CONFIG_DESFECHO_TRATAMENTO
    .OPCOES[numero];
}


function obterAbaDesfecho_(
  ss,
  nome
) {
  const aba =
    ss.getSheetByName(nome);

  if (!aba) {
    throw new Error(
      'A aba necessária "' +
        nome +
        '" não foi encontrada.'
    );
  }

  return aba;
}


function removerHorarioDesfecho_(
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


function chaveDataDesfecho_(
  data
) {
  return Utilities.formatDate(
    new Date(data),
    Session.getScriptTimeZone(),
    'yyyyMMdd'
  );
}


function normalizarTextoDesfecho_(
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
