const CONFIG_DESISTENCIA_TRATAMENTO = {
  ABAS: {
    CADASTRO: 'Cadastro de Pacientes',
    AGENDAMENTOS: 'Agendamentos',
    PENDENCIAS: 'Pendências',
    HISTORICO_PENDENCIAS:
      'Histórico de Pendências',
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
    DATA_INICIO: 19,
    DATA_TERMINO: 20,
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

  PENDENCIAS: {
    ID: 1,
    QUANTIDADE_COLUNAS: 11
  },

  STATUS_PERMITIDOS: [
    'em tratamento',
    'avaliado - aguardando agendamento'
  ],

  STATUS_AGENDAMENTO_ORIGINAL:
    'agendado',

  STATUS_AGENDAMENTO_DESISTENCIA:
    'Cancelado por desistência',

  STATUS_CADASTRO:
    'Inativo',

  DESFECHO:
    'Desistência do tratamento',

  QUANTIDADE_COLUNAS_CADASTRO: 24,
  QUANTIDADE_COLUNAS_AGENDAMENTOS: 22
};


/**
 * Abre a página de pesquisa da desistência.
 */
function abrirFormularioDesistenciaTratamento() {
  const html = HtmlService
    .createHtmlOutputFromFile(
      'FormularioDesistencia'
    )
    .setWidth(760)
    .setHeight(720);

  SpreadsheetApp.getUi().showModalDialog(
    html,
    'SIGAF — Registrar desistência'
  );
}


/**
 * Pesquisa pacientes pelo ID, prontuário,
 * CPF, nome completo ou parte do nome.
 */
function buscarPacientesParaDesistencia(
  termoBusca
) {
  const termoOriginal = String(
    termoBusca || ''
  ).trim();

  if (!termoOriginal) {
    throw new Error(
      'Informe o ID, prontuário, CPF ou nome do paciente.'
    );
  }

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const abaCadastro =
    obterAbaDesistencia_(
      ss,
      CONFIG_DESISTENCIA_TRATAMENTO
        .ABAS
        .CADASTRO
    );

  const ultimaLinha =
    abaCadastro.getLastRow();

  if (ultimaLinha < 2) {
    return [];
  }

  const c =
    CONFIG_DESISTENCIA_TRATAMENTO
      .CADASTRO;

  const dados = abaCadastro
    .getRange(
      2,
      1,
      ultimaLinha - 1,
      CONFIG_DESISTENCIA_TRATAMENTO
        .QUANTIDADE_COLUNAS_CADASTRO
    )
    .getValues();

  const termoNormalizado =
    normalizarTextoDesistencia_(
      termoOriginal
    );

  const termoNumerico =
    termoOriginal.replace(/\D/g, '');

  const resultados = [];

  dados.forEach(function(linha) {
    const id = String(
      linha[c.ID_PACIENTE - 1] || ''
    ).trim();

    if (!id) {
      return;
    }

    const prontuario = String(
      linha[c.PRONTUARIO - 1] || ''
    ).trim();

    const nome = String(
      linha[c.NOME - 1] || ''
    ).trim();

    const cpf = String(
      linha[c.CPF - 1] || ''
    ).trim();

    const cpfNumerico =
      cpf.replace(/\D/g, '');

    const corresponde =
      normalizarTextoDesistencia_(id) ===
        termoNormalizado ||
      prontuario === termoOriginal ||
      (
        termoNumerico &&
        prontuario === termoNumerico
      ) ||
      (
        termoNumerico &&
        cpfNumerico === termoNumerico
      ) ||
      normalizarTextoDesistencia_(nome)
        .indexOf(termoNormalizado) !== -1;

    if (!corresponde) {
      return;
    }

    resultados.push({
      id: id,

      prontuario:
        prontuario,

      nome:
        nome,

      cpf:
        cpf,

      status: String(
        linha[c.STATUS - 1] || ''
      ).trim(),

      desfecho: String(
        linha[c.DESFECHO - 1] || ''
      ).trim(),

      sessoesPrescritas:
        Number(
          linha[
            c.SESSOES_PRESCRITAS - 1
          ]
        ) || 0,

      sessoesRealizadas:
        Number(
          linha[
            c.SESSOES_REALIZADAS - 1
          ]
        ) || 0,

      sessoesRestantes:
        Number(
          linha[
            c.SESSOES_RESTANTES - 1
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

  return resultados.slice(0, 30);
}


/**
 * Carrega e valida o paciente escolhido.
 */
function obterPacienteParaDesistencia(
  idPaciente
) {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const abaCadastro =
    obterAbaDesistencia_(
      ss,
      CONFIG_DESISTENCIA_TRATAMENTO
        .ABAS
        .CADASTRO
    );

  const paciente =
    localizarPacienteDesistencia_(
      abaCadastro,
      idPaciente
    );

  /*
   * Se uma tentativa anterior concluiu a desistência,
   * informa isso sem exigir novamente um paciente ativo.
   */
  if (
    paciente &&
    desistenciaJaRegistrada_(
      paciente
    )
  ) {
    throw new Error(
      'A desistência deste paciente já foi registrada. ' +
      'O paciente está Inativo e o desfecho é "' +
      CONFIG_DESISTENCIA_TRATAMENTO.DESFECHO +
      '".'
    );
  }

  validarPacienteDesistencia_(
    paciente
  );

  const abaAgendamentos =
    obterAbaDesistencia_(
      ss,
      CONFIG_DESISTENCIA_TRATAMENTO
        .ABAS
        .AGENDAMENTOS
    );

  const ciclo =
    obterCicloAtualDesistencia_(
      abaAgendamentos,
      paciente.id
    );

  if (!ciclo.numero) {
    throw new Error(
      'Não foi encontrado um ciclo do paciente na aba "Agendamentos".'
    );
  }

  return {
    idPaciente:
      paciente.id,

    prontuario:
      paciente.prontuario,

    nome:
      paciente.nome,

    telefone:
      paciente.telefone,

    status:
      paciente.status,

    fisioterapeuta:
      paciente.fisioterapeuta,

    sessoesPrescritas:
      paciente.sessoesPrescritas,

    sessoesRealizadas:
      paciente.sessoesRealizadas,

    sessoesRestantes:
      paciente.sessoesRestantes,

    cicloNumero:
      ciclo.numero,

    idCiclo:
      ciclo.id,

    dataInicio:
      formatarDataDesistencia_(
        paciente.dataInicio
      ),

    dataTermino:
      formatarDataDesistencia_(
        paciente.dataTermino
      )
  };
}


/**
 * Registra a desistência, cancela somente
 * as sessões futuras e preserva o histórico.
 */
function registrarDesistenciaTratamento(
  dados
) {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const lock =
    LockService.getDocumentLock();

  let bloqueioObtido = false;

  try {
    lock.waitLock(30000);
    bloqueioObtido = true;

    const entrada =
      validarEntradaDesistencia_(
        dados
      );

    const abaCadastro =
      obterAbaDesistencia_(
        ss,
        CONFIG_DESISTENCIA_TRATAMENTO
          .ABAS
          .CADASTRO
      );

    const paciente =
      localizarPacienteDesistencia_(
        abaCadastro,
        entrada.idPaciente
      );

    if (!paciente || !paciente.id) {
      throw new Error(
        'Paciente não encontrado.'
      );
    }

    /*
     * Proteção contra registro duplicado.
     *
     * Isso também permite recuperar uma tentativa anterior
     * que tenha concluído a desistência, mas apresentado erro
     * durante a atualização geral das pendências.
     */
    if (
      desistenciaJaRegistrada_(
        paciente
      )
    ) {
      return {
        sucesso: true,
        jaRegistrada: true,
        quantidadeCancelada: 0,

        mensagem:
          'A desistência de ' +
          paciente.nome +
          ' já estava registrada.\n\n' +
          'Status atual: Inativo\n' +
          'Desfecho: ' +
          CONFIG_DESISTENCIA_TRATAMENTO
            .DESFECHO +
          '\n\nNenhuma informação foi duplicada.'
      };
    }

    validarPacienteDesistencia_(
      paciente
    );

    const abaAgendamentos =
      obterAbaDesistencia_(
        ss,
        CONFIG_DESISTENCIA_TRATAMENTO
          .ABAS
          .AGENDAMENTOS
      );

    const ciclo =
      obterCicloAtualDesistencia_(
        abaAgendamentos,
        paciente.id
      );

    if (!ciclo.numero) {
      throw new Error(
        'Não foi encontrado um ciclo para registrar a desistência.'
      );
    }

    if (
      entrada.cicloNumero !==
        ciclo.numero ||
      entrada.idCiclo !==
        ciclo.id
    ) {
      throw new Error(
        'O ciclo do paciente mudou. Feche o formulário e abra-o novamente.'
      );
    }

    garantirStatusSessaoDesistencia_(
      ss
    );

    const quantidadeCancelada =
      cancelarSessoesFuturasDesistencia_(
        abaAgendamentos,
        paciente.id,
        ciclo,
        entrada.dataDesistencia,
        entrada.motivo
      );

    const c =
      CONFIG_DESISTENCIA_TRATAMENTO
        .CADASTRO;

    abaCadastro
      .getRange(
        paciente.linha,
        c.STATUS
      )
      .setValue(
        CONFIG_DESISTENCIA_TRATAMENTO
          .STATUS_CADASTRO
      );

    abaCadastro
      .getRange(
        paciente.linha,
        c.DESFECHO
      )
      .setValue(
        CONFIG_DESISTENCIA_TRATAMENTO
          .DESFECHO
      );

    registrarHistoricoDesistencia_(
      ss,
      paciente,
      ciclo,
      entrada,
      quantidadeCancelada
    );

    criarPendenciaEncerramentoDesistencia_(
      ss,
      paciente,
      ciclo,
      entrada.dataDesistencia
    );

    SpreadsheetApp.flush();

    /*
     * A atualização geral das pendências é complementar.
     *
     * Ela não pode desfazer nem interromper uma desistência
     * que já foi registrada corretamente.
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
        'A desistência foi registrada, mas o atualizador geral ' +
        'de pendências apresentou um erro: ' +
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

      quantidadeCancelada:
        quantidadeCancelada,

      mensagem:
        'A desistência de ' +
        paciente.nome +
        ' foi registrada.\n\n' +
        'Ciclo: ' +
        ciclo.numero +
        '\nSessões futuras canceladas: ' +
        quantidadeCancelada +
        '\n\n' +
        'As sessões anteriores e as contagens do ciclo foram preservadas. O paciente agora está Inativo.\n\n' +
        'Foi criada a pendência "Registrar encerramento por desistência" para o fisioterapeuta.'
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


/**
 * Verifica se a desistência já foi registrada.
 */
function desistenciaJaRegistrada_(
  paciente
) {
  if (!paciente) {
    return false;
  }

  return (
    normalizarTextoDesistencia_(
      paciente.status
    ) ===
      normalizarTextoDesistencia_(
        CONFIG_DESISTENCIA_TRATAMENTO
          .STATUS_CADASTRO
      ) &&
    normalizarTextoDesistencia_(
      paciente.desfecho
    ) ===
      normalizarTextoDesistencia_(
        CONFIG_DESISTENCIA_TRATAMENTO
          .DESFECHO
      )
  );
}


/**
 * Localiza o paciente pelo ID.
 */
function localizarPacienteDesistencia_(
  abaCadastro,
  idPaciente
) {
  const idProcurado = String(
    idPaciente || ''
  ).trim();

  if (!idProcurado) {
    throw new Error(
      'O ID do paciente não foi informado.'
    );
  }

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
      CONFIG_DESISTENCIA_TRATAMENTO
        .QUANTIDADE_COLUNAS_CADASTRO
    )
    .getValues();

  const c =
    CONFIG_DESISTENCIA_TRATAMENTO
      .CADASTRO;

  const idNormalizado =
    normalizarTextoDesistencia_(
      idProcurado
    );

  for (
    let indice = 0;
    indice < dados.length;
    indice++
  ) {
    const linha = dados[indice];

    if (
      normalizarTextoDesistencia_(
        linha[c.ID_PACIENTE - 1]
      ) !== idNormalizado
    ) {
      continue;
    }

    return {
      linha:
        indice + 2,

      id: String(
        linha[
          c.ID_PACIENTE - 1
        ] || ''
      ).trim(),

      prontuario: String(
        linha[
          c.PRONTUARIO - 1
        ] || ''
      ).trim(),

      nome: String(
        linha[
          c.NOME - 1
        ] || ''
      ).trim(),

      telefone: String(
        linha[
          c.TELEFONE - 1
        ] || ''
      ).trim(),

      sessoesPrescritas:
        Number(
          linha[
            c.SESSOES_PRESCRITAS - 1
          ]
        ) || 0,

      sessoesRealizadas:
        Number(
          linha[
            c.SESSOES_REALIZADAS - 1
          ]
        ) || 0,

      sessoesRestantes:
        Number(
          linha[
            c.SESSOES_RESTANTES - 1
          ]
        ) || 0,

      dataInicio:
        linha[
          c.DATA_INICIO - 1
        ],

      dataTermino:
        linha[
          c.DATA_TERMINO - 1
        ],

      status: String(
        linha[
          c.STATUS - 1
        ] || ''
      ).trim(),

      fisioterapeuta: String(
        linha[
          c.FISIOTERAPEUTA - 1
        ] || ''
      ).trim(),

      desfecho: String(
        linha[
          c.DESFECHO - 1
        ] || ''
      ).trim()
    };
  }

  return null;
}


/**
 * Confere se o paciente está apto
 * para registrar desistência.
 */
function validarPacienteDesistencia_(
  paciente
) {
  if (!paciente || !paciente.id) {
    throw new Error(
      'Paciente não encontrado.'
    );
  }

  const status =
    normalizarTextoDesistencia_(
      paciente.status
    );

  if (
    CONFIG_DESISTENCIA_TRATAMENTO
      .STATUS_PERMITIDOS
      .indexOf(status) === -1
  ) {
    throw new Error(
      'A desistência pode ser registrada somente para pacientes "Em tratamento" ou "Avaliado – aguardando agendamento". Status atual: ' +
      (paciente.status || 'não informado') +
      '.'
    );
  }
}


/**
 * Valida as informações recebidas.
 */
function validarEntradaDesistencia_(
  dados
) {
  if (!dados || typeof dados !== 'object') {
    throw new Error(
      'Os dados da desistência não foram recebidos.'
    );
  }

  const idPaciente = String(
    dados.idPaciente || ''
  ).trim();

  const idCiclo = String(
    dados.idCiclo || ''
  ).trim();

  const cicloNumero =
    Number(dados.cicloNumero);

  const dataDesistencia =
    converterDataDesistencia_(
      dados.dataDesistencia
    );

  if (!idPaciente) {
    throw new Error(
      'O ID do paciente não foi informado.'
    );
  }

  if (
    !idCiclo ||
    !Number.isInteger(cicloNumero) ||
    cicloNumero < 1
  ) {
    throw new Error(
      'O ciclo do paciente é inválido.'
    );
  }

  if (!dataDesistencia) {
    throw new Error(
      'Informe uma data de desistência válida.'
    );
  }

  const hoje = new Date();

  hoje.setHours(0, 0, 0, 0);

  if (
    dataDesistencia.getTime() >
      hoje.getTime()
  ) {
    throw new Error(
      'A data da desistência não pode estar no futuro.'
    );
  }

  const motivo = String(
    dados.motivo || ''
  ).trim();

  if (motivo.length > 500) {
    throw new Error(
      'O motivo deve possuir no máximo 500 caracteres.'
    );
  }

  return {
    idPaciente:
      idPaciente,

    idCiclo:
      idCiclo,

    cicloNumero:
      cicloNumero,

    dataDesistencia:
      dataDesistencia,

    motivo:
      motivo
  };
}


/**
 * Identifica o ciclo mais recente do paciente.
 */
function obterCicloAtualDesistencia_(
  abaAgendamentos,
  idPaciente
) {
  const ultimaLinha =
    abaAgendamentos.getLastRow();

  if (ultimaLinha < 2) {
    return {
      numero: 0,
      id: ''
    };
  }

  const dados = abaAgendamentos
    .getRange(
      2,
      1,
      ultimaLinha - 1,
      CONFIG_DESISTENCIA_TRATAMENTO
        .QUANTIDADE_COLUNAS_AGENDAMENTOS
    )
    .getValues();

  const a =
    CONFIG_DESISTENCIA_TRATAMENTO
      .AGENDAMENTOS;

  const idProcurado =
    normalizarTextoDesistencia_(
      idPaciente
    );

  let maiorCiclo = 0;
  let idCiclo = '';

  dados.forEach(function(linha) {
    if (
      normalizarTextoDesistencia_(
        linha[
          a.ID_PACIENTE - 1
        ]
      ) !== idProcurado
    ) {
      return;
    }

    const numero = Number(
      linha[
        a.CICLO_NUMERO - 1
      ]
    );

    if (
      Number.isFinite(numero) &&
      numero > maiorCiclo
    ) {
      maiorCiclo = numero;

      idCiclo = String(
        linha[
          a.ID_CICLO - 1
        ] || ''
      ).trim();
    }
  });

  return {
    numero:
      maiorCiclo,

    id:
      idCiclo
  };
}


/**
 * Cancela somente sessões futuras ainda
 * com status Agendado no ciclo atual.
 *
 * Esta versão altera apenas as colunas necessárias,
 * evitando regravar a coluna Tipo do Grupo e outras
 * colunas que possuem validações.
 */
function cancelarSessoesFuturasDesistencia_(
  abaAgendamentos,
  idPaciente,
  ciclo,
  dataDesistencia,
  motivoInformado
) {
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
      CONFIG_DESISTENCIA_TRATAMENTO
        .QUANTIDADE_COLUNAS_AGENDAMENTOS
    )
    .getValues();

  const a =
    CONFIG_DESISTENCIA_TRATAMENTO
      .AGENDAMENTOS;

  const idProcurado =
    normalizarTextoDesistencia_(
      idPaciente
    );

  const idCicloProcurado =
    normalizarTextoDesistencia_(
      ciclo.id
    );

  const dataLimite =
    removerHorarioDesistencia_(
      dataDesistencia
    ).getTime();

  const agora = new Date();

  const motivoBase =
    'Desistência do tratamento em ' +
    formatarDataDesistencia_(
      dataDesistencia
    );

  const motivoFinal = motivoInformado
    ? motivoBase +
      ' — Motivo informado: ' +
      motivoInformado
    : motivoBase;

  const linhasParaCancelar = [];

  dados.forEach(function(linha, indice) {
    const data =
      linha[a.DATA - 1];

    const corresponde =
      normalizarTextoDesistencia_(
        linha[
          a.ID_PACIENTE - 1
        ]
      ) === idProcurado &&

      normalizarTextoDesistencia_(
        linha[
          a.ID_CICLO - 1
        ]
      ) === idCicloProcurado &&

      Number(
        linha[
          a.CICLO_NUMERO - 1
        ]
      ) === ciclo.numero &&

      normalizarTextoDesistencia_(
        linha[
          a.EVENTO - 1
        ]
      ) === 'sessao' &&

      normalizarTextoDesistencia_(
        linha[
          a.STATUS - 1
        ]
      ) ===
        CONFIG_DESISTENCIA_TRATAMENTO
          .STATUS_AGENDAMENTO_ORIGINAL &&

      dataValidaDesistencia_(data) &&

      removerHorarioDesistencia_(
        data
      ).getTime() >= dataLimite;

    if (!corresponde) {
      return;
    }

    linhasParaCancelar.push(
      indice + 2
    );
  });

  linhasParaCancelar.forEach(
    function(numeroLinha) {
      /*
       * Status.
       */
      abaAgendamentos
        .getRange(
          numeroLinha,
          a.STATUS
        )
        .setValue(
          CONFIG_DESISTENCIA_TRATAMENTO
            .STATUS_AGENDAMENTO_DESISTENCIA
        );

      /*
       * Motivo.
       */
      abaAgendamentos
        .getRange(
          numeroLinha,
          a.MOTIVO
        )
        .setValue(
          motivoFinal
        );

      /*
       * Conta como sessão.
       */
      abaAgendamentos
        .getRange(
          numeroLinha,
          a.CONTA_COMO_SESSAO
        )
        .setValue(
          'Não'
        );

      /*
       * Avisar paciente.
       */
      abaAgendamentos
        .getRange(
          numeroLinha,
          a.AVISAR_PACIENTE
        )
        .setValue(
          'Não'
        );

      /*
       * Atualizado em.
       */
      abaAgendamentos
        .getRange(
          numeroLinha,
          a.ATUALIZADO_EM
        )
        .setValue(
          agora
        )
        .setNumberFormat(
          'dd/MM/yyyy HH:mm'
        );

      /*
       * Faturável.
       */
      abaAgendamentos
        .getRange(
          numeroLinha,
          a.FATURAVEL
        )
        .setValue(
          'Não'
        );
    }
  );

  return linhasParaCancelar.length;
}


/**
 * Registra a desistência em histórico próprio.
 */
function registrarHistoricoDesistencia_(
  ss,
  paciente,
  ciclo,
  entrada,
  quantidadeCancelada
) {
  const nomeAba =
    CONFIG_DESISTENCIA_TRATAMENTO
      .ABAS
      .HISTORICO_DESFECHOS;

  let aba =
    ss.getSheetByName(nomeAba);

  const cabecalhos = [
    'ID do Registro',
    'ID Paciente',
    'Prontuário',
    'Paciente',
    'ID Ciclo',
    'Ciclo Nº',
    'Desfecho',
    'Data da Desistência',
    'Sessões Prescritas',
    'Sessões Realizadas',
    'Sessões Restantes',
    'Sessões Futuras Canceladas',
    'Motivo Informado',
    'Data e Hora do Registro'
  ];

  if (!aba) {
    aba = ss.insertSheet(nomeAba);

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

    aba.setRowHeight(
      1,
      42
    );
  }

  const idRegistro =
    criarIdRegistroDesistencia_(
      paciente.id,
      ciclo.id
    );

  if (
    existeIdNaPrimeiraColunaDesistencia_(
      aba,
      idRegistro
    )
  ) {
    return;
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
      cabecalhos.length
    )
    .setValues([[
      idRegistro,
      paciente.id,
      paciente.prontuario,
      paciente.nome,
      ciclo.id,
      ciclo.numero,
      CONFIG_DESISTENCIA_TRATAMENTO
        .DESFECHO,
      entrada.dataDesistencia,
      paciente.sessoesPrescritas,
      paciente.sessoesRealizadas,
      paciente.sessoesRestantes,
      quantidadeCancelada,
      entrada.motivo,
      new Date()
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
      14
    )
    .setNumberFormat(
      'dd/MM/yyyy HH:mm'
    );

  aba
    .getRange(
      novaLinha,
      9,
      1,
      4
    )
    .setNumberFormat('0');

  aba.autoResizeColumns(
    1,
    cabecalhos.length
  );
}


/**
 * Cria a pendência do fisioterapeuta.
 */
function criarPendenciaEncerramentoDesistencia_(
  ss,
  paciente,
  ciclo,
  dataDesistencia
) {
  const abaPendencias =
    obterAbaDesistencia_(
      ss,
      CONFIG_DESISTENCIA_TRATAMENTO
        .ABAS
        .PENDENCIAS
    );

  const abaHistorico =
    obterAbaDesistencia_(
      ss,
      CONFIG_DESISTENCIA_TRATAMENTO
        .ABAS
        .HISTORICO_PENDENCIAS
    );

  const idPendencia =
    criarIdPendenciaDesistencia_(
      paciente.id,
      ciclo.id
    );

  if (
    existeIdNaPrimeiraColunaDesistencia_(
      abaPendencias,
      idPendencia
    ) ||
    existeIdNaPrimeiraColunaDesistencia_(
      abaHistorico,
      idPendencia
    )
  ) {
    return;
  }

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
      CONFIG_DESISTENCIA_TRATAMENTO
        .PENDENCIAS
        .QUANTIDADE_COLUNAS
    )
    .setValues([[
      idPendencia,
      'Alta',
      'Registrar encerramento por desistência',
      paciente.nome,
      paciente.prontuario,
      paciente.telefone,
      paciente.fisioterapeuta ||
        'Fisioterapeuta',
      dataDesistencia,
      dataDesistencia,
      'Pendente',
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
 * Inclui o novo status na lista oficial.
 */
function garantirStatusSessaoDesistencia_(ss) {
  const abaStatus =
    obterAbaDesistencia_(
      ss,
      CONFIG_DESISTENCIA_TRATAMENTO
        .ABAS
        .STATUS_SESSAO
    );

  const ultimaLinha =
    abaStatus.getLastRow();

  const valores = ultimaLinha >= 2
    ? abaStatus
        .getRange(
          2,
          1,
          ultimaLinha - 1,
          1
        )
        .getDisplayValues()
    : [];

  const statusProcurado =
    normalizarTextoDesistencia_(
      CONFIG_DESISTENCIA_TRATAMENTO
        .STATUS_AGENDAMENTO_DESISTENCIA
    );

  const existe = valores.some(
    function(linha) {
      return (
        normalizarTextoDesistencia_(
          linha[0]
        ) === statusProcurado
      );
    }
  );

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
        CONFIG_DESISTENCIA_TRATAMENTO
          .STATUS_AGENDAMENTO_DESISTENCIA
      );
  }
}


/**
 * Confere se um ID já existe na primeira coluna.
 */
function existeIdNaPrimeiraColunaDesistencia_(
  aba,
  id
) {
  const ultimaLinha =
    aba.getLastRow();

  if (ultimaLinha < 2) {
    return false;
  }

  const ids = aba
    .getRange(
      2,
      1,
      ultimaLinha - 1,
      1
    )
    .getDisplayValues();

  const procurado =
    normalizarTextoDesistencia_(
      id
    );

  return ids.some(function(linha) {
    return (
      normalizarTextoDesistencia_(
        linha[0]
      ) === procurado
    );
  });
}


/**
 * Cria o ID do histórico.
 */
function criarIdRegistroDesistencia_(
  idPaciente,
  idCiclo
) {
  return (
    'DESISTENCIA-' +
    limparIdDesistencia_(
      idPaciente
    ) +
    '-' +
    limparIdDesistencia_(
      idCiclo
    )
  );
}


/**
 * Cria o ID fixo da pendência.
 */
function criarIdPendenciaDesistencia_(
  idPaciente,
  idCiclo
) {
  return (
    'PEND-ENCERRAMENTO-DESISTENCIA-' +
    limparIdDesistencia_(
      idCiclo
    ) +
    '-' +
    limparIdDesistencia_(
      idPaciente
    )
  );
}


/**
 * Obtém uma aba obrigatória.
 */
function obterAbaDesistencia_(
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


/**
 * Converte a data enviada pelo formulário.
 */
function converterDataDesistencia_(texto) {
  const resultado = String(
    texto || ''
  ).trim().match(
    /^(\d{4})-(\d{2})-(\d{2})$/
  );

  if (!resultado) {
    return null;
  }

  const ano =
    Number(resultado[1]);

  const mes =
    Number(resultado[2]) - 1;

  const dia =
    Number(resultado[3]);

  const data = new Date(
    ano,
    mes,
    dia
  );

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
 * Formata a data para exibição.
 */
function formatarDataDesistencia_(data) {
  if (!dataValidaDesistencia_(data)) {
    return '';
  }

  return Utilities.formatDate(
    new Date(data),
    Session.getScriptTimeZone(),
    'dd/MM/yyyy'
  );
}


/**
 * Remove horas da data.
 */
function removerHorarioDesistencia_(data) {
  const copia = new Date(data);

  copia.setHours(0, 0, 0, 0);

  return copia;
}


/**
 * Confere se o valor é uma data.
 */
function dataValidaDesistencia_(data) {
  return (
    data instanceof Date &&
    !isNaN(data.getTime())
  );
}


/**
 * Padroniza identificadores.
 */
function limparIdDesistencia_(valor) {
  return normalizarTextoDesistencia_(
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


/**
 * Padroniza textos para comparação.
 */
function normalizarTextoDesistencia_(
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
