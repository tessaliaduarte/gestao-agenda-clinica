const CONFIG_CONFIRMAR_AGENDAMENTO = {
  ABAS: {
    VAGAS: 'Vagas para Sessões',
    CADASTRO: 'Cadastro de Pacientes',
    AGENDAMENTOS: 'Agendamentos',
    HORARIOS: 'Horários',
    FISIOTERAPEUTAS: 'Fisioterapeutas',
    FERIADOS: 'Calendário da Prefeitura',
    BLOQUEIOS: 'Bloqueios'
  },

  COLUNAS_VAGAS: {
    OPCAO: 1,
    ID_PACIENTE: 2,
    PRONTUARIO: 3,
    PACIENTE: 4,
    DIAS: 5,
    HORARIO: 6,
    FISIOTERAPEUTA: 7,
    TIPO_GRUPO: 8,
    OCUPACAO: 9,
    CAPACIDADE: 10,
    VAGAS: 11,
    DATA_INICIO: 12,
    DATA_TERMINO: 13,
    SITUACAO: 14
  },

  COLUNAS_CADASTRO: {
    ID: 1,
    PRONTUARIO: 2,
    NOME: 3,
    HORARIO: 6,
    TIPO_ATENDIMENTO: 7,
    LIMITE_GRUPO: 8,
    SEGUNDA: 9,
    TERCA: 10,
    QUARTA: 11,
    QUINTA: 12,
    SEXTA: 13,
    SESSOES_PRESCRITAS: 14,
    SESSOES_REALIZADAS: 15,
    SESSOES_RESTANTES: 16,
    DATA_INICIO: 19,
    DATA_TERMINO: 20,
    STATUS: 21,
    FISIOTERAPEUTA: 22,
    DESFECHO: 24
  },

  QUANTIDADE_COLUNAS_AGENDAMENTOS: 22,

  STATUS_ATIVOS: [
    'agendado'
  ]
};


/**
 * Confirma a opção selecionada na aba
 * "Vagas para Sessões".
 */
function confirmarAgendamentoSelecionado() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const lock = LockService.getDocumentLock();
  let bloqueioObtido = false;

  try {
    lock.waitLock(30000);
    bloqueioObtido = true;

    const abas =
      obterAbasConfirmarAgendamento_(ss);

    const selecao =
      obterOpcaoSelecionadaConfirmacao_(
        ss,
        abas.vagas
      );

    const paciente =
      localizarPacienteConfirmacao_(
        abas.cadastro,
        selecao.idPaciente
      );

    validarPacienteConfirmacao_(
      paciente
    );

    validarTipoGrupoSelecaoConfirmacao_(
      selecao,
      paciente
    );

    verificarCicloAtivoPacienteConfirmacao_(
      abas.agendamentos,
      paciente.id
    );

    const diasSemana =
      converterDiasTextoConfirmacao_(
        selecao.dias
      );

    if (diasSemana.length === 0) {
      throw new Error(
        'Não foi possível identificar os dias da opção selecionada.'
      );
    }

    const resumo =
      montarResumoConfirmacao_(
        selecao,
        paciente
      );

    const resposta = ui.alert(
      'Confirmar agendamento',
      resumo,
      ui.ButtonSet.YES_NO
    );

    if (resposta !== ui.Button.YES) {
      return;
    }

    /*
     * Recria o contexto da busca para verificar se
     * as vagas continuam disponíveis.
     */
    const abasVagas = {
      cadastro: abas.cadastro,
      resultados: abas.vagas,
      agendamentos: abas.agendamentos,
      horarios: abas.horarios,
      fisioterapeutas: abas.fisioterapeutas,
      feriados: abas.feriados,
      bloqueios: abas.bloqueios
    };

    /*
     * Usa a mesma quantidade calculada pelo módulo
     * de consulta de vagas.
     */
    validarPacienteParaBuscaSessoes_(
      paciente
    );

    const contexto =
      construirContextoVagasSessoes_(
        abasVagas,
        paciente,
        selecao.dataInicio
      );

    const horarioOpcao = {
      valor: selecao.horario,

      chave:
        chaveHorarioVagasSessoes_(
          selecao.horario
        ),

      exibicao:
        formatarHorarioVagasSessoes_(
          selecao.horario
        ),

      turno:
        localizarTurnoHorarioConfirmacao_(
          contexto.horarios,
          selecao.horario
        )
    };

    const sequencia =
      simularSequenciaVagasSessoes_(
        contexto,
        diasSemana,
        horarioOpcao,
        selecao.fisioterapeuta
      );

    if (!sequencia.valida) {
      throw new Error(
        'A opção selecionada não está mais disponível. ' +
        'Faça uma nova consulta de vagas.'
      );
    }

    const cicloNumero =
      obterProximoNumeroCicloConfirmacao_(
        abas.agendamentos,
        paciente.id
      );

    const idCiclo =
      gerarIdCicloConfirmacao_(
        paciente.id,
        cicloNumero
      );

    const primeiroNumeroAgendamento =
      obterProximoNumeroAgendamentoConfirmacao_(
        abas.agendamentos
      );

    const agora = new Date();

    const linhasAgendamentos =
      montarLinhasAgendamentosConfirmacao_(
        paciente,
        selecao,
        sequencia.datas,
        idCiclo,
        cicloNumero,
        primeiroNumeroAgendamento,
        agora
      );

    gravarAgendamentosConfirmacao_(
      abas.agendamentos,
      linhasAgendamentos
    );

    atualizarCadastroPacienteConfirmacao_(
      abas.cadastro,
      paciente,
      selecao,
      diasSemana,
      sequencia.datas
    );

    marcarOpcaoConfirmada_(
      abas.vagas,
      selecao.linha
    );

    SpreadsheetApp.flush();

    if (
      typeof atualizarPendenciasAutomaticas ===
      'function'
    ) {
      atualizarPendenciasAutomaticas();
    }

    ui.alert(
      'Agendamento confirmado',
      'O tratamento foi agendado com sucesso.\n\n' +
        'Paciente: ' +
        paciente.nome +
        '\nCiclo: ' +
        cicloNumero +
        '\nSessões geradas: ' +
        sequencia.datas.length +
        '\nInício: ' +
        formatarDataConfirmacao_(
          sequencia.datas[0]
        ) +
        '\nTérmino previsto: ' +
        formatarDataConfirmacao_(
          sequencia.datas[
            sequencia.datas.length - 1
          ]
        ),
      ui.ButtonSet.OK
    );

    ss.setActiveSheet(
      abas.agendamentos
    );
  } catch (erro) {
    ui.alert(
      'Erro ao confirmar agendamento',
      erro && erro.message
        ? erro.message
        : String(erro),
      ui.ButtonSet.OK
    );
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
 * Localiza todas as abas necessárias.
 */function obterAbasConfirmarAgendamento_(ss) {
  const nomes =
    CONFIG_CONFIRMAR_AGENDAMENTO.ABAS;

  const abas = {
    vagas:
      ss.getSheetByName(
        nomes.VAGAS
      ),

    cadastro:
      ss.getSheetByName(
        nomes.CADASTRO
      ),

    agendamentos:
      ss.getSheetByName(
        nomes.AGENDAMENTOS
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
      )
  };

  Object.keys(abas).forEach(function(chave) {
    if (!abas[chave]) {
      throw new Error(
        'A aba necessária "' +
        nomes[
          {
            vagas: 'VAGAS',
            cadastro: 'CADASTRO',
            agendamentos: 'AGENDAMENTOS',
            horarios: 'HORARIOS',
            fisioterapeutas: 'FISIOTERAPEUTAS',
            feriados: 'FERIADOS',
            bloqueios: 'BLOQUEIOS'
          }[chave]
        ] +
        '" não foi encontrada.'
      );
    }
  });

  return abas;
}


/**
 * Obtém os dados da linha selecionada.
 */
function obterOpcaoSelecionadaConfirmacao_(
  ss,
  abaVagas
) {
  const abaAtiva =
    ss.getActiveSheet();

  if (
    abaAtiva.getName() !==
    CONFIG_CONFIRMAR_AGENDAMENTO.ABAS.VAGAS
  ) {
    throw new Error(
      'Abra a aba "Vagas para Sessões" e selecione uma opção.'
    );
  }

  const intervaloAtivo =
    abaAtiva.getActiveRange();

  if (!intervaloAtivo) {
    throw new Error(
      'Selecione uma célula da opção que deseja confirmar.'
    );
  }

  const linha =
    intervaloAtivo.getRow();

  if (linha < 2) {
    throw new Error(
      'Selecione uma linha de opção abaixo do cabeçalho.'
    );
  }

  const valores = abaVagas
    .getRange(
      linha,
      1,
      1,
      14
    )
    .getValues()[0];

  const idPaciente =
    String(
      valores[
        CONFIG_CONFIRMAR_AGENDAMENTO
          .COLUNAS_VAGAS
          .ID_PACIENTE - 1
      ] || ''
    ).trim();

  const paciente =
    String(
      valores[
        CONFIG_CONFIRMAR_AGENDAMENTO
          .COLUNAS_VAGAS
          .PACIENTE - 1
      ] || ''
    ).trim();

  const dias =
    String(
      valores[
        CONFIG_CONFIRMAR_AGENDAMENTO
          .COLUNAS_VAGAS
          .DIAS - 1
      ] || ''
    ).trim();

  const horario =
    valores[
      CONFIG_CONFIRMAR_AGENDAMENTO
        .COLUNAS_VAGAS
        .HORARIO - 1
    ];

  const fisioterapeuta =
    String(
      valores[
        CONFIG_CONFIRMAR_AGENDAMENTO
          .COLUNAS_VAGAS
          .FISIOTERAPEUTA - 1
      ] || ''
    ).trim();

  const tipoGrupo =
    String(
      valores[
        CONFIG_CONFIRMAR_AGENDAMENTO
          .COLUNAS_VAGAS
          .TIPO_GRUPO - 1
      ] || ''
    ).trim();

  const dataInicio =
    valores[
      CONFIG_CONFIRMAR_AGENDAMENTO
        .COLUNAS_VAGAS
        .DATA_INICIO - 1
    ];

  const dataTermino =
    valores[
      CONFIG_CONFIRMAR_AGENDAMENTO
        .COLUNAS_VAGAS
        .DATA_TERMINO - 1
    ];

  const situacao =
    String(
      valores[
        CONFIG_CONFIRMAR_AGENDAMENTO
          .COLUNAS_VAGAS
          .SITUACAO - 1
      ] || ''
    ).trim();

  if (
    !idPaciente ||
    !paciente ||
    !dias ||
    !horario ||
    !fisioterapeuta ||
    !tipoGrupo ||
    !(dataInicio instanceof Date)
  ) {
    throw new Error(
      'A linha selecionada não contém uma opção válida.'
    );
  }

  if (
    normalizarTextoConfirmacao_(
      situacao
    ) ===
    'agendamento confirmado'
  ) {
    throw new Error(
      'Esta opção já foi confirmada.'
    );
  }

  return {
    linha: linha,

    opcao:
      valores[
        CONFIG_CONFIRMAR_AGENDAMENTO
          .COLUNAS_VAGAS
          .OPCAO - 1
      ],

    idPaciente: idPaciente,

    prontuario:
      String(
        valores[
          CONFIG_CONFIRMAR_AGENDAMENTO
            .COLUNAS_VAGAS
            .PRONTUARIO - 1
        ] || ''
      ).trim(),

    paciente: paciente,
    dias: dias,
    horario: horario,
    fisioterapeuta: fisioterapeuta,
    tipoGrupo: tipoGrupo,
    dataInicio: dataInicio,
    dataTermino: dataTermino,
    situacao: situacao
  };
}


/**
 * Localiza o paciente pelo ID.
 */
function localizarPacienteConfirmacao_(
  abaCadastro,
  idPaciente
) {
  const ultimaLinha =
    abaCadastro.getLastRow();

  if (ultimaLinha < 2) {
    throw new Error(
      'Não existem pacientes cadastrados.'
    );
  }

  const dados = abaCadastro
    .getRange(
      2,
      1,
      ultimaLinha - 1,
      24
    )
    .getValues();

  const idProcurado =
    normalizarTextoConfirmacao_(
      idPaciente
    );

  for (
    let indice = 0;
    indice < dados.length;
    indice++
  ) {
    const linha = dados[indice];

    const idAtual =
      normalizarTextoConfirmacao_(
        linha[0]
      );

    if (idAtual === idProcurado) {
      return montarPacienteVagasSessoes_(
        linha,
        indice + 2
      );
    }
  }

  throw new Error(
    'O paciente da opção selecionada não foi encontrado no cadastro.'
  );
}


/**
 * Valida os dados essenciais do paciente.
 */function validarPacienteConfirmacao_(
  paciente
) {
  if (!paciente) {
    throw new Error(
      'Paciente não encontrado.'
    );
  }

  if (!paciente.id) {
    throw new Error(
      'O paciente não possui ID.'
    );
  }

  if (!paciente.nome) {
    throw new Error(
      'O paciente não possui nome.'
    );
  }

  if (
    paciente.sessoesPrescritas < 1
  ) {
    throw new Error(
      'O paciente não possui sessões prescritas.'
    );
  }

  if (
    paciente.sessoesRestantes <= 0 &&
    paciente.sessoesRealizadas >=
      paciente.sessoesPrescritas
  ) {
    throw new Error(
      'O paciente não possui sessões restantes.'
    );
  }

  const status =
    normalizarTextoConfirmacao_(
      paciente.status
    );

  if (
    status === 'ciclo concluido' ||
    status === 'inativo'
  ) {
    throw new Error(
      'O paciente está com status "' +
      paciente.status +
      '" e não pode ser agendado.'
    );
  }
}


/**
 * Impede a confirmação de uma opção antiga quando
 * o tipo de atendimento do paciente foi alterado
 * depois da consulta de vagas.
 */
function validarTipoGrupoSelecaoConfirmacao_(
  selecao,
  paciente
) {
  const tipoSelecao =
    normalizarTextoConfirmacao_(
      selecao.tipoGrupo
    );

  const tipoPaciente =
    normalizarTextoConfirmacao_(
      paciente.tipoGrupo
    );

  if (
    !tipoSelecao ||
    !tipoPaciente ||
    tipoSelecao !== tipoPaciente
  ) {
    throw new Error(
      'O tipo de atendimento do paciente foi alterado depois da consulta. Faça uma nova consulta de vagas antes de confirmar.'
    );
  }
}


/**
 * Impede a criação de dois ciclos ativos
 * para o mesmo paciente.
 */
function verificarCicloAtivoPacienteConfirmacao_(
  abaAgendamentos,
  idPaciente
) {
  const ultimaLinha =
    abaAgendamentos.getLastRow();

  if (ultimaLinha < 2) {
    return;
  }

  const dados = abaAgendamentos
    .getRange(
      2,
      1,
      ultimaLinha - 1,
      CONFIG_CONFIRMAR_AGENDAMENTO
        .QUANTIDADE_COLUNAS_AGENDAMENTOS
    )
    .getValues();

  const idProcurado =
    normalizarTextoConfirmacao_(
      idPaciente
    );

  const possuiCicloAtivo =
    dados.some(function(linha) {
      const idAtual =
        normalizarTextoConfirmacao_(
          linha[1]
        );

      const evento =
        normalizarTextoConfirmacao_(
          linha[11]
        );

      const status =
        normalizarTextoConfirmacao_(
          linha[15]
        );

      return (
        idAtual === idProcurado &&
        evento === 'sessao' &&
        CONFIG_CONFIRMAR_AGENDAMENTO
          .STATUS_ATIVOS
          .indexOf(status) !== -1
      );
    });

  if (possuiCicloAtivo) {
    throw new Error(
      'Este paciente já possui um ciclo de tratamento ativo na aba "Agendamentos".'
    );
  }
}


/**
 * Converte "Seg / Qua" em [1, 3].
 */
function converterDiasTextoConfirmacao_(
  texto
) {
  const mapa = {
    seg: 1,
    segunda: 1,
    'segunda-feira': 1,

    ter: 2,
    terca: 2,
    'terca-feira': 2,

    qua: 3,
    quarta: 3,
    'quarta-feira': 3,

    qui: 4,
    quinta: 4,
    'quinta-feira': 4,

    sex: 5,
    sexta: 5,
    'sexta-feira': 5
  };

  const partes = String(
    texto || ''
  )
    .split(/[\/,;|]+/)
    .map(function(parte) {
      return normalizarTextoConfirmacao_(
        parte
      );
    })
    .filter(String);

  const dias = [];

  partes.forEach(function(parte) {
    const numero =
      mapa[parte];

    if (
      numero &&
      dias.indexOf(numero) === -1
    ) {
      dias.push(numero);
    }
  });

  return dias.sort();
}


/**
 * Encontra o turno correspondente ao horário.
 */
function localizarTurnoHorarioConfirmacao_(
  horarios,
  horarioSelecionado
) {
  const chaveSelecionada =
    chaveHorarioVagasSessoes_(
      horarioSelecionado
    );

  const encontrado =
    horarios.find(function(horario) {
      return (
        horario.chave ===
        chaveSelecionada
      );
    });

  if (!encontrado) {
    throw new Error(
      'O horário selecionado não foi encontrado na aba "Horários".'
    );
  }

  return encontrado.turno;
}


/**
 * Monta o texto apresentado antes da confirmação.
 */function montarResumoConfirmacao_(
  selecao,
  paciente
) {
  const quantidade =
    paciente.sessoesRestantes > 0
      ? paciente.sessoesRestantes
      : paciente.sessoesPrescritas;

  return (
    'Confira os dados antes de confirmar:\n\n' +
    'Paciente: ' +
    paciente.nome +
    '\nProntuário: ' +
    paciente.prontuario +
    '\nDias: ' +
    selecao.dias +
    '\nHorário: ' +
    formatarHorarioConfirmacao_(
      selecao.horario
    ) +
    '\nFisioterapeuta: ' +
    selecao.fisioterapeuta +
    '\nTipo de grupo: ' +
    selecao.tipoGrupo +
    '\nSessões: ' +
    quantidade +
    '\nInício: ' +
    formatarDataConfirmacao_(
      selecao.dataInicio
    ) +
    '\nTérmino previsto: ' +
    formatarDataConfirmacao_(
      selecao.dataTermino
    ) +
    '\n\nDeseja confirmar este agendamento?'
  );
}


/**
 * Obtém o próximo número do ciclo do paciente.
 */
function obterProximoNumeroCicloConfirmacao_(
  abaAgendamentos,
  idPaciente
) {
  const ultimaLinha =
    abaAgendamentos.getLastRow();

  if (ultimaLinha < 2) {
    return 1;
  }

  const dados = abaAgendamentos
    .getRange(
      2,
      2,
      ultimaLinha - 1,
      5
    )
    .getValues();

  const idProcurado =
    normalizarTextoConfirmacao_(
      idPaciente
    );

  let maiorNumero = 0;

  dados.forEach(function(linha) {
    const idAtual =
      normalizarTextoConfirmacao_(
        linha[0]
      );

    const numeroCiclo =
      Number(linha[4]) || 0;

    if (
      idAtual === idProcurado &&
      numeroCiclo > maiorNumero
    ) {
      maiorNumero = numeroCiclo;
    }
  });

  return maiorNumero + 1;
}


/**
 * Gera o ID do ciclo.
 */
function gerarIdCicloConfirmacao_(
  idPaciente,
  numeroCiclo
) {
  return (
    'CIC-' +
    String(idPaciente)
      .replace(/^PAC-/i, '') +
    '-' +
    String(numeroCiclo)
      .padStart(2, '0')
  );
}


/**
 * Encontra o próximo número usado no ID
 * de agendamento.
 */
function obterProximoNumeroAgendamentoConfirmacao_(
  abaAgendamentos
) {
  const ultimaLinha =
    abaAgendamentos.getLastRow();

  if (ultimaLinha < 2) {
    return 1;
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

  ids.forEach(function(linha) {
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
      Number(resultado[1]);

    if (
      Number.isFinite(numero) &&
      numero > maiorNumero
    ) {
      maiorNumero = numero;
    }
  });

  return maiorNumero + 1;
}


/**
 * Cria todas as linhas do ciclo.
 */
function montarLinhasAgendamentosConfirmacao_(
  paciente,
  selecao,
  datas,
  idCiclo,
  cicloNumero,
  primeiroNumeroAgendamento,
  agora
) {
  const total =
    datas.length;

  return datas.map(
    function(data, indice) {
      const idAgendamento =
        'AG-' +
        String(
          primeiroNumeroAgendamento +
          indice
        ).padStart(6, '0');

      return [
        idAgendamento,
        paciente.id,
        paciente.prontuario,
        paciente.nome,
        idCiclo,
        cicloNumero,
        new Date(data),
        obterNomeDiaConfirmacao_(data),
        selecao.horario,
        selecao.fisioterapeuta,
        selecao.tipoGrupo,
        'Sessão',
        indice + 1,
        total,
        paciente.capacidade,
        'Agendado',
        '',
        'Não',
        'Sim',
        agora,
        agora,
        'Não'
      ];
    }
  );
}


/**
 * Grava os agendamentos na próxima linha livre.
 */
function gravarAgendamentosConfirmacao_(
  abaAgendamentos,
  linhas
) {
  if (
    !linhas ||
    linhas.length === 0
  ) {
    throw new Error(
      'Nenhuma sessão foi gerada.'
    );
  }

  const primeiraLinha =
    Math.max(
      abaAgendamentos.getLastRow() + 1,
      2
    );

  abaAgendamentos
    .getRange(
      primeiraLinha,
      1,
      linhas.length,
      CONFIG_CONFIRMAR_AGENDAMENTO
        .QUANTIDADE_COLUNAS_AGENDAMENTOS
    )
    .setValues(linhas);

  abaAgendamentos
    .getRange(
      primeiraLinha,
      7,
      linhas.length,
      1
    )
    .setNumberFormat(
      'dd/MM/yyyy'
    );

  abaAgendamentos
    .getRange(
      primeiraLinha,
      9,
      linhas.length,
      1
    )
    .setNumberFormat(
      'HH:mm'
    );

  abaAgendamentos
    .getRange(
      primeiraLinha,
      13,
      linhas.length,
      3
    )
    .setNumberFormat(
      '0'
    );

  abaAgendamentos
    .getRange(
      primeiraLinha,
      20,
      linhas.length,
      2
    )
    .setNumberFormat(
      'dd/MM/yyyy HH:mm'
    );

  abaAgendamentos
    .getRange(
      primeiraLinha,
      1,
      linhas.length,
      CONFIG_CONFIRMAR_AGENDAMENTO
        .QUANTIDADE_COLUNAS_AGENDAMENTOS
    )
    .setVerticalAlignment(
      'middle'
    );
}


/**
 * Atualiza o cadastro após a criação do ciclo.
 */function atualizarCadastroPacienteConfirmacao_(
  abaCadastro,
  paciente,
  selecao,
  diasSemana,
  datas
) {
  const linha =
    paciente.linha;

  const primeiraData =
    datas[0];

  const ultimaData =
    datas[datas.length - 1];

  const diasMarcados = {
    1: diasSemana.indexOf(1) !== -1,
    2: diasSemana.indexOf(2) !== -1,
    3: diasSemana.indexOf(3) !== -1,
    4: diasSemana.indexOf(4) !== -1,
    5: diasSemana.indexOf(5) !== -1
  };

  abaCadastro
    .getRange(
      linha,
      CONFIG_CONFIRMAR_AGENDAMENTO
        .COLUNAS_CADASTRO
        .HORARIO
    )
    .setValue(
      selecao.horario
    )
    .setNumberFormat(
      'HH:mm'
    );

  abaCadastro
    .getRange(
      linha,
      CONFIG_CONFIRMAR_AGENDAMENTO
        .COLUNAS_CADASTRO
        .TIPO_ATENDIMENTO
    )
    .setValue(
      selecao.tipoGrupo
    );

  abaCadastro
    .getRange(
      linha,
      CONFIG_CONFIRMAR_AGENDAMENTO
        .COLUNAS_CADASTRO
        .LIMITE_GRUPO
    )
    .setValue(
      paciente.capacidade
    );

  abaCadastro
    .getRange(
      linha,
      CONFIG_CONFIRMAR_AGENDAMENTO
        .COLUNAS_CADASTRO
        .SEGUNDA,
      1,
      5
    )
    .setValues([[
      diasMarcados[1],
      diasMarcados[2],
      diasMarcados[3],
      diasMarcados[4],
      diasMarcados[5]
    ]]);

  abaCadastro
    .getRange(
      linha,
      CONFIG_CONFIRMAR_AGENDAMENTO
        .COLUNAS_CADASTRO
        .DATA_INICIO
    )
    .setValue(
      primeiraData
    )
    .setNumberFormat(
      'dd/MM/yyyy'
    );

  abaCadastro
    .getRange(
      linha,
      CONFIG_CONFIRMAR_AGENDAMENTO
        .COLUNAS_CADASTRO
        .DATA_TERMINO
    )
    .setValue(
      ultimaData
    )
    .setNumberFormat(
      'dd/MM/yyyy'
    );

  abaCadastro
    .getRange(
      linha,
      CONFIG_CONFIRMAR_AGENDAMENTO
        .COLUNAS_CADASTRO
        .STATUS
    )
    .setValue(
      'Em tratamento'
    );

  abaCadastro
    .getRange(
      linha,
      CONFIG_CONFIRMAR_AGENDAMENTO
        .COLUNAS_CADASTRO
        .FISIOTERAPEUTA
    )
    .setValue(
      selecao.fisioterapeuta
    );

  /*
   * O desfecho "Renovação" pertence ao ciclo anterior.
   * Após o novo ciclo ser criado, a coluna volta a ficar
   * disponível para o desfecho futuro.
   */
  abaCadastro
    .getRange(
      linha,
      CONFIG_CONFIRMAR_AGENDAMENTO
        .COLUNAS_CADASTRO
        .DESFECHO
    )
    .clearContent();
}


/**
 * Marca a linha da vaga como confirmada.
 */
function marcarOpcaoConfirmada_(
  abaVagas,
  linha
) {
  abaVagas
    .getRange(
      linha,
      CONFIG_CONFIRMAR_AGENDAMENTO
        .COLUNAS_VAGAS
        .SITUACAO
    )
    .setValue(
      'Agendamento confirmado'
    )
    .setBackground(
      '#b6d7a8'
    )
    .setFontWeight(
      'bold'
    );

  abaVagas
    .getRange(
      linha,
      1,
      1,
      14
    )
    .setBackground(
      '#d9ead3'
    );
}


/**
 * Nome do dia da semana.
 */
function obterNomeDiaConfirmacao_(
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


/**
 * Formata uma data para exibição.
 */
function formatarDataConfirmacao_(
  data
) {
  if (!(data instanceof Date)) {
    return '';
  }

  return Utilities.formatDate(
    data,
    Session.getScriptTimeZone(),
    'dd/MM/yyyy'
  );
}


/**
 * Formata um horário para exibição.
 */
function formatarHorarioConfirmacao_(
  horario
) {
  if (horario instanceof Date) {
    return Utilities.formatDate(
      horario,
      Session.getScriptTimeZone(),
      'HH:mm'
    );
  }

  return String(
    horario || ''
  ).trim();
}


/**
 * Normaliza textos para comparações.
 */
function normalizarTextoConfirmacao_(
  valor
) {
  return String(
    valor || ''
  )
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(
      /[\u0300-\u036f]/g,
      ''
    );
}
