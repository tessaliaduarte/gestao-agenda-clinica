const CONFIG_RENOVACAO_TRATAMENTO = {
  ABAS: {
    CADASTRO: 'Cadastro de Pacientes',
    AGENDAMENTOS: 'Agendamentos',
    HORARIOS: 'Horários',
    TIPOS_GRUPO: 'Tipos de Grupo',
    FISIOTERAPEUTAS: 'Fisioterapeutas'
  },

  COLUNAS_CADASTRO: {
    ID_PACIENTE: 1,
    PRONTUARIO: 2,
    NOME: 3,
    CPF: 4,
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

  COLUNAS_AGENDAMENTOS: {
    ID_PACIENTE: 2,
    CICLO_NUMERO: 6
  },

  STATUS_EXIGIDO: 'Ciclo concluído',
  DESFECHO_EXIGIDO: 'Renovação',
  STATUS_APOS_PLANEJAMENTO:
    'Avaliado – aguardando agendamento',
  MAXIMO_SESSOES: 20
};


/**
 * Abre a página de pesquisa para localizar
 * o paciente que terá o ciclo renovado.
 */
function abrirFormularioPlanejamentoRenovacao() {
  const html = HtmlService
    .createHtmlOutputFromFile(
      'FormularioRenovacao'
    )
    .setWidth(760)
    .setHeight(720);

  SpreadsheetApp.getUi().showModalDialog(
    html,
    'SIGAF — Planejar renovação'
  );
}


/**
 * Pesquisa pacientes por ID, prontuário,
 * CPF, nome completo ou parte do nome.
 */
function buscarPacientesParaRenovacao(
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
    obterAbaObrigatoriaRenovacao_(
      ss,
      CONFIG_RENOVACAO_TRATAMENTO
        .ABAS
        .CADASTRO
    );

  const ultimaLinha =
    abaCadastro.getLastRow();

  if (ultimaLinha < 2) {
    return [];
  }

  const c =
    CONFIG_RENOVACAO_TRATAMENTO
      .COLUNAS_CADASTRO;

  const dados = abaCadastro
    .getRange(
      2,
      1,
      ultimaLinha - 1,
      c.DESFECHO
    )
    .getValues();

  const termoNormalizado =
    normalizarTextoRenovacao_(
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
      normalizarTextoRenovacao_(id) ===
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
      normalizarTextoRenovacao_(nome)
        .indexOf(termoNormalizado) !== -1;

    if (!corresponde) {
      return;
    }

    resultados.push({
      id: id,
      prontuario: prontuario,
      nome: nome,
      cpf: cpf,
      status: String(
        linha[c.STATUS - 1] || ''
      ).trim(),
      desfecho: String(
        linha[c.DESFECHO - 1] || ''
      ).trim()
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
 * Carrega e valida o paciente escolhido
 * antes de exibir o planejamento.
 */
function obterPacienteParaPlanejamentoRenovacao(
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

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const abaCadastro =
    obterAbaObrigatoriaRenovacao_(
      ss,
      CONFIG_RENOVACAO_TRATAMENTO
        .ABAS
        .CADASTRO
    );

  const linhaPaciente =
    localizarLinhaPacienteRenovacao_(
      abaCadastro,
      idProcurado
    );

  if (!linhaPaciente) {
    throw new Error(
      'Paciente não encontrado.'
    );
  }

  const paciente =
    obterPacienteCadastroRenovacao_(
      abaCadastro,
      linhaPaciente
    );

  validarPacienteParaRenovacao_(
    paciente
  );

  const abaAgendamentos =
    obterAbaObrigatoriaRenovacao_(
      ss,
      CONFIG_RENOVACAO_TRATAMENTO
        .ABAS
        .AGENDAMENTOS
    );

  const cicloAnterior =
    obterUltimoCicloPacienteRenovacao_(
      abaAgendamentos,
      paciente.id
    );

  if (cicloAnterior < 1) {
    throw new Error(
      'Não foi encontrado um ciclo anterior deste paciente na aba "Agendamentos".'
    );
  }

  const dadosIniciais = {
    idPaciente: paciente.id,
    prontuario: paciente.prontuario,
    nome: paciente.nome,
    cicloAnterior: cicloAnterior,
    novoCiclo: cicloAnterior + 1,
    sessoesAnteriores:
      paciente.sessoesPrescritas,
    horarioAtual:
      chaveHorarioRenovacao_(
        paciente.horario
      ),
    tipoAtual:
      paciente.tipoAtendimento,
    fisioterapeutaAtual:
      paciente.fisioterapeuta,
    diasAtuais:
      paciente.diasSemana,
    horarios:
      lerOpcoesColunaRenovacao_(
        obterAbaObrigatoriaRenovacao_(
          ss,
          CONFIG_RENOVACAO_TRATAMENTO
            .ABAS
            .HORARIOS
        ),
        1,
        true
      ),
    tiposAtendimento:
      lerOpcoesColunaRenovacao_(
        obterAbaObrigatoriaRenovacao_(
          ss,
          CONFIG_RENOVACAO_TRATAMENTO
            .ABAS
            .TIPOS_GRUPO
        ),
        1,
        false
      ),
    fisioterapeutas:
      lerOpcoesColunaRenovacao_(
        obterAbaObrigatoriaRenovacao_(
          ss,
          CONFIG_RENOVACAO_TRATAMENTO
            .ABAS
            .FISIOTERAPEUTAS
        ),
        2,
        false
      )
  };

  validarListasPlanejamentoRenovacao_(
    dadosIniciais
  );

  return dadosIniciais;
}


/**
 * Salva o planejamento do novo ciclo.
 * Não altera nem exclui registros antigos de Agendamentos.
 */function salvarPlanejamentoRenovacao(dados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getDocumentLock();
  let bloqueioObtido = false;

  try {
    lock.waitLock(30000);
    bloqueioObtido = true;

    const entrada =
      validarEntradaPlanejamentoRenovacao_(
        dados
      );

    const abaCadastro =
      obterAbaObrigatoriaRenovacao_(
        ss,
        CONFIG_RENOVACAO_TRATAMENTO
          .ABAS
          .CADASTRO
      );

    const linhaPaciente =
      localizarLinhaPacienteRenovacao_(
        abaCadastro,
        entrada.idPaciente
      );

    if (!linhaPaciente) {
      throw new Error(
        'O paciente não foi encontrado no Cadastro de Pacientes.'
      );
    }

    const paciente =
      obterPacienteCadastroRenovacao_(
        abaCadastro,
        linhaPaciente
      );

    validarPacienteParaRenovacao_(
      paciente
    );

    const abaAgendamentos =
      obterAbaObrigatoriaRenovacao_(
        ss,
        CONFIG_RENOVACAO_TRATAMENTO
          .ABAS
          .AGENDAMENTOS
      );

    const cicloAnterior =
      obterUltimoCicloPacienteRenovacao_(
        abaAgendamentos,
        paciente.id
      );

    const novoCicloEsperado =
      cicloAnterior + 1;

    if (
      cicloAnterior < 1 ||
      entrada.novoCiclo !==
        novoCicloEsperado
    ) {
      throw new Error(
        'O número do novo ciclo mudou. Feche o formulário e abra-o novamente.'
      );
    }

    const horario =
      validarOpcaoCadastradaRenovacao_(
        entrada.horario,
        lerOpcoesColunaRenovacao_(
          obterAbaObrigatoriaRenovacao_(
            ss,
            CONFIG_RENOVACAO_TRATAMENTO
              .ABAS
              .HORARIOS
          ),
          1,
          true
        ),
        'horário'
      );

    const tipoAtendimento =
      validarOpcaoCadastradaRenovacao_(
        entrada.tipoAtendimento,
        lerOpcoesColunaRenovacao_(
          obterAbaObrigatoriaRenovacao_(
            ss,
            CONFIG_RENOVACAO_TRATAMENTO
              .ABAS
              .TIPOS_GRUPO
          ),
          1,
          false
        ),
        'tipo de atendimento'
      );

    const fisioterapeuta =
      validarOpcaoCadastradaRenovacao_(
        entrada.fisioterapeuta,
        lerOpcoesColunaRenovacao_(
          obterAbaObrigatoriaRenovacao_(
            ss,
            CONFIG_RENOVACAO_TRATAMENTO
              .ABAS
              .FISIOTERAPEUTAS
          ),
          2,
          false
        ),
        'fisioterapeuta'
      );

    const capacidade =
      obterCapacidadeRenovacao_(
        tipoAtendimento
      );

    const c =
      CONFIG_RENOVACAO_TRATAMENTO
        .COLUNAS_CADASTRO;

    abaCadastro
      .getRange(
        linhaPaciente,
        c.HORARIO
      )
      .setValue(horario)
      .setNumberFormat('HH:mm');

    abaCadastro
      .getRange(
        linhaPaciente,
        c.TIPO_ATENDIMENTO
      )
      .setValue(tipoAtendimento);

    abaCadastro
      .getRange(
        linhaPaciente,
        c.LIMITE_GRUPO
      )
      .setValue(capacidade)
      .setNumberFormat('0');

    abaCadastro
      .getRange(
        linhaPaciente,
        c.SEGUNDA,
        1,
        5
      )
      .setValues([[
        entrada.diasSemana[1],
        entrada.diasSemana[2],
        entrada.diasSemana[3],
        entrada.diasSemana[4],
        entrada.diasSemana[5]
      ]]);

    abaCadastro
      .getRange(
        linhaPaciente,
        c.SESSOES_PRESCRITAS
      )
      .setValue(
        entrada.sessoesPrescritas
      )
      .setNumberFormat('0');

    abaCadastro
      .getRange(
        linhaPaciente,
        c.SESSOES_REALIZADAS
      )
      .setValue(0)
      .setNumberFormat('0');

    abaCadastro
      .getRange(
        linhaPaciente,
        c.SESSOES_RESTANTES
      )
      .setValue(
        entrada.sessoesPrescritas
      )
      .setNumberFormat('0');

    abaCadastro
      .getRange(
        linhaPaciente,
        c.DATA_INICIO,
        1,
        2
      )
      .clearContent()
      .setNumberFormat('dd/MM/yyyy');

    abaCadastro
      .getRange(
        linhaPaciente,
        c.STATUS
      )
      .setValue(
        CONFIG_RENOVACAO_TRATAMENTO
          .STATUS_APOS_PLANEJAMENTO
      );

    abaCadastro
      .getRange(
        linhaPaciente,
        c.FISIOTERAPEUTA
      )
      .setValue(fisioterapeuta);

    /*
     * O desfecho é mantido até a criação efetiva
     * do novo ciclo. ConfirmarAgendamento.gs o limpa
     * depois de gerar as novas sessões.
     */
    abaCadastro
      .getRange(
        linhaPaciente,
        c.DESFECHO
      )
      .setValue(
        CONFIG_RENOVACAO_TRATAMENTO
          .DESFECHO_EXIGIDO
      );

    SpreadsheetApp.flush();

    if (
      typeof atualizarPendenciasAutomaticas ===
      'function'
    ) {
      atualizarPendenciasAutomaticas();
    }

    return {
      sucesso: true,
      mensagem:
        'O planejamento do ciclo ' +
        novoCicloEsperado +
        ' foi salvo para ' +
        paciente.nome +
        '.\n\n' +
        'As sessões do ciclo anterior permanecem intactas em Agendamentos.\n\n' +
        'Próximas etapas:\n' +
        '1. Na aba Pendências, conclua "Definir planejamento da renovação" e informe a data da conclusão.\n' +
        '2. Consulte a disponibilidade de sessões.\n' +
        '3. Confirme a vaga para gerar o ciclo ' +
        novoCicloEsperado +
        '.'
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
        // O bloqueio será liberado automaticamente.
      }
    }
  }
}


function obterPacienteCadastroRenovacao_(
  aba,
  linha
) {
  const c =
    CONFIG_RENOVACAO_TRATAMENTO
      .COLUNAS_CADASTRO;

  const dados = aba
    .getRange(
      linha,
      1,
      1,
      c.DESFECHO
    )
    .getValues()[0];

  return {
    linha: linha,
    id: String(
      dados[c.ID_PACIENTE - 1] || ''
    ).trim(),
    prontuario: String(
      dados[c.PRONTUARIO - 1] || ''
    ).trim(),
    nome: String(
      dados[c.NOME - 1] || ''
    ).trim(),
    horario:
      dados[c.HORARIO - 1],
    tipoAtendimento: String(
      dados[c.TIPO_ATENDIMENTO - 1] || ''
    ).trim(),
    diasSemana: {
      1: valorAtivoRenovacao_(
        dados[c.SEGUNDA - 1]
      ),
      2: valorAtivoRenovacao_(
        dados[c.TERCA - 1]
      ),
      3: valorAtivoRenovacao_(
        dados[c.QUARTA - 1]
      ),
      4: valorAtivoRenovacao_(
        dados[c.QUINTA - 1]
      ),
      5: valorAtivoRenovacao_(
        dados[c.SEXTA - 1]
      )
    },
    sessoesPrescritas:
      Number(
        dados[
          c.SESSOES_PRESCRITAS - 1
        ]
      ) || 0,
    status: String(
      dados[c.STATUS - 1] || ''
    ).trim(),
    fisioterapeuta: String(
      dados[c.FISIOTERAPEUTA - 1] || ''
    ).trim(),
    desfecho: String(
      dados[c.DESFECHO - 1] || ''
    ).trim()
  };
}


function validarPacienteParaRenovacao_(
  paciente
) {
  if (!paciente.id) {
    throw new Error(
      'A linha selecionada não possui paciente.'
    );
  }

  if (
    normalizarTextoRenovacao_(
      paciente.desfecho
    ) !==
    normalizarTextoRenovacao_(
      CONFIG_RENOVACAO_TRATAMENTO
        .DESFECHO_EXIGIDO
    )
  ) {
    throw new Error(
      'Registre primeiro o desfecho "Renovação" para este paciente.'
    );
  }

  if (
    normalizarTextoRenovacao_(
      paciente.status
    ) !==
    normalizarTextoRenovacao_(
      CONFIG_RENOVACAO_TRATAMENTO
        .STATUS_EXIGIDO
    )
  ) {
    throw new Error(
      'O paciente precisa estar com status "Ciclo concluído" para planejar a renovação.'
    );
  }
}
function validarEntradaPlanejamentoRenovacao_(
  dados
) {
  if (!dados || typeof dados !== 'object') {
    throw new Error(
      'Os dados do formulário não foram recebidos.'
    );
  }

  const idPaciente = String(
    dados.idPaciente || ''
  ).trim();

  if (!idPaciente) {
    throw new Error(
      'O ID do paciente não foi informado.'
    );
  }

  const novoCiclo =
    Number(dados.novoCiclo);

  if (
    !Number.isInteger(novoCiclo) ||
    novoCiclo < 2
  ) {
    throw new Error(
      'O número do novo ciclo é inválido.'
    );
  }

  const sessoesPrescritas =
    Number(dados.sessoesPrescritas);

  if (
    !Number.isInteger(
      sessoesPrescritas
    ) ||
    sessoesPrescritas < 1 ||
    sessoesPrescritas >
      CONFIG_RENOVACAO_TRATAMENTO
        .MAXIMO_SESSOES
  ) {
    throw new Error(
      'Informe de 1 a ' +
      CONFIG_RENOVACAO_TRATAMENTO
        .MAXIMO_SESSOES +
      ' sessões para o novo ciclo.'
    );
  }

  const diasRecebidos =
    dados.diasSemana || {};

  const diasSemana = {
    1: diasRecebidos[1] === true ||
      diasRecebidos['1'] === true,
    2: diasRecebidos[2] === true ||
      diasRecebidos['2'] === true,
    3: diasRecebidos[3] === true ||
      diasRecebidos['3'] === true,
    4: diasRecebidos[4] === true ||
      diasRecebidos['4'] === true,
    5: diasRecebidos[5] === true ||
      diasRecebidos['5'] === true
  };

  const possuiDia = Object.keys(
    diasSemana
  ).some(function(dia) {
    return diasSemana[dia];
  });

  if (!possuiDia) {
    throw new Error(
      'Selecione pelo menos um dia da semana.'
    );
  }

  const horario =
    chaveHorarioRenovacao_(
      dados.horario
    );

  if (!horario) {
    throw new Error(
      'Selecione o horário do novo ciclo.'
    );
  }

  const tipoAtendimento = String(
    dados.tipoAtendimento || ''
  ).trim();

  if (!tipoAtendimento) {
    throw new Error(
      'Selecione o tipo de atendimento.'
    );
  }

  const fisioterapeuta = String(
    dados.fisioterapeuta || ''
  ).trim();

  if (!fisioterapeuta) {
    throw new Error(
      'Selecione o fisioterapeuta.'
    );
  }

  return {
    idPaciente: idPaciente,
    novoCiclo: novoCiclo,
    sessoesPrescritas:
      sessoesPrescritas,
    diasSemana: diasSemana,
    horario: horario,
    tipoAtendimento:
      tipoAtendimento,
    fisioterapeuta:
      fisioterapeuta
  };
}


function obterUltimoCicloPacienteRenovacao_(
  abaAgendamentos,
  idPaciente
) {
  const ultimaLinha =
    abaAgendamentos.getLastRow();

  if (ultimaLinha < 2) {
    return 0;
  }

  const dados = abaAgendamentos
    .getRange(
      2,
      CONFIG_RENOVACAO_TRATAMENTO
        .COLUNAS_AGENDAMENTOS
        .ID_PACIENTE,
      ultimaLinha - 1,
      CONFIG_RENOVACAO_TRATAMENTO
        .COLUNAS_AGENDAMENTOS
        .CICLO_NUMERO -
        CONFIG_RENOVACAO_TRATAMENTO
          .COLUNAS_AGENDAMENTOS
          .ID_PACIENTE +
        1
    )
    .getValues();

  const idProcurado =
    normalizarTextoRenovacao_(
      idPaciente
    );

  let maiorCiclo = 0;

  dados.forEach(function(linha) {
    const idAtual =
      normalizarTextoRenovacao_(
        linha[0]
      );

    if (idAtual !== idProcurado) {
      return;
    }

    const ciclo = Number(
      linha[
        CONFIG_RENOVACAO_TRATAMENTO
          .COLUNAS_AGENDAMENTOS
          .CICLO_NUMERO -
          CONFIG_RENOVACAO_TRATAMENTO
            .COLUNAS_AGENDAMENTOS
            .ID_PACIENTE
      ]
    );

    if (
      Number.isFinite(ciclo) &&
      ciclo > maiorCiclo
    ) {
      maiorCiclo = ciclo;
    }
  });

  return maiorCiclo;
}


function localizarLinhaPacienteRenovacao_(
  abaCadastro,
  idPaciente
) {
  const ultimaLinha =
    abaCadastro.getLastRow();

  if (ultimaLinha < 2) {
    return 0;
  }

  const ids = abaCadastro
    .getRange(
      2,
      CONFIG_RENOVACAO_TRATAMENTO
        .COLUNAS_CADASTRO
        .ID_PACIENTE,
      ultimaLinha - 1,
      1
    )
    .getDisplayValues();

  const idProcurado =
    normalizarTextoRenovacao_(
      idPaciente
    );

  for (
    let indice = 0;
    indice < ids.length;
    indice++
  ) {
    if (
      normalizarTextoRenovacao_(
        ids[indice][0]
      ) === idProcurado
    ) {
      return indice + 2;
    }
  }

  return 0;
}


function lerOpcoesColunaRenovacao_(
  aba,
  coluna,
  formatarComoHorario
) {
  const ultimaLinha = aba.getLastRow();

  if (ultimaLinha < 2) {
    return [];
  }

  const valores = formatarComoHorario
    ? aba
        .getRange(
          2,
          coluna,
          ultimaLinha - 1,
          1
        )
        .getValues()
    : aba
        .getRange(
          2,
          coluna,
          ultimaLinha - 1,
          1
        )
        .getDisplayValues();

  const unicos = {};
  const opcoes = [];

  valores.forEach(function(linha) {
    const valor = formatarComoHorario
      ? chaveHorarioRenovacao_(linha[0])
      : String(linha[0] || '').trim();

    if (!valor) {
      return;
    }

    const chave =
      normalizarTextoRenovacao_(
        valor
      );

    if (unicos[chave]) {
      return;
    }

    unicos[chave] = true;
    opcoes.push(valor);
  });

  return opcoes;
}
function validarListasPlanejamentoRenovacao_(
  dadosIniciais
) {
  if (dadosIniciais.horarios.length === 0) {
    throw new Error(
      'Não existem horários cadastrados na aba "Horários".'
    );
  }

  if (
    dadosIniciais
      .tiposAtendimento
      .length === 0
  ) {
    throw new Error(
      'Não existem tipos cadastrados na aba "Tipos de Grupo".'
    );
  }

  if (
    dadosIniciais
      .fisioterapeutas
      .length === 0
  ) {
    throw new Error(
      'Não existem fisioterapeutas cadastrados na aba "Fisioterapeutas".'
    );
  }
}


function validarOpcaoCadastradaRenovacao_(
  valorRecebido,
  opcoes,
  nomeCampo
) {
  const procurado =
    normalizarTextoRenovacao_(
      valorRecebido
    );

  for (
    let indice = 0;
    indice < opcoes.length;
    indice++
  ) {
    if (
      normalizarTextoRenovacao_(
        opcoes[indice]
      ) === procurado
    ) {
      return opcoes[indice];
    }
  }

  throw new Error(
    'O ' +
    nomeCampo +
    ' selecionado não está mais disponível. Feche o formulário e tente novamente.'
  );
}


function obterCapacidadeRenovacao_(
  tipoAtendimento
) {
  const tipo =
    normalizarTextoRenovacao_(
      tipoAtendimento
    );

  if (
    tipo ===
      'atendimento com maior supervisao' ||
    tipo === 'maior supervisao'
  ) {
    return 2;
  }

  const gruposComCapacidadeSeis = [
    'grupo de mmss',
    'grupo de mmii',
    'grupo de coluna'
  ];

  if (
    gruposComCapacidadeSeis.indexOf(
      tipo
    ) !== -1
  ) {
    return 6;
  }

  throw new Error(
    'O tipo de atendimento informado não é reconhecido.'
  );
}


function obterAbaObrigatoriaRenovacao_(
  ss,
  nome
) {
  const aba = ss.getSheetByName(nome);

  if (!aba) {
    throw new Error(
      'A aba necessária "' +
      nome +
      '" não foi encontrada.'
    );
  }

  return aba;
}


function valorAtivoRenovacao_(valor) {
  if (
    valor === true ||
    Number(valor) === 1
  ) {
    return true;
  }

  const texto =
    normalizarTextoRenovacao_(
      valor
    );

  return (
    texto === 'sim' ||
    texto === 'x' ||
    texto === 'verdadeiro'
  );
}


function chaveHorarioRenovacao_(valor) {
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

  const texto = String(
    valor || ''
  ).trim();

  const resultado = texto.match(
    /(\d{1,2}):(\d{2})/
  );

  if (!resultado) {
    return '';
  }

  const horas =
    Number(resultado[1]);
  const minutos =
    Number(resultado[2]);

  if (
    horas < 0 ||
    horas > 23 ||
    minutos < 0 ||
    minutos > 59
  ) {
    return '';
  }

  return (
    String(horas).padStart(2, '0') +
    ':' +
    String(minutos).padStart(2, '0')
  );
}


function normalizarTextoRenovacao_(valor) {
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
