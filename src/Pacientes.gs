const CONFIG_MODULO_PACIENTES = {
  ABA_CADASTRO: 'Cadastro de Pacientes',
  ABA_HORARIOS: 'Horários',
  ABA_TIPOS_GRUPO: 'Tipos de Grupo',
  ABA_FISIOTERAPEUTAS: 'Fisioterapeutas',
  ABA_AGENDAMENTOS: 'Agendamentos',
  ABA_VAGAS_REGULACAO: 'Vagas para Regulação',

  COLUNAS: {
    ID: 1,
    PRONTUARIO: 2,
    NOME: 3,
    CPF: 4,
    TELEFONE: 5,
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

    DATA_AVALIACAO: 17,
    HORARIO_AVALIACAO: 18,
    DATA_INICIO: 19,
    DATA_TERMINO: 20,

    STATUS: 21,
    FISIOTERAPEUTA: 22,
    OBSERVACAO: 23,
    DESFECHO: 24
  }
};


/**
 * Abre o formulário de cadastro.
 */
function abrirFormularioNovoPaciente() {
  const html = HtmlService
    .createHtmlOutputFromFile('NovoPaciente')
    .setWidth(780)
    .setHeight(680);

  SpreadsheetApp.getUi().showModalDialog(
    html,
    'SIGAF — Novo paciente'
  );
}


/**
 * Carrega as opções do formulário.
 */
function obterOpcoesFormularioPaciente() {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const abaHorarios = ss.getSheetByName(
    CONFIG_MODULO_PACIENTES.ABA_HORARIOS
  );

  const abaTiposGrupo = ss.getSheetByName(
    CONFIG_MODULO_PACIENTES
      .ABA_TIPOS_GRUPO
  );

  const abaFisioterapeutas =
    ss.getSheetByName(
      CONFIG_MODULO_PACIENTES
        .ABA_FISIOTERAPEUTAS
    );

  if (!abaHorarios) {
    throw new Error(
      'A aba "Horários" não foi encontrada.'
    );
  }

  if (!abaTiposGrupo) {
    throw new Error(
      'A aba "Tipos de Grupo" não foi encontrada.'
    );
  }

  if (!abaFisioterapeutas) {
    throw new Error(
      'A aba "Fisioterapeutas" não foi encontrada.'
    );
  }

  const horarios =
    obterValoresUnicosColuna_(
      abaHorarios,
      1
    )
      .map(function(valor) {
        return formatarHorarioFormulario_(
          valor
        );
      })
      .filter(function(valor) {
        return Boolean(valor);
      });

  const tiposAtendimento =
    obterValoresUnicosColuna_(
      abaTiposGrupo,
      1
    )
      .map(function(valor) {
        return String(valor).trim();
      })
      .filter(function(valor) {
        return Boolean(valor);
      });

  const fisioterapeutas =
    obterValoresUnicosColuna_(
      abaFisioterapeutas,
      2
    )
      .map(function(valor) {
        return String(valor).trim();
      })
      .filter(function(valor) {
        return Boolean(valor);
      });

  return {
    horarios: horarios,
    tiposAtendimento: tiposAtendimento,
    fisioterapeutas: fisioterapeutas
  };
}


/**
 * Salva o paciente no Cadastro e registra
 * automaticamente a avaliação em Agendamentos.
 */
function salvarNovoPaciente(dados) {
  const lock =
    LockService.getDocumentLock();

  let bloqueioObtido = false;

  try {
    lock.waitLock(30000);
    bloqueioObtido = true;

    const ss =
      SpreadsheetApp.getActiveSpreadsheet();

    const aba = ss.getSheetByName(
      CONFIG_MODULO_PACIENTES.ABA_CADASTRO
    );

    if (!aba) {
      throw new Error(
        'A aba "Cadastro de Pacientes" não foi encontrada.'
      );
    }

    const paciente =
      validarDadosNovoPaciente_(dados);

    verificarDuplicidadeNovoPaciente_(
      aba,
      paciente.cpf
    );

    const id =
      gerarProximoIdPacienteSeguro_(aba);

    const prontuario =
      gerarProximoProntuarioPaciente_(aba);

    const linhaDestino =
      encontrarProximaLinhaCadastro_(aba);

    const limiteGrupo =
      obterLimiteGrupoPaciente_(
        paciente.tipoAtendimento
      );

    const sessoesPrescritas =
      paciente.sessoesPrescritas;

    const sessoesRealizadas = 0;

    const sessoesRestantes =
      sessoesPrescritas;

    const status =
      definirStatusInicialPaciente_(
        paciente.dataAvaliacao,
        sessoesPrescritas
      );

    const registro = [[
      id,
      prontuario,
      paciente.nome,
      paciente.cpfFormatado,
      paciente.telefone,
      paciente.horarioSessao || '',
      paciente.tipoAtendimento || '',
      limiteGrupo,

      paciente.dias.segunda,
      paciente.dias.terca,
      paciente.dias.quarta,
      paciente.dias.quinta,
      paciente.dias.sexta,

      sessoesPrescritas,
      sessoesRealizadas,
      sessoesRestantes,

      paciente.dataAvaliacao,
      paciente.horarioAvaliacao,
      '',
      '',

      status,
      paciente.fisioterapeuta || '',
      paciente.observacao || '',
      ''
    ]];

    aba.getRange(
      linhaDestino,
      1,
      1,
      24
    ).setValues(registro);

    formatarLinhaNovoPaciente_(
      aba,
      linhaDestino
    );

    try {
      registrarAvaliacaoPacienteNovo_(
        ss,
        {
          id: id,
          prontuario: prontuario,
          nome: paciente.nome,
          dataAvaliacao:
            paciente.dataAvaliacao,
          horarioAvaliacao:
            paciente.horarioAvaliacao,
          fisioterapeuta:
            paciente.fisioterapeuta,
          exigirVagaRegulacao: true
        }
      );
    } catch (erroAvaliacao) {
      /*
       * Evita deixar o paciente cadastrado pela
       * metade caso a avaliação não seja criada.
       */
      aba.getRange(
        linhaDestino,
        1,
        1,
        24
      ).clearContent();

      throw new Error(
        'O paciente não foi cadastrado porque a avaliação não pôde ser registrada. ' +
          (
            erroAvaliacao &&
            erroAvaliacao.message
              ? erroAvaliacao.message
              : String(erroAvaliacao)
          )
      );
    }

    SpreadsheetApp.flush();

    return {
      sucesso: true,
      id: id,
      prontuario: prontuario,
      linha: linhaDestino,
      nome: paciente.nome,
      status: status,

      mensagem:
        'Paciente cadastrado com sucesso.\n\n' +
        'ID: ' + id +
        '\nProntuário: ' + prontuario +
        '\nPaciente: ' + paciente.nome +
        '\nStatus: ' + status
    };
  } catch (erro) {
    return {
      sucesso: false,

      mensagem:
        erro && erro.message
          ? erro.message
          : String(erro)
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


/**
 * Valida os dados recebidos do formulário.
 */function validarDadosNovoPaciente_(dados) {
  if (!dados) {
    throw new Error(
      'Nenhuma informação foi recebida do formulário.'
    );
  }

  const nome = String(
    dados.nome || ''
  ).trim();

  const telefone = String(
    dados.telefone || ''
  ).trim();

  const cpfNumeros = String(
    dados.cpf || ''
  ).replace(/\D/g, '');

  const dataAvaliacao =
    converterDataFormulario_(
      dados.dataAvaliacao
    );

  const horarioAvaliacao =
    converterHorarioFormulario_(
      dados.horarioAvaliacao
    );

  const horarioSessao =
    dados.horarioSessao
      ? converterHorarioFormulario_(
          dados.horarioSessao
        )
      : '';

  const tipoAtendimento = String(
    dados.tipoAtendimento || ''
  ).trim();

  const fisioterapeuta = String(
    dados.fisioterapeuta || ''
  ).trim();

  const observacao = String(
    dados.observacao || ''
  ).trim();

  const sessoesPrescritas = Number(
    dados.sessoesPrescritas || 0
  );

  if (!nome) {
    throw new Error(
      'Informe o nome completo do paciente.'
    );
  }

  if (!telefone) {
    throw new Error(
      'Informe o telefone do paciente.'
    );
  }

  if (!dataAvaliacao) {
    throw new Error(
      'Informe uma data de avaliação válida.'
    );
  }

  if (!horarioAvaliacao) {
    throw new Error(
      'Informe o horário da avaliação.'
    );
  }

  if (!fisioterapeuta) {
    throw new Error(
      'Informe o fisioterapeuta responsável pela avaliação.'
    );
  }

  if (
    cpfNumeros &&
    (
      cpfNumeros.length !== 11 ||
      !validarCpfNovoPaciente_(
        cpfNumeros
      )
    )
  ) {
    throw new Error(
      'O CPF informado é inválido.'
    );
  }

  if (
    !Number.isInteger(
      sessoesPrescritas
    ) ||
    sessoesPrescritas < 0 ||
    sessoesPrescritas > 20
  ) {
    throw new Error(
      'As sessões prescritas devem ser um número inteiro entre 0 e 20.'
    );
  }

  const dias = {
    segunda: Boolean(dados.segunda),
    terca: Boolean(dados.terca),
    quarta: Boolean(dados.quarta),
    quinta: Boolean(dados.quinta),
    sexta: Boolean(dados.sexta)
  };

  const possuiDiaSelecionado =
    Object.keys(dias)
      .some(function(chave) {
        return dias[chave];
      });

  if (
    sessoesPrescritas > 0 &&
    !tipoAtendimento
  ) {
    throw new Error(
      'Selecione o tipo de atendimento quando houver sessões prescritas.'
    );
  }

  if (
    sessoesPrescritas > 0 &&
    !horarioSessao
  ) {
    throw new Error(
      'Selecione um horário preferencial para as sessões.'
    );
  }

  if (
    sessoesPrescritas > 0 &&
    !possuiDiaSelecionado
  ) {
    throw new Error(
      'Marque pelo menos um dia de preferência para as sessões.'
    );
  }

  return {
    nome:
      normalizarNomePaciente_(nome),

    cpf:
      cpfNumeros,

    cpfFormatado:
      cpfNumeros
        ? formatarCpfPaciente_(
            cpfNumeros
          )
        : '',

    telefone:
      telefone,

    dataAvaliacao:
      dataAvaliacao,

    horarioAvaliacao:
      horarioAvaliacao,

    horarioSessao:
      horarioSessao,

    tipoAtendimento:
      tipoAtendimento,

    fisioterapeuta:
      fisioterapeuta,

    sessoesPrescritas:
      sessoesPrescritas,

    dias:
      dias,

    observacao:
      observacao
  };
}


/**
 * Verifica se o CPF já está cadastrado.
 */
function verificarDuplicidadeNovoPaciente_(
  aba,
  cpf
) {
  const cpfProcurado = String(
    cpf || ''
  ).replace(/\D/g, '');

  if (!cpfProcurado) {
    return;
  }

  const ultimaLinha =
    aba.getLastRow();

  if (ultimaLinha < 2) {
    return;
  }

  const dados = aba
    .getRange(
      2,
      CONFIG_MODULO_PACIENTES
        .COLUNAS.ID,
      ultimaLinha - 1,
      4
    )
    .getDisplayValues();

  dados.forEach(function(
    linha,
    indice
  ) {
    const linhaPlanilha =
      indice + 2;

    const idExistente = String(
      linha[0] || ''
    ).trim();

    if (!idExistente) {
      return;
    }

    const cpfExistente = String(
      linha[3] || ''
    ).replace(/\D/g, '');

    if (
      cpfExistente &&
      cpfExistente === cpfProcurado
    ) {
      throw new Error(
        'Este CPF já pertence a um paciente cadastrado na linha ' +
          linhaPlanilha +
          '.'
      );
    }
  });
}


/**
 * Gera o próximo ID do paciente.
 */
function gerarProximoIdPacienteSeguro_(
  aba
) {
  const ultimaLinha =
    aba.getLastRow();

  if (ultimaLinha < 2) {
    return 'PAC-00001';
  }

  const ids = aba
    .getRange(
      2,
      CONFIG_MODULO_PACIENTES
        .COLUNAS.ID,
      ultimaLinha - 1,
      1
    )
    .getDisplayValues()
    .flat();

  let maiorNumero = 0;

  ids.forEach(function(id) {
    const resultado = String(
      id || ''
    )
      .trim()
      .match(/^PAC-(\d+)$/i);

    if (!resultado) {
      return;
    }

    maiorNumero = Math.max(
      maiorNumero,
      Number(resultado[1])
    );
  });

  return (
    'PAC-' +
    String(maiorNumero + 1)
      .padStart(5, '0')
  );
}


/**
 * Gera o próximo número de prontuário.
 */
function gerarProximoProntuarioPaciente_(
  aba
) {
  const ultimaLinha =
    aba.getLastRow();

  if (ultimaLinha < 2) {
    return 1;
  }

  const prontuarios = aba
    .getRange(
      2,
      CONFIG_MODULO_PACIENTES
        .COLUNAS.PRONTUARIO,
      ultimaLinha - 1,
      1
    )
    .getDisplayValues()
    .flat();

  let maiorProntuario = 0;

  prontuarios.forEach(function(valor) {
    const texto = String(
      valor || ''
    ).trim();

    if (!/^\d+$/.test(texto)) {
      return;
    }

    const numero = Number(texto);

    if (
      Number.isInteger(numero) &&
      numero > maiorProntuario
    ) {
      maiorProntuario = numero;
    }
  });

  return maiorProntuario + 1;
}


/**
 * Encontra a primeira linha vazia.
 */function encontrarProximaLinhaCadastro_(
  aba
) {
  const ultimaLinha =
    aba.getLastRow();

  if (ultimaLinha < 2) {
    return 2;
  }

  const ids = aba
    .getRange(
      2,
      CONFIG_MODULO_PACIENTES
        .COLUNAS.ID,
      ultimaLinha - 1,
      1
    )
    .getDisplayValues();

  for (
    let indice = 0;
    indice < ids.length;
    indice++
  ) {
    const id = String(
      ids[indice][0] || ''
    ).trim();

    if (!id) {
      return indice + 2;
    }
  }

  return ultimaLinha + 1;
}


/**
 * Define a capacidade pelo tipo de atendimento.
 */
function obterLimiteGrupoPaciente_(
  tipoAtendimento
) {
  const tipo =
    normalizarTextoPaciente_(
      tipoAtendimento
    );

  if (!tipo) {
    return '';
  }

  /*
   * A aba "Tipos de Grupo" é a fonte oficial da capacidade.
   * Assim, Cadastro, Edição e demais módulos que reutilizam
   * esta função não ficam com um número fixo diferente da
   * configuração da planilha.
   */
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const abaTiposGrupo =
    ss.getSheetByName(
      CONFIG_MODULO_PACIENTES
        .ABA_TIPOS_GRUPO
    );

  if (!abaTiposGrupo) {
    throw new Error(
      'A aba "Tipos de Grupo" não foi encontrada.'
    );
  }

  const ultimaLinha =
    abaTiposGrupo.getLastRow();

  if (ultimaLinha < 2) {
    throw new Error(
      'A aba "Tipos de Grupo" não possui configurações.'
    );
  }

  const dados =
    abaTiposGrupo
      .getRange(
        2,
        1,
        ultimaLinha - 1,
        2
      )
      .getDisplayValues();

  for (
    let indice = 0;
    indice < dados.length;
    indice++
  ) {
    const tipoConfigurado =
      normalizarTextoPaciente_(
        dados[indice][0]
      );

    if (
      tipoConfigurado !== tipo
    ) {
      continue;
    }

    const capacidade =
      Number(
        String(
          dados[indice][1] || ''
        ).replace(',', '.')
      );

    if (
      !Number.isInteger(
        capacidade
      ) ||
      capacidade < 1
    ) {
      throw new Error(
        'A capacidade configurada para "' +
        tipoAtendimento +
        '" é inválida na aba "Tipos de Grupo".'
      );
    }

    return capacidade;
  }

  throw new Error(
    'O tipo de atendimento "' +
    tipoAtendimento +
    '" não foi encontrado na aba "Tipos de Grupo".'
  );
}


/**
 * Define o status inicial.
 */
function definirStatusInicialPaciente_(
  dataAvaliacao,
  sessoesPrescritas
) {
  if (sessoesPrescritas > 0) {
    return (
      'Avaliado – aguardando agendamento'
    );
  }

  if (dataAvaliacao) {
    return 'Avaliação agendada';
  }

  return 'Inativo';
}


/**
 * Formata a linha inserida no Cadastro.
 */
function formatarLinhaNovoPaciente_(
  aba,
  linha
) {
  aba.getRange(
    linha,
    1,
    1,
    24
  )
    .setVerticalAlignment('middle')
    .setWrap(true);

  aba.getRange(
    linha,
    CONFIG_MODULO_PACIENTES.COLUNAS.ID
  ).setNumberFormat('@');

  aba.getRange(
    linha,
    CONFIG_MODULO_PACIENTES
      .COLUNAS.PRONTUARIO
  )
    .setNumberFormat('0')
    .setBackground('#eeeeee')
    .setNote(
      'Prontuário gerado automaticamente pelo SIGAF.'
    );

  aba.getRange(
    linha,
    CONFIG_MODULO_PACIENTES.COLUNAS.CPF
  ).setNumberFormat('@');

  aba.getRange(
    linha,
    CONFIG_MODULO_PACIENTES
      .COLUNAS.TELEFONE
  ).setNumberFormat('@');

  aba.getRange(
    linha,
    CONFIG_MODULO_PACIENTES
      .COLUNAS.HORARIO
  ).setNumberFormat('HH:mm');

  aba.getRange(
    linha,
    CONFIG_MODULO_PACIENTES
      .COLUNAS.DATA_AVALIACAO
  ).setNumberFormat('dd/MM/yyyy');

  aba.getRange(
    linha,
    CONFIG_MODULO_PACIENTES
      .COLUNAS.HORARIO_AVALIACAO
  ).setNumberFormat('HH:mm');

  aba.getRange(
    linha,
    CONFIG_MODULO_PACIENTES
      .COLUNAS.DATA_INICIO,
    1,
    2
  ).setNumberFormat('dd/MM/yyyy');

  const intervaloDias =
    aba.getRange(
      linha,
      CONFIG_MODULO_PACIENTES
        .COLUNAS.SEGUNDA,
      1,
      5
    );

  const valoresDias =
    intervaloDias.getValues();

  intervaloDias.insertCheckboxes();
  intervaloDias.setValues(valoresDias);

  aba.getRange(
    linha,
    CONFIG_MODULO_PACIENTES
      .COLUNAS.SESSOES_REALIZADAS
  )
    .setBackground('#eeeeee')
    .setNote(
      'Campo controlado automaticamente pelo SIGAF.'
    );

  aba.getRange(
    linha,
    CONFIG_MODULO_PACIENTES
      .COLUNAS.SESSOES_RESTANTES
  )
    .setBackground('#eeeeee')
    .setNote(
      'Campo calculado automaticamente pelo SIGAF.'
    );

  aba.getRange(
    linha,
    CONFIG_MODULO_PACIENTES
      .COLUNAS.LIMITE_GRUPO
  )
    .setBackground('#eeeeee')
    .setNote(
      'Campo calculado automaticamente conforme o tipo de atendimento.'
    );
}


/**
 * Obtém valores únicos de uma coluna.
 */
function obterValoresUnicosColuna_(
  aba,
  coluna
) {
  if (!aba) {
    return [];
  }

  const ultimaLinha =
    aba.getLastRow();

  if (ultimaLinha < 2) {
    return [];
  }

  const valores = aba
    .getRange(
      2,
      coluna,
      ultimaLinha - 1,
      1
    )
    .getValues()
    .flat()
    .filter(function(valor) {
      return (
        String(valor || '').trim() !== ''
      );
    });

  const chavesEncontradas = {};
  const resultado = [];

  valores.forEach(function(valor) {
    const valorParaChave =
      valor instanceof Date
        ? formatarHorarioFormulario_(valor)
        : String(valor);

    const chave =
      normalizarTextoPaciente_(
        valorParaChave
      );

    if (
      !chave ||
      chavesEncontradas[chave]
    ) {
      return;
    }

    chavesEncontradas[chave] = true;
    resultado.push(valor);
  });

  return resultado;
}


/**
 * Converte a data AAAA-MM-DD.
 */
function converterDataFormulario_(texto) {
  const resultado = String(
    texto || ''
  ).match(
    /^(\d{4})-(\d{2})-(\d{2})$/
  );

  if (!resultado) {
    return null;
  }

  const ano = Number(resultado[1]);
  const mes = Number(resultado[2]) - 1;
  const dia = Number(resultado[3]);

  const data = new Date(
    ano,
    mes,
    dia
  );

  data.setHours(0, 0, 0, 0);

  if (
    data.getFullYear() !== ano ||
    data.getMonth() !== mes ||
    data.getDate() !== dia
  ) {
    return null;
  }

  return data;
}


/**
 * Converte HH:mm em horário da planilha.
 */function converterHorarioFormulario_(texto) {
  const resultado = String(
    texto || ''
  ).match(
    /^(\d{1,2}):(\d{2})$/
  );

  if (!resultado) {
    return '';
  }

  const horas = Number(resultado[1]);
  const minutos = Number(resultado[2]);

  if (
    horas < 0 ||
    horas > 23 ||
    minutos < 0 ||
    minutos > 59
  ) {
    return '';
  }

  return new Date(
    1899,
    11,
    30,
    horas,
    minutos,
    0
  );
}


/**
 * Formata um horário como HH:mm.
 */
function formatarHorarioFormulario_(valor) {
  if (valor instanceof Date) {
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
    return texto;
  }

  return (
    String(
      Number(resultado[1])
    ).padStart(2, '0') +
    ':' +
    resultado[2]
  );
}


/**
 * Formata o CPF.
 */
function formatarCpfPaciente_(cpf) {
  return (
    cpf.substring(0, 3) +
    '.' +
    cpf.substring(3, 6) +
    '.' +
    cpf.substring(6, 9) +
    '-' +
    cpf.substring(9, 11)
  );
}


/**
 * Valida o CPF.
 */
function validarCpfNovoPaciente_(cpf) {
  if (!/^\d{11}$/.test(cpf)) {
    return false;
  }

  if (/^(\d)\1{10}$/.test(cpf)) {
    return false;
  }

  let soma = 0;

  for (
    let indice = 0;
    indice < 9;
    indice++
  ) {
    soma +=
      Number(cpf.charAt(indice)) *
      (10 - indice);
  }

  let digito1 =
    11 - (soma % 11);

  if (digito1 >= 10) {
    digito1 = 0;
  }

  if (
    digito1 !==
    Number(cpf.charAt(9))
  ) {
    return false;
  }

  soma = 0;

  for (
    let indice = 0;
    indice < 10;
    indice++
  ) {
    soma +=
      Number(cpf.charAt(indice)) *
      (11 - indice);
  }

  let digito2 =
    11 - (soma % 11);

  if (digito2 >= 10) {
    digito2 = 0;
  }

  return (
    digito2 ===
    Number(cpf.charAt(10))
  );
}


/**
 * Padroniza o nome do paciente.
 */
function normalizarNomePaciente_(nome) {
  return String(nome || '')
    .trim()
    .toLocaleLowerCase('pt-BR')
    .replace(
      /(^|\s)([a-záàâãéèêíïóôõöúç])/g,
      function(
        textoCompleto,
        espaco,
        letra
      ) {
        return (
          espaco +
          letra.toLocaleUpperCase(
            'pt-BR'
          )
        );
      }
    );
}


/**
 * Normaliza textos para comparação.
 */
function normalizarTextoPaciente_(valor) {
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
      /\s*-\s*/g,
      ' - '
    )
    .replace(
      /\s+/g,
      ' '
    );
}


/**
 * Cria as avaliações ausentes dos pacientes
 * que já estão cadastrados.
 */
function sincronizarAvaliacoesAgendadasAusentes() {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const ui =
    SpreadsheetApp.getUi();

  const lock =
    LockService.getDocumentLock();

  let bloqueioObtido = false;

  try {
    lock.waitLock(30000);
    bloqueioObtido = true;

    const abaCadastro =
      ss.getSheetByName(
        CONFIG_MODULO_PACIENTES
          .ABA_CADASTRO
      );

    if (!abaCadastro) {
      throw new Error(
        'A aba "Cadastro de Pacientes" não foi encontrada.'
      );
    }

    const ultimaLinha =
      abaCadastro.getLastRow();

    if (ultimaLinha < 2) {
      ui.alert(
        'Nenhum paciente cadastrado.',
        ui.ButtonSet.OK
      );

      return;
    }

    const dados =
      abaCadastro.getRange(
        2,
        1,
        ultimaLinha - 1,
        24
      ).getValues();

    let criadas = 0;
    let jaExistentes = 0;

    dados.forEach(function(linha) {
      const id = String(
        linha[
          CONFIG_MODULO_PACIENTES
            .COLUNAS.ID - 1
        ] || ''
      ).trim();

      const status =
        normalizarTextoPaciente_(
          linha[
            CONFIG_MODULO_PACIENTES
              .COLUNAS.STATUS - 1
          ]
        );

      const dataAvaliacao =
        linha[
          CONFIG_MODULO_PACIENTES
            .COLUNAS.DATA_AVALIACAO - 1
        ];

      const horarioAvaliacao =
        linha[
          CONFIG_MODULO_PACIENTES
            .COLUNAS.HORARIO_AVALIACAO - 1
        ];

      if (
        !id ||
        status !==
          'avaliacao agendada' ||
        !(
          dataAvaliacao instanceof Date
        ) ||
        !horarioAvaliacao
      ) {
        return;
      }

      const resultado =
        registrarAvaliacaoPacienteNovo_(
          ss,
          {
            id: id,

            prontuario: String(
              linha[
                CONFIG_MODULO_PACIENTES
                  .COLUNAS.PRONTUARIO - 1
              ] || ''
            ).trim(),

            nome: String(
              linha[
                CONFIG_MODULO_PACIENTES
                  .COLUNAS.NOME - 1
              ] || ''
            ).trim(),

            dataAvaliacao:
              dataAvaliacao,

            horarioAvaliacao:
              horarioAvaliacao,

            fisioterapeuta: String(
              linha[
                CONFIG_MODULO_PACIENTES
                  .COLUNAS.FISIOTERAPEUTA -
                  1
              ] || ''
            ).trim()
          }
        );

      if (resultado.criada) {
        criadas++;
      } else {
        jaExistentes++;
      }
    });

    SpreadsheetApp.flush();

    ui.alert(
      'Avaliações sincronizadas',
      'Avaliações criadas: ' +
        criadas +
        '\nAvaliações que já existiam: ' +
        jaExistentes,
      ui.ButtonSet.OK
    );
  } catch (erro) {
    ui.alert(
      'Erro ao sincronizar avaliações',
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
        // Liberação automática.
      }
    }
  }
}


/**
 * Cria uma avaliação em Agendamentos.
 */function registrarAvaliacaoPacienteNovo_(
  ss,
  paciente
) {
  const abaAgendamentos =
    ss.getSheetByName(
      CONFIG_MODULO_PACIENTES
        .ABA_AGENDAMENTOS
    );

  if (!abaAgendamentos) {
    throw new Error(
      'A aba "Agendamentos" não foi encontrada.'
    );
  }

  if (
    !(
      paciente.dataAvaliacao
      instanceof Date
    )
  ) {
    throw new Error(
      'A data da avaliação é inválida.'
    );
  }

  if (!paciente.horarioAvaliacao) {
    throw new Error(
      'O horário da avaliação não foi informado.'
    );
  }

  if (
    paciente.exigirVagaRegulacao &&
    !String(
      paciente.fisioterapeuta || ''
    ).trim()
  ) {
    throw new Error(
      'O fisioterapeuta da avaliação não foi informado.'
    );
  }

  const existente =
    localizarAvaliacaoPacienteNovo_(
      abaAgendamentos,
      paciente.id,
      paciente.dataAvaliacao,
      paciente.horarioAvaliacao
    );

  if (existente) {
    return {
      criada: false,
      idAgendamento:
        existente.idAgendamento,
      linha: existente.linha
    };
  }

  let vagaRegulacao = null;

  if (paciente.exigirVagaRegulacao) {
    vagaRegulacao =
      localizarVagaDisponivelRegulacaoPacienteNovo_(
        ss,
        paciente.dataAvaliacao,
        paciente.horarioAvaliacao,
        paciente.fisioterapeuta
      );

    verificarHorarioAvaliacaoLivrePacienteNovo_(
      abaAgendamentos,
      paciente.dataAvaliacao,
      paciente.horarioAvaliacao,
      paciente.fisioterapeuta
    );
  }

  const idAgendamento =
    gerarProximoIdAgendamentoPacienteNovo_(
      abaAgendamentos
    );

  const agora = new Date();

  const dataAvaliacao =
    new Date(
      paciente.dataAvaliacao
    );

  dataAvaliacao.setHours(
    0,
    0,
    0,
    0
  );

  const registro = [[
    idAgendamento,
    paciente.id,
    paciente.prontuario,
    paciente.nome,
    '',
    '',
    dataAvaliacao,
    obterNomeDiaPacienteNovo_(
      dataAvaliacao
    ),
    paciente.horarioAvaliacao,
    paciente.fisioterapeuta || '',
    '',
    'Avaliação',
    '',
    '',
    1,
    'Agendado',
    '',
    'Não',
    'Sim',
    agora,
    agora,
    'Não'
  ]];

  const linhaDestino = Math.max(
    abaAgendamentos.getLastRow() + 1,
    2
  );

  abaAgendamentos.getRange(
    linhaDestino,
    1,
    1,
    22
  ).setValues(registro);

  abaAgendamentos.getRange(
    linhaDestino,
    7
  ).setNumberFormat('dd/MM/yyyy');

  abaAgendamentos.getRange(
    linhaDestino,
    9
  ).setNumberFormat('HH:mm');

  abaAgendamentos.getRange(
    linhaDestino,
    20,
    1,
    2
  ).setNumberFormat(
    'dd/MM/yyyy HH:mm'
  );

  if (vagaRegulacao) {
    try {
      vagaRegulacao.aba.getRange(
        vagaRegulacao.linha,
        6
      ).setValue('Agendada');
    } catch (erroVaga) {
      /*
       * Se a atualização da vaga falhar, desfaz
       * o agendamento para não deixar registros
       * incompletos ou divergentes.
       */
      abaAgendamentos.getRange(
        linhaDestino,
        1,
        1,
        22
      ).clearContent();

      throw new Error(
        'A avaliação não foi criada porque não foi possível reservar a vaga da Regulação. ' +
          (
            erroVaga && erroVaga.message
              ? erroVaga.message
              : String(erroVaga)
          )
      );
    }
  }

  return {
    criada: true,
    idAgendamento: idAgendamento,
    linha: linhaDestino
  };
}


/**
 * Localiza a vaga correspondente na aba
 * Vagas para Regulação.
 */
function localizarVagaDisponivelRegulacaoPacienteNovo_(
  ss,
  dataAvaliacao,
  horarioAvaliacao,
  fisioterapeuta
) {
  const abaVagas =
    ss.getSheetByName(
      CONFIG_MODULO_PACIENTES
        .ABA_VAGAS_REGULACAO
    );

  if (!abaVagas) {
    throw new Error(
      'A aba "Vagas para Regulação" não foi encontrada.'
    );
  }

  const ultimaLinha =
    abaVagas.getLastRow();

  if (ultimaLinha < 2) {
    throw new Error(
      'Não existem vagas geradas para a Regulação. Gere as vagas mensais antes de cadastrar o paciente.'
    );
  }

  const dados =
    abaVagas.getRange(
      2,
      1,
      ultimaLinha - 1,
      6
    ).getValues();

  const dataProcurada =
    chaveDataPacienteNovo_(
      dataAvaliacao
    );

  const horarioProcurado =
    chaveHorarioPacienteNovo_(
      horarioAvaliacao
    );

  const profissionalProcurado =
    normalizarTextoPaciente_(
      fisioterapeuta
    );

  for (
    let indice = 0;
    indice < dados.length;
    indice++
  ) {
    const linha = dados[indice];

    const data = linha[0];

    if (!(data instanceof Date)) {
      continue;
    }

    const corresponde =
      chaveDataPacienteNovo_(data) ===
        dataProcurada &&
      chaveHorarioPacienteNovo_(
        linha[2]
      ) === horarioProcurado &&
      normalizarTextoPaciente_(
        linha[3]
      ) === profissionalProcurado;

    if (!corresponde) {
      continue;
    }

    const situacao =
      normalizarTextoPaciente_(
        linha[5]
      );

    if (situacao !== 'disponivel') {
      throw new Error(
        'A vaga selecionada já está reservada ou indisponível. Atualize as vagas da Regulação e escolha outra data ou horário.'
      );
    }

    return {
      aba: abaVagas,
      linha: indice + 2
    };
  }

  throw new Error(
    'A data, o horário e o fisioterapeuta informados não correspondem a uma vaga disponível da Regulação.'
  );
}


/**
 * Confere se outro atendimento já ocupa
 * o mesmo fisioterapeuta, data e horário.
 */
function verificarHorarioAvaliacaoLivrePacienteNovo_(
  abaAgendamentos,
  dataAvaliacao,
  horarioAvaliacao,
  fisioterapeuta
) {
  const ultimaLinha =
    abaAgendamentos.getLastRow();

  if (ultimaLinha < 2) {
    return;
  }

  const dados =
    abaAgendamentos.getRange(
      2,
      1,
      ultimaLinha - 1,
      22
    ).getValues();

  const dataProcurada =
    chaveDataPacienteNovo_(
      dataAvaliacao
    );

  const horarioProcurado =
    chaveHorarioPacienteNovo_(
      horarioAvaliacao
    );

  const profissionalProcurado =
    normalizarTextoPaciente_(
      fisioterapeuta
    );

  const statusQueOcupam = [
    'agendado',
    'compareceu',
    'falta justificada',
    'falta nao justificada'
  ];

  const ocupado = dados.some(
    function(linha) {
      const data = linha[6];

      if (!(data instanceof Date)) {
        return false;
      }

      const status =
        normalizarTextoPaciente_(
          linha[15]
        );

      if (
        statusQueOcupam.indexOf(
          status
        ) === -1
      ) {
        return false;
      }

      return (
        chaveDataPacienteNovo_(data) ===
          dataProcurada &&
        chaveHorarioPacienteNovo_(
          linha[8]
        ) === horarioProcurado &&
        normalizarTextoPaciente_(
          linha[9]
        ) === profissionalProcurado
      );
    }
  );

  if (ocupado) {
    throw new Error(
      'O fisioterapeuta já possui um atendimento nesse dia e horário. Atualize as vagas da Regulação e selecione outra vaga.'
    );
  }
}


/**
 * Verifica se a avaliação já existe.
 */function localizarAvaliacaoPacienteNovo_(
  abaAgendamentos,
  idPaciente,
  dataAvaliacao,
  horarioAvaliacao
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
      22
    ).getValues();

  const idProcurado =
    normalizarTextoPaciente_(
      idPaciente
    );

  const dataProcurada =
    chaveDataPacienteNovo_(
      dataAvaliacao
    );

  const horarioProcurado =
    chaveHorarioPacienteNovo_(
      horarioAvaliacao
    );

  for (
    let indice = 0;
    indice < dados.length;
    indice++
  ) {
    const linha = dados[indice];

    const idAtual =
      normalizarTextoPaciente_(
        linha[1]
      );

    const evento =
      normalizarTextoPaciente_(
        linha[11]
      );

    const data = linha[6];
    const horario = linha[8];

    if (
      idAtual === idProcurado &&
      evento === 'avaliacao' &&
      data instanceof Date &&
      chaveDataPacienteNovo_(data) ===
        dataProcurada &&
      chaveHorarioPacienteNovo_(
        horario
      ) === horarioProcurado
    ) {
      return {
        linha: indice + 2,

        idAgendamento: String(
          linha[0] || ''
        ).trim()
      };
    }
  }

  return null;
}


/**
 * Gera o próximo ID de agendamento.
 */
function gerarProximoIdAgendamentoPacienteNovo_(
  abaAgendamentos
) {
  const ultimaLinha =
    abaAgendamentos.getLastRow();

  let maiorNumero = 0;

  if (ultimaLinha >= 2) {
    const ids =
      abaAgendamentos.getRange(
        2,
        1,
        ultimaLinha - 1,
        1
      ).getDisplayValues();

    ids.forEach(function(linha) {
      const resultado = String(
        linha[0] || ''
      ).match(/(\d+)$/);

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
  }

  return (
    'AG-' +
    String(maiorNumero + 1)
      .padStart(6, '0')
  );
}


/**
 * Retorna o nome do dia da semana.
 */
function obterNomeDiaPacienteNovo_(data) {
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
 * Cria uma chave para comparação de datas.
 */
function chaveDataPacienteNovo_(data) {
  return Utilities.formatDate(
    new Date(data),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );
}


/**
 * Cria uma chave para comparação de horários.
 */
function chaveHorarioPacienteNovo_(valor) {
  if (valor instanceof Date) {
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
    return texto;
  }

  return (
    String(
      Number(resultado[1])
    ).padStart(2, '0') +
    ':' +
    resultado[2]
  );
}
