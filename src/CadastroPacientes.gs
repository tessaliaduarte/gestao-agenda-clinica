const CONFIG_CADASTRO_PACIENTES = {
  ABAS: {
    CADASTRO: 'Cadastro de Pacientes',
    HORARIOS: 'Horários',
    TIPOS_GRUPO: 'Tipos de Grupo',
    FISIOTERAPEUTAS: 'Fisioterapeutas'
  },

  TIPOS_ATENDIMENTO: [
    'Atendimento com maior supervisão',
    'Grupo de MMSS',
    'Grupo de MMII',
    'Grupo de Coluna'
  ],

  STATUS: [
    'Avaliação agendada',
    'Avaliado – aguardando agendamento',
    'Em tratamento',
    'Ciclo concluído',
    'Inativo'
  ]
};


/**
 * Nome utilizado pelo menu do SIGAF.
 */
function configurarCadastroPacientes() {
  corrigirConfiguracaoCadastroPacientes();
}


/**
 * Reaplica as configurações do Cadastro de Pacientes.
 */
function corrigirConfiguracaoCadastroPacientes() {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const ui =
    SpreadsheetApp.getUi();

  const abaCadastro =
    ss.getSheetByName(
      CONFIG_CADASTRO_PACIENTES
        .ABAS.CADASTRO
    );

  if (!abaCadastro) {
    ui.alert(
      'A aba "Cadastro de Pacientes" não foi encontrada.'
    );

    return;
  }

  try {
    /*
     * Atualiza primeiro a fonte oficial dos tipos.
     * Isso impede que a validação da coluna G rejeite
     * Grupo de MMSS, MMII ou Coluna.
     */
    sincronizarTiposAtendimentoCadastro_(
      ss
    );

    corrigirCabecalhoCadastro_(
      abaCadastro
    );

    corrigirListasCadastro_(
      ss,
      abaCadastro
    );

    corrigirValidacaoSessoes_(
      abaCadastro
    );

    corrigirCoresStatusCadastro_(
      abaCadastro
    );

    SpreadsheetApp.flush();

    ui.alert(
      'Configuração corrigida',
      'Os tipos de atendimento, listas suspensas, validações e cores do cadastro foram atualizados.',
      ui.ButtonSet.OK
    );
  } catch (erro) {
    ui.alert(
      'Erro ao corrigir o cadastro',
      erro && erro.message
        ? erro.message
        : String(erro),
      ui.ButtonSet.OK
    );

    throw erro;
  }
}


/**
 * Mantém a aba Tipos de Grupo com os valores oficiais.
 */
function sincronizarTiposAtendimentoCadastro_(
  ss
) {
  const nomeAba =
    CONFIG_CADASTRO_PACIENTES
      .ABAS.TIPOS_GRUPO;

  let aba =
    ss.getSheetByName(
      nomeAba
    );

  if (!aba) {
    aba =
      ss.insertSheet(
        nomeAba
      );
  }

  const tipos =
    CONFIG_CADASTRO_PACIENTES
      .TIPOS_ATENDIMENTO;

  const linhasNecessarias =
    tipos.length + 1;

  if (
    aba.getMaxRows() <
    linhasNecessarias
  ) {
    aba.insertRowsAfter(
      aba.getMaxRows(),
      linhasNecessarias -
        aba.getMaxRows()
    );
  }

  /*
   * Remove somente os valores antigos da coluna A.
   * As outras colunas da aba são preservadas.
   */
  const quantidadeLinhas =
    Math.max(
      aba.getMaxRows(),
      linhasNecessarias
    );

  aba.getRange(
    1,
    1,
    quantidadeLinhas,
    1
  ).clearContent();

  aba.getRange(
    1,
    1
  ).setValue(
    'Tipo'
  );

  aba.getRange(
    2,
    1,
    tipos.length,
    1
  ).setValues(
    tipos.map(function(tipo) {
      return [tipo];
    })
  );

  aba.getRange(
    1,
    1,
    1,
    1
  )
    .setFontWeight('bold')
    .setHorizontalAlignment(
      'center'
    )
    .setBackground('#d9ead3');

  aba.autoResizeColumn(1);
  aba.setFrozenRows(1);
}


/**
 * Formata o cabeçalho do Cadastro de Pacientes.
 */
function corrigirCabecalhoCadastro_(
  aba
) {
  const intervalo =
    aba.getRange(
      'A1:X1'
    );

  intervalo
    .setFontWeight('bold')
    .setHorizontalAlignment(
      'center'
    )
    .setVerticalAlignment(
      'middle'
    )
    .setWrap(true)
    .setBackground('#d9ead3');

  aba.setFrozenRows(1);
  aba.setRowHeight(1, 42);
}


/**
 * Reaplica as listas suspensas do cadastro.
 */
function corrigirListasCadastro_(
  ss,
  aba
) {
  const totalLinhas =
    Math.max(
      aba.getMaxRows() - 1,
      1
    );

  const abaHorarios =
    ss.getSheetByName(
      CONFIG_CADASTRO_PACIENTES
        .ABAS.HORARIOS
    );

  const abaTiposGrupo =
    ss.getSheetByName(
      CONFIG_CADASTRO_PACIENTES
        .ABAS.TIPOS_GRUPO
    );

  const abaFisioterapeutas =
    ss.getSheetByName(
      CONFIG_CADASTRO_PACIENTES
        .ABAS.FISIOTERAPEUTAS
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

  const ultimaLinhaHorarios =
    abaHorarios.getLastRow();

  const ultimaLinhaFisioterapeutas =
    abaFisioterapeutas.getLastRow();

  if (ultimaLinhaHorarios < 2) {
    throw new Error(
      'Não existem horários cadastrados.'
    );
  }

  if (
    ultimaLinhaFisioterapeutas < 2
  ) {
    throw new Error(
      'Não existem fisioterapeutas cadastrados.'
    );
  }

  /*
   * Horário das sessões — coluna F.
   */
  const validacaoHorarioSessao =
    SpreadsheetApp
      .newDataValidation()
      .requireValueInRange(
        abaHorarios.getRange(
          2,
          1,
          ultimaLinhaHorarios - 1,
          1
        ),
        true
      )
      .setAllowInvalid(false)
      .setHelpText(
        'Escolha um horário cadastrado na aba Horários.'
      )
      .build();

  aba.getRange(
    2,
    6,
    totalLinhas,
    1
  )
    .clearDataValidations()
    .setDataValidation(
      validacaoHorarioSessao
    )
    .setNumberFormat('HH:mm');

  /*
   * Tipo de atendimento — coluna G.
   * A origem contém exatamente os quatro tipos oficiais.
   */
  const tipos =
    CONFIG_CADASTRO_PACIENTES
      .TIPOS_ATENDIMENTO;

  const validacaoTipoAtendimento =
    SpreadsheetApp
      .newDataValidation()
      .requireValueInRange(
        abaTiposGrupo.getRange(
          2,
          1,
          tipos.length,
          1
        ),
        true
      )
      .setAllowInvalid(false)
      .setHelpText(
        'Escolha um dos tipos de atendimento cadastrados.'
      )
      .build();

  aba.getRange(
    2,
    7,
    totalLinhas,
    1
  )
    .clearDataValidations()
    .setDataValidation(
      validacaoTipoAtendimento
    );

  /*
   * Horário da avaliação — coluna R.
   */
  const validacaoHorarioAvaliacao =
    SpreadsheetApp
      .newDataValidation()
      .requireValueInRange(
        abaHorarios.getRange(
          2,
          1,
          ultimaLinhaHorarios - 1,
          1
        ),
        true
      )
      .setAllowInvalid(false)
      .setHelpText(
        'Escolha o horário da avaliação.'
      )
      .build();

  aba.getRange(
    2,
    18,
    totalLinhas,
    1
  )
    .clearDataValidations()
    .setDataValidation(
      validacaoHorarioAvaliacao
    )
    .setNumberFormat('HH:mm');

  /*
   * Status — coluna U.
   */
  const validacaoStatus =
    SpreadsheetApp
      .newDataValidation()
      .requireValueInList(
        CONFIG_CADASTRO_PACIENTES
          .STATUS,
        true
      )
      .setAllowInvalid(false)
      .setHelpText(
        'Escolha o status atual do paciente.'
      )
      .build();

  aba.getRange(
    2,
    21,
    totalLinhas,
    1
  )
    .clearDataValidations()
    .setDataValidation(
      validacaoStatus
    );

  /*
   * Fisioterapeuta — coluna V.
   */
  const validacaoFisioterapeuta =
    SpreadsheetApp
      .newDataValidation()
      .requireValueInRange(
        abaFisioterapeutas.getRange(
          2,
          2,
          ultimaLinhaFisioterapeutas - 1,
          1
        ),
        true
      )
      .setAllowInvalid(false)
      .setHelpText(
        'Escolha o fisioterapeuta responsável.'
      )
      .build();

  aba.getRange(
    2,
    22,
    totalLinhas,
    1
  )
    .clearDataValidations()
    .setDataValidation(
      validacaoFisioterapeuta
    );
}
/**
 * Reaplica as validações das quantidades de sessões.
 */
function corrigirValidacaoSessoes_(
  aba
) {
  const totalLinhas =
    Math.max(
      aba.getMaxRows() - 1,
      1
    );

  /*
   * Permite zero porque o paciente pode estar
   * apenas com a avaliação agendada.
   */
  const validacaoPrescritas =
    SpreadsheetApp
      .newDataValidation()
      .requireNumberBetween(
        0,
        20
      )
      .setAllowInvalid(false)
      .setHelpText(
        'Informe uma quantidade de 0 a 20 sessões prescritas.'
      )
      .build();

  const validacaoRealizadas =
    SpreadsheetApp
      .newDataValidation()
      .requireNumberBetween(
        0,
        20
      )
      .setAllowInvalid(false)
      .setHelpText(
        'Informe uma quantidade de 0 a 20 sessões realizadas.'
      )
      .build();

  /*
   * N — Sessões Prescritas.
   */
  aba.getRange(
    2,
    14,
    totalLinhas,
    1
  )
    .clearDataValidations()
    .setDataValidation(
      validacaoPrescritas
    )
    .setNumberFormat('0');

  /*
   * O — Sessões Realizadas.
   */
  aba.getRange(
    2,
    15,
    totalLinhas,
    1
  )
    .clearDataValidations()
    .setDataValidation(
      validacaoRealizadas
    )
    .setNumberFormat('0');

  /*
   * P — Sessões Restantes.
   * Campo calculado automaticamente.
   */
  aba.getRange(
    2,
    16,
    totalLinhas,
    1
  )
    .setBackground('#eeeeee')
    .setNote(
      'Campo calculado automaticamente pelo SIGAF.'
    );
}


/**
 * Reaplica as cores dos status.
 */
function corrigirCoresStatusCadastro_(
  aba
) {
  const totalLinhas =
    Math.max(
      aba.getMaxRows() - 1,
      1
    );

  const intervaloStatus =
    aba.getRange(
      2,
      21,
      totalLinhas,
      1
    );

  const regrasExistentes =
    aba.getConditionalFormatRules();

  /*
   * Preserva regras de outras colunas e remove
   * somente as regras que atingem a coluna U.
   */
  const regrasPreservadas =
    regrasExistentes.filter(
      function(regra) {
        return !regra
          .getRanges()
          .some(
            function(intervalo) {
              const primeiraColuna =
                intervalo.getColumn();

              const ultimaColuna =
                primeiraColuna +
                intervalo.getNumColumns() -
                1;

              return (
                primeiraColuna <= 21 &&
                ultimaColuna >= 21
              );
            }
          );
      }
    );

  const regrasStatus = [
    criarRegraCorStatus_(
      intervaloStatus,
      'Avaliação agendada',
      '#fff2cc'
    ),

    criarRegraCorStatus_(
      intervaloStatus,
      'Avaliado – aguardando agendamento',
      '#fce5cd'
    ),

    criarRegraCorStatus_(
      intervaloStatus,
      'Em tratamento',
      '#d9ead3'
    ),

    criarRegraCorStatus_(
      intervaloStatus,
      'Ciclo concluído',
      '#cfe2f3'
    ),

    criarRegraCorStatus_(
      intervaloStatus,
      'Inativo',
      '#d9d9d9'
    )
  ];

  aba.setConditionalFormatRules(
    regrasPreservadas.concat(
      regrasStatus
    )
  );
}


/**
 * Cria uma regra de cor para determinado status.
 */
function criarRegraCorStatus_(
  intervalo,
  texto,
  cor
) {
  return SpreadsheetApp
    .newConditionalFormatRule()
    .whenTextEqualTo(
      texto
    )
    .setBackground(
      cor
    )
    .setRanges([
      intervalo
    ])
    .build();
}
