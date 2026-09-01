/**
 * Cria os menus sempre que a planilha é aberta.
 */
function onOpen() {
  criarMenusSIGAF2026();
}


/**
 * Mantém compatibilidade com chamadas antigas.
 */
function criarMenuSIGAF() {
  criarMenusSIGAF2026();
}


/**
 * Cria o menu completo do SIGAF.
 */
function criarMenusSIGAF2026() {
  const ui =
    SpreadsheetApp.getUi();

  const menuPacientes = ui
    .createMenu('Pacientes')
    .addItem(
      'Novo paciente',
      'abrirFormularioNovoPaciente'
    )
    .addItem(
      'Editar paciente',
      'abrirFormularioEditarPaciente'
    )
    .addItem(
      'Registrar desfecho',
      'registrarDesfechoPacienteSelecionado'
    )
    .addItem(
      'Planejar renovação',
      'abrirFormularioPlanejamentoRenovacao'
    )
    .addItem(
      'Registrar desistência',
      'abrirFormularioDesistenciaTratamento'
    )
    .addSeparator()
    .addItem(
      'Configurar Cadastro de Pacientes',
      'configurarCadastroPacientes'
    );

  const menuSessoes = ui
    .createMenu('Sessões')
    .addItem(
      'Consultar disponibilidade',
      'consultarDisponibilidadeSessoes'
    )
    .addItem(
      'Confirmar agendamento selecionado',
      'confirmarAgendamentoSelecionado'
    )
    .addSeparator()
    .addItem(
      'Corrigir agendamento',
      'abrirFormularioCorrigirAgendamento'
    );

  const menuAgenda = ui
    .createMenu('Agenda')
    .addItem(
      'Abrir agenda de hoje',
      'abrirAgendaHoje'
    )
    .addItem(
      'Selecionar outra data',
      'selecionarDataAgenda'
    )
    .addSeparator()
    .addItem(
      'Dia anterior',
      'abrirDiaAnteriorAgenda'
    )
    .addItem(
      'Próximo dia',
      'abrirProximoDiaAgenda'
    )
    .addSeparator()
    .addItem(
      'Atualizar agenda',
      'carregarAgendaDiaria'
    );

  const menuRegulacao = ui
    .createMenu('Regulação')
    .addItem(
      'Gerar/atualizar vagas mensais',
      'gerarVagasMensaisRegulacao'
    )
    .addItem(
      'Exportar vagas mensais para Excel',
      'exportarVagasMensaisRegulacaoExcel'
    );

  const menuRelatorios = ui
    .createMenu('Relatórios')
    .addItem(
      'Relatório do mês atual',
      'abrirRelatorioMesAtualSIGAF'
    )
    .addItem(
      'Escolher outro mês',
      'abrirRelatorioOutroMesSIGAF'
    )
    .addSeparator()
    .addItem(
      'Exportar mês atual para Word',
      'exportarRelatorioMesAtualSIGAFWord'
    )
    .addItem(
      'Exportar outro mês para Word',
      'exportarRelatorioOutroMesSIGAFWord'
    );

  ui
    .createMenu('SIGAF')
    .addSubMenu(
      menuPacientes
    )
    .addSubMenu(
      menuSessoes
    )
    .addSubMenu(
      menuAgenda
    )
    .addSubMenu(
      menuRegulacao
    )
    .addSubMenu(
      menuRelatorios
    )
    .addSeparator()
    .addItem(
      'Atualizar menu',
      'criarMenusSIGAF2026'
    )
    .addToUi();
}
