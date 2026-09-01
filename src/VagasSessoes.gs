const CONFIG_VAGAS_SESSOES = {
  ABAS: {
    CADASTRO: 'Cadastro de Pacientes',
    RESULTADOS: 'Vagas para Sessões',
    AGENDAMENTOS: 'Agendamentos',
    HORARIOS: 'Horários',
    FISIOTERAPEUTAS: 'Fisioterapeutas',
    FERIADOS: 'Calendário da Prefeitura',
    BLOQUEIOS: 'Bloqueios'
  },

  DIAS: [
    {
      cadastroColuna: 9,
      numeroSemana: 1,
      nome: 'Segunda-feira',
      abreviado: 'Seg'
    },
    {
      cadastroColuna: 10,
      numeroSemana: 2,
      nome: 'Terça-feira',
      abreviado: 'Ter'
    },
    {
      cadastroColuna: 11,
      numeroSemana: 3,
      nome: 'Quarta-feira',
      abreviado: 'Qua'
    },
    {
      cadastroColuna: 12,
      numeroSemana: 4,
      nome: 'Quinta-feira',
      abreviado: 'Qui'
    },
    {
      cadastroColuna: 13,
      numeroSemana: 5,
      nome: 'Sexta-feira',
      abreviado: 'Sex'
    }
  ],

  STATUS_QUE_OCUPAM_VAGA: [
    'Agendado',
    'Compareceu',
    'Falta Justificada',
    'Falta Não Justificada'
  ],

  LIMITE_OPCOES_PREFERENCIA: 10,
  LIMITE_OPCOES_ALTERNATIVAS: 10,
  PRAZO_MAXIMO_BUSCA_DIAS: 365
};


/**
 * Função principal da consulta de disponibilidade.
 */
function consultarDisponibilidadeSessoes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    const abas = obterAbasVagasSessoes_(ss);

    const identificacao = solicitarPacienteVagasSessoes_(ui);

    if (!identificacao) {
      return;
    }

    const paciente = localizarPacienteVagasSessoes_(
      abas.cadastro,
      identificacao
    );

    if (!paciente) {
      ui.alert(
        'Paciente não encontrado',
        'Não foi encontrado paciente com o ID, prontuário ou nome informado.',
        ui.ButtonSet.OK
      );
      return;
    }

    validarPacienteParaBuscaSessoes_(paciente);

    const dataInicial = solicitarDataInicialVagasSessoes_(ui);

    if (!dataInicial) {
      return;
    }

    prepararAbaVagasSessoes_(abas.resultados);

    ss.toast(
      'O SIGAF está analisando os horários disponíveis.',
      'Consulta de sessões',
      5
    );

    const contexto = construirContextoVagasSessoes_(
      abas,
      paciente,
      dataInicial
    );

    /*
     * Primeira busca: somente os dias preferenciais.
     */
    let opcoesPreferenciais = buscarOpcoesVagasSessoes_(
      contexto,
      [paciente.diasPreferenciais],
      'Preferência atendida',
      CONFIG_VAGAS_SESSOES.LIMITE_OPCOES_PREFERENCIA
    );

    let opcoesAlternativas = [];

    /*
     * Caso não exista opção totalmente compatível com os dias
     * preferenciais, realiza a busca em outras combinações.
     */
    if (opcoesPreferenciais.length === 0) {
      ui.alert(
        'Preferência indisponível',
        'Não há uma sequência completa de vagas nos dias de preferência ' +
          'do paciente.\n\nO SIGAF buscará opções em outros dias.',
        ui.ButtonSet.OK
      );

      const combinacoesAlternativas =
        gerarCombinacoesAlternativasVagasSessoes_(
          paciente.diasPreferenciais
        );

      opcoesAlternativas = buscarOpcoesVagasSessoes_(
        contexto,
        combinacoesAlternativas,
        '',
        CONFIG_VAGAS_SESSOES.LIMITE_OPCOES_ALTERNATIVAS
      );
    }

    const todasOpcoes =
      opcoesPreferenciais.concat(opcoesAlternativas);

    if (todasOpcoes.length === 0) {
      registrarMensagemSemVagasSessoes_(
        abas.resultados,
        paciente
      );

      ui.alert(
        'Nenhuma sequência encontrada',
        'Não foi encontrada uma sequência completa para as ' +
          paciente.sessoesPrescritas +
          ' sessões dentro do período pesquisado.',
        ui.ButtonSet.OK
      );

      ss.setActiveSheet(abas.resultados);
      return;
    }

    gravarOpcoesVagasSessoes_(
      abas.resultados,
      paciente,
      todasOpcoes
    );

    ss.setActiveSheet(abas.resultados);
    abas.resultados.activate();
    abas.resultados.getRange('A2').activate();

    ui.alert(
      'Consulta concluída',
      todasOpcoes.length +
        ' opção(ões) encontrada(s) para ' +
        paciente.nome +
        '.\n\nConfira a aba "Vagas para Sessões".',
      ui.ButtonSet.OK
    );
  } catch (erro) {
    ui.alert(
      'Erro na consulta de sessões',
      erro.message,
      ui.ButtonSet.OK
    );

    throw erro;
  }
}


/**
 * Obtém e valida as abas necessárias.
 */function obterAbasVagasSessoes_(ss) {
  const nomes = CONFIG_VAGAS_SESSOES.ABAS;

  const abas = {
    cadastro: ss.getSheetByName(nomes.CADASTRO),
    resultados: ss.getSheetByName(nomes.RESULTADOS),
    agendamentos: ss.getSheetByName(nomes.AGENDAMENTOS),
    horarios: ss.getSheetByName(nomes.HORARIOS),
    fisioterapeutas: ss.getSheetByName(nomes.FISIOTERAPEUTAS),
    feriados: ss.getSheetByName(nomes.FERIADOS),
    bloqueios: ss.getSheetByName(nomes.BLOQUEIOS)
  };

  Object.keys(abas).forEach(function (chave) {
    if (!abas[chave]) {
      throw new Error(
        'A aba necessária "' +
          nomes[
            {
              cadastro: 'CADASTRO',
              resultados: 'RESULTADOS',
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
 * Solicita ID, prontuário ou nome do paciente.
 */
function solicitarPacienteVagasSessoes_(ui) {
  const resposta = ui.prompt(
    'Consultar vagas para sessões',
    'Informe o ID, o prontuário ou o nome do paciente:',
    ui.ButtonSet.OK_CANCEL
  );

  if (resposta.getSelectedButton() !== ui.Button.OK) {
    return null;
  }

  const valor = String(resposta.getResponseText()).trim();

  if (!valor) {
    ui.alert(
      'Informe algum dado para localizar o paciente.',
      ui.ButtonSet.OK
    );
    return null;
  }

  return valor;
}


/**
 * Solicita a data inicial da pesquisa.
 */
function solicitarDataInicialVagasSessoes_(ui) {
  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0);

  const sugestao = Utilities.formatDate(
    hoje,
    Session.getScriptTimeZone(),
    'dd/MM/yyyy'
  );

  const resposta = ui.prompt(
    'Data de início da pesquisa',
    'A partir de qual data as sessões podem começar?\n\n' +
      'Use o formato DD/MM/AAAA.\n' +
      'Sugestão: ' +
      sugestao,
    ui.ButtonSet.OK_CANCEL
  );

  if (resposta.getSelectedButton() !== ui.Button.OK) {
    return null;
  }

  const texto = String(resposta.getResponseText()).trim();
  const data = converterTextoParaDataVagasSessoes_(texto);

  if (!data) {
    throw new Error(
      'A data informada é inválida. Use o formato DD/MM/AAAA.'
    );
  }

  data.setHours(0, 0, 0, 0);

  return data;
}


/**
 * Localiza o paciente por ID, prontuário ou nome.
 */
function localizarPacienteVagasSessoes_(
  abaCadastro,
  identificacao
) {
  const ultimaLinha = abaCadastro.getLastRow();

  if (ultimaLinha < 2) {
    return null;
  }

  const dados = abaCadastro
    .getRange(2, 1, ultimaLinha - 1, 24)
    .getValues();

  const procurado =
    normalizarTextoVagasSessoes_(identificacao);

  let correspondenciaExata = null;
  const correspondenciasParciais = [];

  dados.forEach(function (linha, indice) {
    const id = normalizarTextoVagasSessoes_(linha[0]);
    const prontuario =
      normalizarTextoVagasSessoes_(linha[1]);
    const nome = normalizarTextoVagasSessoes_(linha[2]);

    if (!id && !prontuario && !nome) {
      return;
    }

    if (
      procurado === id ||
      procurado === prontuario ||
      procurado === nome
    ) {
      correspondenciaExata = montarPacienteVagasSessoes_(
        linha,
        indice + 2
      );
      return;
    }

    if (nome.indexOf(procurado) !== -1) {
      correspondenciasParciais.push(
        montarPacienteVagasSessoes_(linha, indice + 2)
      );
    }
  });

  if (correspondenciaExata) {
    return correspondenciaExata;
  }

  /*
   * Usa busca parcial somente quando existe um único resultado.
   */
  if (correspondenciasParciais.length === 1) {
    return correspondenciasParciais[0];
  }

  if (correspondenciasParciais.length > 1) {
    throw new Error(
      'Foram encontrados vários pacientes com esse nome. ' +
        'Pesquise pelo ID ou prontuário.'
    );
  }

  return null;
}


/**
 * Converte uma linha do cadastro em objeto.
 */
function montarPacienteVagasSessoes_(linha, numeroLinha) {
  const diasPreferenciais = [];

  CONFIG_VAGAS_SESSOES.DIAS.forEach(function (dia, indice) {
    const valor = linha[8 + indice];

    if (
      valor === true ||
      String(valor).toUpperCase() === 'TRUE' ||
      String(valor) === '1'
    ) {
      diasPreferenciais.push(dia.numeroSemana);
    }
  });

  return {
    linha: numeroLinha,
    id: String(linha[0] || '').trim(),
    prontuario: String(linha[1] || '').trim(),
    nome: String(linha[2] || '').trim(),
    telefone: String(linha[4] || '').trim(),
    horarioPreferencial: linha[5],
    tipoGrupo: String(linha[6] || '').trim(),
    limiteGrupo: Number(linha[7]) || 0,
    diasPreferenciais: diasPreferenciais,
    sessoesPrescritas: Number(linha[13]) || 0,
    sessoesRealizadas: Number(linha[14]) || 0,
    sessoesRestantes: Number(linha[15]) || 0,
    status: String(linha[20] || '').trim(),
    fisioterapeutaAtual: String(linha[21] || '').trim()
  };
}


/**
 * Verifica se há dados suficientes para pesquisar as sessões.
 */
function validarPacienteParaBuscaSessoes_(paciente) {
  if (!paciente.id) {
    throw new Error(
      'O paciente ainda não possui um ID cadastrado.'
    );
  }

  if (!paciente.nome) {
    throw new Error(
      'O nome do paciente não está preenchido.'
    );
  }

  /*
   * Somente pacientes já avaliados e aguardando
   * o agendamento das sessões podem entrar na consulta.
   *
   * Esta regra também atende às renovações já planejadas,
   * pois elas recebem o mesmo status antes da criação
   * efetiva do novo ciclo.
   */
  const statusNormalizado =
    normalizarTextoVagasSessoes_(
      paciente.status
    );

  if (
    statusNormalizado !==
      'avaliado - aguardando agendamento'
  ) {
    throw new Error(
      'A consulta de vagas é permitida somente para pacientes com status ' +
      '"Avaliado – aguardando agendamento".\n\n' +
      'Status atual: ' +
      (paciente.status || 'não informado') +
      '.'
    );
  }

  if (!paciente.tipoGrupo) {
    throw new Error(
      'O tipo de atendimento do paciente não foi preenchido.'
    );
  }

  const tipoNormalizado =
    normalizarTextoVagasSessoes_(
      paciente.tipoGrupo
    );

  const tiposPermitidos = [
    'atendimento com maior supervisao',
    'grupo de mmss',
    'grupo de mmii',
    'grupo de coluna'
  ];

  if (
    tiposPermitidos.indexOf(
      tipoNormalizado
    ) === -1
  ) {
    throw new Error(
      'O tipo de atendimento do paciente não é válido.'
    );
  }

  if (paciente.diasPreferenciais.length === 0) {
    throw new Error(
      'Marque pelo menos um dia de preferência no cadastro.'
    );
  }

  if (paciente.sessoesPrescritas < 1) {
    throw new Error(
      'Informe a quantidade de sessões prescritas.'
    );
  }

  if (paciente.sessoesPrescritas > 20) {
    throw new Error(
      'O máximo permitido é de 20 sessões prescritas.'
    );
  }

  if (paciente.sessoesRestantes <= 0) {
    throw new Error(
      'O paciente não possui sessões restantes para agendar.'
    );
  }

  paciente.quantidadeParaAgendar =
    paciente.sessoesRestantes;

  paciente.capacidade =
    tipoNormalizado ===
      'atendimento com maior supervisao'
      ? 2
      : 6;
}
/**
 * Monta dados auxiliares usados na pesquisa.
 */
function construirContextoVagasSessoes_(
  abas,
  paciente,
  dataInicial
) {
  return {
    paciente: paciente,
    dataInicial: dataInicial,
    horarios: lerHorariosPermitidosVagasSessoes_(
      abas.horarios
    ),
    fisioterapeutasPorTurno:
      lerFisioterapeutasPorTurnoVagasSessoes_(
        abas.fisioterapeutas
      ),
    feriados: lerFeriadosVagasSessoes_(abas.feriados),
    bloqueios: lerBloqueiosVagasSessoes_(abas.bloqueios),
    ocupacoes: lerOcupacoesVagasSessoes_(
      abas.agendamentos
    )
  };
}


/**
 * Lê os horários que permitem sessões.
 */
function lerHorariosPermitidosVagasSessoes_(aba) {
  const ultimaLinha = aba.getLastRow();

  if (ultimaLinha < 2) {
    throw new Error(
      'Não existem horários cadastrados na aba "Horários".'
    );
  }

  const dados = aba
    .getRange(2, 1, ultimaLinha - 1, 4)
    .getValues();

  const horarios = [];

  dados.forEach(function (linha) {
    const horario = linha[0];
    const turno = String(linha[1] || '').trim();
    const permiteSessao =
      normalizarTextoVagasSessoes_(linha[3]) === 'sim';

    if (horario && turno && permiteSessao) {
      horarios.push({
        valor: horario,
        chave: chaveHorarioVagasSessoes_(horario),
        exibicao: formatarHorarioVagasSessoes_(horario),
        turno: turno
      });
    }
  });

  if (horarios.length === 0) {
    throw new Error(
      'Nenhum horário está marcado como permitido para sessão.'
    );
  }

  return horarios;
}


/**
 * Relaciona os fisioterapeutas aos turnos.
 */
function lerFisioterapeutasPorTurnoVagasSessoes_(aba) {
  const ultimaLinha = aba.getLastRow();

  const mapa = {};

  if (ultimaLinha < 2) {
    return mapa;
  }

  const dados = aba
    .getRange(2, 1, ultimaLinha - 1, 3)
    .getValues();

  dados.forEach(function (linha) {
    const nome = String(linha[1] || '').trim();
    const turno = normalizarTextoVagasSessoes_(linha[2]);

    if (!nome || !turno) {
      return;
    }

    if (!mapa[turno]) {
      mapa[turno] = [];
    }

    mapa[turno].push(nome);
  });

  return mapa;
}


/**
 * Lê datas sem atendimento.
 */
function lerFeriadosVagasSessoes_(aba) {
  const ultimaLinha = aba.getLastRow();
  const feriados = {};

  if (ultimaLinha < 2) {
    return feriados;
  }

  const dados = aba
    .getRange(2, 1, ultimaLinha - 1, 4)
    .getValues();

  dados.forEach(function (linha) {
    const data = linha[0];
    const atendimento =
      normalizarTextoVagasSessoes_(linha[3]);

    if (
      data instanceof Date &&
      atendimento === 'nao'
    ) {
      feriados[chaveDataVagasSessoes_(data)] = true;
    }
  });

  return feriados;
}


/**
 * Lê os bloqueios ativos.
 */
function lerBloqueiosVagasSessoes_(aba) {
  const ultimaLinha = aba.getLastRow();
  const bloqueios = [];

  if (ultimaLinha < 2) {
    return bloqueios;
  }

  const dados = aba
    .getRange(2, 1, ultimaLinha - 1, 7)
    .getValues();

  dados.forEach(function (linha) {
    const data = linha[0];

    if (!(data instanceof Date)) {
      return;
    }

    const status =
      normalizarTextoVagasSessoes_(linha[6]);

    /*
     * Linha vazia, cancelada ou inativa não bloqueia.
     */
    if (
      status &&
      status !== 'ativo' &&
      status !== 'bloqueado'
    ) {
      return;
    }

    bloqueios.push({
      data: chaveDataVagasSessoes_(data),
      horario: linha[1]
        ? chaveHorarioVagasSessoes_(linha[1])
        : '',
      fisioterapeuta:
        normalizarTextoVagasSessoes_(linha[2]),
      abrangencia:
        normalizarTextoVagasSessoes_(linha[3])
    });
  });

  return bloqueios;
}


/**
 * Lê a ocupação atual da aba Agendamentos.
 */
function lerOcupacoesVagasSessoes_(aba) {
  const ultimaLinha = aba.getLastRow();
  const ocupacoes = {};

  if (ultimaLinha < 2) {
    return ocupacoes;
  }

  const dados = aba
    .getRange(2, 1, ultimaLinha - 1, 22)
    .getValues();

  const statusAceitos =
    CONFIG_VAGAS_SESSOES.STATUS_QUE_OCUPAM_VAGA.map(
      normalizarTextoVagasSessoes_
    );

  dados.forEach(function (linha) {
    const data = linha[6];
    const horario = linha[8];
    const fisioterapeuta = String(linha[9] || '').trim();
    const tipoGrupo = String(linha[10] || '').trim();
    const evento =
      normalizarTextoVagasSessoes_(linha[11]);
    const status =
      normalizarTextoVagasSessoes_(linha[15]);

    if (
      !(data instanceof Date) ||
      !horario ||
      !fisioterapeuta ||
      evento !== 'sessao' ||
      statusAceitos.indexOf(status) === -1
    ) {
      return;
    }

    const chave = montarChaveOcupacaoVagasSessoes_(
      data,
      horario,
      fisioterapeuta
    );

    if (!ocupacoes[chave]) {
      ocupacoes[chave] = {
        quantidade: 0,
        tipos: {}
      };
    }

    ocupacoes[chave].quantidade++;

    const tipoNormalizado =
      normalizarTextoVagasSessoes_(tipoGrupo);

    if (tipoNormalizado) {
      ocupacoes[chave].tipos[tipoNormalizado] = true;
    }
  });

  return ocupacoes;
}


/**
 * Busca opções para as combinações de dias recebidas.
 */function buscarOpcoesVagasSessoes_(
  contexto,
  combinacoesDias,
  situacaoFixa,
  limite
) {
  const opcoes = [];
  const chavesUsadas = {};

  for (
    let c = 0;
    c < combinacoesDias.length && opcoes.length < limite;
    c++
  ) {
    const dias = combinacoesDias[c];

    for (
      let h = 0;
      h < contexto.horarios.length &&
      opcoes.length < limite;
      h++
    ) {
      const horario = contexto.horarios[h];
      const turnoNormalizado =
        normalizarTextoVagasSessoes_(horario.turno);

      const fisioterapeutas =
        contexto.fisioterapeutasPorTurno[
          turnoNormalizado
        ] || [];

      for (
        let f = 0;
        f < fisioterapeutas.length &&
        opcoes.length < limite;
        f++
      ) {
        const fisioterapeuta = fisioterapeutas[f];

        const sequencia = simularSequenciaVagasSessoes_(
          contexto,
          dias,
          horario,
          fisioterapeuta
        );

        if (!sequencia.valida) {
          continue;
        }

        const chaveOpcao =
          dias.join('-') +
          '|' +
          horario.chave +
          '|' +
          normalizarTextoVagasSessoes_(fisioterapeuta);

        if (chavesUsadas[chaveOpcao]) {
          continue;
        }

        chavesUsadas[chaveOpcao] = true;

        const situacao =
          situacaoFixa ||
          classificarOpcaoVagasSessoes_(
            dias,
            contexto.paciente.diasPreferenciais
          );

        opcoes.push({
          dias: dias,
          horario: horario,
          fisioterapeuta: fisioterapeuta,
          sequencia: sequencia.datas,
          ocupacaoMaxima: sequencia.ocupacaoMaxima,
          vagasMinimas: sequencia.vagasMinimas,
          dataInicio: sequencia.datas[0],
          dataTermino:
            sequencia.datas[sequencia.datas.length - 1],
          situacao: situacao
        });
      }
    }
  }

  return opcoes;
}


/**
 * Simula todas as datas necessárias para uma opção.
 */
function simularSequenciaVagasSessoes_(
  contexto,
  diasSemana,
  horario,
  fisioterapeuta
) {
  const datas = [];
  let ocupacaoMaxima = 0;
  let vagasMinimas = contexto.paciente.capacidade;

  const data = new Date(contexto.dataInicial);
  data.setHours(0, 0, 0, 0);

  let diasAnalisados = 0;

  while (
    datas.length <
      contexto.paciente.quantidadeParaAgendar &&
    diasAnalisados <
      CONFIG_VAGAS_SESSOES.PRAZO_MAXIMO_BUSCA_DIAS
  ) {
    diasAnalisados++;

    const diaSemana = data.getDay();

    if (
      diasSemana.indexOf(diaSemana) !== -1 &&
      !contexto.feriados[chaveDataVagasSessoes_(data)] &&
      !estaBloqueadoVagasSessoes_(
        contexto.bloqueios,
        data,
        horario.valor,
        fisioterapeuta
      )
    ) {
      const disponibilidade =
        verificarCapacidadeDataVagasSessoes_(
          contexto,
          data,
          horario.valor,
          fisioterapeuta
        );

      if (!disponibilidade.disponivel) {
        /*
         * A sequência precisa preservar o mesmo horário e
         * padrão semanal. Uma data sem vaga invalida a opção.
         */
        return {
          valida: false,
          datas: []
        };
      }

      datas.push(new Date(data));

      ocupacaoMaxima = Math.max(
        ocupacaoMaxima,
        disponibilidade.ocupacao
      );

      vagasMinimas = Math.min(
        vagasMinimas,
        disponibilidade.vagasDisponiveis
      );
    }

    data.setDate(data.getDate() + 1);
  }

  if (
    datas.length <
      contexto.paciente.quantidadeParaAgendar
  ) {
    return {
      valida: false,
      datas: []
    };
  }

  return {
    valida: true,
    datas: datas,
    ocupacaoMaxima: ocupacaoMaxima,
    vagasMinimas: vagasMinimas
  };
}


/**
 * Verifica se há capacidade e se o grupo é compatível.
 */
function verificarCapacidadeDataVagasSessoes_(
  contexto,
  data,
  horario,
  fisioterapeuta
) {
  const chave = montarChaveOcupacaoVagasSessoes_(
    data,
    horario,
    fisioterapeuta
  );

  const ocupacao = contexto.ocupacoes[chave] || {
    quantidade: 0,
    tipos: {}
  };

  const tipoPaciente =
    normalizarTextoVagasSessoes_(
      contexto.paciente.tipoGrupo
    );

  const tiposExistentes = Object.keys(
    ocupacao.tipos
  );

  /*
   * Se o horário já possui outro tipo de grupo,
   * não é permitido misturar.
   *
   * Se o horário já estiver inconsistente e possuir
   * mais de um tipo, também não será oferecido.
   */
  if (
    tiposExistentes.length > 0 &&
    (
      tiposExistentes.length !== 1 ||
      tiposExistentes[0] !== tipoPaciente
    )
  ) {
    return {
      disponivel: false,
      ocupacao: ocupacao.quantidade,
      vagasDisponiveis: 0
    };
  }

  const vagasDisponiveis =
    contexto.paciente.capacidade -
    ocupacao.quantidade;

  return {
    disponivel: vagasDisponiveis >= 1,
    ocupacao: ocupacao.quantidade,
    vagasDisponiveis: vagasDisponiveis
  };
}


/**
 * Verifica bloqueio por data, horário e profissional.
 */
function estaBloqueadoVagasSessoes_(
  bloqueios,
  data,
  horario,
  fisioterapeuta
) {
  const chaveData =
    chaveDataVagasSessoes_(data);

  const chaveHorario =
    chaveHorarioVagasSessoes_(horario);

  const chaveFisioterapeuta =
    normalizarTextoVagasSessoes_(
      fisioterapeuta
    );

  return bloqueios.some(function (bloqueio) {
    if (bloqueio.data !== chaveData) {
      return false;
    }

    const bloqueiaHorario =
      !bloqueio.horario ||
      bloqueio.horario === chaveHorario ||
      bloqueio.abrangencia === 'dia inteiro' ||
      bloqueio.abrangencia === 'turno inteiro';

    const bloqueiaFisioterapeuta =
      !bloqueio.fisioterapeuta ||
      bloqueio.fisioterapeuta ===
        chaveFisioterapeuta ||
      bloqueio.fisioterapeuta === 'todos';

    return (
      bloqueiaHorario &&
      bloqueiaFisioterapeuta
    );
  });
}


/**
 * Gera combinações alternativas com a mesma quantidade
 * de dias semanais escolhida pelo paciente.
 */
function gerarCombinacoesAlternativasVagasSessoes_(
  diasPreferenciais
) {
  const todosDias = [1, 2, 3, 4, 5];

  const quantidadeDias = Math.max(
    1,
    diasPreferenciais.length
  );

  const combinacoes =
    gerarCombinacoesVagasSessoes_(
      todosDias,
      quantidadeDias
    );

  const chavePreferencia =
    diasPreferenciais
      .slice()
      .sort()
      .join('-');

  return combinacoes
    .filter(function (dias) {
      return (
        dias
          .slice()
          .sort()
          .join('-') !==
        chavePreferencia
      );
    })
    .sort(function (a, b) {
      const pontosA =
        calcularCoincidenciasDiasVagasSessoes_(
          a,
          diasPreferenciais
        );

      const pontosB =
        calcularCoincidenciasDiasVagasSessoes_(
          b,
          diasPreferenciais
        );

      return pontosB - pontosA;
    });
}


/**
 * Gera combinações matemáticas de dias.
 */
function gerarCombinacoesVagasSessoes_(
  elementos,
  quantidade
) {
  const resultado = [];

  function combinar(inicio, atual) {
    if (atual.length === quantidade) {
      resultado.push(atual.slice());
      return;
    }

    for (
      let i = inicio;
      i < elementos.length;
      i++
    ) {
      atual.push(elementos[i]);
      combinar(i + 1, atual);
      atual.pop();
    }
  }

  combinar(0, []);

  return resultado;
}


/**
 * Classifica a proximidade com os dias preferenciais.
 */
function classificarOpcaoVagasSessoes_(
  diasOpcao,
  diasPreferenciais
) {
  const coincidencias =
    calcularCoincidenciasDiasVagasSessoes_(
      diasOpcao,
      diasPreferenciais
    );

  if (
    coincidencias === diasPreferenciais.length &&
    diasOpcao.length === diasPreferenciais.length
  ) {
    return 'Preferência atendida';
  }

  if (coincidencias > 0) {
    return 'Preferência parcialmente atendida';
  }

  return 'Dias alternativos';
}


function calcularCoincidenciasDiasVagasSessoes_(
  diasA,
  diasB
) {
  return diasA.filter(function (dia) {
    return diasB.indexOf(dia) !== -1;
  }).length;
}


/**
 * Limpa e configura a aba de resultados.
 */function prepararAbaVagasSessoes_(aba) {
  const cabecalhos = [[
    'Opção',
    'ID do Paciente',
    'Prontuário',
    'Paciente',
    'Dias da semana',
    'Horário',
    'Fisioterapeuta',
    'Tipo do Grupo',
    'Ocupação máxima',
    'Capacidade',
    'Vagas mínimas disponíveis',
    'Data de início',
    'Data prevista de término',
    'Situação'
  ]];

  aba.clearContents();
  aba.clearFormats();

  /*
   * clearDataValidations não existe diretamente para Sheet.
   * Por isso, a limpeza é aplicada ao intervalo utilizado.
   */
  aba.getDataRange().clearDataValidations();

  aba.setConditionalFormatRules([]);

  aba.getRange(1, 1, 1, 14)
    .setValues(cabecalhos)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true)
    .setBackground('#d9ead3');

  aba.setFrozenRows(1);
  aba.setRowHeight(1, 42);

  aba.setColumnWidth(1, 70);
  aba.setColumnWidth(2, 115);
  aba.setColumnWidth(3, 110);
  aba.setColumnWidth(4, 220);
  aba.setColumnWidth(5, 190);
  aba.setColumnWidth(6, 90);
  aba.setColumnWidth(7, 190);
  aba.setColumnWidth(8, 220);
  aba.setColumnWidth(9, 120);
  aba.setColumnWidth(10, 95);
  aba.setColumnWidth(11, 135);
  aba.setColumnWidth(12, 110);
  aba.setColumnWidth(13, 145);
  aba.setColumnWidth(14, 220);
}


/**
 * Grava as opções encontradas.
 */
function gravarOpcoesVagasSessoes_(
  aba,
  paciente,
  opcoes
) {
  const linhas = opcoes.map(function (opcao, indice) {
    return [
      indice + 1,
      paciente.id,
      paciente.prontuario,
      paciente.nome,
      formatarDiasVagasSessoes_(opcao.dias),
      opcao.horario.valor,
      opcao.fisioterapeuta,
      paciente.tipoGrupo,
      opcao.ocupacaoMaxima,
      paciente.capacidade,
      opcao.vagasMinimas,
      opcao.dataInicio,
      opcao.dataTermino,
      opcao.situacao
    ];
  });

  aba.getRange(2, 1, linhas.length, 14)
    .setValues(linhas)
    .setVerticalAlignment('middle');

  aba.getRange(2, 1, linhas.length, 3)
    .setHorizontalAlignment('center');

  aba.getRange(2, 6, linhas.length, 1)
    .setNumberFormat('HH:mm');

  aba.getRange(2, 9, linhas.length, 3)
    .setNumberFormat('0')
    .setHorizontalAlignment('center');

  aba.getRange(2, 12, linhas.length, 2)
    .setNumberFormat('dd/MM/yyyy')
    .setHorizontalAlignment('center');

  aplicarCoresOpcoesVagasSessoes_(
    aba,
    linhas.length
  );
}


/**
 * Aplica cores conforme a classificação.
 */
function aplicarCoresOpcoesVagasSessoes_(
  aba,
  quantidadeLinhas
) {
  const intervaloSituacao = aba.getRange(
    2,
    14,
    quantidadeLinhas,
    1
  );

  const regras = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Preferência atendida')
      .setBackground('#d9ead3')
      .setRanges([intervaloSituacao])
      .build(),

    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(
        'Preferência parcialmente atendida'
      )
      .setBackground('#fff2cc')
      .setRanges([intervaloSituacao])
      .build(),

    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Dias alternativos')
      .setBackground('#fce5cd')
      .setRanges([intervaloSituacao])
      .build()
  ];

  aba.setConditionalFormatRules(regras);
}


/**
 * Mostra uma mensagem na própria aba quando nada é encontrado.
 */
function registrarMensagemSemVagasSessoes_(
  aba,
  paciente
) {
  aba.getRange('A2:N2').merge();

  aba.getRange('A2')
    .setValue(
      'Nenhuma sequência completa encontrada para ' +
        paciente.nome +
        '.'
    )
    .setHorizontalAlignment('center')
    .setFontWeight('bold')
    .setBackground('#f4cccc');

  aba.setRowHeight(2, 35);
}


/**
 * Formata os dias da semana.
 */
function formatarDiasVagasSessoes_(dias) {
  return dias
    .map(function (numeroDia) {
      const configuracao =
        CONFIG_VAGAS_SESSOES.DIAS.find(
          function (dia) {
            return dia.numeroSemana === numeroDia;
          }
        );

      return configuracao
        ? configuracao.abreviado
        : '';
    })
    .filter(String)
    .join(' / ');
}


function montarChaveOcupacaoVagasSessoes_(
  data,
  horario,
  fisioterapeuta
) {
  return [
    chaveDataVagasSessoes_(data),
    chaveHorarioVagasSessoes_(horario),
    normalizarTextoVagasSessoes_(fisioterapeuta)
  ].join('|');
}


function chaveDataVagasSessoes_(data) {
  return Utilities.formatDate(
    new Date(data),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );
}


function chaveHorarioVagasSessoes_(valor) {
  if (valor instanceof Date) {
    return Utilities.formatDate(
      valor,
      Session.getScriptTimeZone(),
      'HH:mm'
    );
  }

  const texto = String(valor || '').trim();

  const correspondencia = texto.match(
    /(\d{1,2}):(\d{2})/
  );

  if (!correspondencia) {
    return texto;
  }

  return (
    String(
      Number(correspondencia[1])
    ).padStart(2, '0') +
    ':' +
    correspondencia[2]
  );
}


function formatarHorarioVagasSessoes_(valor) {
  return chaveHorarioVagasSessoes_(valor);
}


function converterTextoParaDataVagasSessoes_(texto) {
  const resultado = String(texto || '').match(
    /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/
  );

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

  return data;
}


function normalizarTextoVagasSessoes_(valor) {
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
