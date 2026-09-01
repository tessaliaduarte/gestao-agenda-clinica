const CONFIG_EXPORTACAO_WORD_SIGAF = {
  NOME_PASTA: 'Relatórios Mensais SIGAF',
  COR_PRINCIPAL: '#4f81bd',
  COR_SECUNDARIA: '#d9ead3',
  COR_BORDA: '#b7c7d6',
  COR_TEXTO: '#222222',
  COR_TEXTO_SUAVE: '#666666',
  MIME_WORD:
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
};


/**
 * Exporta diretamente o relatório do mês atual para Word.
 */
function exportarRelatorioMesAtualSIGAFWord() {
  validarDependenciasExportacaoWord_();

  const competencia =
    interpretarCompetenciaRelatorio_(
      ''
    );

  executarExportacaoRelatorioWordSIGAF_(
    competencia
  );
}


/**
 * Mantém compatibilidade com o nome antigo.
 */
function exportarRelatorioMensalSIGAFWord() {
  exportarRelatorioOutroMesSIGAFWord();
}


/**
 * Solicita outra competência e exporta para Word.
 */
function exportarRelatorioOutroMesSIGAFWord() {
  const ui =
    SpreadsheetApp.getUi();

  const resposta = ui.prompt(
    'Exportar outro mês para Word',
    'Informe a competência no formato MM/AAAA.\n\n' +
      'Exemplo: 08/2026.',
    ui.ButtonSet.OK_CANCEL
  );

  if (
    resposta.getSelectedButton() !==
      ui.Button.OK
  ) {
    return;
  }

  validarDependenciasExportacaoWord_();

  const competencia =
    interpretarCompetenciaRelatorio_(
      resposta.getResponseText()
    );

  executarExportacaoRelatorioWordSIGAF_(
    competencia
  );
}


/**
 * Motor único da exportação Word.
 */
function executarExportacaoRelatorioWordSIGAF_(
  competencia
) {
  const ui =
    SpreadsheetApp.getUi();

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const lock =
    LockService.getDocumentLock();

  let bloqueioObtido = false;

  try {
    lock.waitLock(30000);
    bloqueioObtido = true;

    validarDependenciasExportacaoWord_();

    gerarRelatorioMensalSIGAF_(
      ss,
      competencia
    );

    const dados =
      obterDadosExportacaoWord_(
        ss,
        competencia
      );

    const resultado =
      criarRelatorioWordSIGAF_(
        ss,
        competencia,
        dados
      );

    SpreadsheetApp.flush();

    mostrarResultadoExportacaoWord_(
      resultado,
      competencia
    );
  } catch (erro) {
    ui.alert(
      'Erro ao exportar relatório',
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
 * Confere se o módulo RelatoriosSIGAF.gs está disponível.
 */
function validarDependenciasExportacaoWord_() {
  const dependencias = [
    {
      nome: 'interpretarCompetenciaRelatorio_',
      funcao:
        typeof interpretarCompetenciaRelatorio_
    },
    {
      nome: 'gerarRelatorioMensalSIGAF_',
      funcao:
        typeof gerarRelatorioMensalSIGAF_
    },
    {
      nome: 'obterAbasRelatorioSIGAF_',
      funcao:
        typeof obterAbasRelatorioSIGAF_
    },
    {
      nome: 'lerCadastroRelatorio_',
      funcao:
        typeof lerCadastroRelatorio_
    },
    {
      nome: 'lerAgendamentosRelatorio_',
      funcao:
        typeof lerAgendamentosRelatorio_
    },
    {
      nome: 'lerHistoricoDesfechosRelatorio_',
      funcao:
        typeof lerHistoricoDesfechosRelatorio_
    }
  ];

  const ausentes =
    dependencias.filter(
      function (item) {
        return item.funcao !== 'function';
      }
    );

  if (ausentes.length > 0) {
    throw new Error(
      'O módulo RelatoriosSIGAF.gs está incompleto. Função ausente: ' +
        ausentes[0].nome
    );
  }
}


/**
 * Reúne os dados mensais usados na planilha e no Word.
 */
function obterDadosExportacaoWord_(
  ss,
  competencia
) {
  const abas =
    obterAbasRelatorioSIGAF_(
      ss
    );

  const cadastroPorId =
    lerCadastroRelatorio_(
      abas.cadastro
    );

  const agendamentos =
    lerAgendamentosRelatorio_(
      abas.agendamentos
    )
      .filter(function (registro) {
        return dataPertenceCompetenciaRelatorio_(
          registro.data,
          competencia
        );
      })
      .sort(
        compararRegistrosRelatorio_
      );

  const historico =
    removerDesfechosDuplicadosRelatorio_(
      lerHistoricoDesfechosRelatorio_(
        abas.historico
      )
        .filter(function (registro) {
          return dataPertenceCompetenciaRelatorio_(
            registro.data,
            competencia
          );
        })
    )
      .sort(
        compararRegistrosRelatorio_
      );

  const resumo =
    calcularResumoRelatorio_(
      agendamentos,
      historico
    );

  const producao =
    calcularProducaoFisioterapeutasRelatorio_(
      agendamentos
    );

  return {
    cadastroPorId: cadastroPorId,
    agendamentos: agendamentos,
    historico: historico,
    resumo: resumo,
    producao: producao
  };
}


/**
 * Cria o Google Docs e o arquivo Word.
 */
function criarRelatorioWordSIGAF_(
  ss,
  competencia,
  dados
) {
  const nomeBase =
    'Relatório Mensal SIGAF - ' +
    formatarCompetenciaRelatorio_(
      competencia
    ).replace('/', '-');

  const documento =
    DocumentApp.create(
      nomeBase
    );

  const corpo =
    documento.getBody();

  corpo.clear();

  configurarPaginaWord_(
    corpo
  );

  montarConteudoWord_(
    corpo,
    competencia,
    dados
  );

  adicionarRodapeWord_(
    documento,
    competencia
  );

  documento.saveAndClose();

  /*
   * Dá ao Google Drive tempo para finalizar
   * a versão do documento antes da exportação.
   */
  Utilities.sleep(1200);

  const pasta =
    obterPastaRelatoriosWord_(
      ss
    );

  const arquivoGoogleDocs =
    DriveApp.getFileById(
      documento.getId()
    );

  moverArquivoParaPastaWord_(
    arquivoGoogleDocs,
    pasta
  );

  const blobWord =
    exportarGoogleDocsComoWord_(
      documento.getId(),
      nomeBase + '.docx'
    );

  const arquivoWord =
    pasta.createFile(
      blobWord
    );

  return {
    nome: nomeBase,
    pasta: pasta,
    googleDocs:
      arquivoGoogleDocs,
    word: arquivoWord
  };
}


/**
 * Define margens e estilo básico da página.
 */
function configurarPaginaWord_(
  corpo
) {
  corpo
    .setMarginTop(45)
    .setMarginBottom(45)
    .setMarginLeft(50)
    .setMarginRight(50);
}


/**
 * Monta todas as seções do documento.
 */
function montarConteudoWord_(
  corpo,
  competencia,
  dados
) {
  adicionarCabecalhoWord_(
    corpo,
    competencia
  );

  adicionarIntroducaoWord_(
    corpo
  );

  adicionarResumoGeralWord_(
    corpo,
    dados.resumo,
    dados.historico
  );

  adicionarProducaoWord_(
    corpo,
    dados.producao
  );

  const avaliacoes =
    dados.agendamentos.filter(
      function (registro) {
        return (
          normalizarTextoRelatorio_(
            registro.evento
          ) === 'avaliacao'
        );
      }
    );

  const sessoes =
    dados.agendamentos.filter(
      function (registro) {
        return (
          normalizarTextoRelatorio_(
            registro.evento
          ) === 'sessao'
        );
      }
    );

  adicionarAvaliacoesWord_(
    corpo,
    avaliacoes
  );

  adicionarSessoesWord_(
    corpo,
    sessoes
  );

  adicionarFaltasAvaliacaoWord_(
    corpo,
    avaliacoes,
    dados.cadastroPorId
  );

  adicionarDesfechosWord_(
    corpo,
    dados.historico,
    dados.cadastroPorId
  );

  adicionarObservacoesFinaisWord_(
    corpo
  );
}


/**
 * Cria o título e a identificação da competência.
 */function adicionarCabecalhoWord_(
  corpo,
  competencia
) {
  const titulo =
    corpo.appendParagraph(
      'RELATÓRIO MENSAL SIGAF'
    );

  titulo
    .setAlignment(
      DocumentApp
        .HorizontalAlignment
        .CENTER
    )
    .setSpacingAfter(4);

  titulo
    .editAsText()
    .setFontFamily('Arial')
    .setFontSize(20)
    .setBold(true)
    .setItalic(false)
    .setForegroundColor(
      CONFIG_EXPORTACAO_WORD_SIGAF
        .COR_PRINCIPAL
    );

  const subtitulo =
    corpo.appendParagraph(
      obterNomeMesRelatorio_(
        competencia.mes
      ).toUpperCase() +
        ' DE ' +
        competencia.ano
    );

  subtitulo
    .setAlignment(
      DocumentApp
        .HorizontalAlignment
        .CENTER
    )
    .setSpacingAfter(4);

  subtitulo
    .editAsText()
    .setFontFamily('Arial')
    .setFontSize(13)
    .setBold(true)
    .setItalic(false)
    .setForegroundColor(
      CONFIG_EXPORTACAO_WORD_SIGAF
        .COR_TEXTO
    );

  const atualizacao =
    corpo.appendParagraph(
      'Competência: ' +
        formatarCompetenciaRelatorio_(
          competencia
        ) +
        ' | Documento gerado em ' +
        Utilities.formatDate(
          new Date(),
          Session.getScriptTimeZone(),
          'dd/MM/yyyy HH:mm'
        )
    );

  atualizacao
    .setAlignment(
      DocumentApp
        .HorizontalAlignment
        .CENTER
    )
    .setSpacingAfter(14);

  atualizacao
    .editAsText()
    .setFontFamily('Arial')
    .setFontSize(9)
    .setBold(false)
    .setItalic(false)
    .setForegroundColor(
      CONFIG_EXPORTACAO_WORD_SIGAF
        .COR_TEXTO_SUAVE
    );
}


/**
 * Explica as principais categorias para a Regulação.
 */
function adicionarIntroducaoWord_(
  corpo
) {
  adicionarTituloSecaoWord_(
    corpo,
    '1. Como interpretar este relatório'
  );

  adicionarParagrafoWord_(
    corpo,
    'Este relatório apresenta a movimentação mensal da fisioterapia. ' +
      'Os indicadores foram calculados a partir dos registros da Agenda, ' +
      'dos Agendamentos e do Histórico de Desfechos do SIGAF.'
  );

  const definicoes = [
    'Avaliações marcadas: avaliações previstas para a competência informada.',
    'Avaliações comparecidas: avaliações em que o resultado registrado foi Compareceu.',
    'Sessões marcadas: sessões de tratamento previstas para o mês. Linhas históricas canceladas por remanejamento não são contadas novamente como uma sessão marcada.',
    'Sessões remanejadas: sessões futuras cuja programação original foi substituída por nova data/horário, preservando a linha antiga apenas para rastreabilidade.',
    'Sessões comparecidas: sessões realizadas com presença registrada.',
    'Sem resultado — passado: atendimento cuja data já passou, mas que continua com status Agendado.',
    'Alta: encerramento do tratamento após conclusão ou decisão clínica.',
    'Encaminhamento para APS: paciente devolvido à Atenção Primária à Saúde para continuidade do cuidado.',
    'Renovação: conclusão do ciclo atual com indicação de planejamento de um novo ciclo.',
    'Alta por abandono: encerramento motivado pelo limite de faltas não justificadas.',
    'Desistência do tratamento: saída registrada por decisão do paciente antes da conclusão prevista.'
  ];

  definicoes.forEach(
    function (texto) {
      const item =
        corpo.appendListItem(
          texto
        );

      item
        .setGlyphType(
          DocumentApp
            .GlyphType
            .BULLET
        )
        .setSpacingAfter(2);

      item
        .editAsText()
        .setFontFamily('Arial')
        .setFontSize(9)
        .setBold(false)
        .setItalic(false)
        .setForegroundColor(
          CONFIG_EXPORTACAO_WORD_SIGAF
            .COR_TEXTO
        );
    }
  );
}


/**
 * Adiciona os indicadores gerais.
 */
function adicionarResumoGeralWord_(
  corpo,
  resumo,
  historico
) {
  adicionarTituloSecaoWord_(
    corpo,
    '2. Resumo geral dos atendimentos'
  );

  const totaisDesfechos =
    calcularTotaisDesfechosRelatorio_(
      historico
    );

  const linhas = [
    [
      'Avaliações marcadas',
      resumo.avaliacoesMarcadas
    ],
    [
      'Avaliações comparecidas',
      resumo.avaliacoesComparecidas
    ],
    [
      'Faltas justificadas em avaliações',
      resumo.avaliacoesFaltaJustificada
    ],
    [
      'Faltas não justificadas em avaliações',
      resumo.avaliacoesFaltaNaoJustificada
    ],
    [
      'Sessões marcadas',
      resumo.sessoesMarcadas
    ],
    [
      'Sessões remanejadas',
      resumo.sessoesRemanejadas || 0
    ],
    [
      'Sessões comparecidas',
      resumo.sessoesComparecidas
    ],
    [
      'Faltas justificadas em sessões',
      resumo.sessoesFaltaJustificada
    ],
    [
      'Faltas não justificadas em sessões',
      resumo.sessoesFaltaNaoJustificada
    ],
    [
      'Cancelamentos pela clínica',
      resumo.avaliacoesCanceladasClinica +
        resumo.sessoesCanceladasClinica
    ],
    [
      'Atendimentos passados sem resultado',
      resumo.avaliacoesSemResultado +
        resumo.sessoesSemResultado
    ],
    [
      'Alta',
      totaisDesfechos.alta || 0
    ],
    [
      'Encaminhamento para APS',
      totaisDesfechos.aps || 0
    ],
    [
      'Renovação',
      totaisDesfechos.renovacao || 0
    ],
    [
      'Alta por abandono',
      totaisDesfechos.abandono || 0
    ],
    [
      'Desistência do tratamento',
      totaisDesfechos.desistencia || 0
    ]
  ];

  adicionarTabelaWord_(
    corpo,
    [
      'Indicador',
      'Quantidade'
    ],
    linhas,
    [
      390,
      110
    ],
    [
      'left',
      'center'
    ]
  );
}


/**
 * Adiciona a produção por fisioterapeuta.
 */
function adicionarProducaoWord_(
  corpo,
  producao
) {
  adicionarTituloSecaoWord_(
    corpo,
    '3. Produção por fisioterapeuta'
  );

  const linhas =
    producao.map(function (item) {
      return [
        item.fisioterapeuta,
        item.avaliacoes,
        item.sessoes,
        item.total
      ];
    });

  adicionarTabelaOuMensagemWord_(
    corpo,
    [
      'Fisioterapeuta',
      'Avaliações',
      'Sessões',
      'Total'
    ],
    linhas,
    [
      250,
      85,
      80,
      75
    ],
    [
      'left',
      'center',
      'center',
      'center'
    ]
  );
}
/**
 * Adiciona a tabela mensal de avaliações.
 */
function adicionarAvaliacoesWord_(
  corpo,
  avaliacoes
) {
  adicionarTituloSecaoWord_(
    corpo,
    '4. Avaliações do mês'
  );

  const linhas =
    avaliacoes.map(
      function (registro) {
        return [
          formatarDataWord_(
            registro.data
          ),
          chaveHorarioRelatorio_(
            registro.horario
          ),
          registro.prontuario,
          registro.paciente,
          registro.status,
          registro.fisioterapeuta
        ];
      }
    );

  adicionarTabelaOuMensagemWord_(
    corpo,
    [
      'Data',
      'Horário',
      'Prontuário',
      'Paciente',
      'Resultado',
      'Fisioterapeuta'
    ],
    linhas,
    [
      58,
      45,
      55,
      135,
      90,
      117
    ],
    [
      'center',
      'center',
      'center',
      'left',
      'center',
      'left'
    ]
  );
}


/**
 * Adiciona a tabela mensal de sessões.
 */
function adicionarSessoesWord_(
  corpo,
  sessoes
) {
  adicionarTituloSecaoWord_(
    corpo,
    '5. Sessões do mês'
  );

  const linhas =
    sessoes.map(
      function (registro) {
        const numeroSessao =
          registro.numeroSessao
            ? (
                registro.numeroSessao +
                '/' +
                registro.totalPrescrito
              )
            : '';

        return [
          formatarDataWord_(
            registro.data
          ),
          chaveHorarioRelatorio_(
            registro.horario
          ),
          registro.prontuario,
          registro.paciente,
          numeroSessao,
          registro.status,
          registro.fisioterapeuta
        ];
      }
    );

  adicionarTabelaOuMensagemWord_(
    corpo,
    [
      'Data',
      'Horário',
      'Prontuário',
      'Paciente',
      'Sessão',
      'Resultado',
      'Fisioterapeuta'
    ],
    linhas,
    [
      55,
      42,
      52,
      115,
      45,
      85,
      106
    ],
    [
      'center',
      'center',
      'center',
      'left',
      'center',
      'center',
      'left'
    ]
  );
}


/**
 * Adiciona as faltas ocorridas em avaliações.
 */
function adicionarFaltasAvaliacaoWord_(
  corpo,
  avaliacoes,
  cadastroPorId
) {
  adicionarTituloSecaoWord_(
    corpo,
    '6. Faltas em avaliações'
  );

  const faltas =
    avaliacoes.filter(
      function (registro) {
        const status =
          normalizarTextoRelatorio_(
            registro.status
          );

        return (
          status ===
            'falta justificada' ||
          status ===
            'falta nao justificada'
        );
      }
    );

  const linhas =
    faltas.map(function (registro) {
      const cadastro =
        cadastroPorId[
          normalizarTextoRelatorio_(
            registro.idPaciente
          )
        ] || {};

      return [
        formatarDataWord_(
          registro.data
        ),
        registro.prontuario,
        registro.paciente,
        cadastro.telefone || '',
        registro.status
      ];
    });

  adicionarTabelaOuMensagemWord_(
    corpo,
    [
      'Data',
      'Prontuário',
      'Paciente',
      'Telefone',
      'Tipo da falta'
    ],
    linhas,
    [
      62,
      62,
      145,
      105,
      126
    ],
    [
      'center',
      'center',
      'left',
      'center',
      'center'
    ]
  );
}


/**
 * Cria uma seção e uma tabela para cada desfecho.
 */
function adicionarDesfechosWord_(
  corpo,
  historico,
  cadastroPorId
) {
  adicionarTituloSecaoWord_(
    corpo,
    '7. Desfechos registrados no mês'
  );

  CONFIG_RELATORIOS_SIGAF
    .DESFECHOS
    .forEach(function (desfecho) {
      adicionarSubtituloWord_(
        corpo,
        desfecho
      );

      const registros =
        historico.filter(
          function (registro) {
            return (
              normalizarTextoRelatorio_(
                registro.desfecho
              ) ===
              normalizarTextoRelatorio_(
                desfecho
              )
            );
          }
        );

      const linhas =
        registros.map(
          function (registro) {
            const cadastro =
              cadastroPorId[
                normalizarTextoRelatorio_(
                  registro.idPaciente
                )
              ] || {};

            return [
              formatarDataWord_(
                registro.data
              ),
              registro.prontuario,
              registro.paciente,
              cadastro.telefone || '',
              registro.cicloNumero || '',
              registro.prescritas,
              registro.realizadas,
              registro.motivo
            ];
          }
        );

      adicionarTabelaOuMensagemWord_(
        corpo,
        [
          'Data',
          'Pront.',
          'Paciente',
          'Telefone',
          'Ciclo',
          'Prescr.',
          'Realiz.',
          'Motivo'
        ],
        linhas,
        [
          55,
          45,
          110,
          85,
          36,
          44,
          44,
          81
        ],
        [
          'center',
          'center',
          'left',
          'center',
          'center',
          'center',
          'center',
          'left'
        ]
      );
    });
}


/**
 * Acrescenta uma nota de encerramento.
 */
function adicionarObservacoesFinaisWord_(
  corpo
) {
  adicionarTituloSecaoWord_(
    corpo,
    '8. Observações'
  );

  adicionarParagrafoWord_(
    corpo,
    'Os dados refletem os registros existentes no SIGAF no momento da geração. ' +
      'Antes do envio, regularize atendimentos passados que ainda estejam como Agendado.'
  );
}
/**
 * Adiciona o rodapé do documento.
 */
function adicionarRodapeWord_(
  documento,
  competencia
) {
  const rodape =
    documento.addFooter();

  const paragrafo =
    rodape.appendParagraph(
      'SIGAF — Relatório mensal ' +
        formatarCompetenciaRelatorio_(
          competencia
        )
    );

  paragrafo.setAlignment(
    DocumentApp
      .HorizontalAlignment
      .CENTER
  );

  paragrafo
    .editAsText()
    .setFontFamily('Arial')
    .setFontSize(8)
    .setBold(false)
    .setItalic(false)
    .setForegroundColor(
      CONFIG_EXPORTACAO_WORD_SIGAF
        .COR_TEXTO_SUAVE
    );
}


/**
 * Adiciona um título principal de seção.
 */
function adicionarTituloSecaoWord_(
  corpo,
  titulo
) {
  const paragrafo =
    corpo.appendParagraph(
      titulo
    );

  paragrafo
    .setSpacingBefore(12)
    .setSpacingAfter(5);

  paragrafo
    .editAsText()
    .setFontFamily('Arial')
    .setFontSize(13)
    .setBold(true)
    .setItalic(false)
    .setForegroundColor(
      CONFIG_EXPORTACAO_WORD_SIGAF
        .COR_PRINCIPAL
    );
}


/**
 * Adiciona um subtítulo.
 */
function adicionarSubtituloWord_(
  corpo,
  titulo
) {
  const paragrafo =
    corpo.appendParagraph(
      titulo
    );

  paragrafo
    .setSpacingBefore(8)
    .setSpacingAfter(3);

  paragrafo
    .editAsText()
    .setFontFamily('Arial')
    .setFontSize(10)
    .setBold(true)
    .setItalic(false)
    .setForegroundColor(
      CONFIG_EXPORTACAO_WORD_SIGAF
        .COR_TEXTO
    );
}


/**
 * Adiciona um parágrafo comum.
 */
function adicionarParagrafoWord_(
  corpo,
  texto
) {
  const paragrafo =
    corpo.appendParagraph(
      texto
    );

  paragrafo
    .setSpacingAfter(6)
    .setLineSpacing(1.08);

  paragrafo
    .editAsText()
    .setFontFamily('Arial')
    .setFontSize(10)
    .setBold(false)
    .setItalic(false)
    .setForegroundColor(
      CONFIG_EXPORTACAO_WORD_SIGAF
        .COR_TEXTO
    );

  return paragrafo;
}


/**
 * Adiciona uma tabela ou uma mensagem quando não houver dados.
 */
function adicionarTabelaOuMensagemWord_(
  corpo,
  cabecalhos,
  linhas,
  larguras,
  alinhamentos
) {
  if (linhas.length === 0) {
    const vazio =
      corpo.appendParagraph(
        'Nenhum registro encontrado nesta competência.'
      );

    vazio.setSpacingAfter(7);

    vazio
      .editAsText()
      .setFontFamily('Arial')
      .setFontSize(9)
      .setBold(false)
      .setItalic(true)
      .setForegroundColor(
        CONFIG_EXPORTACAO_WORD_SIGAF
          .COR_TEXTO_SUAVE
      );

    return null;
  }

  return adicionarTabelaWord_(
    corpo,
    cabecalhos,
    linhas,
    larguras,
    alinhamentos
  );
}


/**
 * Cria e formata uma tabela.
 */
function adicionarTabelaWord_(
  corpo,
  cabecalhos,
  linhas,
  larguras,
  alinhamentos
) {
  const dadosTabela = [
    cabecalhos.map(String)
  ];

  linhas.forEach(function (linha) {
    dadosTabela.push(
      linha.map(function (valor) {
        return String(
          valor === null ||
          valor === undefined
            ? ''
            : valor
        );
      })
    );
  });

  const tabela =
    corpo.appendTable(
      dadosTabela
    );

  tabela
    .setBorderColor(
      CONFIG_EXPORTACAO_WORD_SIGAF
        .COR_BORDA
    )
    .setBorderWidth(0.5);

  for (
    let indiceLinha = 0;
    indiceLinha <
      tabela.getNumRows();
    indiceLinha++
  ) {
    const linha =
      tabela.getRow(
        indiceLinha
      );

    for (
      let indiceColuna = 0;
      indiceColuna <
        linha.getNumCells();
      indiceColuna++
    ) {
      const celula =
        linha.getCell(
          indiceColuna
        );

      if (
        larguras &&
        larguras[indiceColuna]
      ) {
        celula.setWidth(
          larguras[indiceColuna]
        );
      }

      celula.setVerticalAlignment(
        DocumentApp
          .VerticalAlignment
          .CENTER
      );

      const texto =
        celula.editAsText();

      texto
        .setFontFamily('Arial')
        .setFontSize(8)
        .setForegroundColor(
          indiceLinha === 0
            ? '#ffffff'
            : CONFIG_EXPORTACAO_WORD_SIGAF
                .COR_TEXTO
        )
        .setBold(
          indiceLinha === 0
        )
        .setItalic(false);

      if (indiceLinha === 0) {
        celula.setBackgroundColor(
          CONFIG_EXPORTACAO_WORD_SIGAF
            .COR_PRINCIPAL
        );
      } else if (
        indiceLinha % 2 === 0
      ) {
        celula.setBackgroundColor(
          '#f3f7fb'
        );
      }

      const paragrafo =
        celula
          .getChild(0)
          .asParagraph();

      paragrafo
        .setSpacingBefore(1)
        .setSpacingAfter(1)
        .setLineSpacing(1);

      paragrafo.setAlignment(
        obterAlinhamentoWord_(
          alinhamentos &&
          alinhamentos[
            indiceColuna
          ]
        )
      );
    }
  }

  corpo
    .appendParagraph('')
    .setSpacingAfter(2);

  return tabela;
}


/**
 * Converte o nome do alinhamento para DocumentApp.
 */
function obterAlinhamentoWord_(
  alinhamento
) {
  if (alinhamento === 'center') {
    return DocumentApp
      .HorizontalAlignment
      .CENTER;
  }

  if (alinhamento === 'right') {
    return DocumentApp
      .HorizontalAlignment
      .RIGHT;
  }

  return DocumentApp
    .HorizontalAlignment
    .LEFT;
}


/**
 * Formata uma data para o documento.
 */
function formatarDataWord_(
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
 * Localiza ou cria a pasta de relatórios
 * ao lado da planilha.
 */
function obterPastaRelatoriosWord_(
  ss
) {
  const arquivoPlanilha =
    DriveApp.getFileById(
      ss.getId()
    );

  const pais =
    arquivoPlanilha.getParents();

  const pastaPai =
    pais.hasNext()
      ? pais.next()
      : DriveApp.getRootFolder();

  const existentes =
    pastaPai.getFoldersByName(
      CONFIG_EXPORTACAO_WORD_SIGAF
        .NOME_PASTA
    );

  if (existentes.hasNext()) {
    return existentes.next();
  }

  return pastaPai.createFolder(
    CONFIG_EXPORTACAO_WORD_SIGAF
      .NOME_PASTA
  );
}


/**
 * Move o Google Docs para a pasta dos relatórios.
 */
function moverArquivoParaPastaWord_(
  arquivo,
  pasta
) {
  try {
    arquivo.moveTo(
      pasta
    );
  } catch (erro) {
    /*
     * Em alguns Drives compartilhados, mover o arquivo
     * pode ser bloqueado. Isso não impede a exportação.
     */
    console.warn(
      'Não foi possível mover o Google Docs: ' +
        (
          erro && erro.message
            ? erro.message
            : String(erro)
        )
    );
  }
}


/**
 * Exporta um Google Docs como arquivo .docx.
 */
function exportarGoogleDocsComoWord_(
  idDocumento,
  nomeArquivo
) {
  const url =
    'https://www.googleapis.com/drive/v3/files/' +
    encodeURIComponent(
      idDocumento
    ) +
    '/export?mimeType=' +
    encodeURIComponent(
      CONFIG_EXPORTACAO_WORD_SIGAF
        .MIME_WORD
    );

  const resposta =
    UrlFetchApp.fetch(
      url,
      {
        method: 'get',
        headers: {
          Authorization:
            'Bearer ' +
            ScriptApp.getOAuthToken()
        },
        muteHttpExceptions: true
      }
    );

  const codigo =
    resposta.getResponseCode();

  if (codigo !== 200) {
    throw new Error(
      'O Google Docs foi criado, mas não foi possível gerar o arquivo Word. ' +
        'Código da exportação: ' +
        codigo +
        '.'
    );
  }

  return resposta
    .getBlob()
    .setName(
      nomeArquivo
    );
}


/**
 * Mostra links para os arquivos gerados.
 */
function mostrarResultadoExportacaoWord_(
  resultado,
  competencia
) {
  const html =
    HtmlService
      .createHtmlOutput(
        '<div style="' +
          'font-family:Arial,sans-serif;' +
          'padding:16px;' +
          'font-size:14px;' +
        '">' +
          '<h2 style="' +
            'color:#4f81bd;' +
            'margin-top:0;' +
          '">' +
            'Relatório exportado' +
          '</h2>' +
          '<p>Competência: <strong>' +
            formatarCompetenciaRelatorio_(
              competencia
            ) +
          '</strong></p>' +
          '<p>' +
            '<a href="' +
              resultado.word.getUrl() +
            '" target="_blank">' +
              'Abrir arquivo Word (.docx)' +
            '</a>' +
          '</p>' +
          '<p>' +
            '<a href="' +
              resultado.googleDocs.getUrl() +
            '" target="_blank">' +
              'Abrir versão no Google Docs' +
            '</a>' +
          '</p>' +
          '<p>' +
            '<a href="' +
              resultado.pasta.getUrl() +
            '" target="_blank">' +
              'Abrir pasta de relatórios' +
            '</a>' +
          '</p>' +
        '</div>'
      )
      .setWidth(420)
      .setHeight(280);

  SpreadsheetApp
    .getUi()
    .showModalDialog(
      html,
      'Exportação concluída'
    );
}
