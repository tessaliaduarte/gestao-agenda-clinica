const CONFIG_EXPORTAR_VAGAS_REGULACAO = {
  ABA_ORIGEM: 'Vagas para Regulação',

  NOME_PASTA:
    'SIGAF - Arquivos para Regulação',

  COLUNAS: {
    DATA: 1,
    DIA: 2,
    HORARIO: 3,
    FISIOTERAPEUTA: 4,
    TURNO: 5,
    SITUACAO: 6
  },

  QUANTIDADE_COLUNAS_ORIGEM: 6,
  QUANTIDADE_COLUNAS_EXPORTACAO: 5
};


/**
 * Atualiza as vagas da competência e exporta somente
 * as vagas disponíveis para um arquivo Excel.
 */
function exportarVagasMensaisRegulacaoExcel() {
  const ui =
    SpreadsheetApp.getUi();

  const resposta = ui.prompt(
    'Exportar vagas para a Regulação',
    'Informe a competência no formato MM/AAAA.\n\n' +
      'Se deixar o campo vazio, será usado o próximo mês.',
    ui.ButtonSet.OK_CANCEL
  );

  if (
    resposta.getSelectedButton() !==
      ui.Button.OK
  ) {
    return;
  }

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  try {
    validarDependenciasExportacaoVagas_(
      ss
    );

    const competencia =
      interpretarCompetenciaVagasRegulacao_(
        resposta.getResponseText()
      );

    const resultado =
      prepararExportacaoVagasRegulacao_(
        ss,
        competencia
      );

    exibirResultadoExportacaoVagas_(
      resultado
    );
  } catch (erro) {
    ui.alert(
      'Erro ao exportar vagas',
      erro && erro.message
        ? erro.message
        : String(erro),
      ui.ButtonSet.OK
    );

    throw erro;
  }
}


/**
 * Executa a atualização e a exportação sob bloqueio,
 * evitando duas exportações simultâneas.
 */
function prepararExportacaoVagasRegulacao_(
  ss,
  competencia
) {
  const lock =
    LockService.getDocumentLock();

  let bloqueioObtido = false;

  try {
    lock.waitLock(30000);
    bloqueioObtido = true;

    /*
     * Atualiza a aba antes de exportar.
     */
    gerarVagasCompetenciaRegulacao_(
      ss,
      competencia
    );

    SpreadsheetApp.flush();

    const vagas =
      lerVagasDisponiveisParaExportacao_(
        ss,
        competencia
      );

    if (vagas.length === 0) {
      throw new Error(
        'Não existem vagas disponíveis para a competência ' +
          formatarCompetenciaVagasRegulacao_(
            competencia
          ) +
          '.'
      );
    }

    const arquivo =
      criarArquivoExcelVagasRegulacao_(
        vagas,
        competencia
      );

    return {
      arquivo: arquivo,
      quantidade: vagas.length,
      competencia: competencia
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
 * Lê somente as vagas disponíveis da competência.
 */
function lerVagasDisponiveisParaExportacao_(
  ss,
  competencia
) {
  const aba =
    ss.getSheetByName(
      CONFIG_EXPORTAR_VAGAS_REGULACAO
        .ABA_ORIGEM
    );

  const ultimaLinha =
    aba.getLastRow();

  if (ultimaLinha < 2) {
    return [];
  }

  const dados =
    aba
      .getRange(
        2,
        1,
        ultimaLinha - 1,
        CONFIG_EXPORTAR_VAGAS_REGULACAO
          .QUANTIDADE_COLUNAS_ORIGEM
      )
      .getValues();

  const c =
    CONFIG_EXPORTAR_VAGAS_REGULACAO
      .COLUNAS;

  return dados
    .map(function(linha) {
      return {
        data:
          linha[c.DATA - 1],

        dia: String(
          linha[c.DIA - 1] || ''
        ).trim(),

        horario:
          linha[c.HORARIO - 1],

        fisioterapeuta: String(
          linha[
            c.FISIOTERAPEUTA - 1
          ] || ''
        ).trim(),

        turno: String(
          linha[c.TURNO - 1] || ''
        ).trim(),

        situacao: String(
          linha[c.SITUACAO - 1] || ''
        ).trim()
      };
    })
    .filter(function(vaga) {
      return (
        vaga.data instanceof Date &&

        dataPertenceCompetenciaVagasRegulacao_(
          vaga.data,
          competencia
        ) &&

        normalizarTextoVagasRegulacao_(
          vaga.situacao
        ) === 'disponivel' &&

        Boolean(
          chaveHorarioVagasRegulacao_(
            vaga.horario
          )
        ) &&

        Boolean(
          vaga.fisioterapeuta
        )
      );
    })
    .sort(
      compararVagasRegulacao_
    );
}


/**
 * Cria uma planilha temporária, converte para XLSX
 * e salva o arquivo final na pasta da Regulação.
 */
function criarArquivoExcelVagasRegulacao_(
  vagas,
  competencia
) {  const nomeBase =
    montarNomeArquivoVagasRegulacao_(
      competencia
    );

  const planilhaTemporaria =
    SpreadsheetApp.create(
      nomeBase +
        ' - temporário'
    );

  const idTemporario =
    planilhaTemporaria.getId();

  try {
    montarPlanilhaExportacaoVagas_(
      planilhaTemporaria,
      vagas,
      competencia
    );

    SpreadsheetApp.flush();

    /*
     * Aguarda a consolidação da formatação.
     */
    Utilities.sleep(800);

    const blob =
      exportarPlanilhaTemporariaComoXlsx_(
        idTemporario,
        nomeBase + '.xlsx'
      );

    const pasta =
      obterPastaExportacaoVagas_();

    return pasta.createFile(
      blob
    );
  } finally {
    /*
     * A planilha temporária nunca permanece no Drive.
     */
    try {
      DriveApp
        .getFileById(
          idTemporario
        )
        .setTrashed(true);
    } catch (erroLixeira) {
      console.error(
        erroLixeira
      );
    }
  }
}


/**
 * Monta o conteúdo visual da planilha exportada.
 */
function montarPlanilhaExportacaoVagas_(
  planilha,
  vagas,
  competencia
) {
  const aba =
    planilha.getSheets()[0];

  const nomeMes =
    obterNomeMesExportacaoVagas_(
      competencia.mes
    );

  aba.setName(
    'Vagas ' +
      nomeMes +
      ' ' +
      competencia.ano
  );

  aba
    .getRange('A1:E1')
    .merge()
    .setValue(
      'VAGAS DE AVALIAÇÃO EM FISIOTERAPIA'
    )
    .setBackground('#4f81bd')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(16)
    .setHorizontalAlignment(
      'center'
    )
    .setVerticalAlignment(
      'middle'
    );

  aba.setRowHeight(1, 34);

  aba
    .getRange('A2:E2')
    .merge()
    .setValue(
      'Competência: ' +
        nomeMes +
        ' de ' +
        competencia.ano
    )
    .setFontWeight('bold')
    .setHorizontalAlignment(
      'center'
    );

  aba
    .getRange('A3:E3')
    .merge()
    .setValue(
      'Arquivo gerado pelo SIGAF em ' +
        Utilities.formatDate(
          new Date(),
          Session.getScriptTimeZone(),
          'dd/MM/yyyy HH:mm'
        )
    )
    .setFontColor('#666666')
    .setFontSize(10)
    .setHorizontalAlignment(
      'center'
    );

  const cabecalhos = [[
    'Data',
    'Dia da semana',
    'Horário',
    'Fisioterapeuta',
    'Turno'
  ]];

  aba
    .getRange(
      5,
      1,
      1,
      CONFIG_EXPORTAR_VAGAS_REGULACAO
        .QUANTIDADE_COLUNAS_EXPORTACAO
    )
    .setValues(cabecalhos)
    .setBackground('#d9ead3')
    .setFontWeight('bold')
    .setHorizontalAlignment(
      'center'
    )
    .setVerticalAlignment(
      'middle'
    )
    .setWrap(true);

  const valores =
    vagas.map(function(vaga) {
      return [
        new Date(vaga.data),

        vaga.dia ||
          obterNomeDiaVagasRegulacao_(
            vaga.data
          ),

        vaga.horario,
        vaga.fisioterapeuta,
        vaga.turno
      ];
    });

  aba
    .getRange(
      6,
      1,
      valores.length,
      CONFIG_EXPORTAR_VAGAS_REGULACAO
        .QUANTIDADE_COLUNAS_EXPORTACAO
    )
    .setValues(valores)
    .setVerticalAlignment(
      'middle'
    );

  aba
    .getRange(
      6,
      1,
      valores.length,
      1
    )
    .setNumberFormat(
      'dd/MM/yyyy'
    )
    .setHorizontalAlignment(
      'center'
    );

  aba
    .getRange(
      6,
      3,
      valores.length,
      1
    )
    .setNumberFormat(
      'HH:mm'
    )
    .setHorizontalAlignment(
      'center'
    );

  aba
    .getRange(
      6,
      2,
      valores.length,
      1
    )
    .setHorizontalAlignment(
      'center'
    );

  aba
    .getRange(
      6,
      5,
      valores.length,
      1
    )
    .setHorizontalAlignment(
      'center'
    );

  const ultimaLinha =
    5 + valores.length;

  aba
    .getRange(
      5,
      1,
      valores.length + 1,
      CONFIG_EXPORTAR_VAGAS_REGULACAO
        .QUANTIDADE_COLUNAS_EXPORTACAO
    )
    .setBorder(
      true,
      true,
      true,
      true,
      true,
      true,
      '#b7c5d3',
      SpreadsheetApp.BorderStyle.SOLID
    );

  aba
    .getRange(
      6,
      1,
      valores.length,
      CONFIG_EXPORTAR_VAGAS_REGULACAO
        .QUANTIDADE_COLUNAS_EXPORTACAO
    )
    .applyRowBanding(
      SpreadsheetApp
        .BandingTheme
        .LIGHT_GREY,
      false,
      false
    );

  aba
    .getRange(
      ultimaLinha + 2,
      1,
      1,
      5
    )
    .merge()
    .setValue(
      'Total de vagas disponíveis: ' +
        valores.length
    )
    .setFontWeight('bold');

  aba
    .getRange(
      ultimaLinha + 3,
      1,
      1,
      5
    )
    .merge()
    .setValue(
      'Esta planilha contém somente vagas disponíveis no momento da geração.'
    )
    .setFontColor('#666666')
    .setFontSize(10);

  aba.setFrozenRows(5);

  aba.setColumnWidth(1, 105);
  aba.setColumnWidth(2, 135);
  aba.setColumnWidth(3, 90);
  aba.setColumnWidth(4, 210);
  aba.setColumnWidth(5, 100);

  aba
    .getRange(
      1,
      1,
      ultimaLinha + 3,
      5
    )
    .setFontFamily('Arial')
    .setWrap(true);

  aba
    .getRange(
      5,
      1,
      valores.length + 1,
      5
    )
    .createFilter();
}
/**
 * Exporta a planilha temporária para o formato XLSX.
 */
function exportarPlanilhaTemporariaComoXlsx_(
  idPlanilha,
  nomeArquivo
) {
  const url =
    'https://docs.google.com/spreadsheets/d/' +
    encodeURIComponent(
      idPlanilha
    ) +
    '/export?format=xlsx';

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

  if (
    codigo < 200 ||
    codigo >= 300
  ) {
    throw new Error(
      'O Google Drive não conseguiu gerar o arquivo Excel. Código: ' +
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
 * Localiza ou cria a pasta das vagas exportadas.
 */
function obterPastaExportacaoVagas_() {
  const nome =
    CONFIG_EXPORTAR_VAGAS_REGULACAO
      .NOME_PASTA;

  const pastas =
    DriveApp.getFoldersByName(
      nome
    );

  if (pastas.hasNext()) {
    return pastas.next();
  }

  return DriveApp.createFolder(
    nome
  );
}


/**
 * Monta o nome do arquivo Excel.
 */
function montarNomeArquivoVagasRegulacao_(
  competencia
) {
  const mes =
    obterNomeMesExportacaoVagas_(
      competencia.mes
    );

  return (
    'Vagas Fisioterapia Regulação - ' +
    mes +
    ' ' +
    competencia.ano
  );
}


/**
 * Retorna o nome do mês.
 */
function obterNomeMesExportacaoVagas_(
  indiceMes
) {
  const meses = [
    'Janeiro',
    'Fevereiro',
    'Março',
    'Abril',
    'Maio',
    'Junho',
    'Julho',
    'Agosto',
    'Setembro',
    'Outubro',
    'Novembro',
    'Dezembro'
  ];

  return meses[indiceMes];
}


/**
 * Exibe o resultado da exportação.
 */
function exibirResultadoExportacaoVagas_(
  resultado
) {
  const nome =
    resultado.arquivo.getName();

  const url =
    resultado.arquivo.getUrl();

  const competencia =
    formatarCompetenciaVagasRegulacao_(
      resultado.competencia
    );

  const html =
    HtmlService
      .createHtmlOutput(
        '<div style="font-family:Arial,sans-serif;padding:20px;color:#243447;">' +
          '<h2 style="margin-top:0;color:#315f96;">Exportação concluída</h2>' +
          '<p><strong>Competência:</strong> ' +
          escaparHtmlExportacaoVagas_(
            competencia
          ) +
          '</p>' +
          '<p><strong>Vagas disponíveis:</strong> ' +
          resultado.quantidade +
          '</p>' +
          '<p><strong>Arquivo:</strong> ' +
          escaparHtmlExportacaoVagas_(
            nome
          ) +
          '</p>' +
          '<p style="margin-top:24px;">' +
          '<a href="' +
          escaparAtributoExportacaoVagas_(
            url
          ) +
          '" target="_blank" style="background:#4f81bd;color:#ffffff;text-decoration:none;padding:10px 16px;border-radius:6px;font-weight:bold;">Abrir arquivo no Drive</a>' +
          '</p>' +
          '<p style="margin-top:24px;color:#607080;font-size:12px;">O arquivo foi salvo na pasta "' +
          escaparHtmlExportacaoVagas_(
            CONFIG_EXPORTAR_VAGAS_REGULACAO
              .NOME_PASTA
          ) +
          '".</p>' +
        '</div>'
      )
      .setWidth(480)
      .setHeight(330);

  SpreadsheetApp
    .getUi()
    .showModalDialog(
      html,
      'Vagas para a Regulação'
    );
}
/**
 * Verifica as dependências necessárias para a exportação.
 */
function validarDependenciasExportacaoVagas_(
  ss
) {
  if (
    !ss.getSheetByName(
      CONFIG_EXPORTAR_VAGAS_REGULACAO
        .ABA_ORIGEM
    )
  ) {
    throw new Error(
      'A aba "Vagas para Regulação" não foi encontrada.'
    );
  }

  const dependencias = [
    {
      nome:
        'interpretarCompetenciaVagasRegulacao_',
      tipo:
        typeof interpretarCompetenciaVagasRegulacao_
    },

    {
      nome:
        'formatarCompetenciaVagasRegulacao_',
      tipo:
        typeof formatarCompetenciaVagasRegulacao_
    },

    {
      nome:
        'gerarVagasCompetenciaRegulacao_',
      tipo:
        typeof gerarVagasCompetenciaRegulacao_
    },

    {
      nome:
        'dataPertenceCompetenciaVagasRegulacao_',
      tipo:
        typeof dataPertenceCompetenciaVagasRegulacao_
    },

    {
      nome:
        'normalizarTextoVagasRegulacao_',
      tipo:
        typeof normalizarTextoVagasRegulacao_
    },

    {
      nome:
        'chaveHorarioVagasRegulacao_',
      tipo:
        typeof chaveHorarioVagasRegulacao_
    },

    {
      nome:
        'compararVagasRegulacao_',
      tipo:
        typeof compararVagasRegulacao_
    },

    {
      nome:
        'obterNomeDiaVagasRegulacao_',
      tipo:
        typeof obterNomeDiaVagasRegulacao_
    }
  ];

  const ausente =
    dependencias.find(
      function(item) {
        return (
          item.tipo !== 'function'
        );
      }
    );

  if (ausente) {
    throw new Error(
      'O arquivo VagasRegulacao.gs está incompleto. Função ausente: ' +
        ausente.nome +
        '.'
    );
  }
}


/**
 * Protege textos inseridos no conteúdo HTML.
 */
function escaparHtmlExportacaoVagas_(
  valor
) {
  return String(valor || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}


/**
 * Protege valores utilizados em atributos HTML.
 */
function escaparAtributoExportacaoVagas_(
  valor
) {
  return escaparHtmlExportacaoVagas_(
    valor
  );
}
