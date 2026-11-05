
/**
 * DRIVER CONTROL PRO
 * Módulo 2.3.6A
 *
 * Estrutura e diagnóstico do histórico de status dos veículos.
 *
 * Esta etapa cria somente a estrutura da aba.
 * Não altera registros de veículos ou vínculos.
 */

const DCP_HISTORICO_STATUS_236 = {
  ABA: 'HISTORICO_STATUS_VEICULOS',

  CABECALHOS: [
    'HistoricoStatusID',
    'CodigoHistoricoStatus',
    'VeiculoID',
    'CodigoVeiculo',
    'StatusAnterior',
    'StatusNovo',
    'Motivo',
    'DataHora',
    'UsuarioID',
    'Observacoes',
    'Resultado',
    'AuditoriaID'
  ],

  LARGURA_COLUNAS: 12
};


/**
 * Cria a aba e seus cabeçalhos, caso ainda não existam.
 *
 * Se a aba já existir, confere a estrutura sem apagar dados.
 */
function DCP_CRIAR_ESTRUTURA_HISTORICO_STATUS_236() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();

  let aba = planilha.getSheetByName(
    DCP_HISTORICO_STATUS_236.ABA
  );

  if (!aba) {
    aba = planilha.insertSheet(
      DCP_HISTORICO_STATUS_236.ABA
    );
  }

  const ultimaLinha = aba.getLastRow();
  const ultimaColuna = aba.getLastColumn();

  if (ultimaLinha === 0 || ultimaColuna === 0) {
    aba.getRange(
      1,
      1,
      1,
      DCP_HISTORICO_STATUS_236.LARGURA_COLUNAS
    ).setValues([
      DCP_HISTORICO_STATUS_236.CABECALHOS
    ]);

    aba.setFrozenRows(1);

    return {
      sucesso: true,
      criada: true,
      mensagem: 'Aba de histórico criada com sucesso.',
      cabecalhos: DCP_HISTORICO_STATUS_236.CABECALHOS
    };
  }

  const quantidade = DCP_HISTORICO_STATUS_236.LARGURA_COLUNAS;

  const encontrados = aba
    .getRange(1, 1, 1, quantidade)
    .getDisplayValues()[0]
    .map(function(valor) {
      return String(valor || '').trim();
    });

  const divergencias = [];

  DCP_HISTORICO_STATUS_236.CABECALHOS.forEach(
    function(esperado, indice) {
      if (encontrados[indice] !== esperado) {
        divergencias.push({
          coluna: indice + 1,
          esperado: esperado,
          encontrado: encontrados[indice]
        });
      }
    }
  );

  if (divergencias.length > 0) {
    return {
      sucesso: false,
      criada: false,
      mensagem: 'A aba já existe, mas possui divergências nos cabeçalhos. Nenhum dado foi alterado.',
      divergencias: divergencias
    };
  }

  aba.setFrozenRows(1);

  return {
    sucesso: true,
    criada: false,
    mensagem: 'A estrutura do histórico já existe e está correta.',
    cabecalhos: encontrados
  };
}


/**
 * Diagnostica a estrutura da aba de histórico.
 */
function DCP_DIAGNOSTICAR_HISTORICO_STATUS_236() {
  try {
    const planilha = SpreadsheetApp.getActiveSpreadsheet();

    const aba = planilha.getSheetByName(
      DCP_HISTORICO_STATUS_236.ABA
    );

    if (!aba) {
      throw new Error(
        'A aba HISTORICO_STATUS_VEICULOS não existe.'
      );
    }

    const quantidade =
      DCP_HISTORICO_STATUS_236.LARGURA_COLUNAS;

    const encontrados = aba
      .getRange(1, 1, 1, quantidade)
      .getDisplayValues()[0]
      .map(function(valor) {
        return String(valor || '').trim();
      });

    const divergencias = [];

    DCP_HISTORICO_STATUS_236.CABECALHOS.forEach(
      function(esperado, indice) {
        if (encontrados[indice] !== esperado) {
          divergencias.push({
            coluna: indice + 1,
            esperado: esperado,
            encontrado: encontrados[indice]
          });
        }
      }
    );

    const resultado = {
      sucesso: divergencias.length === 0,
      aba: DCP_HISTORICO_STATUS_236.ABA,
      totalColunasEsperadas: quantidade,
      totalLinhas: aba.getLastRow(),
      estruturaCorreta: divergencias.length === 0,
      divergencias: divergencias,
      mensagem: divergencias.length === 0
        ? 'Estrutura do histórico validada com sucesso.'
        : 'Foram encontradas divergências na estrutura.'
    };

    Logger.log(JSON.stringify(resultado, null, 2));

    return resultado;

  } catch (erro) {
    const resultado = {
      sucesso: false,
      mensagem: erro.message
    };

    Logger.log(JSON.stringify(resultado, null, 2));

    return resultado;
  }
}
