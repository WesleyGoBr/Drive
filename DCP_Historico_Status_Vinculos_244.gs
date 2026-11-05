
/**
 * DRIVER CONTROL PRO V1.0
 * MODULO 2.4.4.1
 * Estrutura do historico de status dos vinculos
 *
 * Responsabilidades:
 * - Criar HISTORICO_STATUS_VINCULOS.
 * - Validar os cabecalhos.
 * - Diagnosticar a estrutura.
 *
 * Este modulo nao encerra vinculos nem altera registros existentes.
 */

const DCP_HISTORICO_VINCULOS_244 = Object.freeze({

  ABA: 'HISTORICO_STATUS_VINCULOS',

  CABECALHOS: Object.freeze([
    'HistoricoStatusID',
    'CodigoHistoricoStatus',
    'VinculoID',
    'CodigoVinculo',
    'StatusAnterior',
    'StatusNovo',
    'Motivo',
    'DataHora',
    'UsuarioID',
    'Observacoes',
    'Resultado',
    'AuditoriaID'
  ]),

  LINHA_CABECALHO: 1
});


/**
 * Retorna a planilha ativa.
 */
function DCP_HV244_PLANILHA() {

  const planilha = SpreadsheetApp.getActiveSpreadsheet();

  if (!planilha) {
    throw new Error('Nao foi possivel identificar a planilha ativa.');
  }

  return planilha;
}


/**
 * Localiza a aba do historico.
 */
function DCP_HV244_OBTER_ABA() {

  const planilha = DCP_HV244_PLANILHA();

  return planilha.getSheetByName(
    DCP_HISTORICO_VINCULOS_244.ABA
  );
}


/**
 * Compara os cabecalhos existentes com os esperados.
 */
function DCP_HV244_VALIDAR_CABECALHOS(aba) {

  if (!aba) {
    throw new Error('A aba do historico nao foi encontrada.');
  }

  const esperados =
    DCP_HISTORICO_VINCULOS_244.CABECALHOS;

  const encontrados = aba
    .getRange(
      DCP_HISTORICO_VINCULOS_244.LINHA_CABECALHO,
      1,
      1,
      esperados.length
    )
    .getDisplayValues()[0];

  const divergencias = [];

  for (let i = 0; i < esperados.length; i++) {

    if (String(encontrados[i]).trim() !== esperados[i]) {

      divergencias.push({
        coluna: i + 1,
        esperado: esperados[i],
        encontrado: encontrados[i] || ''
      });
    }
  }

  return {
    valido: divergencias.length === 0,
    divergencias: divergencias,
    totalColunasEsperadas: esperados.length
  };
}


/**
 * Cria a estrutura do historico.
 *
 * Se a aba ja existir:
 * - Nao apaga dados.
 * - Nao substitui cabecalhos.
 * - Apenas valida a estrutura.
 */
function DCP_CRIAR_ESTRUTURA_HISTORICO_VINCULOS_244() {

  const lock = LockService.getDocumentLock();

  if (!lock.tryLock(15000)) {
    throw new Error(
      'Nao foi possivel obter o bloqueio da planilha. Tente novamente.'
    );
  }

  try {

    const planilha = DCP_HV244_PLANILHA();

    const nomeAba =
      DCP_HISTORICO_VINCULOS_244.ABA;

    let aba = planilha.getSheetByName(nomeAba);

    let criada = false;

    if (!aba) {

      aba = planilha.insertSheet(nomeAba);
      criada = true;

      const cabecalhos =
        DCP_HISTORICO_VINCULOS_244.CABECALHOS;

      aba.getRange(
        1,
        1,
        1,
        cabecalhos.length
      ).setValues([cabecalhos]);

      aba.setFrozenRows(1);

      aba.getRange(
        1,
        1,
        1,
        cabecalhos.length
      ).setFontWeight('bold');

      SpreadsheetApp.flush();
    }

    const validacao =
      DCP_HV244_VALIDAR_CABECALHOS(aba);

    if (!validacao.valido) {

      return {
        sucesso: false,
        modulo: '2.4.4.1',
        aba: nomeAba,
        criada: criada,
        mensagem: 'A estrutura possui divergencias.',
        divergencias: validacao.divergencias
      };
    }

    return {
      sucesso: true,
      modulo: '2.4.4.1',
      aba: nomeAba,
      criada: criada,
      totalColunas: validacao.totalColunasEsperadas,
      ultimaLinha: aba.getLastRow(),
      mensagem: criada
        ? 'Aba de historico criada e validada com sucesso.'
        : 'Aba existente conferida e validada com sucesso.'
    };

  } finally {

    lock.releaseLock();

  }
}


/**
 * Diagnostico somente leitura.
 *
 * Nao cria abas e nao altera dados.
 */
function DCP_DIAGNOSTICAR_ESTRUTURA_HISTORICO_VINCULOS_244() {

  const planilha = DCP_HV244_PLANILHA();

  const aba = planilha.getSheetByName(
    DCP_HISTORICO_VINCULOS_244.ABA
  );

  if (!aba) {

    return {
      sucesso: false,
      somenteLeitura: true,
      abaExiste: false,
      mensagem: 'A aba HISTORICO_STATUS_VINCULOS ainda nao existe.'
    };
  }

  const validacao =
    DCP_HV244_VALIDAR_CABECALHOS(aba);

  return {
    sucesso: validacao.valido,
    somenteLeitura: true,
    abaExiste: true,
    cabecalhosCorretos: validacao.valido,
    totalColunasEsperadas: validacao.totalColunasEsperadas,
    ultimaLinha: aba.getLastRow(),
    divergencias: validacao.divergencias,
    mensagem: validacao.valido
      ? 'Estrutura do historico validada.'
      : 'Foram encontradas divergencias nos cabecalhos.'
  };
}


/**
 * Teste somente leitura dos cabecalhos esperados.
 * Nao acessa nem modifica registros reais.
 */
function DCP_TESTAR_CABECALHOS_HISTORICO_VINCULOS_244() {

  const esperados =
    DCP_HISTORICO_VINCULOS_244.CABECALHOS;

  const resultado = {
    sucesso: true,
    somenteLeitura: true,
    totalCampos: esperados.length,
    campos: esperados,
    camposDuplicados: []
  };

  const encontrados = {};

  esperados.forEach(function(campo) {

    if (encontrados[campo]) {
      resultado.camposDuplicados.push(campo);
    }

    encontrados[campo] = true;

  });

  resultado.sucesso =
    resultado.camposDuplicados.length === 0;

  resultado.mensagem = resultado.sucesso
    ? 'Os 12 campos foram conferidos. Nenhuma duplicidade encontrada.'
    : 'Foram encontrados campos duplicados.';

  return resultado;
}
