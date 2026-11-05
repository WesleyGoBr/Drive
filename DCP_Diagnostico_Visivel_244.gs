/**
 * DRIVER CONTROL PRO
 * Diagnóstico visível do módulo 2.4.4.2
 *
 * Grava o resultado em uma aba própria da planilha.
 * Não altera VINCULOS, HISTORICO_STATUS_VINCULOS,
 * CONFIG ou AUDITORIA.
 */

function DCP_DIAGNOSTICO_VISIVEL_ENCERRAMENTO_244() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error('Não foi possível acessar a planilha ativa.');
  }

  const nomeAba = 'DIAGNOSTICO_244';

  let aba = ss.getSheetByName(nomeAba);

  if (!aba) {
    aba = ss.insertSheet(nomeAba);
  }

  const dataHora = new Date();

  let resultado;

  try {

    resultado =
      DCP_DIAGNOSTICAR_MOTOR_ENCERRAMENTO_VINCULOS_244();

  } catch (erro) {

    resultado = {
      sucesso: false,
      erro: erro.message || String(erro),
      pilha: erro.stack || ''
    };

  }

  let conteudo;

  try {
    conteudo = JSON.stringify(resultado, null, 2);
  } catch (erro) {
    conteudo = String(resultado);
  }

  if (conteudo === undefined) {
    conteudo = 'A função não retornou nenhum resultado.';
  }

  const linhaInicial = aba.getLastRow() + 2;

  aba.getRange(linhaInicial, 1, 1, 2).setValues([
    ['DATA/HORA', dataHora]
  ]);

  aba.getRange(linhaInicial + 1, 1).setValue('RESULTADO');

  aba.getRange(linhaInicial + 1, 2).setValue(conteudo);

  aba.getRange(linhaInicial, 1, 2, 2)
    .setFontWeight('bold');

  aba.getRange(linhaInicial + 1, 2)
    .setWrap(true);

  aba.setColumnWidth(1, 150);
  aba.setColumnWidth(2, 750);

  SpreadsheetApp.flush();

  return {
    sucesso: true,
    mensagem: 'Diagnóstico gravado na aba ' + nomeAba,
    resultado: resultado
  };
}
function DCP_TESTE_DATAS_VISIVEL_244() {

  const resultado =
    DCP_TESTAR_DATAS_ENCERRAMENTO_VINCULOS_244();

  const texto = JSON.stringify(resultado, null, 2);

  console.log('=== TESTE DE DATAS 2.4.4.2 ===');
  console.log(texto);

  Logger.log('=== TESTE DE DATAS 2.4.4.2 ===');
  Logger.log(texto);

  return resultado;
}