function DCP_CORRIGIR_VALIDACAO_TIPO_VINCULO_245() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName('VINCULOS');

  if (!aba) {
    throw new Error('Aba VINCULOS não encontrada.');
  }

  const cabecalho = aba.getRange(1, 5).getValue();

  if (cabecalho !== 'TipoVinculo') {
    throw new Error(
      'A coluna E não corresponde a TipoVinculo. Nenhuma alteração realizada.'
    );
  }

  const regra = SpreadsheetApp.newDataValidation()
    .requireValueInList(
      ['PROPRIO', 'ALUGADO', 'CEDIDO'],
      true
    )
    .setAllowInvalid(false)
    .setHelpText(
      'Selecione PROPRIO, ALUGADO ou CEDIDO.'
    )
    .build();

  const quantidadeLinhas = aba.getMaxRows() - 1;

  aba.getRange(2, 5, quantidadeLinhas, 1)
    .setDataValidation(regra);

  SpreadsheetApp.flush();

  Logger.log(
    'SUCESSO: validação de TipoVinculo corrigida.'
  );

  Logger.log(
    'Valores permitidos: PROPRIO, ALUGADO, CEDIDO.'
  );

  return {
    sucesso: true,
    coluna: 'TipoVinculo',
    valoresPermitidos: ['PROPRIO', 'ALUGADO', 'CEDIDO']
  };
}