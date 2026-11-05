function DCP_DIAGNOSTICAR_VALIDACOES_VINCULOS_245() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName('VINCULOS');

  if (!aba) throw new Error('Aba VINCULOS não encontrada.');

  const cabecalhos = aba.getRange(1, 1, 1, 16).getValues()[0];
  const linha = 2;
  const validacoes = aba.getRange(linha, 1, 1, 16).getDataValidations()[0];

  const resultado = cabecalhos.map((nome, i) => {
    const regra = validacoes[i];

    if (!regra) {
      return {
        coluna: i + 1,
        campo: nome,
        validacao: 'SEM VALIDACAO'
      };
    }

    return {
      coluna: i + 1,
      campo: nome,
      criterio: String(regra.getCriteriaType()),
      valoresPermitidos: regra.getCriteriaValues().map(String),
      rejeitarInvalido: regra.getAllowInvalid() === false
    };
  });

  Logger.log(JSON.stringify(resultado, null, 2));
  return resultado;
}