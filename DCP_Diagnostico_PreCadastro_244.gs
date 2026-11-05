
function DCP_DIAGNOSTICAR_PRE_CADASTRO_244() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName('AUDITORIA');

  if (!aba) {
    throw new Error('A aba AUDITORIA não foi encontrada.');
  }

  const resultado = {
    fusoScript: Session.getScriptTimeZone(),
    fusoPlanilha: ss.getSpreadsheetTimeZone(),
    validacoesAuditoria: {}
  };

  const campos = {
    Acao: 4,
    Resultado: 9,
    Origem: 10
  };

  Object.keys(campos).forEach(function(nome) {
    const coluna = campos[nome];
    const regra = aba.getRange(2, coluna).getDataValidation();

    if (!regra) {
      resultado.validacoesAuditoria[nome] = {
        possuiValidacao: false
      };
      return;
    }

    const criterio = regra.getCriteriaType();
    const argumentos = regra.getCriteriaValues();

    resultado.validacoesAuditoria[nome] = {
      possuiValidacao: true,
      criterio: String(criterio),
      valoresPermitidos: argumentos.map(function(valor) {
        if (Array.isArray(valor)) {
          return valor.map(String);
        }
        return String(valor);
      }),
      rejeitarEntradaInvalida: regra.getAllowInvalid() === false
    };
  });

  Logger.log(JSON.stringify(resultado, null, 2));
  return resultado;
}
