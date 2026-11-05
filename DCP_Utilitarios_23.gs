
function DCP_LIBERAR_EDICAO_TEMPORARIA() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ss.getSheets().forEach(aba => {
    // Remove proteções da aba.
    aba.getProtections(SpreadsheetApp.ProtectionType.SHEET)
      .forEach(protecao => protecao.remove());

    // Remove proteções de intervalos.
    aba.getProtections(SpreadsheetApp.ProtectionType.RANGE)
      .forEach(protecao => protecao.remove());
  });

  SpreadsheetApp.flush();

  Logger.log('Proteções removidas de todas as abas.');
  Logger.log('Agora é possível redimensionar as colunas.');
}
