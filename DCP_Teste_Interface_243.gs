function DCP_TESTAR_CARREGAMENTO_INTERFACE_243() {
  const resultado = DCP_FORM243_CARREGAR_DADOS();

  Logger.log(JSON.stringify(resultado, null, 2));

  return resultado;
}