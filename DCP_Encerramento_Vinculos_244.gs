/**
 * DRIVER CONTROL PRO V1.0
 * MODULO 2.4.4.2 - MOTOR DE ENCERRAMENTO
 *
 * VERSAO REFORCADA:
 * - Idempotencia vinculada ao conteudo da solicitacao.
 * - Fuso America/Sao_Paulo.
 * - Gravacao coordenada do vinculo.
 * - Deteccao de chaves tecnicas duplicadas.
 * - Deteccao de historicos duplicados.
 * - Deteccao de auditorias duplicadas.
 * - Recuperacao por conferencia dos registros.
 * - Diagnosticos somente leitura.
 *
 * IMPORTANTE:
 * Nenhum encerramento automatico.
 * Nenhuma exclusao de registros.
 * Nenhuma reversao automatica.
 */

const DCP_ENCERRAMENTO_VINCULOS_244 = Object.freeze({

  ABAS: Object.freeze({
    VINCULOS: 'VINCULOS',
    HISTORICO: 'HISTORICO_STATUS_VINCULOS',
    CONFIG: 'CONFIG',
    AUDITORIA: 'AUDITORIA'
  }),

  FUSO: 'America/Sao_Paulo',

  PREFIXO_HISTORICO: 'HVI',
  CHAVE_SEQUENCIA: 'SEQ_HVI',
  PREFIXO_IDEMPOTENCIA: 'IDEMP_FIM_VIN_',

  USUARIO: 'ADMIN_CENTRAL',

  MODOS: Object.freeze(['IMEDIATO', 'INFORMADO']),

  STATUS_ATIVO: 'ATIVO',
  STATUS_ENCERRADO: 'ENCERRADO',

  CABECALHOS_VINCULOS: Object.freeze([
    'VinculoID',
    'CodigoVinculo',
    'MotoristaID',
    'VeiculoID',
    'TipoVinculo',
    'InicioEm',
    'FimEm',
    'CompartilhamentoAutorizado',
    'ValorAluguel',
    'PeriodicidadeAluguel',
    'Observacoes',
    'Status',
    'CriadoEm',
    'AtualizadoEm',
    'CriadoPor',
    'AtualizadoPor'
  ]),

  CABECALHOS_HISTORICO: Object.freeze([
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

  CABECALHOS_CONFIG: Object.freeze([
    'Chave',
    'Valor',
    'TipoValor',
    'Descricao',
    'Editavel',
    'AtualizadoEm'
  ]),

  CABECALHOS_AUDITORIA: Object.freeze([
    'AuditoriaID',
    'DataHora',
    'UsuarioID',
    'Acao',
    'Tabela',
    'RegistroID',
    'ValorAnterior',
    'ValorNovo',
    'Resultado',
    'Origem',
    'Detalhes'
  ])

});


/* =====================================================
   FUNCOES INTERNAS
===================================================== */

function DCP_FIM244_TEXTO(valor) {
  return String(valor == null ? '' : valor).trim();
}


function DCP_FIM244_ABA(nome) {

  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const aba = planilha.getSheetByName(nome);

  if (!aba) {
    throw new Error('Aba obrigatoria nao encontrada: ' + nome);
  }

  return aba;
}


function DCP_FIM244_VALIDAR_CABECALHOS(aba, esperados) {

  if (aba.getLastRow() < 1) {
    throw new Error('Aba sem cabecalhos: ' + aba.getName());
  }

  const encontrados = aba
    .getRange(1, 1, 1, esperados.length)
    .getDisplayValues()[0];

  const divergencias = [];

  esperados.forEach(function(campo, indice) {

    if (String(encontrados[indice]).trim() !== campo) {
      divergencias.push({
        coluna: indice + 1,
        esperado: campo,
        encontrado: encontrados[indice] || ''
      });
    }

  });

  if (divergencias.length) {
    throw new Error(
      'Cabecalhos divergentes em ' +
      aba.getName() + ': ' +
      JSON.stringify(divergencias)
    );
  }
}


function DCP_FIM244_UUID() {
  return Utilities.getUuid();
}


function DCP_FIM244_HASH(valor) {

  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    valor,
    Utilities.Charset.UTF_8
  );

  return Utilities.base64EncodeWebSafe(bytes)
    .replace(/=+$/g, '');
}


function DCP_FIM244_CONVERTER_DATA_LOCAL(valor) {

  if (valor instanceof Date && !isNaN(valor.getTime())) {
    return new Date(valor.getTime());
  }

  const texto = DCP_FIM244_TEXTO(valor);

  const match = texto.match(
    /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})(?::(\d{2}))?$/
  );

  if (!match) {
    throw new Error(
      'Data/hora invalida. Utilize YYYY-MM-DDTHH:mm.'
    );
  }

  const ano = Number(match[1]);
  const mes = Number(match[2]);
  const dia = Number(match[3]);
  const hora = Number(match[4]);
  const minuto = Number(match[5]);
  const segundo = Number(match[6] || 0);

  if (
    ano < 2000 || ano > 2100 ||
    mes < 1 || mes > 12 ||
    dia < 1 || dia > 31 ||
    hora < 0 || hora > 23 ||
    minuto < 0 || minuto > 59 ||
    segundo < 0 || segundo > 59
  ) {
    throw new Error('Data/hora fora dos limites permitidos.');
  }

  const formato = match[6]
    ? 'yyyy-MM-dd HH:mm:ss'
    : 'yyyy-MM-dd HH:mm';

  const textoFormatado =
    match[1] + '-' + match[2] + '-' + match[3] +
    ' ' + match[4] + ':' + match[5] +
    (match[6] ? ':' + match[6] : '');

  const data = Utilities.parseDate(
    textoFormatado,
    DCP_ENCERRAMENTO_VINCULOS_244.FUSO,
    formato
  );

  const conferencia = Utilities.formatDate(
    data,
    DCP_ENCERRAMENTO_VINCULOS_244.FUSO,
    formato
  );

  if (
    !data ||
    isNaN(data.getTime()) ||
    conferencia !== textoFormatado
  ) {
    throw new Error('A data/hora informada e inexistente.');
  }

  return data;
}


function DCP_FIM244_DATA_EXISTENTE(valor, campo) {

  if (valor instanceof Date && !isNaN(valor.getTime())) {
    return new Date(valor.getTime());
  }

  throw new Error('Data invalida no campo ' + campo + '.');
}


/**
 * Localiza todas as ocorrencias de uma chave na CONFIG.
 * Mais de uma ocorrencia e considerada inconsistencia.
 */
function DCP_FIM244_LOCALIZAR_CHAVES(aba, chave) {

  const ultimaLinha = aba.getLastRow();

  if (ultimaLinha < 2) return [];

  const valores = aba
    .getRange(2, 1, ultimaLinha - 1, 1)
    .getDisplayValues();

  const linhas = [];

  valores.forEach(function(item, indice) {

    if (DCP_FIM244_TEXTO(item[0]) === chave) {
      linhas.push(indice + 2);
    }

  });

  return linhas;
}


function DCP_FIM244_LOCALIZAR_CHAVE(aba, chave) {

  const linhas = DCP_FIM244_LOCALIZAR_CHAVES(aba, chave);

  if (linhas.length > 1) {
    throw new Error(
      'Chave duplicada na CONFIG: ' + chave +
      '. Necessaria revisao manual.'
    );
  }

  return linhas.length ? linhas[0] : 0;
}


function DCP_FIM244_LER_CONFIG(aba, chave) {

  const linha = DCP_FIM244_LOCALIZAR_CHAVE(aba, chave);

  if (!linha) return null;

  return {
    linha: linha,
    valor: aba.getRange(linha, 2).getValue()
  };
}


/**
 * Deve ser chamada com DocumentLock adquirido.
 */
function DCP_FIM244_GRAVAR_CONFIG(
  aba,
  chave,
  valor,
  tipo,
  descricao,
  editavel
) {

  const linhaExistente = DCP_FIM244_LOCALIZAR_CHAVE(
    aba,
    chave
  );

  const linha = linhaExistente ||
    Math.max(2, aba.getLastRow() + 1);

  const registro = [
    chave,
    valor,
    tipo || 'TEXTO',
    descricao || '',
    editavel === true,
    new Date()
  ];

  aba.getRange(linha, 1, 1, 3).clearDataValidations();

  aba.getRange(linha, 1, 1, 6).setValues([registro]);

  return linha;
}


function DCP_FIM244_GERAR_CODIGO(abaConfig, abaHistorico) {

  const chave = DCP_ENCERRAMENTO_VINCULOS_244.CHAVE_SEQUENCIA;

  const configuracao = DCP_FIM244_LER_CONFIG(
    abaConfig,
    chave
  );

  let sequencia = configuracao
    ? Number(configuracao.valor)
    : 0;

  if (!Number.isSafeInteger(sequencia) || sequencia < 0) {
    throw new Error('Sequencia SEQ_HVI invalida.');
  }

  const ultimaLinha = abaHistorico.getLastRow();

  if (ultimaLinha >= 2) {

    const codigos = abaHistorico
      .getRange(2, 2, ultimaLinha - 1, 1)
      .getDisplayValues();

    codigos.forEach(function(item) {

      const match = DCP_FIM244_TEXTO(item[0])
        .match(/^HVI-(\d{4,})$/);

      if (match) {
        sequencia = Math.max(
          sequencia,
          Number(match[1])
        );
      }

    });

  }

  sequencia++;

  DCP_FIM244_GRAVAR_CONFIG(
    abaConfig,
    chave,
    sequencia,
    'NUMERO',
    'Sequencia de historicos de vinculos.',
    false
  );

  return 'HVI-' + String(sequencia).padStart(4, '0');
}


function DCP_FIM244_LOCALIZAR_VINCULOS(aba, identificador) {

  const procurado = DCP_FIM244_TEXTO(identificador);
  const ultimaLinha = aba.getLastRow();

  if (ultimaLinha < 2) return [];

  const dados = aba
    .getRange(2, 1, ultimaLinha - 1, 16)
    .getValues();

  const encontrados = [];

  dados.forEach(function(registro, indice) {

    if (
      procurado === DCP_FIM244_TEXTO(registro[0]) ||
      procurado === DCP_FIM244_TEXTO(registro[1])
    ) {
      encontrados.push({
        linha: indice + 2,
        dados: registro
      });
    }

  });

  return encontrados;
}


function DCP_FIM244_LOCALIZAR_VINCULO(aba, identificador) {

  const encontrados = DCP_FIM244_LOCALIZAR_VINCULOS(
    aba,
    identificador
  );

  if (encontrados.length > 1) {
    throw new Error(
      'Identificador de vinculo duplicado. Operacao bloqueada.'
    );
  }

  return encontrados.length ? encontrados[0] : null;
}


function DCP_FIM244_LOCALIZAR_HISTORICOS(aba, chaveTentativa) {

  const marcador = '[DCP_FIM244_' + chaveTentativa + ']';
  const ultimaLinha = aba.getLastRow();

  if (ultimaLinha < 2) return [];

  const observacoes = aba
    .getRange(2, 10, ultimaLinha - 1, 1)
    .getDisplayValues();

  const linhas = [];

  observacoes.forEach(function(item, indice) {

    if (String(item[0]).indexOf(marcador) !== -1) {
      linhas.push(indice + 2);
    }

  });

  return linhas;
}


function DCP_FIM244_LOCALIZAR_HISTORICO(aba, chaveTentativa) {

  const linhas = DCP_FIM244_LOCALIZAR_HISTORICOS(
    aba,
    chaveTentativa
  );

  if (linhas.length > 1) {
    throw new Error(
      'Marcador de historico duplicado. Operacao bloqueada.'
    );
  }

  return linhas.length ? linhas[0] : null;
}


function DCP_FIM244_AUDITORIAS_EXISTENTES(aba, auditoriaID) {

  const ultimaLinha = aba.getLastRow();

  if (ultimaLinha < 2) return [];

  const ids = aba
    .getRange(2, 1, ultimaLinha - 1, 1)
    .getDisplayValues();

  const linhas = [];

  ids.forEach(function(item, indice) {

    if (DCP_FIM244_TEXTO(item[0]) === auditoriaID) {
      linhas.push(indice + 2);
    }

  });

  return linhas;
}


function DCP_FIM244_AUDITORIA_EXISTE(aba, auditoriaID) {

  const linhas = DCP_FIM244_AUDITORIAS_EXISTENTES(
    aba,
    auditoriaID
  );

  if (linhas.length > 1) {
    throw new Error(
      'ID de auditoria duplicado. Revisao manual necessaria.'
    );
  }

  return linhas.length === 1;
}


function DCP_FIM244_REGISTRAR_AUDITORIA(aba, evento) {

  const existentes = DCP_FIM244_AUDITORIAS_EXISTENTES(
    aba,
    evento.auditoriaID
  );

  if (existentes.length > 1) {
    throw new Error('Auditoria duplicada.');
  }

  const registro = [
    evento.auditoriaID,
    evento.dataHora,
    DCP_ENCERRAMENTO_VINCULOS_244.USUARIO,
    'ENCERRAMENTO',
    'VINCULOS',
    evento.vinculoID,
    'ATIVO',
    'ENCERRADO',
    'SUCESSO',
    'SCRIPT',
    evento.detalhes
  ];

  if (existentes.length === 0) {
    aba.appendRow(registro);
  } else {

    const existente = aba
      .getRange(existentes[0], 1, 1, 11)
      .getValues()[0];

    const coerente =
      DCP_FIM244_TEXTO(existente[2]) ===
        DCP_ENCERRAMENTO_VINCULOS_244.USUARIO &&
      DCP_FIM244_TEXTO(existente[3]) === 'ENCERRAMENTO' &&
      DCP_FIM244_TEXTO(existente[4]) === 'VINCULOS' &&
      DCP_FIM244_TEXTO(existente[5]) === evento.vinculoID &&
      DCP_FIM244_TEXTO(existente[8]) === 'SUCESSO';

    if (!coerente) {
      throw new Error(
        'Auditoria existente possui conteudo divergente.'
      );
    }

  }

  SpreadsheetApp.flush();

  const confirmacao = DCP_FIM244_AUDITORIAS_EXISTENTES(
    aba,
    evento.auditoriaID
  );

  if (confirmacao.length !== 1) {
    throw new Error('Auditoria nao confirmada de forma unica.');
  }
}


function DCP_FIM244_RESPOSTA_ERRO(
  codigo,
  mensagem,
  gravacaoRealizada,
  tentativaInconclusiva
) {

  return {
    sucesso: false,
    codigo: codigo,
    mensagem: mensagem,
    gravacaoRealizada: gravacaoRealizada === true,
    gravacaoParcial: tentativaInconclusiva === true,
    tentativaInconclusiva: tentativaInconclusiva === true
  };
}


/* =====================================================
   VALIDACAO E IDENTIDADE DA SOLICITACAO
===================================================== */

function DCP_FIM244_VALIDAR_DADOS(dados) {

  if (!dados || typeof dados !== 'object') {
    throw new Error('Dados do encerramento nao informados.');
  }

  const identificador = DCP_FIM244_TEXTO(
    dados.VinculoID || dados.CodigoVinculo
  );

  const chave = DCP_FIM244_TEXTO(dados.ChaveTentativa);

  if (!identificador) {
    throw new Error('Informe o vinculo que sera encerrado.');
  }

  if (!/^[a-zA-Z0-9-]{20,80}$/.test(chave)) {
    throw new Error('ChaveTentativa invalida.');
  }

  const modo = DCP_FIM244_TEXTO(
    dados.ModoEncerramento
  ).toUpperCase();

  if (
    DCP_ENCERRAMENTO_VINCULOS_244.MODOS.indexOf(modo) === -1
  ) {
    throw new Error('Modo de encerramento invalido.');
  }

  const motivo = DCP_FIM244_TEXTO(dados.Motivo);
  const observacoes = DCP_FIM244_TEXTO(dados.Observacoes);

  if (!motivo) {
    throw new Error('O motivo do encerramento e obrigatorio.');
  }

  if (motivo.length > 500) {
    throw new Error('Motivo limitado a 500 caracteres.');
  }

  if (observacoes.length > 1000) {
    throw new Error('Observacoes limitadas a 1000 caracteres.');
  }

  const dataHora = modo === 'IMEDIATO'
    ? new Date()
    : DCP_FIM244_CONVERTER_DATA_LOCAL(
        dados.DataHoraEncerramento
      );

  if (dataHora.getTime() > Date.now()) {
    throw new Error('Nao e permitido registrar encerramento futuro.');
  }

  return {
    identificador: identificador,
    chaveTentativa: chave,
    modo: modo,
    motivo: motivo,
    observacoes: observacoes,
    dataHora: dataHora
  };
}


function DCP_FIM244_HASH_SOLICITACAO(solicitacao) {

  const base = {
    identificador: solicitacao.identificador,
    modo: solicitacao.modo,
    motivo: solicitacao.motivo,
    observacoes: solicitacao.observacoes,
    dataInformada: solicitacao.modo === 'INFORMADO'
      ? solicitacao.dataHora.toISOString()
      : null
  };

  return DCP_FIM244_HASH(JSON.stringify(base));
}


function DCP_FIM244_CHAVE_IDEMPOTENCIA(chaveTentativa) {

  return DCP_ENCERRAMENTO_VINCULOS_244.PREFIXO_IDEMPOTENCIA +
    chaveTentativa;
}


function DCP_FIM244_LER_IDEMPOTENCIA(abaConfig, solicitacao) {

  const chaveConfig = DCP_FIM244_CHAVE_IDEMPOTENCIA(
    solicitacao.chaveTentativa
  );

  const registro = DCP_FIM244_LER_CONFIG(
    abaConfig,
    chaveConfig
  );

  if (!registro) return null;

  let dados;

  try {
    dados = typeof registro.valor === 'string'
      ? JSON.parse(registro.valor)
      : registro.valor;
  } catch (erro) {
    throw new Error(
      'Registro de idempotencia ilegivel. Nao repita a operacao.'
    );
  }

  if (!dados || !dados.hashSolicitacao) {
    throw new Error(
      'Registro de idempotencia sem impressao digital.'
    );
  }

  const hashAtual = DCP_FIM244_HASH_SOLICITACAO(solicitacao);

  if (dados.hashSolicitacao !== hashAtual) {
    throw new Error(
      'Chave de idempotencia utilizada com dados diferentes.'
    );
  }

  return {
    concluido: dados.estado === 'CONCLUIDO',
    estado: dados.estado,
    dados: dados,
    resposta: dados.resposta || null
  };
}


function DCP_FIM244_SALVAR_IDEMPOTENCIA(
  abaConfig,
  solicitacao,
  hashSolicitacao,
  estado,
  resposta,
  identificadores
) {

  const conteudo = {
    versao: 3,
    estado: estado,
    hashSolicitacao: hashSolicitacao,

    solicitacao: {
      identificador: solicitacao.identificador,
      modo: solicitacao.modo,
      motivo: solicitacao.motivo,
      observacoes: solicitacao.observacoes,
      dataHora: solicitacao.dataHora.toISOString()
    },

    identificadores: identificadores || {},
    atualizadoEm: new Date().toISOString(),
    resposta: resposta || null
  };

  DCP_FIM244_GRAVAR_CONFIG(
    abaConfig,
    DCP_FIM244_CHAVE_IDEMPOTENCIA(
      solicitacao.chaveTentativa
    ),
    JSON.stringify(conteudo),
    'JSON',
    'Controle de idempotencia do encerramento de vinculo.',
    false
  );

  SpreadsheetApp.flush();
}


/* =====================================================
   CONFERENCIA DE CONSISTENCIA
===================================================== */

function DCP_FIM244_CONFERIR_ESTADO(
  abaVinculos,
  abaHistorico,
  abaAuditoria,
  solicitacao,
  identificadores
) {

  const problemas = [];

  const codigoVinculo = identificadores.codigoVinculo;
  const vinculoID = identificadores.vinculoID;
  const auditoriaID = identificadores.auditoriaID;

  let vinculo = null;

  try {
    vinculo = DCP_FIM244_LOCALIZAR_VINCULO(
      abaVinculos,
      codigoVinculo
    );
  } catch (erro) {
    problemas.push(erro.message);
  }

  const linhasHistorico = DCP_FIM244_LOCALIZAR_HISTORICOS(
    abaHistorico,
    solicitacao.chaveTentativa
  );

  if (linhasHistorico.length > 1) {
    problemas.push('Marcador de historico duplicado.');
  }

  let historico = null;

  if (linhasHistorico.length === 1) {
    historico = abaHistorico
      .getRange(linhasHistorico[0], 1, 1, 12)
      .getValues()[0];
  }

  const linhasAuditoria = DCP_FIM244_AUDITORIAS_EXISTENTES(
    abaAuditoria,
    auditoriaID
  );

  if (linhasAuditoria.length > 1) {
    problemas.push('Auditoria duplicada.');
  }

  let auditoria = null;

  if (linhasAuditoria.length === 1) {
    auditoria = abaAuditoria
      .getRange(linhasAuditoria[0], 1, 1, 11)
      .getValues()[0];
  }

  const vinculoCoerente = !!(
    vinculo &&
    DCP_FIM244_TEXTO(vinculo.dados[0]) === vinculoID &&
    DCP_FIM244_TEXTO(vinculo.dados[1]) === codigoVinculo &&
    DCP_FIM244_TEXTO(vinculo.dados[11]) === 'ENCERRADO' &&
    vinculo.dados[6] instanceof Date &&
    vinculo.dados[6].getTime() === solicitacao.dataHora.getTime()
  );

  const historicoCoerente = !!(
    historico &&
    DCP_FIM244_TEXTO(historico[0]) === identificadores.historicoID &&
    DCP_FIM244_TEXTO(historico[1]) === identificadores.codigoHistorico &&
    DCP_FIM244_TEXTO(historico[2]) === vinculoID &&
    DCP_FIM244_TEXTO(historico[3]) === codigoVinculo &&
    DCP_FIM244_TEXTO(historico[4]) === 'ATIVO' &&
    DCP_FIM244_TEXTO(historico[5]) === 'ENCERRADO' &&
    DCP_FIM244_TEXTO(historico[6]) === solicitacao.motivo &&
    historico[7] instanceof Date &&
    historico[7].getTime() === solicitacao.dataHora.getTime() &&
    DCP_FIM244_TEXTO(historico[11]) === auditoriaID
  );

  const auditoriaCoerente = !!(
    auditoria &&
    DCP_FIM244_TEXTO(auditoria[2]) ===
      DCP_ENCERRAMENTO_VINCULOS_244.USUARIO &&
    DCP_FIM244_TEXTO(auditoria[3]) === 'ENCERRAMENTO' &&
    DCP_FIM244_TEXTO(auditoria[4]) === 'VINCULOS' &&
    DCP_FIM244_TEXTO(auditoria[5]) === vinculoID &&
    DCP_FIM244_TEXTO(auditoria[6]) === 'ATIVO' &&
    DCP_FIM244_TEXTO(auditoria[7]) === 'ENCERRADO' &&
    DCP_FIM244_TEXTO(auditoria[8]) === 'SUCESSO'
  );

  if (!vinculoCoerente) {
    problemas.push('Vinculo ausente ou divergente.');
  }

  if (!historicoCoerente) {
    problemas.push('Historico ausente ou divergente.');
  }

  if (!auditoriaCoerente) {
    problemas.push('Auditoria ausente ou divergente.');
  }

  return {
    consistente: problemas.length === 0,
    vinculoCoerente: vinculoCoerente,
    historicoCoerente: historicoCoerente,
    auditoriaCoerente: auditoriaCoerente,
    problemas: problemas
  };
}


/* =====================================================
   MOTOR PRINCIPAL
===================================================== */

function DCP_ENCERRAR_VINCULO_244(dados) {

  const lock = LockService.getDocumentLock();

  if (!lock.tryLock(20000)) {
    return DCP_FIM244_RESPOSTA_ERRO(
      'BLOQUEIO_INDISPONIVEL',
      'Outra operacao esta sendo executada.',
      false,
      false
    );
  }

  let gravacaoIniciada = false;
  let solicitacao = null;
  let hashSolicitacao = null;
  let abaConfig = null;
  let abaHistorico = null;
  let abaVinculos = null;
  let linhaHistorico = 0;
  let identificadores = {};

  try {

    solicitacao = DCP_FIM244_VALIDAR_DADOS(dados);
    hashSolicitacao = DCP_FIM244_HASH_SOLICITACAO(solicitacao);

    const planilha = SpreadsheetApp.getActiveSpreadsheet();

    if (
      Session.getScriptTimeZone() !==
      DCP_ENCERRAMENTO_VINCULOS_244.FUSO ||
      planilha.getSpreadsheetTimeZone() !==
      DCP_ENCERRAMENTO_VINCULOS_244.FUSO
    ) {
      throw new Error(
        'O fuso do projeto e da planilha deve ser America/Sao_Paulo.'
      );
    }

    abaVinculos = DCP_FIM244_ABA('VINCULOS');
    abaHistorico = DCP_FIM244_ABA('HISTORICO_STATUS_VINCULOS');
    abaConfig = DCP_FIM244_ABA('CONFIG');
    const abaAuditoria = DCP_FIM244_ABA('AUDITORIA');

    DCP_FIM244_VALIDAR_CABECALHOS(
      abaVinculos,
      DCP_ENCERRAMENTO_VINCULOS_244.CABECALHOS_VINCULOS
    );

    DCP_FIM244_VALIDAR_CABECALHOS(
      abaHistorico,
      DCP_ENCERRAMENTO_VINCULOS_244.CABECALHOS_HISTORICO
    );

    DCP_FIM244_VALIDAR_CABECALHOS(
      abaConfig,
      DCP_ENCERRAMENTO_VINCULOS_244.CABECALHOS_CONFIG
    );

    DCP_FIM244_VALIDAR_CABECALHOS(
      abaAuditoria,
      DCP_ENCERRAMENTO_VINCULOS_244.CABECALHOS_AUDITORIA
    );

    /*
     * 1. IDEMPOTENCIA
     */

    const anterior = DCP_FIM244_LER_IDEMPOTENCIA(
      abaConfig,
      solicitacao
    );

    if (anterior) {

      if (anterior.estado === 'CONCLUIDO') {

        if (!anterior.resposta || !anterior.resposta.sucesso) {
          return DCP_FIM244_RESPOSTA_ERRO(
            'CONCLUIDO_INCONSISTENTE',
            'Idempotencia marcada como concluida sem resposta valida. Diagnostico necessario.',
            true,
            true
          );
        }

        return anterior.resposta;
      }

      return DCP_FIM244_RESPOSTA_ERRO(
        'TENTATIVA_INCONCLUSIVA',
        'Existe tentativa anterior. Execute o diagnostico antes de repetir.',
        true,
        true
      );
    }

    /*
     * 2. LOCALIZAR VINCULO
     */

    const encontrado = DCP_FIM244_LOCALIZAR_VINCULO(
      abaVinculos,
      solicitacao.identificador
    );

    if (!encontrado) {
      return DCP_FIM244_RESPOSTA_ERRO(
        'VINCULO_NAO_ENCONTRADO',
        'Vinculo nao encontrado.',
        false,
        false
      );
    }

    const linhaVinculo = encontrado.linha;
    const registroOriginal = encontrado.dados;

    const vinculoID = DCP_FIM244_TEXTO(registroOriginal[0]);
    const codigoVinculo = DCP_FIM244_TEXTO(registroOriginal[1]);

    if (!vinculoID || !codigoVinculo) {
      throw new Error('Identificadores do vinculo incompletos.');
    }

    if (
      DCP_FIM244_TEXTO(registroOriginal[11]).toUpperCase() !== 'ATIVO'
    ) {
      return DCP_FIM244_RESPOSTA_ERRO(
        'VINCULO_NAO_ATIVO',
        'Somente vinculos ATIVOS podem ser encerrados.',
        false,
        false
      );
    }

    if (registroOriginal[6] !== '' && registroOriginal[6] != null) {
      return DCP_FIM244_RESPOSTA_ERRO(
        'VINCULO_COM_FIM_EXISTENTE',
        'O vinculo ja possui data de encerramento.',
        false,
        false
      );
    }

    const inicio = DCP_FIM244_DATA_EXISTENTE(
      registroOriginal[5],
      'InicioEm'
    );

    if (solicitacao.dataHora.getTime() < inicio.getTime()) {
      return DCP_FIM244_RESPOSTA_ERRO(
        'DATA_ANTERIOR_AO_INICIO',
        'O encerramento nao pode ser anterior ao inicio do vinculo.',
        false,
        false
      );
    }

    /*
     * 3. IDENTIFICADORES E HISTORICO
     */

    const historicosExistentes = DCP_FIM244_LOCALIZAR_HISTORICOS(
      abaHistorico,
      solicitacao.chaveTentativa
    );

    if (historicosExistentes.length > 0) {
      return DCP_FIM244_RESPOSTA_ERRO(
        'HISTORICO_EXISTENTE',
        'Ja existe marcador de historico para esta tentativa. Diagnostico necessario.',
        true,
        true
      );
    }

    identificadores = {
      historicoID: DCP_FIM244_UUID(),
      codigoHistorico: DCP_FIM244_GERAR_CODIGO(
        abaConfig,
        abaHistorico
      ),
      auditoriaID: DCP_FIM244_UUID(),
      vinculoID: vinculoID,
      codigoVinculo: codigoVinculo
    };

    /*
     * 4. INICIAR PERSISTENCIA
     *
     * IMPORTANTE:
     * A sinalizacao ocorre ANTES da primeira gravacao.
     * Se a gravacao falhar parcialmente, o catch nao
     * classificara a operacao como simples validacao.
     */

    gravacaoIniciada = true;

    DCP_FIM244_SALVAR_IDEMPOTENCIA(
      abaConfig,
      solicitacao,
      hashSolicitacao,
      'PENDENTE',
      null,
      identificadores
    );

    /*
     * 5. HISTORICO PENDENTE
     */

    const marcador =
      '[DCP_FIM244_' + solicitacao.chaveTentativa + ']';

    const observacoesHistorico = [
      marcador,
      solicitacao.observacoes
    ].filter(Boolean).join(' | ');

    const historico = [
      identificadores.historicoID,
      identificadores.codigoHistorico,
      vinculoID,
      codigoVinculo,
      'ATIVO',
      'ENCERRADO',
      solicitacao.motivo,
      solicitacao.dataHora,
      DCP_ENCERRAMENTO_VINCULOS_244.USUARIO,
      observacoesHistorico,
      'PENDENTE',
      identificadores.auditoriaID
    ];

    abaHistorico.appendRow(historico);
    SpreadsheetApp.flush();

    const linhasHistorico = DCP_FIM244_LOCALIZAR_HISTORICOS(
      abaHistorico,
      solicitacao.chaveTentativa
    );

    if (linhasHistorico.length !== 1) {
      throw new Error(
        'Historico nao foi localizado de forma unica.'
      );
    }

    linhaHistorico = linhasHistorico[0];

    /*
     * 6. ATUALIZAR VINCULO
     */

    const agora = new Date();
    const novoRegistro = registroOriginal.slice();

    novoRegistro[6] = solicitacao.dataHora;
    novoRegistro[11] = 'ENCERRADO';
    novoRegistro[13] = agora;
    novoRegistro[15] = DCP_ENCERRAMENTO_VINCULOS_244.USUARIO;

    abaVinculos
      .getRange(linhaVinculo, 1, 1, 16)
      .setValues([novoRegistro]);

    SpreadsheetApp.flush();

    /*
     * 7. CONFERIR VINCULO
     */

    const conferencia = abaVinculos
      .getRange(linhaVinculo, 1, 1, 16)
      .getValues()[0];

    if (
      DCP_FIM244_TEXTO(conferencia[0]) !== vinculoID ||
      DCP_FIM244_TEXTO(conferencia[1]) !== codigoVinculo ||
      DCP_FIM244_TEXTO(conferencia[11]) !== 'ENCERRADO' ||
      !(conferencia[6] instanceof Date) ||
      conferencia[6].getTime() !== solicitacao.dataHora.getTime()
    ) {
      throw new Error('A conferencia do vinculo falhou.');
    }

    /*
     * 8. AUDITORIA
     */

    const detalhes = [
      marcador,
      'Encerramento de ' + codigoVinculo,
      'Modo: ' + solicitacao.modo,
      'Motivo: ' + solicitacao.motivo,
      'DataHora: ' + solicitacao.dataHora.toISOString(),
      'Historico: ' + identificadores.codigoHistorico
    ].join('. ');

    DCP_FIM244_REGISTRAR_AUDITORIA(
      abaAuditoria,
      {
        auditoriaID: identificadores.auditoriaID,
        dataHora: agora,
        vinculoID: vinculoID,
        detalhes: detalhes
      }
    );

    /*
     * 9. CONFIRMAR HISTORICO
     */

    abaHistorico
      .getRange(linhaHistorico, 11)
      .setValue('SUCESSO');

    SpreadsheetApp.flush();

    /*
     * 10. CONFERENCIA INTEGRAL
     */

    const consistencia = DCP_FIM244_CONFERIR_ESTADO(
      abaVinculos,
      abaHistorico,
      abaAuditoria,
      solicitacao,
      identificadores
    );

    if (!consistencia.consistente) {
      throw new Error(
        'Conferencia integral falhou: ' +
        consistencia.problemas.join(' | ')
      );
    }

    /*
     * 11. CONCLUIR IDEMPOTENCIA
     */

    const resposta = {
      sucesso: true,
      codigoVinculo: codigoVinculo,
      vinculoID: vinculoID,
      statusAnterior: 'ATIVO',
      statusNovo: 'ENCERRADO',
      modoEncerramento: solicitacao.modo,
      dataHoraEncerramento: solicitacao.dataHora.toISOString(),
      motivo: solicitacao.motivo,
      codigoHistorico: identificadores.codigoHistorico,
      historicoID: identificadores.historicoID,
      auditoriaID: identificadores.auditoriaID,
      auditoriaPendente: false,
      mensagem: 'Vinculo encerrado, conferido e auditado.'
    };

    DCP_FIM244_SALVAR_IDEMPOTENCIA(
      abaConfig,
      solicitacao,
      hashSolicitacao,
      'CONCLUIDO',
      resposta,
      identificadores
    );

    SpreadsheetApp.flush();

    return resposta;

  } catch (erro) {

    /*
     * FALHA ANTES DA PERSISTENCIA
     */

    if (!gravacaoIniciada) {
      return DCP_FIM244_RESPOSTA_ERRO(
        'ERRO_VALIDACAO',
        erro.message,
        false,
        false
      );
    }

    /*
     * FALHA APOS INICIO DA PERSISTENCIA
     *
     * Nao excluir nem reverter automaticamente.
     */

    let estadoReal = null;

    try {

      if (
        abaVinculos &&
        abaHistorico &&
        abaConfig
      ) {

        const vinculoID = identificadores.vinculoID;
        const codigoVinculo = identificadores.codigoVinculo;

        const linhasHistorico = DCP_FIM244_LOCALIZAR_HISTORICOS(
          abaHistorico,
          solicitacao.chaveTentativa
        );

        const historicoValido = linhasHistorico.length === 1;

        const vinculoEncontrado = codigoVinculo
          ? DCP_FIM244_LOCALIZAR_VINCULO(
              abaVinculos,
              codigoVinculo
            )
          : null;

        estadoReal = {
          historicoUnico: historicoValido,
          vinculoEncontrado: !!vinculoEncontrado,
          vinculoID: vinculoID || null
        };
      }

    } catch (erroConferencia) {

      console.error(
        'Falha na conferencia de recuperacao: ' +
        erroConferencia.message
      );

    }

    /*
     * Nao sobrescrever CONCLUIDO se a gravacao final
     * ja tiver sido realizada.
     */

    try {

      if (abaConfig && solicitacao) {

        const registroAtual = DCP_FIM244_LER_CONFIG(
          abaConfig,
          DCP_FIM244_CHAVE_IDEMPOTENCIA(
            solicitacao.chaveTentativa
          )
        );

        let dadosAtuais = null;

        if (registroAtual) {
          try {
            dadosAtuais = JSON.parse(registroAtual.valor);
          } catch (erroJSON) {
            dadosAtuais = null;
          }
        }

        if (
          !dadosAtuais ||
          dadosAtuais.estado !== 'CONCLUIDO'
        ) {

          DCP_FIM244_SALVAR_IDEMPOTENCIA(
            abaConfig,
            solicitacao,
            hashSolicitacao,
            'INCONCLUSIVO',
            {
              sucesso: false,
              mensagem: erro.message,
              conferencia: estadoReal
            },
            identificadores
          );

        }

      }

    } catch (erroIdempotencia) {

      console.error(
        'Falha ao atualizar idempotencia: ' +
        erroIdempotencia.message
      );

    }

    try {

      if (abaHistorico) {

        const linhas = DCP_FIM244_LOCALIZAR_HISTORICOS(
          abaHistorico,
          solicitacao.chaveTentativa
        );

        if (linhas.length === 1) {

          const statusAtual = DCP_FIM244_TEXTO(
            abaHistorico.getRange(linhas[0], 11).getValue()
          );

          if (statusAtual === 'PENDENTE') {
            abaHistorico
              .getRange(linhas[0], 11)
              .setValue('INCONCLUSIVO');
          }

        }

      }

    } catch (erroHistorico) {

      console.error(
        'Falha ao sinalizar historico: ' +
        erroHistorico.message
      );

    }

    return DCP_FIM244_RESPOSTA_ERRO(
      'ENCERRAMENTO_INCONCLUSIVO',
      'A operacao iniciou gravacoes, mas nao foi possivel ' +
      'confirmar todas as etapas. Nao repita. ' +
      'Execute o diagnostico de recuperacao. Detalhe: ' +
      erro.message,
      true,
      true
    );

  } finally {

    lock.releaseLock();

  }

}


/* =====================================================
   DIAGNOSTICO DE RECUPERACAO - SOMENTE LEITURA
===================================================== */

function DCP_DIAGNOSTICAR_RECUPERACAO_ENCERRAMENTOS_244() {

  const config = DCP_FIM244_ABA('CONFIG');
  const vinculos = DCP_FIM244_ABA('VINCULOS');
  const historico = DCP_FIM244_ABA('HISTORICO_STATUS_VINCULOS');
  const auditoria = DCP_FIM244_ABA('AUDITORIA');

  const resultados = [];
  const prefixo = DCP_ENCERRAMENTO_VINCULOS_244.PREFIXO_IDEMPOTENCIA;

  const ultimaLinha = config.getLastRow();

  const ocorrencias = {};

  if (ultimaLinha >= 2) {

    const linhas = config
      .getRange(2, 1, ultimaLinha - 1, 2)
      .getValues();

    linhas.forEach(function(item, indice) {

      const chave = DCP_FIM244_TEXTO(item[0]);

      if (chave.indexOf(prefixo) !== 0) return;

      ocorrencias[chave] = (ocorrencias[chave] || 0) + 1;

      let dados;

      try {
        dados = typeof item[1] === 'string'
          ? JSON.parse(item[1])
          : item[1];
      } catch (erro) {

        resultados.push({
          chave: chave,
          estado: 'ILEGIVEL',
          problema: 'JSON da idempotencia invalido.',
          linhaConfig: indice + 2
        });

        return;
      }

      if (!dados || !dados.estado) {

        resultados.push({
          chave: chave,
          estado: 'ILEGIVEL',
          problema: 'Estrutura de idempotencia incompleta.',
          linhaConfig: indice + 2
        });

        return;
      }

      const ids = dados.identificadores || {};

      const chaveTentativa = chave.substring(prefixo.length);

      const linhasHistorico = DCP_FIM244_LOCALIZAR_HISTORICOS(
        historico,
        chaveTentativa
      );

      const auditorias = ids.auditoriaID
        ? DCP_FIM244_AUDITORIAS_EXISTENTES(
            auditoria,
            ids.auditoriaID
          )
        : [];

      let vinculo = null;

      if (ids.codigoVinculo) {
        try {
          vinculo = DCP_FIM244_LOCALIZAR_VINCULO(
            vinculos,
            ids.codigoVinculo
          );
        } catch (erroVinculo) {
          resultados.push({
            chave: chave,
            estado: dados.estado,
            problema: erroVinculo.message,
            linhaConfig: indice + 2
          });
          return;
        }
      }

      const estadoVinculo = vinculo
        ? DCP_FIM244_TEXTO(vinculo.dados[11])
        : 'NAO_LOCALIZADO';

      const estadoHistorico = linhasHistorico.length === 1
        ? DCP_FIM244_TEXTO(
            historico.getRange(linhasHistorico[0], 11).getValue()
          )
        : linhasHistorico.length > 1
          ? 'DUPLICADO'
          : 'NAO_LOCALIZADO';

      const problemas = [];

      if (ocorrencias[chave] > 1) {
        problemas.push('Chave duplicada na CONFIG.');
      }

      if (linhasHistorico.length > 1) {
        problemas.push('Marcador de historico duplicado.');
      }

      if (auditorias.length > 1) {
        problemas.push('Auditoria duplicada.');
      }

      if (
        dados.estado === 'CONCLUIDO' &&
        (
          estadoVinculo !== 'ENCERRADO' ||
          estadoHistorico !== 'SUCESSO' ||
          auditorias.length !== 1
        )
      ) {
        problemas.push(
          'Estado CONCLUIDO sem correspondencia integral.'
        );
      }

      if (
        dados.estado !== 'CONCLUIDO' &&
        dados.estado !== 'PENDENTE' &&
        dados.estado !== 'INCONCLUSIVO'
      ) {
        problemas.push('Estado desconhecido.');
      }

      if (
        dados.estado === 'PENDENTE' ||
        dados.estado === 'INCONCLUSIVO'
      ) {
        problemas.push('Tentativa exige revisao manual.');
      }

      resultados.push({
        chave: chave,
        estadoIdempotencia: dados.estado,
        codigoVinculo: ids.codigoVinculo || '',
        statusVinculo: estadoVinculo,
        quantidadeHistoricos: linhasHistorico.length,
        statusHistorico: estadoHistorico,
        quantidadeAuditorias: auditorias.length,
        problemas: problemas,
        consistente: problemas.length === 0,
        linhaConfig: indice + 2
      });

    });

  }

  return {
    modulo: '2.4.4.2',
    somenteLeitura: true,
    fusoProjeto: Session.getScriptTimeZone(),
    fusoPlanilha: SpreadsheetApp
      .getActiveSpreadsheet()
      .getSpreadsheetTimeZone(),
    sucesso: true,
    quantidadeAnalisada: resultados.length,
    quantidadeComProblemas: resultados.filter(function(item) {
      return item.consistente === false;
    }).length,
    resultados: resultados,
    mensagem: 'Diagnostico concluido. Nenhum registro foi alterado.'
  };
}


/* =====================================================
   DIAGNOSTICO ESTRUTURAL
===================================================== */

function DCP_DIAGNOSTICAR_MOTOR_ENCERRAMENTO_VINCULOS_244() {

  const planilha = SpreadsheetApp.getActiveSpreadsheet();

  const estruturas = [
    {
      nome: 'VINCULOS',
      cabecalhos: DCP_ENCERRAMENTO_VINCULOS_244.CABECALHOS_VINCULOS
    },
    {
      nome: 'HISTORICO_STATUS_VINCULOS',
      cabecalhos: DCP_ENCERRAMENTO_VINCULOS_244.CABECALHOS_HISTORICO
    },
    {
      nome: 'CONFIG',
      cabecalhos: DCP_ENCERRAMENTO_VINCULOS_244.CABECALHOS_CONFIG
    },
    {
      nome: 'AUDITORIA',
      cabecalhos: DCP_ENCERRAMENTO_VINCULOS_244.CABECALHOS_AUDITORIA
    }
  ];

  const resultados = estruturas.map(function(item) {

    const aba = planilha.getSheetByName(item.nome);

    if (!aba) {
      return {
        aba: item.nome,
        existe: false,
        cabecalhosCorretos: false
      };
    }

    const encontrados = aba
      .getRange(1, 1, 1, item.cabecalhos.length)
      .getDisplayValues()[0];

    const divergencias = [];

    item.cabecalhos.forEach(function(campo, indice) {

      if (String(encontrados[indice]).trim() !== campo) {
        divergencias.push({
          coluna: indice + 1,
          esperado: campo,
          encontrado: encontrados[indice] || ''
        });
      }

    });

    return {
      aba: item.nome,
      existe: true,
      cabecalhosCorretos: divergencias.length === 0,
      ultimaLinha: aba.getLastRow(),
      divergencias: divergencias
    };

  });

  const sucesso = resultados.every(function(item) {
    return item.existe && item.cabecalhosCorretos;
  });

  return {
    modulo: '2.4.4.2',
    somenteLeitura: true,
    sucesso: sucesso,
    estruturas: resultados,
    fusoProjeto: Session.getScriptTimeZone(),
    fusoPlanilha: planilha.getSpreadsheetTimeZone(),
    fusoEsperado: DCP_ENCERRAMENTO_VINCULOS_244.FUSO,
    mensagem: sucesso
      ? 'Estruturas verificadas.'
      : 'Existem estruturas ausentes ou divergentes.'
  };
}


/* =====================================================
   TESTES DE DATA
===================================================== */

function DCP_TESTAR_DATAS_ENCERRAMENTO_VINCULOS_244() {

  const exemplos = [
    {
      nome: 'DATA E HORA',
      valor: '2026-11-05T07:00',
      esperado: true
    },
    {
      nome: 'SEGUNDOS',
      valor: '2026-11-05T07:00:30',
      esperado: true
    },
    {
      nome: 'DATA INVALIDA',
      valor: '2026-15-45T28:90',
      esperado: false
    },
    {
      nome: 'FORMATO INVALIDO',
      valor: '05/11/2026',
      esperado: false
    },
    {
      nome: 'DIA INEXISTENTE',
      valor: '2026-02-30T10:00',
      esperado: false
    }
  ];

  const resultados = exemplos.map(function(item) {

    let valido = false;
    let mensagem = '';

    try {
      DCP_FIM244_CONVERTER_DATA_LOCAL(item.valor);
      valido = true;
    } catch (erro) {
      mensagem = erro.message;
    }

    return {
      teste: item.nome,
      esperado: item.esperado,
      resultado: valido,
      aprovado: valido === item.esperado,
      erro: mensagem
    };

  });

  return {
    modulo: '2.4.4.2',
    somenteLeitura: true,
    fusoUtilizado: DCP_ENCERRAMENTO_VINCULOS_244.FUSO,
    sucesso: resultados.every(function(item) {
      return item.aprovado;
    }),
    resultados: resultados
  };
}


/* =====================================================
   TESTES DE IDEMPOTENCIA - SOMENTE MEMORIA
===================================================== */

function DCP_TESTAR_IDEMPOTENCIA_ENCERRAMENTOS_244() {

  const base = {
    identificador: 'VIN-TESTE',
    modo: 'INFORMADO',
    motivo: 'Teste isolado',
    observacoes: '',
    dataHora: new Date('2026-11-05T10:00:00.000Z')
  };

  const repetida = {
    identificador: 'VIN-TESTE',
    modo: 'INFORMADO',
    motivo: 'Teste isolado',
    observacoes: '',
    dataHora: new Date('2026-11-05T10:00:00.000Z')
  };

  const diferente = {
    identificador: 'VIN-TESTE',
    modo: 'INFORMADO',
    motivo: 'Outro motivo',
    observacoes: '',
    dataHora: new Date('2026-11-05T10:00:00.000Z')
  };

  const imediato1 = {
    identificador: 'VIN-TESTE',
    modo: 'IMEDIATO',
    motivo: 'Teste imediato',
    observacoes: '',
    dataHora: new Date('2026-11-05T10:00:00.000Z')
  };

  const imediato2 = {
    identificador: 'VIN-TESTE',
    modo: 'IMEDIATO',
    motivo: 'Teste imediato',
    observacoes: '',
    dataHora: new Date('2026-11-05T11:00:00.000Z')
  };

  const hash1 = DCP_FIM244_HASH_SOLICITACAO(base);
  const hash2 = DCP_FIM244_HASH_SOLICITACAO(repetida);
  const hash3 = DCP_FIM244_HASH_SOLICITACAO(diferente);

  const hashImediato1 = DCP_FIM244_HASH_SOLICITACAO(imediato1);
  const hashImediato2 = DCP_FIM244_HASH_SOLICITACAO(imediato2);

  const resultados = [
    {
      teste: 'MESMA SOLICITACAO',
      aprovado: hash1 === hash2
    },
    {
      teste: 'MOTIVO DIFERENTE',
      aprovado: hash1 !== hash3
    },
    {
      teste: 'HORA GERADA NO MODO IMEDIATO',
      aprovado: hashImediato1 === hashImediato2
    }
  ];

  return {
    modulo: '2.4.4.2',
    somenteMemoria: true,
    sucesso: resultados.every(function(item) {
      return item.aprovado;
    }),
    resultados: resultados
  };
}


/* =====================================================
   FUNCOES DE VISUALIZACAO
===================================================== */

function DCP_DIAGNOSTICO_VISIVEL_ENCERRAMENTO_244() {

  const resultado =
    DCP_DIAGNOSTICAR_MOTOR_ENCERRAMENTO_VINCULOS_244();

  console.log(
    '=== DIAGNOSTICO ESTRUTURAL 2.4.4.2 ==='
  );

  console.log(JSON.stringify(resultado, null, 2));

  return resultado;
}


function DCP_TESTE_DATAS_VISIVEL_244() {

  const resultado =
    DCP_TESTAR_DATAS_ENCERRAMENTO_VINCULOS_244();

  console.log(
    '=== TESTE DE DATAS 2.4.4.2 ==='
  );

  console.log(JSON.stringify(resultado, null, 2));

  return resultado;
}


function DCP_TESTE_IDEMPOTENCIA_VISIVEL_244() {

  const resultado =
    DCP_TESTAR_IDEMPOTENCIA_ENCERRAMENTOS_244();

  console.log(
    '=== TESTE DE IDEMPOTENCIA 2.4.4.2 ==='
  );

  console.log(JSON.stringify(resultado, null, 2));

  return resultado;
}


function DCP_DIAGNOSTICO_RECUPERACAO_VISIVEL_244() {

  const resultado =
    DCP_DIAGNOSTICAR_RECUPERACAO_ENCERRAMENTOS_244();

  console.log(
    '=== DIAGNOSTICO DE RECUPERACAO 2.4.4.2 ==='
  );

  console.log(JSON.stringify(resultado, null, 2));

  return resultado;
}