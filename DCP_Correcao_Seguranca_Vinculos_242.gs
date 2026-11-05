
/**
 * DRIVER CONTROL PRO
 * CORRECOES DE SEGURANCA DO MODULO 2.4.2
 *
 * Aplicar junto ao arquivo DCP_Cadastro_Vinculos_242.gs.
 *
 * Corrige:
 * 1. Validacoes incompatíveis na CONFIG para chaves IDEMP_VIN.
 * 2. Comparacao exata dos horarios CriadoEm e AtualizadoEm.
 * 3. Consulta de idempotencia com comportamento conservador.
 *
 * Nao realiza cadastro real.
 */


/* =====================================================
   1. RESERVA DE IDEMPOTENCIA NA CONFIG
===================================================== */

/**
 * Reserva uma linha técnica na CONFIG.
 *
 * A validacao de dados e removida somente das celulas
 * A e C da linha destinada a chave e ao tipo JSON.
 *
 * As demais colunas e linhas permanecem intactas.
 */
function DCP_CV242_CRIAR_TENTATIVA(abaConfig, chave, dadosResumo) {
  const chaveConfig =
    DCP_CADASTRO_VINCULOS_242.PREFIXO_IDEMPOTENCIA + chave;

  const anterior = DCP_CV242_LOCALIZAR_TENTATIVA(
    abaConfig,
    chave
  );

  if (anterior) {
    throw new Error(
      'A chave de tentativa ja existe na CONFIG.'
    );
  }

  const ultimaLinha = abaConfig.getLastRow();
  const linhaDestino = ultimaLinha + 1;

  // Confere se a linha de destino realmente esta vazia.
  const valoresDestino = abaConfig
    .getRange(linhaDestino, 1, 1, 6)
    .getValues()[0];

  const linhaOcupada = valoresDestino.some(function(valor) {
    return valor !== '' && valor !== null;
  });

  if (linhaOcupada) {
    throw new Error(
      'A proxima linha da CONFIG nao esta vazia. Reserva bloqueada.'
    );
  }

  // Remove somente as validacoes incompatíveis.
  abaConfig.getRange(linhaDestino, 1).clearDataValidations();
  abaConfig.getRange(linhaDestino, 3).clearDataValidations();

  const agora = new Date();

  const registro = {
    estado: 'PROCESSANDO',
    criadoEm: agora.toISOString(),
    resumo: dadosResumo
  };

  const linha = [
    chaveConfig,
    JSON.stringify(registro),
    'JSON',
    'Controle de idempotencia de cadastro de vinculo.',
    false,
    agora
  ];

  // Grava na linha previamente inspecionada.
  abaConfig.getRange(
    linhaDestino,
    1,
    1,
    6
  ).setValues([linha]);

  SpreadsheetApp.flush();

  // Confirma a gravacao pela chave e pela linha.
  const localizado = DCP_CV242_LOCALIZAR_TENTATIVA(
    abaConfig,
    chave
  );

  if (
    !localizado ||
    localizado.linha !== linhaDestino
  ) {
    throw new Error(
      'A reserva foi iniciada, mas nao foi possivel confirmar sua gravacao.'
    );
  }

  let registroConfirmado;

  try {
    registroConfirmado = JSON.parse(localizado.valor);
  } catch (erro) {
    throw new Error(
      'A reserva foi gravada, mas o JSON nao pode ser interpretado.'
    );
  }

  if (
    registroConfirmado.estado !== 'PROCESSANDO' ||
    registroConfirmado.criadoEm !== agora.toISOString()
  ) {
    throw new Error(
      'A reserva gravada diverge dos dados esperados.'
    );
  }

  return localizado.linha;
}


/* =====================================================
   2. CONFERENCIA EXATA DOS 16 CAMPOS
===================================================== */

/**
 * Confere os campos funcionais e os metadados.
 *
 * O objeto esperado agora precisa conter criadoEm.
 * CriadoEm e AtualizadoEm devem ser exatamente iguais
 * ao horario usado na gravacao inicial.
 */
function DCP_CV242_REGISTRO_COMPLETO(registro, esperado) {
  if (!registro || !registro.dados || !esperado) {
    return false;
  }

  const l = registro.dados;

  const camposTexto = [
    [0, esperado.vinculoID],
    [1, esperado.codigo],
    [2, esperado.motoristaID],
    [3, esperado.veiculoID],
    [4, esperado.tipo],
    [9, esperado.periodicidade],
    [10, esperado.observacoes],
    [11, 'ATIVO'],
    [14, DCP_CADASTRO_VINCULOS_242.USUARIO],
    [15, DCP_CADASTRO_VINCULOS_242.USUARIO]
  ];

  for (let i = 0; i < camposTexto.length; i++) {
    const coluna = camposTexto[i][0];
    const valor = camposTexto[i][1];

    const gravado = String(
      l[coluna] === null || l[coluna] === undefined
        ? ''
        : l[coluna]
    ).trim();

    const previsto = String(
      valor === null || valor === undefined
        ? ''
        : valor
    ).trim();

    if (gravado !== previsto) {
      return false;
    }
  }

  if (!DCP_CV242_DATA_IGUAL(l[5], esperado.inicio)) {
    return false;
  }

  if (!DCP_CV242_DATA_IGUAL(l[6], esperado.fim)) {
    return false;
  }

  if (
    DCP_CV242_BOOLEANO(l[7]) !==
    esperado.compartilhamentoAutorizado
  ) {
    return false;
  }

  const aluguelGravado =
    l[8] === '' || l[8] === null || l[8] === undefined
      ? ''
      : Number(l[8]);

  if (aluguelGravado !== esperado.valorAluguel) {
    return false;
  }

  // Comparacao exata dos dois metadados de horario.
  if (!esperado.criadoEm) {
    return false;
  }

  if (!DCP_CV242_DATA_IGUAL(l[12], esperado.criadoEm)) {
    return false;
  }

  if (!DCP_CV242_DATA_IGUAL(l[13], esperado.criadoEm)) {
    return false;
  }

  return true;
}


/* =====================================================
   3. DIAGNOSTICO DE VALIDACOES DA CONFIG
===================================================== */

/**
 * Inspeciona a proxima linha da CONFIG.
 * Somente leitura. Nao altera validacoes ou dados.
 */
function DCP_CV242_DIAGNOSTICAR_CONFIG_DESTINO() {
  const aba = DCP_CV242_ABA('CONFIG');

  DCP_CV242_CONFERIR_CABECALHOS(
    'CONFIG',
    DCP_CADASTRO_VINCULOS_242.CABECALHOS.CONFIG
  );

  const linhaDestino = aba.getLastRow() + 1;

  const colunas = [
    { nome: 'Chave', coluna: 1 },
    { nome: 'Valor', coluna: 2 },
    { nome: 'TipoValor', coluna: 3 },
    { nome: 'Descricao', coluna: 4 },
    { nome: 'Editavel', coluna: 5 },
    { nome: 'AtualizadoEm', coluna: 6 }
  ];

  const resultado = {
    sucesso: true,
    somenteLeitura: true,
    dadosAlterados: false,
    linhaDestino: linhaDestino,
    colunas: []
  };

  colunas.forEach(function(item) {
    const celula = aba.getRange(
      linhaDestino,
      item.coluna
    );

    const regra = celula.getDataValidation();

    resultado.colunas.push({
      coluna: item.nome,
      possuiValidacao: Boolean(regra),
      tipoValidacao: regra
        ? String(regra.getCriteriaType())
        : 'SEM_VALIDACAO'
    });
  });

  const valores = aba.getRange(
    linhaDestino,
    1,
    1,
    6
  ).getValues()[0];

  resultado.linhaVazia = valores.every(function(valor) {
    return valor === '' || valor === null;
  });

  Logger.log(JSON.stringify(resultado, null, 2));
  return resultado;
}


/* =====================================================
   4. TESTE DA COMPARACAO DE HORARIOS
===================================================== */

/**
 * Testa somente em memoria.
 * Nao altera planilhas.
 */
function DCP_CV242_TESTAR_COMPARACAO_HORARIOS() {
  const agora = new Date('2026-11-04T10:30:00.000Z');
  const umSegundoDepois = new Date(
    '2026-11-04T10:30:01.000Z'
  );

  const base = {
    vinculoID: 'UUID-TESTE',
    codigo: 'VIN-TESTE',
    motoristaID: 'MOT-TESTE',
    veiculoID: 'VEI-TESTE',
    tipo: 'PROPRIO',
    inicio: new Date('2026-11-04T07:00:00.000Z'),
    fim: new Date('2026-11-04T12:00:00.000Z'),
    compartilhamentoAutorizado: true,
    valorAluguel: '',
    periodicidade: '',
    observacoes: 'Teste',
    criadoEm: agora
  };

  function criarLinha(criadoEm, atualizadoEm) {
    return [
      base.vinculoID,
      base.codigo,
      base.motoristaID,
      base.veiculoID,
      base.tipo,
      base.inicio,
      base.fim,
      true,
      '',
      '',
      base.observacoes,
      'ATIVO',
      criadoEm,
      atualizadoEm,
      DCP_CADASTRO_VINCULOS_242.USUARIO,
      DCP_CADASTRO_VINCULOS_242.USUARIO
    ];
  }

  const casos = [
    {
      nome: 'Horarios exatamente iguais',
      linha: criarLinha(agora, agora),
      esperado: true
    },
    {
      nome: 'CriadoEm divergente',
      linha: criarLinha(umSegundoDepois, agora),
      esperado: false
    },
    {
      nome: 'AtualizadoEm divergente',
      linha: criarLinha(agora, umSegundoDepois),
      esperado: false
    },
    {
      nome: 'CriadoEm ausente',
      linha: criarLinha('', agora),
      esperado: false
    },
    {
      nome: 'AtualizadoEm ausente',
      linha: criarLinha(agora, ''),
      esperado: false
    }
  ];

  const resultados = casos.map(function(caso) {
    const registro = {
      linha: 2,
      dados: caso.linha
    };

    const obtido = DCP_CV242_REGISTRO_COMPLETO(
      registro,
      base
    );

    return {
      teste: caso.nome,
      esperado: caso.esperado,
      obtido: obtido,
      aprovado: obtido === caso.esperado
    };
  });

  const aprovados = resultados.filter(function(item) {
    return item.aprovado;
  }).length;

  const resposta = {
    total: resultados.length,
    aprovados: aprovados,
    reprovados: resultados.length - aprovados,
    resultados: resultados,
    somenteMemoria: true,
    planilhasAlteradas: false
  };

  Logger.log(JSON.stringify(resposta, null, 2));
  return resposta;
}
