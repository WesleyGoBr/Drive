
/**
 * DRIVER CONTROL PRO
 * MODULO 2.4.3 - INTEGRACAO DO FORMULARIO DE VINCULOS
 *
 * Responsabilidades:
 * 1. Carregar motoristas e veiculos ativos para o formulario.
 * 2. Entregar os dados no formato esperado pelo HTML.
 * 3. Encaminhar o cadastro para DCP_CADASTRAR_VINCULO_242.
 *
 * SEGURANCA:
 * - O carregamento e somente leitura.
 * - Nenhum identificador ou codigo e reservado neste modulo.
 * - Nenhuma linha e gravada ou alterada neste modulo.
 * - O cadastro real permanece sob responsabilidade do modulo 2.4.2.
 * - A resposta do cadastro principal e preservada integralmente.
 */

const DCP_INTEGRACAO_FORMULARIO_VINCULOS_243 = {
  ABAS: {
    MOTORISTAS: 'MOTORISTAS',
    VEICULOS: 'VEICULOS'
  },

  CABECALHOS: {
    MOTORISTAS: [
      'MotoristaID',
      'CodigoMotorista',
      'NomeCompleto',
      'NomeExibicao',
      'Telefone',
      'Email',
      'DataCadastro',
      'Observacoes',
      'Status',
      'CriadoEm',
      'AtualizadoEm',
      'CriadoPor',
      'AtualizadoPor'
    ],

    VEICULOS: [
      'VeiculoID',
      'CodigoVeiculo',
      'Placa',
      'Marca',
      'Modelo',
      'AnoFabricacao',
      'AnoModelo',
      'TipoVeiculo',
      'TipoPropulsao',
      'CombustivelPrincipal',
      'CombustivelSecundario',
      'ProprietarioTipo',
      'ProprietarioMotoristaID',
      'ProprietarioNome',
      'CapacidadeTanqueLitros',
      'CapacidadeBateriaKWh',
      'HodometroAtual',
      'DataCadastro',
      'Observacoes',
      'Status',
      'CriadoEm',
      'AtualizadoEm',
      'CriadoPor',
      'AtualizadoPor'
    ]
  },

  TEMPO_LOCK_MS: 15000
};


/* =====================================================
   1. AUXILIARES
===================================================== */

function DCP_F243_NORMALIZAR_TEXTO(valor) {
  return String(
    valor === null || valor === undefined ? '' : valor
  ).trim();
}


function DCP_F243_ABA(nome) {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();

  if (!planilha) {
    throw new Error(
      'Nao foi possivel acessar a planilha ativa.'
    );
  }

  const aba = planilha.getSheetByName(nome);

  if (!aba) {
    throw new Error(
      'A aba ' + nome + ' nao foi encontrada.'
    );
  }

  return aba;
}


function DCP_F243_CONFERIR_CABECALHOS(nomeAba, esperados) {
  const aba = DCP_F243_ABA(nomeAba);

  if (aba.getLastColumn() < esperados.length) {
    throw new Error(
      'A aba ' + nomeAba +
      ' nao possui todas as colunas esperadas.'
    );
  }

  const encontrados = aba.getRange(
    1,
    1,
    1,
    esperados.length
  ).getValues()[0];

  const divergencias = [];

  esperados.forEach(function(nome, indice) {
    const encontrado = DCP_F243_NORMALIZAR_TEXTO(
      encontrados[indice]
    );

    if (encontrado !== nome) {
      divergencias.push({
        coluna: indice + 1,
        esperado: nome,
        encontrado: encontrado
      });
    }
  });

  if (divergencias.length) {
    throw new Error(
      'Cabecalhos divergentes na aba ' + nomeAba +
      ': ' + JSON.stringify(divergencias)
    );
  }

  return aba;
}


function DCP_F243_COMPARAR_CODIGOS(a, b) {
  return String(a.codigo || '').localeCompare(
    String(b.codigo || ''),
    'pt-BR',
    {
      numeric: true,
      sensitivity: 'base'
    }
  );
}


/* =====================================================
   2. LISTAGEM DE MOTORISTAS ATIVOS
===================================================== */

function DCP_F243_LISTAR_MOTORISTAS_ATIVOS(aba) {
  const ultimaLinha = aba.getLastRow();

  if (ultimaLinha < 2) {
    return [];
  }

  const dados = aba.getRange(
    2,
    1,
    ultimaLinha - 1,
    13
  ).getValues();

  const resultado = [];

  dados.forEach(function(linha, indice) {
    const id = DCP_F243_NORMALIZAR_TEXTO(linha[0]);
    const codigo = DCP_F243_NORMALIZAR_TEXTO(linha[1]);
    const nome = DCP_F243_NORMALIZAR_TEXTO(linha[2]);
    const status = DCP_F243_NORMALIZAR_TEXTO(
      linha[8]
    ).toUpperCase();

    if (status !== 'ATIVO') {
      return;
    }

    if (!id || !codigo || !nome) {
      throw new Error(
        'Motorista ativo com dados obrigatorios incompletos ' +
        'na linha ' + (indice + 2) + '.'
      );
    }

    resultado.push({
      id: id,
      codigo: codigo,
      nome: nome
    });
  });

  resultado.sort(DCP_F243_COMPARAR_CODIGOS);

  return resultado;
}


/* =====================================================
   3. LISTAGEM DE VEICULOS ATIVOS
===================================================== */

function DCP_F243_LISTAR_VEICULOS_ATIVOS(aba) {
  const ultimaLinha = aba.getLastRow();

  if (ultimaLinha < 2) {
    return [];
  }

  const dados = aba.getRange(
    2,
    1,
    ultimaLinha - 1,
    24
  ).getValues();

  const resultado = [];

  dados.forEach(function(linha, indice) {
    const id = DCP_F243_NORMALIZAR_TEXTO(linha[0]);
    const codigo = DCP_F243_NORMALIZAR_TEXTO(linha[1]);
    const placa = DCP_F243_NORMALIZAR_TEXTO(linha[2]);
    const marca = DCP_F243_NORMALIZAR_TEXTO(linha[3]);
    const modelo = DCP_F243_NORMALIZAR_TEXTO(linha[4]);

    const proprietarioTipo =
      DCP_F243_NORMALIZAR_TEXTO(linha[11]).toUpperCase();

    const proprietarioMotoristaID =
      DCP_F243_NORMALIZAR_TEXTO(linha[12]);

    const proprietarioNome =
      DCP_F243_NORMALIZAR_TEXTO(linha[13]);

    const status = DCP_F243_NORMALIZAR_TEXTO(
      linha[19]
    ).toUpperCase();

    if (status !== 'ATIVO') {
      return;
    }

    if (!id || !codigo || !placa) {
      throw new Error(
        'Veiculo ativo com dados obrigatorios incompletos ' +
        'na linha ' + (indice + 2) + '.'
      );
    }

    resultado.push({
      id: id,
      codigo: codigo,
      placa: placa,
      marca: marca,
      modelo: modelo,
      proprietarioTipo: proprietarioTipo,
      proprietarioMotoristaID: proprietarioMotoristaID,
      proprietarioNome: proprietarioNome
    });
  });

  resultado.sort(DCP_F243_COMPARAR_CODIGOS);

  return resultado;
}


/* =====================================================
   4. CARREGAMENTO DOS DADOS DO FORMULARIO
===================================================== */

/**
 * Funcao chamada pelo HTML:
 * DCP_FORM243_CARREGAR_DADOS()
 *
 * Operacao exclusivamente de leitura.
 */

function DCP_FORM243_CARREGAR_DADOS() {
  const lock = LockService.getDocumentLock();

  if (!lock.tryLock(
    DCP_INTEGRACAO_FORMULARIO_VINCULOS_243.TEMPO_LOCK_MS
  )) {
    return {
      sucesso: false,
      somenteLeitura: true,
      motoristas: [],
      veiculos: [],
      mensagem:
        'O sistema esta realizando outra operacao. ' +
        'Aguarde alguns instantes e atualize o formulario.'
    };
  }

  try {
    const abaMotoristas = DCP_F243_CONFERIR_CABECALHOS(
      DCP_INTEGRACAO_FORMULARIO_VINCULOS_243.ABAS.MOTORISTAS,
      DCP_INTEGRACAO_FORMULARIO_VINCULOS_243.CABECALHOS.MOTORISTAS
    );

    const abaVeiculos = DCP_F243_CONFERIR_CABECALHOS(
      DCP_INTEGRACAO_FORMULARIO_VINCULOS_243.ABAS.VEICULOS,
      DCP_INTEGRACAO_FORMULARIO_VINCULOS_243.CABECALHOS.VEICULOS
    );

    const motoristas =
      DCP_F243_LISTAR_MOTORISTAS_ATIVOS(abaMotoristas);

    const veiculos =
      DCP_F243_LISTAR_VEICULOS_ATIVOS(abaVeiculos);

    return {
      sucesso: true,
      somenteLeitura: true,
      motoristas: motoristas,
      veiculos: veiculos,
      totalMotoristas: motoristas.length,
      totalVeiculos: veiculos.length,
      mensagem:
        'Dados carregados. ' +
        motoristas.length + ' motorista(s) ativo(s) e ' +
        veiculos.length + ' veiculo(s) ativo(s).'
    };

  } catch (erro) {
    return {
      sucesso: false,
      somenteLeitura: true,
      motoristas: [],
      veiculos: [],
      mensagem:
        'Falha ao carregar os dados do formulario: ' +
        erro.message
    };

  } finally {
    lock.releaseLock();
  }
}


/* =====================================================
   5. ENCAMINHAMENTO DO CADASTRO
===================================================== */

/**
 * Funcao chamada pelo HTML:
 * DCP_FORM243_CADASTRAR_VINCULO(dados)
 *
 * Nao realiza gravacao diretamente.
 * Encaminha os dados ao modulo principal 2.4.2.
 *
 * A resposta original e preservada para que o HTML
 * consiga identificar sucesso, conflito, falha segura,
 * tentativa inconclusiva ou gravacao parcial.
 */

function DCP_FORM243_CADASTRAR_VINCULO(dados) {
  if (!dados || typeof dados !== 'object') {
    return {
      sucesso: false,
      gravacaoRealizada: false,
      tentativaInconclusiva: false,
      codigo: 'DADOS_INVALIDOS',
      mensagem:
        'Os dados do formulario nao foram recebidos corretamente.'
    };
  }

  if (
    typeof DCP_CADASTRAR_VINCULO_242 !== 'function'
  ) {
    return {
      sucesso: false,
      gravacaoRealizada: false,
      tentativaInconclusiva: false,
      codigo: 'MODULO_PRINCIPAL_INDISPONIVEL',
      mensagem:
        'A funcao principal DCP_CADASTRAR_VINCULO_242 ' +
        'nao foi encontrada. Nenhum cadastro foi encaminhado.'
    };
  }

  // Encaminhamento direto.
  // Nao alterar a chave de idempotencia nem os dados.
  return DCP_CADASTRAR_VINCULO_242(dados);
}


/* =====================================================
   6. DIAGNOSTICO DA INTEGRACAO
===================================================== */

/**
 * Diagnostico somente de leitura.
 * Nao cadastra vinculos e nao reserva codigos.
 */

function DCP_DIAGNOSTICAR_INTEGRACAO_FORMULARIO_243() {
  const resultado = {
    modulo: '2.4.3',
    somenteLeitura: true,
    funcoes: {},
    estrutura: {},
    sucesso: false
  };

  resultado.funcoes = {
    carregarDados:
      typeof DCP_FORM243_CARREGAR_DADOS === 'function',

    encaminharCadastro:
      typeof DCP_FORM243_CADASTRAR_VINCULO === 'function',

    cadastroPrincipal:
      typeof DCP_CADASTRAR_VINCULO_242 === 'function'
  };

  try {
    const abaMotoristas = DCP_F243_CONFERIR_CABECALHOS(
      DCP_INTEGRACAO_FORMULARIO_VINCULOS_243.ABAS.MOTORISTAS,
      DCP_INTEGRACAO_FORMULARIO_VINCULOS_243.CABECALHOS.MOTORISTAS
    );

    const abaVeiculos = DCP_F243_CONFERIR_CABECALHOS(
      DCP_INTEGRACAO_FORMULARIO_VINCULOS_243.ABAS.VEICULOS,
      DCP_INTEGRACAO_FORMULARIO_VINCULOS_243.CABECALHOS.VEICULOS
    );

    resultado.estrutura = {
      motoristas: {
        existe: true,
        cabecalhosCorretos: true,
        ultimaLinha: abaMotoristas.getLastRow()
      },
      veiculos: {
        existe: true,
        cabecalhosCorretos: true,
        ultimaLinha: abaVeiculos.getLastRow()
      }
    };

    resultado.sucesso =
      resultado.funcoes.carregarDados &&
      resultado.funcoes.encaminharCadastro &&
      resultado.funcoes.cadastroPrincipal;

  } catch (erro) {
    resultado.erro = erro.message;
  }

  Logger.log(JSON.stringify(resultado, null, 2));

  return resultado;
}


/* =====================================================
   7. TESTE DE CARREGAMENTO
===================================================== */

/**
 * Executa a listagem real em modo somente leitura.
 * Nao grava nem altera qualquer aba.
 */

function DCP_TESTAR_CARREGAMENTO_FORMULARIO_243() {
  const resposta = DCP_FORM243_CARREGAR_DADOS();

  const resultado = {
    sucesso: Boolean(resposta && resposta.sucesso),
    somenteLeitura: true,
    totalMotoristas: resposta
      ? resposta.totalMotoristas || 0
      : 0,
    totalVeiculos: resposta
      ? resposta.totalVeiculos || 0
      : 0,
    mensagem: resposta
      ? resposta.mensagem
      : 'Nenhuma resposta recebida.'
  };

  Logger.log(JSON.stringify(resultado, null, 2));

  return resultado;
}
