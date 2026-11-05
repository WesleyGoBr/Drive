const DCP_VEICULOS_235A = {
  ABA: 'VEICULOS',
  MOTORISTAS: 'MOTORISTAS',
  ADMIN: 'ADMIN_CENTRAL',

  TIPOS_PROPULSAO: [
    'COMBUSTAO',
    'FLEX',
    'HIBRIDO',
    'HIBRIDO_PLUGIN',
    'ELETRICO'
  ],

  TIPOS_PROPRIETARIO: [
    'MOTORISTA',
    'TERCEIRO'
  ]
};


/**
 * Normaliza placas antigas e Mercosul.
 */
function DCP_NORMALIZAR_PLACA_235A(valor) {
  return String(valor || '')
    .toUpperCase()
    .replace(/[\s-]/g, '');
}


/**
 * Valida os dados cadastrais de um veiculo.
 * Funcao somente de leitura.
 */
function DCP_VALIDAR_DADOS_VEICULO_235A(dados, veiculoIDIgnorado) {
  const erros = [];

  if (!dados || typeof dados !== 'object') {
    return {
      valido: false,
      erros: ['Dados do veiculo invalidos.']
    };
  }

  const placa = DCP_NORMALIZAR_PLACA_235A(dados.placa);
  const marca = String(dados.marca || '').trim();
  const modelo = String(dados.modelo || '').trim();

  if (!placa) {
    erros.push('Placa obrigatoria.');
  } else if (!/^[A-Z]{3}[0-9]{4}$/.test(placa) &&
             !/^[A-Z]{3}[0-9][A-Z][0-9]{2}$/.test(placa)) {
    erros.push('Formato de placa invalido.');
  }

  if (!marca) {
    erros.push('Marca obrigatoria.');
  } else if (marca.length > 80) {
    erros.push('Marca excede 80 caracteres.');
  }

  if (!modelo) {
    erros.push('Modelo obrigatorio.');
  } else if (modelo.length > 100) {
    erros.push('Modelo excede 100 caracteres.');
  }

  // Campos opcionais de ano.
  ['anoFabricacao', 'anoModelo'].forEach(campo => {
    const valor = dados[campo];

    if (valor === '' || valor === null || valor === undefined) {
      return;
    }

    const numero = Number(valor);

    if (
      !Number.isInteger(numero) ||
      numero < 1886 ||
      numero > 2100
    ) {
      erros.push(campo + ' deve ser um ano valido entre 1886 e 2100.');
    }
  });

  // Tipo de veiculo opcional.
  const tipoVeiculo = String(dados.tipoVeiculo || '').trim();

  if (tipoVeiculo.length > 50) {
    erros.push('Tipo de veiculo excede 50 caracteres.');
  }

  // Propulsao opcional, mas controlada quando informada.
  const propulsao = String(dados.tipoPropulsao || '')
    .trim()
    .toUpperCase();

  if (
    propulsao &&
    !DCP_VEICULOS_235A.TIPOS_PROPULSAO.includes(propulsao)
  ) {
    erros.push('Tipo de propulsao invalido.');
  }

  // Propriedade opcional, com consistencia entre os campos.
  const proprietarioTipo = String(dados.proprietarioTipo || '')
    .trim()
    .toUpperCase();

  const proprietarioMotoristaID = String(
    dados.proprietarioMotoristaID || ''
  ).trim();

  const proprietarioNome = String(
    dados.proprietarioNome || ''
  ).trim();

  if (
    proprietarioTipo &&
    !DCP_VEICULOS_235A.TIPOS_PROPRIETARIO.includes(proprietarioTipo)
  ) {
    erros.push('Tipo de proprietario invalido.');
  }

  if (proprietarioNome.length > 150) {
    erros.push('Nome do proprietario excede 150 caracteres.');
  }

  if (proprietarioTipo === 'MOTORISTA') {
    if (!proprietarioMotoristaID) {
      erros.push('Selecione o motorista proprietario.');
    } else {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const aba = ss.getSheetByName('MOTORISTAS');

      if (!aba) {
        erros.push('Aba MOTORISTAS nao encontrada.');
      } else {
        const registros = aba.getDataRange().getValues();

        const motorista = registros.find((linha, indice) =>
          indice > 0 &&
          String(linha[0]).trim() === proprietarioMotoristaID
        );

        if (!motorista) {
          erros.push('Motorista proprietario nao encontrado.');
        } else if (String(motorista[8]).trim() !== 'ATIVO') {
          erros.push('Motorista proprietario nao esta ativo.');
        }
      }
    }

    if (proprietarioNome) {
      erros.push(
        'O nome do proprietario motorista sera obtido automaticamente.'
      );
    }
  }

  if (proprietarioTipo === 'TERCEIRO') {
    if (!proprietarioNome) {
      erros.push('Nome do proprietario terceiro obrigatorio.');
    }

    if (proprietarioMotoristaID) {
      erros.push(
        'Terceiro externo nao deve possuir MotoristaID.'
      );
    }
  }

  // Capacidades e hodometro opcionais.
  [
    ['capacidadeTanqueLitros', 'Capacidade do tanque'],
    ['capacidadeBateriaKWh', 'Capacidade da bateria'],
    ['hodometroAtual', 'Hodometro']
  ].forEach(item => {
    const valor = dados[item[0]];

    if (valor === '' || valor === null || valor === undefined) {
      return;
    }

    const numero = Number(valor);

    if (!Number.isFinite(numero) || numero < 0) {
      erros.push(item[1] + ' deve ser um numero maior ou igual a zero.');
    }
  });

  const observacoes = String(dados.observacoes || '');

  if (observacoes.length > 2000) {
    erros.push('Observacoes excedem 2000 caracteres.');
  }

  // Verifica placa duplicada em registros ativos e inativos.
  if (
    placa &&
    (/^[A-Z]{3}[0-9]{4}$/.test(placa) ||
     /^[A-Z]{3}[0-9][A-Z][0-9]{2}$/.test(placa))
  ) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName('VEICULOS');

    if (!aba) {
      erros.push('Aba VEICULOS nao encontrada.');
    } else {
      const registros = aba.getDataRange().getValues();

      const duplicada = registros.some((linha, indice) => {
        if (indice === 0) return false;

        const placaExistente =
          DCP_NORMALIZAR_PLACA_235A(linha[2]);

        return (
          placaExistente === placa &&
          String(linha[0] || '').trim() !== String(veiculoIDIgnorado || '').trim()
        );
      });

      if (duplicada) {
        erros.push('Esta placa ja esta cadastrada.');
      }
    }
  }

  return {
    valido: erros.length === 0,
    placaNormalizada: placa,
    erros: erros
  };
}


/**
 * Testes sem gravacao de dados.
 */
function DCP_TESTAR_VALIDACOES_VEICULO_235A() {
  const testes = [
    {
      nome: 'Cadastro minimo valido',
      dados: {
        placa: 'ABC1D23',
        marca: 'Toyota',
        modelo: 'Corolla'
      },
      esperado: true
    },
    {
      nome: 'Placa antiga valida',
      dados: {
        placa: 'ABC-1234',
        marca: 'Fiat',
        modelo: 'Argo'
      },
      esperado: true
    },
    {
      nome: 'Placa vazia rejeitada',
      dados: {
        placa: '',
        marca: 'Toyota',
        modelo: 'Corolla'
      },
      esperado: false
    },
    {
      nome: 'Marca vazia rejeitada',
      dados: {
        placa: 'ABC1D23',
        marca: '',
        modelo: 'Corolla'
      },
      esperado: false
    },
    {
      nome: 'Modelo vazio rejeitado',
      dados: {
        placa: 'ABC1D23',
        marca: 'Toyota',
        modelo: ''
      },
      esperado: false
    },
    {
      nome: 'Placa invalida rejeitada',
      dados: {
        placa: 'AB123',
        marca: 'Toyota',
        modelo: 'Corolla'
      },
      esperado: false
    },
    {
      nome: 'Propulsao invalida rejeitada',
      dados: {
        placa: 'ABC1D23',
        marca: 'Toyota',
        modelo: 'Corolla',
        tipoPropulsao: 'NUCLEAR'
      },
      esperado: false
    },
    {
      nome: 'Ano invalido rejeitado',
      dados: {
        placa: 'ABC1D23',
        marca: 'Toyota',
        modelo: 'Corolla',
        anoFabricacao: 1800
      },
      esperado: false
    },
    {
      nome: 'Terceiro sem nome rejeitado',
      dados: {
        placa: 'ABC1D23',
        marca: 'Toyota',
        modelo: 'Corolla',
        proprietarioTipo: 'TERCEIRO'
      },
      esperado: false
    }
  ];

  let aprovados = 0;
  let falhas = 0;

  testes.forEach(teste => {
    try {
      const resultado =
        DCP_VALIDAR_DADOS_VEICULO_235A(teste.dados);

      if (resultado.valido === teste.esperado) {
        Logger.log('OK: ' + teste.nome);
        aprovados++;
      } else {
        Logger.log(
          'FALHA: ' + teste.nome +
          ' | ' + JSON.stringify(resultado)
        );
        falhas++;
      }
    } catch (erro) {
      Logger.log(
        'ERRO: ' + teste.nome +
        ' | ' + erro.message
      );
      falhas++;
    }
  });

  Logger.log('APROVADOS: ' + aprovados);
  Logger.log('FALHAS: ' + falhas);
  Logger.log('TOTAL: ' + testes.length);
}
/**
 * MODULO 2.3.5B
 * Cadastro controlado de veiculos com auditoria.
 */

const DCP_CADASTRO_VEICULOS_235B = {
  ABA: 'VEICULOS',
  CONFIG: 'CONFIG',
  AUDITORIA: 'AUDITORIA',
  MOTORISTAS: 'MOTORISTAS',
  PREFIXO: 'VEI',
  SEQUENCIA: 'SEQ_VEI',
  ADMIN: 'ADMIN_CENTRAL',
  DIGITOS: 4,

  CABECALHOS: [
    'VeiculoID', 'CodigoVeiculo', 'Placa', 'Marca',
    'Modelo', 'AnoFabricacao', 'AnoModelo',
    'TipoVeiculo', 'TipoPropulsao',
    'CombustivelPrincipal', 'CombustivelSecundario',
    'ProprietarioTipo', 'ProprietarioMotoristaID',
    'ProprietarioNome', 'CapacidadeTanqueLitros',
    'CapacidadeBateriaKWh', 'HodometroAtual',
    'DataCadastro', 'Observacoes', 'Status',
    'CriadoEm', 'AtualizadoEm', 'CriadoPor', 'AtualizadoPor'
  ],

  CABECALHOS_CONFIG: [
    'Chave', 'Valor', 'TipoValor',
    'Descricao', 'Editavel', 'AtualizadoEm'
  ],

  CABECALHOS_AUDITORIA: [
    'AuditoriaID', 'DataHora', 'UsuarioID', 'Acao',
    'Tabela', 'RegistroID', 'ValorAnterior',
    'ValorNovo', 'Resultado', 'Origem', 'Detalhes'
  ]
};


/**
 * Registra auditoria sem adquirir outro lock.
 * Deve ser chamada somente dentro do lock principal.
 */
function DCP_AUDITAR_VEICULO_235B(evento) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(
    DCP_CADASTRO_VEICULOS_235B.AUDITORIA
  );

  const dados = [
    Utilities.getUuid(),
    new Date(),
    DCP_CADASTRO_VEICULOS_235B.ADMIN,
    evento.acao,
    'VEICULOS',
    evento.registroID,
    JSON.stringify(evento.valorAnterior || {}),
    JSON.stringify(evento.valorNovo || {}),
    evento.resultado,
    'APPS_SCRIPT',
    evento.detalhes || ''
  ];

  aba.appendRow(dados);
}


/**
 * Cadastra um veiculo.
 */
function DCP_CADASTRAR_VEICULO_235B(dados) {
  const lock = LockService.getDocumentLock();

  if (!lock.tryLock(30000)) {
    return {
      sucesso: false,
      mensagem: 'Nao foi possivel obter o bloqueio. Tente novamente.'
    };
  }

  let tentativaRegistrada = false;
  let gravacaoRealizada = false;
  let veiculoID = '';
  let codigoVeiculo = '';

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cfg = DCP_CADASTRO_VEICULOS_235B;

    const aba = ss.getSheetByName(cfg.ABA);
    const config = ss.getSheetByName(cfg.CONFIG);
    const auditoria = ss.getSheetByName(cfg.AUDITORIA);
    const motoristas = ss.getSheetByName(cfg.MOTORISTAS);

    if (!aba || !config || !auditoria || !motoristas) {
      throw new Error('Uma ou mais abas obrigatorias nao existem.');
    }

    // Confere os cabecalhos.
    const conferir = (planilha, esperados, nome) => {
      const atuais = planilha.getRange(
        1, 1, 1, esperados.length
      ).getValues()[0];

      const iguais = esperados.every(
        (campo, indice) => atuais[indice] === campo
      );

      if (!iguais) {
        throw new Error('Estrutura invalida na aba ' + nome);
      }
    };

    conferir(aba, cfg.CABECALHOS, cfg.ABA);
    conferir(config, cfg.CABECALHOS_CONFIG, cfg.CONFIG);
    conferir(auditoria, cfg.CABECALHOS_AUDITORIA, cfg.AUDITORIA);

    // Validacao integral sob o mesmo lock.
    const validacao = DCP_VALIDAR_DADOS_VEICULO_235A(dados);

    if (!validacao.valido) {
      return {
        sucesso: false,
        erros: validacao.erros,
        mensagem: 'Dados do veiculo nao passaram na validacao.'
      };
    }

    // Recupera a maior sequencia ja utilizada.
    const ultimaLinha = aba.getLastRow();
    let maiorSequencia = 0;

    if (ultimaLinha > 1) {
      const codigos = aba.getRange(
        2, 2, ultimaLinha - 1, 1
      ).getValues();

      codigos.forEach(linha => {
        const match = String(linha[0] || '').match(/^VEI-(\d+)$/);

        if (match) {
          maiorSequencia = Math.max(
            maiorSequencia,
            Number(match[1])
          );
        }
      });
    }

    // Le a sequencia registrada na CONFIG.
    const dadosConfig = config.getDataRange().getValues();
    let linhaSequencia = -1;
    let sequenciaConfig = 0;

    dadosConfig.forEach((linha, indice) => {
      if (indice > 0 && linha[0] === cfg.SEQUENCIA) {
        linhaSequencia = indice + 1;
        sequenciaConfig = Number(linha[1]) || 0;
      }
    });

    const proximaSequencia =
      Math.max(maiorSequencia, sequenciaConfig) + 1;

    codigoVeiculo =
      cfg.PREFIXO + '-' +
      String(proximaSequencia).padStart(cfg.DIGITOS, '0');

    // Confere se o codigo ainda nao existe.
    if (ultimaLinha > 1) {
      const codigosExistentes = aba.getRange(
        2, 2, ultimaLinha - 1, 1
      ).getValues().flat();

      if (codigosExistentes.includes(codigoVeiculo)) {
        throw new Error('Codigo de veiculo ja existente.');
      }
    }

    veiculoID = Utilities.getUuid();

    // Identifica o proprietario automaticamente quando for motorista.
    let proprietarioNome = String(dados.proprietarioNome || '').trim();

    if (
      String(dados.proprietarioTipo || '').trim().toUpperCase() ===
      'MOTORISTA'
    ) {
      const registrosMotoristas = motoristas.getDataRange().getValues();

      const motorista = registrosMotoristas.find(
        (linha, indice) =>
          indice > 0 &&
          String(linha[0]).trim() ===
          String(dados.proprietarioMotoristaID).trim()
      );

      if (!motorista || String(motorista[8]).trim() !== 'ATIVO') {
        throw new Error('Motorista proprietario nao esta ativo.');
      }

      proprietarioNome = String(motorista[2]).trim();
    }

    const agora = new Date();

    const numeroOuVazio = valor =>
      valor === '' || valor === null || valor === undefined
        ? ''
        : Number(valor);

    const registro = [
      veiculoID,
      codigoVeiculo,
      validacao.placaNormalizada,
      String(dados.marca).trim(),
      String(dados.modelo).trim(),
      numeroOuVazio(dados.anoFabricacao),
      numeroOuVazio(dados.anoModelo),
      String(dados.tipoVeiculo || '').trim(),
      String(dados.tipoPropulsao || '').trim().toUpperCase(),
      '', // Compatibilidade sera registrada em tabela propria.
      '', // Compatibilidade sera registrada em tabela propria.
      String(dados.proprietarioTipo || '').trim().toUpperCase(),
      String(dados.proprietarioMotoristaID || '').trim(),
      proprietarioNome,
      numeroOuVazio(dados.capacidadeTanqueLitros),
      numeroOuVazio(dados.capacidadeBateriaKWh),
      numeroOuVazio(dados.hodometroAtual),
      agora,
      String(dados.observacoes || '').trim(),
      'ATIVO',
      agora,
      agora,
      cfg.ADMIN,
      cfg.ADMIN
    ];

    // Reserva a sequencia.
    if (linhaSequencia > 0) {
      config.getRange(linhaSequencia, 2, 1, 5).setValues([[
        proximaSequencia,
        'NUMERO',
        'Sequencia de codigos de veiculos',
        false,
        agora
      ]]);
    } else {
      config.appendRow([
        cfg.SEQUENCIA,
        proximaSequencia,
        'NUMERO',
        'Sequencia de codigos de veiculos',
        false,
        agora
      ]);
    }

    // Registra a tentativa antes da gravacao.
    DCP_AUDITAR_VEICULO_235B({
      acao: 'TENTATIVA_CRIACAO',
      registroID: veiculoID,
      valorAnterior: {},
      valorNovo: {
        VeiculoID: veiculoID,
        CodigoVeiculo: codigoVeiculo,
        Placa: validacao.placaNormalizada,
        Marca: registro[3],
        Modelo: registro[4],
        Status: 'ATIVO'
      },
      resultado: 'PENDENTE',
      detalhes: 'Tentativa de cadastro de veiculo.'
    });

    tentativaRegistrada = true;

    // Grava o registro.
    aba.getRange(
      aba.getLastRow() + 1,
      1,
      1,
      registro.length
    ).setValues([registro]);

    // A gravacao foi aceita pelo Sheets. Marcar antes do flush
    // evita tratar um erro posterior como se nada tivesse sido gravado.
    gravacaoRealizada = true;
    SpreadsheetApp.flush();

    // Confirma a operacao na auditoria.
    DCP_AUDITAR_VEICULO_235B({
      acao: 'CONFIRMACAO_CRIACAO',
      registroID: veiculoID,
      valorAnterior: {},
      valorNovo: {
        VeiculoID: veiculoID,
        CodigoVeiculo: codigoVeiculo,
        Placa: validacao.placaNormalizada,
        Marca: registro[3],
        Modelo: registro[4],
        Status: 'ATIVO'
      },
      resultado: 'SUCESSO',
      detalhes: 'Veiculo cadastrado com sucesso.'
    });

    return {
      sucesso: true,
      auditoriaPendente: false,
      id: veiculoID,
      codigo: codigoVeiculo,
      placa: validacao.placaNormalizada,
      mensagem: 'Veiculo cadastrado com sucesso.'
    };

  } catch (erro) {
    // Confere se o registro chegou a ser persistido mesmo que uma
    // excecao tenha ocorrido durante a confirmacao da gravacao.
    if (!gravacaoRealizada && veiculoID) {
      try {
        const ssVerificacao = SpreadsheetApp.getActiveSpreadsheet();
        const abaVerificacao = ssVerificacao.getSheetByName(
          DCP_CADASTRO_VEICULOS_235B.ABA
        );
        if (abaVerificacao && abaVerificacao.getLastRow() > 1) {
          const idsVerificacao = abaVerificacao.getRange(
            2, 1, abaVerificacao.getLastRow() - 1, 1
          ).getValues();
          gravacaoRealizada = idsVerificacao.some(
            linha => String(linha[0]) === veiculoID
          );
        }
      } catch (erroVerificacao) {
        Logger.log('Falha ao verificar persistencia: ' + erroVerificacao.message);
      }
    }

    if (gravacaoRealizada) {
      return {
        sucesso: true,
        auditoriaPendente: true,
        id: veiculoID,
        codigo: codigoVeiculo,
        mensagem:
          'Veiculo gravado, mas a confirmacao da auditoria falhou. ' +
          'Nao repita o cadastro. Confira a planilha e a auditoria.',
        detalhe: erro.message
      };
    }

    if (tentativaRegistrada) {
      try {
        DCP_AUDITAR_VEICULO_235B({
          acao: 'FALHA_CRIACAO',
          registroID: veiculoID,
          valorAnterior: {},
          valorNovo: {
            CodigoVeiculo: codigoVeiculo
          },
          resultado: 'FALHA',
          detalhes: erro.message
        });
      } catch (erroAuditoria) {
        Logger.log(
          'Falha adicional ao registrar auditoria: ' +
          erroAuditoria.message
        );
      }
    }

    return {
      sucesso: false,
      mensagem: 'Falha no cadastro do veiculo.',
      detalhe: erro.message
    };

  } finally {
    lock.releaseLock();
  }
}
function DCP_TESTAR_ESTRUTURA_CADASTRO_VEICULO_235B() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const estruturas = [
    {
      nome: 'VEICULOS',
      cabecalhos: DCP_CADASTRO_VEICULOS_235B.CABECALHOS
    },
    {
      nome: 'CONFIG',
      cabecalhos: DCP_CADASTRO_VEICULOS_235B.CABECALHOS_CONFIG
    },
    {
      nome: 'AUDITORIA',
      cabecalhos: DCP_CADASTRO_VEICULOS_235B.CABECALHOS_AUDITORIA
    }
  ];

  let aprovados = 0;
  let falhas = 0;

  estruturas.forEach(item => {
    const aba = ss.getSheetByName(item.nome);

    if (!aba) {
      Logger.log('FALHA: Aba ausente: ' + item.nome);
      falhas++;
      return;
    }

    const atuais = aba.getRange(
      1, 1, 1, item.cabecalhos.length
    ).getValues()[0];

    const correto = item.cabecalhos.every(
      (campo, indice) => atuais[indice] === campo
    );

    if (correto) {
      Logger.log('OK: ' + item.nome);
      aprovados++;
    } else {
      Logger.log('FALHA: Cabecalhos divergentes: ' + item.nome);
      falhas++;
    }
  });

  Logger.log('APROVADOS: ' + aprovados);
  Logger.log('FALHAS: ' + falhas);
  Logger.log('TOTAL: ' + estruturas.length);
}

function DCP_TESTAR_SEGURANCA_CADASTRO_VEICULO_235B() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const veiculos = ss.getSheetByName('VEICULOS');
  const config = ss.getSheetByName('CONFIG');
  const auditoria = ss.getSheetByName('AUDITORIA');

  function lerSequencia() {
    const dados = config.getDataRange().getValues();
    const linha = dados.find(
      (item, indice) =>
        indice > 0 && item[0] === 'SEQ_VEI'
    );
    return linha ? String(linha[1]) : '';
  }

  const linhasAntes = veiculos.getLastRow();
  const auditoriaAntes = auditoria.getLastRow();
  const sequenciaAntes = lerSequencia();

  const resultado = DCP_CADASTRAR_VEICULO_235B({
    placa: '',
    marca: '',
    modelo: ''
  });

  const linhasDepois = veiculos.getLastRow();
  const auditoriaDepois = auditoria.getLastRow();
  const sequenciaDepois = lerSequencia();

  const rejeitado = resultado.sucesso === false;
  const semVeiculoNovo = linhasAntes === linhasDepois;
  const semAuditoriaNova = auditoriaAntes === auditoriaDepois;
  const semConsumoSequencia = sequenciaAntes === sequenciaDepois;

  Logger.log(
    (rejeitado ? 'OK' : 'FALHA') +
    ': Dados invalidos rejeitados.'
  );

  Logger.log(
    (semVeiculoNovo ? 'OK' : 'FALHA') +
    ': Nenhum veiculo gravado.'
  );

  Logger.log(
    (semAuditoriaNova ? 'OK' : 'FALHA') +
    ': Nenhum evento de auditoria criado.'
  );

  Logger.log(
    (semConsumoSequencia ? 'OK' : 'FALHA') +
    ': Sequencia preservada.'
  );

  Logger.log('Mensagem: ' + resultado.mensagem);

  const aprovados = [
    rejeitado,
    semVeiculoNovo,
    semAuditoriaNova,
    semConsumoSequencia
  ].filter(Boolean).length;

  Logger.log('APROVADOS: ' + aprovados);
  Logger.log('FALHAS: ' + (4 - aprovados));
}

function DCP_LISTAR_MOTORISTAS_ATIVOS_235C() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName('MOTORISTAS');

  if (!aba) {
    throw new Error('A aba MOTORISTAS não foi encontrada.');
  }

  const cabecalhos = [
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
  ];

  const atuais = aba.getRange(
    1, 1, 1, cabecalhos.length
  ).getValues()[0];

  if (!cabecalhos.every(
    (campo, indice) => atuais[indice] === campo
  )) {
    throw new Error('Cabeçalhos da aba MOTORISTAS divergentes.');
  }

  const ultimaLinha = aba.getLastRow();

  if (ultimaLinha < 2) {
    return [];
  }

  const dados = aba.getRange(
    2, 1, ultimaLinha - 1, cabecalhos.length
  ).getValues();

  return dados
    .filter(linha => String(linha[8]).trim().toUpperCase() === 'ATIVO')
    .map(linha => ({
      id: String(linha[0]),
      codigo: String(linha[1]),
      nome: String(linha[2])
    }))
    .filter(item => item.id && item.nome)
    .sort((a, b) => a.nome.localeCompare(b.nome, 'pt-BR'));
}

function DCP_TESTAR_LISTA_MOTORISTAS_235C() {
  const motoristas = DCP_LISTAR_MOTORISTAS_ATIVOS_235C();

  Logger.log('Motoristas ativos encontrados: ' + motoristas.length);

  motoristas.forEach(motorista => {
    Logger.log(
      motorista.codigo + ' | ' + motorista.nome
    );
  });

  Logger.log('RESULTADO: CONSULTA CONCLUIDA.');
}

function DCP_ABRIR_FORM_VEICULO_235C() {
  const html = HtmlService
    .createHtmlOutputFromFile('DCP_FormVeiculo')
    .setWidth(650)
    .setHeight(750);

  SpreadsheetApp.getUi().showModalDialog(
    html,
    'DRIVER CONTROL PRO | Novo veículo'
  );
}

function DCP_DIAGNOSTICAR_VALIDACOES_VEICULO_235B() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName('VEICULOS');

  if (!aba) {
    throw new Error('Aba VEICULOS não encontrada.');
  }

  const uuid = '0d9a4136-fc7c-4469-b6ea-223c4f746480';
  const ultimaLinha = aba.getLastRow();

  const ids = ultimaLinha >= 2
    ? aba.getRange(2, 1, ultimaLinha - 1, 1).getValues()
    : [];

  const indice = ids.findIndex(linha => String(linha[0]) === uuid);

  if (indice === -1) {
    Logger.log('REGISTRO NÃO ENCONTRADO.');
    return;
  }

  const linha = indice + 2;
  const cabecalhos = aba.getRange(1, 1, 1, 24).getValues()[0];
  const valores = aba.getRange(linha, 1, 1, 24).getValues()[0];
  const validacoes = aba.getRange(linha, 1, 1, 24).getDataValidations()[0];

  Logger.log('===== DIAGNÓSTICO DE VALIDAÇÕES =====');
  Logger.log('Linha do veículo: ' + linha);

  for (let i = 0; i < 24; i++) {
    const regra = validacoes[i];

    if (regra) {
      Logger.log(
        'COLUNA ' + (i + 1) +
        ' | ' + cabecalhos[i] +
        ' | Valor atual: ' + valores[i] +
        ' | Critério: ' + regra.getCriteriaType() +
        ' | Valores permitidos: ' +
        JSON.stringify(regra.getCriteriaValues()) +
        ' | Permite valor inválido: ' + regra.getAllowInvalid()
      );
    }
  }

  Logger.log('===== FIM DO DIAGNÓSTICO =====');
}

function DCP_CORRIGIR_VALIDACOES_VEICULOS_235B() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName('VEICULOS');

  if (!aba) {
    throw new Error('Aba VEICULOS não encontrada.');
  }

  const ultimaLinha = Math.max(aba.getMaxRows(), 2);

  // TipoVeiculo: texto livre.
  aba.getRange(2, 8, ultimaLinha - 1, 1)
    .clearDataValidations();

  // TipoPropulsao: lista oficial do sistema.
  const propulsao = SpreadsheetApp.newDataValidation()
    .requireValueInList([
      'COMBUSTAO',
      'FLEX',
      'HIBRIDO',
      'HIBRIDO_PLUGIN',
      'ELETRICO'
    ], true)
    .setAllowInvalid(false)
    .build();

  aba.getRange(2, 9, ultimaLinha - 1, 1)
    .setDataValidation(propulsao);

  // Campos legados de combustível: sem validação.
  aba.getRange(2, 10, ultimaLinha - 1, 2)
    .clearDataValidations();

  // ProprietarioTipo: aceita vazio ou um dos dois tipos.
  const proprietario = SpreadsheetApp.newDataValidation()
    .requireValueInList([
      'MOTORISTA',
      'TERCEIRO'
    ], true)
    .setAllowInvalid(true)
    .build();

  aba.getRange(2, 12, ultimaLinha - 1, 1)
    .setDataValidation(proprietario);

  SpreadsheetApp.flush();

  Logger.log('VALIDAÇÕES ATUALIZADAS.');
  Logger.log('TipoVeiculo: texto livre.');
  Logger.log('TipoPropulsao: lista oficial.');
  Logger.log('Campos legados: validações removidas.');
  Logger.log('ProprietarioTipo: opcional.');
  Logger.log('Status: preservado.');
}
function DCP_CONFERIR_RECUPERACAO_VEI0001() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName('VEICULOS');
  const auditoria = ss.getSheetByName('AUDITORIA');

  const uuid = '0d9a4136-fc7c-4469-b6ea-223c4f746480';

  const cabecalhos = aba.getRange(1, 1, 1, 24).getValues()[0];
  const dados = aba.getDataRange().getValues();

  const indice = dados.findIndex(
    (linha, i) => i > 0 && String(linha[0]) === uuid
  );

  if (indice === -1) {
    Logger.log('FALHA: VEI-0001 não encontrado.');
    return;
  }

  Logger.log('===== DADOS ATUAIS DO VEI-0001 =====');

  cabecalhos.forEach((campo, i) => {
    Logger.log(campo + ': ' + dados[indice][i]);
  });

  Logger.log('===== AUDITORIA =====');

  const eventos = auditoria.getDataRange().getValues();

  eventos.forEach((linha, i) => {
    if (i > 0 && String(linha[5]) === uuid) {
      Logger.log(
        linha[3] + ' | ' +
        linha[8] + ' | ' +
        linha[10]
      );
    }
  });

  Logger.log('===== FIM DA CONFERÊNCIA =====');
}


/**
 * Recupera o registro incompleto VEI-0001 sem gerar novo UUID ou codigo.
 * Usa os dados ficticios previamente aprovados no teste do formulario.
 * Executar somente uma vez, apos conferir os dados e a auditoria.
 */
function DCP_RECUPERAR_VEI0001_235C() {
  const lock = LockService.getDocumentLock();

  if (!lock.tryLock(30000)) {
    return {
      sucesso: false,
      mensagem: 'Nao foi possivel obter o bloqueio. Tente novamente.'
    };
  }

  const uuidAlvo = '0d9a4136-fc7c-4469-b6ea-223c4f746480';
  const codigoAlvo = 'VEI-0001';
  let tentativaRegistrada = false;
  let gravacaoRealizada = false;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cfg = DCP_CADASTRO_VEICULOS_235B;
    const aba = ss.getSheetByName(cfg.ABA);
    const auditoria = ss.getSheetByName(cfg.AUDITORIA);
    const motoristas = ss.getSheetByName(cfg.MOTORISTAS);

    if (!aba || !auditoria || !motoristas) {
      throw new Error('Uma ou mais abas obrigatorias nao existem.');
    }

    const conferir = (planilha, esperados, nome) => {
      const atuais = planilha.getRange(1, 1, 1, esperados.length).getValues()[0];
      if (!esperados.every((campo, indice) => atuais[indice] === campo)) {
        throw new Error('Estrutura invalida na aba ' + nome + '.');
      }
    };

    conferir(aba, cfg.CABECALHOS, cfg.ABA);
    conferir(auditoria, cfg.CABECALHOS_AUDITORIA, cfg.AUDITORIA);

    const dadosAtuais = aba.getDataRange().getValues();
    const indices = [];
    for (let i = 1; i < dadosAtuais.length; i++) {
      if (String(dadosAtuais[i][0]).trim() === uuidAlvo) indices.push(i);
    }

    if (indices.length !== 1) {
      throw new Error(
        'Esperado exatamente um registro com o UUID de VEI-0001; encontrados: ' + indices.length
      );
    }

    const indiceRegistro = indices[0];
    const linhaPlanilha = indiceRegistro + 1;
    const anterior = dadosAtuais[indiceRegistro];

    if (
      String(anterior[1]).trim() !== codigoAlvo ||
      DCP_NORMALIZAR_PLACA_235A(anterior[2]) !== 'ABC1D23' ||
      String(anterior[3]).trim() !== 'Toyota' ||
      String(anterior[4]).trim() !== 'Corolla'
    ) {
      throw new Error('O registro encontrado nao corresponde ao VEI-0001 ficticio esperado. Recuperacao cancelada.');
    }

    if (String(anterior[19]).trim() === 'ATIVO') {
      return {
        sucesso: false,
        jaRecuperado: true,
        mensagem: 'O VEI-0001 ja consta como ATIVO. Nenhuma alteracao foi realizada.'
      };
    }

    if (String(anterior[19]).trim() !== '') {
      throw new Error('O registro possui Status preenchido diferente de ATIVO. Recuperacao cancelada.');
    }

    // Evita sobrescrever qualquer registro que ja tenha sido completado.
    const camposQueDevemEstarVazios = [7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 20, 21, 22, 23];
    const preenchimentosInesperados = camposQueDevemEstarVazios.filter(
      indice => anterior[indice] !== '' && anterior[indice] !== null
    );
    if (preenchimentosInesperados.length) {
      throw new Error(
        'O registro contem campos ja preenchidos fora do esperado. Recuperacao automatica cancelada. Colunas: ' +
        preenchimentosInesperados.map(i => i + 1).join(', ')
      );
    }

    // Recupera o proprietario pelo codigo publico do motorista, nunca pelo nome digitado.
    const dadosMotoristas = motoristas.getDataRange().getValues();
    const encontradosMotorista = dadosMotoristas.filter((linha, indice) =>
      indice > 0 && String(linha[1]).trim() === 'MOT-0001'
    );

    if (encontradosMotorista.length !== 1) {
      throw new Error('Nao foi encontrado exatamente um motorista MOT-0001.');
    }

    const motorista = encontradosMotorista[0];
    if (String(motorista[8]).trim().toUpperCase() !== 'ATIVO') {
      throw new Error('O motorista MOT-0001 nao esta ATIVO.');
    }

    const dadosRecuperacao = {
      placa: 'ABC1D23',
      marca: 'Toyota',
      modelo: 'Corolla',
      anoFabricacao: 2023,
      anoModelo: 2024,
      tipoVeiculo: 'Sedan',
      tipoPropulsao: 'FLEX',
      proprietarioTipo: 'MOTORISTA',
      proprietarioMotoristaID: String(motorista[0]),
      proprietarioNome: '',
      capacidadeTanqueLitros: 50,
      capacidadeBateriaKWh: '',
      hodometroAtual: 25000,
      observacoes: 'VEICULO FICTICIO PARA TESTES DO SISTEMA'
    };

    const validacao = DCP_VALIDAR_DADOS_VEICULO_235A(
      dadosRecuperacao,
      uuidAlvo
    );

    if (!validacao.valido) {
      throw new Error('Validacao da recuperacao falhou: ' + validacao.erros.join(' | '));
    }

    const agora = new Date();
    const numeroOuVazio = valor =>
      valor === '' || valor === null || valor === undefined ? '' : Number(valor);
    const proprietarioNome = String(motorista[2]).trim();

    const registro = [
      uuidAlvo,
      codigoAlvo,
      validacao.placaNormalizada,
      'Toyota',
      'Corolla',
      2023,
      2024,
      'Sedan',
      'FLEX',
      '',
      '',
      'MOTORISTA',
      String(motorista[0]),
      proprietarioNome,
      numeroOuVazio(50),
      '',
      numeroOuVazio(25000),
      agora,
      'VEICULO FICTICIO PARA TESTES DO SISTEMA',
      'ATIVO',
      agora,
      agora,
      cfg.ADMIN,
      cfg.ADMIN
    ];

    const objetoAnterior = {};
    cfg.CABECALHOS.forEach((campo, i) => {
      objetoAnterior[campo] = anterior[i] instanceof Date
        ? anterior[i].toISOString()
        : anterior[i];
    });

    const objetoNovo = {};
    cfg.CABECALHOS.forEach((campo, i) => {
      objetoNovo[campo] = registro[i] instanceof Date
        ? registro[i].toISOString()
        : registro[i];
    });

    DCP_AUDITAR_VEICULO_235B({
      acao: 'TENTATIVA_RECUPERACAO',
      registroID: uuidAlvo,
      valorAnterior: objetoAnterior,
      valorNovo: objetoNovo,
      resultado: 'PENDENTE',
      detalhes: 'Tentativa controlada de recuperacao do registro incompleto VEI-0001.'
    });
    tentativaRegistrada = true;

    aba.getRange(linhaPlanilha, 1, 1, registro.length).setValues([registro]);
    gravacaoRealizada = true;
    SpreadsheetApp.flush();

    DCP_AUDITAR_VEICULO_235B({
      acao: 'CONFIRMACAO_RECUPERACAO',
      registroID: uuidAlvo,
      valorAnterior: objetoAnterior,
      valorNovo: objetoNovo,
      resultado: 'SUCESSO',
      detalhes: 'Registro VEI-0001 recuperado com sucesso, preservando UUID e codigo.'
    });

    return {
      sucesso: true,
      auditoriaPendente: false,
      id: uuidAlvo,
      codigo: codigoAlvo,
      placa: 'ABC1D23',
      mensagem: 'VEI-0001 recuperado com sucesso. UUID e codigo preservados.'
    };

  } catch (erro) {
    if (gravacaoRealizada) {
      return {
        sucesso: true,
        auditoriaPendente: true,
        id: uuidAlvo,
        codigo: codigoAlvo,
        mensagem: 'Registro recuperado, mas a confirmacao da auditoria falhou. Nao execute novamente; confira a linha e a auditoria.',
        detalhe: erro.message
      };
    }

    if (tentativaRegistrada) {
      try {
        DCP_AUDITAR_VEICULO_235B({
          acao: 'FALHA_RECUPERACAO',
          registroID: uuidAlvo,
          valorAnterior: {},
          valorNovo: { CodigoVeiculo: codigoAlvo },
          resultado: 'FALHA',
          detalhes: erro.message
        });
      } catch (erroAuditoria) {
        Logger.log('Falha adicional ao registrar auditoria da recuperacao: ' + erroAuditoria.message);
      }
    }

    return {
      sucesso: false,
      mensagem: 'Falha na recuperacao do VEI-0001.',
      detalhe: erro.message
    };

  } finally {
    lock.releaseLock();
  }
}
function DCP_CONFERIR_RECUPERACAO_VEI0001() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName('VEICULOS');

  const uuid = '0d9a4136-fc7c-4469-b6ea-223c4f746480';
  const dados = aba.getDataRange().getValues();

  const encontrados = dados.filter(
    (linha, i) => i > 0 && String(linha[0]) === uuid
  );

  Logger.log('REGISTROS ENCONTRADOS: ' + encontrados.length);

  encontrados.forEach(linha => {
    Logger.log('CODIGO: ' + linha[1]);
    Logger.log('PLACA: ' + linha[2]);
    Logger.log('MARCA: ' + linha[3]);
    Logger.log('MODELO: ' + linha[4]);
    Logger.log('PROPRIETARIO: ' + linha[13]);
    Logger.log('STATUS: ' + linha[19]);
  });
}
function DCP_VALIDAR_INTEGRIDADE_VEI0001() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName('VEICULOS');
  const auditoria = ss.getSheetByName('AUDITORIA');

  const uuid = '0d9a4136-fc7c-4469-b6ea-223c4f746480';

  const dados = aba.getDataRange().getValues();
  const registros = dados.filter(
    (linha, i) => i > 0 && String(linha[0]) === uuid
  );

  Logger.log('===== INTEGRIDADE DO VEICULO =====');

  if (registros.length !== 1) {
    Logger.log('FALHA: quantidade de registros diferente de 1.');
    return;
  }

  const linha = registros[0];
  const cabecalhos = dados[0];
  let falhas = 0;

  cabecalhos.forEach((campo, i) => {
    const vazio = linha[i] === '' || linha[i] === null;

    if (vazio) {
      Logger.log('ATENCAO: campo vazio: ' + campo);
      falhas++;
    } else {
      Logger.log('OK: ' + campo);
    }
  });

  Logger.log('===== EVENTOS DE AUDITORIA =====');

  const eventos = auditoria.getDataRange().getValues();

  eventos.forEach((evento, i) => {
    if (i > 0 && String(evento[5]) === uuid) {
      Logger.log(
        evento[3] + ' | ' +
        evento[8] + ' | ' +
        evento[10]
      );
    }
  });

  Logger.log('===== RESULTADO =====');
  Logger.log('Campos vazios: ' + falhas);
}
function DCP_TESTAR_DUPLICIDADE_VEICULOS_235D() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName('VEICULOS');

  const uuid = '0d9a4136-fc7c-4469-b6ea-223c4f746480';
  const placa = 'ABC1D23';

  Logger.log('===== TESTE DE DUPLICIDADE =====');

  // Localiza o veículo original.
  const registros = aba.getDataRange().getValues();

  const encontrados = registros.filter(
    (linha, indice) =>
      indice > 0 &&
      String(linha[0]).trim() === uuid
  );

  if (encontrados.length !== 1) {
    Logger.log('FALHA: VEI-0001 nao encontrado de forma unica.');
    return;
  }

  const veiculo = encontrados[0];

  // Monta os dados usando o próprio cadastro.
  const dados = {
    placa: placa,
    marca: veiculo[3],
    modelo: veiculo[4],
    anoFabricacao: veiculo[5],
    anoModelo: veiculo[6],
    tipoVeiculo: veiculo[7],
    tipoPropulsao: veiculo[8],
    proprietarioTipo: veiculo[11],
    proprietarioMotoristaID: veiculo[12],
    proprietarioNome: '',
    capacidadeTanqueLitros: veiculo[14],
    capacidadeBateriaKWh: veiculo[15],
    hodometroAtual: veiculo[16],
    observacoes: veiculo[18]
  };

  // TESTE 1: novo cadastro com placa já utilizada.
  const testeNovo =
    DCP_VALIDAR_DADOS_VEICULO_235A(dados);

  const rejeitouDuplicidade =
    testeNovo.valido === false &&
    testeNovo.erros.some(
      erro => erro.includes('placa ja esta cadastrada')
    );

  Logger.log(
    (rejeitouDuplicidade ? 'OK' : 'FALHA') +
    ': Novo cadastro com placa duplicada.'
  );

  // TESTE 2: edição ignorando o próprio UUID.
  const testeEdicao =
    DCP_VALIDAR_DADOS_VEICULO_235A(dados, uuid);

  Logger.log(
    (testeEdicao.valido ? 'OK' : 'FALHA') +
    ': Proprio veiculo ignorado durante edicao.'
  );

  if (!testeEdicao.valido) {
    Logger.log(
      'Detalhes da validacao: ' +
      JSON.stringify(testeEdicao.erros)
    );
  }

  // TESTE 3: verifica se a busca considera todos os status.
  const duplicatas = registros.filter(
    (linha, indice) =>
      indice > 0 &&
      DCP_NORMALIZAR_PLACA_235A(linha[2]) === placa
  );

  const statusEncontrados = duplicatas.map(
    linha => String(linha[19]).trim().toUpperCase()
  );

  const consideraTodosStatus = duplicatas.every(
    linha => String(linha[0]).trim() !== ''
  );

  Logger.log(
    (consideraTodosStatus ? 'OK' : 'FALHA') +
    ': Busca da placa considera registros independentemente do status.'
  );

  statusEncontrados.forEach(status => {
    Logger.log('Status encontrado para a placa: ' + status);
  });

  Logger.log('Quantidade de registros com a placa: ' + duplicatas.length);

  // Resultado.
  const aprovados = [
    rejeitouDuplicidade,
    testeEdicao.valido,
    consideraTodosStatus
  ].filter(Boolean).length;

  Logger.log('===== RESULTADO FINAL =====');
  Logger.log('APROVADOS: ' + aprovados + '/3');
  Logger.log('FALHAS: ' + (3 - aprovados));
  Logger.log('TESTE SEM ALTERACAO DE DADOS.');
}