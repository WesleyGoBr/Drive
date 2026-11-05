/************************************************************
 * DRIVER CONTROL PRO
 * MODULO 2.3.5D - EDICAO DE VEICULOS
 *
 * Edicao com UUID permanente, validacao e auditoria.
 ************************************************************/

const DCP_EDICAO_VEICULOS_235D = {
  ABA_VEICULOS: 'VEICULOS',
  ABA_AUDITORIA: 'AUDITORIA',

  CABECALHOS_VEICULOS: [
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
  ],

  CABECALHOS_AUDITORIA: [
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
  ],

  USUARIO: 'ADMIN_CENTRAL',
  ORIGEM: 'SCRIPT'
};


/**
 * Confere cabecalhos de uma aba.
 */
function DCP_EV235D_CONFERIR_CABECALHOS(aba, esperados) {
  if (!aba) {
    throw new Error('Aba nao encontrada.');
  }

  const encontrados = aba
    .getRange(1, 1, 1, esperados.length)
    .getDisplayValues()[0];

  for (let i = 0; i < esperados.length; i++) {
    if (String(encontrados[i]).trim() !== esperados[i]) {
      throw new Error(
        'Cabecalho divergente na coluna ' + (i + 1) +
        '. Esperado: ' + esperados[i] +
        '. Encontrado: ' + encontrados[i]
      );
    }
  }
}


/**
 * Serializa dados para auditoria.
 */
function DCP_EV235D_SERIALIZAR(valor) {
  return JSON.stringify(valor, function(chave, item) {
    if (item instanceof Date) {
      return Utilities.formatDate(
        item,
        Session.getScriptTimeZone(),
        "yyyy-MM-dd'T'HH:mm:ss"
      );
    }
    return item;
  });
}


/**
 * Registra auditoria sem adquirir outro lock.
 *
 * IMPORTANTE:
 * Esta funcao deve ser chamada somente dentro
 * da secao protegida pelo DocumentLock da operacao.
 */
function DCP_EV235D_AUDITAR_INTERNO(evento) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(
    DCP_EDICAO_VEICULOS_235D.ABA_AUDITORIA
  );

  DCP_EV235D_CONFERIR_CABECALHOS(
    aba,
    DCP_EDICAO_VEICULOS_235D.CABECALHOS_AUDITORIA
  );

  const linha = [
    DCP_GERAR_UUID(),
    new Date(),
    DCP_EDICAO_VEICULOS_235D.USUARIO,
    evento.acao,
    'VEICULOS',
    evento.registroID,
    evento.valorAnterior || '',
    evento.valorNovo || '',
    evento.resultado,
    DCP_EDICAO_VEICULOS_235D.ORIGEM,
    evento.detalhes || ''
  ];

  aba.getRange(
    aba.getLastRow() + 1,
    1,
    1,
    linha.length
  ).setValues([linha]);
}


/**
 * Recupera nome do motorista proprietario.
 */
function DCP_EV235D_OBTER_NOME_MOTORISTA(motoristaID) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName('MOTORISTAS');

  if (!aba) {
    throw new Error('Aba MOTORISTAS nao encontrada.');
  }

  const dados = aba.getDataRange().getValues();

  const encontrados = dados.filter(function(linha, indice) {
    return indice > 0 &&
      String(linha[0]).trim() === String(motoristaID).trim();
  });

  if (encontrados.length !== 1) {
    throw new Error(
      'Motorista proprietario nao encontrado de forma unica.'
    );
  }

  const motorista = encontrados[0];

  if (
    String(motorista[8]).trim().toUpperCase() !== 'ATIVO'
  ) {
    throw new Error('O motorista proprietario nao esta ativo.');
  }

  return String(motorista[2]).trim();
}


/**
 * Converte valores numericos opcionais.
 */
function DCP_EV235D_NUMERO_OPCIONAL(valor) {
  if (valor === '' || valor === null || valor === undefined) {
    return '';
  }

  if (typeof valor === 'number') {
    return valor;
  }

  const numero = Number(
    String(valor).trim().replace(',', '.')
  );

  return Number.isFinite(numero) ? numero : valor;
}


/**
 * Executa a edicao de um veiculo.
 *
 * dados deve conter:
 * veiculoID
 * placa
 * marca
 * modelo
 * anoFabricacao
 * anoModelo
 * tipoVeiculo
 * tipoPropulsao
 * proprietarioTipo
 * proprietarioMotoristaID
 * proprietarioNome
 * capacidadeTanqueLitros
 * capacidadeBateriaKWh
 * hodometroAtual
 * observacoes
 *
 * CombustivelPrincipal, CombustivelSecundario,
 * Status e DataCadastro sao preservados.
 */
function DCP_EDITAR_VEICULO_235D(dados) {
  const lock = LockService.getDocumentLock();

  if (!lock.tryLock(30000)) {
    return {
      sucesso: false,
      mensagem: 'Nao foi possivel obter o bloqueio. Tente novamente.'
    };
  }

  let gravacaoRealizada = false;
  let uuid = '';

  try {
    if (!dados || typeof dados !== 'object') {
      throw new Error('Dados de edicao invalidos.');
    }

    uuid = String(dados.veiculoID || '').trim();

    if (!uuid) {
      throw new Error('UUID do veiculo nao informado.');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName(
      DCP_EDICAO_VEICULOS_235D.ABA_VEICULOS
    );

    DCP_EV235D_CONFERIR_CABECALHOS(
      aba,
      DCP_EDICAO_VEICULOS_235D.CABECALHOS_VEICULOS
    );

    const auditoria = ss.getSheetByName(
      DCP_EDICAO_VEICULOS_235D.ABA_AUDITORIA
    );

    DCP_EV235D_CONFERIR_CABECALHOS(
      auditoria,
      DCP_EDICAO_VEICULOS_235D.CABECALHOS_AUDITORIA
    );

    const ultimaLinha = aba.getLastRow();

    if (ultimaLinha < 2) {
      throw new Error('Nenhum veiculo cadastrado.');
    }

    const registros = aba.getRange(
      2,
      1,
      ultimaLinha - 1,
      24
    ).getValues();

    const encontrados = [];

    registros.forEach(function(linha, indice) {
      if (String(linha[0]).trim() === uuid) {
        encontrados.push({
          linhaPlanilha: indice + 2,
          dados: linha
        });
      }
    });

    if (encontrados.length !== 1) {
      throw new Error(
        'Veiculo nao encontrado de forma unica pelo UUID.'
      );
    }

    const encontrado = encontrados[0];
    const anterior = encontrado.dados;
    const numeroLinha = encontrado.linhaPlanilha;

    const statusAtual = String(anterior[19] || '')
      .trim()
      .toUpperCase();

    if (!['ATIVO', 'INATIVO'].includes(statusAtual)) {
      throw new Error('Status atual do veiculo invalido.');
    }

    function obter(campo, indice) {
      return Object.prototype.hasOwnProperty.call(dados, campo)
        ? dados[campo]
        : anterior[indice];
    }

    const dadosValidacao = {
      placa: String(obter('placa', 2) || ''),
      marca: String(obter('marca', 3) || ''),
      modelo: String(obter('modelo', 4) || ''),
      anoFabricacao: obter('anoFabricacao', 5),
      anoModelo: obter('anoModelo', 6),
      tipoVeiculo: String(obter('tipoVeiculo', 7) || ''),
      tipoPropulsao: String(obter('tipoPropulsao', 8) || ''),
      proprietarioTipo: String(obter('proprietarioTipo', 11) || ''),
proprietarioMotoristaID:
  String(obter('proprietarioMotoristaID', 12) || ''),

proprietarioNome:
  String(obter('proprietarioTipo', 11) || '')
    .trim()
    .toUpperCase() === 'MOTORISTA'
      ? ''
      : String(obter('proprietarioNome', 13) || ''),
      capacidadeTanqueLitros:
        obter('capacidadeTanqueLitros', 14),
      capacidadeBateriaKWh:
        obter('capacidadeBateriaKWh', 15),
      hodometroAtual: obter('hodometroAtual', 16),
      observacoes: String(obter('observacoes', 18) || '')
    };

    /*
     * A validacao ignora o UUID do proprio veiculo
     * na verificacao de duplicidade da placa.
     */
    const validacao = DCP_VALIDAR_DADOS_VEICULO_235A(
      dadosValidacao,
      uuid
    );

    if (!validacao || validacao.valido !== true) {
      return {
        sucesso: false,
        mensagem: 'Existem erros de validacao.',
        erros: validacao && validacao.erros
          ? validacao.erros
          : ['Falha desconhecida na validacao.']
      };
    }

    const proprietarioTipo = String(
      dadosValidacao.proprietarioTipo || ''
    ).trim().toUpperCase();

    let proprietarioID = '';
    let proprietarioNome = '';

    if (proprietarioTipo === 'MOTORISTA') {
      proprietarioID = dadosValidacao.proprietarioMotoristaID;
      proprietarioNome =
        DCP_EV235D_OBTER_NOME_MOTORISTA(proprietarioID);

    } else if (proprietarioTipo === 'TERCEIRO') {
      proprietarioNome = dadosValidacao.proprietarioNome;
    }

    const novaLinha = anterior.slice();

    // Identificadores permanentes.
    novaLinha[0] = anterior[0];
    novaLinha[1] = anterior[1];

    // Dados editaveis.
    novaLinha[2] = DCP_NORMALIZAR_PLACA_235A(
      dadosValidacao.placa
    );
    novaLinha[3] = dadosValidacao.marca;
    novaLinha[4] = dadosValidacao.modelo;
    novaLinha[5] = DCP_EV235D_NUMERO_OPCIONAL(
      dadosValidacao.anoFabricacao
    );
    novaLinha[6] = DCP_EV235D_NUMERO_OPCIONAL(
      dadosValidacao.anoModelo
    );
    novaLinha[7] = dadosValidacao.tipoVeiculo;
    novaLinha[8] = dadosValidacao.tipoPropulsao;

    // Combustiveis legados preservados.
    novaLinha[9] = anterior[9];
    novaLinha[10] = anterior[10];

    novaLinha[11] = proprietarioTipo;
    novaLinha[12] = proprietarioID;
    novaLinha[13] = proprietarioNome;

    novaLinha[14] = DCP_EV235D_NUMERO_OPCIONAL(
      dadosValidacao.capacidadeTanqueLitros
    );
    novaLinha[15] = DCP_EV235D_NUMERO_OPCIONAL(
      dadosValidacao.capacidadeBateriaKWh
    );
    novaLinha[16] = DCP_EV235D_NUMERO_OPCIONAL(
      dadosValidacao.hodometroAtual
    );

    // DataCadastro e Status preservados.
    novaLinha[17] = anterior[17];
    novaLinha[18] = dadosValidacao.observacoes;
    novaLinha[19] = anterior[19];

    // Campos de criacao preservados.
    novaLinha[20] = anterior[20];

    // Campos de atualizacao.
    novaLinha[21] = new Date();
    novaLinha[22] = anterior[22];
    novaLinha[23] = DCP_EDICAO_VEICULOS_235D.USUARIO;

    const valorAnterior = DCP_EV235D_SERIALIZAR(anterior);
    const valorNovo = DCP_EV235D_SERIALIZAR(novaLinha);


    // Grava a linha inteira preservando campos protegidos.
    aba.getRange(numeroLinha, 1, 1, 24)
      .setValues([novaLinha]);

    gravacaoRealizada = true;

    SpreadsheetApp.flush();

    // Confirmacao da auditoria.
    try {
      DCP_EV235D_AUDITAR_INTERNO({
        acao: 'ALTERACAO',
        registroID: uuid,
        valorAnterior: valorAnterior,
        valorNovo: valorNovo,
        resultado: 'SUCESSO',
        detalhes: 'Edicao de veiculo concluida.'
      });

    } catch (erroAuditoria) {
      return {
        sucesso: true,
        auditoriaPendente: true,
        mensagem:
          'Veiculo atualizado, mas a confirmacao da auditoria precisa ser verificada.',
        veiculoID: uuid,
        codigoVeiculo: anterior[1]
      };
    }

    return {
      sucesso: true,
      auditoriaPendente: false,
      mensagem: 'Veiculo atualizado com sucesso.',
      veiculoID: uuid,
      codigoVeiculo: anterior[1]
    };

  } catch (erro) {
    if (gravacaoRealizada) {
      return {
        sucesso: true,
        auditoriaPendente: true,
        mensagem:
          'A gravacao foi realizada, mas ocorreu uma falha posterior. Verifique a auditoria.',
        veiculoID: uuid,
        detalhe: erro.message
      };
    }

    try {
      if (uuid) {
        DCP_EV235D_AUDITAR_INTERNO({
          acao: 'ERRO',
          registroID: uuid,
          resultado: 'FALHA',
          detalhes: String(erro.message || erro)
        });
      }
    } catch (erroAuditoria) {
      Logger.log(
        'Falha adicional ao registrar auditoria: ' +
        erroAuditoria.message
      );
    }

    return {
      sucesso: false,
      mensagem: 'Edicao nao realizada.',
      detalhe: erro.message
    };

  } finally {
    lock.releaseLock();
  }
}
function DCP_TESTAR_VALIDACAO_EDICAO_VEI0001_235D() {
  Logger.log('===== TESTE VALIDACAO DE EDICAO =====');

  const consulta = DCP_OBTER_VEICULO_235D('VEI-0001');

  if (!consulta.encontrado) {
    Logger.log('FALHA: VEI-0001 nao encontrado.');
    return;
  }

  const v = consulta.veiculo;

  const dados = {
    placa: v.Placa,
    marca: v.Marca,
    modelo: v.Modelo,
    anoFabricacao: v.AnoFabricacao,
    anoModelo: v.AnoModelo,
    tipoVeiculo: v.TipoVeiculo,
    tipoPropulsao: v.TipoPropulsao,
    proprietarioTipo: v.ProprietarioTipo,
    proprietarioMotoristaID: v.ProprietarioMotoristaID,
    proprietarioNome: '',
    capacidadeTanqueLitros: v.CapacidadeTanqueLitros,
    capacidadeBateriaKWh: v.CapacidadeBateriaKWh,
    hodometroAtual: v.HodometroAtual,
    observacoes: v.Observacoes
  };

  const resultado = DCP_VALIDAR_DADOS_VEICULO_235A(
    dados,
    v.VeiculoID
  );

  Logger.log(
    'VALIDACAO: ' +
    (resultado.valido ? 'APROVADA' : 'REPROVADA')
  );

  if (resultado.erros) {
    resultado.erros.forEach(function(erro) {
      Logger.log('ERRO: ' + erro);
    });
  }

  Logger.log('UUID PRESERVADO: ' + v.VeiculoID);
  Logger.log('CODIGO PRESERVADO: ' + v.CodigoVeiculo);
  Logger.log('PLACA TESTADA: ' + v.Placa);
  Logger.log('NENHUM DADO FOI ALTERADO.');
}
function DCP_TESTAR_EDICAO_REAL_VEI0001_235D() {
  Logger.log('===== TESTE REAL DE EDICAO =====');

  const consultaAntes = DCP_OBTER_VEICULO_235D('VEI-0001');

  if (!consultaAntes || !consultaAntes.encontrado) {
    throw new Error('VEI-0001 nao encontrado.');
  }

  const antes = consultaAntes.veiculo;
  const marcador = '[TESTE EDICAO 235D]';

  const observacoesAntes = String(antes.Observacoes || '');

  const novasObservacoes = observacoesAntes.includes(marcador)
    ? observacoesAntes
    : (observacoesAntes
        ? observacoesAntes + ' ' + marcador
        : marcador);

  const resultado = DCP_EDITAR_VEICULO_235D({
    veiculoID: antes.VeiculoID,
    observacoes: novasObservacoes
  });

  Logger.log('RESULTADO DA EDICAO: ' + JSON.stringify(resultado));

  if (!resultado.sucesso) {
    Logger.log('FALHA: A edicao nao foi concluida.');
    return;
  }

  const consultaDepois = DCP_OBTER_VEICULO_235D('VEI-0001');

  if (!consultaDepois || !consultaDepois.encontrado) {
    throw new Error('Nao foi possivel consultar o veiculo apos a edicao.');
  }

  const depois = consultaDepois.veiculo;

  function conferir(nome, condicao) {
    Logger.log((condicao ? 'OK: ' : 'FALHA: ') + nome);
  }

  conferir(
    'UUID preservado',
    antes.VeiculoID === depois.VeiculoID
  );

  conferir(
    'Codigo preservado',
    antes.CodigoVeiculo === depois.CodigoVeiculo
  );

  conferir(
    'Status preservado',
    antes.Status === depois.Status
  );

  conferir(
    'Data de cadastro preservada',
    DCP_EV235D_SERIALIZAR(antes.DataCadastro) ===
    DCP_EV235D_SERIALIZAR(depois.DataCadastro)
  );

  conferir(
    'CriadoEm preservado',
    DCP_EV235D_SERIALIZAR(antes.CriadoEm) ===
    DCP_EV235D_SERIALIZAR(depois.CriadoEm)
  );

  conferir(
    'CriadoPor preservado',
    antes.CriadoPor === depois.CriadoPor
  );

  conferir(
    'Observacoes atualizadas',
    String(depois.Observacoes || '').includes(marcador)
  );

  conferir(
    'AtualizadoEm preenchido',
    Boolean(depois.AtualizadoEm)
  );

  conferir(
    'AtualizadoPor definido',
    depois.AtualizadoPor === 'ADMIN_CENTRAL'
  );

  // Confere os eventos de auditoria.
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaAuditoria = ss.getSheetByName('AUDITORIA');
  const ultimaLinha = abaAuditoria.getLastRow();

  const eventos = ultimaLinha >= 2
    ? abaAuditoria.getRange(2, 1, ultimaLinha - 1, 11).getValues()
    : [];

const alteracoes = eventos.filter(function(linha) {
  return String(linha[5]) === antes.VeiculoID &&
    String(linha[3]) === 'ALTERACAO' &&
    String(linha[8]) === 'SUCESSO' &&
    String(linha[9]) === 'SCRIPT';
});

conferir(
  'Evento ALTERACAO registrado com sucesso',
  alteracoes.length > 0
);

  Logger.log('UUID: ' + depois.VeiculoID);
  Logger.log('CODIGO: ' + depois.CodigoVeiculo);
  Logger.log('PLACA: ' + depois.Placa);
  Logger.log('OBSERVACOES: ' + depois.Observacoes);

  Logger.log('===== FIM DO TESTE =====');
}
function DCP_DIAGNOSTICAR_FALHA_EDICAO_235D() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Logger.log('===== DIAGNOSTICO DA EDICAO =====');

  // Inspeciona o veiculo.
  const aba = ss.getSheetByName('VEICULOS');
  const dados = aba.getDataRange().getValues();
  const cabecalhos = dados[0];

  const indice = dados.findIndex(function(linha, i) {
    return i > 0 &&
      String(linha[0]).trim() ===
      '0d9a4136-fc7c-4469-b6ea-223c4f746480';
  });

  if (indice < 1) {
    Logger.log('VEICULO NAO ENCONTRADO.');
    return;
  }

  const numeroLinha = indice + 1;
  const valores = dados[indice];

  Logger.log('===== DADOS ATUAIS DO VEICULO =====');

  const validacoes = aba.getRange(
    numeroLinha, 1, 1, cabecalhos.length
  ).getDataValidations()[0];

  cabecalhos.forEach(function(cabecalho, i) {
    const regra = validacoes[i];

    Logger.log(
      cabecalho +
      ' | VALOR: ' + String(valores[i]) +
      ' | VALIDACAO: ' +
      (regra
        ? regra.getCriteriaType() +
          ' | ' + JSON.stringify(regra.getCriteriaValues())
        : 'SEM VALIDACAO')
    );
  });

  // Inspeciona os eventos recentes da auditoria.
  const aud = ss.getSheetByName('AUDITORIA');
  const ultima = aud.getLastRow();

  Logger.log('===== ULTIMOS EVENTOS DE AUDITORIA =====');

  if (ultima >= 2) {
    const inicio = Math.max(2, ultima - 9);
    const quantidade = ultima - inicio + 1;

    const eventos = aud.getRange(
      inicio, 1, quantidade, 11
    ).getDisplayValues();

    eventos.forEach(function(linha) {
      Logger.log(JSON.stringify({
        AuditoriaID: linha[0],
        Acao: linha[3],
        Tabela: linha[4],
        RegistroID: linha[5],
        Resultado: linha[8],
        Detalhes: linha[10]
      }));
    });
  }

  Logger.log('===== FIM DO DIAGNOSTICO =====');
  Logger.log('OPERACAO SOMENTE DE LEITURA.');
}
function DCP_DIAGNOSTICAR_VALIDACAO_AUDITORIA_235D() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName('AUDITORIA');

  Logger.log('===== VALIDACAO DA AUDITORIA =====');

  const ultimaLinha = Math.max(2, aba.getLastRow());

  const cabecalhos = aba.getRange(1, 1, 1, 11)
    .getDisplayValues()[0];

  const validacoes = aba.getRange(2, 1, 1, 11)
    .getDataValidations()[0];

  cabecalhos.forEach(function(campo, i) {
    const regra = validacoes[i];

    Logger.log(
      campo + ' | ' +
      (regra
        ? regra.getCriteriaType() +
          ' | ' +
          JSON.stringify(regra.getCriteriaValues())
        : 'SEM VALIDACAO')
    );
  });

  Logger.log('===== ULTIMA LINHA =====');
  Logger.log('Numero: ' + ultimaLinha);

  Logger.log(
    JSON.stringify(
      aba.getRange(ultimaLinha, 1, 1, 11).getDisplayValues()[0]
    )
  );

  Logger.log('===== FIM =====');
  Logger.log('OPERACAO SOMENTE DE LEITURA.');
}
function DCP_CORRIGIR_AUDITORIA_PARCIAL_235D() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName('AUDITORIA');

  const uuidVeiculo =
    '0d9a4136-fc7c-4469-b6ea-223c4f746480';

  const ultima = aba.getLastRow();
  const dados = aba.getRange(2, 1, ultima - 1, 11).getValues();

  const encontradas = [];

  dados.forEach(function(linha, i) {
    if (
      String(linha[2]).trim() === 'ADMIN_CENTRAL' &&
      String(linha[0]).trim() !== '' &&
      String(linha[3]).trim() === '' &&
      String(linha[5]).trim() === ''
    ) {
      encontradas.push({
        numeroLinha: i + 2,
        linha: linha
      });
    }
  });

  const alvo = encontradas.filter(function(item) {
    return item.numeroLinha === 14;
  });

  if (alvo.length !== 1) {
    throw new Error(
      'A linha parcial esperada nao foi encontrada de forma unica.'
    );
  }

  const linha = alvo[0].linha;

  if (
    String(linha[0]).trim() !==
    '33c177e2-d34a-46e2-bc2e-f95327690ccb'
  ) {
    throw new Error('Identificador da linha inesperado.');
  }

  aba.getRange(14, 4, 1, 8).setValues([[
    'ERRO',
    'VEICULOS',
    uuidVeiculo,
    '',
    '',
    'FALHA',
    'SCRIPT',
    'Tentativa de edicao interrompida por validacao da auditoria.'
  ]]);

  SpreadsheetApp.flush();

  Logger.log('Linha parcial corrigida.');
  Logger.log('Registro de falha documentado na auditoria.');
}