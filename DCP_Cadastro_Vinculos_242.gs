/**
 * DRIVER CONTROL PRO
 * MODULO 2.4.2.2 - CADASTRO DE VINCULOS
 *
 * Revisao de seguranca:
 * - Recuperacao conservadora de tentativas.
 * - Conferencia integral dos 16 campos.
 * - Conferencia exata de CriadoEm e AtualizadoEm.
 * - Verificacao da chave antes de atualizar CONFIG.
 * - Bloqueio de reenvio inconclusivo.
 * - Preservacao de registros parciais.
 * - Correcao das validacoes da linha tecnica de CONFIG.
 * - Confirmacao de gravacao da reserva de idempotencia.
 *
 * IMPORTANTE:
 * - VIN-0001 permanece CANCELADO.
 * - Nao reutilizar chaves reservadas.
 * - Nao executar cadastro real durante os testes.
 */

const DCP_CADASTRO_VINCULOS_242 = {
  ABAS: {
    VINCULOS: 'VINCULOS',
    MOTORISTAS: 'MOTORISTAS',
    VEICULOS: 'VEICULOS',
    CONFIG: 'CONFIG',
    AUDITORIA: 'AUDITORIA'
  },
  PREFIXO: 'VIN',
  CHAVE_SEQUENCIA: 'SEQ_VIN',
  PREFIXO_IDEMPOTENCIA: 'IDEMP_VIN_',
  USUARIO: 'ADMIN_CENTRAL',
  TIPOS: ['PROPRIO', 'ALUGADO', 'CEDIDO'],
  PERIODICIDADES: ['DIARIA', 'SEMANAL', 'MENSAL'],
  CABECALHOS: {
    VINCULOS: [
      'VinculoID', 'CodigoVinculo', 'MotoristaID', 'VeiculoID',
      'TipoVinculo', 'InicioEm', 'FimEm',
      'CompartilhamentoAutorizado', 'ValorAluguel',
      'PeriodicidadeAluguel', 'Observacoes', 'Status',
      'CriadoEm', 'AtualizadoEm', 'CriadoPor', 'AtualizadoPor'
    ],
    MOTORISTAS: [
      'MotoristaID', 'CodigoMotorista', 'NomeCompleto',
      'NomeExibicao', 'Telefone', 'Email', 'DataCadastro',
      'Observacoes', 'Status', 'CriadoEm', 'AtualizadoEm',
      'CriadoPor', 'AtualizadoPor'
    ],
    VEICULOS: [
      'VeiculoID', 'CodigoVeiculo', 'Placa', 'Marca', 'Modelo',
      'AnoFabricacao', 'AnoModelo', 'TipoVeiculo', 'TipoPropulsao',
      'CombustivelPrincipal', 'CombustivelSecundario',
      'ProprietarioTipo', 'ProprietarioMotoristaID',
      'ProprietarioNome', 'CapacidadeTanqueLitros',
      'CapacidadeBateriaKWh', 'HodometroAtual', 'DataCadastro',
      'Observacoes', 'Status', 'CriadoEm', 'AtualizadoEm',
      'CriadoPor', 'AtualizadoPor'
    ],
    CONFIG: [
      'Chave', 'Valor', 'TipoValor', 'Descricao',
      'Editavel', 'AtualizadoEm'
    ],
    AUDITORIA: [
      'AuditoriaID', 'DataHora', 'UsuarioID', 'Acao', 'Tabela',
      'RegistroID', 'ValorAnterior', 'ValorNovo', 'Resultado',
      'Origem', 'Detalhes'
    ]
  }
};


/* =====================================================
   1. AUXILIARES
===================================================== */

function DCP_CV242_ABA(nome) {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nome);
  if (!aba) throw new Error('A aba ' + nome + ' nao foi encontrada.');
  return aba;
}


function DCP_CV242_CONFERIR_CABECALHOS(nomeAba, esperados) {
  const aba = DCP_CV242_ABA(nomeAba);
  const encontrados = aba.getRange(1, 1, 1, esperados.length).getValues()[0];
  const divergencias = [];

  esperados.forEach(function(nome, i) {
    if (String(encontrados[i] || '').trim() !== nome) {
      divergencias.push({
        coluna: i + 1,
        esperado: nome,
        encontrado: encontrados[i] || ''
      });
    }
  });

  if (divergencias.length) {
    throw new Error(
      'Cabecalhos divergentes na aba ' + nomeAba + ': ' +
      JSON.stringify(divergencias)
    );
  }

  return aba;
}


function DCP_CV242_UUID() {
  return Utilities.getUuid();
}


function DCP_CV242_NORMALIZAR_TEXTO(valor) {
  return String(valor === null || valor === undefined ? '' : valor).trim();
}


function DCP_CV242_BOOLEANO(valor) {
  return valor === true || valor === 1 ||
    String(valor).trim().toUpperCase() === 'TRUE' ||
    String(valor).trim().toUpperCase() === 'SIM';
}


// Retorna true/false somente para representacoes explicitas; null para qualquer outro valor.
function DCP_CV242_BOOLEANO_ESTRITO(valor) {
  if (valor === true || valor === 1) return true;
  if (valor === false || valor === 0) return false;
  const texto = String(valor === null || valor === undefined ? '' : valor)
    .trim().toUpperCase();
  if (texto === 'TRUE' || texto === 'SIM') return true;
  if (texto === 'FALSE' || texto === 'NAO' || texto === 'NÃO') return false;
  return null;
}


function DCP_CV242_CLONAR_OBJETO(objeto) {
  return JSON.parse(JSON.stringify(objeto));
}


function DCP_CV242_DATA(valor, campo) {
  if (valor === null || valor === undefined || valor === '') {
    return { valida: true, valor: null };
  }

  if (valor instanceof Date && !isNaN(valor.getTime())) {
    return { valida: true, valor: new Date(valor.getTime()) };
  }

  if (typeof valor === 'string') {
    const p = valor.trim().match(
      /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})(?::(\d{2}))?$/
    );

    if (p) {
      const ano = Number(p[1]);
      const mes = Number(p[2]);
      const dia = Number(p[3]);
      const hora = Number(p[4]);
      const minuto = Number(p[5]);
      const segundo = Number(p[6] || 0);

      const data = new Date(
        ano, mes - 1, dia, hora, minuto, segundo
      );

      const valida =
        data.getFullYear() === ano &&
        data.getMonth() === mes - 1 &&
        data.getDate() === dia &&
        data.getHours() === hora &&
        data.getMinutes() === minuto &&
        data.getSeconds() === segundo;

      return {
        valida: valida,
        valor: valida ? data : null
      };
    }
  }

  return {
    valida: false,
    valor: null,
    mensagem: 'Data/hora invalida: ' + campo
  };
}


function DCP_CV242_PERIODOS_SOBREPOEM(inicioA, fimA, inicioB, fimB) {
  const limiteA = fimA ? fimA.getTime() : Infinity;
  const limiteB = fimB ? fimB.getTime() : Infinity;

  return inicioA.getTime() < limiteB &&
    inicioB.getTime() < limiteA;
}


function DCP_CV242_DATA_IGUAL(valor, esperado) {
  const convertido = DCP_CV242_DATA(valor, 'comparacao');

  if (!convertido.valida) return false;
  if (!convertido.valor && !esperado) return true;
  if (!convertido.valor || !esperado) return false;

  return convertido.valor.getTime() === esperado.getTime();
}


/* =====================================================
   2. BUSCA E VALIDACAO DE CADASTROS
===================================================== */

function DCP_CV242_BUSCAR_UNICO(aba, colunaID, identificador, nomeCampo) {
  const ultimaLinha = aba.getLastRow();

  if (ultimaLinha < 2) {
    throw new Error(nomeCampo + ' nao encontrado.');
  }

  const dados = aba.getRange(
    2, 1, ultimaLinha - 1, aba.getLastColumn()
  ).getValues();

  const encontrados = dados.filter(function(linha) {
    return String(linha[colunaID] || '').trim() === identificador;
  });

  if (encontrados.length !== 1) {
    throw new Error(
      nomeCampo + ' deve corresponder a exatamente um registro. Encontrados: ' +
      encontrados.length
    );
  }

  return encontrados[0];
}


function DCP_CV242_VALIDAR_MOTORISTA(aba, motoristaID) {
  const linha = DCP_CV242_BUSCAR_UNICO(
    aba, 0, motoristaID, 'Motorista'
  );

  if (String(linha[8] || '').trim().toUpperCase() !== 'ATIVO') {
    throw new Error('O motorista selecionado nao esta ATIVO.');
  }

  return {
    id: String(linha[0]).trim(),
    codigo: String(linha[1]).trim(),
    nome: String(linha[2]).trim()
  };
}


function DCP_CV242_VALIDAR_VEICULO(aba, veiculoID) {
  const linha = DCP_CV242_BUSCAR_UNICO(
    aba, 0, veiculoID, 'Veiculo'
  );

  if (String(linha[19] || '').trim().toUpperCase() !== 'ATIVO') {
    throw new Error('O veiculo selecionado nao esta ATIVO.');
  }

  if (!String(linha[2] || '').trim()) {
    throw new Error('O veiculo nao possui placa cadastrada.');
  }

  return {
    id: String(linha[0]).trim(),
    codigo: String(linha[1]).trim(),
    placa: String(linha[2]).trim(),
    proprietarioTipo: String(linha[11] || '').trim().toUpperCase(),
    proprietarioMotoristaID: String(linha[12] || '').trim(),
    proprietarioNome: String(linha[13] || '').trim()
  };
}


/* =====================================================
   3. SEQUENCIA VIN
   Chamado sob o DocumentLock principal.
===================================================== */

function DCP_CV242_GERAR_CODIGO(abaConfig) {
  const abaVinculos = DCP_CV242_ABA('VINCULOS');
  const ultimaLinhaConfig = abaConfig.getLastRow();
  let linhaSequencia = 0;
  let valorSequencia = null;

  if (ultimaLinhaConfig >= 2) {
    const dadosConfig = abaConfig.getRange(
      2, 1, ultimaLinhaConfig - 1, 6
    ).getValues();

    dadosConfig.forEach(function(linha, indice) {
      if (String(linha[0] || '').trim() === 'SEQ_VIN') {
        if (linhaSequencia) {
          throw new Error('Chave SEQ_VIN duplicada na CONFIG.');
        }
        linhaSequencia = indice + 2;
        valorSequencia = linha[1];
      }
    });
  }

  // Descobre o maior VIN efetivamente existente. Qualquer código VIN
  // malformado torna a sequência ambígua e bloqueia a geração.
  let maiorCodigoExistente = 0;
  const ultimaLinhaVinculos = abaVinculos.getLastRow();

  if (ultimaLinhaVinculos >= 2) {
    const codigos = abaVinculos.getRange(
      2, 2, ultimaLinhaVinculos - 1, 1
    ).getValues();

    codigos.forEach(function(item, indice) {
      const codigoExistente = String(item[0] || '').trim();
      if (!codigoExistente || codigoExistente.indexOf('VIN-') !== 0) return;

      const correspondencia = codigoExistente.match(/^VIN-(\d{4,})$/);
      if (!correspondencia) {
        throw new Error(
          'Codigo de vinculo malformado na linha ' + (indice + 2) +
          '. Sequencia bloqueada para evitar reutilizacao.'
        );
      }

      maiorCodigoExistente = Math.max(
        maiorCodigoExistente,
        Number(correspondencia[1])
      );
    });
  }

  let sequenciaAtual;

  if (linhaSequencia) {
    if (
      valorSequencia === '' || valorSequencia === null ||
      valorSequencia === undefined ||
      typeof valorSequencia === 'boolean' ||
      !Number.isSafeInteger(Number(valorSequencia)) ||
      Number(valorSequencia) < 1
    ) {
      throw new Error(
        'SEQ_VIN possui valor invalido. Corrija manualmente apos conferencia; nenhum codigo foi reservado.'
      );
    }
    sequenciaAtual = Number(valorSequencia);
  } else {
    // Se a chave ainda nao existe, parte do maior VIN já gravado.
    sequenciaAtual = maiorCodigoExistente;
  }

  if (sequenciaAtual < maiorCodigoExistente) {
    throw new Error(
      'SEQ_VIN (' + sequenciaAtual + ') esta abaixo do maior codigo existente (' +
      maiorCodigoExistente + '). Geracao bloqueada para impedir reutilizacao.'
    );
  }

  const sequencia = sequenciaAtual + 1;
  if (!Number.isSafeInteger(sequencia)) {
    throw new Error('Limite numerico da sequencia VIN excedido.');
  }

  const codigo = 'VIN-' + String(sequencia).padStart(4, '0');
  const agora = new Date();

  if (linhaSequencia) {
    abaConfig.getRange(linhaSequencia, 2).setValue(sequencia);
    abaConfig.getRange(linhaSequencia, 6).setValue(agora);
  } else {
    abaConfig.appendRow([
      'SEQ_VIN', sequencia, 'NUMERO',
      'Sequencia independente de codigos de vinculos.', false, agora
    ]);
  }

  SpreadsheetApp.flush();

  // Leitura posterior obrigatória: não retornar código sem confirmar contador.
  const linhaConfirmacao = linhaSequencia || abaConfig.getLastRow();
  const confirmacao = abaConfig.getRange(linhaConfirmacao, 1, 1, 2).getValues()[0];
  if (
    String(confirmacao[0] || '').trim() !== 'SEQ_VIN' ||
    Number(confirmacao[1]) !== sequencia
  ) {
    throw new Error(
      'A reserva do codigo VIN nao foi confirmada. Nao tente repetir sem conferencia administrativa.'
    );
  }

  return codigo;
}

/* =====================================================
   4. IDEMPOTENCIA
===================================================== */

function DCP_CV242_LOCALIZAR_TENTATIVA(abaConfig, chave) {
  const chaveConfig =
    DCP_CADASTRO_VINCULOS_242.PREFIXO_IDEMPOTENCIA + chave;

  const ultimaLinha = abaConfig.getLastRow();

  if (ultimaLinha < 2) return null;

  const dados = abaConfig.getRange(
    2, 1, ultimaLinha - 1, 6
  ).getValues();

  const encontrados = [];

  dados.forEach(function(linha, indice) {
    if (String(linha[0] || '').trim() === chaveConfig) {
      encontrados.push({
        linha: indice + 2,
        valor: String(linha[1] || '')
      });
    }
  });

  if (encontrados.length > 1) {
    throw new Error(
      'Chave de tentativa duplicada na CONFIG. Exige analise administrativa.'
    );
  }

  return encontrados.length ? encontrados[0] : null;
}


/**
 * Cria uma reserva de idempotencia na CONFIG.
 *
 * A linha tecnica recebe a chave na coluna A e JSON na coluna C.
 * As validacoes sao removidas somente dessas duas celulas,
 * sem alterar as validacoes das demais linhas ou colunas.
 */
function DCP_CV242_CRIAR_TENTATIVA(abaConfig, chave, dadosResumo) {
  const anterior = DCP_CV242_LOCALIZAR_TENTATIVA(
    abaConfig, chave
  );

  if (anterior) {
    throw new Error('A chave de tentativa ja existe na CONFIG.');
  }

  const registro = {
    estado: 'PROCESSANDO',
    criadoEm: new Date().toISOString(),
    resumo: dadosResumo
  };

  const chaveConfig =
    DCP_CADASTRO_VINCULOS_242.PREFIXO_IDEMPOTENCIA + chave;

  const linhaDestino = abaConfig.getLastRow() + 1;

  // Protecao contra sobrescrita de qualquer conteudo existente.
  const destinoAtual = abaConfig.getRange(
    linhaDestino, 1, 1, 6
  ).getValues()[0];

  const linhaOcupada = destinoAtual.some(function(valor) {
    return valor !== '' && valor !== null && valor !== undefined;
  });

  if (linhaOcupada) {
    throw new Error(
      'A linha tecnica escolhida na CONFIG ja possui conteudo. Gravacao bloqueada.'
    );
  }

  // Remove somente as validacoes incompatíveis com a linha tecnica.
  abaConfig.getRange(linhaDestino, 1).clearDataValidations();
  abaConfig.getRange(linhaDestino, 3).clearDataValidations();

  abaConfig.getRange(linhaDestino, 1, 1, 6).setValues([[
    chaveConfig,
    JSON.stringify(registro),
    'JSON',
    'Controle de idempotencia de cadastro de vinculo.',
    false,
    new Date()
  ]]);

  SpreadsheetApp.flush();

  // Confirma a chave e o JSON diretamente na linha gravada.
  const confirmacaoLinha = abaConfig.getRange(
    linhaDestino, 1, 1, 6
  ).getValues()[0];

  if (String(confirmacaoLinha[0] || '').trim() !== chaveConfig) {
    throw new Error(
      'A reserva foi iniciada, mas a chave gravada diverge da chave esperada.'
    );
  }

  let jsonConfirmado;

  try {
    jsonConfirmado = JSON.parse(String(confirmacaoLinha[1] || ''));
  } catch (erro) {
    throw new Error(
      'A chave foi gravada, mas o JSON da reserva nao pode ser interpretado.'
    );
  }

  if (
    jsonConfirmado.estado !== 'PROCESSANDO' ||
    !jsonConfirmado.criadoEm
  ) {
    throw new Error(
      'A reserva foi gravada, mas seu conteudo nao corresponde ao esperado.'
    );
  }

  // Confirma também pela rotina geral de localização.
  const localizado = DCP_CV242_LOCALIZAR_TENTATIVA(
    abaConfig, chave
  );

  if (!localizado || localizado.linha !== linhaDestino) {
    throw new Error(
      'A reserva foi gravada, mas nao foi possivel confirmar sua localizacao.'
    );
  }

  return localizado.linha;
}


/**
 * Atualiza somente se a linha ainda pertencer a chave esperada.
 * A chave e conferida imediatamente antes da escrita.
 */
function DCP_CV242_ATUALIZAR_TENTATIVA(
  abaConfig, linha, chave, estado, resposta
) {
  if (!linha || linha < 2 || !chave) {
    throw new Error('Parametros invalidos para atualizar tentativa.');
  }

  const chaveEsperada =
    DCP_CADASTRO_VINCULOS_242.PREFIXO_IDEMPOTENCIA + chave;

  const chaveAtual = String(
    abaConfig.getRange(linha, 1).getValue() || ''
  ).trim();

  if (chaveAtual !== chaveEsperada) {
    throw new Error(
      'A linha da tentativa nao corresponde a chave esperada. Atualizacao bloqueada.'
          );
  }

  const registro = {
    estado: estado,
    atualizadoEm: new Date().toISOString(),
    resposta: resposta || null
  };

  abaConfig.getRange(linha, 2).setValue(
    JSON.stringify(registro)
  );

  abaConfig.getRange(linha, 6).setValue(new Date());

  SpreadsheetApp.flush();

  const confirmacao = DCP_CV242_LOCALIZAR_TENTATIVA(
    abaConfig, chave
  );

  if (!confirmacao || confirmacao.linha !== linha) {
    throw new Error(
      'Nao foi possivel confirmar a atualizacao da tentativa.'
    );
  }

  let registroConfirmado;

  try {
    registroConfirmado = JSON.parse(confirmacao.valor);
  } catch (erro) {
    throw new Error('Estado atualizado nao pode ser confirmado.');
  }

  if (registroConfirmado.estado !== estado) {
    throw new Error('Estado da tentativa diverge do esperado.');
  }

  return true;
}


function DCP_CV242_TRATAR_REENVIO(abaConfig, chave) {
  const existente = DCP_CV242_LOCALIZAR_TENTATIVA(
    abaConfig, chave
  );

  if (!existente) return null;

  let registro;

  try {
    registro = JSON.parse(existente.valor);
  } catch (erro) {
    return {
      sucesso: false,
      tentativaInconclusiva: true,
      codigo: 'TENTATIVA_INCONCLUSIVA',
      mensagem:
        'A chave existe na CONFIG, mas seu estado nao pode ser interpretado. ' +
        'Nao repita o cadastro. Solicite conferencia administrativa.'
    };
  }

  if (registro.estado === 'CONCLUIDO' && registro.resposta) {
    const respostaValida =
      registro.resposta &&
      typeof registro.resposta === 'object' &&
      registro.resposta.sucesso === true &&
      typeof registro.resposta.codigoVinculo === 'string' &&
      /^VIN-\d{4,}$/.test(registro.resposta.codigoVinculo) &&
      typeof registro.resposta.vinculoID === 'string' &&
      registro.resposta.vinculoID.trim() !== '';

    if (!respostaValida) {
      return {
        sucesso: false,
        tentativaInconclusiva: true,
        codigo: 'TENTATIVA_INCONCLUSIVA',
        mensagem: 'A chave esta marcada como CONCLUIDO, mas a resposta salva e invalida ou incompleta. Nao repita o cadastro; solicite conferencia administrativa.'
      };
    }

    const resposta = DCP_CV242_CLONAR_OBJETO(registro.resposta);

    resposta.reenvioIdentificado = true;
    resposta.mensagem =
      'Esta tentativa ja foi processada anteriormente. Nenhum novo vinculo foi criado. ' +
      (resposta.mensagem || '');

    return resposta;
  }

  return {
    sucesso: false,
    tentativaInconclusiva: true,
    codigo: 'TENTATIVA_INCONCLUSIVA',
    codigoVinculo: registro.resposta &&
      registro.resposta.codigoVinculo || '',
    mensagem:
      'Esta chave ja foi recebida pelo servidor e nao possui confirmacao conclusiva. ' +
      'Nenhum novo VIN sera gerado. Confira VINCULOS, CONFIG e AUDITORIA.'
  };
}


/**
 * Recupera a chave depois de qualquer falha.
 */
function DCP_CV242_RECUPERAR_TENTATIVA(abaConfig, chave) {
  if (!abaConfig || !chave) {
    return { estado: 'incerta', linha: 0 };
  }

  try {
    const tentativa = DCP_CV242_LOCALIZAR_TENTATIVA(
      abaConfig, chave
    );

    if (!tentativa) {
      return { estado: 'nao_localizada', linha: 0 };
    }

    return {
      estado: 'localizada',
      linha: tentativa.linha,
      valor: tentativa.valor
    };

  } catch (erro) {
    return {
      estado: 'incerta',
      linha: 0,
      erro: erro.message
    };
  }
}


/* =====================================================
   5. AUDITORIA
===================================================== */

function DCP_CV242_REGISTRAR_AUDITORIA(aba, evento) {
  const id = DCP_CV242_UUID();
  const agora = new Date();
  const linhaDestino = aba.getLastRow() + 1;

  aba.getRange(linhaDestino, 1, 1, 11).setValues([[
    id,
    agora,
    DCP_CADASTRO_VINCULOS_242.USUARIO,
    evento.acao,
    'VINCULOS',
    evento.registroID || '',
    evento.valorAnterior || '',
    evento.valorNovo || '',
    evento.resultado,
    'SCRIPT',
    evento.detalhes || ''
  ]]);

  SpreadsheetApp.flush();

  if (
    String(aba.getRange(linhaDestino, 1).getValue()).trim() !== id
  ) {
    throw new Error('Nao foi possivel confirmar a gravacao da auditoria.');
  }

  return id;
}


function DCP_CV242_AUDITORIA_EXISTE(
  aba, vinculoID, marcador, acaoEsperada
) {
  const ultimaLinha = aba.getLastRow();

  if (ultimaLinha < 2) return false;

  const dados = aba.getRange(
    2, 1, ultimaLinha - 1, 11
  ).getValues();

  return dados.some(function(linha) {
    return (
      String(linha[3] || '').trim() === (acaoEsperada || 'ERRO') &&
      String(linha[4] || '').trim() === 'VINCULOS' &&
      String(linha[5] || '').trim() === vinculoID &&
      String(linha[10] || '').indexOf(marcador) !== -1
    );
  });
}


/* =====================================================
   6. CONFLITOS DE PERIODOS
===================================================== */

function DCP_CV242_VERIFICAR_CONFLITOS(
  abaVinculos,
  motoristaID,
  veiculoID,
  inicioNovo,
  fimNovo,
  autorizacaoNova
) {
  const ultimaLinha = abaVinculos.getLastRow();

  if (ultimaLinha < 2) return [];

  const registros = abaVinculos.getRange(
    2, 1, ultimaLinha - 1, 16
  ).getValues();

  const conflitos = [];

  registros.forEach(function(linha, indice) {
    if (
      String(linha[11] || '').trim().toUpperCase() !== 'ATIVO'
    ) return;

    const motoristaExistente = String(linha[2] || '').trim();
    const veiculoExistente = String(linha[3] || '').trim();

    const inicio = DCP_CV242_DATA(
      linha[5], 'InicioEm do vinculo existente'
    );

    const fim = DCP_CV242_DATA(
      linha[6], 'FimEm do vinculo existente'
    );

    if (!inicio.valida || !fim.valida || !inicio.valor) {
      conflitos.push({
        tipo: 'REGISTRO_EXISTENTE_INVALIDO',
        codigo: String(linha[1] || ''),
        linha: indice + 2,
        mensagem:
          'Existe um vinculo ATIVO com datas invalidas. Corrija antes de continuar.'
      });
      return;
    }

    const sobrepoe = DCP_CV242_PERIODOS_SOBREPOEM(
      inicioNovo,
      fimNovo,
      inicio.valor,
      fim.valor
    );

    if (motoristaExistente === motoristaID && sobrepoe) {
      conflitos.push({
        tipo: 'CONFLITO_MOTORISTA',
        codigo: String(linha[1] || ''),
        linha: indice + 2,
        mensagem:
          'O motorista ja possui outro vinculo ativo nesse periodo.'
      });
    }

    if (veiculoExistente === veiculoID) {
      if (sobrepoe) {
        conflitos.push({
          tipo: 'CONFLITO_VEICULO',
          codigo: String(linha[1] || ''),
          linha: indice + 2,
          mensagem:
            'O veiculo ja possui outro vinculo ativo nesse periodo.'
        });
      } else {
        const autorizacaoExistente =
          DCP_CV242_BOOLEANO(linha[7]);

        if (!autorizacaoNova || !autorizacaoExistente) {
          conflitos.push({
            tipo: 'AUTORIZACAO_COMPARTILHAMENTO',
            codigo: String(linha[1] || ''),
            linha: indice + 2,
            mensagem:
              'Os vinculos envolvidos precisam de autorizacao explicita de compartilhamento.'
          });
        }
      }
    }
  });

  return conflitos;
}


/* =====================================================
   7. VALIDACAO DOS DADOS
===================================================== */

function DCP_CV242_VALIDAR_DADOS(dados) {
  if (!dados || typeof dados !== 'object') {
    throw new Error('Dados do vinculo nao informados.');
  }

  const motoristaID = DCP_CV242_NORMALIZAR_TEXTO(
    dados.MotoristaID
  );

  const veiculoID = DCP_CV242_NORMALIZAR_TEXTO(
    dados.VeiculoID
  );

  const tipo = DCP_CV242_NORMALIZAR_TEXTO(
    dados.TipoVinculo
  ).toUpperCase();

  const inicio = DCP_CV242_DATA(
    dados.InicioEm, 'InicioEm'
  );

  const fim = DCP_CV242_DATA(
    dados.FimEm, 'FimEm'
  );

  const observacoes = DCP_CV242_NORMALIZAR_TEXTO(
    dados.Observacoes
  );

  if (!motoristaID) throw new Error('Selecione um motorista.');
  if (!veiculoID) throw new Error('Selecione um veiculo.');

  if (DCP_CADASTRO_VINCULOS_242.TIPOS.indexOf(tipo) === -1) {
    throw new Error(
      'Tipo invalido. Utilize PROPRIO, ALUGADO ou CEDIDO.'
    );
  }

  if (!dados.InicioEm || !inicio.valida || !inicio.valor) {
    throw new Error(
      inicio.mensagem ||
      'Data e horario de inicio sao obrigatorios.'
    );
  }

  if (!fim.valida) throw new Error(fim.mensagem);

  if (
    fim.valor &&
    fim.valor.getTime() <= inicio.valor.getTime()
  ) {
    throw new Error('O termino deve ser posterior ao inicio.');
  }

  const autorizacao = dados.CompartilhamentoAutorizado;

  const autorizacaoInformada =
    autorizacao === true ||
    autorizacao === false ||
    autorizacao === 1 ||
    autorizacao === 0 ||
    ['TRUE', 'FALSE', 'SIM', 'NAO'].indexOf(
      String(autorizacao).trim().toUpperCase()
    ) !== -1;

  if (!autorizacaoInformada) {
    throw new Error(
      'Informe explicitamente se o compartilhamento esta autorizado.'
    );
  }

  const compartilhamentoAutorizado =
    DCP_CV242_BOOLEANO(autorizacao);

  let valorAluguel = '';

  if (
    dados.ValorAluguel !== '' &&
    dados.ValorAluguel !== null &&
    dados.ValorAluguel !== undefined
  ) {
    valorAluguel = Number(dados.ValorAluguel);

    if (!isFinite(valorAluguel) || valorAluguel <= 0) {
      throw new Error(
        'Valor do aluguel deve ser maior que zero.'
      );
    }
  }

  const periodicidade = DCP_CV242_NORMALIZAR_TEXTO(
    dados.PeriodicidadeAluguel
  ).toUpperCase();

  if (
    periodicidade &&
    DCP_CADASTRO_VINCULOS_242.PERIODICIDADES.indexOf(
      periodicidade
    ) === -1
  ) {
    throw new Error(
      'Periodicidade invalida. Utilize DIARIA, SEMANAL ou MENSAL.'
    );
  }

  if (Boolean(valorAluguel) !== Boolean(periodicidade)) {
    throw new Error(
      'Valor e periodicidade do aluguel devem ser informados juntos.'
    );
  }

  if (periodicidade && tipo !== 'ALUGADO') {
    throw new Error(
      'Dados de aluguel so podem ser informados para vinculo ALUGADO.'
    );
  }

  if (observacoes.length > 2000) {
    throw new Error(
      'Observacoes devem ter no maximo 2.000 caracteres.'
    );
  }

  return {
    motoristaID: motoristaID,
    veiculoID: veiculoID,
    tipo: tipo,
    inicio: inicio.valor,
    fim: fim.valor,
    compartilhamentoAutorizado: compartilhamentoAutorizado,
    valorAluguel: valorAluguel,
    periodicidade: periodicidade,
    observacoes: observacoes
  };
}


/* =====================================================
   8. VALIDACAO DAS LISTAS SUSPENSAS
===================================================== */

function DCP_CV242_VALOR_ACEITO_VALIDACAO(
  aba, linha, coluna, valor
) {
  const regra = aba.getRange(
    linha, coluna
  ).getDataValidation();

  if (!regra) return true;

  const tipo = String(regra.getCriteriaType());
  const criterios = regra.getCriteriaValues();

  if (tipo === 'VALUE_IN_LIST') {
    const lista = criterios && criterios[0]
      ? criterios[0]
      : [];

    return lista.some(function(item) {
      return String(item).trim().toUpperCase() ===
        String(valor).trim().toUpperCase();
    });
  }

  if (tipo === 'VALUE_IN_RANGE') {
    const intervalo = criterios && criterios[0];

    if (!intervalo || typeof intervalo.getValues !== 'function') {
      throw new Error(
        'Nao foi possivel interpretar a lista suspensa da coluna ' +
        coluna + '.'
      );
    }

    const lista = intervalo.getValues().flat();

    return lista.some(function(item) {
      return String(item).trim().toUpperCase() ===
        String(valor).trim().toUpperCase();
    });
  }

  return true;
}


function DCP_CV242_VALIDAR_DESTINO_GRAVACAO(
  abaVinculos,
  abaAuditoria,
  linhaDestino,
  validado
) {
  const verificacoesVinculo = [
    { coluna: 5, valor: validado.tipo, campo: 'TipoVinculo' },
    { coluna: 12, valor: 'ATIVO', campo: 'Status' }
  ];

  if (validado.periodicidade) {
    verificacoesVinculo.push({
      coluna: 10,
      valor: validado.periodicidade,
      campo: 'PeriodicidadeAluguel'
    });
  }

  verificacoesVinculo.forEach(function(item) {
    if (!DCP_CV242_VALOR_ACEITO_VALIDACAO(
      abaVinculos,
      linhaDestino,
      item.coluna,
      item.valor
    )) {
      throw new Error(
        'A validacao da planilha nao aceita "' +
        item.valor + '" no campo ' + item.campo +
        '. Nenhum vinculo foi gravado.'
      );
    }
  });

  const linhaAuditoria = Math.max(
    2, abaAuditoria.getLastRow() + 1
  );

  const verificacoesAuditoria = [
    { coluna: 4, valor: 'INSERCAO', campo: 'Acao INSERCAO' },
    { coluna: 4, valor: 'ERRO', campo: 'Acao ERRO' },
    { coluna: 9, valor: 'SUCESSO', campo: 'Resultado SUCESSO' },
    { coluna: 9, valor: 'FALHA', campo: 'Resultado FALHA' },
    { coluna: 10, valor: 'SCRIPT', campo: 'Origem SCRIPT' }
  ];

  verificacoesAuditoria.forEach(function(item) {
    if (!DCP_CV242_VALOR_ACEITO_VALIDACAO(
      abaAuditoria,
      linhaAuditoria,
      item.coluna,
      item.valor
    )) {
      throw new Error(
        'A validacao da aba AUDITORIA nao aceita "' +
        item.valor + '" no campo ' + item.campo +
        '. Cadastro interrompido antes da gravacao.'
      );
    }
  });

  return true;
}


/* =====================================================
   9. LOCALIZACAO E CONFERENCIA INTEGRAL
===================================================== */

function DCP_CV242_LOCALIZAR_REGISTRO(
  aba, vinculoID, codigo
) {
  const ultimaLinha = aba.getLastRow();

  if (ultimaLinha < 2) return [];

  const dados = aba.getRange(
        2, 1, ultimaLinha - 1, 16
  ).getValues();

  const encontrados = [];

  dados.forEach(function(linha, indice) {
    const id = String(linha[0] || '').trim();
    const cod = String(linha[1] || '').trim();

    if (
      (vinculoID && id === vinculoID) ||
      (codigo && cod === codigo)
    ) {
      encontrados.push({
        linha: indice + 2,
        dados: linha
      });
    }
  });

  return encontrados;
}


/**
 * Confere os campos funcionais e os metadados do registro.
 * Compara CriadoEm e AtualizadoEm com a data exata esperada.
 */
function DCP_CV242_REGISTRO_COMPLETO(registro, esperado) {
  if (!registro || !registro.dados || !esperado) return false;

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

    if (
      String(
        l[coluna] === null || l[coluna] === undefined
          ? ''
          : l[coluna]
      ).trim() !==
      String(
        valor === null || valor === undefined
          ? ''
          : valor
      ).trim()
    ) {
      return false;
    }
  }

  if (!DCP_CV242_DATA_IGUAL(l[5], esperado.inicio)) return false;
  if (!DCP_CV242_DATA_IGUAL(l[6], esperado.fim)) return false;

  const compartilhamentoGravado = DCP_CV242_BOOLEANO_ESTRITO(l[7]);
  if (compartilhamentoGravado === null) return false;
  if (compartilhamentoGravado !== esperado.compartilhamentoAutorizado) return false;

  const aluguelGravado =
    l[8] === '' || l[8] === null || l[8] === undefined
      ? ''
      : Number(l[8]);

  if (aluguelGravado !== esperado.valorAluguel) return false;

  if (!esperado.criadoEm) return false;

  if (!DCP_CV242_DATA_IGUAL(l[12], esperado.criadoEm)) {
    return false;
  }

  if (!DCP_CV242_DATA_IGUAL(l[13], esperado.criadoEm)) {
    return false;
  }

  return true;
}


/* =====================================================
   10. RESPOSTA DE CADASTRO
===================================================== */

function DCP_CV242_MONTAR_RESPOSTA_SUCESSO(
  dados, auditoriaPendente, recuperacaoAutomatica
) {
  return {
    sucesso: true,
    auditoriaPendente: Boolean(auditoriaPendente),
    codigoVinculo: dados.codigo,
    vinculoID: dados.vinculoID,
    auditoriaID: dados.auditoriaID || '',
    motorista: dados.motorista,
    veiculo: dados.veiculo,
    tipoVinculo: dados.tipo,
    status: 'ATIVO',
    recuperacaoAutomatica: Boolean(recuperacaoAutomatica),
    mensagem: auditoriaPendente
      ? 'Vinculo gravado e conferido. A auditoria precisa ser verificada. Nao repita o cadastro.'
      : 'Vinculo ' + dados.codigo +
        ' cadastrado, conferido e auditado com sucesso.'
  };
}


/* =====================================================
   11. CADASTRO PRINCIPAL
===================================================== */

function DCP_CADASTRAR_VINCULO_242(dados) {
  const lock = LockService.getDocumentLock();

  if (!lock.tryLock(30000)) {
    return {
      sucesso: false,
      gravacaoRealizada: false,
      tentativaInconclusiva: false,
      mensagem:
        'Outra operacao esta em andamento. Nenhum cadastro foi iniciado. Aguarde e tente novamente.'
    };
  }

  let abaVinculos = null;
  let abaConfig = null;
  let abaAuditoria = null;

  let chaveTentativa = '';
  let linhaTentativa = 0;

  let reservaTentada = false;
  let gravacaoIniciada = false;

  let codigo = '';
  let vinculoID = '';
  let linhaGravada = 0;

  let validado = null;
  let motorista = null;
  let veiculo = null;

  // Variavel externa ao try para permitir recuperacao no catch.
  let dataCriacaoRegistro = null;

  try {
    abaVinculos = DCP_CV242_CONFERIR_CABECALHOS(
      'VINCULOS',
      DCP_CADASTRO_VINCULOS_242.CABECALHOS.VINCULOS
    );

    const abaMotoristas = DCP_CV242_CONFERIR_CABECALHOS(
      'MOTORISTAS',
      DCP_CADASTRO_VINCULOS_242.CABECALHOS.MOTORISTAS
    );

    const abaVeiculos = DCP_CV242_CONFERIR_CABECALHOS(
      'VEICULOS',
      DCP_CADASTRO_VINCULOS_242.CABECALHOS.VEICULOS
    );

    abaConfig = DCP_CV242_CONFERIR_CABECALHOS(
      'CONFIG',
      DCP_CADASTRO_VINCULOS_242.CABECALHOS.CONFIG
    );

    abaAuditoria = DCP_CV242_CONFERIR_CABECALHOS(
      'AUDITORIA',
      DCP_CADASTRO_VINCULOS_242.CABECALHOS.AUDITORIA
    );

    chaveTentativa = DCP_CV242_NORMALIZAR_TEXTO(
      dados && dados.ChaveTentativa
    );

    if (!/^[a-zA-Z0-9-]{20,80}$/.test(chaveTentativa)) {
      throw new Error(
        'Chave de tentativa ausente ou invalida. Feche e abra um novo formulario.'
      );
    }

    /*
     * A consulta da chave tambem e uma operacao critica.
     * Se a leitura falhar, o resultado sera inconclusivo.
     */
    reservaTentada = true;

    const reenvio = DCP_CV242_TRATAR_REENVIO(
      abaConfig,
      chaveTentativa
    );

    if (reenvio) return reenvio;

    /*
     * A chave foi consultada e nao foi encontrada.
     * Ainda nao existe reserva.
     */
    reservaTentada = false;

    validado = DCP_CV242_VALIDAR_DADOS(dados);

    motorista = DCP_CV242_VALIDAR_MOTORISTA(
      abaMotoristas,
      validado.motoristaID
    );

    veiculo = DCP_CV242_VALIDAR_VEICULO(
      abaVeiculos,
      validado.veiculoID
    );

    if (
      validado.tipo === 'PROPRIO' &&
      (
        veiculo.proprietarioTipo !== 'MOTORISTA' ||
        veiculo.proprietarioMotoristaID !== motorista.id
      )
    ) {
      throw new Error(
        'Vinculo PROPRIO nao permitido: o motorista nao e o proprietario cadastrado do veiculo.'
      );
    }

    const conflitos = DCP_CV242_VERIFICAR_CONFLITOS(
      abaVinculos,
      motorista.id,
      veiculo.id,
      validado.inicio,
      validado.fim,
      validado.compartilhamentoAutorizado
    );

    if (conflitos.length) {
      return {
        sucesso: false,
        codigo: 'CONFLITO_DE_VINCULO',
        gravacaoRealizada: false,
        tentativaInconclusiva: false,
        mensagem:
          'Cadastro bloqueado por conflito ou falta de autorizacao.',
        conflitos: conflitos
      };
    }

    linhaGravada = abaVinculos.getLastRow() + 1;

    DCP_CV242_VALIDAR_DESTINO_GRAVACAO(
      abaVinculos,
      abaAuditoria,
      linhaGravada,
      validado
    );

    /*
     * A partir deste ponto, qualquer falha e inconclusiva.
     * Nao classificamos como falha segura para nova tentativa.
     */
    reservaTentada = true;

    linhaTentativa = DCP_CV242_CRIAR_TENTATIVA(
      abaConfig,
      chaveTentativa,
      {
        motoristaID: motorista.id,
        veiculoID: veiculo.id,
        tipo: validado.tipo,
        inicio: validado.inicio.toISOString()
      }
    );

    codigo = DCP_CV242_GERAR_CODIGO(abaConfig);
    vinculoID = DCP_CV242_UUID();

    dataCriacaoRegistro = new Date();
    const agora = dataCriacaoRegistro;

    const novaLinha = [
      vinculoID,
      codigo,
      motorista.id,
      veiculo.id,
      validado.tipo,
      validado.inicio,
      validado.fim || '',
      validado.compartilhamentoAutorizado,
      validado.valorAluguel,
      validado.periodicidade,
      validado.observacoes,
      'ATIVO',
      agora,
      agora,
      DCP_CADASTRO_VINCULOS_242.USUARIO,
      DCP_CADASTRO_VINCULOS_242.USUARIO
    ];

    gravacaoIniciada = true;

    abaVinculos.getRange(
      linhaGravada, 1, 1, 16
    ).setValues([novaLinha]);

    SpreadsheetApp.flush();

    const registroLido = abaVinculos.getRange(
      linhaGravada, 1, 1, 16
    ).getValues()[0];

    const registroConfirmado = {
      linha: linhaGravada,
      dados: registroLido
    };

    const esperado = {
      vinculoID: vinculoID,
      codigo: codigo,
      motoristaID: motorista.id,
      veiculoID: veiculo.id,
      tipo: validado.tipo,
      inicio: validado.inicio,
      fim: validado.fim,
      compartilhamentoAutorizado:
        validado.compartilhamentoAutorizado,
      valorAluguel: validado.valorAluguel,
      periodicidade: validado.periodicidade,
      observacoes: validado.observacoes,
      criadoEm: agora
    };

    if (!DCP_CV242_REGISTRO_COMPLETO(
      registroConfirmado,
      esperado
    )) {
      throw new Error(
        'A gravacao foi iniciada, mas a conferencia integral dos 16 campos falhou.'
      );
    }

    const marcador = '[DCP_VINCULO_' + vinculoID + ']';

    const detalhes =
      marcador +
      ' Cadastro ' + codigo +
      '. Motorista: ' + motorista.codigo +
      '. Veiculo: ' + veiculo.codigo +
      '. Tipo: ' + validado.tipo +
      '. Inicio: ' + validado.inicio.toISOString() +
      '. Termino: ' +
      (validado.fim
        ? validado.fim.toISOString()
        : 'SEM TERMINO') +
      '. Compartilhamento autorizado: ' +
      validado.compartilhamentoAutorizado;

    let auditoriaID = '';
    let auditoriaPendente = false;

    try {
      auditoriaID = DCP_CV242_REGISTRAR_AUDITORIA(
        abaAuditoria,
        {
          acao: 'INSERCAO',
          registroID: vinculoID,
          valorAnterior: '',
          valorNovo: 'ATIVO',
          resultado: 'SUCESSO',
          detalhes: detalhes
        }
      );

    } catch (erroAuditoria) {
      auditoriaPendente = !DCP_CV242_AUDITORIA_EXISTE(
        abaAuditoria,
        vinculoID,
        marcador,
        'INSERCAO'
      );
    }

    const resposta = DCP_CV242_MONTAR_RESPOSTA_SUCESSO(
      {
        codigo: codigo,
        vinculoID: vinculoID,
        auditoriaID: auditoriaID,
        motorista: motorista.nome,
        veiculo: veiculo.codigo + ' | ' + veiculo.placa,
        tipo: validado.tipo
      },
      auditoriaPendente,
      false
    );

    DCP_CV242_ATUALIZAR_TENTATIVA(
      abaConfig,
      linhaTentativa,
      chaveTentativa,
      'CONCLUIDO',
      resposta
    );

    return resposta;

  } catch (erro) {

    let recuperacao = {
      estado: reservaTentada ? 'incerta' : 'nao_localizada',
      linha: 0
    };

    /*
     * A recuperacao e executada mesmo se a criacao da tentativa
     * falhou antes de devolver o numero da linha.
     */
    if (reservaTentada && abaConfig && chaveTentativa) {
      recuperacao = DCP_CV242_RECUPERAR_TENTATIVA(
        abaConfig,
        chaveTentativa
      );

      if (recuperacao.estado === 'localizada') {
        linhaTentativa = recuperacao.linha;
      }
    }

    let registrosEncontrados = [];

    if (gravacaoIniciada && vinculoID && codigo) {
      try {
        registrosEncontrados = DCP_CV242_LOCALIZAR_REGISTRO(
          DCP_CV242_ABA('VINCULOS'),
          vinculoID,
          codigo
        );
      } catch (erroConsulta) {
        registrosEncontrados = [];
      }
    }

    /*
     * Se o registro completo existe e a chave tambem existe,
     * tentamos recuperar o estado CONCLUIDO.
     */
    if (
      registrosEncontrados.length === 1 &&
      recuperacao.estado === 'localizada' &&
      validado &&
      motorista &&
      veiculo &&
      dataCriacaoRegistro
    ) {
      const esperado = {
        vinculoID: vinculoID,
        codigo: codigo,
        motoristaID: motorista.id,
        veiculoID: veiculo.id,
        tipo: validado.tipo,
        inicio: validado.inicio,
        fim: validado.fim,
        compartilhamentoAutorizado:
          validado.compartilhamentoAutorizado,
        valorAluguel: validado.valorAluguel,
        periodicidade: validado.periodicidade,
        observacoes: validado.observacoes,
        criadoEm: dataCriacaoRegistro
      };

      if (DCP_CV242_REGISTRO_COMPLETO(
        registrosEncontrados[0],
        esperado
      )) {
        let auditoriaPendente = true;

        try {
          const marcador =
            '[DCP_VINCULO_' + vinculoID + ']';

          auditoriaPendente = !DCP_CV242_AUDITORIA_EXISTE(
            DCP_CV242_ABA('AUDITORIA'),
            vinculoID,
            marcador,
            'INSERCAO'
          );
        } catch (erroAuditoriaConsulta) {
          auditoriaPendente = true;
        }

        const respostaRecuperada =
          DCP_CV242_MONTAR_RESPOSTA_SUCESSO(
            {
              codigo: codigo,
              vinculoID: vinculoID,
              motorista: motorista.nome,
              veiculo: veiculo.codigo + ' | ' + veiculo.placa,
              tipo: validado.tipo
            },
            auditoriaPendente,
            true
          );

        try {
          DCP_CV242_ATUALIZAR_TENTATIVA(
            abaConfig,
            linhaTentativa,
            chaveTentativa,
            'CONCLUIDO',
            respostaRecuperada
          );

          return respostaRecuperada;

        } catch (erroConclusao) {
          return {
            sucesso: false,
            tentativaInconclusiva: true,
            gravacaoParcial: true,
            codigoVinculo: codigo,
            vinculoID: vinculoID,
            mensagem:
              'O vinculo foi localizado e conferido integralmente, mas nao foi possivel confirmar o estado da chave. ' +
              'Nao repita o cadastro. Exige conferencia administrativa.',
            detalhe: erroConclusao.message
          };
        }
      }
    }

    let respostaErro;

    if (registrosEncontrados.length > 1) {
      respostaErro = {
        sucesso: false,
        tentativaInconclusiva: true,
        gravacaoParcial: true,
        codigoVinculo: codigo,
        vinculoID: vinculoID,
        codigo: 'REGISTROS_DUPLICADOS',
        mensagem:
          'Foram encontrados registros duplicados com o mesmo identificador. Nao tente cadastrar novamente. Exige analise manual.'
      };

    } else if (registrosEncontrados.length === 1) {
      respostaErro = {
        sucesso: false,
        tentativaInconclusiva: true,
        gravacaoParcial: true,
        codigoVinculo: codigo,
        vinculoID: vinculoID,
                linha: registrosEncontrados[0].linha,
        codigo: 'GRAVACAO_PARCIAL',
        mensagem:
          'Foi detectada uma gravacao parcial ou inconclusiva. A linha foi preservada. Nao tente cadastrar novamente.',
        detalhe: erro.message
      };

    } else if (reservaTentada) {
      respostaErro = {
        sucesso: false,
        tentativaInconclusiva: true,
        gravacaoRealizada: null,
        codigoVinculo: codigo,
        vinculoID: vinculoID,
        codigo: 'ESTADO_DA_TENTATIVA_INCERTO',
        mensagem:
          'A reserva da chave foi iniciada, mas nao foi possivel confirmar integralmente o resultado. ' +
          'Por seguranca, nao repita o cadastro. Solicite conferencia administrativa.',
        detalhe: erro.message,
        detalheRecuperacao: recuperacao.erro || ''
      };

    } else {
      respostaErro = {
        sucesso: false,
        tentativaInconclusiva: false,
        gravacaoRealizada: false,
        codigo: 'FALHA_ANTES_DA_RESERVA',
        mensagem: erro.message,
        orientacao:
          'A tentativa nao chegou a iniciar a reserva da chave. Corrija a causa antes de uma nova operacao.'
      };
    }

    /*
     * Se a chave foi localizada, tentamos marcar INCONCLUSIVO.
     * Nunca alteramos uma linha sem conferir a chave.
     */
    if (
      reservaTentada &&
      recuperacao.estado === 'localizada' &&
      linhaTentativa
    ) {
      try {
        DCP_CV242_ATUALIZAR_TENTATIVA(
          abaConfig,
          linhaTentativa,
          chaveTentativa,
          'INCONCLUSIVO',
          respostaErro
        );

      } catch (erroAtualizacao) {
        respostaErro.detalheControleTentativa =
          erroAtualizacao.message;
      }
    }

    /*
     * Se houve inicio de gravacao e a linha ficou parcial,
     * registramos a falha na auditoria sem apagar o registro.
     */
    if (
      registrosEncontrados.length === 1 &&
      vinculoID &&
      abaAuditoria
    ) {
      try {
        const marcador =
          '[DCP_FALHA_PARCIAL_' + vinculoID + ']';

        if (!DCP_CV242_AUDITORIA_EXISTE(
          abaAuditoria,
          vinculoID,
          marcador,
          'ERRO'
        )) {
          DCP_CV242_REGISTRAR_AUDITORIA(
            abaAuditoria,
            {
              acao: 'ERRO',
              registroID: vinculoID,
              valorAnterior: '',
              valorNovo: 'GRAVACAO_PARCIAL',
              resultado: 'FALHA',
              detalhes:
                marcador +
                ' Falha no cadastro ' + codigo +
                '. Erro: ' + erro.message
            }
          );
        }

      } catch (erroAuditoriaParcial) {
        respostaErro.detalheAuditoria =
          erroAuditoriaParcial.message;
      }
    }

    return respostaErro;

  } finally {
    lock.releaseLock();
  }
}


/* =====================================================
   12. DIAGNOSTICO DE ESTRUTURA
===================================================== */

function DCP_DIAGNOSTICAR_ESTRUTURA_CADASTRO_VINCULOS_242() {
  const resultado = {};

  Object.keys(
    DCP_CADASTRO_VINCULOS_242.CABECALHOS
  ).forEach(function(chave) {
    const nome = DCP_CADASTRO_VINCULOS_242.ABAS[chave];

    try {
      const aba = DCP_CV242_CONFERIR_CABECALHOS(
        nome,
        DCP_CADASTRO_VINCULOS_242.CABECALHOS[chave]
      );

      resultado[chave] = {
        existe: true,
        cabecalhosCorretos: true,
        ultimaLinha: aba.getLastRow()
      };

    } catch (erro) {
      resultado[chave] = {
        existe: false,
        cabecalhosCorretos: false,
        erro: erro.message
      };
    }
  });

  const sucesso = Object.keys(resultado).every(function(chave) {
    return resultado[chave].cabecalhosCorretos;
  });

  const resposta = {
    sucesso: sucesso,
    somenteLeitura: true,
    estrutura: resultado
  };

  Logger.log(JSON.stringify(resposta, null, 2));
  return resposta;
}


/* =====================================================
   13. TESTES DE VALIDACAO
===================================================== */

function DCP_TESTAR_VALIDACOES_CADASTRO_VINCULOS_242() {
  const base = {
    MotoristaID: 'MOTORISTA_TESTE',
    VeiculoID: 'VEICULO_TESTE',
    TipoVinculo: 'PROPRIO',
    InicioEm: '2026-11-10T06:00',
    FimEm: '2026-11-10T14:00',
    CompartilhamentoAutorizado: false,
    ValorAluguel: '',
    PeriodicidadeAluguel: '',
    Observacoes: ''
  };

  const casos = [
    {
      nome: 'Cadastro basico valido',
      dados: Object.assign({}, base),
      esperado: true
    },
    {
      nome: 'Tipo invalido',
      dados: Object.assign({}, base, {
        TipoVinculo: 'OUTRO'
      }),
      esperado: false
    },
    {
      nome: 'Inicio ausente',
      dados: Object.assign({}, base, {
        InicioEm: ''
      }),
      esperado: false
    },
    {
      nome: 'Termino anterior',
      dados: Object.assign({}, base, {
        FimEm: '2026-11-10T05:00'
      }),
      esperado: false
    },
    {
      nome: 'Autorizacao ausente',
      dados: Object.assign({}, base, {
        CompartilhamentoAutorizado: ''
      }),
      esperado: false
    },
    {
      nome: 'Aluguel sem periodicidade',
      dados: Object.assign({}, base, {
        TipoVinculo: 'ALUGADO',
        ValorAluguel: 1000
      }),
      esperado: false
    },
    {
      nome: 'Aluguel coerente',
      dados: Object.assign({}, base, {
        TipoVinculo: 'ALUGADO',
        ValorAluguel: 1000,
        PeriodicidadeAluguel: 'MENSAL'
      }),
      esperado: true
    },
    {
      nome: 'Periodicidade invalida',
      dados: Object.assign({}, base, {
        TipoVinculo: 'ALUGADO',
        ValorAluguel: 1000,
        PeriodicidadeAluguel: 'ANUAL'
      }),
      esperado: false
    },
    {
      nome: 'Observacoes acima do limite',
      dados: Object.assign({}, base, {
        Observacoes: 'X'.repeat(2001)
      }),
      esperado: false
    }
  ];

  const resultados = casos.map(function(caso) {
    let obtido = true;
    let mensagemErro = '';

    try {
      DCP_CV242_VALIDAR_DADOS(caso.dados);
    } catch (erro) {
      obtido = false;
      mensagemErro = erro.message;
    }

    return {
      teste: caso.nome,
      esperado: caso.esperado,
      obtido: obtido,
      aprovado: obtido === caso.esperado,
      mensagem: mensagemErro
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
    somenteValidacaoEmMemoria: true,
    dadosDaPlanilhaAlterados: false
  };

  Logger.log(JSON.stringify(resposta, null, 2));
  return resposta;
}


/* =====================================================
   14. DIAGNOSTICOS HISTORICOS VIN-0001
===================================================== */

function DCP_DIAGNOSTICAR_VIN_FALHO_245() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaV = ss.getSheetByName('VINCULOS');
  const abaC = ss.getSheetByName('CONFIG');

  if (!abaV || !abaC) {
    throw new Error('VINCULOS ou CONFIG nao encontrada.');
  }

  const dados = abaV.getDataRange().getValues();
  const encontrados = [];

  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][1]).trim() === 'VIN-0001') {
      encontrados.push({
        linha: i + 1,
        dados: dados[i].map(function(v) {
          return v instanceof Date ? v.toISOString() : v;
        })
      });
    }
  }

  const config = abaC.getDataRange().getValues();

  const sequencias = config.filter(function(linha) {
    return String(linha[0]).trim() === 'SEQ_VIN';
  }).map(function(linha) {
    return {
      valor: linha[1],
      tipo: linha[2]
    };
  });

  const resultado = {
    somenteLeitura: true,
    registrosVIN0001: encontrados,
    sequenciaVIN: sequencias
  };

  Logger.log(JSON.stringify(resultado, null, 2));
  return resultado;
}


function DCP_CONFERIR_AUDITORIA_VIN_245() {
  const aba = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('AUDITORIA');

  if (!aba) {
    throw new Error('Aba AUDITORIA nao encontrada.');
  }

  const dados = aba.getDataRange().getValues();
  const encontrados = [];

  for (let i = 1; i < dados.length; i++) {
    if (
      String(dados[i][3]).trim() === 'ERRO' &&
      String(dados[i][4]).trim() === 'VINCULOS' &&
      String(dados[i][5]).trim() ===
        '7cbe4eb2-0f5d-42d6-a7cd-6e4a5be30d4d'
    ) {
      encontrados.push({
        linha: i + 1,
        dados: dados[i].map(function(v) {
          return v instanceof Date ? v.toISOString() : v;
        })
      });
    }
  }

  const resultado = {
    somenteLeitura: true,
    eventosEncontrados: encontrados.length,
    eventos: encontrados
  };

  Logger.log(JSON.stringify(resultado, null, 2));
  return resultado;
}


/* =====================================================
   15. DIAGNOSTICO DA LINHA TECNICA CONFIG
   SOMENTE LEITURA
===================================================== */

function DCP_CV242_DIAGNOSTICAR_CONFIG_DESTINO() {
  const aba = DCP_CV242_CONFERIR_CABECALHOS(
    'CONFIG',
    DCP_CADASTRO_VINCULOS_242.CABECALHOS.CONFIG
  );

  const linha = aba.getLastRow() + 1;
  const colunas = [1, 2, 3, 4, 5, 6];

  const diagnostico = colunas.map(function(coluna) {
    const celula = aba.getRange(linha, coluna);
    const regra = celula.getDataValidation();

    return {
      coluna: coluna,
      valorAtual: celula.getValue(),
      possuiValidacao: Boolean(regra),
      tipoValidacao: regra
        ? String(regra.getCriteriaType())
        : '',
      criterios: regra
        ? regra.getCriteriaValues().map(function(item) {
            if (Array.isArray(item)) return item;
            if (item && typeof item.getA1Notation === 'function') {
              return item.getA1Notation();
            }
            return String(item);
          })
        : []
    };
  });

  const resposta = {
    sucesso: true,
    somenteLeitura: true,
    linhaTecnicaPrevista: linha,
    colunas: diagnostico
  };

  Logger.log(JSON.stringify(resposta, null, 2));
  return resposta;
}


/* =====================================================
   16. TESTE DE COMPARACAO DE HORARIOS
   SOMENTE MEMORIA
===================================================== */

function DCP_CV242_TESTAR_COMPARACAO_HORARIOS() {
  const instante = new Date(2026, 10, 4, 7, 0, 0, 0);

  const registroBase = [
    'UUID-TESTE',
    'VIN-TESTE',
    'MOTORISTA-TESTE',
    'VEICULO-TESTE',
    'PROPRIO',
    new Date(2026, 10, 4, 7, 0, 0, 0),
    new Date(2026, 10, 4, 12, 0, 0, 0),
    true,
    '',
    '',
    'Teste',
    'ATIVO',
    new Date(instante.getTime()),
    new Date(instante.getTime()),
    'ADMIN_CENTRAL',
    'ADMIN_CENTRAL'
  ];

  const esperado = {
    vinculoID: 'UUID-TESTE',
    codigo: 'VIN-TESTE',
    motoristaID: 'MOTORISTA-TESTE',
    veiculoID: 'VEICULO-TESTE',
    tipo: 'PROPRIO',
    inicio: new Date(2026, 10, 4, 7, 0, 0, 0),
    fim: new Date(2026, 10, 4, 12, 0, 0, 0),
    compartilhamentoAutorizado: true,
    valorAluguel: '',
    periodicidade: '',
    observacoes: 'Teste',
    criadoEm: instante
  };

  const casos = [];

  function executar(nome, linha, objetoEsperado, resultadoEsperado) {
    const obtido = DCP_CV242_REGISTRO_COMPLETO(
      { linha: 2, dados: linha },
      objetoEsperado
    );

    casos.push({
      teste: nome,
      esperado: resultadoEsperado,
      obtido: obtido,
      aprovado: obtido === resultadoEsperado
    });
  }

  executar(
    'Datas de criacao iguais',
    registroBase.slice(),
    esperado,
    true
  );

  const alteradoCriadoEm = registroBase.slice();
  alteradoCriadoEm[12] = new Date(instante.getTime() + 1000);

  executar(
    'CriadoEm diferente',
    alteradoCriadoEm,
    esperado,
    false
  );

  const alteradoAtualizadoEm = registroBase.slice();
  alteradoAtualizadoEm[13] = new Date(instante.getTime() + 1000);

  executar(
    'AtualizadoEm diferente',
    alteradoAtualizadoEm,
    esperado,
    false
  );

  const ausenteCriadoEm = Object.assign({}, esperado);
  delete ausenteCriadoEm.criadoEm;

  executar(
    'Data esperada ausente',
    registroBase.slice(),
    ausenteCriadoEm,
    false
  );

  const alteradoCampo = registroBase.slice();
  alteradoCampo[10] = 'Observacao divergente';

  executar(
    'Campo funcional divergente',
    alteradoCampo,
    esperado,
    false
  );

  const booleanoInvalido = registroBase.slice();
  booleanoInvalido[7] = 'QUALQUER TEXTO';
  executar(
    'Compartilhamento com valor invalido',
    booleanoInvalido,
    esperado,
    false
  );

  const booleanoNumerico = registroBase.slice();
  booleanoNumerico[7] = 1;
  executar(
    'Compartilhamento com booleano numerico explicito',
    booleanoNumerico,
    esperado,
    true
  );

  const aprovados = casos.filter(function(item) {
    return item.aprovado;
  }).length;

  const resposta = {
    sucesso: aprovados === casos.length,
    total: casos.length,
    aprovados: aprovados,
    reprovados: casos.length - aprovados,
    resultados: casos,
    somenteMemoria: true,
    planilhasAlteradas: false,
    sequenciaVINAlterada: false,
    cadastroRealExecutado: false
  };

  Logger.log(JSON.stringify(resposta, null, 2));
  return resposta;
}