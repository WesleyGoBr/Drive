
/************************************************************
 * DRIVER CONTROL PRO
 * MODULO 2.3 - COMBUSTIVEIS
 * Versao: 2.3.2
 *
 * Funcoes:
 * 1. Atualizar estrutura da aba COMBUSTIVEIS
 * 2. Diagnosticar estrutura
 * 3. Validar dados
 * 4. Cadastrar combustiveis
 * 5. Gerar UUID e codigo sequencial
 * 6. Registrar auditoria
 *
 * IMPORTANTE:
 * Este modulo utiliza DocumentLock.
 * As operacoes de sequencia, cadastro e auditoria
 * sao controladas por um unico bloqueio.
 ************************************************************/


/************************************************************
 * 1. CONSTANTES DO MODULO
 ************************************************************/

const DCP_COMBUSTIVEIS_232 = {
  ABA: 'COMBUSTIVEIS',
  CONFIG: 'CONFIG',
  AUDITORIA: 'AUDITORIA',

  PREFIXO: 'COM',
  CHAVE_SEQUENCIA: 'SEQ_COM',

  ADMIN: 'ADMIN_CENTRAL',

  CABECALHOS: [
    'CombustivelID',
    'CodigoCombustivel',
    'NomeCombustivel',
    'TipoEnergia',
    'UnidadeMedida',
    'Observacoes',
    'Status',
    'CriadoEm',
    'AtualizadoEm',
    'CriadoPor',
    'AtualizadoPor',
    'CodigoReferencia'
  ],

  CABECALHOS_CONFIG: [
    'Chave',
    'Valor',
    'TipoValor',
    'Descricao',
    'Editavel',
    'AtualizadoEm'
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

  TIPOS_ENERGIA: [
    'FOSSIL',
    'BIOCOMBUSTIVEL',
    'OUTRO'
  ],

  UNIDADES: [
    'L',
    'M3',
    'KG'
  ]
};


/************************************************************
 * 2. ATUALIZAR ESTRUTURA DA ABA
 *
 * Acrescenta CodigoReferencia na coluna L.
 * Nao apaga dados existentes.
 ************************************************************/

function DCP_ATUALIZAR_ESTRUTURA_COMBUSTIVEIS_232() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const aba = ss.getSheetByName(
    DCP_COMBUSTIVEIS_232.ABA
  );

  if (!aba) {
    throw new Error(
      'Aba COMBUSTIVEIS nao encontrada.'
    );
  }

  const cabecalhosAntigos =
    DCP_COMBUSTIVEIS_232.CABECALHOS.slice(0, 11);

  const encontrados = aba
    .getRange(1, 1, 1, 11)
    .getValues()[0];

  const estruturaCorreta = cabecalhosAntigos.every(
    (campo, indice) => campo === encontrados[indice]
  );

  if (!estruturaCorreta) {
    throw new Error(
      'Cabecalhos existentes diferentes do esperado. ' +
      'Atualizacao cancelada.'
    );
  }

  const colunaL = aba.getRange(1, 12).getValue();

  if (colunaL === 'CodigoReferencia') {
    Logger.log(
      'OK: CodigoReferencia ja existe.'
    );
    return;
  }

  if (colunaL !== '') {
    throw new Error(
      'A coluna L ja possui outro conteudo. ' +
      'Atualizacao cancelada.'
    );
  }

  aba.getRange(1, 12)
    .setValue('CodigoReferencia');

  aba.getRange(1, 12)
    .setFontWeight('bold');

  aba.autoResizeColumn(12);

  SpreadsheetApp.flush();

  Logger.log(
    'OK: CodigoReferencia adicionada.'
  );

  Logger.log(
    'Total de colunas: 12'
  );
}


/************************************************************
 * 3. DIAGNOSTICO DA ESTRUTURA
 *
 * Funcao somente leitura.
 ************************************************************/

function DCP_DIAGNOSTICAR_COMBUSTIVEIS_232() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const estruturas = [
    {
      nome: 'COMBUSTIVEIS',
      cabecalhos: DCP_COMBUSTIVEIS_232.CABECALHOS
    },
    {
      nome: 'CONFIG',
      cabecalhos: DCP_COMBUSTIVEIS_232.CABECALHOS_CONFIG
    },
    {
      nome: 'AUDITORIA',
      cabecalhos: DCP_COMBUSTIVEIS_232.CABECALHOS_AUDITORIA
    }
  ];

  let problemas = 0;

  Logger.log(
    '===== DIAGNOSTICO COMBUSTIVEIS 2.3.2 ====='
  );

  estruturas.forEach(item => {

    const aba = ss.getSheetByName(item.nome);

    if (!aba) {
      Logger.log(
        'ERRO: Aba inexistente: ' + item.nome
      );

      problemas++;
      return;
    }

    const encontrados = aba
      .getRange(
        1,
        1,
        1,
        item.cabecalhos.length
      )
      .getValues()[0];

    const correto = item.cabecalhos.every(
      (campo, indice) =>
        campo === encontrados[indice]
    );

    if (correto) {
      Logger.log(
        'OK: Estrutura ' + item.nome
      );
    } else {
      Logger.log(
        'ERRO: Cabecalhos divergentes em ' + item.nome
      );

      Logger.log(
        'Esperados: ' + item.cabecalhos.join(' | ')
      );

      Logger.log(
        'Encontrados: ' + encontrados.join(' | ')
      );

      problemas++;
    }
  });

  Logger.log(
    'TOTAL DE PROBLEMAS: ' + problemas
  );

  Logger.log(
    problemas === 0
      ? 'RESULTADO: ESTRUTURA APROVADA.'
      : 'RESULTADO: NECESSITA CORRECAO.'
  );
}


/************************************************************
 * 4. VALIDAR DADOS DO COMBUSTIVEL
 *
 * Nao grava dados.
 ************************************************************/

function DCP_VALIDAR_COMBUSTIVEL_232(dados) {

  const erros = [];

  if (!dados || typeof dados !== 'object') {
    return {
      valido: false,
      erros: ['Dados nao informados.']
    };
  }

  const nome = String(
    dados.nomeCombustivel || ''
  ).trim();

  const referencia = String(
    dados.codigoReferencia || ''
  ).trim().toUpperCase();

  const tipo = String(
    dados.tipoEnergia || ''
  ).trim().toUpperCase();

  const unidade = String(
    dados.unidadeMedida || ''
  ).trim().toUpperCase();

  const observacoes = String(
    dados.observacoes || ''
  ).trim();

  if (nome.length < 2 || nome.length > 100) {
    erros.push(
      'Nome deve ter entre 2 e 100 caracteres.'
    );
  }

  if (!/^[A-Z0-9_]{2,30}$/.test(referencia)) {
    erros.push(
      'CodigoReferencia invalido. Use letras, numeros e _.'
    );
  }

  if (
    !DCP_COMBUSTIVEIS_232.TIPOS_ENERGIA.includes(tipo)
  ) {
    erros.push(
      'TipoEnergia invalido.'
    );
  }

  if (
    !DCP_COMBUSTIVEIS_232.UNIDADES.includes(unidade)
  ) {
    erros.push(
      'UnidadeMedida invalida. Utilize L, M3 ou KG.'
    );
  }

  if (observacoes.length > 2000) {
    erros.push(
      'Observacoes devem ter no maximo 2000 caracteres.'
    );
  }

  return {
    valido: erros.length === 0,
    erros: erros,
    dadosNormalizados: {
      nomeCombustivel: nome,
      codigoReferencia: referencia,
      tipoEnergia: tipo,
      unidadeMedida: unidade,
      observacoes: observacoes
    }
  };
}


/************************************************************
 * 5. CADASTRAR COMBUSTIVEL
 *
 * Exemplo de chamada:
 *
 * DCP_CADASTRAR_COMBUSTIVEL_232({
 *   nomeCombustivel: 'Gasolina',
 *   codigoReferencia: 'GASOLINA',
 *   tipoEnergia: 'FOSSIL',
 *   unidadeMedida: 'L',
 *   observacoes: ''
 * });
 ************************************************************/

function DCP_CADASTRAR_COMBUSTIVEL_232(dados) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const aba = ss.getSheetByName(
    DCP_COMBUSTIVEIS_232.ABA
  );

  const config = ss.getSheetByName(
    DCP_COMBUSTIVEIS_232.CONFIG
  );

  const auditoria = ss.getSheetByName(
    DCP_COMBUSTIVEIS_232.AUDITORIA
  );

  let lock;
  let bloqueioAdquirido = false;
  let registroGravado = false;
  let tentativaAuditada = false;

  let combustivelID = '';
  let codigoCombustivel = '';

  const agora = new Date();
  const usuario = DCP_COMBUSTIVEIS_232.ADMIN;

  try {

    /******************************************************
     * A. VALIDACAO INICIAL
     ******************************************************/

    const validacao =
      DCP_VALIDAR_COMBUSTIVEL_232(dados);

    if (!validacao.valido) {
      throw new Error(
        validacao.erros.join(' ')
      );
    }

    const normalizados =
      validacao.dadosNormalizados;

    const nome = normalizados.nomeCombustivel;
    const referencia = normalizados.codigoReferencia;
    const tipo = normalizados.tipoEnergia;
    const unidade = normalizados.unidadeMedida;
    const observacoes = normalizados.observacoes;


    /******************************************************
     * B. VALIDAR ABAS
     ******************************************************/

    if (!aba || !config || !auditoria) {
      throw new Error(
        'Uma ou mais abas necessarias nao existem.'
      );
    }

    function conferirCabecalhos(
      planilha,
      esperados
    ) {

      const encontrados = planilha
        .getRange(
          1,
          1,
          1,
          esperados.length
        )
        .getValues()[0];

      return esperados.every(
        (campo, indice) =>
          campo === encontrados[indice]
      );
    }

    if (
      !conferirCabecalhos(
        aba,
        DCP_COMBUSTIVEIS_232.CABECALHOS
      )
    ) {
      throw new Error(
        'Cabecalhos de COMBUSTIVEIS incorretos.'
      );
    }

    if (
      !conferirCabecalhos(
        config,
        DCP_COMBUSTIVEIS_232.CABECALHOS_CONFIG
      )
    ) {
      throw new Error(
        'Cabecalhos de CONFIG incorretos.'
      );
    }

    if (
      !conferirCabecalhos(
        auditoria,
        DCP_COMBUSTIVEIS_232.CABECALHOS_AUDITORIA
      )
    ) {
      throw new Error(
        'Cabecalhos de AUDITORIA incorretos.'
      );
    }


    /******************************************************
     * C. ADQUIRIR BLOQUEIO UNICO
     ******************************************************/

    lock = LockService.getDocumentLock();

    lock.waitLock(30000);

    bloqueioAdquirido = true;


    /******************************************************
     * D. VERIFICAR REGISTROS EXISTENTES
     ******************************************************/

    const ultimaLinha = aba.getLastRow();

    let maiorSequenciaExistente = 0;

    if (ultimaLinha >= 2) {

      const registros = aba
        .getRange(
          2,
          1,
          ultimaLinha - 1,
          12
        )
        .getValues();

      for (const linha of registros) {

        const idExistente =
          String(linha[0] || '').trim();

        const codigoExistente =
          String(linha[1] || '').trim().toUpperCase();

        const nomeExistente =
          String(linha[2] || '').trim().toUpperCase();

        const referenciaExistente =
          String(linha[11] || '').trim().toUpperCase();


        // Ignora linhas completamente vazias.

        if (
          !idExistente &&
          !codigoExistente &&
          !nomeExistente &&
          !referenciaExistente
        ) {
          continue;
        }


        // Impede registros incompletos.

        if (!idExistente) {
          throw new Error(
            'Registro existente sem UUID. Verifique a aba COMBUSTIVEIS.'
          );
        }

        if (!codigoExistente) {
          throw new Error(
            'Registro existente sem CodigoCombustivel.'
          );
        }


        // Verifica formato e recupera sequencia.

        const correspondencia =
          codigoExistente.match(/^COM-(\d{4,})$/);

        if (!correspondencia) {
          throw new Error(
            'Codigo sequencial invalido: ' +
            codigoExistente
          );
        }

        maiorSequenciaExistente = Math.max(
          maiorSequenciaExistente,
          Number(correspondencia[1])
        );


        // Impede duplicidade de referencia.

        if (
          referenciaExistente &&
          referenciaExistente === referencia
        ) {
          throw new Error(
            'Codigo de referencia ja cadastrado.'
          );
        }


        // Impede duplicidade de nome.

        if (
          nomeExistente &&
          nomeExistente === nome.toUpperCase()
        ) {
          throw new Error(
            'Ja existe um combustivel com esse nome.'
          );
        }
      }
    }


    /******************************************************
     * E. LOCALIZAR SEQUENCIA NA CONFIG
     ******************************************************/

    const chave =
      DCP_COMBUSTIVEIS_232.CHAVE_SEQUENCIA;

    const ultimaLinhaConfig = config.getLastRow();

    let linhaSequencia = 0;
    let maiorSequenciaConfig = 0;

    if (ultimaLinhaConfig >= 2) {

      const registrosConfig = config
        .getRange(
          2,
          1,
          ultimaLinhaConfig - 1,
          6
        )
        .getValues();

      registrosConfig.forEach(
        (linha, indice) => {

          if (
            String(linha[0] || '').trim() === chave
          ) {

            if (linhaSequencia !== 0) {
              throw new Error(
                'Chave SEQ_COM duplicada na CONFIG.'
              );
            }

            linhaSequencia = indice + 2;

            maiorSequenciaConfig =
              Number(linha[1]);

            if (
              !Number.isInteger(maiorSequenciaConfig) ||
              maiorSequenciaConfig < 0
            ) {
              throw new Error(
                'Valor de SEQ_COM invalido na CONFIG.'
              );
            }
          }
        }
      );
    }


    /******************************************************
     * F. GERAR IDENTIFICADORES
     ******************************************************/

    const novoNumero =
      Math.max(
        maiorSequenciaConfig,
        maiorSequenciaExistente
      ) + 1;

    codigoCombustivel =
      'COM-' +
      String(novoNumero).padStart(4, '0');

    combustivelID = Utilities.getUuid();


    /******************************************************
     * G. MONTAR NOVO REGISTRO
     ******************************************************/

    const novoRegistro = {
      CombustivelID: combustivelID,
      CodigoCombustivel: codigoCombustivel,
      NomeCombustivel: nome,
      TipoEnergia: tipo,
      UnidadeMedida: unidade,
      Observacoes: observacoes,
      Status: 'ATIVO',
      CodigoReferencia: referencia
    };


    /******************************************************
     * H. RESERVAR SEQUENCIA
     ******************************************************/

    if (linhaSequencia > 0) {

      config
        .getRange(linhaSequencia, 2)
        .setValue(novoNumero);

      config
        .getRange(linhaSequencia, 6)
        .setValue(agora);

    } else {

      config.appendRow([
        chave,
        novoNumero,
        'NUMERO',
        'Sequencia de combustiveis',
        false,
        agora
      ]);
    }

    SpreadsheetApp.flush();


    /******************************************************
     * I. REGISTRAR TENTATIVA DE CRIACAO
     ******************************************************/

    auditoria.appendRow([
      Utilities.getUuid(),
      new Date(),
      usuario,
      'TENTATIVA_CRIACAO',
      'COMBUSTIVEIS',
      combustivelID,
      '{}',
      JSON.stringify(novoRegistro),
      'PENDENTE',
      'APPS_SCRIPT',
      'Tentativa de cadastro de combustivel.'
    ]);

    tentativaAuditada = true;

    SpreadsheetApp.flush();


    /******************************************************
     * J. GRAVAR COMBUSTIVEL
     ******************************************************/

    aba.appendRow([
      combustivelID,
      codigoCombustivel,
      nome,
      tipo,
      unidade,
      observacoes,
      'ATIVO',
      agora,
      agora,
      usuario,
      usuario,
      referencia
    ]);

    registroGravado = true;

    SpreadsheetApp.flush();


    /******************************************************
     * K. CONFIRMAR AUDITORIA
     ******************************************************/

    try {

      auditoria.appendRow([
        Utilities.getUuid(),
        new Date(),
        usuario,
        'CONFIRMACAO_CRIACAO',
        'COMBUSTIVEIS',
        combustivelID,
        '{}',
        JSON.stringify(novoRegistro),
        'SUCESSO',
        'APPS_SCRIPT',
        'Combustivel cadastrado com sucesso.'
      ]);

      SpreadsheetApp.flush();

    } catch (erroAuditoria) {

      Logger.log(
        'FALHA NA CONFIRMACAO DA AUDITORIA: ' +
        erroAuditoria.message
      );

      return {
        sucesso: true,
        auditoriaPendente: true,
        id: combustivelID,
        codigo: codigoCombustivel,
        mensagem:
          'Combustivel cadastrado, mas a confirmacao ' +
          'da auditoria falhou. Confira a aba AUDITORIA.'
      };
    }


    /******************************************************
     * L. RETORNO DE SUCESSO
     ******************************************************/

    return {
      sucesso: true,
      auditoriaPendente: false,
      id: combustivelID,
      codigo: codigoCombustivel,
      mensagem:
        'Combustivel cadastrado com sucesso.'
    };


  } catch (erro) {

    Logger.log(
      'ERRO NO CADASTRO DE COMBUSTIVEL: ' +
      erro.message
    );


    /******************************************************
     * M. REGISTRAR FALHA, SE A TENTATIVA FOI AUDITADA
     ******************************************************/

    if (
      bloqueioAdquirido &&
      tentativaAuditada &&
      !registroGravado
    ) {

      try {

        auditoria.appendRow([
          Utilities.getUuid(),
          new Date(),
          usuario,
          'FALHA_CRIACAO',
          'COMBUSTIVEIS',
          combustivelID,
          '{}',
          JSON.stringify({
            CodigoCombustivel: codigoCombustivel
          }),
          'FALHA',
          'APPS_SCRIPT',
          String(erro.message || erro)
        ]);

        SpreadsheetApp.flush();

      } catch (erroSecundario) {

        Logger.log(
          'FALHA AO REGISTRAR ERRO NA AUDITORIA: ' +
          erroSecundario.message
        );
      }
    }


    return {
      sucesso: false,
      auditoriaPendente: false,
      mensagem: String(
        erro.message || erro
      )
    };


  } finally {

    /******************************************************
     * N. LIBERAR BLOQUEIO
     ******************************************************/

    if (bloqueioAdquirido) {
      lock.releaseLock();
    }
  }
}


/************************************************************
 * 6. TESTAR VALIDACOES
 *
 * Teste somente de validacao.
 * Nao cadastra nenhum combustivel.
 ************************************************************/

function DCP_TESTAR_VALIDACOES_COMBUSTIVEIS_232() {

  const testes = [
    {
      nome: 'Gasolina valida',

      dados: {
        nomeCombustivel: 'Gasolina',
        codigoReferencia: 'GASOLINA',
        tipoEnergia: 'FOSSIL',
        unidadeMedida: 'L'
      },

      esperado: true
    },

    {
      nome: 'Etanol valido',

      dados: {
        nomeCombustivel: 'Etanol',
        codigoReferencia: 'ETANOL',
        tipoEnergia: 'BIOCOMBUSTIVEL',
        unidadeMedida: 'L'
      },

      esperado: true
    },

    {
      nome: 'Nome vazio',

      dados: {
        nomeCombustivel: '',
        codigoReferencia: 'TESTE',
        tipoEnergia: 'FOSSIL',
        unidadeMedida: 'L'
      },

      esperado: false
    },

    {
      nome: 'Referencia invalida',

      dados: {
        nomeCombustivel: 'Combustivel teste',
        codigoReferencia: 'GASOLINA-TESTE',
        tipoEnergia: 'FOSSIL',
        unidadeMedida: 'L'
      },

      esperado: false
    },

    {
      nome: 'Tipo de energia invalido',

      dados: {
        nomeCombustivel: 'Combustivel teste',
        codigoReferencia: 'TESTE',
        tipoEnergia: 'INVALIDO',
        unidadeMedida: 'L'
      },

      esperado: false
    },

    {
      nome: 'Unidade invalida',

      dados: {
        nomeCombustivel: 'Combustivel teste',
        codigoReferencia: 'TESTE',
        tipoEnergia: 'FOSSIL',
        unidadeMedida: 'KM'
      },

      esperado: false
    }
  ];

  let aprovados = 0;
  let falhas = 0;

  Logger.log(
    '===== TESTE VALIDACOES COMBUSTIVEIS ====='
  );

  testes.forEach(teste => {

    const resultado =
      DCP_VALIDAR_COMBUSTIVEL_232(teste.dados);

    if (resultado.valido === teste.esperado) {

      Logger.log(
        'OK: ' + teste.nome
      );

      aprovados++;

    } else {

      Logger.log(
        'FALHA: ' + teste.nome
      );

      falhas++;
    }
  });

  Logger.log(
    'APROVADOS: ' + aprovados
  );

  Logger.log(
    'FALHAS: ' + falhas
  );

  Logger.log(
    'TOTAL: ' + testes.length
  );

  Logger.log(
    falhas === 0
      ? 'RESULTADO: TESTES APROVADOS.'
      : 'RESULTADO: EXISTEM FALHAS.'
  );
}

function DCP_TESTAR_CADASTRO_GASOLINA_233() {

  const resultado = DCP_CADASTRAR_COMBUSTIVEL_232({
    nomeCombustivel: 'Gasolina',
    codigoReferencia: 'GASOLINA',
    tipoEnergia: 'FOSSIL',
    unidadeMedida: 'L',
    observacoes: 'Combustivel para motores de combustao.'
  });

  Logger.log('===== RESULTADO CADASTRO GASOLINA =====');
  Logger.log(JSON.stringify(resultado, null, 2));

  if (!resultado.sucesso) {
    throw new Error(resultado.mensagem);
  }
}
