
/**
 * ============================================================
 * DRIVER CONTROL PRO
 * MÓDULO 2.3.6B
 * INATIVAÇÃO SEGURA DE VEÍCULOS
 * ============================================================
 *
 * FUNCIONALIDADES:
 * - Localizar veículo por UUID ou código.
 * - Exigir motivo de inativação.
 * - Impedir inativação de veículo com vínculo ativo.
 * - Alterar somente o campo Status.
 * - Gerar UUID e código HVS para o histórico.
 * - Registrar histórico de status.
 * - Registrar auditoria.
 * - Utilizar um único DocumentLock.
 * - Tentar rollback em caso de falha.
 *
 * IMPORTANTE:
 * Este módulo não executa inativação automaticamente.
 * A alteração só ocorre quando a função de inativação
 * é chamada explicitamente.
 *
 * DEPENDÊNCIA:
 * Utiliza a aba CONFIG com seis colunas:
 * Chave | Valor | TipoValor | Descricao | Editavel | AtualizadoEm
 * ============================================================
 */

const DCP_INATIVACAO_VEICULOS_236 = {
  ABA_VEICULOS: 'VEICULOS',
  ABA_VINCULOS: 'VINCULOS',
  ABA_HISTORICO: 'HISTORICO_STATUS_VEICULOS',
  ABA_AUDITORIA: 'AUDITORIA',
  ABA_CONFIG: 'CONFIG',

  PREFIXO_HISTORICO: 'HVS',
  CHAVE_SEQUENCIA: 'SEQ_HVS',

  USUARIO_PADRAO: 'ADMIN_CENTRAL',

  STATUS_ORIGEM: 'ATIVO',
  STATUS_DESTINO: 'INATIVO',

  TAMANHO_MAXIMO_MOTIVO: 500
};


/**
 * CABEÇALHOS ESPERADOS
 */

const DCP_IB236_CABECALHOS_VEICULOS = [
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
];

const DCP_IB236_CABECALHOS_VINCULOS = [
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
];

const DCP_IB236_CABECALHOS_HISTORICO = [
  'HistoricoStatusID',
  'CodigoHistoricoStatus',
  'VeiculoID',
  'CodigoVeiculo',
  'StatusAnterior',
  'StatusNovo',
  'Motivo',
  'DataHora',
  'UsuarioID',
  'Observacoes',
  'Resultado',
  'AuditoriaID'
];

const DCP_IB236_CABECALHOS_AUDITORIA = [
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
];

const DCP_IB236_CABECALHOS_CONFIG = [
  'Chave',
  'Valor',
  'TipoValor',
  'Descricao',
  'Editavel',
  'AtualizadoEm'
];


/**
 * ============================================================
 * 1. CONFERÊNCIA DE ESTRUTURA
 * ============================================================
 */

function DCP_IB236_CONFERIR_CABECALHOS(nomeAba, esperados) {

  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const aba = planilha.getSheetByName(nomeAba);

  if (!aba) {
    throw new Error(
      'A aba "' + nomeAba + '" não foi encontrada.'
    );
  }

  if (aba.getLastColumn() < esperados.length) {
    throw new Error(
      'A aba "' + nomeAba +
      '" possui menos colunas do que o esperado.'
    );
  }

  const encontrados = aba
    .getRange(1, 1, 1, esperados.length)
    .getDisplayValues()[0]
    .map(function(valor) {
      return String(valor || '').trim();
    });

  const divergencias = [];

  esperados.forEach(function(esperado, indice) {

    if (encontrados[indice] !== esperado) {

      divergencias.push(
        'Coluna ' + (indice + 1) +
        ': esperado "' + esperado +
        '", encontrado "' + encontrados[indice] + '"'
      );

    }

  });

  if (divergencias.length > 0) {

    throw new Error(
      'Estrutura incorreta na aba "' + nomeAba + '":\n' +
      divergencias.join('\n')
    );

  }

  return aba;
}


/**
 * ============================================================
 * 2. GERAÇÃO DO CÓDIGO HVS
 * ============================================================
 *
 * Esta função deve ser chamada somente quando o DocumentLock
 * já estiver adquirido.
 *
 * Não utiliza DCP_GERAR_CODIGO para evitar lock aninhado.
 */

function DCP_IB236_GERAR_CODIGO_HISTORICO(abaConfig) {

  const ultimaLinha = abaConfig.getLastRow();

  const dados = ultimaLinha >= 2
    ? abaConfig.getRange(
        2,
        1,
        ultimaLinha - 1,
        6
      ).getValues()
    : [];

  const chave =
    DCP_INATIVACAO_VEICULOS_236.CHAVE_SEQUENCIA;

  let linhaEncontrada = -1;
  let numeroAtual = 0;

  dados.forEach(function(linha, indice) {

    if (String(linha[0] || '').trim() === chave) {

      if (linhaEncontrada !== -1) {
        throw new Error(
          'Chave de sequência duplicada na CONFIG: ' + chave
        );
      }

      linhaEncontrada = indice + 2;
      numeroAtual = Number(linha[1]) || 0;

    }

  });

  const proximo = numeroAtual + 1;

  if (!Number.isSafeInteger(proximo) || proximo <= 0) {
    throw new Error('Número de sequência inválido.');
  }

  if (linhaEncontrada === -1) {

    abaConfig.appendRow([
      chave,
      proximo,
      'NUMERO',
      'Sequência de históricos de status de veículos',
      'NAO',
      new Date()
    ]);

  } else {

    abaConfig
      .getRange(linhaEncontrada, 2)
      .setValue(proximo);

    abaConfig
      .getRange(linhaEncontrada, 6)
      .setValue(new Date());

  }

  SpreadsheetApp.flush();

  return (
    DCP_INATIVACAO_VEICULOS_236.PREFIXO_HISTORICO +
    '-' +
    String(proximo).padStart(4, '0')
  );

}


/**
 * ============================================================
 * 3. LOCALIZAÇÃO DO VEÍCULO
 * ============================================================
 *
 * Aceita UUID ou código do veículo.
 */

function DCP_IB236_LOCALIZAR_VEICULO(
  abaVeiculos,
  identificador
) {

  const ultimaLinha = abaVeiculos.getLastRow();

  if (ultimaLinha < 2) {
    throw new Error('Não existem veículos cadastrados.');
  }

  const dados = abaVeiculos
    .getRange(
      2,
      1,
      ultimaLinha - 1,
      24
    )
    .getValues();

  const chave = String(identificador || '')
    .trim()
    .toUpperCase();

  const encontrados = [];

  dados.forEach(function(linha, indice) {

    const id = String(linha[0] || '')
      .trim()
      .toUpperCase();

    const codigo = String(linha[1] || '')
      .trim()
      .toUpperCase();

    if (id === chave || codigo === chave) {

      encontrados.push({
        linhaPlanilha: indice + 2,
        dados: linha
      });

    }

  });

  if (encontrados.length === 0) {
    throw new Error('Veículo não encontrado.');
  }

  if (encontrados.length > 1) {
    throw new Error(
      'Identificador corresponde a mais de um veículo.'
    );
  }

  return encontrados[0];

}


/**
 * ============================================================
 * 4. VERIFICAÇÃO DE VÍNCULOS ATIVOS
 * ============================================================
 *
 * Não permite inativação enquanto existir vínculo ATIVO
 * associado ao veículo.
 */

function DCP_IB236_BUSCAR_VINCULOS_ATIVOS(
  abaVinculos,
  veiculoID
) {

  const ultimaLinha = abaVinculos.getLastRow();

  if (ultimaLinha < 2) {
    return [];
  }

  const dados = abaVinculos
    .getRange(
      2,
      1,
      ultimaLinha - 1,
      16
    )
    .getValues();

  const identificador = String(veiculoID || '').trim();

  return dados
    .filter(function(linha) {

      const idVeiculo = String(linha[3] || '').trim();

      const status = String(linha[11] || '')
        .trim()
        .toUpperCase();

      return (
        idVeiculo === identificador &&
        status === 'ATIVO'
      );

    })
    .map(function(linha) {

      return {
        VinculoID: linha[0],
        CodigoVinculo: linha[1],
        MotoristaID: linha[2],
        VeiculoID: linha[3],
        TipoVinculo: linha[4],
        InicioEm: linha[5],
        FimEm: linha[6],
        Status: linha[11]
      };

    });

}


/**
 * ============================================================
 * 5. CONFERÊNCIA DA GRAVAÇÃO
 * ============================================================
 */

function DCP_IB236_CONFIRMAR_REGISTRO(
  aba,
  linha,
  coluna,
  valorEsperado
) {

  const valorLido = String(
    aba.getRange(linha, coluna).getDisplayValue()
  ).trim();

  return valorLido === String(valorEsperado).trim();

}


/**
 * ============================================================
 * 6. OPERAÇÃO PRINCIPAL DE INATIVAÇÃO
 * ============================================================
 *
 * EXEMPLO:
 *
 * DCP_INATIVAR_VEICULO_236({
 *   identificador: 'VEI-0001',
 *   motivo: 'Veículo temporariamente fora de operação'
 * });
 *
 * ATENÇÃO:
 * Esta função altera efetivamente o status do veículo.
 */

function DCP_INATIVAR_VEICULO_236(dados) {

  const bloqueio = LockService.getDocumentLock();

  if (!bloqueio.tryLock(30000)) {

    throw new Error(
      'Não foi possível obter o bloqueio. Tente novamente.'
    );

  }

  let abaVeiculos = null;
  let abaVinculos = null;
  let abaHistorico = null;
  let abaAuditoria = null;
  let abaConfig = null;

  let linhaVeiculo = null;
  let linhaHistorico = null;
  let linhaAuditoria = null;

  let veiculoID = null;
  let codigoVeiculo = null;

  let historicoID = null;
  let codigoHistorico = null;
  let auditoriaID = null;

  let statusAlterado = false;
  let historicoGravado = false;
  let auditoriaGravada = false;

  let statusFinal = null;

  try {

    // --------------------------------------------------------
    // 6.1 VALIDAÇÃO DOS DADOS RECEBIDOS
    // --------------------------------------------------------

    if (!dados || typeof dados !== 'object') {
      throw new Error(
        'Dados da operação não foram informados.'
      );
    }

    const identificador = String(
      dados.identificador || ''
    ).trim();

    const motivo = String(
      dados.motivo || ''
    ).trim();

    if (!identificador) {
      throw new Error(
        'Informe o código ou UUID do veículo.'
      );
    }

    if (!motivo) {
      throw new Error(
        'O motivo da inativação é obrigatório.'
      );
    }

    if (
      motivo.length >
      DCP_INATIVACAO_VEICULOS_236.TAMANHO_MAXIMO_MOTIVO
    ) {

      throw new Error(
        'O motivo deve ter no máximo 500 caracteres.'
      );

    }

    // --------------------------------------------------------
    // 6.2 CONFERÊNCIA DAS ESTRUTURAS
    // --------------------------------------------------------

    abaVeiculos = DCP_IB236_CONFERIR_CABECALHOS(
      DCP_INATIVACAO_VEICULOS_236.ABA_VEICULOS,
      DCP_IB236_CABECALHOS_VEICULOS
    );

    abaVinculos = DCP_IB236_CONFERIR_CABECALHOS(
      DCP_INATIVACAO_VEICULOS_236.ABA_VINCULOS,
      DCP_IB236_CABECALHOS_VINCULOS
    );

    abaHistorico = DCP_IB236_CONFERIR_CABECALHOS(
      DCP_INATIVACAO_VEICULOS_236.ABA_HISTORICO,
      DCP_IB236_CABECALHOS_HISTORICO
    );

    abaAuditoria = DCP_IB236_CONFERIR_CABECALHOS(
      DCP_INATIVACAO_VEICULOS_236.ABA_AUDITORIA,
      DCP_IB236_CABECALHOS_AUDITORIA
    );

    abaConfig = DCP_IB236_CONFERIR_CABECALHOS(
      DCP_INATIVACAO_VEICULOS_236.ABA_CONFIG,
      DCP_IB236_CABECALHOS_CONFIG
    );

    // --------------------------------------------------------
    // 6.3 LOCALIZAÇÃO DO VEÍCULO
    // --------------------------------------------------------

    const encontrado = DCP_IB236_LOCALIZAR_VEICULO(
      abaVeiculos,
      identificador
    );

    linhaVeiculo = encontrado.linhaPlanilha;

    const veiculo = encontrado.dados;

    veiculoID = String(veiculo[0] || '').trim();

    codigoVeiculo = String(veiculo[1] || '').trim();

    const statusAtual = String(veiculo[19] || '')
      .trim()
      .toUpperCase();

    if (!veiculoID || !codigoVeiculo) {
      throw new Error(
        'O veículo possui identificador ou código vazio.'
      );
    }

    if (
      statusAtual !==
      DCP_INATIVACAO_VEICULOS_236.STATUS_ORIGEM
    ) {

      throw new Error(
        'O veículo precisa estar ATIVO para ser inativado. ' +
        'Status atual: ' + statusAtual
      );

    }

    // --------------------------------------------------------
    // 6.4 VERIFICAÇÃO DOS VÍNCULOS
    // --------------------------------------------------------

    const vinculos = DCP_IB236_BUSCAR_VINCULOS_ATIVOS(
      abaVinculos,
      veiculoID
    );

    if (vinculos.length > 0) {

      throw new Error(
        'Inativação bloqueada: o veículo possui ' +
        vinculos.length +
        ' vínculo(s) ativo(s). Encerre-os antes de continuar.'
      );

    }

    // --------------------------------------------------------
    // 6.5 GERAÇÃO DOS IDENTIFICADORES
    // --------------------------------------------------------

    const agora = new Date();

    historicoID = Utilities.getUuid();

    auditoriaID = Utilities.getUuid();

    codigoHistorico =
      DCP_IB236_GERAR_CODIGO_HISTORICO(abaConfig);

    const usuarioID =
      DCP_INATIVACAO_VEICULOS_236.USUARIO_PADRAO;

    // --------------------------------------------------------
    // 6.6 ALTERAÇÃO DO STATUS
    // --------------------------------------------------------

    abaVeiculos
      .getRange(linhaVeiculo, 20)
      .setValue(
        DCP_INATIVACAO_VEICULOS_236.STATUS_DESTINO
      );

    SpreadsheetApp.flush();

    statusFinal = String(
      abaVeiculos
        .getRange(linhaVeiculo, 20)
        .getDisplayValue()
    ).trim().toUpperCase();

    if (statusFinal !== 'INATIVO') {

      throw new Error(
        'Não foi possível confirmar a alteração do status.'
      );

    }

    statusAlterado = true;

    // --------------------------------------------------------
    // 6.7 GRAVAÇÃO DO HISTÓRICO
    // --------------------------------------------------------

    const ultimaLinhaHistoricoAntes =
      abaHistorico.getLastRow();

    abaHistorico.appendRow([
      historicoID,
      codigoHistorico,
      veiculoID,
      codigoVeiculo,
      'ATIVO',
      'INATIVO',
      motivo,
      agora,
      usuarioID,
      '',
      'SUCESSO',
      auditoriaID
    ]);

    linhaHistorico =
      ultimaLinhaHistoricoAntes + 1;

    SpreadsheetApp.flush();

    historicoGravado =
      DCP_IB236_CONFIRMAR_REGISTRO(
        abaHistorico,
        linhaHistorico,
        1,
        historicoID
      );

    if (!historicoGravado) {

      throw new Error(
        'Não foi possível confirmar a gravação do histórico.'
      );

    }

    // --------------------------------------------------------
    // 6.8 GRAVAÇÃO DA AUDITORIA
    // --------------------------------------------------------

    const ultimaLinhaAuditoriaAntes =
      abaAuditoria.getLastRow();

    abaAuditoria.appendRow([
      auditoriaID,
      agora,
      usuarioID,
      'INATIVACAO',
      'VEICULOS',
      veiculoID,
      'ATIVO',
      'INATIVO',
      'SUCESSO',
      'SCRIPT',
      'Histórico ' + codigoHistorico +
      '. Motivo: ' + motivo
    ]);

    linhaAuditoria =
      ultimaLinhaAuditoriaAntes + 1;

    SpreadsheetApp.flush();

    auditoriaGravada =
      DCP_IB236_CONFIRMAR_REGISTRO(
        abaAuditoria,
        linhaAuditoria,
        1,
        auditoriaID
      );

    if (!auditoriaGravada) {

      throw new Error(
        'Não foi possível confirmar a gravação da auditoria.'
      );

    }

    // --------------------------------------------------------
    // 6.9 RESULTADO FINAL
    // --------------------------------------------------------

    return {

      sucesso: true,

      codigoVeiculo: codigoVeiculo,

      veiculoID: veiculoID,

      historicoID: historicoID,

      codigoHistorico: codigoHistorico,

      auditoriaID: auditoriaID,

      statusAnterior: 'ATIVO',

      statusNovo: 'INATIVO',

      motivo: motivo,

      historicoConfirmado: true,

      auditoriaConfirmada: true,

      mensagem:
        'Veículo ' + codigoVeiculo +
        ' inativado. Histórico e auditoria confirmados.'

    };

  } catch (erro) {

    // --------------------------------------------------------
    // 6.10 TRATAMENTO DE FALHAS E ROLLBACK
    // --------------------------------------------------------

    let rollbackTentado = false;
    let rollbackConfirmado = false;
    let rollbackErro = null;

    // Tenta restaurar o status original.
    if (
      statusAlterado &&
      abaVeiculos &&
      linhaVeiculo
    ) {

      rollbackTentado = true;

      try {

        abaVeiculos
          .getRange(linhaVeiculo, 20)
          .setValue('ATIVO');

        SpreadsheetApp.flush();

        statusFinal = String(
          abaVeiculos
            .getRange(linhaVeiculo, 20)
            .getDisplayValue()
        ).trim().toUpperCase();

        rollbackConfirmado =
          statusFinal === 'ATIVO';

        if (!rollbackConfirmado) {

          rollbackErro =
            'A leitura posterior não confirmou o status ATIVO.';

        }

      } catch (erroRollback) {

        rollbackErro = erroRollback.message;

      }

    }

    // Marca o histórico como FALHA, se existir.
    if (
      linhaHistorico &&
      abaHistorico
    ) {

      try {

        const idNaLinha = String(
          abaHistorico
            .getRange(linhaHistorico, 1)
            .getDisplayValue()
        ).trim();

        if (idNaLinha === String(historicoID)) {

          abaHistorico
            .getRange(linhaHistorico, 11)
            .setValue('FALHA');

          abaHistorico
            .getRange(linhaHistorico, 10)
            .setValue(
              'Operação não concluída. Erro: ' +
              erro.message +
              '. Rollback confirmado: ' +
              (rollbackConfirmado ? 'SIM' : 'NÃO')
            );

          SpreadsheetApp.flush();

          historicoGravado = true;

        }

      } catch (erroHistorico) {

        // A falha será informada no retorno.

      }

    }

    // Marca a auditoria como FALHA, se ela chegou a ser gravada.
    if (
      linhaAuditoria &&
      abaAuditoria
    ) {

      try {

        const idNaLinha = String(
          abaAuditoria
            .getRange(linhaAuditoria, 1)
            .getDisplayValue()
        ).trim();

        if (idNaLinha === String(auditoriaID)) {

          abaAuditoria
            .getRange(linhaAuditoria, 9)
            .setValue('FALHA');

          abaAuditoria
            .getRange(linhaAuditoria, 11)
            .setValue(
              'Operação não concluída. Erro: ' +
              erro.message +
              '. Rollback confirmado: ' +
              (rollbackConfirmado ? 'SIM' : 'NÃO')
            );

          SpreadsheetApp.flush();

          auditoriaGravada = true;

        }

      } catch (erroAuditoria) {

        // A falha será informada no retorno.

      }

    }

    // Confere o estado final real do veículo.
    if (
      abaVeiculos &&
      linhaVeiculo
    ) {

      try {

        statusFinal = String(
          abaVeiculos
            .getRange(linhaVeiculo, 20)
            .getDisplayValue()
        ).trim().toUpperCase();

      } catch (erroLeitura) {

        statusFinal = 'NAO_CONFIRMADO';

      }

    }

    const estadoOperacao =
      statusFinal === 'ATIVO'
        ? 'REVERTIDA'
        : statusFinal === 'INATIVO'
          ? 'INATIVACAO_PERSISTIU'
          : 'ESTADO_NAO_CONFIRMADO';

    return {

      sucesso: false,

      mensagem: erro.message,

      estadoOperacao: estadoOperacao,

      statusFinal: statusFinal,

      rollbackTentado: rollbackTentado,

      rollbackConfirmado: rollbackConfirmado,

      rollbackErro: rollbackErro,

      historicoGravado: historicoGravado,

      auditoriaGravada: auditoriaGravada,

      historicoID: historicoID,

      codigoHistorico: codigoHistorico,

      auditoriaID: auditoriaID,

      observacao:
        estadoOperacao === 'REVERTIDA'
          ? 'O veículo permaneceu ATIVO. Confira histórico e auditoria.'
          : estadoOperacao === 'INATIVACAO_PERSISTIU'
            ? 'ATENÇÃO: o veículo continua INATIVO apesar da falha. Verifique histórico e auditoria antes de qualquer nova operação.'
            : 'ATENÇÃO: estado final não confirmado. Faça conferência manual antes de tentar novamente.'

    };

  } finally {

    bloqueio.releaseLock();

  }

}


/**
 * ============================================================
 * 7. DIAGNÓSTICO ESTRUTURAL
 * ============================================================
 *
 * Não altera dados.
 */

function DCP_TESTAR_ESTRUTURA_INATIVACAO_236() {

  const planilha =
    SpreadsheetApp.getActiveSpreadsheet();

  const nomes = [
    'VEICULOS',
    'VINCULOS',
    'HISTORICO_STATUS_VEICULOS',
    'AUDITORIA',
    'CONFIG'
  ];

  const resultado = nomes.map(function(nome) {

    const aba = planilha.getSheetByName(nome);

    return {

      aba: nome,

      existe: !!aba,

      ultimaLinha: aba
        ? aba.getLastRow()
        : null

    };

  });

  Logger.log(
    JSON.stringify(resultado, null, 2)
  );

  return resultado;

}


/**
 * ============================================================
 * 8. DIAGNÓSTICO INDIVIDUAL DE INATIVAÇÃO
 * ============================================================
 *
 * Não altera dados.
 *
 * Exemplo:
 * DCP_DIAGNOSTICAR_INATIVACAO_VEI0001_236()
 */

function DCP_DIAGNOSTICAR_INATIVACAO_VEI0001_236() {

  const bloqueio = LockService.getDocumentLock();

  if (!bloqueio.tryLock(30000)) {

    throw new Error(
      'Não foi possível obter o bloqueio.'
    );

  }

  try {

    const abaVeiculos = DCP_IB236_CONFERIR_CABECALHOS(
      'VEICULOS',
      DCP_IB236_CABECALHOS_VEICULOS
    );

    const abaVinculos = DCP_IB236_CONFERIR_CABECALHOS(
      'VINCULOS',
      DCP_IB236_CABECALHOS_VINCULOS
    );

    const encontrado = DCP_IB236_LOCALIZAR_VEICULO(
      abaVeiculos,
      'VEI-0001'
    );

    const veiculo = encontrado.dados;

    const veiculoID = String(veiculo[0] || '').trim();

    const status = String(veiculo[19] || '')
      .trim()
      .toUpperCase();

    const vinculos = DCP_IB236_BUSCAR_VINCULOS_ATIVOS(
      abaVinculos,
      veiculoID
    );

    const resultado = {

      sucesso: true,

      codigoVeiculo: veiculo[1],

      placa: veiculo[2],

      marca: veiculo[3],

      modelo: veiculo[4],

      statusAtual: status,

      quantidadeVinculosAtivos: vinculos.length,

      podeSerInativado:
        status === 'ATIVO' &&
        vinculos.length === 0,

      vinculosAtivos: vinculos,

      mensagem:
        status !== 'ATIVO'
          ? 'O veículo não está ATIVO.'
          : vinculos.length > 0
            ? 'Existem vínculos ativos. Inativação bloqueada.'
            : 'Veículo apto para inativação.'

    };

    Logger.log(
      JSON.stringify(resultado, null, 2)
    );

    return resultado;

  } finally {

    bloqueio.releaseLock();

  }

}


/**
 * ============================================================
 * 9. EXECUÇÃO CONTROLADA DE INATIVAÇÃO
 * ============================================================
 *
 * ATENÇÃO:
 * Esta função altera efetivamente o VEI-0001.
 *
 * Execute somente quando autorizar a inativação.
 */

function DCP_TESTAR_INATIVACAO_REAL_VEI0001_236() {

  const resultado = DCP_INATIVAR_VEICULO_236({

    identificador: 'VEI-0001',

    motivo:
      'Inativação controlada para teste do módulo 2.3.6'

  });

  Logger.log(
    JSON.stringify(resultado, null, 2)
  );

  return resultado;

}


/**
 * ============================================================
 * FIM DO MÓDULO 2.3.6B
 * ============================================================
 */
