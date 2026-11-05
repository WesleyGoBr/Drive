
/**
 * DRIVER CONTROL PRO
 * MODULO 2.3.3 - GESTAO DO CATALOGO DE COMBUSTIVEIS
 *
 * Funcoes:
 * - Listar combustiveis ativos
 * - Listar todos os combustiveis
 * - Inativar combustivel
 *
 * Nao exclui registros fisicamente.
 */

const DCP_GESTAO_COMBUSTIVEIS_233 = {
  ABA: 'COMBUSTIVEIS',
  AUDITORIA: 'AUDITORIA',
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
  ]
};


/**
 * Confere se a estrutura da aba esta correta.
 */
function DCP_CONFERIR_ESTRUTURA_GESTAO_COMBUSTIVEIS_233() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const aba = ss.getSheetByName(
    DCP_GESTAO_COMBUSTIVEIS_233.ABA
  );

  const auditoria = ss.getSheetByName(
    DCP_GESTAO_COMBUSTIVEIS_233.AUDITORIA
  );

  if (!aba || !auditoria) {
    throw new Error(
      'Aba COMBUSTIVEIS ou AUDITORIA nao encontrada.'
    );
  }

  const cabecalhosCombustiveis = aba
    .getRange(
      1,
      1,
      1,
      DCP_GESTAO_COMBUSTIVEIS_233.CABECALHOS.length
    )
    .getDisplayValues()[0];

  const cabecalhosAuditoria = auditoria
    .getRange(
      1,
      1,
      1,
      DCP_GESTAO_COMBUSTIVEIS_233.CABECALHOS_AUDITORIA.length
    )
    .getDisplayValues()[0];

  if (
    JSON.stringify(cabecalhosCombustiveis) !==
    JSON.stringify(DCP_GESTAO_COMBUSTIVEIS_233.CABECALHOS)
  ) {
    throw new Error(
      'Cabecalhos da aba COMBUSTIVEIS divergentes.'
    );
  }

  if (
    JSON.stringify(cabecalhosAuditoria) !==
    JSON.stringify(DCP_GESTAO_COMBUSTIVEIS_233.CABECALHOS_AUDITORIA)
  ) {
    throw new Error(
      'Cabecalhos da aba AUDITORIA divergentes.'
    );
  }

  return {
    aba: aba,
    auditoria: auditoria
  };
}


/**
 * Le os registros do catalogo.
 *
 * incluirInativos = false: somente ATIVOS.
 * incluirInativos = true: todos os registros.
 */
function DCP_LISTAR_COMBUSTIVEIS_233(incluirInativos) {
  try {
    const estrutura =
      DCP_CONFERIR_ESTRUTURA_GESTAO_COMBUSTIVEIS_233();

    const aba = estrutura.aba;
    const ultimaLinha = aba.getLastRow();

    if (ultimaLinha < 2) {
      return {
        sucesso: true,
        quantidade: 0,
        dados: [],
        mensagem: 'Nenhum combustivel cadastrado.'
      };
    }

    const valores = aba
      .getRange(2, 1, ultimaLinha - 1, 12)
      .getValues();

    const dados = valores
      .filter(linha => {
        const id = String(linha[0] || '').trim();
        const status = String(linha[6] || '')
          .trim()
          .toUpperCase();

        if (!id) return false;

        return incluirInativos === true
          ? true
          : status === 'ATIVO';
      })
      .map(linha => ({
        combustivelID: linha[0],
        codigoCombustivel: linha[1],
        nomeCombustivel: linha[2],
        tipoEnergia: linha[3],
        unidadeMedida: linha[4],
        observacoes: linha[5],
        status: linha[6],
        criadoEm: linha[7],
        atualizadoEm: linha[8],
        criadoPor: linha[9],
        atualizadoPor: linha[10],
        codigoReferencia: linha[11]
      }));

    return {
      sucesso: true,
      quantidade: dados.length,
      dados: dados,
      mensagem: 'Consulta realizada com sucesso.'
    };

  } catch (erro) {
    return {
      sucesso: false,
      quantidade: 0,
      dados: [],
      mensagem: erro.message
    };
  }
}


/**
 * Atalho para listar somente combustiveis ativos.
 */
function DCP_LISTAR_COMBUSTIVEIS_ATIVOS_233() {
  return DCP_LISTAR_COMBUSTIVEIS_233(false);
}


/**
 * Atalho para listar todos os combustiveis,
 * incluindo os inativos.
 */
function DCP_LISTAR_TODOS_COMBUSTIVEIS_233() {
  return DCP_LISTAR_COMBUSTIVEIS_233(true);
}


/**
 * Inativa um combustivel pelo UUID.
 *
 * Nao apaga o registro.
 * Nao reutiliza o codigo.
 * Registra tentativa e confirmacao na auditoria.
 */
function DCP_INATIVAR_COMBUSTIVEL_233(combustivelID) {

  const id = String(combustivelID || '').trim();

  if (!id) {
    return {
      sucesso: false,
      mensagem: 'Informe o ID do combustivel.'
    };
  }

  const lock = LockService.getDocumentLock();

  if (!lock.tryLock(30000)) {
    return {
      sucesso: false,
      mensagem: 'Sistema ocupado. Tente novamente.'
    };
  }

  let tentativaRegistrada = false;
  let alteracaoRealizada = false;

  try {
    const estrutura =
      DCP_CONFERIR_ESTRUTURA_GESTAO_COMBUSTIVEIS_233();

    const aba = estrutura.aba;
    const auditoria = estrutura.auditoria;

    const ultimaLinha = aba.getLastRow();

    if (ultimaLinha < 2) {
      throw new Error('Nao existem combustiveis cadastrados.');
    }

    const valores = aba
      .getRange(2, 1, ultimaLinha - 1, 12)
      .getValues();

    const indice = valores.findIndex(
      linha => String(linha[0]).trim() === id
    );

    if (indice === -1) {
      throw new Error('Combustivel nao encontrado.');
    }

    const linhaPlanilha = indice + 2;
    const registro = valores[indice];

    const statusAtual = String(registro[6] || '')
      .trim()
      .toUpperCase();

    if (statusAtual === 'INATIVO') {
      return {
        sucesso: true,
        jaInativo: true,
        mensagem: 'Este combustivel ja esta inativo.'
      };
    }

    if (statusAtual !== 'ATIVO') {
      throw new Error(
        'Status atual nao permite inativacao: ' + statusAtual
      );
    }

    const agora = new Date();

    const registroAnterior = {
      CombustivelID: registro[0],
      CodigoCombustivel: registro[1],
      NomeCombustivel: registro[2],
      Status: registro[6],
      CodigoReferencia: registro[11]
    };

    const registroNovo = {
      CombustivelID: registro[0],
      CodigoCombustivel: registro[1],
      NomeCombustivel: registro[2],
      Status: 'INATIVO',
      CodigoReferencia: registro[11]
    };

    // Registra a tentativa antes da alteracao.
    auditoria.appendRow([
      Utilities.getUuid(),
      agora,
      DCP_GESTAO_COMBUSTIVEIS_233.ADMIN,
      'TENTATIVA_INATIVACAO',
      'COMBUSTIVEIS',
      id,
      JSON.stringify(registroAnterior),
      JSON.stringify(registroNovo),
      'PENDENTE',
      'APPS_SCRIPT',
      'Tentativa de inativacao de combustivel.'
    ]);

    tentativaRegistrada = true;
    SpreadsheetApp.flush();

    // Atualiza somente os campos de controle necessarios.
    aba.getRange(linhaPlanilha, 7).setValue('INATIVO');
    aba.getRange(linhaPlanilha, 9).setValue(agora);
    aba.getRange(linhaPlanilha, 11).setValue(
      DCP_GESTAO_COMBUSTIVEIS_233.ADMIN
    );

    SpreadsheetApp.flush();
    alteracaoRealizada = true;

    // Registra a confirmacao.
    auditoria.appendRow([
      Utilities.getUuid(),
      new Date(),
      DCP_GESTAO_COMBUSTIVEIS_233.ADMIN,
      'CONFIRMACAO_INATIVACAO',
      'COMBUSTIVEIS',
      id,
      JSON.stringify(registroAnterior),
      JSON.stringify(registroNovo),
      'SUCESSO',
      'APPS_SCRIPT',
      'Combustivel inativado com sucesso.'
    ]);

    SpreadsheetApp.flush();

    return {
      sucesso: true,
      auditoriaPendente: false,
      id: id,
      codigo: registro[1],
      mensagem: 'Combustivel inativado com sucesso.'
    };

  } catch (erro) {

    // Se a alteracao ja ocorreu, nao declarar que ela falhou.
    if (alteracaoRealizada) {
      return {
        sucesso: true,
        auditoriaPendente: true,
        id: id,
        mensagem:
          'Combustivel inativado, mas a confirmacao da auditoria ' +
          'precisa ser verificada. Nao execute novamente.'
      };
    }

    // Registra falha somente quando a tentativa foi registrada.
    if (tentativaRegistrada) {
      try {
        const auditoria = SpreadsheetApp
          .getActiveSpreadsheet()
          .getSheetByName(
            DCP_GESTAO_COMBUSTIVEIS_233.AUDITORIA
          );

        auditoria.appendRow([
          Utilities.getUuid(),
          new Date(),
          DCP_GESTAO_COMBUSTIVEIS_233.ADMIN,
          'FALHA_INATIVACAO',
          'COMBUSTIVEIS',
          id,
          '{}',
          '{}',
          'FALHA',
          'APPS_SCRIPT',
          String(erro.message)
        ]);

        SpreadsheetApp.flush();

      } catch (erroAuditoria) {
        Logger.log(
          'Falha adicional na auditoria: ' +
          erroAuditoria.message
        );
      }
    }

    return {
      sucesso: false,
      mensagem: erro.message
    };

  } finally {
    lock.releaseLock();
  }
}


/**
 * Teste de leitura do catalogo.
 * Nao altera nenhum dado.
 */
function DCP_TESTAR_CONSULTA_COMBUSTIVEIS_233() {

  Logger.log('===== TESTE CONSULTA COMBUSTIVEIS =====');

  const ativos = DCP_LISTAR_COMBUSTIVEIS_ATIVOS_233();

  Logger.log('CONSULTA DE ATIVOS:');
  Logger.log(JSON.stringify(ativos, null, 2));

  const todos = DCP_LISTAR_TODOS_COMBUSTIVEIS_233();

  Logger.log('CONSULTA DE TODOS:');
  Logger.log(JSON.stringify(todos, null, 2));

  if (!ativos.sucesso || !todos.sucesso) {
    throw new Error('Falha na consulta do catalogo.');
  }

  Logger.log('TOTAL DE ATIVOS: ' + ativos.quantidade);
  Logger.log('TOTAL DE REGISTROS: ' + todos.quantidade);
  Logger.log('RESULTADO: CONSULTAS APROVADAS.');
}

function DCP_TESTAR_INATIVACAO_GASOLINA_233() {

  const id = 'feb4ced9-05be-481b-9545-7e452da66afe';

  Logger.log('===== TESTE INATIVACAO GASOLINA =====');

  const resultado = DCP_INATIVAR_COMBUSTIVEL_233(id);

  Logger.log('RESULTADO:');
  Logger.log(JSON.stringify(resultado, null, 2));

  if (!resultado.sucesso) {
    throw new Error(resultado.mensagem);
  }

  Logger.log('TESTE DE INATIVACAO EXECUTADO.');
}

/**
 * REATIVA UM COMBUSTIVEL EXISTENTE.
 *
 * Preserva UUID, codigo, data de criacao e criador.
 * Verifica duplicidades entre registros ativos.
 * Registra tentativa e confirmacao na auditoria.
 */
function DCP_REATIVAR_COMBUSTIVEL_233(combustivelID) {

  const id = String(combustivelID || '').trim();

  if (!id) {
    return {
      sucesso: false,
      mensagem: 'Informe o ID do combustivel.'
    };
  }

  const lock = LockService.getDocumentLock();

  if (!lock.tryLock(30000)) {
    return {
      sucesso: false,
      mensagem: 'Sistema ocupado. Tente novamente.'
    };
  }

  let tentativaRegistrada = false;
  let alteracaoTentada = false;

  try {

    const estrutura =
      DCP_CONFERIR_ESTRUTURA_GESTAO_COMBUSTIVEIS_233();

    const aba = estrutura.aba;
    const auditoria = estrutura.auditoria;

    const ultimaLinha = aba.getLastRow();

    if (ultimaLinha < 2) {
      throw new Error('Nenhum combustivel cadastrado.');
    }

    const valores = aba
      .getRange(2, 1, ultimaLinha - 1, 12)
      .getValues();

    const indice = valores.findIndex(
      linha => String(linha[0]).trim() === id
    );

    if (indice === -1) {
      throw new Error('Combustivel nao encontrado.');
    }

    const linhaPlanilha = indice + 2;
    const registro = valores[indice];

    const statusAtual = String(registro[6] || '')
      .trim()
      .toUpperCase();

    if (statusAtual === 'ATIVO') {
      return {
        sucesso: true,
        jaAtivo: true,
        mensagem: 'Este combustivel ja esta ativo.'
      };
    }

    if (statusAtual !== 'INATIVO') {
      throw new Error(
        'Somente combustiveis INATIVOS podem ser reativados.'
      );
    }

    // Verifica duplicidade de nome e codigo de referencia.
    const nomeAtual = String(registro[2] || '')
      .trim()
      .toUpperCase();

    const referenciaAtual = String(registro[11] || '')
      .trim()
      .toUpperCase();

    for (let i = 0; i < valores.length; i++) {

      if (i === indice) continue;

      const outro = valores[i];

      const outroID = String(outro[0] || '').trim();
      const outroStatus = String(outro[6] || '')
        .trim()
        .toUpperCase();

      if (!outroID || outroStatus !== 'ATIVO') continue;

      const outroNome = String(outro[2] || '')
        .trim()
        .toUpperCase();

      const outraReferencia = String(outro[11] || '')
        .trim()
        .toUpperCase();

      if (nomeAtual && nomeAtual === outroNome) {
        throw new Error(
          'Ja existe outro combustivel ativo com este nome.'
        );
      }

      if (
        referenciaAtual &&
        referenciaAtual === outraReferencia
      ) {
        throw new Error(
          'Ja existe outro combustivel ativo com este codigo de referencia.'
        );
      }
    }

    const agora = new Date();

    const registroAnterior = {
      CombustivelID: registro[0],
      CodigoCombustivel: registro[1],
      NomeCombustivel: registro[2],
      Status: registro[6],
      CodigoReferencia: registro[11]
    };

    const registroNovo = {
      CombustivelID: registro[0],
      CodigoCombustivel: registro[1],
      NomeCombustivel: registro[2],
      Status: 'ATIVO',
      CodigoReferencia: registro[11]
    };

    // Registra a tentativa antes de alterar o cadastro.
    auditoria.appendRow([
      Utilities.getUuid(),
      agora,
      DCP_GESTAO_COMBUSTIVEIS_233.ADMIN,
      'TENTATIVA_REATIVACAO',
      'COMBUSTIVEIS',
      id,
      JSON.stringify(registroAnterior),
      JSON.stringify(registroNovo),
      'PENDENTE',
      'APPS_SCRIPT',
      'Tentativa de reativacao de combustivel.'
    ]);

    tentativaRegistrada = true;
    SpreadsheetApp.flush();

    // Atualiza G:K preservando os dados de criacao.
    const dadosControle = [[
      'ATIVO',
      registro[7],
      agora,
      registro[9],
      DCP_GESTAO_COMBUSTIVEIS_233.ADMIN
    ]];

    alteracaoTentada = true;

    aba.getRange(linhaPlanilha, 7, 1, 5)
      .setValues(dadosControle);

    SpreadsheetApp.flush();

    // Registra a confirmacao.
    auditoria.appendRow([
      Utilities.getUuid(),
      new Date(),
      DCP_GESTAO_COMBUSTIVEIS_233.ADMIN,
      'CONFIRMACAO_REATIVACAO',
      'COMBUSTIVEIS',
      id,
      JSON.stringify(registroAnterior),
      JSON.stringify(registroNovo),
      'SUCESSO',
      'APPS_SCRIPT',
      'Combustivel reativado com sucesso.'
    ]);

    SpreadsheetApp.flush();

    return {
      sucesso: true,
      auditoriaPendente: false,
      id: id,
      codigo: registro[1],
      mensagem: 'Combustivel reativado com sucesso.'
    };

  } catch (erro) {

    // Se a alteracao foi tentada, confere o estado real.
    if (alteracaoTentada) {
      try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const aba = ss.getSheetByName(
          DCP_GESTAO_COMBUSTIVEIS_233.ABA
        );

        const ultimaLinha = aba.getLastRow();

        if (ultimaLinha >= 2) {
          const dados = aba
            .getRange(2, 1, ultimaLinha - 1, 7)
            .getValues();

          const registroAtual = dados.find(
            linha => String(linha[0]).trim() === id
          );

          if (
            registroAtual &&
            String(registroAtual[6]).trim().toUpperCase() === 'ATIVO'
          ) {
            return {
              sucesso: true,
              auditoriaPendente: true,
              id: id,
              mensagem:
                'Combustivel reativado. Verifique a confirmacao ' +
                'da auditoria antes de repetir a operacao.'
            };
          }
        }
      } catch (erroVerificacao) {
        Logger.log(
          'Falha ao conferir estado: ' +
          erroVerificacao.message
        );
      }
    }

    if (tentativaRegistrada) {
      try {
        auditoriaFalha = SpreadsheetApp
          .getActiveSpreadsheet()
          .getSheetByName(
            DCP_GESTAO_COMBUSTIVEIS_233.AUDITORIA
          );

        auditoriaFalha.appendRow([
          Utilities.getUuid(),
          new Date(),
          DCP_GESTAO_COMBUSTIVEIS_233.ADMIN,
          'FALHA_REATIVACAO',
          'COMBUSTIVEIS',
          id,
          '{}',
          '{}',
          'FALHA',
          'APPS_SCRIPT',
          String(erro.message)
        ]);

        SpreadsheetApp.flush();

      } catch (erroAuditoria) {
        Logger.log(
          'Falha ao registrar erro: ' +
          erroAuditoria.message
        );
      }
    }

    return {
      sucesso: false,
      mensagem: erro.message
    };

  } finally {
    lock.releaseLock();
  }
}


/**
 * TESTE CONTROLADO DE REATIVACAO DA GASOLINA.
 */
function DCP_TESTAR_REATIVACAO_GASOLINA_233() {

  const id =
    'feb4ced9-05be-481b-9545-7e452da66afe';

  Logger.log('===== TESTE REATIVACAO GASOLINA =====');

  const resultado =
    DCP_REATIVAR_COMBUSTIVEL_233(id);

  Logger.log(JSON.stringify(resultado, null, 2));

  if (!resultado.sucesso) {
    throw new Error(resultado.mensagem);
  }

  Logger.log('TESTE DE REATIVACAO EXECUTADO.');
}
