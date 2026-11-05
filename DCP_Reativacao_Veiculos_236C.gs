
/**
 * DRIVER CONTROL PRO
 * MÓDULO 2.3.6C
 * REATIVAÇÃO DE VEÍCULOS
 *
 * Reativa veículos inativos mediante validação cadastral.
 * Registra histórico e auditoria.
 *
 * IMPORTANTE:
 * A função operacional é DCP_REATIVAR_VEICULO_236C(dados).
 */

const DCP_REATIVACAO_VEICULOS_236C = {
  ABA_VEICULOS: 'VEICULOS',
  ABA_VINCULOS: 'VINCULOS',
  ABA_MOTORISTAS: 'MOTORISTAS',
  ABA_HISTORICO: 'HISTORICO_STATUS_VEICULOS',
  ABA_AUDITORIA: 'AUDITORIA',
  ABA_CONFIG: 'CONFIG',

  PREFIXO_HISTORICO: 'HVS',
  CHAVE_SEQUENCIA: 'SEQ_HVS',
  USUARIO: 'ADMIN_CENTRAL',

  STATUS_ORIGEM: 'INATIVO',
  STATUS_DESTINO: 'ATIVO',
  LIMITE_MOTIVO: 500
};

const DCP_RC236C_CAB_VEICULOS = [
  'VeiculoID', 'CodigoVeiculo', 'Placa', 'Marca', 'Modelo',
  'AnoFabricacao', 'AnoModelo', 'TipoVeiculo', 'TipoPropulsao',
  'CombustivelPrincipal', 'CombustivelSecundario',
  'ProprietarioTipo', 'ProprietarioMotoristaID', 'ProprietarioNome',
  'CapacidadeTanqueLitros', 'CapacidadeBateriaKWh',
  'HodometroAtual', 'DataCadastro', 'Observacoes', 'Status',
  'CriadoEm', 'AtualizadoEm', 'CriadoPor', 'AtualizadoPor'
];

const DCP_RC236C_CAB_VINCULOS = [
  'VinculoID', 'CodigoVinculo', 'MotoristaID', 'VeiculoID',
  'TipoVinculo', 'InicioEm', 'FimEm', 'CompartilhamentoAutorizado',
  'ValorAluguel', 'PeriodicidadeAluguel', 'Observacoes', 'Status',
  'CriadoEm', 'AtualizadoEm', 'CriadoPor', 'AtualizadoPor'
];

const DCP_RC236C_CAB_MOTORISTAS = [
  'MotoristaID', 'CodigoMotorista', 'NomeCompleto', 'NomeExibicao',
  'Telefone', 'Email', 'DataCadastro', 'Observacoes', 'Status',
  'CriadoEm', 'AtualizadoEm', 'CriadoPor', 'AtualizadoPor'
];

const DCP_RC236C_CAB_HISTORICO = [
  'HistoricoStatusID', 'CodigoHistoricoStatus', 'VeiculoID',
  'CodigoVeiculo', 'StatusAnterior', 'StatusNovo', 'Motivo',
  'DataHora', 'UsuarioID', 'Observacoes', 'Resultado', 'AuditoriaID'
];

const DCP_RC236C_CAB_AUDITORIA = [
  'AuditoriaID', 'DataHora', 'UsuarioID', 'Acao', 'Tabela',
  'RegistroID', 'ValorAnterior', 'ValorNovo', 'Resultado',
  'Origem', 'Detalhes'
];

const DCP_RC236C_CAB_CONFIG = [
  'Chave', 'Valor', 'TipoValor', 'Descricao',
  'Editavel', 'AtualizadoEm'
];


/**
 * Confere a existência de uma aba e seus cabeçalhos.
 */
function DCP_RC236C_CONFERIR_ABA(nome, cabecalhos) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(nome);

  if (!aba) {
    throw new Error('Aba não encontrada: ' + nome);
  }

  const encontrados = aba
    .getRange(1, 1, 1, cabecalhos.length)
    .getValues()[0]
    .map(v => String(v || '').trim());

  const divergencias = [];

  cabecalhos.forEach((campo, i) => {
    if (encontrados[i] !== campo) {
      divergencias.push({
        coluna: i + 1,
        esperado: campo,
        encontrado: encontrados[i]
      });
    }
  });

  if (divergencias.length) {
    throw new Error(
      'Cabeçalhos divergentes em ' + nome + ': ' +
      JSON.stringify(divergencias)
    );
  }

  return aba;
}


/**
 * Localiza veículo por UUID ou código.
 * Exige resultado único.
 */
function DCP_RC236C_LOCALIZAR_VEICULO(aba, identificador) {
  const ultimaLinha = aba.getLastRow();

  if (ultimaLinha < 2) {
    throw new Error('Não existem veículos cadastrados.');
  }

  const dados = aba.getRange(2, 1, ultimaLinha - 1, 24).getValues();
  const encontrados = [];

  dados.forEach((linha, i) => {
    const uuid = String(linha[0] || '').trim();
    const codigo = String(linha[1] || '').trim();

    if (uuid === identificador || codigo === identificador) {
      encontrados.push({
        linha: i + 2,
        dados: linha
      });
    }
  });

  if (encontrados.length === 0) {
    throw new Error('Veículo não encontrado: ' + identificador);
  }

  if (encontrados.length !== 1) {
    throw new Error('Identificador duplicado. Operação bloqueada.');
  }

  return encontrados[0];
}


/**
 * Valida a placa brasileira.
 */
function DCP_RC236C_PLACA_VALIDA(placa) {
  return (
    /^[A-Z]{3}[0-9]{4}$/.test(placa) ||
    /^[A-Z]{3}[0-9][A-Z][0-9]{2}$/.test(placa)
  );
}


/**
 * Verifica duplicidade da placa em todos os veículos,
 * independentemente do status, exceto o próprio veículo.
 */
function DCP_RC236C_VERIFICAR_PLACA(aba, placa, veiculoID) {
  const ultimaLinha = aba.getLastRow();

  if (ultimaLinha < 2) return;

  const dados = aba.getRange(2, 1, ultimaLinha - 1, 24).getValues();

  const duplicado = dados.some(linha => {
    const placaExistente = String(linha[2] || '')
      .trim()
      .toUpperCase()
      .replace(/[\s-]/g, '');

    const idExistente = String(linha[0] || '').trim();

    return (
      placaExistente === placa &&
      idExistente !== veiculoID
    );
  });

  if (duplicado) {
    throw new Error(
      'A placa já está cadastrada em outro veículo. Reativação bloqueada.'
    );
  }
}


/**
 * Confere o proprietário do veículo.
 */
function DCP_RC236C_VALIDAR_PROPRIETARIO(
  abaMotoristas,
  tipo,
  motoristaID,
  nomeProprietario
) {
  tipo = String(tipo || '').trim().toUpperCase();
  motoristaID = String(motoristaID || '').trim();
  nomeProprietario = String(nomeProprietario || '').trim();

  if (tipo === 'TERCEIRO') {
    if (
      !nomeProprietario ||
      nomeProprietario.length > 150 ||
      motoristaID !== ''
    ) {
      throw new Error('Dados do proprietário terceiro inconsistentes.');
    }

    return true;
  }

  if (tipo !== 'MOTORISTA') {
    throw new Error('Tipo de proprietário inválido.');
  }

  if (!motoristaID) {
    throw new Error('Proprietário motorista não informado.');
  }

  const ultimaLinha = abaMotoristas.getLastRow();
  const motoristas = ultimaLinha >= 2
    ? abaMotoristas.getRange(2, 1, ultimaLinha - 1, 13).getValues()
    : [];

  const encontrados = motoristas.filter(linha =>
    String(linha[0] || '').trim() === motoristaID
  );

  if (encontrados.length !== 1) {
    throw new Error('Proprietário motorista não encontrado ou duplicado.');
  }

  const motorista = encontrados[0];
  const status = String(motorista[8] || '').trim().toUpperCase();
  const nome = String(motorista[2] || '').trim();

  if (status !== 'ATIVO') {
    throw new Error('O proprietário motorista não está ATIVO.');
  }

  if (!nome || nome !== nomeProprietario) {
    throw new Error('Nome do proprietário diverge do cadastro do motorista.');
  }

  return true;
}


/**
 * Confere se existem vínculos ativos para o veículo.
 */
function DCP_RC236C_VERIFICAR_VINCULOS(aba, veiculoID) {
  const ultimaLinha = aba.getLastRow();

  if (ultimaLinha < 2) return;

  const dados = aba.getRange(2, 1, ultimaLinha - 1, 16).getValues();

  const ativos = dados.filter(linha =>
    String(linha[3] || '').trim() === veiculoID &&
    String(linha[11] || '').trim().toUpperCase() === 'ATIVO'
  );

  if (ativos.length > 0) {
    throw new Error(
      'Existem ' + ativos.length +
      ' vínculo(s) ativo(s). Reativação bloqueada.'
    );
  }
}


/**
 * Reserva o próximo código HVS.
 * Deve ser chamada com DocumentLock já adquirido.
 * Lacunas na sequência em caso de falha são permitidas.
 */
function DCP_RC236C_GERAR_CODIGO_HISTORICO(abaConfig, abaHistorico) {
  const ultimaLinhaConfig = abaConfig.getLastRow();

  const config = ultimaLinhaConfig >= 2
    ? abaConfig.getRange(2, 1, ultimaLinhaConfig - 1, 6).getValues()
    : [];

  let linhaConfig = -1;
  let sequenciaConfig = 0;

  config.forEach((linha, i) => {
    if (String(linha[0] || '').trim() ===
        DCP_REATIVACAO_VEICULOS_236C.CHAVE_SEQUENCIA) {
      if (linhaConfig !== -1) {
        throw new Error('Chave SEQ_HVS duplicada na CONFIG.');
      }

      linhaConfig = i + 2;
      sequenciaConfig = Number(linha[1]) || 0;
    }
  });

  let maiorHistorico = 0;
  const ultimaLinhaHistorico = abaHistorico.getLastRow();

  if (ultimaLinhaHistorico >= 2) {
    const codigos = abaHistorico
      .getRange(2, 2, ultimaLinhaHistorico - 1, 1)
      .getValues();

    codigos.forEach(linha => {
      const codigo = String(linha[0] || '').trim();
      const match = codigo.match(/^HVS-(\d+)$/);

      if (match) {
        maiorHistorico = Math.max(
          maiorHistorico,
          Number(match[1])
        );
      }
    });
  }

  const proxima = Math.max(sequenciaConfig, maiorHistorico) + 1;
  const codigo = 'HVS-' + String(proxima).padStart(4, '0');
  const agora = new Date();

  if (linhaConfig === -1) {
    abaConfig.appendRow([
      DCP_REATIVACAO_VEICULOS_236C.CHAVE_SEQUENCIA,
      proxima,
      'NUMERO',
      'Sequência de histórico de status de veículos',
      false,
      agora
    ]);
  } else {
    abaConfig.getRange(linhaConfig, 2).setValue(proxima);
    abaConfig.getRange(linhaConfig, 6).setValue(agora);
  }

  return codigo;
}


/**
 * Confirma que um UUID foi gravado na linha esperada.
 */
function DCP_RC236C_CONFIRMAR_UUID(aba, linha, coluna, uuid) {
  SpreadsheetApp.flush();

  const valor = String(
    aba.getRange(linha, coluna).getValue() || ''
  ).trim();

  if (valor !== uuid) {
    throw new Error('Não foi possível confirmar a gravação do registro.');
  }
}


/**
 * Função operacional de reativação.
 *
 * Exemplo:
 * DCP_REATIVAR_VEICULO_236C({
 *   identificador: 'VEI-0001',
 *   motivo: 'Retorno do veículo à operação.'
 * });
 */
function DCP_REATIVAR_VEICULO_236C(dados) {
  const lock = LockService.getDocumentLock();

  if (!lock.tryLock(30000)) {
    throw new Error('Não foi possível obter o bloqueio da planilha.');
  }

  let abaVeiculos;
  let abaHistorico;
  let abaAuditoria;
  let linhaVeiculo = -1;
  let veiculoID = '';
  let codigoVeiculo = '';
  let statusAlterado = false;
  let linhaHistorico = -1;
  let historicoID = '';
  let codigoHistorico = '';
  let auditoriaID = '';
  let historicoGravado = false;

  try {
    dados = dados || {};

    const identificador = String(dados.identificador || '').trim();
    const motivo = String(dados.motivo || '').trim();

    if (!identificador) {
      throw new Error('Informe o UUID ou código do veículo.');
    }

    if (!motivo) {
      throw new Error('O motivo da reativação é obrigatório.');
    }

    if (motivo.length > DCP_REATIVACAO_VEICULOS_236C.LIMITE_MOTIVO) {
      throw new Error('O motivo ultrapassou 500 caracteres.');
    }

    abaVeiculos = DCP_RC236C_CONFERIR_ABA(
      DCP_REATIVACAO_VEICULOS_236C.ABA_VEICULOS,
      DCP_RC236C_CAB_VEICULOS
    );

    const abaVinculos = DCP_RC236C_CONFERIR_ABA(
      DCP_REATIVACAO_VEICULOS_236C.ABA_VINCULOS,
      DCP_RC236C_CAB_VINCULOS
    );

    const abaMotoristas = DCP_RC236C_CONFERIR_ABA(
      DCP_REATIVACAO_VEICULOS_236C.ABA_MOTORISTAS,
      DCP_RC236C_CAB_MOTORISTAS
    );

    abaHistorico = DCP_RC236C_CONFERIR_ABA(
      DCP_REATIVACAO_VEICULOS_236C.ABA_HISTORICO,
      DCP_RC236C_CAB_HISTORICO
    );

    abaAuditoria = DCP_RC236C_CONFERIR_ABA(
      DCP_REATIVACAO_VEICULOS_236C.ABA_AUDITORIA,
      DCP_RC236C_CAB_AUDITORIA
    );

    const abaConfig = DCP_RC236C_CONFERIR_ABA(
      DCP_REATIVACAO_VEICULOS_236C.ABA_CONFIG,
      DCP_RC236C_CAB_CONFIG
    );

    const localizado = DCP_RC236C_LOCALIZAR_VEICULO(
      abaVeiculos,
      identificador
    );

    linhaVeiculo = localizado.linha;

    const v = localizado.dados;

    veiculoID = String(v[0] || '').trim();
    codigoVeiculo = String(v[1] || '').trim();

    const statusAtual = String(v[19] || '').trim().toUpperCase();

    if (statusAtual !== 'INATIVO') {
      throw new Error(
        'Somente veículos INATIVOS podem ser reativados. Status atual: ' +
        statusAtual
      );
    }

    const placa = String(v[2] || '')
      .trim()
      .toUpperCase()
      .replace(/[\s-]/g, '');

    if (!DCP_RC236C_PLACA_VALIDA(placa)) {
      throw new Error('A placa do veículo é inválida.');
    }

    if (
      !String(v[3] || '').trim() ||
      !String(v[4] || '').trim()
    ) {
      throw new Error('Marca e modelo são obrigatórios.');
    }

    DCP_RC236C_VERIFICAR_PLACA(
      abaVeiculos,
      placa,
      veiculoID
    );

    DCP_RC236C_VALIDAR_PROPRIETARIO(
      abaMotoristas,
      v[11],
      v[12],
      v[13]
    );

    DCP_RC236C_VERIFICAR_VINCULOS(
      abaVinculos,
      veiculoID
    );

    // Reserva identificadores antes das gravações.
    historicoID = Utilities.getUuid();
    auditoriaID = Utilities.getUuid();

    codigoHistorico = DCP_RC236C_GERAR_CODIGO_HISTORICO(
      abaConfig,
      abaHistorico
    );

    const agora = new Date();
    const usuario = DCP_REATIVACAO_VEICULOS_236C.USUARIO;

    // Atualiza somente o status e os campos de atualização.
    abaVeiculos.getRange(linhaVeiculo, 20).setValue('ATIVO');

    // Marcado imediatamente após setValue para permitir rollback
    // caso flush ou confirmação apresentem erro.
    statusAlterado = true;

    abaVeiculos.getRange(linhaVeiculo, 22).setValue(agora);
    abaVeiculos.getRange(linhaVeiculo, 24).setValue(usuario);

    SpreadsheetApp.flush();

    const statusConfirmado = String(
      abaVeiculos.getRange(linhaVeiculo, 20).getValue() || ''
    ).trim().toUpperCase();

    if (statusConfirmado !== 'ATIVO') {
      throw new Error('Não foi possível confirmar o novo status ATIVO.');
    }

    // Registra histórico.
    const novaLinhaHistorico = [
      historicoID,
      codigoHistorico,
      veiculoID,
      codigoVeiculo,
      'INATIVO',
      'ATIVO',
      motivo,
      agora,
      usuario,
      'Reativação de veículo.',
      'SUCESSO',
      auditoriaID
    ];

    abaHistorico.appendRow(novaLinhaHistorico);
    linhaHistorico = abaHistorico.getLastRow();

    DCP_RC236C_CONFIRMAR_UUID(
      abaHistorico,
      linhaHistorico,
      1,
      historicoID
    );

    historicoGravado = true;

    // Registra auditoria.
    const detalhes =
      'Histórico ' + codigoHistorico +
      '. Motivo: ' + motivo;

    const novaLinhaAuditoria = [
      auditoriaID,
      agora,
      usuario,
      'ALTERACAO',
      'VEICULOS',
      veiculoID,
      'INATIVO',
      'ATIVO',
      'SUCESSO',
      'SCRIPT',
      detalhes
    ];

    abaAuditoria.appendRow(novaLinhaAuditoria);
    const linhaAuditoria = abaAuditoria.getLastRow();

    DCP_RC236C_CONFIRMAR_UUID(
      abaAuditoria,
      linhaAuditoria,
      1,
      auditoriaID
    );

    return {
      sucesso: true,
      codigoVeiculo: codigoVeiculo,
      veiculoID: veiculoID,
      statusAnterior: 'INATIVO',
      statusNovo: 'ATIVO',
      historicoID: historicoID,
      codigoHistorico: codigoHistorico,
      auditoriaID: auditoriaID,
      historicoConfirmado: true,
      auditoriaConfirmada: true,
      motivo: motivo,
      mensagem:
        'Veículo ' + codigoVeiculo +
        ' reativado. Histórico e auditoria confirmados.'
    };

  } catch (erro) {
    let rollbackConfirmado = false;
    let statusFinal = 'NAO_CONFIRMADO';

    // Se o status chegou a ser alterado, tenta restaurar INATIVO.
    if (statusAlterado && abaVeiculos && linhaVeiculo > 0) {
      try {
        abaVeiculos.getRange(linhaVeiculo, 20).setValue('INATIVO');
        SpreadsheetApp.flush();

        statusFinal = String(
          abaVeiculos.getRange(linhaVeiculo, 20).getValue() || ''
        ).trim().toUpperCase();

        rollbackConfirmado = statusFinal === 'INATIVO';
      } catch (erroRollback) {
        statusFinal = 'ERRO_NO_ROLLBACK';
      }
    }

    // Se o histórico chegou a ser gravado, identifica a operação como falha.
    if (historicoGravado && linhaHistorico > 0) {
      try {
        abaHistorico.getRange(linhaHistorico, 11).setValue('FALHA');
        SpreadsheetApp.flush();
      } catch (erroHistorico) {
        // A conferência manual será necessária.
      }
    }

    // Registra o erro como novo evento, sem editar auditorias anteriores.
    if (abaAuditoria && veiculoID) {
      try {
        const auditoriaErroID = Utilities.getUuid();
        const agoraErro = new Date();

        abaAuditoria.appendRow([
          auditoriaErroID,
          agoraErro,
          DCP_REATIVACAO_VEICULOS_236C.USUARIO,
          'ERRO',
          'VEICULOS',
          veiculoID,
          'INATIVO',
          statusFinal,
          'FALHA',
          'SCRIPT',
          'Falha na reativação. ' + String(erro.message || erro) +
          '. Rollback confirmado: ' + rollbackConfirmado +
          '. Histórico: ' + (codigoHistorico || 'Não gerado')
        ]);

        SpreadsheetApp.flush();
      } catch (erroAuditoria) {
        // O estado deverá ser conferido manualmente.
      }
    }

    return {
      sucesso: false,
      codigoVeiculo: codigoVeiculo,
      veiculoID: veiculoID,
      codigoHistorico: codigoHistorico,
      rollbackConfirmado: rollbackConfirmado,
      statusFinal: statusFinal,
      mensagem: String(erro.message || erro),
      orientacao:
        'Confira VEICULOS, HISTORICO_STATUS_VEICULOS e AUDITORIA antes de tentar novamente.'
    };

  } finally {
    lock.releaseLock();
  }
}



function DCP_TESTE_OPERACIONAL_REATIVACAO_236C() {
  const resultado = DCP_REATIVAR_VEICULO_236C({
    identificador: 'VEI-0001',
    motivo: 'Retorno do veículo à operação após teste de inativação.'
  });

  Logger.log(JSON.stringify(resultado, null, 2));
}
