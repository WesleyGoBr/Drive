/**
 * DRIVER CONTROL PRO
 * MÓDULO 2.3.6C
 * DIAGNÓSTICO DE REATIVAÇÃO DE VEÍCULOS
 *
 * Somente leitura.
 * Não altera veículos, histórico, auditoria ou configurações.
 */

const DCP_DIAG_REATIVACAO_236C = {
  ABA_VEICULOS: 'VEICULOS',
  ABA_VINCULOS: 'VINCULOS',
  ABA_MOTORISTAS: 'MOTORISTAS',

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

  CABECALHOS_VINCULOS: [
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
  ],

  CABECALHOS_MOTORISTAS: [
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
  ]
};


/**
 * Confere os cabeçalhos de uma aba.
 */
function DCP_DC236C_CONFERIR_ABA(nomeAba, cabecalhosEsperados) {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const aba = planilha.getSheetByName(nomeAba);

  if (!aba) {
    throw new Error('Aba não encontrada: ' + nomeAba);
  }

  const encontrados = aba
    .getRange(1, 1, 1, cabecalhosEsperados.length)
    .getValues()[0]
    .map(function(valor) {
      return String(valor || '').trim();
    });

  const divergencias = [];

  cabecalhosEsperados.forEach(function(cabecalho, indice) {
    if (encontrados[indice] !== cabecalho) {
      divergencias.push({
        coluna: indice + 1,
        esperado: cabecalho,
        encontrado: encontrados[indice]
      });
    }
  });

  if (divergencias.length > 0) {
    throw new Error(
      'Cabeçalhos divergentes em ' + nomeAba + ': ' +
      JSON.stringify(divergencias)
    );
  }

  return aba;
}


/**
 * Localiza veículo pelo código ou UUID.
 */
function DCP_DC236C_LOCALIZAR_VEICULO(aba, identificador) {
  const ultimaLinha = aba.getLastRow();

  if (ultimaLinha < 2) {
    throw new Error('Não existem veículos cadastrados.');
  }

  const dados = aba
    .getRange(2, 1, ultimaLinha - 1, 24)
    .getValues();

  const procurado = String(identificador || '').trim();

  const encontrados = [];

  dados.forEach(function(linha, indice) {
    const id = String(linha[0] || '').trim();
    const codigo = String(linha[1] || '').trim();

    if (id === procurado || codigo === procurado) {
      encontrados.push({
        linha: indice + 2,
        dados: linha
      });
    }
  });

  if (encontrados.length === 0) {
    throw new Error('Veículo não encontrado: ' + procurado);
  }

  if (encontrados.length > 1) {
    throw new Error('Identificador duplicado. Operação bloqueada.');
  }

  return encontrados[0];
}


/**
 * Diagnóstico somente de leitura.
 */
function DCP_DIAGNOSTICAR_REATIVACAO_VEI0001_236C() {
  const lock = LockService.getDocumentLock();

  if (!lock.tryLock(30000)) {
    throw new Error('Não foi possível obter o bloqueio da planilha.');
  }

  try {
    const abaVeiculos = DCP_DC236C_CONFERIR_ABA(
      DCP_DIAG_REATIVACAO_236C.ABA_VEICULOS,
      DCP_DIAG_REATIVACAO_236C.CABECALHOS_VEICULOS
    );

    const abaVinculos = DCP_DC236C_CONFERIR_ABA(
      DCP_DIAG_REATIVACAO_236C.ABA_VINCULOS,
      DCP_DIAG_REATIVACAO_236C.CABECALHOS_VINCULOS
    );

    const abaMotoristas = DCP_DC236C_CONFERIR_ABA(
      DCP_DIAG_REATIVACAO_236C.ABA_MOTORISTAS,
      DCP_DIAG_REATIVACAO_236C.CABECALHOS_MOTORISTAS
    );

    const localizado = DCP_DC236C_LOCALIZAR_VEICULO(
      abaVeiculos,
      'VEI-0001'
    );

    const v = localizado.dados;

    const veiculoID = String(v[0] || '').trim();
    const codigoVeiculo = String(v[1] || '').trim();
    const placa = String(v[2] || '').trim().toUpperCase();
    const status = String(v[19] || '').trim().toUpperCase();

    const placaValida =
      /^[A-Z]{3}[0-9]{4}$/.test(placa) ||
      /^[A-Z]{3}[0-9][A-Z][0-9]{2}$/.test(placa);

    const ultimaLinhaVeiculos = abaVeiculos.getLastRow();

    const todosVeiculos = ultimaLinhaVeiculos >= 2
      ? abaVeiculos.getRange(
          2, 1, ultimaLinhaVeiculos - 1, 24
        ).getValues()
      : [];

    const placasDuplicadas = todosVeiculos.filter(function(linha) {
      return (
        String(linha[2] || '').trim().toUpperCase() === placa &&
        String(linha[0] || '').trim() !== veiculoID
      );
    });

    const tipoProprietario = String(v[11] || '')
      .trim()
      .toUpperCase();

    const proprietarioMotoristaID = String(v[12] || '').trim();
    const proprietarioNome = String(v[13] || '').trim();

    let proprietarioValido = false;
    let dadosProprietario = null;

    if (tipoProprietario === 'TERCEIRO') {
      proprietarioValido = (
        proprietarioNome.length > 0 &&
        proprietarioNome.length <= 150 &&
        proprietarioMotoristaID === ''
      );

      dadosProprietario = {
        tipo: 'TERCEIRO',
        nome: proprietarioNome,
        valido: proprietarioValido
      };

    } else if (tipoProprietario === 'MOTORISTA') {
      const ultimaLinhaMotoristas = abaMotoristas.getLastRow();

      const motoristas = ultimaLinhaMotoristas >= 2
        ? abaMotoristas.getRange(
            2, 1, ultimaLinhaMotoristas - 1, 13
          ).getValues()
        : [];

      const encontrados = motoristas.filter(function(linha) {
        return String(linha[0] || '').trim() ===
          proprietarioMotoristaID;
      });

      if (encontrados.length === 1) {
        const motorista = encontrados[0];

        const statusMotorista = String(motorista[8] || '')
          .trim()
          .toUpperCase();

        const nomeMotorista = String(motorista[2] || '').trim();

        proprietarioValido = (
          statusMotorista === 'ATIVO' &&
          nomeMotorista.length >= 3 &&
          nomeMotorista === proprietarioNome
        );

        dadosProprietario = {
          tipo: 'MOTORISTA',
          motoristaID: proprietarioMotoristaID,
          nome: nomeMotorista,
          status: statusMotorista,
          valido: proprietarioValido
        };
      } else {
        dadosProprietario = {
          tipo: 'MOTORISTA',
          motoristaID: proprietarioMotoristaID,
          quantidadeEncontrada: encontrados.length,
          valido: false
        };
      }
    }

    const ultimaLinhaVinculos = abaVinculos.getLastRow();

    const vinculos = ultimaLinhaVinculos >= 2
      ? abaVinculos.getRange(
          2, 1, ultimaLinhaVinculos - 1, 16
        ).getValues()
      : [];

    const vinculosAtivos = vinculos.filter(function(linha) {
      return (
        String(linha[3] || '').trim() === veiculoID &&
        String(linha[11] || '').trim().toUpperCase() === 'ATIVO'
      );
    });

    const verificacoes = {
      statusInativo: status === 'INATIVO',
      placaValida: placaValida,
      placaSemDuplicidade: placasDuplicadas.length === 0,
      proprietarioValido: proprietarioValido,
      semVinculosAtivos: vinculosAtivos.length === 0
    };

    const podeSerReativado = Object.keys(verificacoes)
      .every(function(chave) {
        return verificacoes[chave] === true;
      });

    const resultado = {
      sucesso: true,
      codigoVeiculo: codigoVeiculo,
      veiculoID: veiculoID,
      placa: placa,
      marca: v[3],
      modelo: v[4],
      statusAtual: status,
      proprietario: dadosProprietario,
      quantidadeVinculosAtivos: vinculosAtivos.length,
      verificacoes: verificacoes,
      podeSerReativado: podeSerReativado,
      mensagem: podeSerReativado
        ? 'Veículo apto para reativação.'
        : 'Veículo não está apto. Verifique os critérios reprovados.'
    };

    Logger.log(JSON.stringify(resultado, null, 2));

    return resultado;

  } finally {
    lock.releaseLock();
  }
}

