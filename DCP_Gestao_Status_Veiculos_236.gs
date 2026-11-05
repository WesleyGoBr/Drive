
/**
 * DRIVER CONTROL PRO
 * Módulo 2.3.6
 * Diagnóstico de vínculos ativos de veículos
 *
 * Esta etapa é somente de leitura.
 * Nenhum dado será alterado.
 */

const DCP_STATUS_VEICULOS_236 = {
  ABA_VEICULOS: 'VEICULOS',
  ABA_VINCULOS: 'VINCULOS',

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
  ]
};


/**
 * Confere se a aba existe e se os cabeçalhos estão corretos.
 */
function DCP_SV236_CONFERIR_ABA(abaNome, cabecalhosEsperados) {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const aba = planilha.getSheetByName(abaNome);

  if (!aba) {
    throw new Error('A aba "' + abaNome + '" não foi encontrada.');
  }

  const quantidade = cabecalhosEsperados.length;
  const encontrados = aba
    .getRange(1, 1, 1, quantidade)
    .getDisplayValues()[0]
    .map(function(valor) {
      return String(valor || '').trim();
    });

  const divergencias = [];

  cabecalhosEsperados.forEach(function(esperado, indice) {
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
      'Estrutura incorreta na aba "' + abaNome + '":\n' +
      divergencias.join('\n')
    );
  }

  return aba;
}


/**
 * Retorna os vínculos ativos associados a um veículo.
 */
function DCP_SV236_BUSCAR_VINCULOS_ATIVOS(veiculoID) {
  const aba = DCP_SV236_CONFERIR_ABA(
    DCP_STATUS_VEICULOS_236.ABA_VINCULOS,
    DCP_STATUS_VEICULOS_236.CABECALHOS_VINCULOS
  );

  const ultimaLinha = aba.getLastRow();

  if (ultimaLinha < 2) {
    return [];
  }

  const dados = aba
    .getRange(2, 1, ultimaLinha - 1, 16)
    .getValues();

  const identificador = String(veiculoID || '').trim();

  return dados
    .filter(function(linha) {
      const idVeiculo = String(linha[3] || '').trim();
      const status = String(linha[11] || '').trim().toUpperCase();

      return idVeiculo === identificador && status === 'ATIVO';
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
 * Diagnóstico geral:
 * lista veículos e quantidade de vínculos ativos.
 */
function DCP_DIAGNOSTICAR_STATUS_VEICULOS_236() {
  try {
    const abaVeiculos = DCP_SV236_CONFERIR_ABA(
      DCP_STATUS_VEICULOS_236.ABA_VEICULOS,
      DCP_STATUS_VEICULOS_236.CABECALHOS_VEICULOS
    );

    DCP_SV236_CONFERIR_ABA(
      DCP_STATUS_VEICULOS_236.ABA_VINCULOS,
      DCP_STATUS_VEICULOS_236.CABECALHOS_VINCULOS
    );

    const ultimaLinha = abaVeiculos.getLastRow();

    if (ultimaLinha < 2) {
      const vazio = {
        sucesso: true,
        mensagem: 'Estrutura correta. Nenhum veículo cadastrado.',
        totalVeiculos: 0,
        veiculos: []
      };

      Logger.log(JSON.stringify(vazio, null, 2));
      return vazio;
    }

    const dados = abaVeiculos
      .getRange(2, 1, ultimaLinha - 1, 24)
      .getValues();

    const veiculos = dados
      .filter(function(linha) {
        return String(linha[0] || '').trim() !== '';
      })
      .map(function(linha) {
        const vinculos = DCP_SV236_BUSCAR_VINCULOS_ATIVOS(linha[0]);

        return {
          VeiculoID: linha[0],
          CodigoVeiculo: linha[1],
          Placa: linha[2],
          Marca: linha[3],
          Modelo: linha[4],
          Status: String(linha[19] || '').trim().toUpperCase(),
          QuantidadeVinculosAtivos: vinculos.length,
          PodeSerInativado: (
            String(linha[19] || '').trim().toUpperCase() === 'ATIVO' &&
            vinculos.length === 0
          )
        };
      });

    const resultado = {
      sucesso: true,
      mensagem: 'Diagnóstico concluído. Nenhum dado foi alterado.',
      totalVeiculos: veiculos.length,
      veiculos: veiculos
    };

    Logger.log(JSON.stringify(resultado, null, 2));
    return resultado;

  } catch (erro) {
    const resultado = {
      sucesso: false,
      mensagem: erro.message
    };

    Logger.log(JSON.stringify(resultado, null, 2));
    return resultado;
  }
}


/**
 * Diagnóstico individual por código ou UUID do veículo.
 */
function DCP_DIAGNOSTICAR_VEICULO_STATUS_236(identificador) {
  try {
    if (!identificador) {
      throw new Error('Informe o código ou UUID do veículo.');
    }

    const aba = DCP_SV236_CONFERIR_ABA(
      DCP_STATUS_VEICULOS_236.ABA_VEICULOS,
      DCP_STATUS_VEICULOS_236.CABECALHOS_VEICULOS
    );

    DCP_SV236_CONFERIR_ABA(
      DCP_STATUS_VEICULOS_236.ABA_VINCULOS,
      DCP_STATUS_VEICULOS_236.CABECALHOS_VINCULOS
    );

    const ultimaLinha = aba.getLastRow();

    if (ultimaLinha < 2) {
      throw new Error('Não existem veículos cadastrados.');
    }

    const dados = aba.getRange(2, 1, ultimaLinha - 1, 24).getValues();
    const chave = String(identificador).trim().toUpperCase();

    const encontrados = dados.filter(function(linha) {
      return (
        String(linha[0] || '').trim().toUpperCase() === chave ||
        String(linha[1] || '').trim().toUpperCase() === chave
      );
    });

    if (encontrados.length === 0) {
      throw new Error('Veículo não encontrado.');
    }

    if (encontrados.length > 1) {
      throw new Error('Mais de um veículo corresponde ao identificador.');
    }

    const linha = encontrados[0];
    const vinculos = DCP_SV236_BUSCAR_VINCULOS_ATIVOS(linha[0]);
    const status = String(linha[19] || '').trim().toUpperCase();

    const resultado = {
      sucesso: true,
      veiculo: {
        VeiculoID: linha[0],
        CodigoVeiculo: linha[1],
        Placa: linha[2],
        Marca: linha[3],
        Modelo: linha[4],
        Status: status
      },
      vinculosAtivos: vinculos,
      quantidadeVinculosAtivos: vinculos.length,
      podeSerInativado: status === 'ATIVO' && vinculos.length === 0,
      mensagem: vinculos.length > 0
        ? 'O veículo possui vínculo(s) ativo(s). Encerre-os antes de inativar.'
        : status !== 'ATIVO'
          ? 'O veículo não está ATIVO.'
          : 'O veículo não possui vínculos ativos e está apto para análise de inativação.'
    };

    Logger.log(JSON.stringify(resultado, null, 2));
    return resultado;

  } catch (erro) {
    const resultado = {
      sucesso: false,
      mensagem: erro.message
    };

    Logger.log(JSON.stringify(resultado, null, 2));
    return resultado;
  }
}
function DCP_TESTAR_DIAGNOSTICO_VEI0001_236() {
  const resultado = DCP_DIAGNOSTICAR_VEICULO_STATUS_236('VEI-0001');

  Logger.log(JSON.stringify(resultado, null, 2));

  return resultado;
}
