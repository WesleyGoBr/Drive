/************************************************************
 * DRIVER CONTROL PRO
 * MODULO 2.3.5D - CONSULTA E LISTAGEM DE VEICULOS
 *
 * Versao: 1.0
 * Operacoes: SOMENTE LEITURA
 ************************************************************/

const DCP_GESTAO_VEICULOS_235D = {
  ABA: 'VEICULOS',

  CABECALHOS: [
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
};


/**
 * Confere a estrutura da aba VEICULOS.
 * Nao altera nenhum dado.
 */
function DCP_CONFERIR_ESTRUTURA_GESTAO_VEICULOS_235D() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(
    DCP_GESTAO_VEICULOS_235D.ABA
  );

  if (!aba) {
    throw new Error('A aba VEICULOS nao foi encontrada.');
  }

  const esperados = DCP_GESTAO_VEICULOS_235D.CABECALHOS;
  const encontrados = aba
    .getRange(1, 1, 1, esperados.length)
    .getDisplayValues()[0];

  const divergencias = [];

  esperados.forEach(function(cabecalho, indice) {
    if (String(encontrados[indice]).trim() !== cabecalho) {
      divergencias.push({
        coluna: indice + 1,
        esperado: cabecalho,
        encontrado: encontrados[indice]
      });
    }
  });

  const resultado = {
    sucesso: divergencias.length === 0,
    aba: DCP_GESTAO_VEICULOS_235D.ABA,
    quantidadeColunas: esperados.length,
    divergencias: divergencias
  };

  Logger.log(JSON.stringify(resultado, null, 2));

  if (!resultado.sucesso) {
    throw new Error(
      'Estrutura da aba VEICULOS divergente. Consulte o log.'
    );
  }

  return resultado;
}


/**
 * Converte datas para texto antes de devolver os dados.
 * Isso evita problemas de serializacao em formularios HTML.
 */
function DCP_GV235D_SERIALIZAR_VALOR(valor) {
  if (valor instanceof Date) {
    return Utilities.formatDate(
      valor,
      Session.getScriptTimeZone(),
      'yyyy-MM-dd HH:mm:ss'
    );
  }

  return valor;
}


/**
 * Converte uma linha da planilha em objeto.
 */
function DCP_GV235D_CONVERTER_REGISTRO(linha) {
  const campos = DCP_GESTAO_VEICULOS_235D.CABECALHOS;
  const registro = {};

  campos.forEach(function(campo, indice) {
    registro[campo] = DCP_GV235D_SERIALIZAR_VALOR(
      linha[indice]
    );
  });

  registro.Status = String(registro.Status || '')
    .trim()
    .toUpperCase();

  return registro;
}


/**
 * Lista veiculos.
 *
 * filtroStatus:
 * ATIVO    = somente ativos
 * INATIVO  = somente inativos
 * TODOS    = todos
 *
 * pesquisa:
 * Busca por codigo, placa, marca, modelo ou proprietario.
 */
function DCP_LISTAR_VEICULOS_235D(filtroStatus, pesquisa) {
  DCP_CONFERIR_ESTRUTURA_GESTAO_VEICULOS_235D();

  const status = String(filtroStatus || 'ATIVO')
    .trim()
    .toUpperCase();

  if (!['ATIVO', 'INATIVO', 'TODOS'].includes(status)) {
    throw new Error(
      'Filtro de status invalido. Use ATIVO, INATIVO ou TODOS.'
    );
  }

  const termo = String(pesquisa || '')
    .trim()
    .toUpperCase();

  if (termo.length > 100) {
    throw new Error('A pesquisa deve ter no maximo 100 caracteres.');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(
    DCP_GESTAO_VEICULOS_235D.ABA
  );

  const ultimaLinha = aba.getLastRow();

  if (ultimaLinha < 2) {
    return {
      sucesso: true,
      quantidade: 0,
      filtroStatus: status,
      pesquisa: termo,
      veiculos: []
    };
  }

  const dados = aba
    .getRange(
      2,
      1,
      ultimaLinha - 1,
      DCP_GESTAO_VEICULOS_235D.CABECALHOS.length
    )
    .getValues();

  const veiculos = [];

  dados.forEach(function(linha) {
    if (!linha[0]) {
      return;
    }

    const registro = DCP_GV235D_CONVERTER_REGISTRO(linha);

    if (
      status !== 'TODOS' &&
      registro.Status !== status
    ) {
      return;
    }

    if (termo) {
      const camposPesquisa = [
        registro.CodigoVeiculo,
        registro.Placa,
        registro.Marca,
        registro.Modelo,
        registro.ProprietarioNome
      ];

      const corresponde = camposPesquisa.some(function(valor) {
        return String(valor || '')
          .toUpperCase()
          .includes(termo);
      });

      if (!corresponde) {
        return;
      }
    }

    veiculos.push(registro);
  });

  veiculos.sort(function(a, b) {
    return String(a.CodigoVeiculo)
      .localeCompare(String(b.CodigoVeiculo));
  });

  return {
    sucesso: true,
    quantidade: veiculos.length,
    filtroStatus: status,
    pesquisa: termo,
    veiculos: veiculos
  };
}


/**
 * Consulta um veiculo por UUID ou codigo.
 * Exemplo: VEI-0001
 */
function DCP_OBTER_VEICULO_235D(identificador) {
  const chave = String(identificador || '').trim();

  if (!chave) {
    throw new Error('Informe o UUID ou codigo do veiculo.');
  }

  DCP_CONFERIR_ESTRUTURA_GESTAO_VEICULOS_235D();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(
    DCP_GESTAO_VEICULOS_235D.ABA
  );

  const ultimaLinha = aba.getLastRow();

  if (ultimaLinha < 2) {
    return {
      sucesso: true,
      encontrado: false,
      veiculo: null
    };
  }

  const dados = aba
    .getRange(
      2,
      1,
      ultimaLinha - 1,
      DCP_GESTAO_VEICULOS_235D.CABECALHOS.length
    )
    .getValues();

  const encontrados = dados.filter(function(linha) {
    return (
      String(linha[0]).trim() === chave ||
      String(linha[1]).trim().toUpperCase() === chave.toUpperCase()
    );
  });

  if (encontrados.length > 1) {
    throw new Error(
      'Integridade comprometida: identificador encontrado mais de uma vez.'
    );
  }

  if (encontrados.length === 0) {
    return {
      sucesso: true,
      encontrado: false,
      veiculo: null
    };
  }

  return {
    sucesso: true,
    encontrado: true,
    veiculo: DCP_GV235D_CONVERTER_REGISTRO(encontrados[0])
  };
}


/**
 * Teste de consulta do modulo.
 * Nao altera dados.
 */
function DCP_TESTAR_CONSULTA_VEICULOS_235D() {
  Logger.log('===== TESTE CONSULTA DE VEICULOS =====');

  const estrutura =
    DCP_CONFERIR_ESTRUTURA_GESTAO_VEICULOS_235D();

  Logger.log(
    (estrutura.sucesso ? 'OK' : 'FALHA') +
    ': Estrutura da aba.'
  );

  const ativos = DCP_LISTAR_VEICULOS_235D('ATIVO');
  Logger.log('Veiculos ativos: ' + ativos.quantidade);

  const inativos = DCP_LISTAR_VEICULOS_235D('INATIVO');
  Logger.log('Veiculos inativos: ' + inativos.quantidade);

  const todos = DCP_LISTAR_VEICULOS_235D('TODOS');
  Logger.log('Total de veiculos: ' + todos.quantidade);

  const consulta = DCP_OBTER_VEICULO_235D('VEI-0001');

  Logger.log(
    (consulta.encontrado ? 'OK' : 'ATENCAO') +
    ': Consulta individual do VEI-0001.'
  );

  if (consulta.encontrado) {
    Logger.log('CODIGO: ' + consulta.veiculo.CodigoVeiculo);
    Logger.log('PLACA: ' + consulta.veiculo.Placa);
    Logger.log('MARCA: ' + consulta.veiculo.Marca);
    Logger.log('MODELO: ' + consulta.veiculo.Modelo);
    Logger.log('STATUS: ' + consulta.veiculo.Status);
  }

  Logger.log('===== FIM DO TESTE =====');
  Logger.log('OPERACAO SOMENTE DE LEITURA.');
}