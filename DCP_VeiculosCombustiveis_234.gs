
const DCP_VEICULOS_COMBUSTIVEIS_234 = {
  ABA: 'VEICULOS_COMBUSTIVEIS',
  PREFIXO: 'VCO',
  CHAVE_SEQUENCIA: 'SEQ_VCO',
  ADMIN: 'ADMIN_CENTRAL',

  CABECALHOS: [
    'VeiculoCombustivelID',
    'CodigoVeiculoCombustivel',
    'VeiculoID',
    'CombustivelID',
    'InicioVigencia',
    'FimVigencia',
    'Observacoes',
    'Status',
    'CriadoEm',
    'AtualizadoEm',
    'CriadoPor',
    'AtualizadoPor'
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
  ]
};


/**
 * Cria a estrutura da tabela de compatibilidade.
 * Não cadastra veículos nem combustíveis.
 */
function DCP_CRIAR_ESTRUTURA_VEICULOS_COMBUSTIVEIS_234() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const existentes = [
    'VEICULOS',
    'COMBUSTIVEIS',
    'CONFIG',
    'AUDITORIA'
  ];

  existentes.forEach(nome => {
    if (!ss.getSheetByName(nome)) {
      throw new Error(
        'Aba obrigatória não encontrada: ' + nome
      );
    }
  });

  const nome = DCP_VEICULOS_COMBUSTIVEIS_234.ABA;
  let aba = ss.getSheetByName(nome);

  if (!aba) {
    aba = ss.insertSheet(nome);
  }

  const cabecalhos =
    DCP_VEICULOS_COMBUSTIVEIS_234.CABECALHOS;

  const ultimaColuna = aba.getLastColumn();

  if (ultimaColuna > 0) {
    const atuais = aba
      .getRange(1, 1, 1, Math.max(
        ultimaColuna,
        cabecalhos.length
      ))
      .getValues()[0];

    const possuiDados =
      aba.getLastRow() > 1;

    if (possuiDados) {
      throw new Error(
        'A aba já possui registros. ' +
        'A criação foi interrompida para preservar os dados.'
      );
    }

    const preenchidos = atuais.filter(
      valor => valor !== ''
    );

    if (
      preenchidos.length > 0 &&
      JSON.stringify(preenchidos) !==
      JSON.stringify(cabecalhos)
    ) {
      throw new Error(
        'A aba já possui cabeçalhos diferentes. ' +
        'Nenhuma alteração foi realizada.'
      );
    }
  }

  aba.getRange(
    1,
    1,
    1,
    cabecalhos.length
  ).setValues([cabecalhos]);

  aba.setFrozenRows(1);

  aba.getRange(
    1,
    1,
    1,
    cabecalhos.length
  ).setFontWeight('bold');

  aba.autoResizeColumns(
    1,
    cabecalhos.length
  );

  SpreadsheetApp.flush();

  Logger.log(
    'ESTRUTURA VEICULOS_COMBUSTIVEIS CRIADA/CONFERIDA.'
  );
  Logger.log('Total de colunas: ' + cabecalhos.length);
  Logger.log('Código visual: VCO-0001 em diante.');
}


/**
 * Confere os cabeçalhos das tabelas relacionadas.
 * Não modifica dados.
 */
function DCP_DIAGNOSTICAR_ESTRUTURA_VEICULOS_COMBUSTIVEIS_234() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const estruturas = [
    {
      nome: 'VEICULOS_COMBUSTIVEIS',
      cabecalhos:
        DCP_VEICULOS_COMBUSTIVEIS_234.CABECALHOS
    },
    {
      nome: 'VEICULOS',
      cabecalhos: [
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
    },
    {
      nome: 'COMBUSTIVEIS',
      cabecalhos: [
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
      ]
    },
    {
      nome: 'CONFIG',
      cabecalhos:
        DCP_VEICULOS_COMBUSTIVEIS_234.CABECALHOS_CONFIG
    },
    {
      nome: 'AUDITORIA',
      cabecalhos:
        DCP_VEICULOS_COMBUSTIVEIS_234.CABECALHOS_AUDITORIA
    }
  ];

  let falhas = 0;

  estruturas.forEach(item => {
    const aba = ss.getSheetByName(item.nome);

    if (!aba) {
      Logger.log('ERRO: Aba ausente: ' + item.nome);
      falhas++;
      return;
    }

    const atuais = aba.getRange(
      1,
      1,
      1,
      item.cabecalhos.length
    ).getValues()[0];

    const iguais = item.cabecalhos.every(
      (campo, indice) => atuais[indice] === campo
    );

    if (iguais) {
      Logger.log('OK: ' + item.nome);
    } else {
      Logger.log('ERRO: Cabeçalhos divergentes: ' + item.nome);
      falhas++;
    }
  });

  Logger.log('TOTAL DE FALHAS: ' + falhas);

  Logger.log(
    falhas === 0
      ? 'RESULTADO: ESTRUTURA APROVADA.'
      : 'RESULTADO: ESTRUTURA COM PENDÊNCIAS.'
  );
}

/**
 * MODULO 2.3.4A
 * Validacao de compatibilidade entre veiculo e combustivel.
 * Somente leitura: nao grava dados.
 */

function DCP_VALIDAR_COMPATIBILIDADE_234(dados) {
  const erros = [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!dados || typeof dados !== 'object') {
    return {
      valido: false,
      erros: ['Dados de entrada invalidos.']
    };
  }

  const veiculoID = String(dados.veiculoID || '').trim();
  const combustivelID = String(dados.combustivelID || '').trim();
  const observacoes = String(dados.observacoes || '').trim();

  if (!veiculoID) erros.push('Veiculo obrigatorio.');
  if (!combustivelID) erros.push('Combustivel obrigatorio.');

  if (observacoes.length > 2000) {
    erros.push('Observacoes excedem 2000 caracteres.');
  }

  const inicio = DCP_CONVERTER_DATA_COMPATIBILIDADE_234(
    dados.inicioVigencia
  );

  if (!inicio) {
    erros.push('Data de inicio invalida ou nao informada.');
  }

  if (erros.length > 0) {
    return { valido: false, erros: erros };
  }

  const abaVeiculos = ss.getSheetByName('VEICULOS');
  const abaCombustiveis = ss.getSheetByName('COMBUSTIVEIS');
  const abaVinculos = ss.getSheetByName('VEICULOS_COMBUSTIVEIS');

  if (!abaVeiculos || !abaCombustiveis || !abaVinculos) {
    return {
      valido: false,
      erros: ['Uma ou mais abas obrigatorias nao foram encontradas.']
    };
  }

  // Confere se os identificadores existem e se os registros estao ativos.
  const veiculos = abaVeiculos.getDataRange().getValues();
  const combustiveis = abaCombustiveis.getDataRange().getValues();

  const indiceVeiculo = veiculos.findIndex(
    (linha, i) => i > 0 &&
      String(linha[0]).trim() === veiculoID
  );

  if (indiceVeiculo < 0) {
    erros.push('Veiculo nao encontrado.');
  } else if (
    String(veiculos[indiceVeiculo][19]).trim() !== 'ATIVO'
  ) {
    erros.push('Veiculo nao esta ativo.');
  }

  const indiceCombustivel = combustiveis.findIndex(
    (linha, i) => i > 0 &&
      String(linha[0]).trim() === combustivelID
  );

  if (indiceCombustivel < 0) {
    erros.push('Combustivel nao encontrado.');
  } else if (
    String(combustiveis[indiceCombustivel][6]).trim() !== 'ATIVO'
  ) {
    erros.push('Combustivel nao esta ativo.');
  }

  if (erros.length > 0) {
    return { valido: false, erros: erros };
  }

  // Impede dois vinculos ativos para o mesmo par.
  const vinculos = abaVinculos.getDataRange().getValues();

  const duplicado = vinculos.some((linha, i) =>
    i > 0 &&
    String(linha[2]).trim() === veiculoID &&
    String(linha[3]).trim() === combustivelID &&
    String(linha[7]).trim() === 'ATIVO'
  );

  if (duplicado) {
    erros.push(
      'Este veiculo ja possui uma compatibilidade ativa com esse combustivel.'
    );
  }

  return {
    valido: erros.length === 0,
    erros: erros
  };
}


/**
 * Converte data recebida em objeto Date valido.
 */
function DCP_CONVERTER_DATA_COMPATIBILIDADE_234(valor) {
  if (valor instanceof Date && !isNaN(valor.getTime())) {
    return valor;
  }

  if (typeof valor !== 'string' || !valor.trim()) {
    return null;
  }

  const texto = valor.trim();
  const partes = texto.match(/^(\d{4})-(\d{2})-(\d{2})$/);

  if (!partes) return null;

  const ano = Number(partes[1]);
  const mes = Number(partes[2]);
  const dia = Number(partes[3]);

  const data = new Date(ano, mes - 1, dia);

  if (
    data.getFullYear() !== ano ||
    data.getMonth() !== mes - 1 ||
    data.getDate() !== dia
  ) {
    return null;
  }

  data.setHours(0, 0, 0, 0);
  return data;
}

function DCP_TESTAR_VALIDACOES_COMPATIBILIDADE_234() {
  const testes = [
    {
      nome: 'Dados vazios rejeitados',
      dados: {},
      esperado: false
    },
    {
      nome: 'Veiculo inexistente rejeitado',
      dados: {
        veiculoID: 'VEICULO_INEXISTENTE',
        combustivelID: 'feb4ced9-05be-481b-9545-7e452da66afe',
        inicioVigencia: '2026-11-03'
      },
      esperado: false
    },
    {
      nome: 'Combustivel inexistente rejeitado',
      dados: {
        veiculoID: 'VEICULO_INEXISTENTE',
        combustivelID: 'COMBUSTIVEL_INEXISTENTE',
        inicioVigencia: '2026-11-03'
      },
      esperado: false
    },
    {
      nome: 'Data invalida rejeitada',
      dados: {
        veiculoID: 'VEICULO_INEXISTENTE',
        combustivelID: 'feb4ced9-05be-481b-9545-7e452da66afe',
        inicioVigencia: '2026-02-30'
      },
      esperado: false
    },
    {
      nome: 'Data valida reconhecida',
      dados: {
        veiculoID: 'VEICULO_INEXISTENTE',
        combustivelID: 'feb4ced9-05be-481b-9545-7e452da66afe',
        inicioVigencia: '2026-11-03'
      },
      esperado: false
    }
  ];

  let aprovados = 0;
  let falhas = 0;

  testes.forEach(teste => {
    try {
      const resultado =
        DCP_VALIDAR_COMPATIBILIDADE_234(teste.dados);

      const passou =
        resultado.valido === teste.esperado;

      if (passou) {
        Logger.log('OK: ' + teste.nome);
        aprovados++;
      } else {
        Logger.log(
          'FALHA: ' + teste.nome +
          ' | Resultado: ' + JSON.stringify(resultado)
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

function DCP_DIAGNOSTICAR_REGISTROS_COMPATIBILIDADE_234() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const abaVeiculos = ss.getSheetByName('VEICULOS');
  const abaCombustiveis = ss.getSheetByName('COMBUSTIVEIS');
  const abaVinculos = ss.getSheetByName('VEICULOS_COMBUSTIVEIS');

  if (!abaVeiculos || !abaCombustiveis || !abaVinculos) {
    throw new Error('Uma ou mais abas obrigatorias nao existem.');
  }

  const veiculos = abaVeiculos.getDataRange().getValues();
  const combustiveis = abaCombustiveis.getDataRange().getValues();
  const vinculos = abaVinculos.getDataRange().getValues();

  const veiculosAtivos = veiculos.slice(1).filter(
    linha => String(linha[19]).trim() === 'ATIVO'
  );

  const combustiveisAtivos = combustiveis.slice(1).filter(
    linha => String(linha[6]).trim() === 'ATIVO'
  );

  const veiculosInativos = veiculos.slice(1).filter(
    linha => String(linha[19]).trim() === 'INATIVO'
  );

  const combustiveisInativos = combustiveis.slice(1).filter(
    linha => String(linha[6]).trim() === 'INATIVO'
  );

  const vinculosAtivos = vinculos.slice(1).filter(
    linha => String(linha[7]).trim() === 'ATIVO'
  );

  Logger.log('===== DIAGNOSTICO DE COMPATIBILIDADE =====');

  Logger.log('VEICULOS ATIVOS: ' + veiculosAtivos.length);
  Logger.log('VEICULOS INATIVOS: ' + veiculosInativos.length);

  veiculosAtivos.forEach(linha => {
    Logger.log(
      'VEICULO: ' + linha[1] +
      ' | PLACA: ' + linha[2] +
      ' | MODELO: ' + linha[4] +
      ' | UUID: ' + linha[0]
    );
  });

  Logger.log('COMBUSTIVEIS ATIVOS: ' + combustiveisAtivos.length);
  Logger.log('COMBUSTIVEIS INATIVOS: ' + combustiveisInativos.length);

  combustiveisAtivos.forEach(linha => {
    Logger.log(
      'COMBUSTIVEL: ' + linha[1] +
      ' | NOME: ' + linha[2] +
      ' | UUID: ' + linha[0]
    );
  });

  Logger.log('COMPATIBILIDADES ATIVAS: ' + vinculosAtivos.length);

  Logger.log(
    'TOTAL DE REGISTROS NA TABELA DE COMPATIBILIDADE: ' +
    Math.max(0, vinculos.length - 1)
  );

  Logger.log('===== FIM DO DIAGNOSTICO =====');
  Logger.log('DIAGNOSTICO CONCLUIDO. NENHUM DADO ALTERADO.');
}

function DCP_DIAGNOSTICAR_VEICULOS_235A() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const esperados = [
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

  const obrigatorias = [
    'VEICULOS',
    'MOTORISTAS',
    'CONFIG',
    'AUDITORIA',
    'COMBUSTIVEIS'
  ];

  let falhas = 0;

  obrigatorias.forEach(nome => {
    if (ss.getSheetByName(nome)) {
      Logger.log('OK: Aba ' + nome);
    } else {
      Logger.log('ERRO: Aba ausente: ' + nome);
      falhas++;
    }
  });

  const aba = ss.getSheetByName('VEICULOS');

  if (!aba) {
    Logger.log('Diagnostico interrompido: VEICULOS ausente.');
    return;
  }

  const cabecalhos = aba.getRange(
    1, 1, 1, esperados.length
  ).getValues()[0];

  esperados.forEach((campo, indice) => {
    if (cabecalhos[indice] === campo) {
      Logger.log('OK: Coluna ' + (indice + 1) + ' - ' + campo);
    } else {
      Logger.log(
        'ERRO: Coluna ' + (indice + 1) +
        ' | Esperado: ' + campo +
        ' | Encontrado: ' + cabecalhos[indice]
      );
      falhas++;
    }
  });

  const ultimaLinha = aba.getLastRow();
  const totalRegistros = Math.max(0, ultimaLinha - 1);

  let ativos = 0;
  let inativos = 0;
  let outrosStatus = 0;

  if (totalRegistros > 0) {
    const registros = aba.getRange(
      2, 1, totalRegistros, esperados.length
    ).getValues();

    registros.forEach(linha => {
      if (!linha.some(valor => valor !== '')) return;

      const status = String(linha[19]).trim();

      if (status === 'ATIVO') ativos++;
      else if (status === 'INATIVO') inativos++;
      else outrosStatus++;
    });
  }

  Logger.log('===== RESUMO VEICULOS =====');
  Logger.log('Registros com dados: ' + totalRegistros);
  Logger.log('Ativos: ' + ativos);
  Logger.log('Inativos: ' + inativos);
  Logger.log('Outros status: ' + outrosStatus);
  Logger.log('TOTAL DE FALHAS: ' + falhas);

  Logger.log(
    falhas === 0
      ? 'RESULTADO: ESTRUTURA APROVADA.'
      : 'RESULTADO: ESTRUTURA COM PENDENCIAS.'
  );

  Logger.log('Nenhum dado foi alterado.');
}

