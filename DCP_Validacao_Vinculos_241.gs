
/**
 * DRIVER CONTROL PRO
 * Módulo 2.4.1 - Validação de períodos e conflitos de vínculos
 *
 * Somente validação: não grava nem altera registros.
 */

const DCP_VALIDACAO_VINCULOS_241 = {
  ABA_VINCULOS: 'VINCULOS',
  CABECALHOS: [
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
 * Converte valores de data/hora para Date.
 * Strings no formato YYYY-MM-DDTHH:mm são interpretadas
 * como horário local da planilha/script.
 */
function DCP_V241_CONVERTER_DATA(valor, campo) {
  if (valor === null || valor === undefined || valor === '') {
    return {
      valida: true,
      valor: null
    };
  }

  if (valor instanceof Date && !isNaN(valor.getTime())) {
    return {
      valida: true,
      valor: new Date(valor.getTime())
    };
  }

  if (typeof valor === 'string') {
    const texto = valor.trim();

    const correspondencia = texto.match(
      /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})(?::(\d{2}))?$/
    );

    if (correspondencia) {
      const ano = Number(correspondencia[1]);
      const mes = Number(correspondencia[2]);
      const dia = Number(correspondencia[3]);
      const hora = Number(correspondencia[4]);
      const minuto = Number(correspondencia[5]);
      const segundo = Number(correspondencia[6] || 0);

      const data = new Date(
        ano,
        mes - 1,
        dia,
        hora,
        minuto,
        segundo
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
    mensagem: 'Data/hora inválida no campo ' + campo + '.'
  };
}


/**
 * Confere os cabeçalhos da aba VINCULOS.
 */
function DCP_V241_CONFERIR_ESTRUTURA() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const aba = planilha.getSheetByName(
    DCP_VALIDACAO_VINCULOS_241.ABA_VINCULOS
  );

  if (!aba) {
    throw new Error('A aba VINCULOS não foi encontrada.');
  }

  const esperados = DCP_VALIDACAO_VINCULOS_241.CABECALHOS;
  const encontrados = aba
    .getRange(1, 1, 1, esperados.length)
    .getValues()[0];

  const divergencias = [];

  esperados.forEach(function(cabecalho, indice) {
    if (String(encontrados[indice] || '').trim() !== cabecalho) {
      divergencias.push({
        coluna: indice + 1,
        esperado: cabecalho,
        encontrado: encontrados[indice] || ''
      });
    }
  });

  return {
    sucesso: divergencias.length === 0,
    aba: aba.getName(),
    quantidadeColunas: esperados.length,
    divergencias: divergencias
  };
}


/**
 * Valida os dados básicos do período de um vínculo.
 */
function DCP_VALIDAR_PERIODO_VINCULO_241(dados) {
  if (!dados || typeof dados !== 'object') {
    return {
      valido: false,
      erros: ['Dados do vínculo não informados.']
    };
  }

  const erros = [];

  const motoristaID = String(dados.MotoristaID || '').trim();
  const veiculoID = String(dados.VeiculoID || '').trim();

  const inicio = DCP_V241_CONVERTER_DATA(
    dados.InicioEm,
    'InicioEm'
  );

  const fim = DCP_V241_CONVERTER_DATA(
    dados.FimEm,
    'FimEm'
  );

  if (!motoristaID) {
    erros.push('MotoristaID é obrigatório.');
  }

  if (!veiculoID) {
    erros.push('VeiculoID é obrigatório.');
  }

  if (!dados.InicioEm) {
    erros.push('Data e horário de início são obrigatórios.');
  } else if (!inicio.valida) {
    erros.push(inicio.mensagem);
  }

  if (!fim.valida) {
    erros.push(fim.mensagem);
  }

  if (
    inicio.valida &&
    fim.valida &&
    inicio.valor &&
    fim.valor &&
    fim.valor.getTime() <= inicio.valor.getTime()
  ) {
    erros.push(
      'O término deve ser posterior ao início do vínculo.'
    );
  }

  const autorizacao = dados.CompartilhamentoAutorizado;

  const autorizacaoValida =
    autorizacao === true ||
    autorizacao === false ||
    String(autorizacao).trim().toUpperCase() === 'TRUE' ||
    String(autorizacao).trim().toUpperCase() === 'FALSE' ||
    autorizacao === 1 ||
    autorizacao === 0;

  if (!autorizacaoValida) {
    erros.push(
      'CompartilhamentoAutorizado deve ser informado explicitamente como verdadeiro ou falso.'
    );
  }

  return {
    valido: erros.length === 0,
    erros: erros,
    periodo: {
      inicio: inicio.valor,
      fim: fim.valor
    }
  };
}


/**
 * Verifica se dois períodos se sobrepõem.
 *
 * O término é exclusivo:
 * 06:00 até 14:00 e 14:00 até 22:00 são permitidos.
 *
 * FimEm vazio representa período sem término definido.
 */
function DCP_V241_PERIODOS_SOBREPOEM(
  inicioA,
  fimA,
  inicioB,
  fimB
) {
  const limiteA = fimA
    ? fimA.getTime()
    : Infinity;

  const limiteB = fimB
    ? fimB.getTime()
    : Infinity;

  return (
    inicioA.getTime() < limiteB &&
    inicioB.getTime() < limiteA
  );
}


/**
 * Verifica conflitos com vínculos já cadastrados.
 *
 * Regras:
 * 1. Motorista não pode ter períodos sobrepostos.
 * 2. Veículo não pode ter períodos sobrepostos.
 * 3. Para um veículo ter vários vínculos ativos,
 *    o vínculo novo e os existentes precisam ter
 *    autorização explícita de compartilhamento.
 *
 * identificadorIgnorado permite futuramente editar
 * um vínculo sem gerar conflito consigo mesmo.
 */
function DCP_VERIFICAR_CONFLITOS_VINCULO_241(dados) {
  const validacao = DCP_VALIDAR_PERIODO_VINCULO_241(dados);

  if (!validacao.valido) {
    return {
      sucesso: false,
      semConflitos: false,
      erros: validacao.erros,
      conflitos: []
    };
  }

  const estrutura = DCP_V241_CONFERIR_ESTRUTURA();

  if (!estrutura.sucesso) {
    return {
      sucesso: false,
      semConflitos: false,
      erros: ['Estrutura da aba VINCULOS divergente.'],
      divergencias: estrutura.divergencias,
      conflitos: []
    };
  }

  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const aba = planilha.getSheetByName(
    DCP_VALIDACAO_VINCULOS_241.ABA_VINCULOS
  );

  const ultimaLinha = aba.getLastRow();

  if (ultimaLinha < 2) {
    return {
      sucesso: true,
      semConflitos: true,
      conflitos: [],
      mensagem: 'Nenhum vínculo existente para gerar conflito.'
    };
  }

  const registros = aba.getRange(
    2,
    1,
    ultimaLinha - 1,
    DCP_VALIDACAO_VINCULOS_241.CABECALHOS.length
  ).getValues();

  const motoristaID = String(dados.MotoristaID).trim();
  const veiculoID = String(dados.VeiculoID).trim();

  const identificadorIgnorado = String(
    dados.VinculoIDIgnorado || ''
  ).trim();

  const inicioNovo = validacao.periodo.inicio;
  const fimNovo = validacao.periodo.fim;

  const autorizacaoNova =
    dados.CompartilhamentoAutorizado === true ||
    String(dados.CompartilhamentoAutorizado)
      .trim().toUpperCase() === 'TRUE' ||
    dados.CompartilhamentoAutorizado === 1;

  const conflitos = [];

  registros.forEach(function(linha, indice) {
    const vinculoID = String(linha[0] || '').trim();
    const codigoVinculo = String(linha[1] || '').trim();
    const motoristaExistente = String(linha[2] || '').trim();
    const veiculoExistente = String(linha[3] || '').trim();
    const status = String(linha[11] || '')
      .trim()
      .toUpperCase();

    if (status !== 'ATIVO') {
      return;
    }

    if (
      identificadorIgnorado &&
      (
        vinculoID === identificadorIgnorado ||
        codigoVinculo === identificadorIgnorado
      )
    ) {
      return;
    }

    const inicioExistente = DCP_V241_CONVERTER_DATA(
      linha[5],
      'InicioEm do vínculo ' + codigoVinculo
    );

    const fimExistente = DCP_V241_CONVERTER_DATA(
      linha[6],
      'FimEm do vínculo ' + codigoVinculo
    );

    if (
      !inicioExistente.valida ||
      !fimExistente.valida ||
      !inicioExistente.valor
    ) {
      conflitos.push({
        tipo: 'REGISTRO_EXISTENTE_INVALIDO',
        codigoVinculo: codigoVinculo,
        linha: indice + 2,
        mensagem: 'Vínculo existente possui período inválido.'
      });
      return;
    }

    const sobreposicao = DCP_V241_PERIODOS_SOBREPOEM(
      inicioNovo,
      fimNovo,
      inicioExistente.valor,
      fimExistente.valor
    );

    if (!sobreposicao) {
      // Períodos sequenciais são permitidos.
      return;
    }

    if (motoristaExistente === motoristaID) {
      conflitos.push({
        tipo: 'CONFLITO_MOTORISTA',
        codigoVinculo: codigoVinculo,
        linha: indice + 2,
        mensagem:
          'O motorista já possui outro vínculo ativo nesse período.'
      });
    }

    if (veiculoExistente === veiculoID) {
      conflitos.push({
        tipo: 'CONFLITO_VEICULO',
        codigoVinculo: codigoVinculo,
        linha: indice + 2,
        mensagem:
          'O veículo já possui outro vínculo ativo nesse período.'
      });
    }
  });

  return {
    sucesso: true,
    semConflitos: conflitos.length === 0,
    conflitos: conflitos,
    mensagem: conflitos.length === 0
      ? 'Nenhum conflito de período encontrado.'
      : 'Foram encontrados conflitos que impedem o vínculo.'
  };
}


/**
 * Testes isolados da lógica de períodos.
 * Não acessam nem alteram a planilha.
 */
function DCP_TESTAR_PERIODOS_VINCULOS_241() {
  const casos = [
    {
      nome: 'Períodos sequenciais',
      inicioA: new Date(2026, 10, 10, 6, 0),
      fimA: new Date(2026, 10, 10, 14, 0),
      inicioB: new Date(2026, 10, 10, 14, 0),
      fimB: new Date(2026, 10, 10, 22, 0),
      esperado: false
    },
    {
      nome: 'Sobreposição parcial',
      inicioA: new Date(2026, 10, 10, 6, 0),
      fimA: new Date(2026, 10, 10, 14, 0),
      inicioB: new Date(2026, 10, 10, 13, 0),
      fimB: new Date(2026, 10, 10, 22, 0),
      esperado: true
    },
    {
      nome: 'Sobreposição total',
      inicioA: new Date(2026, 10, 10, 6, 0),
      fimA: new Date(2026, 10, 10, 22, 0),
      inicioB: new Date(2026, 10, 10, 8, 0),
      fimB: new Date(2026, 10, 10, 12, 0),
      esperado: true
    },
    {
      nome: 'Vínculo sem término',
      inicioA: new Date(2026, 10, 10, 6, 0),
      fimA: null,
      inicioB: new Date(2026, 10, 11, 8, 0),
      fimB: new Date(2026, 10, 11, 12, 0),
      esperado: true
    },
    {
      nome: 'Períodos em dias diferentes',
      inicioA: new Date(2026, 10, 10, 6, 0),
      fimA: new Date(2026, 10, 10, 22, 0),
      inicioB: new Date(2026, 10, 11, 6, 0),
      fimB: new Date(2026, 10, 11, 22, 0),
      esperado: false
    }
  ];

  const resultados = casos.map(function(caso) {
    const obtido = DCP_V241_PERIODOS_SOBREPOEM(
      caso.inicioA,
      caso.fimA,
      caso.inicioB,
      caso.fimB
    );

    return {
      teste: caso.nome,
      esperado: caso.esperado,
      obtido: obtido,
      aprovado: obtido === caso.esperado
    };
  });

  const aprovados = resultados.filter(function(item) {
    return item.aprovado;
  }).length;

  const resultado = {
    total: resultados.length,
    aprovados: aprovados,
    reprovados: resultados.length - aprovados,
    resultados: resultados
  };

  Logger.log(JSON.stringify(resultado, null, 2));
  return resultado;
}


/**
 * Testa a estrutura da aba VINCULOS.
 */
function DCP_TESTAR_ESTRUTURA_VINCULOS_241() {
  const resultado = DCP_V241_CONFERIR_ESTRUTURA();

  Logger.log(JSON.stringify(resultado, null, 2));
  return resultado;
}
