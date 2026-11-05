
/**
 * DRIVER CONTROL PRO
 * MODULO 2.4.6
 * TESTE DE VALIDACOES DAS LISTAS SUSPENSAS
 *
 * Arquivo independente.
 *
 * OBJETIVO:
 * Verificar se as celulas que receberiam o proximo vinculo
 * e o proximo evento de auditoria aceitam os valores usados
 * pelo modulo DCP_Cadastro_Vinculos_242.gs.
 *
 * SEGURANCA:
 * - Somente leitura.
 * - Nao grava dados.
 * - Nao altera validacoes.
 * - Nao consome sequencias.
 * - Nao cria vinculos.
 */

const DCP_TESTE_LISTAS_VINCULOS_246 = {
  ABAS: {
    VINCULOS: 'VINCULOS',
    AUDITORIA: 'AUDITORIA'
  },

  TESTES: [
    {
      aba: 'VINCULOS',
      coluna: 5,
      campo: 'TipoVinculo',
      valores: ['PROPRIO', 'ALUGADO', 'CEDIDO']
    },
    {
      aba: 'VINCULOS',
      coluna: 12,
      campo: 'Status',
      valores: ['ATIVO']
    },
    {
      aba: 'VINCULOS',
      coluna: 10,
      campo: 'PeriodicidadeAluguel',
      valores: ['DIARIA', 'SEMANAL', 'MENSAL']
    },
    {
      aba: 'AUDITORIA',
      coluna: 4,
      campo: 'Acao',
      valores: ['INSERCAO', 'ERRO']
    },
    {
      aba: 'AUDITORIA',
      coluna: 9,
      campo: 'Resultado',
      valores: ['SUCESSO', 'FALHA']
    },
    {
      aba: 'AUDITORIA',
      coluna: 10,
      campo: 'Origem',
      valores: ['SCRIPT']
    }
  ]
};


/**
 * Confere se uma celula aceita determinado valor
 * conforme sua regra de validacao de dados.
 *
 * Nao modifica a celula.
 */
function DCP_TL246_CONFERIR_VALOR(
  aba,
  linha,
  coluna,
  valor
) {
  const celula = aba.getRange(linha, coluna);
  const regra = celula.getDataValidation();

  if (!regra) {
    return {
      aceito: true,
      tipoRegra: 'SEM_RESTRICAO',
      detalhe: 'A celula nao possui lista suspensa ou regra de validacao.'
    };
  }

  const tipo = regra.getCriteriaType();
  const criterios = regra.getCriteriaValues();

  if (
    tipo === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST
  ) {
    const lista = criterios[0] || [];

    const aceito = lista.some(function(item) {
      return String(item).trim().toUpperCase() ===
        String(valor).trim().toUpperCase();
    });

    return {
      aceito: aceito,
      tipoRegra: 'VALUE_IN_LIST',
      valoresPermitidos: lista,
      detalhe: aceito
        ? 'Valor encontrado na lista.'
        : 'Valor nao encontrado na lista.'
    };
  }

  if (
    tipo === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE
  ) {
    const intervalo = criterios[0];

    if (!intervalo || typeof intervalo.getValues !== 'function') {
      return {
        aceito: null,
        tipoRegra: 'VALUE_IN_RANGE',
        detalhe: 'Nao foi possivel ler o intervalo da lista.'
      };
    }

    const lista = intervalo.getValues().flat();

    const aceito = lista.some(function(item) {
      return String(item).trim().toUpperCase() ===
        String(valor).trim().toUpperCase();
    });

    return {
      aceito: aceito,
      tipoRegra: 'VALUE_IN_RANGE',
      intervalo: intervalo.getA1Notation(),
      valoresPermitidos: lista,
      detalhe: aceito
        ? 'Valor encontrado no intervalo.'
        : 'Valor nao encontrado no intervalo.'
    };
  }

  return {
    aceito: null,
    tipoRegra: String(tipo),
    detalhe:
      'Regra diferente de lista direta ou intervalo. ' +
      'Necessaria verificacao manual.'
  };
}


/**
 * Executa a verificacao das listas nas linhas
 * que seriam utilizadas na proxima gravacao.
 *
 * Somente leitura.
 */
function DCP_TESTAR_LISTAS_VINCULOS_246() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const abaVinculos = ss.getSheetByName(
    DCP_TESTE_LISTAS_VINCULOS_246.ABAS.VINCULOS
  );

  const abaAuditoria = ss.getSheetByName(
    DCP_TESTE_LISTAS_VINCULOS_246.ABAS.AUDITORIA
  );

  if (!abaVinculos || !abaAuditoria) {
    throw new Error(
      'As abas VINCULOS e AUDITORIA precisam existir.'
    );
  }

  // Mesmas linhas de destino calculadas pelo modulo de cadastro.
  const linhaVinculo = Math.max(2, abaVinculos.getLastRow() + 1);
  const linhaAuditoria = Math.max(2, abaAuditoria.getLastRow() + 1);

  const resultados = [];

  DCP_TESTE_LISTAS_VINCULOS_246.TESTES.forEach(function(teste) {
    const aba = teste.aba === 'VINCULOS'
      ? abaVinculos
      : abaAuditoria;

    const linha = teste.aba === 'VINCULOS'
      ? linhaVinculo
      : linhaAuditoria;

    teste.valores.forEach(function(valor) {
      const verificacao = DCP_TL246_CONFERIR_VALOR(
        aba,
        linha,
        teste.coluna,
        valor
      );

      resultados.push({
        aba: teste.aba,
        celula: aba.getRange(linha, teste.coluna).getA1Notation(),
        campo: teste.campo,
        valorTestado: valor,
        aceito: verificacao.aceito,
        tipoRegra: verificacao.tipoRegra,
        detalhe: verificacao.detalhe
      });
    });
  });

  const aprovados = resultados.filter(function(item) {
    return item.aceito === true;
  }).length;

  const reprovados = resultados.filter(function(item) {
    return item.aceito === false;
  }).length;

  const inconclusivos = resultados.filter(function(item) {
    return item.aceito === null;
  }).length;

  const sucesso =
    reprovados === 0 &&
    inconclusivos === 0;

  const resposta = {
    sucesso: sucesso,
    somenteLeitura: true,
    dadosDaPlanilhaAlterados: false,
    sequenciasAlteradas: false,
    linhaDestinoVinculos: linhaVinculo,
    linhaDestinoAuditoria: linhaAuditoria,
    total: resultados.length,
    aprovados: aprovados,
    reprovados: reprovados,
    inconclusivos: inconclusivos,
    resultados: resultados
  };

  Logger.log(JSON.stringify(resposta, null, 2));

  return resposta;
}
