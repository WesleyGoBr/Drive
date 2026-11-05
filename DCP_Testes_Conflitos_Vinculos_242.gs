
/**
 * DRIVER CONTROL PRO
 * TESTES DE CONFLITOS DE VÍNCULOS - MÓDULO 2.4.2
 *
 * Testes totalmente simulados.
 * NÃO lê nem grava dados nas abas da planilha.
 * NÃO cria vínculos reais.
 */

function DCP_TESTAR_CONFLITOS_VINCULOS_242() {
  const motoristaA = 'MOTORISTA_A';
  const motoristaB = 'MOTORISTA_B';
  const veiculoA = 'VEICULO_A';
  const veiculoB = 'VEICULO_B';

  function data(texto) {
    const resultado = DCP_CV242_DATA(texto, 'teste');

    if (!resultado.valida || !resultado.valor) {
      throw new Error('Data inválida no teste: ' + texto);
    }

    return resultado.valor;
  }

  function vinculo(
    motoristaID,
    veiculoID,
    inicio,
    fim,
    autorizado
  ) {
    return [
      'UUID_TESTE',
      'VIN-TESTE',
      motoristaID,
      veiculoID,
      'CEDIDO',
      data(inicio),
      fim ? data(fim) : '',
      autorizado,
      '',
      '',
      'REGISTRO SIMULADO',
      'ATIVO',
      new Date(),
      new Date(),
      'TESTE',
      'TESTE'
    ];
  }

  function simular(
    existentes,
    motoristaID,
    veiculoID,
    inicio,
    fim,
    autorizacaoNova
  ) {
    const linhas = [
      [
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
    ].concat(existentes);

    const abaFalsa = {
      getLastRow: function() {
        return linhas.length;
      },

      getRange: function(linha, coluna, quantidadeLinhas, quantidadeColunas) {
        return {
          getValues: function() {
            return linhas
              .slice(linha - 1, linha - 1 + quantidadeLinhas)
              .map(function(registro) {
                return registro.slice(
                  coluna - 1,
                  coluna - 1 + quantidadeColunas
                );
              });
          }
        };
      }
    };

    return DCP_CV242_VERIFICAR_CONFLITOS(
      abaFalsa,
      motoristaID,
      veiculoID,
      data(inicio),
      fim ? data(fim) : null,
      autorizacaoNova
    );
  }

  const casos = [
    {
      nome: '1. Mesmo motorista em dois veículos no mesmo período',
      existentes: [
        vinculo(
          motoristaA,
          veiculoA,
          '2026-11-10T06:00',
          '2026-11-10T14:00',
          false
        )
      ],
      motorista: motoristaA,
      veiculo: veiculoB,
      inicio: '2026-11-10T08:00',
      fim: '2026-11-10T12:00',
      autorizacao: false,
      esperado: 'CONFLITO_MOTORISTA'
    },

    {
      nome: '2. Mesmo veículo com dois motoristas no mesmo período',
      existentes: [
        vinculo(
          motoristaA,
          veiculoA,
          '2026-11-10T06:00',
          '2026-11-10T14:00',
          true
        )
      ],
      motorista: motoristaB,
      veiculo: veiculoA,
      inicio: '2026-11-10T08:00',
      fim: '2026-11-10T12:00',
      autorizacao: true,
      esperado: 'CONFLITO_VEICULO'
    },

    {
      nome: '3. Mesmo veículo em períodos diferentes, ambos autorizados',
      existentes: [
        vinculo(
          motoristaA,
          veiculoA,
          '2026-11-10T06:00',
          '2026-11-10T14:00',
          true
        )
      ],
      motorista: motoristaB,
      veiculo: veiculoA,
      inicio: '2026-11-10T14:00',
      fim: '2026-11-10T22:00',
      autorizacao: true,
      esperado: 'SEM_CONFLITO'
    },

    {
      nome: '4. Mesmo veículo em períodos diferentes, vínculo existente sem autorização',
      existentes: [
        vinculo(
          motoristaA,
          veiculoA,
          '2026-11-10T06:00',
          '2026-11-10T14:00',
          false
        )
      ],
      motorista: motoristaB,
      veiculo: veiculoA,
      inicio: '2026-11-10T14:00',
      fim: '2026-11-10T22:00',
      autorizacao: true,
      esperado: 'AUTORIZACAO_COMPARTILHAMENTO'
    },

    {
      nome: '5. Mesmo veículo em períodos diferentes, novo vínculo sem autorização',
      existentes: [
        vinculo(
          motoristaA,
          veiculoA,
          '2026-11-10T06:00',
          '2026-11-10T14:00',
          true
        )
      ],
      motorista: motoristaB,
      veiculo: veiculoA,
      inicio: '2026-11-10T14:00',
      fim: '2026-11-10T22:00',
      autorizacao: false,
      esperado: 'AUTORIZACAO_COMPARTILHAMENTO'
    }
  ];

  const resultados = casos.map(function(caso) {
    try {
      const conflitos = simular(
        caso.existentes,
        caso.motorista,
        caso.veiculo,
        caso.inicio,
        caso.fim,
        caso.autorizacao
      );

      const tipos = conflitos.map(function(item) {
        return item.tipo;
      });

      const obtido = tipos.length
        ? tipos[0]
        : 'SEM_CONFLITO';

      return {
        teste: caso.nome,
        esperado: caso.esperado,
        obtido: obtido,
        aprovado: obtido === caso.esperado,
        conflitos: tipos
      };

    } catch (erro) {
      return {
        teste: caso.nome,
        esperado: caso.esperado,
        obtido: 'ERRO',
        aprovado: false,
        mensagem: erro.message
      };
    }
  });

  const aprovados = resultados.filter(function(item) {
    return item.aprovado;
  }).length;

  const resposta = {
    sucesso: aprovados === casos.length,
    somenteSimulacao: true,
    dadosDaPlanilhaAlterados: false,
    total: casos.length,
    aprovados: aprovados,
    reprovados: casos.length - aprovados,
    resultados: resultados
  };

  Logger.log(JSON.stringify(resposta, null, 2));

  return resposta;
}
