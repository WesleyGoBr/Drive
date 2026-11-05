
/**
 * DRIVER CONTROL PRO
 * MÓDULO 2.4
 * DIAGNÓSTICO INICIAL DE VÍNCULOS
 *
 * SOMENTE LEITURA.
 * Não cadastra, altera ou exclui informações.
 */

const DCP_DIAGNOSTICO_VINCULOS_240 = {
  ABAS: {
    VINCULOS: 'VINCULOS',
    MOTORISTAS: 'MOTORISTAS',
    VEICULOS: 'VEICULOS'
  },

  CABECALHOS: {
    VINCULOS: [
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

    MOTORISTAS: [
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
    ],

    VEICULOS: [
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
  }
};


/**
 * Confere os cabeçalhos de uma aba.
 */
function DCP_DV240_CONFERIR_ABA(nomeAba, esperados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(nomeAba);

  if (!aba) {
    return {
      existe: false,
      aba: nomeAba,
      divergencias: ['Aba não encontrada']
    };
  }

  const encontrados = aba
    .getRange(1, 1, 1, esperados.length)
    .getValues()[0]
    .map(v => String(v || '').trim());

  const divergencias = [];

  esperados.forEach((campo, i) => {
    if (encontrados[i] !== campo) {
      divergencias.push({
        coluna: i + 1,
        esperado: campo,
        encontrado: encontrados[i]
      });
    }
  });

  return {
    existe: true,
    aba: nomeAba,
    ultimaLinha: aba.getLastRow(),
    quantidadeColunasEsperadas: esperados.length,
    cabecalhosCorretos: divergencias.length === 0,
    divergencias: divergencias
  };
}


/**
 * Diagnóstico geral da estrutura e dos registros.
 */
function DCP_DIAGNOSTICAR_ESTRUTURA_VINCULOS_240() {
  const lock = LockService.getDocumentLock();

  if (!lock || !lock.tryLock(30000)) {
    throw new Error('Não foi possível obter o bloqueio da planilha.');
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cfg = DCP_DIAGNOSTICO_VINCULOS_240;

    const estrutura = {
      VINCULOS: DCP_DV240_CONFERIR_ABA(
        cfg.ABAS.VINCULOS,
        cfg.CABECALHOS.VINCULOS
      ),

      MOTORISTAS: DCP_DV240_CONFERIR_ABA(
        cfg.ABAS.MOTORISTAS,
        cfg.CABECALHOS.MOTORISTAS
      ),

      VEICULOS: DCP_DV240_CONFERIR_ABA(
        cfg.ABAS.VEICULOS,
        cfg.CABECALHOS.VEICULOS
      )
    };

    const estruturaOK = Object.values(estrutura).every(
      item => item.existe && item.cabecalhosCorretos
    );

    if (!estruturaOK) {
      const resultadoEstrutura = {
        sucesso: false,
        etapa: 'ESTRUTURA',
        estrutura: estrutura,
        mensagem: 'Existem problemas estruturais. Nenhum registro foi analisado.'
      };

      Logger.log(JSON.stringify(resultadoEstrutura, null, 2));
      return resultadoEstrutura;
    }

    const abaVinculos = ss.getSheetByName(cfg.ABAS.VINCULOS);
    const abaMotoristas = ss.getSheetByName(cfg.ABAS.MOTORISTAS);
    const abaVeiculos = ss.getSheetByName(cfg.ABAS.VEICULOS);

    const vinculos = abaVinculos.getLastRow() >= 2
      ? abaVinculos.getRange(
          2, 1, abaVinculos.getLastRow() - 1, 16
        ).getValues()
      : [];

    const motoristas = abaMotoristas.getLastRow() >= 2
      ? abaMotoristas.getRange(
          2, 1, abaMotoristas.getLastRow() - 1, 13
        ).getValues()
      : [];

    const veiculos = abaVeiculos.getLastRow() >= 2
      ? abaVeiculos.getRange(
          2, 1, abaVeiculos.getLastRow() - 1, 24
        ).getValues()
      : [];

    const idsMotoristas = new Set(
      motoristas.map(l => String(l[0] || '').trim()).filter(Boolean)
    );

    const idsVeiculos = new Set(
      veiculos.map(l => String(l[0] || '').trim()).filter(Boolean)
    );

    const codigosVinculos = new Set();
    const idsVinculos = new Set();

    const problemas = [];
    const resumoVinculos = [];

    let ativos = 0;
    let inativos = 0;
    let outrosStatus = 0;

    vinculos.forEach((linha, indice) => {
      const numeroLinha = indice + 2;

      const id = String(linha[0] || '').trim();
      const codigo = String(linha[1] || '').trim();
      const motoristaID = String(linha[2] || '').trim();
      const veiculoID = String(linha[3] || '').trim();
      const inicio = linha[5];
      const fim = linha[6];
      const status = String(linha[11] || '').trim().toUpperCase();

      if (!id) {
        problemas.push({
          linha: numeroLinha,
          problema: 'UUID do vínculo vazio'
        });
      } else if (idsVinculos.has(id)) {
        problemas.push({
          linha: numeroLinha,
          problema: 'UUID de vínculo duplicado',
          identificador: id
        });
      }
      idsVinculos.add(id);

      if (!codigo) {
        problemas.push({
          linha: numeroLinha,
          problema: 'Código do vínculo vazio'
        });
      } else if (codigosVinculos.has(codigo)) {
        problemas.push({
          linha: numeroLinha,
          problema: 'Código de vínculo duplicado',
          identificador: codigo
        });
      }
      codigosVinculos.add(codigo);

      if (!motoristaID || !idsMotoristas.has(motoristaID)) {
        problemas.push({
          linha: numeroLinha,
          problema: 'Motorista inexistente ou não informado',
          codigoVinculo: codigo
        });
      }

      if (!veiculoID || !idsVeiculos.has(veiculoID)) {
        problemas.push({
          linha: numeroLinha,
          problema: 'Veículo inexistente ou não informado',
          codigoVinculo: codigo
        });
      }

      if (!inicio || !(inicio instanceof Date) ||
          isNaN(inicio.getTime())) {
        problemas.push({
          linha: numeroLinha,
          problema: 'Data de início inválida ou vazia',
          codigoVinculo: codigo
        });
      }

      if (fim && (!(fim instanceof Date) || isNaN(fim.getTime()))) {
        problemas.push({
          linha: numeroLinha,
          problema: 'Data de fim inválida',
          codigoVinculo: codigo
        });
      }

      if (inicio instanceof Date && fim instanceof Date &&
          !isNaN(inicio.getTime()) && !isNaN(fim.getTime()) &&
          fim < inicio) {
        problemas.push({
          linha: numeroLinha,
          problema: 'Data de fim anterior à data de início',
          codigoVinculo: codigo
        });
      }

      if (status === 'ATIVO') {
        ativos++;
      } else if (status === 'INATIVO') {
        inativos++;
      } else {
        outrosStatus++;
        problemas.push({
          linha: numeroLinha,
          problema: 'Status vazio ou não reconhecido',
          codigoVinculo: codigo,
          status: status
        });
      }

      resumoVinculos.push({
        linha: numeroLinha,
        codigoVinculo: codigo,
        motoristaExiste: idsMotoristas.has(motoristaID),
        veiculoExiste: idsVeiculos.has(veiculoID),
        status: status,
        inicio: inicio instanceof Date
          ? inicio.toISOString()
          : String(inicio || ''),
        fim: fim instanceof Date
          ? fim.toISOString()
          : String(fim || '')
      });
    });

    const resultado = {
      sucesso: true,
      somenteLeitura: true,

      estrutura: estrutura,

      quantidadeMotoristas: motoristas.length,
      quantidadeVeiculos: veiculos.length,
      quantidadeVinculos: vinculos.length,

      vinculosAtivos: ativos,
      vinculosInativos: inativos,
      vinculosComStatusIrregular: outrosStatus,

      problemasEncontrados: problemas.length,
      problemas: problemas,

      registros: resumoVinculos,

      mensagem: problemas.length === 0
        ? 'Estrutura e registros analisados sem inconsistências detectadas.'
        : 'Diagnóstico concluído. Existem pontos que precisam ser avaliados.'
    };

    Logger.log(JSON.stringify(resultado, null, 2));

    return resultado;

  } finally {
    lock.releaseLock();
  }
}
