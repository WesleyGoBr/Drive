
/**
 * DRIVER CONTROL PRO
 * MODULO 2.4.3 - INTERFACE DE VINCULOS
 *
 * Carregamento de dados e encaminhamento ao cadastro oficial.
 * A gravacao definitiva ocorre no modulo 2.4.2.
 *
 * VERSAO REVISADA:
 * - Mantem as estruturas e indices originais.
 * - Diferencia erros de validacao de falhas inesperadas.
 * - Bloqueia reenvio quando o resultado do servidor e incerto.
 */

const DCP_INTERFACE_VINCULOS_243 = {
  TITULO: 'DRIVER CONTROL PRO - Cadastro de Vinculo',
  LARGURA: 700,
  ALTURA: 800
};


/**
 * Abre a janela do formulario.
 */
function DCP_ABRIR_FORM_VINCULO_243() {
  const html = HtmlService
    .createHtmlOutputFromFile('DCP_FormVinculo')
    .setWidth(DCP_INTERFACE_VINCULOS_243.LARGURA)
    .setHeight(DCP_INTERFACE_VINCULOS_243.ALTURA);

  SpreadsheetApp.getUi().showModalDialog(
    html,
    DCP_INTERFACE_VINCULOS_243.TITULO
  );
}


/**
 * Lista somente motoristas ativos.
 */
function DCP_FORM243_LISTAR_MOTORISTAS() {
  const cabecalhos =
    DCP_CADASTRO_VINCULOS_242.CABECALHOS.MOTORISTAS;

  const aba = DCP_CV242_CONFERIR_CABECALHOS(
    'MOTORISTAS',
    cabecalhos
  );

  const ultimaLinha = aba.getLastRow();

  if (ultimaLinha < 2) {
    return {
      sucesso: true,
      registros: [],
      mensagem: 'Nenhum motorista cadastrado.'
    };
  }

  const dados = aba.getRange(
    2,
    1,
    ultimaLinha - 1,
    cabecalhos.length
  ).getValues();

  const registros = dados
    .filter(function(linha) {
      return String(linha[8] || '')
        .trim()
        .toUpperCase() === 'ATIVO';
    })
    .map(function(linha) {
      return {
        id: String(linha[0] || '').trim(),
        codigo: String(linha[1] || '').trim(),
        nome: String(linha[2] || '').trim(),
        nomeExibicao: String(linha[3] || '').trim()
      };
    })
    .filter(function(item) {
      return item.id && item.codigo && item.nome;
    });

  return {
    sucesso: true,
    registros: registros,
    mensagem: registros.length
      ? registros.length + ' motorista(s) ativo(s) encontrado(s).'
      : 'Nenhum motorista ativo disponivel.'
  };
}


/**
 * Lista somente veiculos ativos.
 */
function DCP_FORM243_LISTAR_VEICULOS() {
  const cabecalhos =
    DCP_CADASTRO_VINCULOS_242.CABECALHOS.VEICULOS;

  const aba = DCP_CV242_CONFERIR_CABECALHOS(
    'VEICULOS',
    cabecalhos
  );

  const ultimaLinha = aba.getLastRow();

  if (ultimaLinha < 2) {
    return {
      sucesso: true,
      registros: [],
      mensagem: 'Nenhum veiculo cadastrado.'
    };
  }

  const dados = aba.getRange(
    2,
    1,
    ultimaLinha - 1,
    cabecalhos.length
  ).getValues();

  const registros = dados
    .filter(function(linha) {
      return String(linha[19] || '')
        .trim()
        .toUpperCase() === 'ATIVO';
    })
    .map(function(linha) {
      return {
        id: String(linha[0] || '').trim(),
        codigo: String(linha[1] || '').trim(),
        placa: String(linha[2] || '').trim(),
        marca: String(linha[3] || '').trim(),
        modelo: String(linha[4] || '').trim(),
        proprietarioTipo: String(linha[11] || '')
          .trim()
          .toUpperCase(),
        proprietarioMotoristaID: String(linha[12] || '').trim(),
        proprietarioNome: String(linha[13] || '').trim()
      };
    })
    .filter(function(item) {
      return item.id && item.codigo && item.placa;
    });

  return {
    sucesso: true,
    registros: registros,
    mensagem: registros.length
      ? registros.length + ' veiculo(s) ativo(s) encontrado(s).'
      : 'Nenhum veiculo ativo disponivel.'
  };
}


/**
 * Carrega os dados necessarios para o formulario.
 */
function DCP_FORM243_CARREGAR_DADOS() {
  try {
    const motoristas = DCP_FORM243_LISTAR_MOTORISTAS();
    const veiculos = DCP_FORM243_LISTAR_VEICULOS();

    return {
      sucesso: true,
      motoristas: motoristas.registros,
      veiculos: veiculos.registros,
      mensagem:
        motoristas.registros.length + ' motorista(s) e ' +
        veiculos.registros.length + ' veiculo(s) ativo(s) carregados.'
    };

  } catch (erro) {
    return {
      sucesso: false,
      motoristas: [],
      veiculos: [],
      mensagem: erro && erro.message
        ? erro.message
        : 'Erro ao carregar os dados.'
    };
  }
}


/**
 * Encaminha o cadastro para o modulo oficial 2.4.2.
 *
 * Validacoes feitas antes da chamada oficial sao consideradas
 * seguras para nova tentativa, pois ainda nao houve encaminhamento.
 *
 * Excecoes durante a chamada oficial sao tratadas como inconclusivas:
 * nao presumimos que nenhuma gravacao tenha ocorrido.
 */
function DCP_FORM243_CADASTRAR_VINCULO(dados) {

  // Validacoes locais do servidor, antes do cadastro oficial.
  if (!dados || typeof dados !== 'object' || Array.isArray(dados)) {
    return {
      sucesso: false,
      gravacaoRealizada: false,
      tentativaInconclusiva: false,
      mensagem: 'Dados do formulario nao recebidos corretamente.'
    };
  }

  if (
    typeof dados.ChaveTentativa !== 'string' ||
    !dados.ChaveTentativa.trim()
  ) {
    return {
      sucesso: false,
      gravacaoRealizada: false,
      tentativaInconclusiva: false,
      mensagem:
        'Chave da tentativa nao recebida. Feche e abra um novo formulario.'
    };
  }

  // A partir daqui, o cadastro oficial pode iniciar operacoes.
  try {
    return DCP_CADASTRAR_VINCULO_242(dados);

  } catch (erro) {
    return {
      sucesso: false,
      erroComunicacao: false,
      tentativaInconclusiva: true,
      gravacaoRealizada: null,
      gravacaoParcial: true,
      mensagem:
        'Ocorreu uma falha inesperada durante o processamento. ' +
        'Nao envie novamente. A situacao da tentativa precisa ser ' +
        'verificada antes de qualquer nova operacao. Detalhe: ' +
        (erro && erro.message
          ? erro.message
          : 'Erro sem descricao.')
    };
  }
}
