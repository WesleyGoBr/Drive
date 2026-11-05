
/**
 * DRIVER CONTROL PRO
 * Interface de edicao de veiculos - Modulo 2.3.5D
 *
 * Este arquivo integra a interface HTML aos modulos
 * de consulta e edicao ja existentes.
 */

/**
 * Abre a janela de edicao.
 */
function DCP_ABRIR_FORM_EDICAO_VEICULO_235D() {
  const html = HtmlService
    .createHtmlOutputFromFile('DCP_FormEdicaoVeiculo')
    .setWidth(1000)
    .setHeight(780);

  SpreadsheetApp.getUi().showModalDialog(
    html,
    'DRIVER CONTROL PRO - Editar veiculo'
  );
}


/**
 * Extrai a lista de registros retornada pelo modulo de consulta.
 */
function DCP_FORM235D_EXTRAIR_LISTA(resposta) {
  if (Array.isArray(resposta)) {
    return resposta;
  }

  if (!resposta || typeof resposta !== 'object') {
    return [];
  }

  const propriedades = [
    'veiculos',
    'registros',
    'dados',
    'lista',
    'resultados',
    'itens'
  ];

  for (let i = 0; i < propriedades.length; i++) {
    const valor = resposta[propriedades[i]];

    if (Array.isArray(valor)) {
      return valor;
    }
  }

  return [];
}


/**
 * Normaliza os campos de um veiculo para a interface.
 */
function DCP_FORM235D_NORMALIZAR_VEICULO(registro) {
  if (!registro || typeof registro !== 'object') {
    return null;
  }

  if (registro.veiculo && typeof registro.veiculo === 'object') {
    registro = registro.veiculo;
  }

  const campos = [
    'VeiculoID',
    'CodigoVeiculo',
    'Placa',
    'Marca',
    'Modelo',
    'AnoFabricacao',
    'AnoModelo',
    'TipoVeiculo',
    'TipoPropulsao',
    'ProprietarioTipo',
    'ProprietarioMotoristaID',
    'ProprietarioNome',
    'CapacidadeTanqueLitros',
    'CapacidadeBateriaKWh',
    'HodometroAtual',
    'Observacoes',
    'Status'
  ];

  const normalizado = {};

  campos.forEach(function(campo) {
    if (Object.prototype.hasOwnProperty.call(registro, campo)) {
      normalizado[campo] = registro[campo];
      return;
    }

    const primeiraMinuscula =
      campo.charAt(0).toLowerCase() + campo.slice(1);

    normalizado[campo] = registro[primeiraMinuscula] || '';
  });

  return normalizado;
}


/**
 * Pesquisa veiculos para a interface.
 */
function DCP_FORM235D_PESQUISAR_VEICULOS(pesquisa) {
  try {
    const resposta = DCP_LISTAR_VEICULOS_235D(
      'TODOS',
      String(pesquisa || '').trim()
    );

    if (resposta && resposta.sucesso === false) {
      throw new Error(
        resposta.mensagem || 'Falha na consulta de veiculos.'
      );
    }

    const listaBruta = DCP_FORM235D_EXTRAIR_LISTA(resposta);

    const veiculos = listaBruta
      .map(DCP_FORM235D_NORMALIZAR_VEICULO)
      .filter(function(v) {
        return v && v.VeiculoID;
      });

    return {
      sucesso: true,
      veiculos: veiculos
    };

  } catch (erro) {
    return {
      sucesso: false,
      mensagem: erro.message || String(erro),
      veiculos: []
    };
  }
}


/**
 * Consulta motoristas ativos para o campo de proprietario.
 */
function DCP_FORM235D_LISTAR_MOTORISTAS_ATIVOS() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName('MOTORISTAS');

  if (!aba) {
    throw new Error('A aba MOTORISTAS nao foi encontrada.');
  }

  const ultimaLinha = aba.getLastRow();

  if (ultimaLinha < 2) {
    return [];
  }

  const dados = aba.getRange(
    2,
    1,
    ultimaLinha - 1,
    13
  ).getValues();

  const motoristas = [];

  dados.forEach(function(linha) {
    const id = String(linha[0] || '').trim();
    const codigo = String(linha[1] || '').trim();
    const nome = String(linha[2] || '').trim();
    const status = String(linha[8] || '')
      .trim()
      .toUpperCase();

    if (id && nome && status === 'ATIVO') {
      motoristas.push({
        id: id,
        codigo: codigo,
        nome: nome
      });
    }
  });

  motoristas.sort(function(a, b) {
    return a.nome.localeCompare(b.nome, 'pt-BR');
  });

  return motoristas;
}


/**
 * Carrega os dados completos de um veiculo e os motoristas ativos.
 */
function DCP_FORM235D_CARREGAR_VEICULO(identificador) {
  try {
    const id = String(identificador || '').trim();

    if (!id) {
      throw new Error('Identificador do veiculo nao informado.');
    }

    const resposta = DCP_OBTER_VEICULO_235D(id);

    if (
      !resposta ||
      resposta.encontrado !== true ||
      !resposta.veiculo
    ) {
      throw new Error('Veiculo nao encontrado.');
    }

    const veiculo = DCP_FORM235D_NORMALIZAR_VEICULO(
      resposta.veiculo
    );

    if (!veiculo || !veiculo.VeiculoID) {
      throw new Error('Os dados do veiculo estao incompletos.');
    }

    const motoristas =
      DCP_FORM235D_LISTAR_MOTORISTAS_ATIVOS();

    return {
      sucesso: true,
      veiculo: veiculo,
      motoristas: motoristas
    };

  } catch (erro) {
    return {
      sucesso: false,
      mensagem: erro.message || String(erro)
    };
  }
}
