
/**
 * DRIVER CONTROL PRO V1.0
 * INTEGRACAO DO FORMULARIO DE ENCERRAMENTO - 2.4.4.3
 *
 * Diagnostico reforcado da listagem de vinculos.
 *
 * IMPORTANTE:
 * - Nao altera registros.
 * - Nao encerra vinculos.
 * - Nao modifica CONFIG, HISTORICO ou AUDITORIA.
 * - Mantem o motor principal separado.
 */

const DCP_FORM_ENC244 = Object.freeze({
  ABA: 'VINCULOS',
  NOME_HTML: 'DCP_FormEncerramentoVinculo',
  FUSO: 'America/Sao_Paulo',
  USUARIO: 'ADMIN_CENTRAL',

  CABECALHOS: Object.freeze([
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
  ])
});


function DCP_FORM244_TEXTO(valor) {
  return String(valor == null ? '' : valor).trim();
}


function DCP_FORM244_NORMALIZAR(valor) {
  return DCP_FORM244_TEXTO(valor)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toUpperCase();
}


function DCP_FORM244_ESTA_VAZIO(valor) {
  return valor === null ||
    valor === undefined ||
    String(valor).trim() === '';
}


function DCP_FORM244_ABA(nome) {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const aba = planilha.getSheetByName(nome);

  if (!aba) {
    throw new Error('Aba nao encontrada: ' + nome);
  }

  return aba;
}


function DCP_FORM244_VALIDAR_ESTRUTURA(aba) {
  const encontrados = aba
    .getRange(1, 1, 1, DCP_FORM_ENC244.CABECALHOS.length)
    .getDisplayValues()[0];

  const divergencias = [];

  DCP_FORM_ENC244.CABECALHOS.forEach(function(campo, indice) {
    if (String(encontrados[indice]).trim() !== campo) {
      divergencias.push({
        coluna: indice + 1,
        esperado: campo,
        encontrado: encontrados[indice] || ''
      });
    }
  });

  if (divergencias.length) {
    throw new Error(
      'Estrutura VINCULOS divergente: ' +
      JSON.stringify(divergencias)
    );
  }

  return encontrados;
}


function DCP_FORM244_FORMATAR_DATA(valor) {
  if (!(valor instanceof Date) || isNaN(valor.getTime())) {
    return '';
  }

  return Utilities.formatDate(
    valor,
    DCP_FORM_ENC244.FUSO,
    'dd/MM/yyyy HH:mm'
  );
}


function DCP_FORM244_OBTER_DADOS_ABA() {
  const aba = DCP_FORM244_ABA(DCP_FORM_ENC244.ABA);

  DCP_FORM244_VALIDAR_ESTRUTURA(aba);

  const ultimaLinha = aba.getLastRow();

  if (ultimaLinha < 2) {
    return {
      aba: aba,
      ultimaLinha: ultimaLinha,
      dados: [],
      exibidos: [],
      formulas: []
    };
  }

  const quantidade = ultimaLinha - 1;
  const intervalo = aba.getRange(2, 1, quantidade, 16);

  return {
    aba: aba,
    ultimaLinha: ultimaLinha,
    dados: intervalo.getValues(),
    exibidos: intervalo.getDisplayValues(),
    formulas: intervalo.getFormulas()
  };
}


/**
 * Lista somente vinculos ativos e sem data de encerramento.
 * Operacao somente leitura.
 */
function DCP_FORM244_LISTAR_VINCULOS_ATIVOS() {

  const lock = LockService.getDocumentLock();

  if (!lock.tryLock(15000)) {
    throw new Error(
      'Sistema ocupado. Aguarde e tente carregar novamente.'
    );
  }

  try {
    const planilha = SpreadsheetApp.getActiveSpreadsheet();

    if (
      Session.getScriptTimeZone() !== DCP_FORM_ENC244.FUSO ||
      planilha.getSpreadsheetTimeZone() !== DCP_FORM_ENC244.FUSO
    ) {
      throw new Error(
        'O fuso do projeto e da planilha deve ser America/Sao_Paulo.'
      );
    }

    const leitura = DCP_FORM244_OBTER_DADOS_ABA();
    const vinculos = [];

    leitura.dados.forEach(function(registro) {

      const status = DCP_FORM244_NORMALIZAR(registro[11]);
      const fim = registro[6];

      if (status !== 'ATIVO') return;
      if (!DCP_FORM244_ESTA_VAZIO(fim)) return;

      const vinculoID = DCP_FORM244_TEXTO(registro[0]);
      const codigoVinculo = DCP_FORM244_TEXTO(registro[1]);

      if (!vinculoID || !codigoVinculo) return;

      vinculos.push({
        vinculoID: vinculoID,
        codigoVinculo: codigoVinculo,
        motoristaID: DCP_FORM244_TEXTO(registro[2]),
        veiculoID: DCP_FORM244_TEXTO(registro[3]),
        tipoVinculo: DCP_FORM244_TEXTO(registro[4]),
        inicioEm: DCP_FORM244_FORMATAR_DATA(registro[5]),
        compartilhamentoAutorizado:
          registro[7] === true ||
          DCP_FORM244_NORMALIZAR(registro[7]) === 'TRUE'
      });
    });

    vinculos.sort(function(a, b) {
      return a.codigoVinculo.localeCompare(b.codigoVinculo);
    });

    return {
      sucesso: true,
      quantidade: vinculos.length,
      vinculos: vinculos
    };

  } finally {
    lock.releaseLock();
  }
}


/**
 * Carrega os dados para o formulario.
 */
function DCP_FORM244_CARREGAR_DADOS() {
  try {
    const resultado = DCP_FORM244_LISTAR_VINCULOS_ATIVOS();

    return {
      sucesso: true,
      quantidade: resultado.quantidade,
      vinculos: resultado.vinculos,
      mensagem: resultado.quantidade
        ? 'Vinculos ativos carregados.'
        : 'Nenhum vinculo ativo disponivel para encerramento.'
    };

  } catch (erro) {
    return {
      sucesso: false,
      quantidade: 0,
      mensagem: erro.message,
      vinculos: []
    };
  }
}


/**
 * Encaminha a solicitacao ao motor principal.
 * O controle transacional permanece no motor principal.
 */
function DCP_FORM244_ENCERRAR_VINCULO(dados) {

  if (!dados || typeof dados !== 'object') {
    return {
      sucesso: false,
      codigo: 'DADOS_INVALIDOS',
      mensagem: 'Solicitacao nao informada.',
      gravacaoRealizada: false
    };
  }

  try {
    return DCP_ENCERRAR_VINCULO_244(dados);

  } catch (erro) {
    console.error(
      'Erro na integracao do encerramento: ' + erro.message
    );

    return {
      sucesso: false,
      codigo: 'ERRO_INTEGRACAO',
      mensagem: erro.message,
      gravacaoRealizada: false,
      gravacaoParcial: false,
      tentativaInconclusiva: false
    };
  }
}


/**
 * Abre o formulario em janela modal.
 */
function DCP_ABRIR_FORM_ENCERRAMENTO_VINCULO_244() {

  const html = HtmlService
    .createHtmlOutputFromFile(
      DCP_FORM_ENC244.NOME_HTML
    )
    .setWidth(720)
    .setHeight(760);

  SpreadsheetApp.getUi().showModalDialog(
    html,
    'Encerramento de vinculo'
  );
}


/**
 * DIAGNOSTICO DETALHADO.
 *
 * Examina todas as linhas e informa por que cada uma
 * foi aceita ou descartada pelo filtro.
 *
 * Somente leitura.
 */
function DCP_DIAGNOSTICAR_INTEGRACAO_ENCERRAMENTO_244() {

  const relatorio = {
    modulo: '2.4.4.3',
    somenteLeitura: true,
    sucesso: false,
    abaVinculosExiste: false,
    estruturaValida: false,
    ultimaLinha: 0,
    fusoProjeto: Session.getScriptTimeZone(),
    fusoPlanilha: '',
    fusoEsperado: DCP_FORM_ENC244.FUSO,
    totalLinhasDados: 0,
    statusAtivoEncontrado: 0,
    fimVazioEncontrado: 0,
    candidatosValidos: 0,
    linhas: [],
    mensagem: ''
  };

  try {
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    relatorio.fusoPlanilha = planilha.getSpreadsheetTimeZone();

    const aba = planilha.getSheetByName('VINCULOS');

    if (!aba) {
      relatorio.mensagem = 'Aba VINCULOS nao encontrada.';
      console.log(JSON.stringify(relatorio, null, 2));
      return relatorio;
    }

    relatorio.abaVinculosExiste = true;
    relatorio.ultimaLinha = aba.getLastRow();

    DCP_FORM244_VALIDAR_ESTRUTURA(aba);
    relatorio.estruturaValida = true;

    const leitura = DCP_FORM244_OBTER_DADOS_ABA();
    relatorio.totalLinhasDados = leitura.dados.length;

    leitura.dados.forEach(function(registro, indice) {

      const linhaPlanilha = indice + 2;

      const statusBruto = registro[11];
      const statusExibido = leitura.exibidos[indice][11];

      const fimBruto = registro[6];
      const fimExibido = leitura.exibidos[indice][6];
      const formulaFim = leitura.formulas[indice][6];

      const statusNormalizado = DCP_FORM244_NORMALIZAR(statusBruto);
      const fimVazio = DCP_FORM244_ESTA_VAZIO(fimBruto);

      const id = DCP_FORM244_TEXTO(registro[0]);
      const codigo = DCP_FORM244_TEXTO(registro[1]);

      const motivos = [];

      if (statusNormalizado === 'ATIVO') {
        relatorio.statusAtivoEncontrado++;
      } else {
        motivos.push('STATUS_DIFERENTE_DE_ATIVO');
      }

      if (fimVazio) {
        relatorio.fimVazioEncontrado++;
      } else {
        motivos.push('FIM_EM_PREENCHIDO');
      }

      if (!id) motivos.push('VINCULO_ID_VAZIO');
      if (!codigo) motivos.push('CODIGO_VINCULO_VAZIO');

      const elegivel =
        statusNormalizado === 'ATIVO' &&
        fimVazio &&
        Boolean(id) &&
        Boolean(codigo);

      if (elegivel) {
        relatorio.candidatosValidos++;
      }

      relatorio.linhas.push({
        linha: linhaPlanilha,
        vinculoID: id,
        codigoVinculo: codigo,
        statusBruto: String(statusBruto == null ? '' : statusBruto),
        statusExibido: statusExibido,
        statusNormalizado: statusNormalizado,
        fimEmBruto:
          fimBruto instanceof Date
            ? fimBruto.toISOString()
            : String(fimBruto == null ? '' : fimBruto),
        fimEmExibido: fimExibido,
        formulaFimEm: formulaFim,
        fimVazio: fimVazio,
        elegivel: elegivel,
        motivos: motivos
      });
    });

    const fusoCorreto =
      relatorio.fusoProjeto === DCP_FORM_ENC244.FUSO &&
      relatorio.fusoPlanilha === DCP_FORM_ENC244.FUSO;

    relatorio.fusoCorreto = fusoCorreto;
    relatorio.sucesso = relatorio.estruturaValida && fusoCorreto;

    relatorio.mensagem = relatorio.candidatosValidos
      ? 'Foram encontrados vinculos elegiveis para encerramento.'
      : 'Nenhum vinculo passou por todos os criterios do filtro. Consulte as linhas do relatorio.';

  } catch (erro) {
    relatorio.erro = erro.message;
    relatorio.mensagem = 'Falha durante o diagnostico.';
  }

  console.log(
    '=== DIAGNOSTICO DETALHADO DA INTEGRACAO 2.4.4.3 ==='
  );

  console.log(JSON.stringify(relatorio, null, 2));

  return relatorio;
}


/**
 * Testa o carregamento e mostra o resultado nos logs.
 */
function DCP_TESTAR_CARREGAMENTO_FORMULARIO_ENCERRAMENTO_244() {

  const resultado = DCP_FORM244_CARREGAR_DADOS();

  console.log(
    '=== TESTE DE CARREGAMENTO FORMULARIO 2.4.4.3 ==='
  );

  console.log(JSON.stringify(resultado, null, 2));

  return resultado;
}
