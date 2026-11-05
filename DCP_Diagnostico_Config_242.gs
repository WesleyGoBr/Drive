
/**
 * DRIVER CONTROL PRO
 * Diagnóstico da aba CONFIG para o módulo de vínculos 2.4.2
 *
 * Objetivo:
 * - Conferir os cabeçalhos.
 * - Inspecionar as validações de dados.
 * - Verificar a próxima linha disponível.
 * - Não alterar dados, formatação ou validações.
 */

function DCP_DIAGNOSTICAR_CONFIG_VINCULOS_242() {
  const resultado = {
    sucesso: false,
    somenteLeitura: true,
    dadosAlterados: false,
    testes: [],
    erros: []
  };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName('CONFIG');

    if (!aba) {
      throw new Error('A aba CONFIG não foi encontrada.');
    }

    const cabecalhosEsperados = [
      'Chave',
      'Valor',
      'TipoValor',
      'Descricao',
      'Editavel',
      'AtualizadoEm'
    ];

    const ultimaColuna = cabecalhosEsperados.length;
    const ultimaLinha = Math.max(aba.getLastRow(), 1);

    const cabecalhos = aba
      .getRange(1, 1, 1, ultimaColuna)
      .getDisplayValues()[0];

    const cabecalhosCorretos = cabecalhos.every(
      (valor, indice) =>
        String(valor).trim() === cabecalhosEsperados[indice]
    );

    resultado.testes.push({
      teste: 'Cabeçalhos da CONFIG',
      aprovado: cabecalhosCorretos,
      esperado: cabecalhosEsperados,
      encontrado: cabecalhos
    });

    const proximaLinha = ultimaLinha + 1;

    const dadosExistentes = aba
      .getRange(2, 1, Math.max(ultimaLinha - 1, 1), ultimaColuna)
      .getDisplayValues();

    const sequencias = dadosExistentes.filter(
      linha => String(linha[0]).trim() === 'SEQ_VIN'
    );

    resultado.testes.push({
      teste: 'Identificação da sequência SEQ_VIN',
      aprovado: sequencias.length === 1,
      quantidadeEncontrada: sequencias.length,
      valorEncontrado: sequencias.length === 1
        ? sequencias[0][1]
        : null
    });

    const linhaDestino = aba.getRange(
      proximaLinha,
      1,
      1,
      ultimaColuna
    );

    const validacoes = linhaDestino.getDataValidations()[0];

    const nomesColunas = [
      'Chave',
      'Valor',
      'TipoValor',
      'Descricao',
      'Editavel',
      'AtualizadoEm'
    ];

    const validacoesDestino = [];

    for (let i = 0; i < validacoes.length; i++) {
      const regra = validacoes[i];

      let descricao = 'SEM_VALIDACAO';

      if (regra) {
        descricao = String(regra.getCriteriaType());

        try {
          const argumentos = regra.getCriteriaValues();

          descricao += ' | ' + JSON.stringify(
            argumentos.map(item => {
              if (item && typeof item.getA1Notation === 'function') {
                return item.getA1Notation();
              }
              return String(item);
            })
          );
        } catch (e) {
          descricao += ' | argumentos não disponíveis';
        }
      }

      validacoesDestino.push({
        coluna: nomesColunas[i],
        validacao: descricao
      });
    }

    resultado.testes.push({
      teste: 'Validações na próxima linha',
      aprovado: true,
      linhaInspecionada: proximaLinha,
      colunas: validacoesDestino
    });

    const linhasComChaveIdemp = dadosExistentes.filter(
      linha => String(linha[0]).trim().indexOf('IDEMP_VIN_') === 0
    );

    resultado.testes.push({
      teste: 'Chaves de idempotência existentes',
      aprovado: true,
      quantidade: linhasComChaveIdemp.length,
      linhas: linhasComChaveIdemp.map((linha, indice) => ({
        linha: indice + 2,
        chave: linha[0],
        estado: linha[1]
      }))
    });

    resultado.proximaLinhaDisponivel = proximaLinha;
    resultado.ultimaLinhaComConteudo = ultimaLinha;
    resultado.sucesso = cabecalhosCorretos && sequencias.length === 1;

  } catch (erro) {
    resultado.erros.push({
      mensagem: erro.message,
      detalhe: erro.stack || ''
    });
  }

  Logger.log(JSON.stringify(resultado, null, 2));
  return resultado;
}
