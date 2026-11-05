
/**
 * DRIVER CONTROL PRO
 * TESTES DE IDEMPOTENCIA - MODULO 2.4.2.2
 *
 * Testes exclusivamente em memoria.
 * Nao acessa planilhas reais.
 * Nao altera CONFIG, SEQ_VIN, VINCULOS ou AUDITORIA.
 */

function DCP_TESTAR_IDEMPOTENCIA_VINCULOS_242() {
  const resultados = [];

  function testar(nome, funcao) {
    try {
      const retorno = funcao();

      resultados.push({
        teste: nome,
        aprovado: retorno === true,
        detalhe: retorno === true ? 'OK' : 'Resultado inesperado.'
      });
    } catch (erro) {
      resultados.push({
        teste: nome,
        aprovado: false,
        detalhe: erro.message
      });
    }
  }

  /*
   * SIMULADOR DE INTERVALO
   */
  function FakeRange(aba, linha, coluna, quantidadeLinhas, quantidadeColunas) {
    this.aba = aba;
    this.linha = linha;
    this.coluna = coluna;
    this.quantidadeLinhas = quantidadeLinhas || 1;
    this.quantidadeColunas = quantidadeColunas || 1;
  }

  FakeRange.prototype.getValue = function() {
    return this.aba.dados[this.linha - 1][this.coluna - 1];
  };

  FakeRange.prototype.setValue = function(valor) {
    this.aba.dados[this.linha - 1][this.coluna - 1] = valor;
    return this;
  };

  FakeRange.prototype.getValues = function() {
    const resultado = [];

    for (let i = 0; i < this.quantidadeLinhas; i++) {
      const linha = [];

      for (let j = 0; j < this.quantidadeColunas; j++) {
        linha.push(
          this.aba.dados[this.linha - 1 + i][this.coluna - 1 + j]
        );
      }

      resultado.push(linha);
    }

    return resultado;
  };

  FakeRange.prototype.setValues = function(valores) {
    for (let i = 0; i < valores.length; i++) {
      for (let j = 0; j < valores[i].length; j++) {
        this.aba.dados[this.linha - 1 + i][this.coluna - 1 + j] =
          valores[i][j];
      }
    }

    return this;
  };

  FakeRange.prototype.clearDataValidations = function() {
    return this;
  };

  /*
   * SIMULADOR DE ABA
   */
  function FakeSheet() {
    this.dados = [
      ['Chave', 'Valor', 'TipoValor', 'Descricao', 'Editavel', 'AtualizadoEm']
    ];
  }

  FakeSheet.prototype.getLastRow = function() {
    return this.dados.length;
  };

  FakeSheet.prototype.getRange = function(
    linha, coluna, quantidadeLinhas, quantidadeColunas
  ) {
    while (this.dados.length < linha + (quantidadeLinhas || 1) - 1) {
      this.dados.push(['', '', '', '', '', '']);
    }

    return new FakeRange(
      this,
      linha,
      coluna,
      quantidadeLinhas,
      quantidadeColunas
    );
  };

  FakeSheet.prototype.appendRow = function(linha) {
    this.dados.push(linha.slice());
    return this;
  };

  const chave = 'teste-chave-12345678901234567890';

  /*
   * 1. CRIAR TENTATIVA
   */
  testar('Criacao de tentativa em memoria', function() {
    const fake = new FakeSheet();

    const linha = DCP_CV242_CRIAR_TENTATIVA(
      fake,
      chave,
      { teste: true }
    );

    return linha === 2 &&
      fake.dados[1][0] === 'IDEMP_VIN_' + chave;
  });

  /*
   * 2. LOCALIZAR TENTATIVA
   */
  testar('Localizacao de chave existente', function() {
    const fake = new FakeSheet();

    DCP_CV242_CRIAR_TENTATIVA(
      fake,
      chave,
      { teste: true }
    );

    const tentativa = DCP_CV242_LOCALIZAR_TENTATIVA(
      fake,
      chave
    );

    return tentativa !== null &&
      tentativa.linha === 2;
  });

  /*
   * 3. ATUALIZAR COM CHAVE CORRETA
   */
  testar('Atualizacao com chave correspondente', function() {
    const fake = new FakeSheet();

    const linha = DCP_CV242_CRIAR_TENTATIVA(
      fake,
      chave,
      { teste: true }
    );

    const sucesso = DCP_CV242_ATUALIZAR_TENTATIVA(
      fake,
      linha,
      chave,
      'INCONCLUSIVO',
      { teste: true }
    );

    const tentativa = DCP_CV242_LOCALIZAR_TENTATIVA(
      fake,
      chave
    );

    const registro = JSON.parse(tentativa.valor);

    return sucesso === true &&
      registro.estado === 'INCONCLUSIVO';
  });

  /*
   * 4. BLOQUEAR CHAVE DIFERENTE
   */
  testar('Bloqueio de atualizacao com chave diferente', function() {
    const fake = new FakeSheet();

    const linha = DCP_CV242_CRIAR_TENTATIVA(
      fake,
      chave,
      { teste: true }
    );

    let bloqueou = false;

    try {
      DCP_CV242_ATUALIZAR_TENTATIVA(
        fake,
        linha,
        'outra-chave-12345678901234567890',
        'CONCLUIDO',
        { teste: true }
      );
    } catch (erro) {
      bloqueou = erro.message.indexOf(
        'nao corresponde a chave esperada'
      ) !== -1;
    }

    const tentativa = DCP_CV242_LOCALIZAR_TENTATIVA(
      fake,
      chave
    );

    const registro = JSON.parse(tentativa.valor);

    return bloqueou && registro.estado === 'PROCESSANDO';
  });

  /*
   * 5. BLOQUEAR REENVIO INCONCLUSIVO
   */
  testar('Bloqueio de reenvio inconclusivo', function() {
    const fake = new FakeSheet();

    const linha = DCP_CV242_CRIAR_TENTATIVA(
      fake,
      chave,
      { teste: true }
    );

    DCP_CV242_ATUALIZAR_TENTATIVA(
      fake,
      linha,
      chave,
      'INCONCLUSIVO',
      { codigoVinculo: 'VIN-0001' }
    );

    const resposta = DCP_CV242_TRATAR_REENVIO(
      fake,
      chave
    );

    return resposta !== null &&
      resposta.sucesso === false &&
      resposta.tentativaInconclusiva === true;
  });

  /*
   * 6. REENVIO CONCLUIDO
   *
   * O codigo VIN precisa obedecer ao formato real:
   * VIN-0001, VIN-0002 etc.
   */
  testar('Reenvio concluido devolve resposta anterior', function() {
    const fake = new FakeSheet();

    const linha = DCP_CV242_CRIAR_TENTATIVA(
      fake,
      chave,
      { teste: true }
    );

    const respostaOriginal = {
      sucesso: true,
      codigoVinculo: 'VIN-0002',
      vinculoID: 'UUID-TESTE-0002',
      mensagem: 'Cadastro confirmado.'
    };

    DCP_CV242_ATUALIZAR_TENTATIVA(
      fake,
      linha,
      chave,
      'CONCLUIDO',
      respostaOriginal
    );

    const antes = fake.getLastRow();

    const resposta = DCP_CV242_TRATAR_REENVIO(
      fake,
      chave
    );

    const depois = fake.getLastRow();

    return resposta !== null &&
      resposta.sucesso === true &&
      resposta.reenvioIdentificado === true &&
      resposta.codigoVinculo === 'VIN-0002' &&
      resposta.vinculoID === 'UUID-TESTE-0002' &&
      antes === depois;
  });

  /*
   * 7. CHAVE INEXISTENTE
   */
  testar('Chave inexistente permanece nao localizada', function() {
    const fake = new FakeSheet();

    const tentativa = DCP_CV242_RECUPERAR_TENTATIVA(
      fake,
      chave
    );

    return tentativa.estado === 'nao_localizada';
  });

  /*
   * 8. REGISTRO INCOMPLETO
   *
   * CriadoPor esta vazio intencionalmente.
   * O registro deve ser rejeitado pela conferencia integral.
   */
  testar('Registro incompleto nao e considerado completo', function() {
    const agora = new Date();
    const inicio = new Date(2026, 10, 10, 6, 0);
    const fim = new Date(2026, 10, 10, 14, 0);

    const esperado = {
      vinculoID: 'UUID-TESTE',
      codigo: 'VIN-0002',
      motoristaID: 'MOT-TESTE',
      veiculoID: 'VEI-TESTE',
      tipo: 'PROPRIO',
      inicio: inicio,
      fim: fim,
      compartilhamentoAutorizado: false,
      valorAluguel: '',
      periodicidade: '',
      observacoes: '',
      criadoEm: agora
    };

    const registroIncompleto = {
      linha: 2,
      dados: [
        'UUID-TESTE',
        'VIN-0002',
        'MOT-TESTE',
        'VEI-TESTE',
        'PROPRIO',
        inicio,
        fim,
        false,
        '',
        '',
        '',
        'ATIVO',
        agora,
        agora,
        '',
        'ADMIN_CENTRAL'
      ]
    };

    return DCP_CV242_REGISTRO_COMPLETO(
      registroIncompleto,
      esperado
    ) === false;
  });

  /*
   * 9. REGISTRO COMPLETO
   */
  testar('Registro completo passa na conferencia', function() {
    const inicio = new Date(2026, 10, 10, 6, 0);
    const fim = new Date(2026, 10, 10, 14, 0);
    const agora = new Date();

    const esperado = {
      vinculoID: 'UUID-TESTE',
      codigo: 'VIN-0002',
      motoristaID: 'MOT-TESTE',
      veiculoID: 'VEI-TESTE',
      tipo: 'PROPRIO',
      inicio: inicio,
      fim: fim,
      compartilhamentoAutorizado: true,
      valorAluguel: '',
      periodicidade: '',
      observacoes: 'Teste integral',
      criadoEm: agora
    };

    const registro = {
      linha: 2,
      dados: [
        'UUID-TESTE',
        'VIN-0002',
        'MOT-TESTE',
        'VEI-TESTE',
        'PROPRIO',
        inicio,
        fim,
        true,
        '',
        '',
        'Teste integral',
        'ATIVO',
        agora,
        agora,
        'ADMIN_CENTRAL',
        'ADMIN_CENTRAL'
      ]
    };

    return DCP_CV242_REGISTRO_COMPLETO(
      registro,
      esperado
    ) === true;
  });

  /*
   * 10. DIFERENCA EM OBSERVACOES
   */
  testar('Diferenca de observacoes invalida recuperacao', function() {
    const inicio = new Date(2026, 10, 10, 6, 0);
    const fim = new Date(2026, 10, 10, 14, 0);
    const agora = new Date();

    const esperado = {
      vinculoID: 'UUID-TESTE',
      codigo: 'VIN-0002',
      motoristaID: 'MOT-TESTE',
      veiculoID: 'VEI-TESTE',
      tipo: 'PROPRIO',
      inicio: inicio,
      fim: fim,
      compartilhamentoAutorizado: true,
      valorAluguel: '',
      periodicidade: '',
      observacoes: 'Observacao esperada',
      criadoEm: agora
    };

    const registro = {
      linha: 2,
      dados: [
        'UUID-TESTE',
        'VIN-0002',
        'MOT-TESTE',
        'VEI-TESTE',
        'PROPRIO',
        inicio,
        fim,
        true,
        '',
        '',
        'Observacao diferente',
        'ATIVO',
        agora,
        agora,
        'ADMIN_CENTRAL',
        'ADMIN_CENTRAL'
      ]
    };

    return DCP_CV242_REGISTRO_COMPLETO(
      registro,
      esperado
    ) === false;
  });

  /*
   * RESULTADO FINAL
   */
  const aprovados = resultados.filter(function(item) {
    return item.aprovado;
  }).length;

  const resposta = {
    total: resultados.length,
    aprovados: aprovados,
    reprovados: resultados.length - aprovados,
    resultados: resultados,
    somenteMemoria: true,
    planilhasReaisAlteradas: false,
    sequenciaVINAlterada: false,
    registrosReaisAlterados: false
  };

  Logger.log(JSON.stringify(resposta, null, 2));
  return resposta;
}
