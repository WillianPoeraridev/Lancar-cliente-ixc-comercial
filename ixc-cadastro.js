import CDP from 'chrome-remote-interface'

const CONFIG = {
  host: 'localhost',
  port: 9222,
  ixcUrl: 'sistema.fenixwireless.com.br',
  opaUrl: 'opasuite.fenixwireless.com.br',
}

// ── Helpers globais ──

async function findTab(urlPart) {
  const resp = await fetch('http://' + CONFIG.host + ':' + CONFIG.port + '/json')
  const tabs = await resp.json()
  return tabs.find(t => t.url && t.url.includes(urlPart))
}

async function connectTab(tab) {
  await CDP.Activate({ id: tab.id, host: CONFIG.host, port: CONFIG.port })
  await new Promise(r => setTimeout(r, 300))
  const client = await CDP({ host: CONFIG.host, port: CONFIG.port, target: tab.id })
  const { Runtime, Page } = client
  await Page.enable()

  const exec = (expression) => Runtime.evaluate({ expression, awaitPromise: true })

  async function waitFor(selector, timeout) {
    timeout = timeout || 10000
    const inicio = Date.now()
    while (Date.now() - inicio < timeout) {
      const result = await exec("!!document.querySelector('" + selector + "')")
      if (result.result && result.result.value === true) return true
      await new Promise(r => setTimeout(r, 200))
    }
    console.warn('⚠️ Timeout esperando: ' + selector)
    return false
  }

  return { client, exec, waitFor }
}

// ══════════════════════════════════════════
//  FASE 1 — EXTRAIR DADOS DO OPA
// ══════════════════════════════════════════

async function extrairDadosOPA() {
  console.log('🔍 Conectando no OPA...')

  const tab = await findTab(CONFIG.opaUrl)
  if (!tab) {
    console.error('❌ Aba do OPA não encontrada.')
    process.exit(1)
  }

  console.log('✅ OPA encontrado:', tab.url)
  const { client, exec, waitFor } = await connectTab(tab)

  // Clica em "Ver todas" pra abrir o modal completo
  console.log('📋 Abrindo observações...')
  await waitFor('button.observacao-btn-listar')
  await exec("document.querySelector('button.observacao-btn-listar').click()")
  await new Promise(r => setTimeout(r, 2000))

  // Aguarda o modal carregar e extrai o texto da PRIMEIRA observação (mais recente)
  console.log('📋 Extraindo dados da observação...')
  const textoObs = await (async function() {
    const inicio = Date.now()
    while (Date.now() - inicio < 10000) {
      const result = await exec(
        "(function() {" +
        "  var msgs = document.querySelectorAll('div.corpo-observacao-lista div.observacao_mensagem');" +
        "  if (msgs.length === 0) msgs = document.querySelectorAll('div.observacao_mensagem');" +
        "  if (msgs.length > 0) {" +
        "    var texto = msgs[0].innerText.trim();" +
        "    if (texto.length > 50) return texto;" +
        "  }" +
        "  return '';" +
        "})()"
      )
      if (result.result && result.result.value) return result.result.value
      await new Promise(r => setTimeout(r, 300))
    }
    return ''
  })()

  if (!textoObs) {
    console.error('❌ Não consegui extrair a observação do OPA.')
    await client.close()
    process.exit(1)
  }

  console.log('✅ Observação extraída!')
  console.log(textoObs)

  // Parseia cada linha "Label: Valor"
  var dados = {}
  var linhas = textoObs.split('\n')
  for (var i = 0; i < linhas.length; i++) {
    var linha = linhas[i]
    var idx = linha.indexOf(':')
    if (idx > -1) {
      var chave = linha.substring(0, idx).trim().toLowerCase()
      var valor = linha.substring(idx + 1).trim()
      dados[chave] = valor
    }
  }

  console.log('📦 Dados parseados:', dados)

  await client.close()

  // Limpa CPF e CEP (remove pontos, traços, espaços)
  var cpfLimpo = (dados['cpf'] || '').replace(/\D/g, '')
  var cepLimpo = (dados['cep'] || '').replace(/\D/g, '')

  return {
    nome: dados['nome'] || '',
    cpf: cpfLimpo,
    dataNascimento: dados['data de nascimento'] || '',
    email: dados['e-mail'] || '',
    celular: dados['celular'] || '',
    vencimento: dados['dia de vencimento preferido'] || '',
    planoVendas: dados['plano escolhido'] || '',
    cep: cepLimpo,
    endereco: dados['endereço'] || dados['endereco'] || '',
    numero: dados['número'] || dados['numero'] || '',
    bairro: dados['bairro'] || '',
    cto: dados['cto'] || '',
    distancia: dados['distância'] || dados['distancia'] || '',
    vendaDia: dados['dia da venda'] || '',
    instalacaoDia: dados['dia da instalação'] || dados['dia da instalacao'] || '',
    turno: dados['turno'] || '',
    roteadorExtra: dados['roteador extra'] || 'Não',
    assuntoOS: dados['assunto os'] || '1',
  }
}

// ══════════════════════════════════════════
//  FASE 2 — CADASTRO NO IXC
// ══════════════════════════════════════════

async function cadastrarNoIXC(DADOS) {
  console.log('🤖 Conectando no IXC...')

  const tab = await findTab(CONFIG.ixcUrl)
  if (!tab) {
    console.error('❌ Aba do IXC não encontrada.')
    process.exit(1)
  }

  console.log('✅ IXC encontrado:', tab.url)
  const { client, exec, waitFor } = await connectTab(tab)

  // Aceita dialog nativo automaticamente
  client.on('Page.javascriptDialogOpening', async () => {
    await client.send('Page.handleJavaScriptDialog', { accept: true })
  })

  // Recarrega
  console.log('🔄 Recarregando IXC...')
  await client.send('Page.reload')
  await new Promise(r => setTimeout(r, 3000))

  // ── CADASTRO DO CLIENTE ──

  // 1. Clica em Cadastros
  console.log('📂 Abrindo Cadastros...')
  await waitFor('div.submenu_title a')
  await exec("document.querySelector('div.submenu_title a').click()")
  await new Promise(r => setTimeout(r, 800))

  // 2. Clica em Clientes
  console.log('👤 Abrindo Clientes...')
  await waitFor('li#menu_item_cliente')
  await exec("document.querySelector('li#menu_item_cliente a').click()")
  await waitFor('button[name=\"novo\"]')

  // 3. Clica em Novo
  console.log('➕ Clicando em Novo...')
  await exec("document.querySelector('button[name=\"novo\"]').click()")
  await waitFor('input#id_tipo_cliente')

  // 4. Tipo de cliente = 1
  console.log('📝 Tipo de cliente = 1...')
  await exec("(function() { var i = document.querySelector('input#id_tipo_cliente'); i.value = '1'; i.dispatchEvent(new Event('input', {bubbles:true})); i.dispatchEvent(new Event('change', {bubbles:true})); })()")
  await new Promise(r => setTimeout(r, 500))

  // 5. Tipo de cliente fiscal = 03
  console.log('📝 Tipo de cliente fiscal = 03...')
  await waitFor('select#tipo_cliente_scm')
  await exec("(function() { var s = document.querySelector('select#tipo_cliente_scm'); s.value = '03'; s.dispatchEvent(new Event('change', {bubbles:true})); })()")

  // 6. CPF
  console.log('📝 Preenchendo CPF...')
  await waitFor('input#cnpj_cpf')
  await exec("(function() { var i = document.querySelector('input#cnpj_cpf'); i.value = '" + DADOS.cpf + "'; i.dispatchEvent(new Event('input', {bubbles:true})); i.dispatchEvent(new Event('change', {bubbles:true})); })()")
  await new Promise(r => setTimeout(r, 500))

  // 7. Data de nascimento
  console.log('📝 Preenchendo data de nascimento...')
  await waitFor('input#data_nascimento')
  await exec("(function() { var i = document.querySelector('input#data_nascimento'); i.value = '" + DADOS.dataNascimento + "'; i.dispatchEvent(new Event('input', {bubbles:true})); i.dispatchEvent(new Event('change', {bubbles:true})); })()")

  // 8. Nome
  console.log('📝 Preenchendo nome...')
  await waitFor('input#razao')
  await exec("(function() { var i = document.querySelector('input#razao'); i.value = '" + DADOS.nome + "'; i.dispatchEvent(new Event('input', {bubbles:true})); i.dispatchEvent(new Event('change', {bubbles:true})); })()")

  // ── Aba Endereço ──
  console.log('📍 Abrindo aba Endereço...')
  await waitFor('a.tabTitle[rel=\"1\"]')
  await exec("document.querySelector('a.tabTitle[rel=\"1\"]').click()")
  await waitFor('input#cep')

  // CEP
  console.log('📝 Preenchendo CEP...')
  await exec("(function() { var i = document.querySelector('input#cep'); i.value = '" + DADOS.cep + "'; i.dispatchEvent(new Event('input', {bubbles:true})); i.dispatchEvent(new Event('change', {bubbles:true})); })()")
  await new Promise(r => setTimeout(r, 300))

  // Validar CEP
  console.log('🔍 Validando CEP...')
  await exec("document.querySelector('button#buscacep').click()")

  // Aguarda cidade
  console.log('⏳ Aguardando cidade...')
  var cidadeOk = await (async function() {
    var inicio = Date.now()
    while (Date.now() - inicio < 10000) {
      var result = await exec("(function() { var el = document.querySelector('input#cidade_label'); return el && el.value && el.value.trim().length > 0 ? el.value : ''; })()")
      if (result.result && result.result.value) {
        console.log('✅ Cidade encontrada:', result.result.value)
        return true
      }
      await new Promise(r => setTimeout(r, 300))
    }
    return false
  })()

  if (!cidadeOk) {
    await exec("alert('⚠️ CEP não encontrado! Verifique o CEP e tente novamente.')")
    console.error('❌ CEP inválido ou não encontrado.')
    await client.close()
    return
  }

  // Endereço
  console.log('📝 Preenchendo endereço...')
  await exec("(function() { var i = document.querySelector('input#endereco'); i.value = '" + DADOS.endereco + "'; i.dispatchEvent(new Event('input', {bubbles:true})); i.dispatchEvent(new Event('change', {bubbles:true})); })()")

  // Número
  console.log('📝 Preenchendo número...')
  await exec("(function() { var i = document.querySelector('input#numero'); i.value = '" + DADOS.numero + "'; i.dispatchEvent(new Event('input', {bubbles:true})); i.dispatchEvent(new Event('change', {bubbles:true})); })()")

  // Bairro
  console.log('📝 Preenchendo bairro...')
  await exec("(function() { var i = document.querySelector('input#bairro'); i.value = '" + DADOS.bairro + "'; i.dispatchEvent(new Event('input', {bubbles:true})); i.dispatchEvent(new Event('change', {bubbles:true})); })()")

  // ── Aba Contato ──
  console.log('📞 Abrindo aba Contato...')
  await waitFor('a.tabTitle[rel=\"2\"]')
  await exec("document.querySelector('a.tabTitle[rel=\"2\"]').click()")
  await waitFor('input#telefone_celular')

  console.log('📝 Preenchendo telefone celular...')
  await exec("(function() { var i = document.querySelector('input#telefone_celular'); i.value = '" + DADOS.celular + "'; i.dispatchEvent(new Event('input', {bubbles:true})); i.dispatchEvent(new Event('change', {bubbles:true})); })()")

  console.log('📝 Preenchendo WhatsApp...')
  await exec("(function() { var i = document.querySelector('input#whatsapp'); i.value = '" + DADOS.celular + "'; i.dispatchEvent(new Event('input', {bubbles:true})); i.dispatchEvent(new Event('change', {bubbles:true})); })()")

  console.log('📝 Preenchendo email...')
  await exec("(function() { var i = document.querySelector('input#email'); i.value = '" + DADOS.email + "'; i.dispatchEvent(new Event('input', {bubbles:true})); i.dispatchEvent(new Event('change', {bubbles:true})); })()")

  // Login e Senha = CPF
  console.log('📝 Preenchendo login...')
  await exec("(function() { var i = document.querySelector('input#hotsite_email'); i.value = '" + DADOS.cpf + "'; i.dispatchEvent(new Event('input', {bubbles:true})); i.dispatchEvent(new Event('change', {bubbles:true})); })()")

  console.log('📝 Preenchendo senha...')
  await exec("(function() { var i = document.querySelector('input#senha'); i.value = '" + DADOS.cpf + "'; i.dispatchEvent(new Event('input', {bubbles:true})); i.dispatchEvent(new Event('change', {bubbles:true})); })()")

  // ── Aba Contratos ──
  console.log('📋 Abrindo aba Contratos...')
  await waitFor('a.tabTitle[rel=\"7\"]')
  await exec("document.querySelector('a.tabTitle[rel=\"7\"]').click()")
  await new Promise(r => setTimeout(r, 800))

  // Salvar cliente
  console.log('💾 Salvando cliente...')
  await exec("document.querySelector('button[type=\"submit\"][title=\"Alt+S\"]').click()")

  console.log('⏳ Aguardando confirmação de salvo...')
  var salvou = await waitForNotification(exec)

  if (!salvou) {
    console.warn('⚠️ Não detectei confirmação — esperando 5s...')
    await new Promise(r => setTimeout(r, 5000))
  }

  // ── CONTRATO ──

  console.log('➕ Clicando em Novo contrato...')
  await new Promise(r => setTimeout(r, 1000))
  // IMPORTANTE: pega o Novo da GRID de contratos (div.tDiv2/gridActions), NÃO o do formulário principal
  await exec(
    "(function() {" +
    "  var btn = null;" +
    "  var gridBtns = document.querySelectorAll('div.tDiv2 button[name=\"novo\"], div.gridActions button[name=\"novo\"]');" +
    "  for (var i = 0; i < gridBtns.length; i++) {" +
    "    if (gridBtns[i].offsetParent !== null && gridBtns[i].id !== 'novo_form') {" +
    "      btn = gridBtns[i]; break;" +
    "    }" +
    "  }" +
    "  if (btn) { console.log('Clicando em:', btn.textContent.trim()); btn.click(); }" +
    "  else console.warn('Botão Novo contrato não encontrado na grid');" +
    "})()"
  )
  await new Promise(r => setTimeout(r, 2000))

  // Debug: verifica se o formulário do contrato abriu
  var formAbriu = await exec("(function() { var el = document.querySelector('input#id_vd_contrato'); return el ? 'sim' : 'não'; })()")
  console.log('🔍 Formulário do contrato abriu?', formAbriu.result ? formAbriu.result.value : 'erro')

  // Verifica cliente preenchido
  console.log('🔍 Verificando campo cliente...')
  var clienteOk = await (async function() {
    var inicio = Date.now()
    while (Date.now() - inicio < 5000) {
      var result = await exec("(function() { var el = document.querySelector('input#id_cliente'); return el && el.value && el.value.trim().length > 0 ? el.value : ''; })()")
      if (result.result && result.result.value) return true
      await new Promise(r => setTimeout(r, 300))
    }
    return false
  })()

  if (!clienteOk) {
    await exec("alert('⚠️ Campo cliente não preenchido! Selecione manualmente e clique OK.')")
  }

  // Plano de vendas
  console.log('📝 Preenchendo plano de vendas...')
  await waitFor('input#id_vd_contrato', 15000)
  await exec("(function() { var i = document.querySelector('input#id_vd_contrato'); i.value = '" + DADOS.planoVendas + "'; i.dispatchEvent(new Event('input', {bubbles:true})); i.dispatchEvent(new Event('change', {bubbles:true})); i.dispatchEvent(new KeyboardEvent('keydown', {keyCode:13, bubbles:true})); })()")
  await new Promise(r => setTimeout(r, 500))

  // Vencimento
  console.log('📝 Preenchendo vencimento...')
  await waitFor('input#id_tipo_contrato')
  await exec("(function() { var i = document.querySelector('input#id_tipo_contrato'); i.value = '" + DADOS.vencimento + "'; i.dispatchEvent(new Event('input', {bubbles:true})); i.dispatchEvent(new Event('change', {bubbles:true})); i.dispatchEvent(new KeyboardEvent('keydown', {keyCode:13, bubbles:true})); })()")
  await new Promise(r => setTimeout(r, 500))

  // Motivo de inclusão — sempre 1
  console.log('📝 Preenchendo motivo de inclusão...')
  await waitFor('input#id_motivo_inclusao')
  await exec("(function() { var i = document.querySelector('input#id_motivo_inclusao'); i.value = '1'; i.dispatchEvent(new Event('input', {bubbles:true})); i.dispatchEvent(new Event('change', {bubbles:true})); i.dispatchEvent(new KeyboardEvent('keydown', {keyCode:13, bubbles:true})); })()")
  await new Promise(r => setTimeout(r, 500))

  // Carteira de cobrança — sempre 7
  console.log('📝 Preenchendo carteira de cobrança...')
  await waitFor('input#id_carteira_cobranca')
  await exec("(function() { var i = document.querySelector('input#id_carteira_cobranca'); i.value = '7'; i.dispatchEvent(new Event('input', {bubbles:true})); i.dispatchEvent(new Event('change', {bubbles:true})); i.dispatchEvent(new KeyboardEvent('keydown', {keyCode:13, bubbles:true})); })()")
  await new Promise(r => setTimeout(r, 500))

  // Salvar contrato
  console.log('💾 Salvando contrato...')
  await exec("document.querySelector('button[type=\"submit\"][title=\"Alt+S\"]').click()")

  console.log('⏳ Aguardando confirmação do contrato...')
  var salvouContrato = await waitForNotification(exec)
  if (!salvouContrato) {
    console.warn('⚠️ Não detectei confirmação do contrato — esperando 5s...')
    await new Promise(r => setTimeout(r, 5000))
  }

  // Fechar contrato (ESC)
  console.log('🔙 Fechando formulário do contrato...')
  await exec("document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', code: 'Escape', keyCode: 27, bubbles: true }))")
  await new Promise(r => setTimeout(r, 500))

  // ── LOGIN ──

  // Aba Logins
  console.log('🔐 Abrindo aba Logins...')
  await waitFor('a.tabTitle[rel=\"10\"]')
  await exec("document.querySelector('a.tabTitle[rel=\"10\"]').click()")
  await new Promise(r => setTimeout(r, 1000))

  // Novo login
  console.log('➕ Clicando em Novo login...')
  await waitFor('div.gridActions button[name=\"novo\"]')
  await exec("document.querySelector('div.gridActions button[name=\"novo\"]').click()")
  await new Promise(r => setTimeout(r, 1500))

  // Lupa do contrato
  console.log('🔍 Abrindo busca de contrato...')
  await waitFor('button[rel=\"id_contrato\"].but_search')
  await exec("document.querySelector('button[rel=\"id_contrato\"].but_search').click()")
  await new Promise(r => setTimeout(r, 1500))

  // Seleciona Pré-contrato (double-click)
  console.log('📋 Selecionando contrato Pré-contrato...')
  var contratoSelecionado = await (async function() {
    var inicio = Date.now()
    while (Date.now() - inicio < 10000) {
      var result = await exec(
        "(function() {" +
        "  var badge = document.querySelector('span.vg-badge-content');" +
        "  if (badge && badge.textContent.trim() === 'Pré-contrato') {" +
        "    var row = badge.closest('tr');" +
        "    if (row) {" +
        "      row.dispatchEvent(new MouseEvent('dblclick', {bubbles:true}));" +
        "      var idCell = row.querySelector('td:first-child');" +
        "      return idCell ? idCell.textContent.trim() : 'selecionado';" +
        "    }" +
        "  }" +
        "  return '';" +
        "})()"
      )
      if (result.result && result.result.value) {
        console.log('✅ Contrato selecionado:', result.result.value)
        return true
      }
      await new Promise(r => setTimeout(r, 300))
    }
    return false
  })()

  if (!contratoSelecionado) {
    await exec("alert('⚠️ Contrato Pré-contrato não encontrado! Selecione manualmente.')")
  }

  await new Promise(r => setTimeout(r, 1000))

  // Captura número do contrato
  console.log('📋 Capturando número do contrato...')
  var numeroContrato = await (async function() {
    var inicio = Date.now()
    while (Date.now() - inicio < 5000) {
      var result = await exec("(function() { var el = document.querySelector('input#id_contrato'); return el && el.value && el.value.trim().length > 0 ? el.value : ''; })()")
      if (result.result && result.result.value) return result.result.value
      await new Promise(r => setTimeout(r, 300))
    }
    return ''
  })()

  if (numeroContrato) {
    console.log('✅ Número do contrato:', numeroContrato)
  } else {
    console.warn('⚠️ Não consegui capturar o número do contrato.')
  }

  // Lupa do plano (grupo)
  console.log('🔍 Abrindo busca de plano...')
  await waitFor('button[rel=\"id_grupo\"].but_search')
  await exec(
    "(function() {" +
    "  var btns = document.querySelectorAll('button[rel=\"id_grupo\"].but_search');" +
    "  var btn = Array.from(btns).find(function(b) { return b.style.marginLeft || b.classList.contains('last-visible'); }) || btns[btns.length - 1];" +
    "  btn.click();" +
    "})()"
  )
  await new Promise(r => setTimeout(r, 1500))

  // Atualizar pra carregar o plano
  console.log('🔄 Atualizando listagem de planos...')
  await waitFor('i.fa-arrows-rotate[title=\"Atualizar\"]')
  await exec("document.querySelector('i.fa-arrows-rotate[title=\"Atualizar\"]').click()")
  await new Promise(r => setTimeout(r, 1500))

  // Double-click no plano
  console.log('📋 Selecionando plano...')
  var planoSelecionado = await (async function() {
    var inicio = Date.now()
    while (Date.now() - inicio < 10000) {
      var result = await exec(
        "(function() {" +
        "  var row = document.querySelector('tbody tr.tableRow');" +
        "  if (row) {" +
        "    row.dispatchEvent(new MouseEvent('dblclick', {bubbles:true}));" +
        "    var id = row.getAttribute('data-valorcampoautoincrement');" +
        "    return id || 'selecionado';" +
        "  }" +
        "  return '';" +
        "})()"
      )
      if (result.result && result.result.value) {
        console.log('✅ Plano selecionado:', result.result.value)
        return true
      }
      await new Promise(r => setTimeout(r, 300))
    }
    return false
  })()

  if (!planoSelecionado) {
    await exec("alert('⚠️ Plano não encontrado! Selecione manualmente.')")
  }

  await new Promise(r => setTimeout(r, 1000))

  // Login = número do contrato
  console.log('📝 Preenchendo login com número do contrato...')
  await waitFor('input#login')
  await exec("(function() { var i = document.querySelector('input#login'); i.value = '" + numeroContrato + "'; i.dispatchEvent(new Event('input', {bubbles:true})); i.dispatchEvent(new Event('change', {bubbles:true})); })()")

  // Senha PPPoE = número do contrato
  console.log('📝 Preenchendo senha PPPoE...')
  await waitFor('input#senha')
  await exec("(function() { var i = document.querySelector('input#senha'); i.value = '" + numeroContrato + "'; i.dispatchEvent(new Event('input', {bubbles:true})); i.dispatchEvent(new Event('change', {bubbles:true})); })()")

  // ── DADOS TÉCNICOS ──

  // Aba Dados técnicos
  console.log('🔧 Abrindo aba Dados técnicos...')
  await exec(
    "(function() {" +
    "  var tabs = document.querySelectorAll('a.tabTitle');" +
    "  var tab = Array.from(tabs).find(function(t) { return t.textContent.trim() === 'Dados técnicos'; });" +
    "  if (tab) tab.click();" +
    "})()"
  )
  await new Promise(r => setTimeout(r, 1000))

  // Lupa da caixa de atendimento
  console.log('🔍 Abrindo busca de caixa de atendimento...')
  await waitFor('button[rel=\"id_caixa_ftth\"].but_search')
  await exec(
    "(function() {" +
    "  var btns = document.querySelectorAll('button[rel=\"id_caixa_ftth\"].but_search');" +
    "  var btn = Array.from(btns).find(function(b) { return b.classList.contains('last-visible'); }) || btns[btns.length - 1];" +
    "  btn.click();" +
    "})()"
  )
  await new Promise(r => setTimeout(r, 1500))

  // Pesquisa a caixa (CTO do OPA)
  console.log('🔍 Pesquisando caixa:', DADOS.cto)
  await waitFor('input.gridActionsSearchInput')
  await exec("(function() { var i = document.querySelector('input.gridActionsSearchInput'); i.value = '" + DADOS.cto + "'; i.dispatchEvent(new Event('input', {bubbles:true})); i.dispatchEvent(new KeyboardEvent('keydown', {key:'Enter', keyCode:13, bubbles:true})); })()")
  await new Promise(r => setTimeout(r, 2000))

  // Double-click na caixa
  console.log('📋 Selecionando caixa...')
  var caixaSelecionada = await (async function() {
    var inicio = Date.now()
    while (Date.now() - inicio < 10000) {
      var result = await exec(
        "(function() {" +
        "  var row = document.querySelector('tbody tr.tableRow');" +
        "  if (row) {" +
        "    row.dispatchEvent(new MouseEvent('dblclick', {bubbles:true}));" +
        "    var id = row.getAttribute('data-valorcampoautoincrement');" +
        "    return id || 'selecionado';" +
        "  }" +
        "  return '';" +
        "})()"
      )
      if (result.result && result.result.value) {
        console.log('✅ Caixa selecionada:', result.result.value)
        return true
      }
      await new Promise(r => setTimeout(r, 300))
    }
    return false
  })()

  if (!caixaSelecionada) {
    await exec("alert('⚠️ Caixa não encontrada! Selecione manualmente.')")
  }

  await new Promise(r => setTimeout(r, 1000))

  // Listar portas
  console.log('📋 Listando portas de atendimento...')
  await exec(
    "(function() {" +
    "  var btns = document.querySelectorAll('button');" +
    "  var listar = Array.from(btns).find(function(b) { return b.textContent.trim() === 'Listar'; });" +
    "  if (listar) listar.click();" +
    "})()"
  )
  await new Promise(r => setTimeout(r, 2000))

  // Seleciona primeira porta Disponível
  console.log('🔍 Procurando porta disponível...')
  var portaAtendimento = ''
  var portaSelecionada = await (async function() {
    var inicio = Date.now()
    while (Date.now() - inicio < 10000) {
      var result = await exec(
        "(function() {" +
        "  var rows = document.querySelectorAll('tbody tr');" +
        "  for (var j = 0; j < rows.length; j++) {" +
        "    var row = rows[j];" +
        "    var badge = row.querySelector('span.vg-badge-content');" +
        "    if (badge && badge.textContent.trim() === 'Disponível') {" +
        "      var cells = row.querySelectorAll('td');" +
        "      var porta = cells[0] ? cells[0].textContent.trim() : '';" +
        "      row.click();" +
        "      return porta || 'disponível';" +
        "    }" +
        "  }" +
        "  return '';" +
        "})()"
      )
      if (result.result && result.result.value) {
        portaAtendimento = result.result.value
        console.log('✅ Porta disponível:', portaAtendimento)
        return true
      }
      await new Promise(r => setTimeout(r, 300))
    }
    return false
  })()

  if (!portaSelecionada) {
    await exec("alert('⚠️ Nenhuma porta disponível! Selecione manualmente.')")
  }

  // Confirmar porta
  console.log('✅ Confirmando porta...')
  await exec(
    "(function() {" +
    "  var btns = document.querySelectorAll('button');" +
    "  var sel = Array.from(btns).find(function(b) { return b.textContent.trim() === 'Selecionar'; });" +
    "  if (sel) sel.click();" +
    "})()"
  )
  await new Promise(r => setTimeout(r, 1000))

  // ── SALVAR LOGIN ──

  console.log('💾 Salvando login...')
  await exec("document.querySelector('button[type=\"submit\"][title=\"Alt+S\"]').click()")

  console.log('⏳ Aguardando confirmação do login...')
  var salvouLogin = await waitForNotification(exec)
  if (!salvouLogin) {
    console.warn('⚠️ Não detectei confirmação do login — esperando 5s...')
    await new Promise(r => setTimeout(r, 5000))
  }

  // ── O.S. ──

  console.log('📋 Abrindo aba O.S...')
  await waitFor('a.tabTitle[rel=\"13\"]')
  await exec("document.querySelector('a.tabTitle[rel=\"13\"]').click()")
  await new Promise(r => setTimeout(r, 1000))

  // Nova OS
  console.log('➕ Clicando em Nova O.S...')
  await waitFor('button[name=\"novo\"]')
  await exec(
    "(function() {" +
    "  var btns = document.querySelectorAll('button[name=\"novo\"]');" +
    "  var btn = Array.from(btns).find(function(b) { return b.textContent.trim() === 'Nova'; });" +
    "  if (btn) btn.click();" +
    "})()"
  )
  await new Promise(r => setTimeout(r, 1500))

  // Assunto
  console.log('📝 Preenchendo assunto da OS...')
  await waitFor('input#id_assunto')
  await exec("(function() { var i = document.querySelector('input#id_assunto'); i.value = '" + DADOS.assuntoOS + "'; i.dispatchEvent(new Event('input', {bubbles:true})); i.dispatchEvent(new Event('change', {bubbles:true})); i.dispatchEvent(new KeyboardEvent('keydown', {keyCode:13, bubbles:true})); })()")
  await new Promise(r => setTimeout(r, 500))

  // Setor — sempre 1
  console.log('📝 Preenchendo setor...')
  await exec("(function() { var i = document.querySelector('input#id_setor') || document.querySelector('input[name=\"id_setor\"]'); if (i) { i.value = '1'; i.dispatchEvent(new Event('input', {bubbles:true})); i.dispatchEvent(new Event('change', {bubbles:true})); i.dispatchEvent(new KeyboardEvent('keydown', {keyCode:13, bubbles:true})); } })()")
  await new Promise(r => setTimeout(r, 500))

  // Descrição da OS
  var descricaoOS = 'Vendedor: WILLIAN POERARI\n'
    + 'Caixa / Metragem / Porta: ' + DADOS.cto + ' - ' + DADOS.distancia + ' - P' + portaAtendimento + '\n'
    + 'Plano Escolhido: ' + DADOS.planoVendas + '\n'
    + 'Valor de Instalação no Ato: ISENTO\n'
    + 'Venda no Dia: ' + DADOS.vendaDia + '\n'
    + 'Agendado Instalação Para: ' + DADOS.instalacaoDia + '\n'
    + 'Turno: ' + DADOS.turno

  console.log('📝 Preenchendo descrição da OS...')
  await waitFor('textarea#mensagem')
  await exec("(function() { var t = document.querySelector('textarea#mensagem'); t.value = " + JSON.stringify(descricaoOS) + "; t.dispatchEvent(new Event('input', {bubbles:true})); t.dispatchEvent(new Event('change', {bubbles:true})); })()")
  await new Promise(r => setTimeout(r, 500))

  // Salvar OS
  console.log('💾 Salvando OS...')
  await exec("document.querySelector('button[type=\"submit\"][title=\"Alt+S\"]').click()")

  console.log('⏳ Aguardando confirmação da OS...')
  var salvouOS = await waitForNotification(exec)
  if (!salvouOS) {
    console.warn('⚠️ Não detectei confirmação da OS — esperando 5s...')
    await new Promise(r => setTimeout(r, 5000))
  }

  // Popup de sucesso
  await exec("alert('✅ CADASTRO COMPLETO!\\n\\nCliente: " + DADOS.nome + "\\nContrato: " + numeroContrato + "\\nPorta: P" + portaAtendimento + "\\nOS aberta e salva.')")

  console.log('══════════════════════════════════════')
  console.log('✅ FLUXO COMPLETO!')
  console.log('📌 Cliente:', DADOS.nome)
  console.log('📌 Contrato:', numeroContrato)
  console.log('📌 Porta:', portaAtendimento)
  console.log('══════════════════════════════════════')

  await client.close()
}

// ── Helper: aguarda notificação do IXC ──

async function waitForNotification(exec) {
  // Limpa notificações antigas antes de esperar a nova
  await exec("(function() { var els = document.querySelectorAll('div.notification-success'); els.forEach(function(e) { e.remove(); }); })()")
  await new Promise(r => setTimeout(r, 300))

  var inicio = Date.now()
  while (Date.now() - inicio < 15000) {
    var result = await exec("(function() { var el = document.querySelector('div.notification-success'); if (el) { var msg = el.querySelector('div.notificationMessage span'); return msg && msg.innerText ? msg.innerText.trim() : 'sucesso'; } return ''; })()")
    if (result.result && result.result.value) {
      console.log('✅ Resposta IXC:', result.result.value)
      return true
    }
    await new Promise(r => setTimeout(r, 300))
  }
  return false
}

// ── MAIN ──

async function main() {
  console.log('══════════════════════════════════════')
  console.log('🚀 AUTOMAÇÃO OPA → IXC')
  console.log('══════════════════════════════════════')

  // Fase 1: Extrai dados do OPA
  var dados = await extrairDadosOPA()

  console.log('══════════════════════════════════════')
  console.log('📦 Dados extraídos do OPA:')
  console.log('  Nome:', dados.nome)
  console.log('  CPF:', dados.cpf)
  console.log('  Plano:', dados.planoVendas)
  console.log('  CTO:', dados.cto)
  console.log('  Vencimento:', dados.vencimento)
  console.log('══════════════════════════════════════')

  // Fase 2: Cadastra no IXC
  await cadastrarNoIXC(dados)
}

main()