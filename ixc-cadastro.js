import CDP from 'chrome-remote-interface'

const CONFIG = {
  host: 'localhost',
  port: 9222,
  ixcUrl: 'sistema.fenixwireless.com.br',
}

async function getIxcTab() {
  const resp = await fetch(`http://${CONFIG.host}:${CONFIG.port}/json`)
  const tabs = await resp.json()
  return tabs.find(t => t.url && t.url.includes(CONFIG.ixcUrl))
}

async function main() {
  console.log('🤖 Conectando no IXC...')

  const tab = await getIxcTab()
  if (!tab) {
    console.error('❌ Aba do IXC não encontrada.')
    process.exit(1)
  }

  console.log('✅ IXC encontrado:', tab.url)

  await CDP.Activate({ id: tab.id, host: CONFIG.host, port: CONFIG.port })
  await new Promise(r => setTimeout(r, 300))

  const client = await CDP({ host: CONFIG.host, port: CONFIG.port, target: tab.id })
  const { Runtime, Page } = client
  await Page.enable()

  const exec = (expression) => Runtime.evaluate({ expression, awaitPromise: true })

  // Aguarda elemento aparecer no DOM (max 10s)
  async function waitFor(selector, timeout = 10000) {
    const inicio = Date.now()
    while (Date.now() - inicio < timeout) {
      const result = await exec(`!!document.querySelector('${selector}')`)
      if (result.result?.value === true) return true
      await new Promise(r => setTimeout(r, 200))
    }
    console.warn(`⚠️ Timeout esperando: ${selector}`)
    return false
  }

  // Aceita dialog nativo automaticamente
  client.on('Page.javascriptDialogOpening', async () => {
    await client.send('Page.handleJavaScriptDialog', { accept: true })
  })

  // Recarrega
  console.log('🔄 Recarregando IXC...')
  await Page.reload()
  await new Promise(r => setTimeout(r, 3000))

  // 1. Clica em Cadastros
  console.log('📂 Abrindo Cadastros...')
  await waitFor('div.submenu_title a')
  await exec(`document.querySelector('div.submenu_title a').click()`)
  await new Promise(r => setTimeout(r, 800))

  // 2. Clica em Clientes
  console.log('👤 Abrindo Clientes...')
  await waitFor('li#menu_item_cliente')
  await exec(`document.querySelector('li#menu_item_cliente a').click()`)
  await waitFor('button[name="novo"]')

  // 3. Clica em Novo
  console.log('➕ Clicando em Novo...')
  await exec(`document.querySelector('button[name="novo"]').click()`)
  await waitFor('input#id_tipo_cliente')

  // 4. Tipo de cliente = 1
  console.log('📝 Tipo de cliente = 1...')
  await exec(`
    (function() {
      const input = document.querySelector('input#id_tipo_cliente')
      input.value = '1'
      input.dispatchEvent(new Event('input', { bubbles: true }))
      input.dispatchEvent(new Event('change', { bubbles: true }))
    })()
  `)
  await new Promise(r => setTimeout(r, 500))

  // 5. Tipo de cliente fiscal = 03
  console.log('📝 Tipo de cliente fiscal = 03...')
  await waitFor('select#tipo_cliente_scm')
  await exec(`
    (function() {
      const select = document.querySelector('select#tipo_cliente_scm')
      select.value = '03'
      select.dispatchEvent(new Event('change', { bubbles: true }))
    })()
  `)

  // 6. CPF
  console.log('📝 Preenchendo CPF...')
  await waitFor('input#cnpj_cpf')
  await exec(`
    (function() {
      const input = document.querySelector('input#cnpj_cpf')
      input.value = '01234567890'
      input.dispatchEvent(new Event('input', { bubbles: true }))
      input.dispatchEvent(new Event('change', { bubbles: true }))
    })()
  `)
  await new Promise(r => setTimeout(r, 500))

  // 7. Data de nascimento
  console.log('📝 Preenchendo data de nascimento...')
  await waitFor('input#data_nascimento')
  await exec(`
    (function() {
      const input = document.querySelector('input#data_nascimento')
      input.value = '01/01/1990'
      input.dispatchEvent(new Event('input', { bubbles: true }))
      input.dispatchEvent(new Event('change', { bubbles: true }))
    })()
  `)

  // 8. Nome do cliente (Razão Social)
  console.log('📝 Preenchendo nome...')
  await waitFor('input#razao')
  await exec(`
    (function() {
      const input = document.querySelector('input#razao')
      input.value = 'João da Silva Teste'
      input.dispatchEvent(new Event('input', { bubbles: true }))
      input.dispatchEvent(new Event('change', { bubbles: true }))
    })()
  `)

  // 9. Clica na aba Endereço
  console.log('📍 Abrindo aba Endereço...')
  await waitFor('a.tabTitle[rel="1"]')
  await exec(`document.querySelector('a.tabTitle[rel="1"]').click()`)
  await waitFor('input#cep')

  // 10. Preenche CEP
  console.log('📝 Preenchendo CEP...')
  await exec(`
    (function() {
      const input = document.querySelector('input#cep')
      input.value = '93310040'
      input.dispatchEvent(new Event('input', { bubbles: true }))
      input.dispatchEvent(new Event('change', { bubbles: true }))
    })()
  `)
  await new Promise(r => setTimeout(r, 300))

  // 11. Clica em Validar
  console.log('🔍 Validando CEP...')
  await exec(`document.querySelector('button#buscacep').click()`)

  // 12. Aguarda cidade aparecer (sinal que o CEP foi validado)
  console.log('⏳ Aguardando cidade...')
  const cidadeOk = await (async () => {
    const inicio = Date.now()
    while (Date.now() - inicio < 10000) {
      const result = await exec(`
        (function() {
          const el = document.querySelector('input#cidade_label')
          return el && el.value && el.value.trim().length > 0 ? el.value : ''
        })()
      `)
      if (result.result?.value) {
        console.log('✅ Cidade encontrada:', result.result.value)
        return true
      }
      await new Promise(r => setTimeout(r, 300))
    }
    return false
  })()

  if (!cidadeOk) {
    await exec(`alert('⚠️ CEP não encontrado! Verifique o CEP e tente novamente.')`)
    console.error('❌ CEP inválido ou não encontrado.')
    await client.close()
    return
  }

  // 13. Preenche Endereço
  console.log('📝 Preenchendo endereço...')
  await exec(`
    (function() {
      const input = document.querySelector('input#endereco')
      input.value = 'Rua Teste'
      input.dispatchEvent(new Event('input', { bubbles: true }))
      input.dispatchEvent(new Event('change', { bubbles: true }))
    })()
  `)

  // 14. Preenche Número
  console.log('📝 Preenchendo número...')
  await exec(`
    (function() {
      const input = document.querySelector('input#numero')
      input.value = '123'
      input.dispatchEvent(new Event('input', { bubbles: true }))
      input.dispatchEvent(new Event('change', { bubbles: true }))
    })()
  `)

  // 15. Preenche Bairro
  console.log('📝 Preenchendo bairro...')
  await exec(`
    (function() {
      const input = document.querySelector('input#bairro')
      input.value = 'Centro'
      input.dispatchEvent(new Event('input', { bubbles: true }))
      input.dispatchEvent(new Event('change', { bubbles: true }))
    })()
  `)

  // Aba Contato
  console.log('📞 Abrindo aba Contato...')
  await waitFor('a.tabTitle[rel="2"]')
  await exec(`document.querySelector('a.tabTitle[rel="2"]').click()`)
  await waitFor('input#telefone_celular')

  // Telefone celular
  console.log('📝 Preenchendo telefone celular...')
  await exec(`
    (function() {
      const input = document.querySelector('input#telefone_celular')
      input.value = '51999999999'
      input.dispatchEvent(new Event('input', { bubbles: true }))
      input.dispatchEvent(new Event('change', { bubbles: true }))
    })()
  `)

  // WhatsApp
  console.log('📝 Preenchendo WhatsApp...')
  await exec(`
    (function() {
      const input = document.querySelector('input#whatsapp')
      input.value = '51999999999'
      input.dispatchEvent(new Event('input', { bubbles: true }))
      input.dispatchEvent(new Event('change', { bubbles: true }))
    })()
  `)

  // Email
  console.log('📝 Preenchendo email...')
  await exec(`
    (function() {
      const input = document.querySelector('input#email')
      input.value = 'teste@email.com'
      input.dispatchEvent(new Event('input', { bubbles: true }))
      input.dispatchEvent(new Event('change', { bubbles: true }))
    })()
  `)

  // Login (CPF sem pontos)
  console.log('📝 Preenchendo login...')
  await exec(`
    (function() {
      const input = document.querySelector('input#hotsite_email')
      input.value = '01234567890'
      input.dispatchEvent(new Event('input', { bubbles: true }))
      input.dispatchEvent(new Event('change', { bubbles: true }))
    })()
  `)

  // Senha (CPF sem pontos)
  console.log('📝 Preenchendo senha...')
  await exec(`
    (function() {
      const input = document.querySelector('input#senha')
      input.value = '01234567890'
      input.dispatchEvent(new Event('input', { bubbles: true }))
      input.dispatchEvent(new Event('change', { bubbles: true }))
    })()
  `)

  // Aba Contratos
  console.log('📋 Abrindo aba Contratos...')
  await waitFor('a.tabTitle[rel="7"]')
  await exec(`document.querySelector('a.tabTitle[rel="7"]').click()`)
  await new Promise(r => setTimeout(r, 800))

  // Salvar cliente
  console.log('💾 Salvando cliente...')
  await exec(`document.querySelector('button[type="submit"][title="Alt+S"]').click()`)

  // Aguarda confirmação de salvo — lê notificação do IXC
  console.log('⏳ Aguardando confirmação de salvo...')
  const salvou = await (async () => {
    const inicio = Date.now()
    while (Date.now() - inicio < 15000) {
      const result = await exec(`
        (function() {
          const el = document.querySelector('div.notificationMessage span')
          return el && el.innerText ? el.innerText.trim() : ''
        })()
      `)
      if (result.result?.value) {
        console.log('✅ Resposta IXC:', result.result.value)
        return true
      }
      await new Promise(r => setTimeout(r, 300))
    }
    return false
  })()

  if (!salvou) {
    console.warn('⚠️ Não detectei confirmação de salvo — esperando 5s e continuando...')
    await new Promise(r => setTimeout(r, 5000))
  }

  // ══════════════════════════════════════════
  //  PARTE 2 — CONTRATO
  // ══════════════════════════════════════════

  // Novo contrato
  console.log('➕ Clicando em Novo contrato...')
  await waitFor('button[name="novo"].fbutton')
  await exec(`document.querySelector('button[name="novo"].fbutton').click()`)
  await new Promise(r => setTimeout(r, 1000))

  // Verifica se campo cliente veio preenchido
  console.log('🔍 Verificando campo cliente...')
  const clienteOk = await (async () => {
    const inicio = Date.now()
    while (Date.now() - inicio < 5000) {
      const result = await exec(`
        (function() {
          const el = document.querySelector('input#id_cliente')
          return el && el.value && el.value.trim().length > 0 ? el.value : ''
        })()
      `)
      if (result.result?.value) return true
      await new Promise(r => setTimeout(r, 300))
    }
    return false
  })()

  if (!clienteOk) {
    await exec(`alert('⚠️ Campo cliente não preenchido automaticamente! Selecione o cliente manualmente e clique OK para continuar.')`)
    console.warn('⚠️ Cliente não preenchido — aguardando ação manual...')
  }

  // Plano de vendas (código vem do OPA ex: 848, 903, 906)
  console.log('📝 Preenchendo plano de vendas...')
  await waitFor('input#id_vd_contrato')
  await exec(`
    (function() {
      const input = document.querySelector('input#id_vd_contrato')
      input.value = '848'
      input.dispatchEvent(new Event('input', { bubbles: true }))
      input.dispatchEvent(new Event('change', { bubbles: true }))
      input.dispatchEvent(new KeyboardEvent('keydown', { keyCode: 13, bubbles: true }))
    })()
  `)
  await new Promise(r => setTimeout(r, 500))

  // Tipo de contrato / vencimento (dia: 7, 12 ou 16 — vem do OPA)
  console.log('📝 Preenchendo tipo de contrato/vencimento...')
  await waitFor('input#id_tipo_contrato')
  await exec(`
    (function() {
      const input = document.querySelector('input#id_tipo_contrato')
      input.value = '7'
      input.dispatchEvent(new Event('input', { bubbles: true }))
      input.dispatchEvent(new Event('change', { bubbles: true }))
      input.dispatchEvent(new KeyboardEvent('keydown', { keyCode: 13, bubbles: true }))
    })()
  `)
  await new Promise(r => setTimeout(r, 500))

  // Motivo de inclusão — sempre 1
  console.log('📝 Preenchendo motivo de inclusão...')
  await waitFor('input#id_motivo_inclusao')
  await exec(`
    (function() {
      const input = document.querySelector('input#id_motivo_inclusao')
      input.value = '1'
      input.dispatchEvent(new Event('input', { bubbles: true }))
      input.dispatchEvent(new Event('change', { bubbles: true }))
      input.dispatchEvent(new KeyboardEvent('keydown', { keyCode: 13, bubbles: true }))
    })()
  `)
  await new Promise(r => setTimeout(r, 500))

  // Carteira de cobrança — sempre 7
  console.log('📝 Preenchendo carteira de cobrança...')
  await waitFor('input#id_carteira_cobranca')
  await exec(`
    (function() {
      const input = document.querySelector('input#id_carteira_cobranca')
      input.value = '7'
      input.dispatchEvent(new Event('input', { bubbles: true }))
      input.dispatchEvent(new Event('change', { bubbles: true }))
      input.dispatchEvent(new KeyboardEvent('keydown', { keyCode: 13, bubbles: true }))
    })()
  `)
  await new Promise(r => setTimeout(r, 500))

  // Salvar contrato
  console.log('💾 Salvando contrato...')
  await exec(`document.querySelector('button[type="submit"][title="Alt+S"]').click()`)

  // Aguarda confirmação de salvo do contrato
  console.log('⏳ Aguardando confirmação de salvo do contrato...')
  const salvouContrato = await (async () => {
    const inicio = Date.now()
    while (Date.now() - inicio < 15000) {
      const result = await exec(`
        (function() {
          const el = document.querySelector('div.notificationMessage span')
          return el && el.innerText ? el.innerText.trim() : ''
        })()
      `)
      if (result.result?.value) {
        console.log('✅ Resposta IXC:', result.result.value)
        return true
      }
      await new Promise(r => setTimeout(r, 300))
    }
    return false
  })()

  if (!salvouContrato) {
    console.warn('⚠️ Não detectei confirmação do contrato — esperando 5s...')
    await new Promise(r => setTimeout(r, 5000))
  }

  // Fechar formulário do contrato (ESC)
  console.log('🔙 Fechando formulário do contrato...')
  await exec(`
    document.dispatchEvent(new KeyboardEvent('keydown', {
      key: 'Escape', code: 'Escape', keyCode: 27, bubbles: true
    }))
  `)
  await new Promise(r => setTimeout(r, 500))

  // ══════════════════════════════════════════
  //  PARTE 3 — LOGIN
  // ══════════════════════════════════════════

  // Aba Logins
  console.log('🔐 Abrindo aba Logins...')
  await waitFor('a.tabTitle[rel="10"]')
  await exec(`document.querySelector('a.tabTitle[rel="10"]').click()`)
  await new Promise(r => setTimeout(r, 1000))

  // Clica em Novo na grid de logins
  console.log('➕ Clicando em Novo login...')
  await waitFor('div.gridActions button[name="novo"]')
  await exec(`document.querySelector('div.gridActions button[name="novo"]').click()`)
  await new Promise(r => setTimeout(r, 1500))

  // Clica na lupa do campo Contrato para abrir a busca
  console.log('🔍 Abrindo busca de contrato...')
  await waitFor('button[rel="id_contrato"].but_search')
  await exec(`document.querySelector('button[rel="id_contrato"].but_search').click()`)
  await new Promise(r => setTimeout(r, 1500))

  // Seleciona o contrato com status "Pré-contrato"
  console.log('📋 Selecionando contrato Pré-contrato...')
  const contratoSelecionado = await (async () => {
    const inicio = Date.now()
    while (Date.now() - inicio < 10000) {
      const result = await exec(`
        (function() {
          // Procura a célula com badge "Pré-contrato" na listagem
          const badge = document.querySelector('span.vg-badge-content')
          if (badge && badge.textContent.trim() === 'Pré-contrato') {
            // Clica na linha (tr) que contém esse badge
            const row = badge.closest('tr')
            if (row) {
              row.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }))
              // Tenta pegar o ID do contrato da primeira célula
              const idCell = row.querySelector('td:first-child')
              return idCell ? idCell.textContent.trim() : 'selecionado'
            }
          }
          return ''
        })()
      `)
      if (result.result?.value) {
        console.log('✅ Contrato selecionado:', result.result.value)
        return true
      }
      await new Promise(r => setTimeout(r, 300))
    }
    return false
  })()

  if (!contratoSelecionado) {
    console.warn('⚠️ Não encontrei contrato com Pré-contrato — selecione manualmente.')
    await exec(`alert('⚠️ Não encontrei contrato Pré-contrato! Selecione manualmente e clique OK.')`)
  }

  // Aguarda um pouco pra modal fechar e campo preencher
  await new Promise(r => setTimeout(r, 1000))

  // Captura o número do contrato preenchido (vai ser usado em campos abaixo)
  console.log('📋 Capturando número do contrato...')
  const numeroContrato = await (async () => {
    const inicio = Date.now()
    while (Date.now() - inicio < 5000) {
      const result = await exec(`
        (function() {
          const el = document.querySelector('input#id_contrato')
          return el && el.value && el.value.trim().length > 0 ? el.value : ''
        })()
      `)
      if (result.result?.value) return result.result.value
      await new Promise(r => setTimeout(r, 300))
    }
    return ''
  })()

  if (numeroContrato) {
    console.log('✅ Número do contrato capturado:', numeroContrato)
  } else {
    console.warn('⚠️ Não consegui capturar o número do contrato.')
  }

  // Verifica se o contrato foi preenchido corretamente no label
  const contratoLabel = await exec(`
    (function() {
      const el = document.querySelector('input#id_contrato_label')
      return el && el.value ? el.value.trim() : ''
    })()
  `)
  if (contratoLabel.result?.value) {
    console.log('✅ Contrato confirmado:', contratoLabel.result.value)
  }

  // Clica na lupa do Plano (grupo)
  console.log('🔍 Abrindo busca de plano...')
  await waitFor('button[rel="id_grupo"].but_search')
  await exec(`
    (function() {
      // Pega a segunda lupa (last-visible) que é a que abre a listagem
      const btns = document.querySelectorAll('button[rel="id_grupo"].but_search')
      const btn = Array.from(btns).find(b => b.style.marginLeft || b.classList.contains('last-visible')) || btns[btns.length - 1]
      btn.click()
    })()
  `)
  await new Promise(r => setTimeout(r, 1500))

  // Clica em Atualizar pra carregar o plano
  console.log('🔄 Atualizando listagem de planos...')
  await waitFor('i.fa-arrows-rotate[title="Atualizar"]')
  await exec(`document.querySelector('i.fa-arrows-rotate[title="Atualizar"]').click()`)
  await new Promise(r => setTimeout(r, 1500))

  // Seleciona o primeiro plano da lista (clica na primeira linha do tbody)
  console.log('📋 Selecionando plano...')
  const planoSelecionado = await (async () => {
    const inicio = Date.now()
    while (Date.now() - inicio < 10000) {
      const result = await exec(`
        (function() {
          const row = document.querySelector('tbody tr.tableRow')
          if (row) {
            row.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }))
            const id = row.getAttribute('data-valorcampoautoincrement')
            return id || 'selecionado'
          }
          return ''
        })()
      `)
      if (result.result?.value) {
        console.log('✅ Plano selecionado:', result.result.value)
        return true
      }
      await new Promise(r => setTimeout(r, 300))
    }
    return false
  })()

  if (!planoSelecionado) {
    console.warn('⚠️ Não encontrei plano na listagem — selecione manualmente.')
    await exec(`alert('⚠️ Plano não encontrado! Selecione manualmente e clique OK.')`)
  }

  await new Promise(r => setTimeout(r, 1000))

  // Preenche Login com número do contrato
  console.log('📝 Preenchendo login com número do contrato...')
  await waitFor('input#login')
  await exec(`
    (function() {
      const input = document.querySelector('input#login')
      input.value = '${numeroContrato}'
      input.dispatchEvent(new Event('input', { bubbles: true }))
      input.dispatchEvent(new Event('change', { bubbles: true }))
    })()
  `)

  // Preenche Senha PPPoE com número do contrato
  console.log('📝 Preenchendo senha PPPoE com número do contrato...')
  await waitFor('input#senha')
  await exec(`
    (function() {
      const input = document.querySelector('input#senha')
      input.value = '${numeroContrato}'
      input.dispatchEvent(new Event('input', { bubbles: true }))
      input.dispatchEvent(new Event('change', { bubbles: true }))
    })()
  `)

  // ══════════════════════════════════════════
  //  PARTE 4 — DADOS TÉCNICOS
  // ══════════════════════════════════════════

  // Abre aba Dados técnicos
  console.log('🔧 Abrindo aba Dados técnicos...')
  await waitFor('a.tabTitle[rel="dados_tecnicos"]')  // ajustar rel se diferente
  await exec(`
    (function() {
      const tabs = document.querySelectorAll('a.tabTitle')
      const tab = Array.from(tabs).find(t => t.textContent.trim() === 'Dados técnicos')
      if (tab) tab.click()
    })()
  `)
  await new Promise(r => setTimeout(r, 1000))

  // Clica na lupa da Caixa de atendimento
  console.log('🔍 Abrindo busca de caixa de atendimento...')
  await waitFor('button[rel="id_caixa_ftth"].but_search')
  await exec(`
    (function() {
      const btns = document.querySelectorAll('button[rel="id_caixa_ftth"].but_search')
      const btn = Array.from(btns).find(b => b.classList.contains('last-visible')) || btns[btns.length - 1]
      btn.click()
    })()
  `)
  await new Promise(r => setTimeout(r, 1500))

  // Pesquisa a caixa de atendimento (dado que vem do OPA)
  const caixaAtendimento = 'SM002-01-15' // TODO: vem do OPA
  console.log('🔍 Pesquisando caixa:', caixaAtendimento)
  await waitFor('input.gridActionsSearchInput')
  await exec(`
    (function() {
      const input = document.querySelector('input.gridActionsSearchInput')
      input.value = '${caixaAtendimento}'
      input.dispatchEvent(new Event('input', { bubbles: true }))
      input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', keyCode: 13, bubbles: true }))
    })()
  `)
  await new Promise(r => setTimeout(r, 2000))

  // Double-click na primeira linha pra selecionar a caixa
  console.log('📋 Selecionando caixa de atendimento...')
  const caixaSelecionada = await (async () => {
    const inicio = Date.now()
    while (Date.now() - inicio < 10000) {
      const result = await exec(`
        (function() {
          const row = document.querySelector('table#grid_rel_id_caixa_ftth3 tbody tr.tableRow, tbody tr.tableRow')
          if (row) {
            row.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }))
            const id = row.getAttribute('data-valorcampoautoincrement')
            return id || 'selecionado'
          }
          return ''
        })()
      `)
      if (result.result?.value) {
        console.log('✅ Caixa selecionada:', result.result.value)
        return true
      }
      await new Promise(r => setTimeout(r, 300))
    }
    return false
  })()

  if (!caixaSelecionada) {
    console.warn('⚠️ Caixa não encontrada — selecione manualmente.')
    await exec(`alert('⚠️ Caixa de atendimento não encontrada! Selecione manualmente e clique OK.')`)
  }

  await new Promise(r => setTimeout(r, 1000))

  // Clica em Listar pra ver portas de atendimento
  console.log('📋 Listando portas de atendimento...')
  await waitFor('button')
  await exec(`
    (function() {
      const btns = document.querySelectorAll('button')
      const listar = Array.from(btns).find(b => b.textContent.trim() === 'Listar')
      if (listar) listar.click()
    })()
  `)
  await new Promise(r => setTimeout(r, 2000))

  // Seleciona a primeira porta com status "Disponível" e captura o número
  console.log('🔍 Procurando porta disponível...')
  let portaAtendimento = ''
  const portaSelecionada = await (async () => {
    const inicio = Date.now()
    while (Date.now() - inicio < 10000) {
      const result = await exec(`
        (function() {
          // Procura todas as linhas da tabela de portas
          const rows = document.querySelectorAll('tbody tr')
          for (const row of rows) {
            const cells = row.querySelectorAll('td')
            // Procura a célula com badge "Disponível"
            const badge = row.querySelector('span.vg-badge-content')
            if (badge && badge.textContent.trim() === 'Disponível') {
              // Pega o número da porta (primeira coluna)
              const porta = cells[0] ? cells[0].textContent.trim() : ''
              // Clica pra selecionar
              row.click()
              return porta || 'disponível'
            }
          }
          return ''
        })()
      `)
      if (result.result?.value) {
        portaAtendimento = result.result.value
        console.log('✅ Porta disponível encontrada:', portaAtendimento)
        return true
      }
      await new Promise(r => setTimeout(r, 300))
    }
    return false
  })()

  if (!portaSelecionada) {
    console.warn('⚠️ Nenhuma porta disponível encontrada!')
    await exec(`alert('⚠️ Nenhuma porta disponível! Selecione manualmente e clique OK.')`)
  }

  // Clica em Selecionar pra confirmar a porta
  console.log('✅ Confirmando porta selecionada...')
  await exec(`
    (function() {
      const btns = document.querySelectorAll('button')
      const selecionar = Array.from(btns).find(b => b.textContent.trim() === 'Selecionar')
      if (selecionar) selecionar.click()
    })()
  `)
  await new Promise(r => setTimeout(r, 1000))

  // ══════════════════════════════════════════
  //  SALVAR LOGIN
  // ══════════════════════════════════════════

  console.log('💾 Salvando login...')
  await exec(`document.querySelector('button[type="submit"][title="Alt+S"]').click()`)

  console.log('⏳ Aguardando confirmação de salvo do login...')
  const salvouLogin = await (async () => {
    const inicio = Date.now()
    while (Date.now() - inicio < 15000) {
      const result = await exec(`
        (function() {
          const el = document.querySelector('div.notificationMessage span')
          return el && el.innerText ? el.innerText.trim() : ''
        })()
      `)
      if (result.result?.value) {
        console.log('✅ Resposta IXC:', result.result.value)
        return true
      }
      await new Promise(r => setTimeout(r, 300))
    }
    return false
  })()

  if (!salvouLogin) {
    console.warn('⚠️ Não detectei confirmação do login — esperando 5s...')
    await new Promise(r => setTimeout(r, 5000))
  }

  // ══════════════════════════════════════════
  //  PARTE 5 — O.S.
  // ══════════════════════════════════════════

  // Aba O.S.
  console.log('📋 Abrindo aba O.S...')
  await waitFor('a.tabTitle[rel="13"]')
  await exec(`document.querySelector('a.tabTitle[rel="13"]').click()`)
  await new Promise(r => setTimeout(r, 1000))

  // Clica em Nova
  console.log('➕ Clicando em Nova O.S...')
  await waitFor('button[name="novo"]')
  await exec(`
    (function() {
      const btns = document.querySelectorAll('button[name="novo"]')
      const btn = Array.from(btns).find(b => b.textContent.trim() === 'Nova')
      if (btn) btn.click()
    })()
  `)
  await new Promise(r => setTimeout(r, 1500))

  // Assunto (código do OPA)
  const assuntoOS = '1' // TODO: vem do OPA
  console.log('📝 Preenchendo assunto da OS...')
  await waitFor('input#id_assunto')
  await exec(`
    (function() {
      const input = document.querySelector('input#id_assunto')
      input.value = '${assuntoOS}'
      input.dispatchEvent(new Event('input', { bubbles: true }))
      input.dispatchEvent(new Event('change', { bubbles: true }))
      input.dispatchEvent(new KeyboardEvent('keydown', { keyCode: 13, bubbles: true }))
    })()
  `)
  await new Promise(r => setTimeout(r, 500))

  // Setor — sempre 1
  console.log('📝 Preenchendo setor...')
  await waitFor('input#id_assunto')  // setor pode ser outro seletor, ajustar se necessário
  await exec(`
    (function() {
      // Tenta pelo seletor mais provável
      const input = document.querySelector('input#id_setor') || document.querySelector('input[name="id_setor"]')
      if (input) {
        input.value = '1'
        input.dispatchEvent(new Event('input', { bubbles: true }))
        input.dispatchEvent(new Event('change', { bubbles: true }))
        input.dispatchEvent(new KeyboardEvent('keydown', { keyCode: 13, bubbles: true }))
      }
    })()
  `)
  await new Promise(r => setTimeout(r, 500))

  // Descrição da OS — monta o texto com dados extraídos
  // TODO: substituir pelos dados reais do OPA
  const caixaOPA = 'TR012-12-14'       // vem do OPA
  const metragemOPA = '50m'             // vem do OPA
  const planoEscolhido = 'PLANO RENASCER FIDELIDADE: 570 MEGA 99,90' // vem do OPA
  const vendaDia = '06/03/2026'         // vem do OPA
  const agendadoPara = '07/03/2026'     // vem do OPA
  const turno = 'MANHÃ'                 // vem do OPA

  const descricaoOS = 'Vendedor: WILLIAN POERARI\n'
    + 'Caixa / Metragem / Porta: ' + caixaOPA + ' - ' + metragemOPA + ' - P' + portaAtendimento + '\n'
    + 'Plano Escolhido: ' + planoEscolhido + '\n'
    + 'Valor de Instalação no Ato: ISENTO\n'
    + 'Venda no Dia: ' + vendaDia + '\n'
    + 'Agendado Instalação Para: ' + agendadoPara + '\n'
    + 'Turno: ' + turno

  console.log('📝 Preenchendo descrição da OS...')
  await waitFor('textarea#mensagem')
  await exec("(function() { const textarea = document.querySelector('textarea#mensagem'); textarea.value = " + JSON.stringify(descricaoOS) + "; textarea.dispatchEvent(new Event('input', { bubbles: true })); textarea.dispatchEvent(new Event('change', { bubbles: true })); })()")
  await new Promise(r => setTimeout(r, 500))

  // Salvar OS
  console.log('💾 Salvando OS...')
  await exec(`document.querySelector('button[type="submit"][title="Alt+S"]').click()`)

  console.log('⏳ Aguardando confirmação de salvo da OS...')
  const salvouOS = await (async () => {
    const inicio = Date.now()
    while (Date.now() - inicio < 15000) {
      const result = await exec(`
        (function() {
          const el = document.querySelector('div.notificationMessage span')
          return el && el.innerText ? el.innerText.trim() : ''
        })()
      `)
      if (result.result?.value) {
        console.log('✅ Resposta IXC:', result.result.value)
        return true
      }
      await new Promise(r => setTimeout(r, 300))
    }
    return false
  })()

  if (!salvouOS) {
    console.warn('⚠️ Não detectei confirmação da OS — esperando 5s...')
    await new Promise(r => setTimeout(r, 5000))
  }

  // Popup de sucesso
  await exec(`alert('✅ CADASTRO COMPLETO!\\n\\nCliente cadastrado com sucesso!\\nContrato: ${numeroContrato}\\nPorta: P${portaAtendimento}\\nOS aberta e salva.')`)

  console.log('══════════════════════════════════════')
  console.log('✅ FLUXO COMPLETO!')
  console.log('📌 Contrato:', numeroContrato)
  console.log('📌 Porta:', portaAtendimento)
  console.log('══════════════════════════════════════')

  await client.close()
}

main()