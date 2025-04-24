function processarAcessos() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const abaForms = planilha.getSheetByName('Log_de_Pedidos'); // Nome da aba com respostas do Forms
  const abaDicionario = planilha.getSheetByName('Deparo'); // Nome da aba com o mapa de canais
  const abaDestino = planilha.getSheetByName('Controle_de_Acessos'); // Aba onde será gravado o acesso processado
  
  const dadosForms = abaForms.getDataRange().getValues(); // Pega todos os dados das respostas do formulário
  const dadosDicionario = abaDicionario.getDataRange().getValues(); // Pega o dicionário de canais
  const dadosDestino = abaDestino.getDataRange().getValues().map(linha => linha.join('|')); // Pega os dados existentes na aba de controle de acessos

  const mapaCanais = {};
  
  // Criando o mapa de canais a partir da aba 'Deparo'
  for (let i = 1; i < dadosDicionario.length; i++) {
    const nomeForms = dadosDicionario[i][0]; // Nome no Formulário
    const nomeReal = dadosDicionario[i][1];  // Nome real do canal
    mapaCanais[nomeForms.trim().toLowerCase()] = nomeReal;
  }

  Logger.log("Mapa de Canais: " + JSON.stringify(mapaCanais)); // Log para verificar o mapa de canais

  // Controlando se o e-mail já foi enviado
  const eMailsEnviados = new Set(); // Usaremos um Set para controlar os e-mails enviados

  // Processando as respostas do Forms
  for (let i = 1; i < dadosForms.length; i++) {
    const email = String(dadosForms[i][1]).trim(); // A coluna B contém o e-mail (começando de B2)
    const canaisForms = String(dadosForms[i][2]).trim(); // A coluna C contém os canais (começando de C2)
    
    // Verificando se a célula de e-mail ou canais está vazia
    if (!email || !canaisForms) {
      Logger.log("E-mail ou Canal vazio na linha " + (i + 1));
      continue;
    }

    Logger.log("Processando linha " + (i + 1) + " - Canal(s): " + canaisForms + ", E-mail: " + email);

    // Verifique o conteúdo de canaisForms (se não estiver vazio)
    if (!canaisForms) {
      Logger.log("ERRO: Canal não definido ou vazio na linha " + (i + 1));
      continue;
    }

    // Dividindo os canais separados por vírgula
    const canaisArray = canaisForms.split(',').map(canal => canal.trim()); // Separa os canais que foram separados por vírgula

    // Para cada canal selecionado, processamos individualmente
    canaisArray.forEach(canalForms => {
      const canalFormsTrimmed = canalForms.toLowerCase(); // Remove espaços e coloca em minúsculas
      const canalPadrao = mapaCanais[canalFormsTrimmed]; // Mapeia o canal para o nome real

      if (!canalPadrao) {
        Logger.log("Canal não mapeado: " + canalForms); // Se não encontrar o canal, registra
        return; // Pula se o canal não estiver mapeado no dicionário
      }

      const linhaNova = [canalPadrao, email, new Date()]; // Nova linha com canal, e-mail e data
      const chaveUnica = linhaNova.join('|'); // Cria uma chave única para verificar duplicatas

      // Adiciona a linha se a chave não existir
      if (!dadosDestino.includes(chaveUnica)) {
        abaDestino.appendRow(linhaNova); // Adiciona a linha à aba Controle_de_Acessos
        Logger.log("Adicionando linha: " + linhaNova); // Log da linha que foi adicionada

        // Verificar se o e-mail já foi enviado antes
        if (!eMailsEnviados.has(email)) {
          // Enviar e-mail de notificação se for um novo acesso
          const assunto = "Novo pedido de acesso";
          const mensagem = `
Novo pedido de acesso recebido:

E-mail: ${email}
Canal(s): ${canaisForms}
Data da solicitação: ${new Date().toLocaleString()}

Verifique na aba Controle_de_Acessos para liberar o acesso.
`;

          // Envia o e-mail para notificação
          MailApp.sendEmail("seu-email@exemplo.com", assunto, mensagem);  // Substitua pelo seu e-mail
          Logger.log(`E-mail de notificação enviado para: seu-email@exemplo.com`);

          // Marca o e-mail como enviado
          eMailsEnviados.add(email); // Adiciona o e-mail ao Set
        }
      } else {
        Logger.log("Linha já existe: " + linhaNova); // Log se a linha já existe
      }
    });
  }
}
