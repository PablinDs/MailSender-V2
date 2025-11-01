function enviarEmails() {

  // CONFIGURAÇÕES GERAIS
  const SPREADSHEET_ID = 'ID_DA_SUA_PLANILHA';
  const RANGE_DADOS = 'Controle!A2:G'; 
  const EMAIL_REMETENTE = 'Nome do Remetente <noreply@empresa.com>';
  const EMAIL_REPLY_TO = 'colaborador@empresa.com';
  const EMAIL_GERENTE = 'gerente@empresa.com';

  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const values = sheet.getRange(RANGE_DADOS).getValues();

  if (!values || values.length === 0) {
    Logger.log('Nenhum dado encontrado na planilha.');
    return;
  }

  const relatorio = [];

  values.forEach((linha, index) => {
    const codigo = linha[0];
    const empresa = linha[1];
    const cnpj = linha[2];
    const vencimento = linha[3];
    const situacao = (linha[4] || '').toString().trim().toUpperCase();
    const email = linha[5];

    if (situacao !== 'SEM RETORNO' || !email) return;

    const assunto = 'AVISO: Vencimento de Certificado Digital';
    const corpo = `
Olá,<br><br>
Estamos entrando em contato para informar que o certificado digital da empresa <b>${empresa}</b>, CNPJ <b>${cnpj}</b>,
expira no dia <b>${vencimento}</b>.<br><br>
É de extrema importância que a renovação seja realizada antes dessa data para evitar interrupções
em serviços como emissão de notas fiscais, declarações junto à Receita Federal e demais obrigações eletrônicas.<br><br>
Caso precise de auxílio no processo de renovação, nossa equipe está à disposição para orientá-lo.<br><br>
Atenciosamente,<br>
<b>Departamento de Certificação Digital</b>
`;

    try {
      GmailApp.sendEmail(email, assunto, '', {
        name: EMAIL_REMETENTE,
        replyTo: EMAIL_REPLY_TO,
        htmlBody: corpo
      });

      relatorio.push({
        empresa: empresa,
        email: email,
        situacao: situacao,
        status: '✔️ Enviado com sucesso'
      });

      Logger.log(`✔️ E-mail enviado para ${empresa} (${email})`);
    } catch (e) {
      relatorio.push({
        empresa: empresa,
        email: email,
        situacao: situacao,
        status: `❌ Erro: ${e.message}`
      });
      Logger.log(`❌ Erro ao enviar e-mail para ${empresa} (${email}): ${e}`);
    }
  });

  // RELATÓRIO PARA O GERENTE
  if (relatorio.length > 0) {
    const tabelaHTML = gerarTabelaRelatorio(relatorio);
    const assuntoRelatorio = 'Relatório de Envio - Certificados Digitais';
    const corpoRelatorio = `
      <h3>Resumo dos envios realizados</h3>
      ${tabelaHTML}
      <br>
      <i>Relatório gerado automaticamente em ${new Date().toLocaleString('pt-BR')}</i>
    `;

    GmailApp.sendEmail(EMAIL_GERENTE, assuntoRelatorio, '', {
      name: EMAIL_REMETENTE,
      htmlBody: corpoRelatorio
    });

    Logger.log(`📧 Relatório enviado para o gerente (${EMAIL_GERENTE}).`);
  } else {
    Logger.log('Nenhum envio realizado — nenhum cliente com situação "SEM RETORNO".');
  }
}


 // Gera uma tabela HTML para o corpo do relatório do gerente.
function gerarTabelaRelatorio(dados) {
  let html = `
    <table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse; font-family:Arial; font-size:13px;">
      <tr style="background-color:#f2f2f2;">
        <th>Empresa</th>
        <th>E-mail</th>
        <th>Situação</th>
        <th>Status do Envio</th>
      </tr>
  `;

  dados.forEach(item => {
    html += `
      <tr>
        <td>${item.empresa}</td>
        <td>${item.email}</td>
        <td>${item.situacao}</td>
        <td>${item.status}</td>
      </tr>
    `;
  });

  html += '</table>';
  return html;
}
