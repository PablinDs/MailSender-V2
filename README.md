# MailSender-V2 - Sistema de Envio Automatizado de E-mails com Integração ao Google Sheets  
**Autor:** Pablo Víctor  
**Data:** Maio de 2025  
---

## 1. Visão Geral

O **MailSender** é um sistema desenvolvido em **Google Apps Script** para automação do envio de e-mails personalizados, integrado diretamente ao **Google Sheets**.  
O objetivo é otimizar a comunicação com clientes e colaboradores, enviando mensagens automáticas com base em dados extraídos de uma planilha hospedada no Google Sheets.

O sistema foi projetado para facilitar o controle de **vencimento de certificados digitais**, automatizando o envio de lembretes e relatórios de status — tornando o processo rápido, seguro.
---

## 2. Funcionalidades

- **Leitura de dados** diretamente de uma planilha no Google Sheets.  
- **Envio automatizado de e-mails** para clientes com situação "SEM RETORNO".  
- **Geração de relatório HTML** contendo o status de cada envio.  
- **Envio automático do relatório** ao e-mail do gerente configurado.  
- **Log interno (Logger)** para rastrear execuções e possíveis erros.  
- Estrutura **modular e comentada**, pronta para uso corporativo e publicação pública.

---

## 3. Tecnologias Utilizadas

- **Google Apps Script** (baseado em JavaScript)  
- **APIs nativas do Google**:
  - `SpreadsheetApp` – leitura de planilhas.
  - `GmailApp` – envio de e-mails.  
- **HTML** – para formatação dos relatórios e corpo das mensagens.  
- **Planilha do Google Sheets** – utilizada como base de dados.  

---

## 4. Estrutura do Código

### 4.1 Função Principal: `enviarEmails()`
Responsável por:
- Ler a planilha configurada.
- Verificar a coluna **SITUAÇÃO**.
- Enviar e-mails apenas para clientes com status **“SEM RETORNO”**.
- Gerar e enviar o relatório HTML ao gerente.

### 4.2 Função Auxiliar: `gerarTabelaRelatorio(dados)`
Cria o corpo HTML do relatório a partir dos resultados dos envios.  
Retorna uma tabela formatada contendo:
- Empresa  
- E-mail  
- Situação  
- Status do envio (✔️ sucesso / ❌ erro)

### 4.3 Configurações do Script
No topo do código, as variáveis principais podem ser ajustadas:

```javascript
const SPREADSHEET_ID = 'ID_DA_SUA_PLANILHA';
const RANGE_DADOS = 'Controle!A2:G';
const EMAIL_REMETENTE = 'Nome do Remetente <noreply@empresa.com>';
const EMAIL_REPLY_TO = 'colaborador@empresa.com';
const EMAIL_GERENTE = 'gerente@empresa.com';
```
Esses valores controlam o comportamento do sistema e devem ser atualizados conforme sua estrutura.

## 6. Estrutura da Planilha (Google Sheets)

A planilha deve conter as seguintes colunas:

| Coluna | Cabeçalho               | Descrição                                   |
|:-------:|--------------------------|---------------------------------------------|
| A | **CÓDIGO** | Código interno da empresa |
| B | **EMPRESA / RAZÃO SOCIAL** | Nome da empresa |
| C | **CNPJ** | Documento da empresa |
| D | **VENCIMENTO** | Data de expiração do certificado |
| E | **SITUAÇÃO** | Status atual (ex: SEM RETORNO, VALIDADO, etc.) |
| F | **E-mail** | Endereço de envio |
| G | **TELEFONE** | Contato (opcional) |

> ⚠️ Apenas os registros com **SITUAÇÃO = “SEM RETORNO”** terão e-mails enviados.

---

## 7. Considerações de Segurança

- Nenhum dado sensível é armazenado diretamente no código.  
- Para publicação pública (ex: GitHub), utilize *placeholders* nos IDs e e-mails.  
- O script usa exclusivamente **APIs oficiais do Google**, sem dependências externas.  
- Todos os acessos e envios seguem as permissões da conta Google autenticada.

---

## 8. Como Usar

### 8.1 Pré-requisitos

- Conta Google com acesso ao **Google Sheets** e **Gmail**.  
- Permissão para criar e editar scripts (**Google Apps Script**).  
- Planilha configurada conforme o modelo indicado.

### 8.2 Passos para Execução Manual

1. Abra sua planilha no **Google Sheets**.  
2. Vá em **Extensões ▸ Apps Script**.  
3. Cole o conteúdo do arquivo `enviarEmails.js`.  
4. Atualize as variáveis de configuração no topo do código.  
5. Clique em **Executar ▶️**.  
6. Conceda as permissões solicitadas (na primeira execução).

Após isso, o sistema:

- Enviará os e-mails necessários.  
- Gerará o relatório.  
- E enviará o resumo ao gerente automaticamente.

---

## 9. Exemplo de Relatório

O gerente receberá um e-mail com uma tabela HTML semelhante a esta:

| Empresa | E-mail | Situação | Status |
|----------|---------|----------|--------|
| EMPRESA 1 ME | cliente@empresa.com | SEM RETORNO | ✔️ Enviado com sucesso |
| EMPRESA 2 LTDA | cliente2@empresacom | SEM RETORNO | ❌ Erro: endereço inválido |

O relatório é gerado automaticamente e contém:

- Data e hora da execução.  
- Status detalhado por empresa.  
- Indicação de erros, se houver.

---




