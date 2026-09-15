# JH Telecom Portal na Vercel

## Variáveis de ambiente

Configure na Vercel:

- `GOOGLE_CLIENT_ID`: OAuth 2.0 Client ID para aplicação Web.
- `GOOGLE_SERVICE_ACCOUNT_JSON`: conteúdo JSON da conta de serviço do Google Cloud, em uma única variável.

Ative a Google Sheets API no mesmo projeto do Google Cloud. Compartilhe as quatro planilhas usadas pelo portal com o e-mail `client_email` da conta de serviço, com permissão de editor para permitir os registros e atualizações.

No OAuth Client ID, adicione o domínio da Vercel em **Origens JavaScript autorizadas**, por exemplo `https://seu-projeto.vercel.app`.

## Deploy

1. Importe este repositório na Vercel.
2. Use a raiz do repositório como diretório do projeto.
3. Cadastre as duas variáveis de ambiente em Development, Preview e Production.
4. Faça o deploy e abra a URL publicada.
5. Entre com uma conta Google. O e-mail precisa estar na aba `PERMISSOES` para acessar sistemas restritos; o administrador atual é definido em `api/rpc.js`.

O arquivo `src/code.gs` fica como referência do projeto original. O deploy da Vercel usa `api/rpc.js` e `src/index.html`.