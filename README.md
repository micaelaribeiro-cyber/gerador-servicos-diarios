# Gerador de Serviços Diários

Aplicativo web (HTML + CSS + JavaScript) para processar o relatório de OS no navegador e gerar um Excel com sete abas por localidade. O conteúdo do relatório não é enviado a servidores.

## Dependência da biblioteca Excel

Esta versão carrega `xlsx-js-style` de duas CDNs em sequência: jsDelivr e unpkg. É necessário que o navegador tenha acesso à internet e que a rede permita pelo menos um destes domínios:
- `cdn.jsdelivr.net`
- `unpkg.com`

A biblioteca é código executado no navegador; o arquivo com dados dos clientes é lido e processado localmente. Se a rede corporativa bloquear as duas CDNs, solicite à TI a liberação desses domínios ou disponibilize o arquivo `xlsx.bundle.js` localmente em `vendor/`.

## Publicar no GitHub Pages

1. Abra o repositório `gerador-servicos-diarios`.
2. Clique em **Add file > Upload files**.
3. Envie o conteúdo desta pasta mantendo a estrutura: `index.html`, `css/`, `js/`, `tests/`, `vendor/`, `README.md` e `.gitignore`. Não envie relatórios reais de clientes.
4. Clique em **Commit changes**.
5. Acesse **Settings > Pages**.
6. Em **Build and deployment**, escolha **Deploy from a branch**.
7. Selecione a branch `main` e a pasta `/ (root)`, depois clique em **Save**.
8. Aguarde a publicação. O endereço será `https://SEU-USUARIO.github.io/gerador-servicos-diarios/`.

## Regras implementadas

- Exclui assuntos que contenham a palavra `retirada`.
- Mantém os demais assuntos.
- Cria sete abas: Cidades Externas, Itabatã, Mucuri, Nova Viçosa, Posto da Mata, Nanuque e Serra dos Aimorés.
- Complementos diferentes das seis localidades principais vão para Cidades Externas.
- Organiza as colunas na ordem definida no projeto.
- Exibe totais, distribuição e avisos de integridade.

## Testes

Os testes sintéticos podem ser executados com `node tests/test-processor.js` em um ambiente que tenha Node.js. Não é necessário instalar Node para usar o aplicativo publicado.
