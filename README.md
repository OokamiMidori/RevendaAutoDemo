# App Revenda — Demo (100% local, sem backend)

A mesma interface de `real/app-revenda`, mas sem nenhuma dependência de servidor: todos os
dados (contas, estoque, vendas, gastos) vivem só na memória do navegador e somem ao
recarregar a página. Feita para demonstrar a interface e o fluxo de telas sem precisar
configurar backend, banco de dados ou certificado nenhum.

## Rodando localmente

```bash
npm install
npm run dev
```

Abre em `http://localhost:5175` — pronto para usar, sem nenhuma configuração.

## Contas de teste

| E-mail | Senha | O que tem |
|---|---|---|
| `contato@revendaauto.com.br` | `123456` | Estoque, vendas e gastos de exemplo já carregados |
| `contato@outrarevenda.com.br` | `123456` | Conta zerada — útil pra ver o isolamento entre "empresas" diferentes |

## O que é (e o que não é) simulado

Toda a navegação, cadastro, edição e os cálculos do dashboard são reais — é só a "rede" que
não existe. Especificamente:

- **Emissão de NF-e**: gera o XML no schema oficial (com CFOP, NCM, impostos, chave de
  acesso calculada pelo algoritmo real de módulo 11) e sorteia aprovação ou recusa — porque
  uma emissão de verdade exige assinatura com certificado A1 e conversa com a SEFAZ, que só
  a versão conectada faz.
- **Cadastro no Renave**: o botão funciona e atualiza o status na tela, mas também é só
  local — a integração real depende de credenciamento junto ao Denatran.
- **Relatório Excel / exportação Domínio Web**: esses são gerados de verdade (arquivo
  `.xlsx`/`.csv` real, para download) — não têm nada de simulado, porque não dependem de
  nenhum servidor externo.

Uma faixa amarela fixa em toda tela lembra que é uma versão de demonstração.

## Stack

Igual à versão conectada — mesmo React 18 + Vite, Recharts, SheetJS, Lucide. A única
diferença de código está na camada de dados (estado local em vez de chamadas `fetch`).
