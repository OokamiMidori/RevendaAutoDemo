import React, { useState, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, LineChart, Line, PieChart, Pie, Cell,
  XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer
} from "recharts";
import {
  LayoutDashboard, Car, ShoppingCart, Receipt, FileText,
  Link2, Plus, Trash2, ShieldCheck, AlertTriangle, TrendingUp,
  TrendingDown, DollarSign, Wallet, LogOut, Lock, Mail, Eye, EyeOff,
  Upload, Download, Copy, Check, FileUp, FileSpreadsheet, FileDown,
  Building2, ImagePlus, X, FlaskConical, RefreshCcw
} from "lucide-react";

// ============================================================================
// VERSÃO DE DEMONSTRAÇÃO — não se conecta a nenhum servidor real.
// Todo dado (contas, estoque, vendas, gastos) vive só na memória desta aba do
// navegador e some ao recarregar a página. Serve para mostrar a interface e o
// fluxo de telas, não para uso real — para a versão conectada ao backend de
// verdade, veja o arquivo gestao-revenda.jsx.
// ============================================================================

// contas de demonstração — duas "empresas" diferentes pra mostrar que os dados
// de uma não aparecem pra outra, mesmo sendo tudo local
const CONTAS_DEMO = [
  { email: "contato@revendaauto.com.br", senha: "123456", nome: "Revenda Auto", razaoSocial: "Revenda Auto Veículos LTDA", cnpj: "", logo: null },
  { email: "contato@outrarevenda.com.br", senha: "123456", nome: "Outra Revenda", razaoSocial: "Outra Revenda Veículos LTDA", cnpj: "", logo: null },
];

// origem: "juridica" (concessionária/loja, com nota fiscal normal) ou "fisica" (pessoa
// física, exige nota fiscal de entrada emitida pela própria revenda)
const ESTOQUE_DEMO = [
  { id: "v1", placa: "RIO2A34", modelo: "Onix LT 1.0", marca: "Chevrolet", ano: 2021, km: 42000, precoCompra: 58000, precoVenda: 72900, status: "disponivel", origem: "juridica", vendedorNome: "", vendedorCpf: "", renaveCadastroStatus: "concluido", nfeEntradaStatus: null, dataCadastro: "2026-07-28" },
  { id: "v2", placa: "BRA1B77", modelo: "HB20 Comfort", marca: "Hyundai", ano: 2020, km: 61000, precoCompra: 49000, precoVenda: 61900, status: "disponivel", origem: "fisica", vendedorNome: "Ana Paula Ferreira", vendedorCpf: "123.456.789-00", renaveCadastroStatus: "concluido", nfeEntradaStatus: "emitida", dataCadastro: "2026-08-02" },
  { id: "v3", placa: "ARR9C12", modelo: "Corolla XEi", marca: "Toyota", ano: 2019, km: 78000, precoCompra: 89000, precoVenda: 104900, status: "vendido", origem: "juridica", vendedorNome: "", vendedorCpf: "", renaveCadastroStatus: "concluido", nfeEntradaStatus: null, dataCadastro: "2026-07-15" },
  { id: "v4", placa: "SPX5D90", modelo: "Compass Longitude", marca: "Jeep", ano: 2022, km: 25000, precoCompra: 132000, precoVenda: 158900, status: "disponivel", origem: "fisica", vendedorNome: "Roberto Nakamura", vendedorCpf: "987.654.321-00", renaveCadastroStatus: "pendente", nfeEntradaStatus: "pendente", dataCadastro: "2026-08-10" },
];
const VENDAS_DEMO = [
  { id: "s1", veiculoId: "v3", cliente: "Marcos Vinícius Andrade", data: "2026-08-05", valor: 104900, nfeStatus: "emitida", renaveStatus: "concluida" },
];
const GASTOS_DEMO = [
  { id: "g1", categoria: "Manutenção", descricao: "Revisão HB20", valor: 850, data: "2026-08-02" },
  { id: "g2", categoria: "Marketing", descricao: "Anúncios OLX/Webmotors", valor: 1200, data: "2026-08-04" },
  { id: "g3", categoria: "Documentação", descricao: "Transferência Corolla", valor: 610, data: "2026-08-05" },
  { id: "g4", categoria: "Estrutura", descricao: "Aluguel do pátio", valor: 4200, data: "2026-08-01" },
];

// cada empresa (chave = e-mail da conta) tem seu próprio estoque/vendas/gastos, isolados
// das demais — só a conta de demonstração principal vem com dados de exemplo. Uma "empresa"
// recém-cadastrada nesta demo sempre começa zerada, inclusive no dashboard.
const DADOS_VAZIOS = { estoque: [], vendas: [], gastos: [] };
const DADOS_INICIAIS_POR_EMPRESA = {
  "contato@revendaauto.com.br": { estoque: ESTOQUE_DEMO, vendas: VENDAS_DEMO, gastos: GASTOS_DEMO },
};

let proximoId = 100; // contador simples pra simular IDs novos sem precisar de um backend
const gerarIdLocal = (prefixo) => `${prefixo}${proximoId++}`;

// monta o histórico dos últimos 6 meses (faturamento/gastos/lucro) a partir das datas reais
// dos registros locais — mesmo cálculo que a versão conectada faz a partir do banco de dados
const MESES_ABREV = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"];
function calcularHistoricoMensal(vendas, gastos, estoque) {
  const hoje = new Date();
  const meses = [];
  for (let i = 5; i >= 0; i--) {
    const d = new Date(hoje.getFullYear(), hoje.getMonth() - i, 1);
    meses.push({ chave: `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`, mes: MESES_ABREV[d.getMonth()] });
  }
  return meses.map(({ chave, mes }) => {
    const faturamento = vendas.filter((v) => v.data?.startsWith(chave)).reduce((s, v) => s + Number(v.valor), 0);
    const gastosOperacionais = gastos.filter((g) => g.data?.startsWith(chave)).reduce((s, g) => s + Number(g.valor), 0);
    const gastoCompraVeiculos = estoque.filter((v) => v.dataCadastro?.startsWith(chave)).reduce((s, v) => s + Number(v.precoCompra), 0);
    const gastosTotal = gastosOperacionais + gastoCompraVeiculos;
    return { mes, faturamento, gastos: gastosTotal, lucro: faturamento - gastosTotal };
  });
}

const CORES_CATEGORIA = { Manutenção: "#2563EB", Marketing: "#F59E0B", Documentação: "#8B5CF6", Estrutura: "#EF4444", "Compra de veículos": "#0EA5E9", Outros: "#64748B" };
const fmt = (v) => v.toLocaleString("pt-BR", { style: "currency", currency: "BRL", maximumFractionDigits: 0 });

// marcas padronizadas — evita erro de digitação e mantém o estoque consistente para relatórios/filtros
const MARCAS = [
  "Audi", "BMW", "BYD", "Chery", "Chevrolet", "Citroën", "Fiat", "Ford",
  "GWM", "Honda", "Hyundai", "Jaguar", "Jeep", "Kia", "Land Rover", "Mercedes-Benz",
  "Mitsubishi", "Nissan", "Peugeot", "Porsche", "Renault", "RAM", "Subaru", "Suzuki",
  "Toyota", "Volkswagen", "Volvo",
].sort();

// CFOPs mais comuns para uma revenda vendendo um veículo usado (mercadoria adquirida de
// terceiros para revenda) — o correto para cada venda depende do estado do cliente e do
// regime tributário da revenda; confirme com o contador antes de usar em uma nota real.
const CFOPS_VENDA_VEICULO = [
  { codigo: "5102", descricao: "5102 — Venda de mercadoria de terceiros, dentro do estado" },
  { codigo: "6102", descricao: "6102 — Venda de mercadoria de terceiros, para outro estado" },
  { codigo: "5405", descricao: "5405 — Venda com ICMS retido por substituição tributária, dentro do estado" },
  { codigo: "6404", descricao: "6404 — Venda com ICMS retido por substituição tributária, para outro estado" },
];

const UFS = ["AC", "AL", "AP", "AM", "BA", "CE", "DF", "ES", "GO", "MA", "MT", "MS", "MG", "PA", "PB", "PR", "PE", "PI", "RJ", "RN", "RS", "RO", "RR", "SC", "SP", "SE", "TO"];

// código do IBGE de cada UF — usado nos dois primeiros dígitos da chave de acesso da NF-e
const CODIGO_UF = { AC: 12, AL: 27, AP: 16, AM: 13, BA: 29, CE: 23, DF: 53, ES: 32, GO: 52, MA: 21, MT: 51, MS: 50, MG: 31, PA: 15, PB: 25, PR: 41, PE: 26, PI: 22, RJ: 33, RN: 24, RS: 43, RO: 11, RR: 14, SC: 42, SP: 35, SE: 28, TO: 17 };

// dígito verificador (módulo 11) dos 43 primeiros dígitos da chave de acesso — mesmo
// algoritmo usado de verdade pela SEFAZ. A chave em si é só simulada (ver gerarXmlVenda),
// mas esse cálculo é o real.
function calcularDvChaveAcesso(chave43) {
  const pesos = [2, 3, 4, 5, 6, 7, 8, 9];
  let soma = 0, pesoIdx = 0;
  for (let i = chave43.length - 1; i >= 0; i--) {
    soma += Number(chave43[i]) * pesos[pesoIdx % 8];
    pesoIdx++;
  }
  const resto = soma % 11;
  return resto < 2 ? 0 : 11 - resto;
}

// monta uma chave de acesso de 44 dígitos no formato correto (cUF+AAMM+CNPJ+mod+serie+
// nNF+tpEmis+cNF+DV) — estruturalmente válida, mas simulada: numa nota real ela só existe
// depois que a SEFAZ autoriza a nota; aqui é só pra completar o exemplo.
function gerarChaveAcessoSimulada({ uf, dataEmissao, cnpj, serie, numero }) {
  const cUF = String(CODIGO_UF[uf] || "35").padStart(2, "0");
  const [ano, mes] = (dataEmissao || new Date().toISOString().slice(0, 10)).split("-");
  const aamm = `${(ano || "2026").slice(2)}${mes || "01"}`;
  const cnpjNumeros = String(cnpj || "").replace(/\D/g, "").padStart(14, "0").slice(0, 14);
  const mod = "55"; // modelo do documento fiscal — 55 é o da NF-e
  const serieFmt = String(serie || "1").padStart(3, "0");
  const nNF = String(numero || "1").padStart(9, "0");
  const tpEmis = "1"; // emissão normal
  const cNF = String(Math.floor(Math.random() * 1e8)).padStart(8, "0"); // código numérico aleatório
  const chave43 = `${cUF}${aamm}${cnpjNumeros}${mod}${serieFmt}${nNF}${tpEmis}${cNF}`;
  const dv = calcularDvChaveAcesso(chave43);
  return `${chave43}${dv}`;
}

// motivos de rejeição no mesmo formato que a SEFAZ costuma devolver — usados para simular
// uma emissão no front (não existe emissão real sem o backend assinando com o certificado
// A1 e conversando com o webservice, então aqui a aceitação/recusa é sorteada só pra dar
// pra testar a tela de "nota recusada -> editar e reemitir")
const MOTIVOS_RECUSA_NFE = [
  "Rejeição 217: CNPJ do emitente não habilitado para emissão de NF-e",
  "Rejeição 226: CFOP incompatível com a natureza da operação",
  "Rejeição 539: Duplicidade de NF-e (chave de acesso já autorizada)",
  "Rejeição 610: Inscrição Estadual do destinatário inválida",
  "Rejeição 778: Valor total da nota diverge do somatório dos itens",
  "Rejeição 999: Rejeição genérica — erro não especificado pela SEFAZ",
];

// simula o resultado de uma emissão: sorteia aprovação ou recusa (com um motivo realista).
// É só para testar o fluxo da tela — a emissão de verdade só existe quando o backend está
// assinando o XML e falando com o webservice da SEFAZ.
function simularResultadoEmissao() {
  const aprovada = Math.random() < 0.7; // 70% de chance de sucesso, só para o teste ficar interessante
  if (aprovada) return { status: "emitida", motivo: null };
  const motivo = MOTIVOS_RECUSA_NFE[Math.floor(Math.random() * MOTIVOS_RECUSA_NFE.length)];
  return { status: "recusada", motivo };
}

// ---------- geração e leitura do XML de notas (formato próprio, usado como base
// de teste no front — a emissão de verdade segue o layout oficial da SEFAZ/Domínio Web) ----------
const tag = (nome, valor) => `<${nome}>${String(valor ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;")}</${nome}>`;

function gerarXmlCompra(d) {
  return `<?xml version="1.0" encoding="UTF-8"?>
<notaFiscalCompra>
  ${tag("numero", d.numero)}
  ${tag("dataEmissao", d.dataEmissao)}
  ${tag("origem", d.origem)}
  <vendedor>
    ${tag("nome", d.vendedorNome)}
    ${tag("documento", d.vendedorDocumento)}
  </vendedor>
  <emitente>
    ${tag("nome", d.emitenteNome)}
  </emitente>
  <veiculo>
    ${tag("placa", d.placa)}
    ${tag("marca", d.marca)}
    ${tag("modelo", d.modelo)}
    ${tag("ano", d.ano)}
    ${tag("km", d.km)}
  </veiculo>
  ${tag("valor", d.valor)}
</notaFiscalCompra>`;
}

function gerarXmlVenda(d) {
  const valor = Number(d.valor) || 0;
  const baseIcms = valor;
  const valorIcms = baseIcms * (Number(d.aliquotaIcms) || 0) / 100;
  const valorPis = valor * (Number(d.aliquotaPis) || 0) / 100;
  const valorCofins = valor * (Number(d.aliquotaCofins) || 0) / 100;
  const formaPagamentoCodigo = { "À vista": "01", "Financiado": "15", "Cartão de crédito": "03", "PIX": "17", "Boleto": "15" }[d.formaPagamento] || "99";

  const chaveAcesso = gerarChaveAcessoSimulada({
    uf: d.emitenteUf, dataEmissao: d.dataEmissao, cnpj: d.emitenteCnpj, serie: d.serie, numero: d.numero,
  });
  // protocolo e data de recebimento só existem de verdade depois que a SEFAZ autoriza a
  // nota — aqui são só um exemplo, com o mesmo formato numérico que o retorno real teria
  const protocoloSimulado = `1${CODIGO_UF[d.emitenteUf] || 35}${new Date().getFullYear()}${String(Math.floor(Math.random() * 1e8)).padStart(8, "0")}`;

  return `<?xml version="1.0" encoding="UTF-8"?>
<!-- XML de EXEMPLO para testar a tela — chave de acesso, protocolo e assinatura abaixo
     são simulados. Uma NF-e real só existe depois que o backend assina este XML com o
     certificado A1 da empresa e a SEFAZ autoriza a emissão. -->
<nfeProc versao="4.00">
  <NFe>
    <infNFe versao="4.00" Id="NFe${chaveAcesso}">
      <ide>
        ${tag("cUF", CODIGO_UF[d.emitenteUf] || 35)}
        ${tag("natOp", "Venda de mercadoria adquirida de terceiros")}
        ${tag("mod", 55)}
        ${tag("serie", d.serie || 1)}
        ${tag("nNF", d.numero)}
        ${tag("dhEmi", d.dataEmissao)}
        ${tag("tpNF", 1)}
        ${tag("tpAmb", 2)}
        ${tag("finNFe", 1)}
      </ide>
      <emit>
        ${tag("CNPJ", d.emitenteCnpj)}
        ${tag("xNome", d.emitenteNome)}
        ${tag("IE", d.emitenteIe)}
        ${tag("CRT", 3)}
        <enderEmit>
          ${tag("xLgr", d.emitenteLogradouro)}
          ${tag("nro", d.emitenteNumero)}
          ${tag("xBairro", d.emitenteBairro)}
          ${tag("xMun", d.emitenteMunicipio)}
          ${tag("UF", d.emitenteUf)}
          ${tag("CEP", d.emitenteCep)}
        </enderEmit>
      </emit>
      <dest>
        ${tag(d.clienteTipoPessoa === "fisica" ? "CPF" : "CNPJ", d.clienteDocumento)}
        ${tag("xNome", d.clienteNome)}
        ${tag("indIEDest", d.clienteTipoPessoa === "juridica" ? 1 : 9)}
        ${tag("IE", d.clienteTipoPessoa === "juridica" ? (d.clienteIe || "ISENTO") : "")}
        ${tag("email", d.clienteEmail)}
        ${tag("fone", d.clienteTelefone)}
        <enderDest>
          ${tag("xLgr", d.clienteLogradouro)}
          ${tag("nro", d.clienteNumero)}
          ${tag("xCpl", d.clienteComplemento)}
          ${tag("xBairro", d.clienteBairro)}
          ${tag("xMun", d.clienteMunicipio)}
          ${tag("UF", d.clienteUf)}
          ${tag("CEP", d.clienteCep)}
        </enderDest>
      </dest>
      <det nItem="1">
        <prod>
          ${tag("cProd", d.placa)}
          ${tag("xProd", d.modelo)}
          ${tag("NCM", d.ncm)}
          ${tag("CFOP", d.cfop)}
          ${tag("uCom", "UN")}
          ${tag("qCom", "1.0000")}
          ${tag("vUnCom", valor.toFixed(2))}
          ${tag("vProd", valor.toFixed(2))}
          ${tag("placaVeiculo", d.placa)}
        </prod>
        <imposto>
          <ICMS>
            <ICMS00>
              ${tag("orig", 0)}
              ${tag("CST", "00")}
              ${tag("modBC", 3)}
              ${tag("vBC", baseIcms.toFixed(2))}
              ${tag("pICMS", Number(d.aliquotaIcms) || 0)}
              ${tag("vICMS", valorIcms.toFixed(2))}
            </ICMS00>
          </ICMS>
          <PIS>
            <PISAliq>
              ${tag("CST", "01")}
              ${tag("vBC", valor.toFixed(2))}
              ${tag("pPIS", Number(d.aliquotaPis) || 0)}
              ${tag("vPIS", valorPis.toFixed(2))}
            </PISAliq>
          </PIS>
          <COFINS>
            <COFINSAliq>
              ${tag("CST", "01")}
              ${tag("vBC", valor.toFixed(2))}
              ${tag("pCOFINS", Number(d.aliquotaCofins) || 0)}
              ${tag("vCOFINS", valorCofins.toFixed(2))}
            </COFINSAliq>
          </COFINS>
        </imposto>
      </det>
      <total>
        <ICMSTot>
          ${tag("vBC", baseIcms.toFixed(2))}
          ${tag("vICMS", valorIcms.toFixed(2))}
          ${tag("vProd", valor.toFixed(2))}
          ${tag("vPIS", valorPis.toFixed(2))}
          ${tag("vCOFINS", valorCofins.toFixed(2))}
          ${tag("vNF", valor.toFixed(2))}
        </ICMSTot>
      </total>
      <transp>${tag("modFrete", 9)}</transp>
      <pag>
        <detPag>
          ${tag("tPag", formaPagamentoCodigo)}
          ${tag("xPag", d.formaPagamento)}
          ${tag("vPag", valor.toFixed(2))}
        </detPag>
      </pag>
    </infNFe>
    <!-- assinatura SIMULADA — a assinatura real usa o certificado A1 da empresa e só pode
         ser gerada no backend; isto aqui é só texto de exemplo, não é uma assinatura válida -->
    <Signature>ASSINATURA_SIMULADA_PARA_TESTES_NAO_VALIDA</Signature>
  </NFe>
  <!-- protocolo SIMULADO — só existe de verdade depois que a SEFAZ autoriza a nota -->
  <protNFe versao="4.00">
    <infProt>
      ${tag("chNFe", chaveAcesso)}
      ${tag("dhRecbto", new Date().toISOString())}
      ${tag("nProt", protocoloSimulado)}
      ${tag("cStat", 100)}
      ${tag("xMotivo", "Autorizado o uso da NF-e (SIMULADO — não é uma autorização real da SEFAZ)")}
    </infProt>
  </protNFe>
</nfeProc>`;
}

// lê um XML de nota de compra (o gerado acima, ou um XML de fornecedor com as mesmas tags)
// e devolve os campos já prontos para preencher o formulário de estoque
function lerXmlCompra(texto) {
  const doc = new DOMParser().parseFromString(texto, "application/xml");
  if (doc.querySelector("parsererror")) throw new Error("XML inválido ou mal formado.");
  const val = (seletor) => doc.querySelector(seletor)?.textContent?.trim() || "";
  const origemLida = val("origem").toLowerCase() === "fisica" ? "fisica" : "juridica";
  return {
    placa: val("veiculo > placa").toUpperCase(),
    marca: val("veiculo > marca"),
    modelo: val("veiculo > modelo"),
    ano: val("veiculo > ano"),
    km: val("veiculo > km"),
    precoCompra: val("valor"),
    origem: origemLida,
    vendedorNome: val("vendedor > nome"),
    vendedorCpf: origemLida === "fisica" ? val("vendedor > documento") : "",
  };
}

// ---------- exportação de arquivos (download via Blob) ----------
// layout de exportação de gastos no padrão de importação de lançamentos da Domínio Web:
// texto delimitado por ponto e vírgula, data DD/MM/AAAA e valor com vírgula decimal.
// Este é o layout deste protótipo — confira com o contador o layout exato aceito pela
// versão da Domínio usada pela revenda antes de importar de verdade.
function gerarCsvGastosDominio(gastos) {
  const dataBr = (iso) => iso.split("-").reverse().join("/");
  const valorBr = (v) => Number(v).toFixed(2).replace(".", ",");
  const linhas = gastos.map((g) =>
    [dataBr(g.data), g.categoria, g.descricao.replace(/;/g, ","), valorBr(g.valor)].join(";")
  );
  return ["Data;Categoria;Historico;Valor", ...linhas].join("\r\n");
}

// gera o relatório de estoque em Excel (.xlsx), filtrado por período de cadastro
// gera o relatório de estoque em Excel (.xlsx), filtrado por período de cadastro —
// devolve a URL do arquivo em memória (não dispara download sozinho: XLSX.writeFile faz
// isso por dentro com um clique sintético, que fica bloqueado no ambiente do Claude.ai;
// quem chama esta função renderiza um link real para o usuário clicar)
function gerarExcelEstoque(veiculos, dataInicio, dataFim) {
  const filtrados = veiculos.filter((v) => {
    if (dataInicio && v.dataCadastro < dataInicio) return false;
    if (dataFim && v.dataCadastro > dataFim) return false;
    return true;
  });

  const linhas = filtrados.map((v) => ({
    Placa: v.placa,
    Modelo: v.modelo,
    Marca: v.marca,
    Ano: v.ano,
    KM: v.km,
    "Data de cadastro": v.dataCadastro,
    "Origem da compra": v.origem === "fisica" ? "Pessoa física" : "Pessoa jurídica",
    "Preço de compra": Number(v.precoCompra),
    "Preço de venda": Number(v.precoVenda),
    Status: v.status === "disponivel" ? "Disponível" : "Vendido",
    "Status Renave": v.renaveCadastroStatus === "concluido" ? "Concluído" : "Pendente",
  }));

  const planilha = XLSX.utils.json_to_sheet(linhas);
  planilha["!cols"] = [
    { wch: 10 }, { wch: 22 }, { wch: 14 }, { wch: 6 }, { wch: 10 }, { wch: 14 },
    { wch: 16 }, { wch: 14 }, { wch: 14 }, { wch: 12 }, { wch: 13 },
  ];
  const livro = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(livro, planilha, "Estoque");
  const nomeArquivo = `relatorio-estoque_${dataInicio || "inicio"}_a_${dataFim || "hoje"}.xlsx`;

  const arrayBuffer = XLSX.write(livro, { bookType: "xlsx", type: "array" });
  const blob = new Blob([arrayBuffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);

  return { total: filtrados.length, url, nomeArquivo };
}

// ---------- shell ----------
const NAV = [
  { id: "dashboard", label: "Dashboard", icon: LayoutDashboard },
  { id: "estoque", label: "Estoque", icon: Car },
  { id: "vendas", label: "Vendas", icon: ShoppingCart },
  { id: "gastos", label: "Gastos", icon: Receipt },
  { id: "notas", label: "Notas fiscais", icon: FileText },
  { id: "renave", label: "Integração Renave", icon: Link2 },
];

// ---------- estilos globais / responsividade ----------
// os componentes usam style inline (valores que não mudam por tamanho de tela); as regras
// aqui cobrem só o que precisa se adaptar entre celular, tablet e desktop.
function GlobalStyles() {
  return (
    <style>{`
      *, *::before, *::after { box-sizing: border-box; }
      .rv-shell { width: 100%; overflow-x: hidden; }
      .rv-main { width: 100%; }

      /* tabelas: em telas estreitas, rola na horizontal em vez de espremer as colunas */
      .rv-table-scroll { overflow-x: auto; -webkit-overflow-scrolling: touch; }
      .rv-table-scroll table { min-width: 640px; }

      /* KPIs do dashboard: grade que reflui sozinha, sem precisar de media query */
      .rv-kpi-row { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 14px; }

      /* gráficos do dashboard: lado a lado no desktop, empilhados no celular */
      .rv-charts-row { display: flex; gap: 16px; flex-wrap: wrap; }
      .rv-chart-wide { flex: 2 1 380px; min-width: 0; }
      .rv-chart-narrow { flex: 1 1 280px; min-width: 0; }

      .rv-auth-wrap { width: 100%; max-width: 380px; padding-left: 16px; padding-right: 16px; }

      .rv-sidebar-nav { display: flex; flex-direction: column; gap: 4px; }

      @media (max-width: 860px) {
        .rv-shell { flex-direction: column; }
        .rv-sidebar {
          width: 100% !important; flex-direction: row !important; align-items: center !important;
          padding: 8px 10px !important; gap: 6px !important; position: sticky; top: 0; z-index: 20;
        }
        .rv-sidebar-user { margin-bottom: 0 !important; flex-shrink: 0; }
        .rv-sidebar-nav { flex-direction: row !important; flex: 1 1 auto; min-width: 0; overflow-x: auto; -webkit-overflow-scrolling: touch; }
        .rv-nav-btn { padding: 8px 10px !important; font-size: 13px !important; }
        .rv-hide-mobile { display: none !important; }
        .rv-main { padding: 16px !important; }
      }

      @media (max-width: 480px) {
        .rv-nav-label { display: none; }
        .rv-nav-btn { padding: 8px !important; }
        .rv-main { padding: 12px !important; }
        /* uma coluna só: os valores em R$ usam espaço não-quebrável entre "R$" e o número,
           então em 2 colunas eles ficam espremidos demais e quebram no meio do valor */
        .rv-kpi-row { grid-template-columns: 1fr; }
      }
    `}</style>
  );
}

export default function App() {
  const [contas, setContas] = useState(CONTAS_DEMO);
  const [telaAuth, setTelaAuth] = useState("login");
  const [usuario, setUsuario] = useState(null);
  const [aba, setAba] = useState("dashboard");
  // dados de todas as "empresas" cadastradas nesta demo, indexados por e-mail — cada login
  // só lê/escreve na própria chave, então nenhuma conta vê estoque, vendas ou gastos de outra,
  // tudo isso só na memória desta aba (nada é salvo em disco ou enviado a lugar nenhum)
  const [dadosEmpresas, setDadosEmpresas] = useState(DADOS_INICIAIS_POR_EMPRESA);

  const emailAtual = usuario?.email;
  const dadosAtuais = (emailAtual && dadosEmpresas[emailAtual]) || DADOS_VAZIOS;
  const { estoque, vendas, gastos } = dadosAtuais;

  const criarSetter = (campo) => (valorOuFn) => {
    if (!emailAtual) return;
    setDadosEmpresas((atual) => {
      const empresaAtual = atual[emailAtual] || DADOS_VAZIOS;
      const novoValor = typeof valorOuFn === "function" ? valorOuFn(empresaAtual[campo]) : valorOuFn;
      return { ...atual, [emailAtual]: { ...empresaAtual, [campo]: novoValor } };
    });
  };
  const setEstoque = criarSetter("estoque");
  const setVendas = criarSetter("vendas");
  const setGastos = criarSetter("gastos");

  const historicoMensal = useMemo(() => calcularHistoricoMensal(vendas, gastos, estoque), [vendas, gastos, estoque]);

  // as funções abaixo têm exatamente a mesma assinatura que a versão conectada ao backend
  // (gestao-revenda.jsx) — só que aqui elas mexem direto no estado local, sem fetch nenhum.
  // Isso é o que permite reaproveitar os componentes de tela sem alterar uma linha deles.
  const criarVeiculo = async (dados) => {
    const novo = {
      id: gerarIdLocal("v"), ...dados,
      ano: Number(dados.ano), km: Number(dados.km), precoCompra: Number(dados.precoCompra), precoVenda: Number(dados.precoVenda),
      status: "disponivel",
      renaveCadastroStatus: "pendente",
      nfeEntradaStatus: dados.origem === "fisica" ? "pendente" : null,
      dataCadastro: new Date().toISOString().slice(0, 10),
    };
    setEstoque((atual) => [novo, ...atual]);
  };
  const removerVeiculo = async (id) => setEstoque((atual) => atual.filter((v) => v.id !== id));
  const marcarRenaveVeiculo = async (id) => setEstoque((atual) => atual.map((v) => (v.id === id ? { ...v, renaveCadastroStatus: "concluido" } : v)));
  const criarVenda = async (dados) => {
    const nova = { id: gerarIdLocal("s"), veiculoId: dados.veiculoId, cliente: dados.cliente, valor: Number(dados.valor), data: new Date().toISOString().slice(0, 10), nfeStatus: "pendente", renaveStatus: "pendente" };
    setVendas((atual) => [nova, ...atual]);
    setEstoque((atual) => atual.map((v) => (v.id === dados.veiculoId ? { ...v, status: "vendido" } : v)));
  };
  const criarGasto = async (dados) => {
    const novo = { id: gerarIdLocal("g"), categoria: dados.categoria, descricao: dados.descricao, valor: Number(dados.valor), data: new Date().toISOString().slice(0, 10) };
    setGastos((atual) => [novo, ...atual]);
  };
  const removerGasto = async (id) => setGastos((atual) => atual.filter((g) => g.id !== id));

  const sair = () => { setUsuario(null); setTelaAuth("login"); };

  if (!usuario) {
    if (telaAuth === "cadastro") {
      return (
        <>
          <GlobalStyles />
          <RegisterScreen
            contas={contas}
            onCadastrar={(conta) => {
              setContas((atual) => [...atual, conta]);
              setDadosEmpresas((atual) => ({ ...atual, [conta.email]: DADOS_VAZIOS }));
              setUsuario(conta);
            }}
            onIrParaLogin={() => setTelaAuth("login")}
          />
        </>
      );
    }
    return (
      <>
        <GlobalStyles />
        <LoginScreen contas={contas} onEntrar={setUsuario} onIrParaCadastro={() => setTelaAuth("cadastro")} />
      </>
    );
  }

  return (
    <div className="rv-shell" style={{ display: "flex", minHeight: "100vh", background: "#F1F5F9", fontFamily: "'Inter', system-ui, sans-serif", color: "#0F172A" }}>
      <GlobalStyles />
      <Sidebar aba={aba} setAba={setAba} usuario={usuario} onSair={sair} />
      <main className="rv-main" style={{ flex: 1, padding: "28px 32px", overflowY: "auto", minWidth: 0 }}>
        <FaixaDemo />
        {aba === "dashboard" && <Dashboard estoque={estoque} vendas={vendas} gastos={gastos} historicoMensal={historicoMensal} />}
        {aba === "estoque" && <Estoque estoque={estoque} criarVeiculo={criarVeiculo} removerVeiculo={removerVeiculo} marcarRenaveVeiculo={marcarRenaveVeiculo} />}
        {aba === "vendas" && <Vendas estoque={estoque} vendas={vendas} criarVenda={criarVenda} />}
        {aba === "gastos" && <Gastos gastos={gastos} criarGasto={criarGasto} removerGasto={removerGasto} />}
        {aba === "notas" && <NotasFiscais vendas={vendas} estoque={estoque} setVendas={setVendas} />}
        {aba === "renave" && <Renave vendas={vendas} estoque={estoque} marcarRenaveVeiculo={marcarRenaveVeiculo} />}
      </main>
    </div>
  );
}

// aviso fixo em toda tela logada, pra nunca deixar dúvida de que isso é só mostra
function FaixaDemo() {
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 10, background: "#FFFBEB", border: "1px solid #FDE68A", color: "#92400E", borderRadius: 10, padding: "10px 16px", marginBottom: 18, fontSize: 13 }}>
      <FlaskConical size={16} style={{ flexShrink: 0 }} />
      <span><strong>Versão de demonstração</strong> — nada aqui é salvo de verdade nem enviado a um servidor. Tudo some ao recarregar a página.</span>
    </div>
  );
}

function CampoAuth({ label, icon: Icon, marginBottom = 16, children }) {
  return (
    <>
      <label style={{ fontSize: 12, color: "#94A3B8", fontWeight: 600, display: "block", marginBottom: 6 }}>{label}</label>
      <div style={{ display: "flex", alignItems: "center", gap: 8, background: "#0F172A", border: "1px solid #334155", borderRadius: 8, padding: "0 12px", marginBottom }}>
        {Icon && <Icon size={15} color="#64748B" />}
        {children}
      </div>
    </>
  );
}
const estiloInputAuth = { flex: 1, background: "transparent", border: "none", outline: "none", padding: "10px 0", color: "#E2E8F0", fontSize: 13.5 };

// ---------- login ----------
function LoginScreen({ contas, onEntrar, onIrParaCadastro }) {
  const [email, setEmail] = useState("");
  const [senha, setSenha] = useState("");
  const [mostrarSenha, setMostrarSenha] = useState(false);
  const [erro, setErro] = useState("");
  const [carregando, setCarregando] = useState(false);

  const entrar = () => {
    setErro("");
    if (!email || !senha) {
      setErro("Preencha e-mail e senha.");
      return;
    }
    setCarregando(true);
    // simula o tempinho de uma chamada de rede, só pelo efeito visual — não fala com
    // nenhum servidor de verdade, é tudo comparado com a lista local CONTAS_DEMO
    setTimeout(() => {
      const conta = contas.find((c) => c.email === email && c.senha === senha);
      setCarregando(false);
      if (!conta) {
        setErro("E-mail ou senha incorretos.");
        return;
      }
      onEntrar(conta);
    }, 400);
  };

  return (
    <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: "#0F172A", fontFamily: "'Inter', system-ui, sans-serif", padding: "24px 0" }}>
      <div className="rv-auth-wrap">
        <div style={{ textAlign: "center", marginBottom: 20 }}>
          <div style={{ display: "inline-flex", alignItems: "center", gap: 6, padding: "5px 12px", borderRadius: 20, background: "#78350F33", border: "1px solid #92400E66", color: "#FCD34D", fontSize: 11.5, fontWeight: 600, marginBottom: 16 }}>
            <FlaskConical size={13} /> VERSÃO DE DEMONSTRAÇÃO
          </div>
          <div style={{ width: 44, height: 44, borderRadius: 12, background: "#2563EB1A", display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 14px" }}>
            <Car size={22} color="#60A5FA" />
          </div>
          <div style={{ fontSize: 12, letterSpacing: 2, color: "#64748B", fontWeight: 600 }}>PAINEL</div>
          <div style={{ fontSize: 22, fontWeight: 700, color: "#fff", letterSpacing: -0.3 }}>Revenda Auto</div>
          <div style={{ fontSize: 13, color: "#94A3B8", marginTop: 6 }}>Entre com a conta da sua revenda</div>
        </div>

        <div onKeyDown={(e) => e.key === "Enter" && entrar()} style={{ background: "#1E293B", borderRadius: 14, padding: 24, border: "1px solid #334155" }}>
          <CampoAuth label="E-mail" icon={Mail}>
            <input
              type="email"
              value={email}
              onChange={(e) => setEmail(e.target.value)}
              placeholder="contato@suarevenda.com.br"
              style={estiloInputAuth}
            />
          </CampoAuth>

          <CampoAuth label="Senha" icon={Lock} marginBottom={erro ? 10 : 20}>
            <input
              type={mostrarSenha ? "text" : "password"}
              value={senha}
              onChange={(e) => setSenha(e.target.value)}
              placeholder="Sua senha"
              style={estiloInputAuth}
            />
            <button type="button" onClick={() => setMostrarSenha(!mostrarSenha)} style={{ background: "none", border: "none", cursor: "pointer", color: "#64748B", display: "flex" }} aria-label="Mostrar senha">
              {mostrarSenha ? <EyeOff size={15} /> : <Eye size={15} />}
            </button>
          </CampoAuth>

          {erro && <div style={{ fontSize: 12.5, color: "#F87171", marginBottom: 14 }}>{erro}</div>}

          <button
            type="button"
            onClick={entrar}
            disabled={carregando}
            style={{ width: "100%", padding: "11px 0", background: "#2563EB", color: "#fff", border: "none", borderRadius: 8, fontSize: 14, fontWeight: 600, cursor: carregando ? "default" : "pointer", opacity: carregando ? 0.7 : 1 }}
          >
            {carregando ? "Entrando..." : "Entrar"}
          </button>

          <div style={{ textAlign: "center", marginTop: 16, paddingTop: 14, borderTop: "1px solid #334155" }}>
            <span style={{ fontSize: 12.5, color: "#94A3B8" }}>Ainda não tem conta? </span>
            <button type="button" onClick={onIrParaCadastro} style={{ background: "none", border: "none", cursor: "pointer", color: "#60A5FA", fontSize: 12.5, fontWeight: 600, padding: 0 }}>
              Criar conta da revenda
            </button>
          </div>
        </div>

        <div style={{ marginTop: 14, fontSize: 11.5, color: "#475569", lineHeight: 1.6, textAlign: "center" }}>
          Contas de demonstração: <strong style={{ color: "#94A3B8" }}>contato@revendaauto.com.br</strong> / <strong style={{ color: "#94A3B8" }}>123456</strong> (com dados de
          exemplo) ou <strong style={{ color: "#94A3B8" }}>contato@outrarevenda.com.br</strong> / <strong style={{ color: "#94A3B8" }}>123456</strong> (zerada).
          Nada aqui se conecta a um servidor real.
        </div>
      </div>
    </div>
  );
}

// ---------- cadastro ----------
function RegisterScreen({ contas, onCadastrar, onIrParaLogin }) {
  const [nome, setNome] = useState("");
  const [razaoSocial, setRazaoSocial] = useState("");
  const [cnpj, setCnpj] = useState("");
  const [email, setEmail] = useState("");
  const [senha, setSenha] = useState("");
  const [confirmarSenha, setConfirmarSenha] = useState("");
  const [logo, setLogo] = useState(null);
  const [mostrarSenha, setMostrarSenha] = useState(false);
  const [erro, setErro] = useState("");
  const [carregando, setCarregando] = useState(false);
  const inputLogoRef = useRef(null);

  const escolherLogo = (e) => {
    const arquivo = e.target.files?.[0];
    e.target.value = "";
    if (!arquivo) return;
    if (!arquivo.type.startsWith("image/")) {
      setErro("A logo precisa ser uma imagem (PNG, JPG ou SVG).");
      return;
    }
    if (arquivo.size > 2 * 1024 * 1024) {
      setErro("Escolha uma imagem de até 2MB.");
      return;
    }
    setErro("");
    const leitor = new FileReader();
    leitor.onload = () => setLogo(String(leitor.result));
    leitor.readAsDataURL(arquivo);
  };

  const cadastrar = () => {
    setErro("");
    if (!nome || !email || !senha) {
      setErro("Preencha nome da revenda, e-mail e senha.");
      return;
    }
    if (!/^\S+@\S+\.\S+$/.test(email)) {
      setErro("Digite um e-mail válido.");
      return;
    }
    if (senha.length < 6) {
      setErro("A senha precisa ter pelo menos 6 caracteres.");
      return;
    }
    if (senha !== confirmarSenha) {
      setErro("As senhas não coincidem.");
      return;
    }
    if (contas.some((c) => c.email === email)) {
      setErro("Já existe uma conta com esse e-mail.");
      return;
    }
    setCarregando(true);
    setTimeout(() => {
      setCarregando(false);
      onCadastrar({ nome, razaoSocial, cnpj, email, senha, logo });
    }, 400);
  };

  return (
    <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: "#0F172A", fontFamily: "'Inter', system-ui, sans-serif", padding: "40px 0" }}>
      <div className="rv-auth-wrap">
        <div style={{ textAlign: "center", marginBottom: 20 }}>
          <div style={{ display: "inline-flex", alignItems: "center", gap: 6, padding: "5px 12px", borderRadius: 20, background: "#78350F33", border: "1px solid #92400E66", color: "#FCD34D", fontSize: 11.5, fontWeight: 600, marginBottom: 16 }}>
            <FlaskConical size={13} /> VERSÃO DE DEMONSTRAÇÃO
          </div>
          <div style={{ width: 44, height: 44, borderRadius: 12, background: "#2563EB1A", display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 14px" }}>
            <Car size={22} color="#60A5FA" />
          </div>
          <div style={{ fontSize: 22, fontWeight: 700, color: "#fff", letterSpacing: -0.3 }}>Criar conta da revenda</div>
          <div style={{ fontSize: 13, color: "#94A3B8", marginTop: 6 }}>Leva menos de um minuto</div>
        </div>

        <div onKeyDown={(e) => e.key === "Enter" && cadastrar()} style={{ background: "#1E293B", borderRadius: 14, padding: 24, border: "1px solid #334155" }}>
          <CampoAuth label="Nome da revenda">
            <input value={nome} onChange={(e) => setNome(e.target.value)} placeholder="Ex.: Revenda Auto" style={estiloInputAuth} />
          </CampoAuth>

          <CampoAuth label="Razão social (opcional)">
            <input value={razaoSocial} onChange={(e) => setRazaoSocial(e.target.value)} placeholder="Razão social LTDA" style={estiloInputAuth} />
          </CampoAuth>

          <CampoAuth label="CNPJ (opcional)">
            <input value={cnpj} onChange={(e) => setCnpj(e.target.value)} placeholder="00.000.000/0001-00" style={estiloInputAuth} />
          </CampoAuth>

          <label style={{ fontSize: 12, color: "#94A3B8", fontWeight: 600, display: "block", marginBottom: 6 }}>Logo da empresa (opcional)</label>
          <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 16 }}>
            <input ref={inputLogoRef} type="file" accept="image/*" onChange={escolherLogo} style={{ display: "none" }} />
            <div style={{ width: 40, height: 40, borderRadius: 8, background: "#0F172A", border: "1px solid #334155", display: "flex", alignItems: "center", justifyContent: "center", overflow: "hidden", flexShrink: 0 }}>
              {logo ? <img src={logo} alt="Logo escolhida" style={{ width: "100%", height: "100%", objectFit: "cover" }} /> : <ImagePlus size={16} color="#64748B" />}
            </div>
            <button type="button" onClick={() => inputLogoRef.current?.click()} style={{ padding: "8px 12px", background: "#0F172A", color: "#CBD5E1", border: "1px solid #334155", borderRadius: 6, fontSize: 12.5, fontWeight: 600, cursor: "pointer" }}>
              {logo ? "Trocar imagem" : "Escolher imagem"}
            </button>
            {logo && (
              <button type="button" onClick={() => setLogo(null)} style={{ padding: "8px", background: "none", border: "none", color: "#64748B", cursor: "pointer", display: "flex" }} aria-label="Remover logo">
                <X size={15} />
              </button>
            )}
          </div>

          <CampoAuth label="E-mail" icon={Mail}>
            <input type="email" value={email} onChange={(e) => setEmail(e.target.value)} placeholder="contato@suarevenda.com.br" style={estiloInputAuth} />
          </CampoAuth>

          <CampoAuth label="Senha" icon={Lock}>
            <input
              type={mostrarSenha ? "text" : "password"}
              value={senha}
              onChange={(e) => setSenha(e.target.value)}
              placeholder="Mínimo de 6 caracteres"
              style={estiloInputAuth}
            />
            <button type="button" onClick={() => setMostrarSenha(!mostrarSenha)} style={{ background: "none", border: "none", cursor: "pointer", color: "#64748B", display: "flex" }} aria-label="Mostrar senha">
              {mostrarSenha ? <EyeOff size={15} /> : <Eye size={15} />}
            </button>
          </CampoAuth>

          <CampoAuth label="Confirmar senha" icon={Lock} marginBottom={erro ? 10 : 20}>
            <input
              type={mostrarSenha ? "text" : "password"}
              value={confirmarSenha}
              onChange={(e) => setConfirmarSenha(e.target.value)}
              placeholder="Repita a senha"
              style={estiloInputAuth}
            />
          </CampoAuth>

          {erro && (
            <div style={{ display: "flex", gap: 8, alignItems: "flex-start", fontSize: 12.5, color: "#FCA5A5", background: "#7F1D1D22", border: "1px solid #7F1D1D55", borderRadius: 8, padding: "10px 12px", marginBottom: 14 }}>
              <AlertTriangle size={14} style={{ marginTop: 1, flexShrink: 0 }} />
              <span>{erro}</span>
            </div>
          )}

          <button
            type="button"
            onClick={cadastrar}
            disabled={carregando}
            style={{ width: "100%", padding: "11px 0", background: "#2563EB", color: "#fff", border: "none", borderRadius: 8, fontSize: 14, fontWeight: 600, cursor: carregando ? "default" : "pointer", opacity: carregando ? 0.7 : 1 }}
          >
            {carregando ? "Criando conta..." : "Criar conta"}
          </button>

          <div style={{ textAlign: "center", marginTop: 16, paddingTop: 14, borderTop: "1px solid #334155" }}>
            <span style={{ fontSize: 12.5, color: "#94A3B8" }}>Já tem conta? </span>
            <button type="button" onClick={onIrParaLogin} style={{ background: "none", border: "none", cursor: "pointer", color: "#60A5FA", fontSize: 12.5, fontWeight: 600, padding: 0 }}>
              Fazer login
            </button>
          </div>
        </div>

        <div style={{ marginTop: 14, fontSize: 11.5, color: "#475569", lineHeight: 1.6, textAlign: "center" }}>
          Versão de demonstração: a conta fica só na memória desta aba do navegador e some ao recarregar a página. Nenhum dado é enviado a um servidor.
        </div>
      </div>
    </div>
  );
}

function Sidebar({ aba, setAba, usuario, onSair }) {
  const iniciais = (usuario?.nome || "?").split(" ").map((p) => p[0]).slice(0, 2).join("").toUpperCase();
  return (
    <aside className="rv-sidebar" style={{ width: 232, background: "#0F172A", color: "#E2E8F0", padding: "24px 14px", display: "flex", flexDirection: "column", gap: 4 }}>
      <div className="rv-hide-mobile" style={{ padding: "0 10px 24px" }}>
        <div style={{ fontSize: 12, letterSpacing: 2, color: "#64748B", fontWeight: 600 }}>PAINEL</div>
        <div style={{ fontSize: 19, fontWeight: 700, letterSpacing: -0.3 }}>Revenda Auto</div>
      </div>

      <div className="rv-sidebar-user" style={{ display: "flex", alignItems: "center", gap: 10, padding: "10px 10px", marginBottom: 12, background: "#1E293B", borderRadius: 10 }}>
        <div style={{ width: 32, height: 32, borderRadius: "50%", background: "#2563EB", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 12, fontWeight: 700, flexShrink: 0, overflow: "hidden" }}>
          {usuario?.logo ? <img src={usuario.logo} alt="Logo da empresa" style={{ width: "100%", height: "100%", objectFit: "cover" }} /> : iniciais}
        </div>
        <div className="rv-hide-mobile" style={{ flex: 1, minWidth: 0 }}>
          <div style={{ fontSize: 13, fontWeight: 600, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{usuario?.nome}</div>
          <div style={{ fontSize: 11, color: "#64748B", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{usuario?.email}</div>
        </div>
        <button onClick={onSair} title="Sair" aria-label="Sair" style={{ border: "none", background: "none", cursor: "pointer", color: "#64748B", display: "flex", flexShrink: 0 }}>
          <LogOut size={15} />
        </button>
      </div>

      <div className="rv-sidebar-nav">
        {NAV.map((item) => {
          const Icon = item.icon;
          const ativo = aba === item.id;
          return (
            <button
              key={item.id}
              onClick={() => setAba(item.id)}
              className="rv-nav-btn"
              style={{
                display: "flex", alignItems: "center", gap: 10, padding: "10px 12px",
                borderRadius: 8, border: "none", cursor: "pointer", textAlign: "left",
                background: ativo ? "#1E293B" : "transparent",
                color: ativo ? "#60A5FA" : "#CBD5E1",
                fontWeight: ativo ? 600 : 500, fontSize: 14, flexShrink: 0,
              }}
            >
              <Icon size={17} /> <span className="rv-nav-label">{item.label}</span>
            </button>
          );
        })}
      </div>
      <div className="rv-hide-mobile" style={{ marginTop: "auto", padding: "14px 12px", background: "#1E293B", borderRadius: 10, fontSize: 12, color: "#94A3B8", lineHeight: 1.5 }}>
        NF-e, Renave e certificado A1 dependem de um serviço de backend seguro — não rodam no navegador. Veja as abas "Notas fiscais" e "Renave" para o desenho da integração.
      </div>
    </aside>
  );
}

function Card({ children, style, className }) {
  return <div className={className} style={{ background: "#fff", borderRadius: 12, border: "1px solid #E2E8F0", padding: 18, minWidth: 0, ...style }}>{children}</div>;
}

// combobox de marca: busca por texto e mostra no máximo 6 opções por vez
function SeletorMarca({ valor, onSelecionar, largura = 150 }) {
  const [texto, setTexto] = useState(valor || "");
  const [aberto, setAberto] = useState(false);
  const ALTURA_ITEM = 34;
  const ITENS_VISIVEIS = 6;

  const opcoes = useMemo(() => {
    const termo = texto.trim().toLowerCase();
    return termo ? MARCAS.filter((m) => m.toLowerCase().includes(termo)) : MARCAS;
  }, [texto]);

  const escolher = (marca) => {
    setTexto(marca);
    onSelecionar(marca);
    setAberto(false);
  };

  return (
    <div style={{ position: "relative", width: largura }}>
      <input
        value={texto}
        onChange={(e) => { setTexto(e.target.value); onSelecionar(""); setAberto(true); }}
        onFocus={() => setAberto(true)}
        onBlur={() => setTimeout(() => setAberto(false), 120)}
        placeholder="Buscar marca"
        style={{ width: "100%", padding: "8px 10px", borderRadius: 6, border: "1px solid #CBD5E1", fontSize: 13, boxSizing: "border-box" }}
      />
      {aberto && (
        <div
          style={{
            position: "absolute", top: "calc(100% + 4px)", left: 0, right: 0, background: "#fff",
            border: "1px solid #E2E8F0", borderRadius: 8, boxShadow: "0 4px 12px rgba(15,23,42,0.08)", zIndex: 10,
            maxHeight: ALTURA_ITEM * ITENS_VISIVEIS, overflowY: "auto",
          }}
        >
          {opcoes.length === 0 && (
            <div style={{ padding: "8px 12px", fontSize: 12.5, color: "#94A3B8" }}>Nenhuma marca encontrada</div>
          )}
          {opcoes.map((marca) => (
            <div
              key={marca}
              onMouseDown={() => escolher(marca)}
              style={{ padding: "8px 12px", height: ALTURA_ITEM, boxSizing: "border-box", display: "flex", alignItems: "center", fontSize: 13, cursor: "pointer", background: marca === valor ? "#EFF6FF" : "transparent", color: marca === valor ? "#1D4ED8" : "#0F172A" }}
            >
              {marca}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function Kpi({ label, valor, icon: Icon, cor }) {
  return (
    <Card style={{ display: "flex", alignItems: "center", gap: 14, flex: 1, minWidth: 0 }}>
      <div style={{ width: 40, height: 40, borderRadius: 10, background: cor + "1A", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
        <Icon size={19} color={cor} />
      </div>
      <div style={{ minWidth: 0 }}>
        <div style={{ fontSize: 12, color: "#64748B", fontWeight: 600 }}>{label}</div>
        <div style={{ fontSize: 20, fontWeight: 700, fontFamily: "'JetBrains Mono', monospace", overflowWrap: "break-word", wordBreak: "break-word" }}>{valor}</div>
      </div>
    </Card>
  );
}

function Titulo({ children, sub }) {
  return (
    <div style={{ marginBottom: 20 }}>
      <h1 style={{ fontSize: 22, fontWeight: 700, margin: 0, letterSpacing: -0.3 }}>{children}</h1>
      {sub && <p style={{ margin: "4px 0 0", color: "#64748B", fontSize: 13.5 }}>{sub}</p>}
    </div>
  );
}

// ---------- dashboard ----------
function Dashboard({ estoque, vendas, gastos, historicoMensal }) {
  const gastosPorCategoria = useMemo(() => {
    const mapa = {};
    gastos.forEach((g) => { mapa[g.categoria] = (mapa[g.categoria] || 0) + g.valor; });
    const compraVeiculos = estoque.reduce((s, v) => s + Number(v.precoCompra), 0);
    if (compraVeiculos > 0) mapa["Compra de veículos"] = compraVeiculos;
    return Object.entries(mapa).map(([categoria, valor]) => ({ categoria, valor }));
  }, [gastos, estoque]);

  // mês atual calculado ao vivo a partir das vendas, gastos e compras de veículos cadastrados;
  // os meses anteriores vêm do histórico da própria empresa (vazio para quem acabou de se
  // cadastrar). Regime de caixa: o custo de compra do veículo entra como gasto assim que ele
  // é cadastrado no estoque, não só quando é vendido.
  const mesAtual = useMemo(() => {
    const faturamento = vendas.reduce((s, v) => s + Number(v.valor), 0);
    const gastosOperacionais = gastos.reduce((s, g) => s + Number(g.valor), 0);
    const gastoCompraVeiculos = estoque.reduce((s, v) => s + Number(v.precoCompra), 0);
    const gastosTotal = gastosOperacionais + gastoCompraVeiculos;
    const lucro = faturamento - gastosTotal;
    return { mes: "Ago", faturamento, gastos: gastosTotal, gastoCompraVeiculos, lucro };
  }, [vendas, gastos, estoque]);

  const dadosMensais = useMemo(
    () => [...(historicoMensal || []), mesAtual],
    [historicoMensal, mesAtual]
  );

  return (
    <div>
      <Titulo sub="Visão geral do faturamento, gastos e lucro">Dashboard</Titulo>

      <div className="rv-kpi-row" style={{ marginBottom: 20 }}>
        <Kpi label="Faturamento (ago)" valor={fmt(mesAtual.faturamento)} icon={DollarSign} cor="#2563EB" />
        <Kpi label="Compra de veículos (ago)" valor={fmt(mesAtual.gastoCompraVeiculos)} icon={Car} cor="#8B5CF6" />
        <Kpi label="Outros gastos (ago)" valor={fmt(mesAtual.gastos - mesAtual.gastoCompraVeiculos)} icon={TrendingDown} cor="#F59E0B" />
        <Kpi label="Lucro (ago)" valor={fmt(mesAtual.lucro)} icon={TrendingUp} cor="#10B981" />
      </div>

      <div className="rv-charts-row" style={{ marginBottom: 16 }}>
        <Card className="rv-chart-wide">
          <div style={{ fontWeight: 600, fontSize: 14, marginBottom: 12 }}>Faturamento, gastos e lucro por mês</div>
          <ResponsiveContainer width="100%" height={260}>
            <BarChart data={dadosMensais}>
              <CartesianGrid strokeDasharray="3 3" stroke="#E2E8F0" vertical={false} />
              <XAxis dataKey="mes" tick={{ fontSize: 12, fill: "#64748B" }} axisLine={false} tickLine={false} />
              <YAxis tick={{ fontSize: 11, fill: "#64748B" }} axisLine={false} tickLine={false} tickFormatter={(v) => `${v / 1000}k`} />
              <Tooltip formatter={(v) => fmt(v)} contentStyle={{ borderRadius: 8, border: "1px solid #E2E8F0", fontSize: 12 }} />
              <Legend wrapperStyle={{ fontSize: 12 }} />
              <Bar dataKey="faturamento" name="Faturamento" fill="#2563EB" radius={[4, 4, 0, 0]} />
              <Bar dataKey="gastos" name="Gastos" fill="#F59E0B" radius={[4, 4, 0, 0]} />
              <Bar dataKey="lucro" name="Lucro" fill="#10B981" radius={[4, 4, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </Card>

        <Card className="rv-chart-narrow">
          <div style={{ fontWeight: 600, fontSize: 14, marginBottom: 12 }}>Gastos por categoria (mês atual)</div>
          <ResponsiveContainer width="100%" height={260}>
            <PieChart>
              <Pie data={gastosPorCategoria} dataKey="valor" nameKey="categoria" innerRadius={45} outerRadius={80} paddingAngle={2}>
                {gastosPorCategoria.map((entry, i) => (
                  <Cell key={i} fill={CORES_CATEGORIA[entry.categoria] || CORES_CATEGORIA.Outros} />
                ))}
              </Pie>
              <Tooltip formatter={(v) => fmt(v)} contentStyle={{ borderRadius: 8, border: "1px solid #E2E8F0", fontSize: 12 }} />
              <Legend wrapperStyle={{ fontSize: 11 }} />
            </PieChart>
          </ResponsiveContainer>
        </Card>
      </div>

      <Card>
        <div style={{ fontWeight: 600, fontSize: 14, marginBottom: 12 }}>Tendência de lucro mensal</div>
        <ResponsiveContainer width="100%" height={200}>
          <LineChart data={dadosMensais}>
            <CartesianGrid strokeDasharray="3 3" stroke="#E2E8F0" vertical={false} />
            <XAxis dataKey="mes" tick={{ fontSize: 12, fill: "#64748B" }} axisLine={false} tickLine={false} />
            <YAxis tick={{ fontSize: 11, fill: "#64748B" }} axisLine={false} tickLine={false} tickFormatter={(v) => `${v / 1000}k`} />
            <Tooltip formatter={(v) => fmt(v)} contentStyle={{ borderRadius: 8, border: "1px solid #E2E8F0", fontSize: 12 }} />
            <Line type="monotone" dataKey="lucro" stroke="#10B981" strokeWidth={2.5} dot={{ r: 3 }} />
          </LineChart>
        </ResponsiveContainer>
      </Card>
    </div>
  );
}

// ---------- estoque ----------
const NOVO_VEICULO_VAZIO = { placa: "", modelo: "", marca: "", ano: "", km: "", precoCompra: "", precoVenda: "", origem: "juridica", vendedorNome: "", vendedorCpf: "" };

function Estoque({ estoque, criarVeiculo, removerVeiculo, marcarRenaveVeiculo }) {
  const [novo, setNovo] = useState(NOVO_VEICULO_VAZIO);
  const [erro, setErro] = useState("");
  const [enviando, setEnviando] = useState(false);
  const [importado, setImportado] = useState(null); // resumo do que veio do arquivo importado
  const [avisoImportacao, setAvisoImportacao] = useState("");
  const inputArquivoRef = useRef(null);
  const [periodoInicio, setPeriodoInicio] = useState("");
  const [periodoFim, setPeriodoFim] = useState("");
  const [relatorioGerado, setRelatorioGerado] = useState(null); // { url, nomeArquivo, total } ou null
  const [erroRelatorio, setErroRelatorio] = useState("");

  const importarArquivo = (e) => {
    const arquivo = e.target.files?.[0];
    e.target.value = ""; // permite reimportar o mesmo arquivo depois
    if (!arquivo) return;
    setAvisoImportacao("");
    setImportado(null);

    const extensao = arquivo.name.split(".").pop().toLowerCase();

    if (extensao === "xml") {
      const leitor = new FileReader();
      leitor.onload = () => {
        try {
          const dados = lerXmlCompra(String(leitor.result));
          setNovo((atual) => ({
            ...atual,
            placa: dados.placa || atual.placa,
            marca: MARCAS.includes(dados.marca) ? dados.marca : atual.marca,
            modelo: dados.modelo || atual.modelo,
            ano: dados.ano || atual.ano,
            km: dados.km || atual.km,
            precoCompra: dados.precoCompra || atual.precoCompra,
            origem: dados.origem,
            vendedorNome: dados.vendedorNome,
            vendedorCpf: dados.vendedorCpf,
          }));
          setImportado(arquivo.name);
          if (dados.marca && !MARCAS.includes(dados.marca)) {
            setAvisoImportacao(`A marca "${dados.marca}" do XML não está na lista padronizada — selecione a marca correta manualmente.`);
          }
        } catch (err) {
          setAvisoImportacao("Não foi possível ler esse XML: " + err.message);
        }
      };
      leitor.readAsText(arquivo);
    } else if (extensao === "pdf") {
      // leitura de texto de PDF exige uma biblioteca de parsing/OCR no backend;
      // aqui só anexamos o arquivo como referência e o preenchimento continua manual
      setImportado(arquivo.name);
      setAvisoImportacao("PDF anexado como referência. A leitura automática dos campos a partir do PDF precisa de processamento no backend — preencha os campos manualmente.");
    } else {
      setAvisoImportacao("Formato não suportado. Envie um arquivo .xml ou .pdf.");
    }
  };

  const adicionar = async () => {
    if (!novo.placa || !novo.modelo) {
      setErro("Preencha ao menos a placa e o modelo.");
      return;
    }
    if (!MARCAS.includes(novo.marca)) {
      setErro("Selecione a marca na lista de busca.");
      return;
    }
    if (novo.origem === "fisica" && (!novo.vendedorNome || !novo.vendedorCpf)) {
      setErro("Compra de pessoa física exige nome e CPF do vendedor, para a nota fiscal de entrada.");
      return;
    }
    setErro("");
    setEnviando(true);
    try {
      await criarVeiculo(novo);
      setNovo(NOVO_VEICULO_VAZIO);
      setImportado(null);
      setAvisoImportacao("");
    } catch (err) {
      setErro(err.message);
    } finally {
      setEnviando(false);
    }
  };

  const remover = async (id) => {
    try {
      await removerVeiculo(id);
    } catch (err) {
      setErro(err.message);
    }
  };

  const cadastrarNoRenave = async (id) => {
    try {
      await marcarRenaveVeiculo(id);
    } catch (err) {
      setErro(err.message);
    }
  };

  const gerarRelatorio = () => {
    if (relatorioGerado?.url) URL.revokeObjectURL(relatorioGerado.url); // libera o anterior, se houver
    setRelatorioGerado(null);

    if (periodoInicio && periodoFim && periodoInicio > periodoFim) {
      setErroRelatorio("A data inicial não pode ser depois da data final.");
      return;
    }
    setErroRelatorio("");
    const resultado = gerarExcelEstoque(estoque, periodoInicio, periodoFim);
    setRelatorioGerado(resultado);
  };

  const campo = (chave, placeholder, largura = 100) => (
    <input
      value={novo[chave]}
      onChange={(e) => setNovo({ ...novo, [chave]: e.target.value })}
      placeholder={placeholder}
      style={{ width: largura, padding: "8px 10px", borderRadius: 6, border: "1px solid #CBD5E1", fontSize: 13 }}
    />
  );

  return (
    <div>
      <Titulo sub="Veículos disponíveis para venda">Estoque</Titulo>

      <Card style={{ marginBottom: 16 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12, paddingBottom: 12, marginBottom: 12, borderBottom: "1px solid #F1F5F9" }}>
          <input ref={inputArquivoRef} type="file" accept=".xml,.pdf" onChange={importarArquivo} style={{ display: "none" }} />
          <button
            onClick={() => inputArquivoRef.current?.click()}
            style={{ display: "flex", alignItems: "center", gap: 6, padding: "8px 14px", background: "#F1F5F9", color: "#334155", border: "1px dashed #CBD5E1", borderRadius: 6, fontSize: 13, fontWeight: 600, cursor: "pointer" }}
          >
            <FileUp size={15} /> Importar nota de compra (XML ou PDF)
          </button>
          {importado && (
            <span style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 12.5, color: "#166534", background: "#DCFCE7", padding: "5px 10px", borderRadius: 20 }}>
              <Check size={13} /> {importado}
            </span>
          )}
          <span style={{ fontSize: 12, color: "#94A3B8" }}>Os campos abaixo são preenchidos automaticamente a partir do XML — confira antes de salvar.</span>
        </div>
        {avisoImportacao && (
          <div style={{ fontSize: 12.5, color: "#92400E", background: "#FEF3C7", padding: "8px 12px", borderRadius: 6, marginBottom: 12 }}>{avisoImportacao}</div>
        )}

        <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "flex-start" }}>
          <input
            value={novo.placa}
            onChange={(e) => setNovo({ ...novo, placa: e.target.value.toUpperCase() })}
            placeholder="Placa"
            style={{ width: 100, padding: "8px 10px", borderRadius: 6, border: "1px solid #CBD5E1", fontSize: 13, textTransform: "uppercase" }}
          />
          {campo("modelo", "Modelo", 150)}
          <SeletorMarca valor={novo.marca} onSelecionar={(marca) => setNovo({ ...novo, marca })} largura={150} />
          {campo("ano", "Ano", 70)}
          {campo("km", "KM", 80)}
          {campo("precoCompra", "Preço de compra", 130)}
          {campo("precoVenda", "Preço de venda", 130)}
        </div>

        <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center", marginTop: 10, paddingTop: 10, borderTop: "1px solid #F1F5F9" }}>
          <span style={{ fontSize: 12.5, color: "#64748B", fontWeight: 600 }}>Comprado de:</span>
          <div style={{ display: "flex", border: "1px solid #CBD5E1", borderRadius: 6, overflow: "hidden" }}>
            {[["juridica", "Pessoa jurídica"], ["fisica", "Pessoa física"]].map(([valor, label]) => (
              <button
                key={valor}
                onClick={() => setNovo({ ...novo, origem: valor })}
                style={{
                  padding: "7px 12px", fontSize: 12.5, border: "none", cursor: "pointer",
                  background: novo.origem === valor ? "#2563EB" : "#fff",
                  color: novo.origem === valor ? "#fff" : "#334155", fontWeight: 600,
                }}
              >
                {label}
              </button>
            ))}
          </div>
          {novo.origem === "fisica" && (
            <>
              <input
                value={novo.vendedorNome}
                onChange={(e) => setNovo({ ...novo, vendedorNome: e.target.value })}
                placeholder="Nome do vendedor"
                style={{ width: 170, padding: "8px 10px", borderRadius: 6, border: "1px solid #CBD5E1", fontSize: 13 }}
              />
              <input
                value={novo.vendedorCpf}
                onChange={(e) => setNovo({ ...novo, vendedorCpf: e.target.value })}
                placeholder="CPF do vendedor"
                style={{ width: 140, padding: "8px 10px", borderRadius: 6, border: "1px solid #CBD5E1", fontSize: 13 }}
              />
            </>
          )}
          <button onClick={adicionar} disabled={enviando} style={{ display: "flex", alignItems: "center", gap: 6, padding: "8px 14px", background: enviando ? "#93C5FD" : "#2563EB", color: "#fff", border: "none", borderRadius: 6, fontSize: 13, fontWeight: 600, cursor: enviando ? "default" : "pointer", marginLeft: "auto" }}>
            <Plus size={15} /> {enviando ? "Salvando..." : "Adicionar"}
          </button>
        </div>

        {novo.origem === "fisica" && (
          <div style={{ fontSize: 12, color: "#64748B", marginTop: 8 }}>
            Como pessoa física não emite nota fiscal, a própria revenda emite a nota de entrada (emitente e destinatário são a revenda) — ela aparece na aba "Notas fiscais".
          </div>
        )}
        {erro && <div style={{ fontSize: 12.5, color: "#DC2626", marginTop: 10 }}>{erro}</div>}
      </Card>

      <Card style={{ marginBottom: 16 }}>
        <div style={{ fontWeight: 600, fontSize: 14, marginBottom: 4 }}>Relatório do estoque por período</div>
        <div style={{ fontSize: 12.5, color: "#64748B", marginBottom: 12 }}>Gera uma planilha Excel (.xlsx) com os veículos cadastrados no período escolhido. Deixe as datas em branco para incluir todo o estoque.</div>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "flex-end" }}>
          <div>
            <label style={{ fontSize: 11.5, color: "#64748B", fontWeight: 600, display: "block", marginBottom: 4 }}>De</label>
            <input type="date" value={periodoInicio} onChange={(e) => setPeriodoInicio(e.target.value)} style={{ padding: "8px 10px", borderRadius: 6, border: "1px solid #CBD5E1", fontSize: 13 }} />
          </div>
          <div>
            <label style={{ fontSize: 11.5, color: "#64748B", fontWeight: 600, display: "block", marginBottom: 4 }}>Até</label>
            <input type="date" value={periodoFim} onChange={(e) => setPeriodoFim(e.target.value)} style={{ padding: "8px 10px", borderRadius: 6, border: "1px solid #CBD5E1", fontSize: 13 }} />
          </div>
          <button onClick={gerarRelatorio} style={{ display: "flex", alignItems: "center", gap: 6, padding: "8px 14px", background: "#10B981", color: "#fff", border: "none", borderRadius: 6, fontSize: 13, fontWeight: 600, cursor: "pointer" }}>
            <FileSpreadsheet size={15} /> Gerar relatório Excel
          </button>
        </div>
        {erroRelatorio && <div style={{ fontSize: 12.5, color: "#DC2626", marginTop: 10 }}>{erroRelatorio}</div>}
        {relatorioGerado && (
          relatorioGerado.total === 0 ? (
            <div style={{ fontSize: 12.5, color: "#64748B", marginTop: 10 }}>Nenhum veículo cadastrado nesse período.</div>
          ) : (
            <div style={{ display: "flex", alignItems: "center", gap: 10, marginTop: 12 }}>
              <a
                href={relatorioGerado.url}
                download={relatorioGerado.nomeArquivo}
                target="_blank"
                rel="noopener noreferrer"
                style={{ display: "flex", alignItems: "center", gap: 6, padding: "8px 14px", background: "#F0FDF4", color: "#166534", border: "1px solid #BBF7D0", borderRadius: 6, fontSize: 13, fontWeight: 600, textDecoration: "none" }}
              >
                <Download size={15} /> Baixar {relatorioGerado.nomeArquivo}
              </a>
              <span style={{ fontSize: 12, color: "#94A3B8" }}>{relatorioGerado.total} veículo(s) — clique para salvar</span>
            </div>
          )
        )}
      </Card>

      <Card style={{ padding: 0 }}>
        <div className="rv-table-scroll"><table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ background: "#F8FAFC", textAlign: "left" }}>
              {["Placa", "Modelo", "Marca", "Ano", "KM", "Cadastro", "Compra", "Venda", "Origem", "Status", "Renave", ""].map((h) => (
                <th key={h} style={{ padding: "10px 14px", color: "#64748B", fontWeight: 600, fontSize: 11.5, textTransform: "uppercase", letterSpacing: 0.4 }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {estoque.map((v) => (
              <tr key={v.id} style={{ borderTop: "1px solid #F1F5F9" }}>
                <td style={{ padding: "10px 14px", fontFamily: "'JetBrains Mono', monospace" }}>{v.placa}</td>
                <td style={{ padding: "10px 14px" }}>{v.modelo}</td>
                <td style={{ padding: "10px 14px" }}>{v.marca}</td>
                <td style={{ padding: "10px 14px" }}>{v.ano}</td>
                <td style={{ padding: "10px 14px" }}>{Number(v.km).toLocaleString("pt-BR")}</td>
                <td style={{ padding: "10px 14px", fontSize: 12.5, color: "#64748B" }}>{v.dataCadastro?.split("-").reverse().join("/")}</td>
                <td style={{ padding: "10px 14px" }}>{fmt(Number(v.precoCompra))}</td>
                <td style={{ padding: "10px 14px" }}>{fmt(Number(v.precoVenda))}</td>
                <td style={{ padding: "10px 14px", fontSize: 12.5, color: "#64748B" }}>{v.origem === "fisica" ? "Pessoa física" : "Pessoa jurídica"}</td>
                <td style={{ padding: "10px 14px" }}>
                  <span style={{ padding: "3px 9px", borderRadius: 20, fontSize: 11.5, fontWeight: 600, background: v.status === "disponivel" ? "#DCFCE7" : "#F1F5F9", color: v.status === "disponivel" ? "#166534" : "#64748B" }}>
                    {v.status === "disponivel" ? "Disponível" : "Vendido"}
                  </span>
                </td>
                <td style={{ padding: "10px 14px" }}>
                  {v.renaveCadastroStatus === "concluido" ? (
                    <Selo status={v.renaveCadastroStatus} />
                  ) : (
                    <button onClick={() => cadastrarNoRenave(v.id)} style={{ padding: "4px 10px", borderRadius: 20, fontSize: 11.5, fontWeight: 600, background: "#FEF3C7", color: "#92400E", border: "none", cursor: "pointer" }}>
                      Cadastrar no Renave
                    </button>
                  )}
                </td>
                <td style={{ padding: "10px 14px" }}>
                  <button onClick={() => remover(v.id)} style={{ border: "none", background: "none", cursor: "pointer", color: "#94A3B8" }}><Trash2 size={15} /></button>
                </td>
              </tr>
            ))}
          </tbody>
        </table></div>
      </Card>
    </div>
  );
}

// ---------- vendas ----------
function Vendas({ estoque, vendas, criarVenda }) {
  const [veiculoId, setVeiculoId] = useState("");
  const [cliente, setCliente] = useState("");
  const [valor, setValor] = useState("");
  const [erro, setErro] = useState("");
  const [enviando, setEnviando] = useState(false);
  const disponiveis = estoque.filter((v) => v.status === "disponivel");

  const registrar = async () => {
    if (!veiculoId || !cliente || !valor) {
      setErro("Selecione o veículo e preencha cliente e valor.");
      return;
    }
    setErro("");
    setEnviando(true);
    try {
      await criarVenda({ veiculoId, cliente, valor: Number(valor) });
      setVeiculoId(""); setCliente(""); setValor("");
    } catch (err) {
      setErro(err.message);
    } finally {
      setEnviando(false);
    }
  };

  return (
    <div>
      <Titulo sub="Registre a venda de um veículo do estoque">Vendas</Titulo>

      <Card style={{ marginBottom: 16 }}>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
          <select value={veiculoId} onChange={(e) => setVeiculoId(e.target.value)} style={{ padding: "8px 10px", borderRadius: 6, border: "1px solid #CBD5E1", fontSize: 13, minWidth: 220 }}>
            <option value="">Selecione o veículo</option>
            {disponiveis.map((v) => <option key={v.id} value={v.id}>{v.placa} — {v.modelo}</option>)}
          </select>
          <input value={cliente} onChange={(e) => setCliente(e.target.value)} placeholder="Nome do cliente" style={{ padding: "8px 10px", borderRadius: 6, border: "1px solid #CBD5E1", fontSize: 13, width: 200 }} />
          <input value={valor} onChange={(e) => setValor(e.target.value)} placeholder="Valor da venda" style={{ padding: "8px 10px", borderRadius: 6, border: "1px solid #CBD5E1", fontSize: 13, width: 140 }} />
          <button onClick={registrar} disabled={enviando} style={{ display: "flex", alignItems: "center", gap: 6, padding: "8px 14px", background: enviando ? "#93C5FD" : "#2563EB", color: "#fff", border: "none", borderRadius: 6, fontSize: 13, fontWeight: 600, cursor: enviando ? "default" : "pointer" }}>
            <Plus size={15} /> {enviando ? "Registrando..." : "Registrar venda"}
          </button>
        </div>
        {erro && <div style={{ fontSize: 12.5, color: "#DC2626", marginTop: 10 }}>{erro}</div>}
      </Card>

      <Card style={{ padding: 0 }}>
        <div className="rv-table-scroll"><table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ background: "#F8FAFC", textAlign: "left" }}>
              {["Veículo", "Cliente", "Data", "Valor", "Lucro", "NF-e", "Renave"].map((h) => (
                <th key={h} style={{ padding: "10px 14px", color: "#64748B", fontWeight: 600, fontSize: 11.5, textTransform: "uppercase", letterSpacing: 0.4 }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {vendas.map((venda) => {
              const veiculo = estoque.find((v) => v.id === venda.veiculoId);
              const lucro = veiculo ? venda.valor - Number(veiculo.precoCompra) : null;
              return (
                <tr key={venda.id} style={{ borderTop: "1px solid #F1F5F9" }}>
                  <td style={{ padding: "10px 14px" }}>{veiculo ? `${veiculo.placa} — ${veiculo.modelo}` : "—"}</td>
                  <td style={{ padding: "10px 14px" }}>{venda.cliente}</td>
                  <td style={{ padding: "10px 14px" }}>{venda.data}</td>
                  <td style={{ padding: "10px 14px" }}>{fmt(venda.valor)}</td>
                  <td style={{ padding: "10px 14px", color: lucro >= 0 ? "#166534" : "#B91C1C", fontWeight: 600 }}>{lucro !== null ? fmt(lucro) : "—"}</td>
                  <td style={{ padding: "10px 14px" }}><Selo status={venda.nfeStatus} /></td>
                  <td style={{ padding: "10px 14px" }}><Selo status={venda.renaveStatus} /></td>
                </tr>
              );
            })}
          </tbody>
        </table></div>
      </Card>
    </div>
  );
}

function Selo({ status }) {
  const mapa = {
    emitida: ["Emitida", "#DCFCE7", "#166534"],
    concluida: ["Concluída", "#DCFCE7", "#166534"],
    concluido: ["Concluído", "#DCFCE7", "#166534"],
    pendente: ["Pendente", "#FEF3C7", "#92400E"],
    recusada: ["Recusada", "#FEE2E2", "#991B1B"],
  };
  const [label, bg, cor] = mapa[status] || mapa.pendente;
  return <span style={{ padding: "3px 9px", borderRadius: 20, fontSize: 11.5, fontWeight: 600, background: bg, color: cor }}>{label}</span>;
}

// ---------- gastos ----------
function Gastos({ gastos, criarGasto, removerGasto }) {
  const [categoria, setCategoria] = useState("Manutenção");
  const [descricao, setDescricao] = useState("");
  const [valor, setValor] = useState("");
  const [csvExportado, setCsvExportado] = useState("");
  const [erro, setErro] = useState("");
  const [enviando, setEnviando] = useState(false);
  const total = gastos.reduce((s, g) => s + g.valor, 0);

  const adicionar = async () => {
    if (!descricao || !valor) {
      setErro("Preencha descrição e valor.");
      return;
    }
    setErro("");
    setEnviando(true);
    try {
      await criarGasto({ categoria, descricao, valor: Number(valor) });
      setDescricao(""); setValor("");
    } catch (err) {
      setErro(err.message);
    } finally {
      setEnviando(false);
    }
  };

  const remover = async (id) => {
    try {
      await removerGasto(id);
    } catch (err) {
      setErro(err.message);
    }
  };

  return (
    <div>
      <Titulo sub="Despesas operacionais da revenda">Gastos</Titulo>

      <div style={{ marginBottom: 16 }}>
        <Kpi label="Total de gastos registrados" valor={fmt(total)} icon={Wallet} cor="#F59E0B" />
      </div>

      <Card style={{ marginBottom: 16 }}>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
          <select value={categoria} onChange={(e) => setCategoria(e.target.value)} style={{ padding: "8px 10px", borderRadius: 6, border: "1px solid #CBD5E1", fontSize: 13 }}>
            {Object.keys(CORES_CATEGORIA).filter((c) => c !== "Compra de veículos").map((c) => <option key={c} value={c}>{c}</option>)}
          </select>
          <input value={descricao} onChange={(e) => setDescricao(e.target.value)} placeholder="Descrição do gasto" style={{ padding: "8px 10px", borderRadius: 6, border: "1px solid #CBD5E1", fontSize: 13, width: 240 }} />
          <input value={valor} onChange={(e) => setValor(e.target.value)} placeholder="Valor" style={{ padding: "8px 10px", borderRadius: 6, border: "1px solid #CBD5E1", fontSize: 13, width: 120 }} />
          <button onClick={adicionar} disabled={enviando} style={{ display: "flex", alignItems: "center", gap: 6, padding: "8px 14px", background: enviando ? "#93C5FD" : "#2563EB", color: "#fff", border: "none", borderRadius: 6, fontSize: 13, fontWeight: 600, cursor: enviando ? "default" : "pointer" }}>
            <Plus size={15} /> {enviando ? "Salvando..." : "Adicionar gasto"}
          </button>
        </div>
        {erro && <div style={{ fontSize: 12.5, color: "#DC2626", marginTop: 10 }}>{erro}</div>}
      </Card>

      <Card style={{ padding: 0 }}>
        <div className="rv-table-scroll"><table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ background: "#F8FAFC", textAlign: "left" }}>
              {["Categoria", "Descrição", "Data", "Valor", ""].map((h) => (
                <th key={h} style={{ padding: "10px 14px", color: "#64748B", fontWeight: 600, fontSize: 11.5, textTransform: "uppercase", letterSpacing: 0.4 }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {gastos.map((g) => (
              <tr key={g.id} style={{ borderTop: "1px solid #F1F5F9" }}>
                <td style={{ padding: "10px 14px" }}>
                  <span style={{ display: "inline-flex", alignItems: "center", gap: 6 }}>
                    <span style={{ width: 8, height: 8, borderRadius: "50%", background: CORES_CATEGORIA[g.categoria] || CORES_CATEGORIA.Outros }} />
                    {g.categoria}
                  </span>
                </td>
                <td style={{ padding: "10px 14px" }}>{g.descricao}</td>
                <td style={{ padding: "10px 14px" }}>{g.data}</td>
                <td style={{ padding: "10px 14px", fontWeight: 600 }}>{fmt(g.valor)}</td>
                <td style={{ padding: "10px 14px" }}>
                  <button onClick={() => remover(g.id)} style={{ border: "none", background: "none", cursor: "pointer", color: "#94A3B8" }}><Trash2 size={15} /></button>
                </td>
              </tr>
            ))}
          </tbody>
        </table></div>
      </Card>

      <Card style={{ marginTop: 16 }}>
        <div style={{ fontWeight: 600, fontSize: 14, marginBottom: 4 }}>Exportar para a Domínio Web</div>
        <div style={{ fontSize: 12.5, color: "#64748B", marginBottom: 12 }}>
          Gera um arquivo com os {gastos.length} gasto(s) cadastrados, no layout de importação de lançamentos deste protótipo (texto separado por ponto e vírgula, data e valor no formato brasileiro). Confira com o contador o layout exato aceito pela versão da Domínio da revenda antes de importar de verdade.
        </div>
        <button
          onClick={() => setCsvExportado(gerarCsvGastosDominio(gastos))}
          disabled={gastos.length === 0}
          style={{ display: "flex", alignItems: "center", gap: 6, padding: "9px 16px", background: gastos.length === 0 ? "#CBD5E1" : "#2563EB", color: "#fff", border: "none", borderRadius: 6, fontSize: 13, fontWeight: 600, cursor: gastos.length === 0 ? "not-allowed" : "pointer", marginBottom: csvExportado ? 14 : 0 }}
        >
          <FileDown size={15} /> Gerar arquivo para a Domínio
        </button>
        {csvExportado && <CaixaArquivo conteudo={csvExportado} nomeArquivo="gastos_dominio_web.csv" mimetype="text/csv;charset=utf-8;" />}
      </Card>
    </div>
  );
}

// ---------- caixa de arquivo gerado (XML ou CSV), com copiar e baixar ----------
function CaixaArquivo({ conteudo, nomeArquivo, mimetype = "application/xml" }) {
  const [copiado, setCopiado] = useState(false);

  // gera a URL do blob uma vez por conteúdo (não a cada render) e libera a anterior quando
  // o conteúdo muda ou o componente desmonta, pra não vazar memória
  const urlDownload = useMemo(() => {
    const blob = new Blob([conteudo], { type: mimetype });
    return URL.createObjectURL(blob);
  }, [conteudo, mimetype]);

  React.useEffect(() => () => URL.revokeObjectURL(urlDownload), [urlDownload]);

  const copiar = () => {
    navigator.clipboard?.writeText(conteudo).then(() => {
      setCopiado(true);
      setTimeout(() => setCopiado(false), 1500);
    }).catch(() => {});
  };

  return (
    <div>
      <div style={{ display: "flex", gap: 8, marginBottom: 8, flexWrap: "wrap" }}>
        <button onClick={copiar} style={{ display: "flex", alignItems: "center", gap: 6, padding: "7px 12px", background: "#fff", color: "#334155", border: "1px solid #CBD5E1", borderRadius: 6, fontSize: 12.5, fontWeight: 600, cursor: "pointer" }}>
          {copiado ? <Check size={14} color="#16A34A" /> : <Copy size={14} />} {copiado ? "Copiado" : "Copiar conteúdo"}
        </button>
        {/* link real (não clique disparado por JS) — em algumas janelas o navegador abre
            o arquivo numa aba em vez de baixar direto; nesse caso, salve por lá (Ctrl+S / ⌘S) */}
        <a
          href={urlDownload}
          download={nomeArquivo}
          target="_blank"
          rel="noopener noreferrer"
          style={{ display: "flex", alignItems: "center", gap: 6, padding: "7px 12px", background: "#fff", color: "#334155", border: "1px solid #CBD5E1", borderRadius: 6, fontSize: 12.5, fontWeight: 600, textDecoration: "none" }}
        >
          <Download size={14} /> Baixar {nomeArquivo}
        </a>
      </div>
      <pre style={{ background: "#0F172A", color: "#E2E8F0", padding: 14, borderRadius: 8, fontSize: 11.5, lineHeight: 1.6, overflowX: "auto", margin: 0, fontFamily: "'JetBrains Mono', monospace" }}>{conteudo}</pre>
    </div>
  );
}

// ---------- formulário de preenchimento da NF-e de compra ----------
function FormNfeCompra() {
  const [d, setD] = useState({
    numero: "", dataEmissao: new Date().toISOString().slice(0, 10), origem: "fisica",
    vendedorNome: "", vendedorDocumento: "", emitenteNome: "Revenda Auto Veículos LTDA",
    placa: "", marca: "", modelo: "", ano: "", km: "", valor: "",
  });
  const [xml, setXml] = useState("");
  const campo = (chave, label, largura = 160) => (
    <div>
      <label style={{ fontSize: 11.5, color: "#64748B", fontWeight: 600, display: "block", marginBottom: 4 }}>{label}</label>
      <input value={d[chave]} onChange={(e) => setD({ ...d, [chave]: e.target.value })} style={{ width: largura, padding: "8px 10px", borderRadius: 6, border: "1px solid #CBD5E1", fontSize: 13, boxSizing: "border-box" }} />
    </div>
  );

  return (
    <Card>
      <div style={{ fontWeight: 600, fontSize: 14, marginBottom: 4 }}>Preencher NF-e de compra</div>
      <div style={{ fontSize: 12.5, color: "#64748B", marginBottom: 14 }}>
        Quando o vendedor é pessoa física, a própria revenda é o emitente e o destinatário desta nota (nota fiscal de entrada).
      </div>

      <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 12 }}>
        {campo("numero", "Número da nota", 120)}
        {campo("dataEmissao", "Data de emissão", 140)}
        <div>
          <label style={{ fontSize: 11.5, color: "#64748B", fontWeight: 600, display: "block", marginBottom: 4 }}>Vendedor é</label>
          <div style={{ display: "flex", border: "1px solid #CBD5E1", borderRadius: 6, overflow: "hidden" }}>
            {[["fisica", "Pessoa física"], ["juridica", "Pessoa jurídica"]].map(([valor, label]) => (
              <button key={valor} onClick={() => setD({ ...d, origem: valor })} style={{ padding: "8px 12px", fontSize: 12.5, border: "none", cursor: "pointer", background: d.origem === valor ? "#2563EB" : "#fff", color: d.origem === valor ? "#fff" : "#334155", fontWeight: 600 }}>{label}</button>
            ))}
          </div>
        </div>
      </div>

      <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 12 }}>
        {campo("vendedorNome", "Nome do vendedor", 200)}
        {campo("vendedorDocumento", d.origem === "fisica" ? "CPF do vendedor" : "CNPJ do vendedor", 160)}
        {campo("emitenteNome", "Emitente (sua revenda)", 220)}
      </div>

      <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 14 }}>
        {campo("placa", "Placa", 100)}
        {campo("marca", "Marca", 130)}
        {campo("modelo", "Modelo", 170)}
        {campo("ano", "Ano", 70)}
        {campo("km", "KM", 90)}
        {campo("valor", "Valor pago", 120)}
      </div>

      <button onClick={() => setXml(gerarXmlCompra(d))} style={{ display: "flex", alignItems: "center", gap: 6, padding: "9px 16px", background: "#2563EB", color: "#fff", border: "none", borderRadius: 6, fontSize: 13, fontWeight: 600, cursor: "pointer", marginBottom: xml ? 14 : 0 }}>
        <FileText size={15} /> Gerar XML da nota
      </button>

      {xml && <CaixaArquivo conteudo={xml} nomeArquivo={`nota-compra-${d.placa || "veiculo"}.xml`} mimetype="application/xml" />}
    </Card>
  );
}

// ---------- formulário de preenchimento da NF-e de venda ----------
function FormNfeVenda({ estoque, vendaEmEdicao, onReemitir }) {
  const [d, setD] = useState({
    numero: "", serie: "1", dataEmissao: new Date().toISOString().slice(0, 10),
    emitenteNome: "Revenda Auto Veículos LTDA", emitenteCnpj: "", emitenteIe: "",
    emitenteLogradouro: "", emitenteNumero: "", emitenteBairro: "", emitenteMunicipio: "", emitenteUf: "", emitenteCep: "",
    placa: vendaEmEdicao?.veiculo?.placa || "", modelo: vendaEmEdicao?.veiculo?.modelo || "", valor: vendaEmEdicao?.venda?.valor || "",
    formaPagamento: "À vista",
    cfop: "5102", ncm: "8703.23.10",
    aliquotaIcms: "12", aliquotaPis: "0.65", aliquotaCofins: "3",
    clienteTipoPessoa: "fisica", clienteNome: vendaEmEdicao?.venda?.cliente || "", clienteDocumento: "", clienteIe: "",
    clienteTelefone: "", clienteEmail: "",
    clienteCep: "", clienteLogradouro: "", clienteNumero: "", clienteComplemento: "",
    clienteBairro: "", clienteMunicipio: "", clienteUf: "",
  });
  const [xml, setXml] = useState("");
  const [resultadoReemissao, setResultadoReemissao] = useState(null);
  const campo = (chave, label, largura = 160) => (
    <div>
      <label style={{ fontSize: 11.5, color: "#64748B", fontWeight: 600, display: "block", marginBottom: 4 }}>{label}</label>
      <input value={d[chave]} onChange={(e) => setD({ ...d, [chave]: e.target.value })} style={{ width: largura, padding: "8px 10px", borderRadius: 6, border: "1px solid #CBD5E1", fontSize: 13, boxSizing: "border-box" }} />
    </div>
  );

  const selecionarVeiculo = (id) => {
    const veiculo = estoque.find((v) => v.id === Number(id));
    if (veiculo) setD({ ...d, placa: veiculo.placa, modelo: veiculo.modelo, valor: veiculo.precoVenda });
  };

  const reemitir = () => {
    const resultado = simularResultadoEmissao();
    setResultadoReemissao(resultado);
    onReemitir(vendaEmEdicao.venda.id, resultado);
  };

  const valor = Number(d.valor) || 0;
  const valorIcms = (valor * (Number(d.aliquotaIcms) || 0)) / 100;
  const valorPis = (valor * (Number(d.aliquotaPis) || 0)) / 100;
  const valorCofins = (valor * (Number(d.aliquotaCofins) || 0)) / 100;

  return (
    <Card>
      <div style={{ fontWeight: 600, fontSize: 14, marginBottom: 4 }}>Preencher NF-e de venda</div>
      <div style={{ fontSize: 12.5, color: "#64748B", marginBottom: 14 }}>
        Emitente é sempre a revenda; destinatário é o cliente que está comprando o veículo.
      </div>

      {vendaEmEdicao && !resultadoReemissao && (
        <div style={{ display: "flex", gap: 10, alignItems: "flex-start", background: "#FEF2F2", border: "1px solid #FECACA", borderRadius: 8, padding: "10px 14px", marginBottom: 16 }}>
          <AlertTriangle size={16} color="#DC2626" style={{ marginTop: 1, flexShrink: 0 }} />
          <div style={{ fontSize: 12.5, color: "#991B1B" }}>
            <strong>Reemitindo nota recusada</strong> — {vendaEmEdicao.veiculo?.placa} / {vendaEmEdicao.venda.cliente}.<br />
            Motivo da recusa anterior: {vendaEmEdicao.venda.nfeMotivoRecusa}
          </div>
        </div>
      )}

      {resultadoReemissao && (
        <div style={{ display: "flex", gap: 10, alignItems: "flex-start", background: resultadoReemissao.status === "emitida" ? "#F0FDF4" : "#FEF2F2", border: `1px solid ${resultadoReemissao.status === "emitida" ? "#BBF7D0" : "#FECACA"}`, borderRadius: 8, padding: "10px 14px", marginBottom: 16 }}>
          {resultadoReemissao.status === "emitida" ? <Check size={16} color="#16A34A" style={{ marginTop: 1, flexShrink: 0 }} /> : <AlertTriangle size={16} color="#DC2626" style={{ marginTop: 1, flexShrink: 0 }} />}
          <div style={{ fontSize: 12.5, color: resultadoReemissao.status === "emitida" ? "#166534" : "#991B1B" }}>
            {resultadoReemissao.status === "emitida"
              ? "Nota reemitida com sucesso — veja o novo status na aba \"Notas de venda\"."
              : <>Recusada de novo: {resultadoReemissao.motivo}. Ajuste os dados e tente reemitir mais uma vez.</>}
          </div>
        </div>
      )}

      <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 14 }}>
        {campo("numero", "Número da nota", 110)}
        {campo("serie", "Série", 70)}
        {campo("dataEmissao", "Data de emissão", 140)}
      </div>

      <div style={{ fontSize: 12.5, fontWeight: 600, color: "#334155", marginBottom: 8, paddingTop: 4, borderTop: "1px solid #F1F5F9" }}>Emitente</div>
      <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 10 }}>
        {campo("emitenteNome", "Nome / razão social", 240)}
        {campo("emitenteCnpj", "CNPJ", 160)}
        {campo("emitenteIe", "Inscrição estadual", 150)}
      </div>
      <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 14 }}>
        {campo("emitenteCep", "CEP", 110)}
        {campo("emitenteLogradouro", "Logradouro", 200)}
        {campo("emitenteNumero", "Número", 80)}
        {campo("emitenteBairro", "Bairro", 150)}
        {campo("emitenteMunicipio", "Município", 150)}
        <div>
          <label style={{ fontSize: 11.5, color: "#64748B", fontWeight: 600, display: "block", marginBottom: 4 }}>UF</label>
          <select value={d.emitenteUf} onChange={(e) => setD({ ...d, emitenteUf: e.target.value })} style={{ padding: "8px 10px", borderRadius: 6, border: "1px solid #CBD5E1", fontSize: 13, width: 80 }}>
            <option value="">—</option>
            {UFS.map((uf) => <option key={uf} value={uf}>{uf}</option>)}
          </select>
        </div>
      </div>

      <div style={{ marginBottom: 12 }}>
        <label style={{ fontSize: 11.5, color: "#64748B", fontWeight: 600, display: "block", marginBottom: 4 }}>Veículo do estoque (preenche placa, modelo e valor)</label>
        <select onChange={(e) => selecionarVeiculo(e.target.value)} defaultValue="" style={{ padding: "8px 10px", borderRadius: 6, border: "1px solid #CBD5E1", fontSize: 13, minWidth: 240 }}>
          <option value="">Selecionar veículo...</option>
          {estoque.map((v) => <option key={v.id} value={v.id}>{v.placa} — {v.modelo}</option>)}
        </select>
      </div>

      <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 14 }}>
        {campo("placa", "Placa", 100)}
        {campo("modelo", "Modelo", 170)}
        {campo("valor", "Valor da venda", 130)}
        <div>
          <label style={{ fontSize: 11.5, color: "#64748B", fontWeight: 600, display: "block", marginBottom: 4 }}>Forma de pagamento</label>
          <select value={d.formaPagamento} onChange={(e) => setD({ ...d, formaPagamento: e.target.value })} style={{ padding: "8px 10px", borderRadius: 6, border: "1px solid #CBD5E1", fontSize: 13 }}>
            {["À vista", "Financiado", "Cartão de crédito", "PIX", "Boleto"].map((f) => <option key={f} value={f}>{f}</option>)}
          </select>
        </div>
      </div>

      <div style={{ fontSize: 12.5, fontWeight: 600, color: "#334155", marginBottom: 8, paddingTop: 4, borderTop: "1px solid #F1F5F9" }}>Dados fiscais do item</div>
      <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 10 }}>
        <div style={{ flex: "1 1 260px", minWidth: 0 }}>
          <label style={{ fontSize: 11.5, color: "#64748B", fontWeight: 600, display: "block", marginBottom: 4 }}>CFOP</label>
          <select value={d.cfop} onChange={(e) => setD({ ...d, cfop: e.target.value })} style={{ width: "100%", padding: "8px 10px", borderRadius: 6, border: "1px solid #CBD5E1", fontSize: 13, boxSizing: "border-box" }}>
            {CFOPS_VENDA_VEICULO.map((c) => <option key={c.codigo} value={c.codigo}>{c.descricao}</option>)}
          </select>
        </div>
        {campo("ncm", "NCM do veículo", 130)}
      </div>
      <div style={{ fontSize: 11.5, color: "#94A3B8", marginBottom: 14 }}>
        O CFOP correto depende do estado do cliente e do regime tributário da revenda — confirme com o contador antes de emitir de verdade.
      </div>

      <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 6 }}>
        {campo("aliquotaIcms", "Alíquota ICMS (%)", 130)}
        {campo("aliquotaPis", "Alíquota PIS (%)", 130)}
        {campo("aliquotaCofins", "Alíquota COFINS (%)", 130)}
      </div>
      <div style={{ display: "flex", gap: 16, flexWrap: "wrap", fontSize: 12, color: "#64748B", background: "#F8FAFC", border: "1px solid #E2E8F0", borderRadius: 8, padding: "10px 14px", marginBottom: 16 }}>
        <span>ICMS: <strong style={{ color: "#0F172A" }}>{fmt(valorIcms)}</strong></span>
        <span>PIS: <strong style={{ color: "#0F172A" }}>{fmt(valorPis)}</strong></span>
        <span>COFINS: <strong style={{ color: "#0F172A" }}>{fmt(valorCofins)}</strong></span>
        <span style={{ marginLeft: "auto", color: "#94A3B8" }}>Alíquotas de exemplo — confirme as reais com o contador.</span>
      </div>

      <div style={{ fontSize: 12.5, fontWeight: 600, color: "#334155", marginBottom: 8, paddingTop: 4, borderTop: "1px solid #F1F5F9" }}>Destinatário</div>
      <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 10 }}>
        <div>
          <label style={{ fontSize: 11.5, color: "#64748B", fontWeight: 600, display: "block", marginBottom: 4 }}>Destinatário é</label>
          <div style={{ display: "flex", border: "1px solid #CBD5E1", borderRadius: 6, overflow: "hidden" }}>
            {[["fisica", "Pessoa física"], ["juridica", "Pessoa jurídica"]].map(([valorTipo, label]) => (
              <button key={valorTipo} type="button" onClick={() => setD({ ...d, clienteTipoPessoa: valorTipo })} style={{ padding: "8px 12px", fontSize: 12.5, border: "none", cursor: "pointer", background: d.clienteTipoPessoa === valorTipo ? "#2563EB" : "#fff", color: d.clienteTipoPessoa === valorTipo ? "#fff" : "#334155", fontWeight: 600 }}>{label}</button>
            ))}
          </div>
        </div>
        {campo("clienteNome", "Nome / razão social", 220)}
        {campo("clienteDocumento", d.clienteTipoPessoa === "fisica" ? "CPF" : "CNPJ", 150)}
        {d.clienteTipoPessoa === "juridica" && campo("clienteIe", "Inscrição estadual", 150)}
      </div>
      <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 10 }}>
        {campo("clienteTelefone", "Telefone", 140)}
        {campo("clienteEmail", "E-mail", 220)}
      </div>
      <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 10 }}>
        {campo("clienteCep", "CEP", 110)}
        {campo("clienteLogradouro", "Logradouro", 220)}
        {campo("clienteNumero", "Número", 90)}
        {campo("clienteComplemento", "Complemento", 140)}
      </div>
      <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 14 }}>
        {campo("clienteBairro", "Bairro", 170)}
        {campo("clienteMunicipio", "Município", 170)}
        <div>
          <label style={{ fontSize: 11.5, color: "#64748B", fontWeight: 600, display: "block", marginBottom: 4 }}>UF</label>
          <select value={d.clienteUf} onChange={(e) => setD({ ...d, clienteUf: e.target.value })} style={{ padding: "8px 10px", borderRadius: 6, border: "1px solid #CBD5E1", fontSize: 13, width: 80 }}>
            <option value="">—</option>
            {UFS.map((uf) => <option key={uf} value={uf}>{uf}</option>)}
          </select>
        </div>
      </div>

      <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: xml ? 14 : 0 }}>
        <button onClick={() => setXml(gerarXmlVenda(d))} style={{ display: "flex", alignItems: "center", gap: 6, padding: "9px 16px", background: "#2563EB", color: "#fff", border: "none", borderRadius: 6, fontSize: 13, fontWeight: 600, cursor: "pointer" }}>
          <FileText size={15} /> Gerar XML da nota
        </button>
        {vendaEmEdicao && (
          <button onClick={reemitir} style={{ display: "flex", alignItems: "center", gap: 6, padding: "9px 16px", background: "#DC2626", color: "#fff", border: "none", borderRadius: 6, fontSize: 13, fontWeight: 600, cursor: "pointer" }}>
            <RefreshCcw size={15} /> Reemitir nota
          </button>
        )}
      </div>

      {xml && <CaixaArquivo conteudo={xml} nomeArquivo={`nota-venda-${d.placa || "veiculo"}.xml`} mimetype="application/xml" />}
    </Card>
  );
}

const ABAS_NOTAS = [
  { id: "venda", label: "Notas de venda" },
  { id: "entrada", label: "Notas de entrada (compra PF)" },
  { id: "preencher-compra", label: "Preencher NF-e de compra" },
  { id: "preencher-venda", label: "Preencher NF-e de venda" },
];

// ---------- notas fiscais ----------
function NotasFiscais({ vendas, estoque, setVendas }) {
  const [subAba, setSubAba] = useState("venda");
  const [vendaEmEdicao, setVendaEmEdicao] = useState(null); // { venda, veiculo } — nota recusada sendo reeditada
  const entradasPF = estoque.filter((v) => v.origem === "fisica");

  const emitirNota = (vendaId) => {
    const resultado = simularResultadoEmissao();
    setVendas((atual) => atual.map((v) => (v.id === vendaId ? { ...v, nfeStatus: resultado.status, nfeMotivoRecusa: resultado.motivo } : v)));
  };

  const editarEReemitir = (venda) => {
    const veiculo = estoque.find((v) => v.id === venda.veiculoId);
    setVendaEmEdicao({ venda, veiculo });
    setSubAba("preencher-venda");
  };

  const aplicarReemissao = (vendaId, resultado) => {
    setVendas((atual) => atual.map((v) => (v.id === vendaId ? { ...v, nfeStatus: resultado.status, nfeMotivoRecusa: resultado.motivo } : v)));
  };

  return (
    <div>
      <Titulo sub="Emissão de NF-e e exportação de XML para a Domínio Web">Notas fiscais</Titulo>

      <Card style={{ marginBottom: 16, borderLeft: "3px solid #F59E0B" }}>
        <div style={{ display: "flex", gap: 10, alignItems: "flex-start" }}>
          <AlertTriangle size={18} color="#F59E0B" style={{ marginTop: 2, flexShrink: 0 }} />
          <div style={{ fontSize: 13, lineHeight: 1.6, color: "#334155" }}>
            <strong>Isso exige um backend.</strong> A emissão real de NF-e assina o XML com o certificado digital da empresa e conversa com o webservice da SEFAZ do estado. Os formulários abaixo geram um XML no formato usado por este protótipo, para você testar o front — inclusive importando de volta na aba Estoque. A emissão oficial só acontece quando o backend estiver plugado.
          </div>
        </div>
      </Card>

      <div style={{ display: "flex", gap: 6, marginBottom: 16, flexWrap: "wrap" }}>
        {ABAS_NOTAS.map((item) => (
          <button
            key={item.id}
            onClick={() => { setSubAba(item.id); setVendaEmEdicao(null); }}
            style={{
              padding: "8px 14px", borderRadius: 20, fontSize: 12.5, fontWeight: 600, cursor: "pointer",
              border: subAba === item.id ? "1px solid #2563EB" : "1px solid #E2E8F0",
              background: subAba === item.id ? "#EFF6FF" : "#fff",
              color: subAba === item.id ? "#1D4ED8" : "#334155",
            }}
          >
            {item.label}
          </button>
        ))}
      </div>

      {subAba === "venda" && (
        <Card style={{ padding: 0 }}>
          <div className="rv-table-scroll"><table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <thead>
              <tr style={{ background: "#F8FAFC", textAlign: "left" }}>
                {["Veículo", "Cliente", "Valor", "Status NF-e", "XML (Domínio Web)", ""].map((h) => (
                  <th key={h} style={{ padding: "10px 14px", color: "#64748B", fontWeight: 600, fontSize: 11.5, textTransform: "uppercase", letterSpacing: 0.4 }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {vendas.map((venda) => {
                const veiculo = estoque.find((v) => v.id === venda.veiculoId);
                return (
                  <tr key={venda.id} style={{ borderTop: "1px solid #F1F5F9" }}>
                    <td style={{ padding: "10px 14px" }}>{veiculo ? `${veiculo.placa} — ${veiculo.modelo}` : "—"}</td>
                    <td style={{ padding: "10px 14px" }}>{venda.cliente}</td>
                    <td style={{ padding: "10px 14px" }}>{fmt(venda.valor)}</td>
                    <td style={{ padding: "10px 14px" }}>
                      <Selo status={venda.nfeStatus} />
                      {venda.nfeStatus === "recusada" && venda.nfeMotivoRecusa && (
                        <div style={{ fontSize: 11.5, color: "#991B1B", marginTop: 4, maxWidth: 240 }}>{venda.nfeMotivoRecusa}</div>
                      )}
                    </td>
                    <td style={{ padding: "10px 14px", color: "#64748B" }}>{venda.nfeStatus === "emitida" ? "dominio_web_export.xml" : "aguardando emissão"}</td>
                    <td style={{ padding: "10px 14px" }}>
                      {venda.nfeStatus === "emitida" && <span style={{ fontSize: 12, color: "#94A3B8" }}>—</span>}
                      {venda.nfeStatus === "pendente" && (
                        <button onClick={() => emitirNota(venda.id)} style={{ padding: "6px 12px", borderRadius: 6, border: "1px solid #93C5FD", background: "#EFF6FF", color: "#1D4ED8", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>
                          Emitir
                        </button>
                      )}
                      {venda.nfeStatus === "recusada" && (
                        <button onClick={() => editarEReemitir(venda)} style={{ display: "flex", alignItems: "center", gap: 5, padding: "6px 12px", borderRadius: 6, border: "1px solid #FCA5A5", background: "#FEF2F2", color: "#991B1B", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>
                          <RefreshCcw size={12} /> Editar e reemitir
                        </button>
                      )}
                    </td>
                  </tr>
                );
              })}
              {vendas.length === 0 && (
                <tr><td colSpan={6} style={{ padding: "16px 14px", color: "#94A3B8", fontSize: 13 }}>Nenhuma venda registrada ainda.</td></tr>
              )}
            </tbody>
          </table></div>
        </Card>
      )}

      {subAba === "entrada" && (
        <Card style={{ padding: 0 }}>
          <div className="rv-table-scroll"><table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <thead>
              <tr style={{ background: "#F8FAFC", textAlign: "left" }}>
                {["Veículo", "Vendedor (PF)", "CPF", "Valor", "Emitente", "Tomador", "Status NF-e entrada", ""].map((h) => (
                  <th key={h} style={{ padding: "10px 14px", color: "#64748B", fontWeight: 600, fontSize: 11.5, textTransform: "uppercase", letterSpacing: 0.4 }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {entradasPF.map((v) => (
                <tr key={v.id} style={{ borderTop: "1px solid #F1F5F9" }}>
                  <td style={{ padding: "10px 14px" }}>{v.placa} — {v.modelo}</td>
                  <td style={{ padding: "10px 14px" }}>{v.vendedorNome}</td>
                  <td style={{ padding: "10px 14px", fontFamily: "'JetBrains Mono', monospace" }}>{v.vendedorCpf}</td>
                  <td style={{ padding: "10px 14px" }}>{fmt(Number(v.precoCompra))}</td>
                  <td style={{ padding: "10px 14px", fontSize: 12.5, color: "#64748B" }}>Revenda (própria empresa)</td>
                  <td style={{ padding: "10px 14px", fontSize: 12.5, color: "#64748B" }}>Revenda (própria empresa)</td>
                  <td style={{ padding: "10px 14px" }}><Selo status={v.nfeEntradaStatus} /></td>
                  <td style={{ padding: "10px 14px" }}>
                    <button disabled style={{ padding: "6px 12px", borderRadius: 6, border: "1px solid #E2E8F0", background: "#F8FAFC", color: "#94A3B8", fontSize: 12, cursor: "not-allowed" }}>
                      Emitir (requer backend)
                    </button>
                  </td>
                </tr>
              ))}
              {entradasPF.length === 0 && (
                <tr><td colSpan={8} style={{ padding: "16px 14px", color: "#94A3B8", fontSize: 13 }}>Nenhuma compra de pessoa física cadastrada ainda.</td></tr>
              )}
            </tbody>
          </table></div>
        </Card>
      )}

      {subAba === "preencher-compra" && <FormNfeCompra />}
      {subAba === "preencher-venda" && (
        <FormNfeVenda
          key={vendaEmEdicao?.venda.id || "novo"}
          estoque={estoque}
          vendaEmEdicao={vendaEmEdicao}
          onReemitir={aplicarReemissao}
        />
      )}

      <Card style={{ marginTop: 16 }}>
        <div style={{ fontWeight: 600, fontSize: 14, marginBottom: 8 }}>Como plugar de verdade</div>
        <ol style={{ margin: 0, paddingLeft: 18, fontSize: 13, color: "#334155", lineHeight: 1.8 }}>
          <li>Backend recebe o pedido de emissão desta tela via API própria (ex.: <code>POST /notas/emitir</code>).</li>
          <li>Backend monta o XML da NF-e/NFC-e, assina com o certificado A1 guardado no servidor.</li>
          <li>Backend envia ao webservice da SEFAZ do estado e trata retorno (autorizada, rejeitada, contingência).</li>
          <li>Backend gera o XML já no layout aceito pela importação da Domínio Web e disponibiliza para download/envio automático.</li>
        </ol>
      </Card>
    </div>
  );
}

// ---------- renave ----------
function Renave({ vendas, estoque, marcarRenaveVeiculo }) {
  const [erro, setErro] = useState("");

  const cadastrarNoRenave = async (id) => {
    try {
      await marcarRenaveVeiculo(id);
    } catch (err) {
      setErro(err.message);
    }
  };

  return (
    <div>
      <Titulo sub="Cadastro de veículos e comunicação de venda ao Renave via certificado A1">Integração Renave</Titulo>

      <Card style={{ marginBottom: 16, borderLeft: "3px solid #EF4444" }}>
        <div style={{ display: "flex", gap: 10, alignItems: "flex-start" }}>
          <ShieldCheck size={18} color="#EF4444" style={{ marginTop: 2, flexShrink: 0 }} />
          <div style={{ fontSize: 13, lineHeight: 1.6, color: "#334155" }}>
            <strong>O certificado A1 nunca deve ficar no navegador.</strong> É um arquivo <code>.pfx</code> com a chave privada da empresa — se ele chegar ao cliente, qualquer pessoa com acesso ao computador pode assinar documentos em nome da revenda. A autenticação mTLS com a API do Renave precisa acontecer inteiramente no backend, com o certificado guardado em um cofre de segredos ou HSM. <strong>Nesta versão de demonstração, o botão abaixo só simula o cadastro localmente</strong> — na versão conectada ao backend real, ele chama a API de verdade.
          </div>
        </div>
      </Card>

      {erro && <div style={{ fontSize: 12.5, color: "#DC2626", marginBottom: 16 }}>{erro}</div>}

      <div style={{ fontWeight: 600, fontSize: 14, marginBottom: 10 }}>Cadastro de veículos no Renave</div>
      <div style={{ fontSize: 12.5, color: "#64748B", marginBottom: 10 }}>
        Pelas regras atuais do Renave, todo veículo precisa ser cadastrado assim que entra no estoque — não apenas no momento da venda.
      </div>
      <Card style={{ padding: 0, marginBottom: 20 }}>
        <div className="rv-table-scroll"><table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ background: "#F8FAFC", textAlign: "left" }}>
              {["Veículo", "Placa", "Origem da compra", "Status Renave", ""].map((h) => (
                <th key={h} style={{ padding: "10px 14px", color: "#64748B", fontWeight: 600, fontSize: 11.5, textTransform: "uppercase", letterSpacing: 0.4 }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {estoque.map((v) => (
              <tr key={v.id} style={{ borderTop: "1px solid #F1F5F9" }}>
                <td style={{ padding: "10px 14px" }}>{v.modelo}</td>
                <td style={{ padding: "10px 14px", fontFamily: "'JetBrains Mono', monospace" }}>{v.placa}</td>
                <td style={{ padding: "10px 14px", fontSize: 12.5, color: "#64748B" }}>{v.origem === "fisica" ? "Pessoa física" : "Pessoa jurídica"}</td>
                <td style={{ padding: "10px 14px" }}><Selo status={v.renaveCadastroStatus} /></td>
                <td style={{ padding: "10px 14px" }}>
                  {v.renaveCadastroStatus === "concluido" ? (
                    <span style={{ fontSize: 12, color: "#94A3B8" }}>—</span>
                  ) : (
                    <button onClick={() => cadastrarNoRenave(v.id)} style={{ padding: "6px 12px", borderRadius: 6, border: "1px solid #FDE68A", background: "#FFFBEB", color: "#92400E", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>
                      Cadastrar no Renave
                    </button>
                  )}
                </td>
              </tr>
            ))}
          </tbody>
        </table></div>
      </Card>

      <div style={{ fontWeight: 600, fontSize: 14, marginBottom: 10 }}>Comunicação de venda ao Renave</div>
      <Card style={{ padding: 0 }}>
        <div className="rv-table-scroll"><table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ background: "#F8FAFC", textAlign: "left" }}>
              {["Veículo", "Placa", "Cliente", "Status Renave", ""].map((h) => (
                <th key={h} style={{ padding: "10px 14px", color: "#64748B", fontWeight: 600, fontSize: 11.5, textTransform: "uppercase", letterSpacing: 0.4 }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {vendas.map((venda) => {
              const veiculo = estoque.find((v) => v.id === venda.veiculoId);
              return (
                <tr key={venda.id} style={{ borderTop: "1px solid #F1F5F9" }}>
                  <td style={{ padding: "10px 14px" }}>{veiculo?.modelo || "—"}</td>
                  <td style={{ padding: "10px 14px", fontFamily: "'JetBrains Mono', monospace" }}>{veiculo?.placa || "—"}</td>
                  <td style={{ padding: "10px 14px" }}>{venda.cliente}</td>
                  <td style={{ padding: "10px 14px" }}><Selo status={venda.renaveStatus} /></td>
                  <td style={{ padding: "10px 14px" }}>
                    <button disabled style={{ padding: "6px 12px", borderRadius: 6, border: "1px solid #E2E8F0", background: "#F8FAFC", color: "#94A3B8", fontSize: 12, cursor: "not-allowed" }}>
                      Enviar ao Renave (requer backend)
                    </button>
                  </td>
                </tr>
              );
            })}
            {vendas.length === 0 && (
              <tr><td colSpan={5} style={{ padding: "16px 14px", color: "#94A3B8", fontSize: 13 }}>Nenhuma venda registrada ainda.</td></tr>
            )}
          </tbody>
        </table></div>
      </Card>

      <Card style={{ marginTop: 16 }}>
        <div style={{ fontWeight: 600, fontSize: 14, marginBottom: 8 }}>Fluxo real de integração</div>
        <ol style={{ margin: 0, paddingLeft: 18, fontSize: 13, color: "#334155", lineHeight: 1.8 }}>
          <li>Empresa credencia-se no Renave e obtém acesso à API (homologação e depois produção).</li>
          <li>Backend carrega o certificado A1 de um armazenamento seguro e autentica via mTLS a cada chamada.</li>
          <li>Ao cadastrar um veículo no estoque, o backend registra o veículo no Renave antes de liberá-lo para venda.</li>
          <li>Ao confirmar a venda, o backend comunica a transferência ao Renave com os dados do comprador e vendedor.</li>
          <li>Backend recebe o protocolo/retorno do Renave e atualiza os status mostrados aqui.</li>
        </ol>
      </Card>
    </div>
  );
}
