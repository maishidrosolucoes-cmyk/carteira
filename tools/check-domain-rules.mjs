import { readFileSync } from "node:fs";
import path from "node:path";
import vm from "node:vm";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const projectRoot = path.resolve(__dirname, "..");
const backendPath = path.join(projectRoot, "assents", "motor", "motorbackend.js");

function fail(message) {
  throw new Error(`[domain] ${message}`);
}

function assert(condition, message) {
  if (!condition) fail(message);
}

function parseDetalhes(linha) {
  try {
    return JSON.parse(linha && linha[18] ? linha[18] : "{}");
  } catch (_) {
    fail("DETALHES_JSON invalido na linha consolidada.");
  }
}

function pedidoBase(overrides = {}) {
  return {
    obra: "26.046",
    cliente: "CLIENTE TESTE",
    item: "ITEM TESTE",
    categoria: "VENDA CONTRIBUINTE",
    data_abertura: "2026-05-01",
    p_total: 100,
    ...overrides
  };
}

function getGrupo(context, linhas, obraKey = "26046") {
  const grupos = context.__domain.montarGruposObrasERP(linhas);
  return grupos[obraKey] || null;
}

const source = readFileSync(backendPath, "utf8");
const context = {
  window: {},
  console,
  fetch: async () => {
    throw new Error("check-domain-rules nao deve acessar rede.");
  }
};

vm.runInNewContext(
  `${source}
;globalThis.__domain = {
  montarGruposObrasERP,
  consolidarGrupoObra,
  isPedidoInvalidoERP,
  isPedidoFrustradoExibivelERP,
  gerarMetaPedidosOperacionais,
  gerarMetaCarteiraPedidos,
  gerarMetaTodasConsolidado
};`,
  context,
  { filename: "motorbackend.js" }
);

const { __domain } = context;

assert(__domain.isPedidoInvalidoERP({ status_pedido: "C" }) === true, "status C deve ser invalido.");
assert(__domain.isPedidoFrustradoExibivelERP(pedidoBase({
  numero_pedido: "100001",
  status_pedido: "C",
  condicao: "PROPOSTA FRUSTRADA",
  p_total: 999
})) === true, "status C com condicao frustrada, obra 2026 e valor > 1 deve ser exibivel em Frustradas.");

const grupoSomenteFrustrado = getGrupo(context, [
  pedidoBase({
    numero_pedido: "100001",
    status_pedido: "C",
    situacao_pedido: "P",
    condicao: "PROPOSTA FRUSTRADA",
    p_total: 999
  })
]);
assert(grupoSomenteFrustrado, "obra apenas com status C frustrado valido deve ser agrupada para Frustradas.");

const detalhesSomenteFrustrado = parseDetalhes(__domain.consolidarGrupoObra(grupoSomenteFrustrado));
assert(detalhesSomenteFrustrado.meta_pedidos_operacionais.length === 1, "Frustradas deve receber o pedido C valido.");
assert(detalhesSomenteFrustrado.meta_pedidos_operacionais[0].status_operacional === "FRUSTRADAS", "Pedido C valido deve aparecer como FRUSTRADAS.");
assert(detalhesSomenteFrustrado.meta_pedidos_operacionais[0].valor > 1, "Pedido C frustrado deve manter valor acima de 1.");

const grupoCCondicaoNaoFrustrada = getGrupo(context, [
  pedidoBase({
    numero_pedido: "100010",
    status_pedido: "C",
    situacao_pedido: "P",
    condicao: "PROPOSTA FIRMADA",
    p_total: 999
  })
]);
assert(!grupoCCondicaoNaoFrustrada, "status C sem condicao frustrada nao pode ser exibido.");

const grupoCValorIrrelevante = getGrupo(context, [
  pedidoBase({
    numero_pedido: "100011",
    status_pedido: "C",
    situacao_pedido: "P",
    condicao: "PROPOSTA FRUSTRADA",
    p_total: 1
  })
]);
assert(!grupoCValorIrrelevante, "status C frustrado com valor ate 1 real nao pode ser exibido.");

const grupoCObraFora2026 = getGrupo(context, [
  pedidoBase({
    obra: "25.046",
    numero_pedido: "100012",
    status_pedido: "C",
    situacao_pedido: "P",
    condicao: "PROPOSTA FRUSTRADA",
    p_total: 999
  })
], "25046");
assert(!grupoCObraFora2026, "status C frustrado fora de 2026 nao pode ser exibido.");

const grupoComPedidoValido = getGrupo(context, [
  pedidoBase({
    numero_pedido: "100001",
    status_pedido: "C",
    situacao_pedido: "P",
    condicao: "PROPOSTA FRUSTRADA",
    p_total: 999
  }),
  pedidoBase({
    numero_pedido: "100002",
    status_pedido: "A",
    situacao_pedido: "P",
    condicao: "PROPOSTA FIRMADA",
    data_firmada: "2026-05-03",
    p_total: 500
  })
]);

assert(grupoComPedidoValido, "obra com pedido valido deve permanecer exibivel.");
assert(grupoComPedidoValido.itens.length === 2, "pedido C valido para Frustradas pode ser avaliado junto da obra.");

const linhaValida = __domain.consolidarGrupoObra(grupoComPedidoValido);
const detalhesValidos = parseDetalhes(linhaValida);
const detalhesTexto = JSON.stringify(detalhesValidos);

assert(!detalhesTexto.includes("100001"), "pedido status C nao pode aparecer nos metadados da obra.");
assert(detalhesValidos.meta_carteira_pedidos.length === 1, "carteira deve manter somente pedido valido.");
assert(detalhesValidos.meta_carteira_pedidos[0].numero_pedido === "100002", "carteira nao pode herdar pedido status C.");
assert(detalhesValidos.meta_pedidos_operacionais.length === 1, "abas consolidadas devem usar somente pedidos validos.");
assert(detalhesValidos.meta_pedidos_operacionais[0].status_operacional === "FIRMADAS", "pedido A/P deve entrar como carteira.");

function assertStatusCExcluidoPorPedidoValido(pedidoValido, statusEsperado) {
  const grupo = getGrupo(context, [
    pedidoBase({
      numero_pedido: "100001",
      status_pedido: "C",
      situacao_pedido: "P",
      condicao: "PROPOSTA FRUSTRADA",
      p_total: 999
    }),
    pedidoBase(pedidoValido)
  ]);

  assert(grupo, `obra com pedido ${statusEsperado} deve permanecer exibivel.`);
  const detalhes = parseDetalhes(__domain.consolidarGrupoObra(grupo));
  const texto = JSON.stringify(detalhes);
  const status = detalhes.meta_pedidos_operacionais.map(pedido => pedido.status_operacional);

  assert(!texto.includes("100001"), `pedido C nao pode aparecer quando existe ${statusEsperado}.`);
  assert(status.includes(statusEsperado), `pedido valido deve aparecer como ${statusEsperado}.`);
}

assertStatusCExcluidoPorPedidoValido({
  numero_pedido: "100003",
  status_pedido: "A",
  situacao_pedido: "S",
  condicao: "PROPOSTA ENVIADA",
  data_enviada: "2026-05-04",
  p_total: 300
}, "ENVIADAS");

assertStatusCExcluidoPorPedidoValido({
  numero_pedido: "100004",
  status_pedido: "L",
  situacao_pedido: "P",
  condicao: "ENTREGUE",
  nf: "800",
  data_faturam: "2026-05-22",
  p_total: 400
}, "ENTREGUE");

const grupoParcial = getGrupo(context, [
  pedidoBase({
    obra: "26.088",
    numero_pedido: "834",
    status_pedido: "A",
    situacao_pedido: "P",
    condicao: "PROPOSTA FIRMADA",
    data_firmada: "2026-05-07",
    p_total: 26570
  }),
  pedidoBase({
    obra: "26.088",
    numero_pedido: "100246",
    status_pedido: "L",
    situacao_pedido: "P",
    condicao: "ENTREGUE",
    nf: "800",
    data_faturam: "2026-05-22",
    p_total: 37412.2
  })
], "26088");

assert(grupoParcial, "obra parcial deve permanecer agrupada por obra.");
const detalhesParcial = parseDetalhes(__domain.consolidarGrupoObra(grupoParcial));
const statusParcial = detalhesParcial.meta_pedidos_operacionais
  .map(pedido => pedido.status_operacional)
  .sort();

assert(statusParcial.includes("FIRMADAS"), "obra parcial deve preservar pedido valido em carteira.");
assert(statusParcial.includes("ENTREGUE"), "obra parcial deve preservar pedido faturado em entregue.");
assert(detalhesParcial.meta_carteira_pedidos.length === 1, "carteira deve exibir somente a parte ainda firmada.");
assert(detalhesParcial.meta_concluidas_nf.length === 1, "entregue deve consolidar a NF faturada.");

console.log("[domain] Regras de obra/pedido validadas.");
