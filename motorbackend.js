// ==========================================
// motorbackend.js
// CONEXÃO LOCAL (NODE.JS) - SUBSTITUINDO O SUPABASE
// ARQUITETURA: ERP-FIRST com OVERRIDE MANUAL e EXIBIÇÃO TOTAL
// ==========================================

const ITENS_ORDEM = ["BBA/ELET.", "MT", "FLUT.", "M FV.", "AD. FLEX", "AD. RIG.", "FIXADORES", "SIST. ELÉT.", "PEÇAS REP.", "SERV.", "MONT.", "FATUR."];

function getSafeId(str) {
  if (!str) return "";
  return str.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9]/g, '_');
}

function normalizeSpaces(value) {
  return String(value || "").replace(/\s+/g, " ").trim();
}

function normalizeCompareText(value) {
  return normalizeSpaces(value)
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

function parseDateSafe(value) {
  if (!value) return null;
  if (value instanceof Date && !isNaN(value.getTime())) return new Date(value.getTime());

  const txt = String(value).trim();
  if (!txt) return null;

  let m = txt.match(/^(\d{2})\/(\d{2})\/(\d{2,4})$/);
  if (m) {
    const ano = Number(m[3].length === 2 ? `20${m[3]}` : m[3]);
    const dt = new Date(ano, Number(m[2]) - 1, Number(m[1]), 12, 0, 0, 0);
    return isNaN(dt.getTime()) ? null : dt;
  }

  m = txt.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) {
    const dt = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12, 0, 0, 0);
    return isNaN(dt.getTime()) ? null : dt;
  }

  const dt = new Date(txt);
  if (isNaN(dt.getTime())) return null;
  dt.setHours(12, 0, 0, 0);
  return dt;
}

function getYearFromDate(value) {
  const dt = parseDateSafe(value);
  if (!dt) return "";
  return String(dt.getFullYear()).slice(-2);
}

function parseMoneyFlexible(value) {
  if (value === null || value === undefined || value === "") return 0;
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;

  let str = String(value).trim();
  if (!str) return 0;

  str = str.replace(/\s/g, "").replace(/[R$r$\u00A0]/g, "");

  if (str.includes(",")) {
    str = str.replace(/\./g, "").replace(",", ".");
  } else {
    const dotCount = (str.match(/\./g) || []).length;
    if (dotCount > 1) str = str.replace(/\./g, "");
  }

  str = str.replace(/[^\d.-]/g, "");
  const n = parseFloat(str);
  return Number.isFinite(n) ? n : 0;
}

function joinUniqueText(baseValue, newValue) {
  const items = [];

  function pushParts(value) {
    normalizeSpaces(value)
      .split(/\s*\/\s*/)
      .map(v => normalizeSpaces(v))
      .filter(Boolean)
      .forEach(part => {
        const key = normalizeCompareText(part);
        if (!key) return;
        if (!items.some(existing => normalizeCompareText(existing) === key)) {
          items.push(part);
        }
      });
  }

  pushParts(baseValue);
  pushParts(newValue);

  return items.join(" / ");
}

function joinUniqueNF(baseValue, newValue) {
  const items = [];

  function pushParts(value) {
    normalizeSpaces(value)
      .split(/[\/,;|]+/)
      .map(v => normalizeSpaces(v))
      .filter(Boolean)
      .forEach(part => {
        const key = normalizeCompareText(part);
        if (!key) return;
        if (!items.some(existing => normalizeCompareText(existing) === key)) {
          items.push(part);
        }
      });
  }

  pushParts(baseValue);
  pushParts(newValue);

  return items.join(" / ");
}

function pickFirstFilled(currentValue, nextValue) {
  return normalizeSpaces(currentValue) || normalizeSpaces(nextValue) || "";
}

function pickLongerFilled(currentValue, nextValue) {
  const current = normalizeSpaces(currentValue);
  const next = normalizeSpaces(nextValue);
  if (!current) return next;
  if (!next) return current;
  return next.length > current.length ? next : current;
}

function pickEarlierDate(currentValue, nextValue) {
  const a = parseDateSafe(currentValue);
  const b = parseDateSafe(nextValue);
  if (!a && !b) return normalizeSpaces(currentValue) || normalizeSpaces(nextValue) || "";
  if (!a) return normalizeSpaces(nextValue);
  if (!b) return normalizeSpaces(currentValue);
  return a.getTime() <= b.getTime() ? normalizeSpaces(currentValue) : normalizeSpaces(nextValue);
}

function pickLaterDate(currentValue, nextValue) {
  const a = parseDateSafe(currentValue);
  const b = parseDateSafe(nextValue);
  if (!a && !b) return normalizeSpaces(currentValue) || normalizeSpaces(nextValue) || "";
  if (!a) return normalizeSpaces(nextValue);
  if (!b) return normalizeSpaces(currentValue);
  return a.getTime() >= b.getTime() ? normalizeSpaces(currentValue) : normalizeSpaces(nextValue);
}

function inferirAnoPorDatas(erp) {
  return (
    getYearFromDate(erp.data_abertura) ||
    getYearFromDate(erp.data_firmada) ||
    getYearFromDate(erp.data_enviada) ||
    getYearFromDate(erp.data_faturam || erp.data_faturamento) ||
    getYearFromDate(erp.data_frustrada) ||
    ""
  );
}

function extrairNumeroObraCanonico(valorObra) {
  const raw = normalizeSpaces(valorObra);
  if (!raw) {
    return { numero: "", ano: "", bruto: "" };
  }

  const padroes = [
    /(?:^|[^0-9])((20\d{2})[.\-\/](\d{3,5}))(?!\d)/i,
    /(?:^|[^0-9])((\d{2})[.\-\/](\d{3,5}))(?!\d)/i,
    /(?:^|[^0-9])((\d{2})\s+(\d{3,5}))(?!\d)/i
  ];

  for (const regex of padroes) {
    const match = raw.match(regex);
    if (!match) continue;

    let ano = String(match[2] || "").trim();
    const sequencia = String(match[3] || "").trim();

    if (!ano || !sequencia) continue;
    if (ano.length === 4) ano = ano.slice(-2);

    const seqNorm = String(Number(sequencia)).padStart(3, "0");
    return {
      numero: `${ano}.${seqNorm}`,
      ano,
      bruto: raw
    };
  }

  return { numero: "", ano: "", bruto: raw };
}

function resolverIdentidadeObra(erp) {
  const obraBruta = normalizeSpaces(erp.obra);
  const extraida = extrairNumeroObraCanonico(obraBruta);
  const anoInferido = extraida.ano || inferirAnoPorDatas(erp) || "";
  const clienteNorm = normalizeCompareText(erp.cliente || "").slice(0, 120);
  const obraFallback = normalizeCompareText(obraBruta).slice(0, 180);

  let chave = "";
  let exibicao = "";

  if (extraida.numero) {
    chave = `obra:${extraida.numero}`;
    exibicao = extraida.numero;
  } else {
    chave = `legado:${anoInferido || "sem-ano"}:${clienteNorm}:${obraFallback}`;
    exibicao = obraBruta || `OBRA ${anoInferido || "S/ANO"}`;
  }

  return {
    obraExibicao: exibicao,
    obraChave: chave,
    anoObra: anoInferido,
    obraBruta
  };
}

function calcularStatusProposta(erp) {
  let statusProposta = "ENVIADAS";
  const etapaUp = String(erp.etapa || "").toUpperCase();

  if (erp.data_frustrada) {
    statusProposta = "FRUSTRADAS";
  } else if (etapaUp.includes("CONCLU") || erp.data_faturam || erp.data_faturamento) {
    statusProposta = "CONCLUIDAS";
  } else if (etapaUp.includes("ENTREGUE")) {
    statusProposta = "ENTREGUES";
  } else if (erp.data_firmada) {
    statusProposta = "FIRMADAS";
  }

  return statusProposta;
}

function getStatusWeight(status) {
  const mapa = {
    ENVIADAS: 1,
    FIRMADAS: 2,
    ENTREGUES: 3,
    CONCLUIDAS: 4,
    FRUSTRADAS: 5
  };
  return mapa[String(status || "").trim()] || 0;
}

function atualizarStatusMaisForte(linhaExistente, statusNovo) {
  const atual = String(linhaExistente[22] || "").trim();
  if (getStatusWeight(statusNovo) >= getStatusWeight(atual)) {
    linhaExistente[22] = statusNovo;
  }
}

const motorBackend = {

  sincronizarEFetch: async function(anoFiltro = 'TODOS') {
    try {
      // 1. Conecta no servidor da empresa usando o Túnel Cloudflare (Seguro, HTTPS e Público)
      const response = await fetch('https://thumbzilla-modern-refrigerator-simon.trycloudflare.com/api/carteira');

      if (!response.ok) {
        throw new Error('Erro ao conectar no servidor. Verifique se o túnel e o motor estão rodando.');
      }

      const erpData = await response.json();

      // 2. Prepara o cabeçalho que o script.js espera ler
      const resultado = [
        ["DATA", "OBRA", "CLIENTE", "VALOR", "DIAS PRAZO", ...ITENS_ORDEM, "OBSERVAÇÕES", "DETALHES_JSON", "CPMV", "ITEM", "CATEGORIA"]
      ];

      // Dicionário (memória) para evitar duplicação visual de obras
      const obrasProcessadas = {};

      // 3. Varre os dados do JSON e traduz para a matriz do painel
      if (Array.isArray(erpData) && erpData.length > 0) {
        erpData.forEach(erp => {
          const identidade = resolverIdentidadeObra(erp);
          if (!identidade.obraExibicao) return;

          // Filtro de ano usando prefixo canônico ou datas do ERP
          if (anoFiltro !== 'TODOS' && identidade.anoObra !== String(anoFiltro)) {
            return;
          }

          const chaveObra = identidade.obraChave;
          const valorERP = erp.p_total !== null ? parseMoneyFlexible(erp.p_total) : 0;
          const statusProposta = calcularStatusProposta(erp);

          if (obrasProcessadas[chaveObra]) {
            const linhaExistente = obrasProcessadas[chaveObra];

            linhaExistente[3] = parseMoneyFlexible(linhaExistente[3]) + valorERP;
            linhaExistente[20] = joinUniqueText(linhaExistente[20], erp.item || "");
            linhaExistente[21] = joinUniqueText(linhaExistente[21], erp.categoria || "");
            linhaExistente[29] = joinUniqueNF(linhaExistente[29], erp.nf || "");

            linhaExistente[0] = pickEarlierDate(linhaExistente[0], erp.data_firmada || "");
            linhaExistente[4] = pickFirstFilled(linhaExistente[4], erp.praz || erp.pz || "");
            linhaExistente[17] = pickLongerFilled(linhaExistente[17], erp.observacoes || erp.obs || "");
            linhaExistente[19] = pickFirstFilled(linhaExistente[19], erp.cpmv || 0);

            linhaExistente[2] = pickFirstFilled(linhaExistente[2], erp.cliente || "");
            linhaExistente[23] = pickEarlierDate(linhaExistente[23], erp.data_abertura || "");
            linhaExistente[24] = pickFirstFilled(linhaExistente[24], erp.segmento || "");
            linhaExistente[25] = pickFirstFilled(linhaExistente[25], erp.vendedor || erp.responsavel || "");
            linhaExistente[26] = pickFirstFilled(linhaExistente[26], erp.complexidade || "");
            linhaExistente[27] = pickFirstFilled(linhaExistente[27], erp.uf || "");
            linhaExistente[28] = pickLongerFilled(linhaExistente[28], erp.etapa || "");
            linhaExistente[30] = pickEarlierDate(linhaExistente[30], erp.data_frustrada || "");
            linhaExistente[31] = pickEarlierDate(linhaExistente[31], erp.data_enviada || "");
            linhaExistente[32] = pickLaterDate(linhaExistente[32], erp.data_faturam || erp.data_faturamento || "");

            atualizarStatusMaisForte(linhaExistente, statusProposta);
            return;
          }

          const novaLinha = [
            erp.data_firmada || "", // 0: DATA FIRMADA
            identidade.obraExibicao, // 1: OBRA EXIBIDA / CANÔNICA
            erp.cliente || "", // 2: CLIENTE
            valorERP, // 3: VALOR
            erp.praz || erp.pz || "", // 4: DIAS_PRAZO

            // 5 a 16: Itens de controle em branco
            "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A",

            normalizeSpaces(erp.observacoes || erp.obs || ""), // 17: OBSERVAÇÕES
            "{}", // 18: DETALHES JSON
            erp.cpmv || 0, // 19: CPMV
            normalizeSpaces(erp.item || ""), // 20: ITEM
            normalizeSpaces(erp.categoria || ""), // 21: CATEGORIA

            // 22 a 32: INFORMAÇÕES EXTRAS
            statusProposta, // 22: STATUS GERAL DA PROPOSTA
            erp.data_abertura || "", // 23: ABERTURA
            erp.segmento || "", // 24: SEGMENTO
            erp.vendedor || erp.responsavel || "", // 25: RESPONSAVEL
            erp.complexidade || "", // 26: COMPLEXIDADE
            erp.uf || "", // 27: UF
            erp.etapa || "", // 28: ETAPA
            normalizeSpaces(erp.nf || ""), // 29: NF
            erp.data_frustrada || "", // 30: FRUSTRADA
            erp.data_enviada || "", // 31: ENVIADA
            erp.data_faturam || erp.data_faturamento || "" // 32: FATURAMENTO
          ];

          obrasProcessadas[chaveObra] = novaLinha;
        });

        // Ordenação crescente e definitiva
        const listaObras = Object.values(obrasProcessadas);
        listaObras.sort((a, b) => {
          return String(a[1] || "").localeCompare(String(b[1] || ""), 'pt-BR', { numeric: true });
        });

        listaObras.forEach(linha => resultado.push(linha));
      }

      return resultado;

    } catch (e) {
      console.error("Erro na comunicação local:", e);
      throw e;
    }
  },

  salvarProjeto: async function(obj) {
    console.log("Simulação local de salvamento:", obj);
    return "✅ (Modo Local) Dados processados na sessão!";
  },

  getResumoGeralObra: async function(numObra) {
    return { encontrado: false };
  },

  getDadosGeralSimplificado: async function(numObra) {
    return null;
  },

  excluirObra: async function(numObra) {
    return "🗑️ (Modo Local) Simulação de exclusão concluída.";
  }
};

window.motorBackend = motorBackend;
