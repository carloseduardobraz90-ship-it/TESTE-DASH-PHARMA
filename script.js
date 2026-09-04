/* ============================================================
 * VARIÁVEIS GLOBAIS DE ESTADO
 * ============================================================ */
let dadosBrutos = [];
let dadosFiltrados = [];
let graficoRegiaoInstance = null;
let graficoTipoInstance = null;

/* Mapeamento amigável de siglas de regiões para nomes por extenso */
const DEPARA_REGIOES = {
    'SP': 'São Paulo',
    'RJ': 'Rio de Janeiro',
    'MG': 'Minas Gerais',
    'ES': 'Espírito Santo',
    'PR': 'Paraná',
    'SC': 'Santa Catarina',
    'RS': 'Rio Grande do Sul',
    'BA': 'Bahia',
    'PE': 'Pernambuco',
    'CE': 'Ceará',
    'GO': 'Goiás',
    'DF': 'Distrito Federal'
};

/* ============================================================
 * EVENTOS INICIAIS DA PÁGINA
 * ============================================================ */
document.addEventListener("DOMContentLoaded", () => {
    const btnUpdate = document.getElementById("btnUpdate");
    const btnClear = document.getElementById("btnClear");
    const excelFileInput = document.getElementById("excelFile");
    const statusFilter = document.getElementById("statusFilter");
    const regionFilter = document.getElementById("regionFilter");
    const startDate = document.getElementById("startDate");
    const endDate = document.getElementById("endDate");

    if (excelFileInput) {
        excelFileInput.addEventListener("change", processarArquivoExcel);
    }

    if (btnUpdate) {
        btnUpdate.addEventListener("click", processarEAtualizar);
    }

    if (btnClear) {
        btnClear.addEventListener("click", limparFiltros);
    }

    // Atualização automática ao alterar filtros
    if (statusFilter) statusFilter.addEventListener("change", processarEAtualizar);
    if (regionFilter) regionFilter.addEventListener("change", processarEAtualizar);
    if (startDate) startDate.addEventListener("change", processarEAtualizar);
    if (endDate) endDate.addEventListener("change", processarEAtualizar);
});

/* ============================================================
 * FUNÇÕES AUXILIARES DE TRATAMENTO DE COLUNAS E STRINGS
 * ============================================================ */

// Normaliza texto para busca (remove acentos, espaços e deixa em minúsculo)
function normalizarChave(str) {
    return String(str || "")
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .toLowerCase()
        .replace(/[^a-z0-9]/g, "");
}

// Procura o valor de uma coluna independentemente de como foi escrita na planilha
function extrairValorColuna(row, nomesPossiveis) {
    if (!row) return "";
    const chaves = Object.keys(row);

    for (const nome of nomesPossiveis) {
        const nomeAlvo = normalizarChave(nome);
        const chaveEncontrada = chaves.find(k => normalizarChave(k) === nomeAlvo);
        if (chaveEncontrada && row[chaveEncontrada] !== undefined && row[chaveEncontrada] !== "") {
            return String(row[chaveEncontrada]).trim();
        }
    }
    return "";
}

function nomeRegiao(sigla) {
    if (!sigla) return 'Outros';
    const chave = String(sigla).toUpperCase().trim();
    return DEPARA_REGIOES[chave] || sigla;
}

/* ============================================================
 * TRATAMENTO DE DATAS
 * ============================================================ */
function obterTimestampZerado(valorData) {
    if (!valorData) return null;

    let d = null;

    // Se for número serial de data do Excel
    if (typeof valorData === 'number') {
        d = new Date(Math.round((valorData - 25569) * 86400 * 1000));
    } else if (typeof valorData === 'string') {
        const partes = valorData.split(/[-/]/);
        if (partes.length === 3) {
            if (partes[0].length === 4) {
                // Formato YYYY-MM-DD
                d = new Date(partes[0], partes[1] - 1, partes[2]);
            } else {
                // Formato DD/MM/YYYY
                d = new Date(partes[2], partes[1] - 1, partes[0]);
            }
        } else {
            d = new Date(valorData);
        }
    } else if (valorData instanceof Date) {
        d = valorData;
    }

    if (!d || isNaN(d.getTime())) return null;

    d.setHours(0, 0, 0, 0);
    return d.getTime();
}

function formatarData(timestamp) {
    if (!timestamp || timestamp === 9999999999999) return "Sem Data";
    const d = new Date(timestamp);
    const dia = String(d.getDate()).padStart(2, '0');
    const mes = String(d.getMonth() + 1).padStart(2, '0');
    const ano = d.getFullYear();
    return `${dia}/${mes}/${ano}`;
}

/* ============================================================
 * LEITURA DO ARQUIVO EXCEL / CSV
 * ============================================================ */
function processarArquivoExcel(event) {
    const file = event.target.files[0];
    const statusBase = document.getElementById("statusBase");
    const statusIcon = document.getElementById("statusIcon");
    const statusMensagem = document.getElementById("statusMensagem");
    const statusDetalhes = document.getElementById("statusDetalhes");
    const painelErro = document.getElementById("painelErro");

    if (!file) return;

    if (painelErro) painelErro.style.display = "none";

    const reader = new FileReader();

    reader.onload = function (e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: "array" });

            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];

            dadosBrutos = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

            if (dadosBrutos.length === 0) {
                throw new Error("A planilha selecionada está vazia.");
            }

            if (statusIcon) statusIcon.innerText = "🟢";
            if (statusMensagem) statusMensagem.innerText = "Base carregada com sucesso!";
            if (statusDetalhes) statusDetalhes.innerText = `${dadosBrutos.length} registros importados.`;

            popularFiltrosSelect(dadosBrutos);
            processarEAtualizar();

        } catch (error) {
            if (statusIcon) statusIcon.innerText = "🔴";
            if (statusMensagem) statusMensagem.innerText = "Erro ao carregar a base";
            if (statusDetalhes) statusDetalhes.innerText = "Não foi possível ler o arquivo enviado.";

            if (painelErro) {
                painelErro.style.display = "block";
                document.getElementById("erroDetalhes").innerText = error.message || "Erro desconhecido ao processar planilha.";
            }
        }
    };

    reader.readAsArrayBuffer(file);
}

/* ============================================================
 * POPULAR SELECTS DE FILTROS DINAMICAMENTE
 * ============================================================ */
function popularFiltrosSelect(dados) {
    const statusFilter = document.getElementById("statusFilter");
    const regionFilter = document.getElementById("regionFilter");

    if (!statusFilter || !regionFilter) return;

    const statusSet = new Set();
    const regionSet = new Set();

    dados.forEach((row) => {
        const st = extrairValorColuna(row, ["STATUS", "SITUACAO", "STATE", "SITUAÇÃO"]);
        const rg = extrairValorColuna(row, ["REGIAO", "REGIAÕ", "UF", "ESTADO", "REGIONAL"]);

        if (st) statusSet.add(st);
        if (rg) regionSet.add(rg);
    });

    statusFilter.innerHTML = '<option value="TODOS">Todos os Status</option>';
    Array.from(statusSet).sort().forEach((st) => {
        statusFilter.innerHTML += `<option value="${st}">${st}</option>`;
    });

    regionFilter.innerHTML = '<option value="TODAS">Todas as Regiões</option>';
    Array.from(regionSet).sort().forEach((rg) => {
        regionFilter.innerHTML += `<option value="${rg}">${nomeRegiao(rg)}</option>`;
    });
}

/* ============================================================
 * FILTRAGEM E ATUALIZAÇÃO GERAL
 * ============================================================ */
function processarEAtualizar() {
    if (!dadosBrutos || dadosBrutos.length === 0) return;

    const statusSel = document.getElementById("statusFilter")?.value || "TODOS";
    const regionSel = document.getElementById("regionFilter")?.value || "TODAS";
    const startVal = document.getElementById("startDate")?.value;
    const endVal = document.getElementById("endDate")?.value;

    const startTimestamp = obterTimestampZerado(startVal);
    const endTimestamp = obterTimestampZerado(endVal);

    dadosFiltrados = dadosBrutos.filter((row) => {
        const st = extrairValorColuna(row, ["STATUS", "SITUACAO", "STATE", "SITUAÇÃO"]);
        if (statusSel !== "TODOS" && st !== statusSel) return false;

        const rg = extrairValorColuna(row, ["REGIAO", "REGIAÕ", "UF", "ESTADO", "REGIONAL"]);
        if (regionSel !== "TODAS" && rg !== regionSel) return false;

        const rawData = extrairValorColuna(row, ["DATA", "DATE", "DATA AGENDAMENTO", "AGENDAMENTO", "DATA_AGENDAMENTO"]);
        const dataTimestamp = obterTimestampZerado(rawData);

        if (startTimestamp && (!dataTimestamp || dataTimestamp < startTimestamp)) return false;
        if (endTimestamp && (!dataTimestamp || dataTimestamp > endTimestamp)) return false;

        return true;
    });

    atualizarKPIs(dadosFiltrados);
    atualizarGraficos(dadosFiltrados);
    atualizarTabelaAgenda(dadosFiltrados);
}

/* ============================================================
 * ATUALIZAÇÃO DOS KPIS
 * ============================================================ */
function atualizarKPIs(dados) {
    const painelKPIs = document.getElementById("painelKPIs");
    if (!painelKPIs) return;

    let total = dados.length;
    let pendentes = 0;
    let concluidos = 0;

    dados.forEach((row) => {
        const st = normalizarChave(extrairValorColuna(row, ["STATUS", "SITUACAO", "STATE", "SITUAÇÃO"]));
        if (st.includes("conclu") || st.includes("entreg") || st.includes("finaliz") || st.includes("ok")) {
            concluidos++;
        } else if (st.includes("pend") || st.includes("andament") || st.includes("aguard") || st.includes("abert")) {
            pendentes++;
        }
    });

    painelKPIs.innerHTML = `
        <div class="kpi-card">
            <h4>Total de Pedidos</h4>
            <span class="kpi-value">${total}</span>
        </div>
        <div class="kpi-card">
            <h4>Pendentes</h4>
            <span class="kpi-value">${pendentes}</span>
        </div>
        <div class="kpi-card">
            <h4>Concluídos</h4>
            <span class="kpi-value">${concluidos}</span>
        </div>
    `;
}

/* ============================================================
 * ATUALIZAÇÃO DOS GRÁFICOS (CHART.JS)
 * ============================================================ */
function atualizarGraficos(dados) {
    const ctxRegiao = document.getElementById("chartRegiao")?.getContext("2d");
    const ctxTipo = document.getElementById("chartTipo")?.getContext("2d");

    const regiaoCounts = {};
    const tipoCounts = {};

    dados.forEach((row) => {
        const rgBruta = extrairValorColuna(row, ["REGIAO", "REGIAÕ", "UF", "ESTADO", "REGIONAL"]) || "N/A";
        const rg = nomeRegiao(rgBruta);

        const tp = extrairValorColuna(row, ["TIPO", "OPERACAO", "OPERAÇÃO", "TIPO OPERACAO", "SERVICO", "SERVIÇO"]) || "Outros";

        regiaoCounts[rg] = (regiaoCounts[rg] || 0) + 1;
        tipoCounts[tp] = (tipoCounts[tp] || 0) + 1;
    });

    if (ctxRegiao) {
        if (graficoRegiaoInstance) graficoRegiaoInstance.destroy();
        graficoRegiaoInstance = new Chart(ctxRegiao, {
            type: "bar",
            data: {
                labels: Object.keys(regiaoCounts),
                datasets: [{
                    label: "Volume por Região",
                    data: Object.values(regiaoCounts),
                    backgroundColor: "#2563eb"
                }]
            },
            options: { responsive: true, maintainAspectRatio: false }
        });
    }

    if (ctxTipo) {
        if (graficoTipoInstance) graficoTipoInstance.destroy();
        graficoTipoInstance = new Chart(ctxTipo, {
            type: "doughnut",
            data: {
                labels: Object.keys(tipoCounts),
                datasets: [{
                    data: Object.values(tipoCounts),
                    backgroundColor: ["#2563eb", "#16a34a", "#dc2626", "#eab308", "#9333ea", "#64748b"]
                }]
            },
            options: { responsive: true, maintainAspectRatio: false }
        });
    }
}

/* ============================================================
 * ATUALIZAÇÃO DA TABELA DE AGENDA
 * ============================================================ */
function atualizarTabelaAgenda(dados) {
    const tbody = document.querySelector("#tabelaAgenda tbody");
    if (!tbody) return;

    const grupos = {};

    dados.forEach((row) => {
        const rawData = extrairValorColuna(row, ["DATA", "DATE", "DATA AGENDAMENTO", "AGENDAMENTO", "DATA_AGENDAMENTO"]);
        const ts = obterTimestampZerado(rawData) || 9999999999999;
        const regiao = extrairValorColuna(row, ["REGIAO", "REGIAÕ", "UF", "ESTADO", "REGIONAL"]) || "N/A";
        const status = extrairValorColuna(row, ["STATUS", "SITUACAO", "STATE", "SITUAÇÃO"]) || "Indefinido";

        const chave = `${ts}_${regiao}_${status}`;

        if (!grupos[chave]) {
            grupos[chave] = {
                timestamp: ts,
                data: formatarData(ts),
                regiao: regiao,
                status: status,
                quantidade: 0
            };
        }

        grupos[chave].quantidade++;
    });

    const gruposOrdenados = Object.values(grupos).sort((a, b) => {
        if (a.timestamp !== b.timestamp) {
            return a.timestamp - b.timestamp;
        }
        return a.regiao.localeCompare(b.regiao);
    });

    tbody.innerHTML = "";

    if (gruposOrdenados.length === 0) {
        tbody.innerHTML = `<tr><td colspan="4" style="text-align:center;">Nenhum registro encontrado para os filtros selecionados.</td></tr>`;
        return;
    }

    gruposOrdenados.forEach((grupo) => {
        const tr = document.createElement("tr");

        tr.innerHTML = `
            <td>${grupo.data}</td>
            <td>${nomeRegiao(grupo.regiao)}</td>
            <td><span class="badge status-${normalizarChave(grupo.status)}">${grupo.status}</span></td>
            <td><strong>${grupo.quantidade}</strong></td>
        `;

        tbody.appendChild(tr);
    });
}

/* ============================================================
 * LIMPEZA DE FILTROS
 * ============================================================ */
function limparFiltros() {
    const statusFilter = document.getElementById("statusFilter");
    const regionFilter = document.getElementById("regionFilter");
    const startDate = document.getElementById("startDate");
    const endDate = document.getElementById("endDate");

    if (statusFilter) statusFilter.value = "TODOS";
    if (regionFilter) regionFilter.value = "TODAS";
    if (startDate) startDate.value = "";
    if (endDate) endDate.value = "";

    processarEAtualizar();
}