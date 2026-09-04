/* ============================================================
 * VARIÁVEIS GLOBAIS DE ESTADO
 * ============================================================ */
let dadosBrutos = [];
let dadosFiltrados = [];
let graficoTransportadoraInstance = null;
let graficoStatusInstance = null;

/* ============================================================
 * EVENTOS INICIAIS DA PÁGINA
 * ============================================================ */
document.addEventListener("DOMContentLoaded", () => {
    const btnUpdate = document.getElementById("btnUpdate");
    const btnClear = document.getElementById("btnClear");
    const excelFileInput = document.getElementById("excelFile");
    const statusFilter = document.getElementById("statusFilter");
    const transportadoraFilter = document.getElementById("transportadoraFilter");
    const monthFilter = document.getElementById("monthFilter");

    if (excelFileInput) {
        excelFileInput.addEventListener("change", processarArquivoExcel);
    }

    if (btnUpdate) {
        btnUpdate.addEventListener("click", processarEAtualizar);
    }

    if (btnClear) {
        btnClear.addEventListener("click", limparFiltros);
    }

    if (statusFilter) statusFilter.addEventListener("change", processarEAtualizar);
    if (transportadoraFilter) transportadoraFilter.addEventListener("change", processarEAtualizar);
    if (monthFilter) monthFilter.addEventListener("change", processarEAtualizar);
});

/* ============================================================
 * FUNÇÕES AUXILIARES DE TRATAMENTO DE COLUNAS E STRINGS
 * ============================================================ */
function normalizarChave(str) {
    return String(str || "")
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .toLowerCase()
        .replace(/[^a-z0-9]/g, "");
}

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

function formatarMoeda(valor) {
    return Number(valor || 0).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
}

/* ============================================================
 * LEITURA DO ARQUIVO EXCEL
 * ============================================================ */
function processarArquivoExcel(event) {
    const file = event.target.files[0];
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

            // Busca a aba BASE_COTACOES ou pega a primeira se não existir
            let sheetName = workbook.SheetNames.find(s => normalizarChave(s) === "basecotacoes") || workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];

            dadosBrutos = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

            if (dadosBrutos.length === 0) {
                throw new Error("A planilha selecionada está vazia.");
            }

            if (statusIcon) statusIcon.innerText = "🟢";
            if (statusMensagem) statusMensagem.innerText = "Base de Fretes Carregada!";
            if (statusDetalhes) statusDetalhes.innerText = `${dadosBrutos.length} linhas de cotação lidas.`;

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
 * POPULAR FILTROS SELECT DINAMICAMENTE
 * ============================================================ */
function popularFiltrosSelect(dados) {
    const statusFilter = document.getElementById("statusFilter");
    const transportadoraFilter = document.getElementById("transportadoraFilter");
    const monthFilter = document.getElementById("monthFilter");

    const statusSet = new Set();
    const transpSet = new Set();
    const mesSet = new Set();

    dados.forEach((row) => {
        const st = extrairValorColuna(row, ["Status"]);
        const tr = extrairValorColuna(row, ["Transportadora"]);
        const ms = extrairValorColuna(row, ["MÊS", "Mes"]);

        if (st) statusSet.add(st);
        if (tr) transpSet.add(tr);
        if (ms) mesSet.add(ms);
    });

    if (statusFilter) {
        statusFilter.innerHTML = '<option value="TODOS">Todos os Status</option>';
        Array.from(statusSet).sort().forEach(st => statusFilter.innerHTML += `<option value="${st}">${st}</option>`);
    }

    if (transportadoraFilter) {
        transportadoraFilter.innerHTML = '<option value="TODAS">Todas as Transportadoras</option>';
        Array.from(transpSet).sort().forEach(tr => transportadoraFilter.innerHTML += `<option value="${tr}">${tr}</option>`);
    }

    if (monthFilter) {
        monthFilter.innerHTML = '<option value="TODOS">Todos os Meses</option>';
        Array.from(mesSet).sort().forEach(ms => monthFilter.innerHTML += `<option value="${ms}">${ms}</option>`);
    }
}

/* ============================================================
 * FILTRAGEM E ATUALIZAÇÃO GERAL
 * ============================================================ */
function processarEAtualizar() {
    if (!dadosBrutos || dadosBrutos.length === 0) return;

    const statusSel = document.getElementById("statusFilter")?.value || "TODOS";
    const transpSel = document.getElementById("transportadoraFilter")?.value || "TODAS";
    const mesSel = document.getElementById("monthFilter")?.value || "TODOS";

    dadosFiltrados = dadosBrutos.filter((row) => {
        const st = extrairValorColuna(row, ["Status"]);
        if (statusSel !== "TODOS" && st !== statusSel) return false;

        const tr = extrairValorColuna(row, ["Transportadora"]);
        if (transpSel !== "TODAS" && tr !== transpSel) return false;

        const ms = extrairValorColuna(row, ["MÊS", "Mes"]);
        if (mesSel !== "TODOS" && ms !== mesSel) return false;

        return true;
    });

    atualizarKPIs(dadosFiltrados);
    atualizarGraficos(dadosFiltrados);
    atualizarTabelaAgenda(dadosFiltrados);
}

/* ============================================================
 * ATUALIZAÇÃO DOS KPIS (DEDUPLICANDO SOLICITAÇÕES)
 * ============================================================ */
function atualizarKPIs(dados) {
    const painelKPIs = document.getElementById("painelKPIs");
    if (!painelKPIs) return;

    const pedidosUnicosMap = new Map();
    let totalCotacoesRecebidas = dados.length;
    let valorTotalAprovado = 0;
    let somaValoresCotacoes = 0;

    dados.forEach((row) => {
        const idSolicitacao = extrairValorColuna(row, ["Solicitação", "Solicitacao", "PEDIDO"]);
        const status = extrairValorColuna(row, ["Status"]);
        const valorCotacao = parseFloat(extrairValorColuna(row, ["Valor Cotação", "Valor Cotacao"])) || 0;
        const verba = parseFloat(extrairValorColuna(row, ["Verba"])) || 0;

        somaValoresCotacoes += valorCotacao;

        if (status.toUpperCase() === "APROVADO") {
            valorTotalAprovado += valorCotacao;
        }

        // Agrupamento por ID de Solicitação único
        if (idSolicitacao) {
            if (!pedidosUnicosMap.has(idSolicitacao)) {
                pedidosUnicosMap.set(idSolicitacao, {
                    id: idSolicitacao,
                    status: 'EM ANALISE',
                    verba: verba,
                    menorCotacao: valorCotacao
                });
            }

            const ped = pedidosUnicosMap.get(idSolicitacao);
            if (valorCotacao < ped.menorCotacao && valorCotacao > 0) {
                ped.menorCotacao = valorCotacao;
            }

            if (status.toUpperCase() === "APROVADO") {
                ped.status = 'APROVADO';
            }
        }
    });

    const totalPedidosGerados = pedidosUnicosMap.size;
    let pedidosAprovadosCount = 0;
    let economiaPotencial = 0;

    pedidosUnicosMap.forEach((ped) => {
        if (ped.status === 'APROVADO') pedidosAprovadosCount++;
        if (ped.verba > ped.menorCotacao) {
            economiaPotencial += (ped.verba - ped.menorCotacao);
        }
    });

    const valorMedioCotacao = totalCotacoesRecebidas > 0 ? (somaValoresCotacoes / totalCotacoesRecebidas) : 0;
    const taxaAprovacao = totalCotacoesRecebidas > 0 ? ((pedidosAprovadosCount / totalCotacoesRecebidas) * 100).toFixed(1) : 0;

    painelKPIs.innerHTML = `
        <div class="kpi-card">
            <h4>Total de Solicitações</h4>
            <span class="kpi-value">${totalPedidosGerados}</span>
        </div>
        <div class="kpi-card">
            <h4>Cotações Recebidas</h4>
            <span class="kpi-value">${totalCotacoesRecebidas}</span>
        </div>
        <div class="kpi-card">
            <h4>Valor Médio / Cotação</h4>
            <span class="kpi-value">${formatarMoeda(valorMedioCotacao)}</span>
        </div>
        <div class="kpi-card">
            <h4>% Cotações Aprovadas</h4>
            <span class="kpi-value">${taxaAprovacao}%</span>
        </div>
        <div class="kpi-card">
            <h4>Valor Total Aprovado</h4>
            <span class="kpi-value">${formatarMoeda(valorTotalAprovado)}</span>
        </div>
        <div class="kpi-card">
            <h4>Economia Potencial</h4>
            <span class="kpi-value">${formatarMoeda(economiaPotencial)}</span>
        </div>
    `;
}

/* ============================================================
 * ATUALIZAÇÃO DOS GRÁFICOS (CHART.JS)
 * ============================================================ */
function atualizarGraficos(dados) {
    const ctxTransp = document.getElementById("chartRegiao")?.getContext("2d"); // Usando id existente
    const ctxStatus = document.getElementById("chartTipo")?.getContext("2d");

    const transpMedias = {};
    const statusCounts = {};

    dados.forEach((row) => {
        const tr = extrairValorColuna(row, ["Transportadora"]) || "Outras";
        const st = extrairValorColuna(row, ["Status"]) || "Indefinido";
        const val = parseFloat(extrairValorColuna(row, ["Valor Cotação", "Valor Cotacao"])) || 0;

        if (!transpMedias[tr]) transpMedias[tr] = { soma: 0, qtd: 0 };
        transpMedias[tr].soma += val;
        transpMedias[tr].qtd += 1;

        statusCounts[st] = (statusCounts[st] || 0) + 1;
    });

    const transpLabels = Object.keys(transpMedias);
    const transpValores = transpLabels.map(tr => (transpMedias[tr].soma / transpMedias[tr].qtd).toFixed(2));

    if (ctxTransp) {
        if (graficoTransportadoraInstance) graficoTransportadoraInstance.destroy();
        graficoTransportadoraInstance = new Chart(ctxTransp, {
            type: "bar",
            data: {
                labels: transpLabels,
                datasets: [{
                    label: "Valor Médio da Cotação (R$)",
                    data: transpValores,
                    backgroundColor: "#2563eb"
                }]
            },
            options: { responsive: true, maintainAspectRatio: false }
        });
    }

    if (ctxStatus) {
        if (graficoStatusInstance) graficoStatusInstance.destroy();
        graficoStatusInstance = new Chart(ctxStatus, {
            type: "doughnut",
            data: {
                labels: Object.keys(statusCounts),
                datasets: [{
                    data: Object.values(statusCounts),
                    backgroundColor: ["#eab308", "#dc2626", "#16a34a", "#2563eb"]
                }]
            },
            options: { responsive: true, maintainAspectRatio: false }
        });
    }
}

/* ============================================================
 * ATUALIZAÇÃO DA TABELA DE SOLICITAÇÕES
 * ============================================================ */
function atualizarTabelaAgenda(dados) {
    const tbody = document.querySelector("#tabelaAgenda tbody");
    if (!tbody) return;

    tbody.innerHTML = "";

    if (dados.length === 0) {
        tbody.innerHTML = `<tr><td colspan="6" style="text-align:center;">Nenhuma cotação encontrada.</td></tr>`;
        return;
    }

    dados.forEach((row) => {
        const tr = document.createElement("tr");
        const idPed = extrairValorColuna(row, ["Solicitação", "Solicitacao"]);
        const origDest = `${extrairValorColuna(row, ["Origem"])} ➔ ${extrairValorColuna(row, ["Destino"])}`;
        const transp = extrairValorColuna(row, ["Transportadora"]);
        const valor = formatarMoeda(extrairValorColuna(row, ["Valor Cotação", "Valor Cotacao"]));
        const verba = formatarMoeda(extrairValorColuna(row, ["Verba"]));
        const status = extrairValorColuna(row, ["Status"]);

        tr.innerHTML = `
            <td><strong>#${idPed}</strong></td>
            <td>${origDest}</td>
            <td>${transp}</td>
            <td>${verba}</td>
            <td>${valor}</td>
            <td><span class="badge status-${normalizarChave(status)}">${status}</span></td>
        `;

        tbody.appendChild(tr);
    });
}

/* ============================================================
 * LIMPEZA DE FILTROS
 * ============================================================ */
function limparFiltros() {
    const statusFilter = document.getElementById("statusFilter");
    const transportadoraFilter = document.getElementById("transportadoraFilter");
    const monthFilter = document.getElementById("monthFilter");

    if (statusFilter) statusFilter.value = "TODOS";
    if (transportadoraFilter) transportadoraFilter.value = "TODAS";
    if (monthFilter) monthFilter.value = "TODOS";

    processarEAtualizar();
}