/* ============================================================
 * DASHBOARD OPERACIONAL & LOGÍSTICO - SCRIPT CORRIGIDO
 * ============================================================ */

// Estado Global da Aplicação
let dadosBrutos = [];
let dadosFiltrados = [];
let graficoRegiaoInstance = null;
let graficoTipoInstance = null;

/* ============================================================
 * INICIALIZAÇÃO E EVENTOS
 * ============================================================ */

document.addEventListener("DOMContentLoaded", function () {
    const excelFile = document.getElementById("excelFile");
    const btnUpdate = document.getElementById("btnUpdate");
    const statusFilter = document.getElementById("statusFilter");
    const regionFilter = document.getElementById("regionFilter");
    const startDate = document.getElementById("startDate");
    const endDate = document.getElementById("endDate");
    const btnClear = document.getElementById("btnClear");

    if (excelFile) excelFile.addEventListener("change", lerArquivo);
    if (btnUpdate) btnUpdate.addEventListener("click", processarEAtualizar);
    if (statusFilter) statusFilter.addEventListener("change", processarEAtualizar);
    if (regionFilter) regionFilter.addEventListener("change", processarEAtualizar);
    if (startDate) startDate.addEventListener("change", processarEAtualizar);
    if (endDate) endDate.addEventListener("change", processarEAtualizar);
    if (btnClear) btnClear.addEventListener("click", limparFiltros);
});

/* ============================================================
 * LEITURA DE ARQUIVOS (XLSX / CSV)
 * ============================================================ */

function lerArquivo(evento) {
    const arquivo = evento.target.files[0];
    if (!arquivo) return;

    const leitor = new FileReader();

    leitor.onload = function (e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: "array", cellDates: true });
            const primeiraAba = workbook.SheetNames[0];
            const planilha = workbook.Sheets[primeiraAba];

            // Converte para JSON
            dadosBrutos = XLSX.utils.sheet_to_json(planilha, { defval: "" });

            // Atualiza status na interface
            const statusIcon = document.getElementById("statusIcon");
            const statusMensagem = document.getElementById("statusMensagem");
            const statusDetalhes = document.getElementById("statusDetalhes");

            if (statusIcon) statusIcon.innerText = "🟢";
            if (statusMensagem) statusMensagem.innerText = "Base carregada com sucesso!";
            if (statusDetalhes) statusDetalhes.innerText = `${dadosBrutos.length} registros importados.`;

            // Oculta painel de erro se estiver visível
            const painelErro = document.getElementById("painelErro");
            if (painelErro) painelErro.style.display = "none";

            popularFiltrosSelect(dadosBrutos);
            processarEAtualizar();
        } catch (erro) {
            console.error("Erro ao ler arquivo:", erro);
            const painelErro = document.getElementById("painelErro");
            const erroDetalhes = document.getElementById("erroDetalhes");

            if (painelErro && erroDetalhes) {
                painelErro.style.display = "block";
                erroDetalhes.innerText = erro.message || "Falha ao processar a planilha.";
            } else {
                alert("Falha ao processar o arquivo. Verifique se é um arquivo Excel/CSV válido.");
            }
        }
    };

    leitor.readAsArrayBuffer(arquivo);
}

/* ============================================================
 * POPULAR OPÇÕES DOS FILTROS (DROPDOWNS)
 * ============================================================ */

function popularFiltrosSelect(dados) {
    const statusFilter = document.getElementById("statusFilter");
    const regionFilter = document.getElementById("regionFilter");

    if (!statusFilter || !regionFilter) return;

    const statusSet = new Set();
    const regionSet = new Set();

    dados.forEach((row) => {
        const st = normalizarTexto(row.STATUS || row.Status || row.status);
        const rg = normalizarTexto(row.REGIAO || row.Regiao || row.regiao || row.UF || row.Estado);

        if (st) statusSet.add(st);
        if (rg) regionSet.add(rg);
    });

    // Atualiza Select de Status
    statusFilter.innerHTML = '<option value="TODOS">Todos os Status</option>';
    Array.from(statusSet).sort().forEach((st) => {
        statusFilter.innerHTML += `<option value="${st}">${st}</option>`;
    });

    // Atualiza Select de Região
    regionFilter.innerHTML = '<option value="TODAS">Todas as Regiões</option>';
    Array.from(regionSet).sort().forEach((rg) => {
        regionFilter.innerHTML += `<option value="${rg}">${nomeRegiao(rg)}</option>`;
    });
}

/* ============================================================
 * PROCESSAMENTO E FILTRAGEM
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
        // Status
        const st = normalizarTexto(row.STATUS || row.Status || row.status);
        if (statusSel !== "TODOS" && st !== statusSel) return false;

        // Região
        const rg = normalizarTexto(row.REGIAO || row.Regiao || row.regiao || row.UF || row.Estado);
        if (regionSel !== "TODAS" && rg !== regionSel) return false;

        // Data
        const rawData = row.DATA || row.Data || row.data || row.DATA_AGENDAMENTO || row.DataAgendamento;
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
 * ATUALIZAÇÃO DE KPIS (CARDS DE RESUMO)
 * ============================================================ */

function atualizarKPIs(dados) {
    const painelKPIs = document.getElementById("painelKPIs");
    if (!painelKPIs) return;

    let total = dados.length;
    let pendentes = 0;
    let concluidos = 0;

    dados.forEach((row) => {
        const st = normalizarTexto(row.STATUS || row.Status || row.status).toUpperCase();
        if (st.includes("CONCLU") || st.includes("ENTREGUE") || st.includes("FINALIZADO")) {
            concluidos++;
        } else if (st.includes("PENDENTE") || st.includes("EM ANDAMENTO") || st.includes("AGUARDANDO")) {
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
 * GRÁFICOS (CHART.JS)
 * ============================================================ */

function atualizarGraficos(dados) {
    const ctxRegiao = document.getElementById("chartRegiao")?.getContext("2d");
    const ctxTipo = document.getElementById("chartTipo")?.getContext("2d");

    const regiaoCounts = {};
    const tipoCounts = {};

    dados.forEach((row) => {
        const rg = nomeRegiao(normalizarTexto(row.REGIAO || row.Regiao || row.regiao || row.UF || row.Estado) || "N/A");
        const tp = normalizarTexto(row.TIPO || row.Tipo || row.tipo || row.OPERACAO || row.Operacao) || "Outros";

        regiaoCounts[rg] = (regiaoCounts[rg] || 0) + 1;
        tipoCounts[tp] = (tipoCounts[tp] || 0) + 1;
    });

    // Gráfico por Região (Barras)
    if (ctxRegiao) {
        if (graficoRegiaoInstance) graficoRegiaoInstance.destroy();
        graficoRegiaoInstance = new Chart(ctxRegiao, {
            type: "bar",
            data: {
                labels: Object.keys(regiaoCounts),
                datasets: [{
                    label: "Volume por Região",
                    data: Object.values(regiaoCounts),
                    backgroundColor: "#3b82f6"
                }]
            },
            options: { responsive: true, maintainAspectRatio: false }
        });
    }

    // Gráfico por Tipo (Rosca/Doughnut)
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
 * TABELA PRINCIPAL / AGENDA
 * ============================================================ */

function atualizarTabelaAgenda(dados) {
    const tbody = document.querySelector("#tabelaAgenda tbody");
    if (!tbody) return;

    const grupos = {};

    dados.forEach((row) => {
        const rawData = row.DATA || row.Data || row.data || row.DATA_AGENDAMENTO || row.DataAgendamento;
        const ts = obterTimestampZerado(rawData) || 9999999999999;
        const regiao = normalizarTexto(row.REGIAO || row.Regiao || row.regiao || row.UF || row.Estado) || "N/A";
        const status = normalizarTexto(row.STATUS || row.Status || row.status) || "Indefinido";

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
            <td><span class="badge status-${grupo.status.toLowerCase().replace(/\s+/g, '-')}">${grupo.status}</span></td>
            <td><strong>${grupo.quantidade}</strong></td>
        `;

        tbody.appendChild(tr);
    });
}

/* ============================================================
 * MANIPULAÇÃO DE DATAS & UTILITÁRIOS
 * ============================================================ */

function obterTimestampZerado(valorData) {
    if (!valorData) return null;

    let data = null;

    if (valorData instanceof Date) {
        data = new Date(valorData);
    } else if (typeof valorData === "number") {
        data = new Date(Math.round((valorData - 25569) * 86400 * 1000));
    } else if (typeof valorData === "string") {
        const texto = valorData.trim();
        if (!texto) return null;

        if (texto.includes("/")) {
            const partes = texto.split("/");
            if (partes.length === 3) {
                data = new Date(partes[2], partes[1] - 1, partes[0]);
            }
        } else if (texto.includes("-")) {
            const partes = texto.split("-");
            if (partes.length === 3) {
                data = new Date(partes[0], partes[1] - 1, partes[2]);
            }
        }
    }

    if (!data || isNaN(data.getTime())) {
        return null;
    }

    data.setHours(0, 0, 0, 0);
    return data.getTime();
}

function formatarData(timestamp) {
    if (!timestamp || timestamp === 9999999999999) return "Sem Data";
    const data = new Date(timestamp);
    const dia = String(data.getDate()).padStart(2, "0");
    const mes = String(data.getMonth() + 1).padStart(2, "0");
    const ano = data.getFullYear();
    return `${dia}/${mes}/${ano}`;
}

function normalizarTexto(texto) {
    if (!texto) return "";
    return String(texto).trim();
}

function nomeRegiao(codigoOuNome) {
    if (!codigoOuNome) return "N/A";
    const mapa = {
        "MG": "Minas Gerais",
        "SP": "São Paulo",
        "RJ": "Rio de Janeiro",
        "PR": "Paraná",
        "SC": "Santa Catarina",
        "RS": "Rio Grande do Sul"
    };
    return mapa[codigoOuNome.toUpperCase()] || codigoOuNome;
}

/* ============================================================
 * LIMPAR FILTROS
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