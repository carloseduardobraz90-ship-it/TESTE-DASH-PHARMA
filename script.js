/* ============================================================
 * VARIÁVEIS GLOBAIS DE ESTADO
 * ============================================================ */
let dadosBrutos = [];
let dadosFiltrados = [];
let graficoStatusInstance = null;
let graficoCategoriaInstance = null;
let graficoEvolucaoValorInstance = null; // NOVO

/* ============================================================
 * EVENTOS INICIAIS DA PÁGINA
 * ============================================================ */
document.addEventListener("DOMContentLoaded", () => {
    const btnUpdate = document.getElementById("btnUpdate");
    const btnClear = document.getElementById("btnClear");
    const excelFileInput = document.getElementById("excelFile");
    const statusFilter = document.getElementById("statusFilter");
    const categoryFilter = document.getElementById("categoryFilter");
    const buyerFilter = document.getElementById("buyerFilter");
    const serviceFilter = document.getElementById("serviceFilter"); // NOVO

    if (excelFileInput) {
        excelFileInput.addEventListener("change", processarArquivo);
    }

    if (btnUpdate) {
        btnUpdate.addEventListener("click", processarEAtualizar);
    }

    if (btnClear) {
        btnClear.addEventListener("click", limparFiltros);
    }

    if (statusFilter) statusFilter.addEventListener("change", processarEAtualizar);
    if (categoryFilter) categoryFilter.addEventListener("change", processarEAtualizar);
    if (buyerFilter) buyerFilter.addEventListener("change", processarEAtualizar);
    if (serviceFilter) serviceFilter.addEventListener("change", processarEAtualizar); // NOVO
});

/* ============================================================
 * FUNÇÕES AUXILIARES E CONVERSÃO DE VALORES
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

function converterMoedaParaNumero(valor) {
    if (!valor) return 0;
    if (typeof valor === 'number') return valor;

    const valorStr = String(valor).trim();
    
    if (valorStr.includes(',') && valorStr.includes('.')) {
        const limpo = valorStr.replace(/\./g, '').replace(',', '.').replace(/[^0-9.-]/g, '');
        return parseFloat(limpo) || 0;
    }
    if (valorStr.includes(',') && !valorStr.includes('.')) {
        const limpo = valorStr.replace(',', '.').replace(/[^0-9.-]/g, '');
        return parseFloat(limpo) || 0;
    }

    const limpo = valorStr.replace(/[^0-9.-]/g, '');
    return parseFloat(limpo) || 0;
}

function formatarMoedaBRL(valor) {
    return Number(valor || 0).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
}

function parseDataBR(dataStr) {
    if (!dataStr) return null;
    
    // Se vier no formato Excel serial number ou string de data
    if (typeof dataStr === 'number') {
        const utc_days = Math.floor(dataStr - 25569);
        const utc_value = utc_days * 86400;
        return new Date(utc_value * 1000);
    }

    const limpo = String(dataStr).trim();
    
    // Formato DD/MM/AAAA ou DD-MM-AAAA
    const partes = limpo.split(/[\/\-]/);
    if (partes.length === 3) {
        // Se o primeiro tiver 4 dígitos (AAAA-MM-DD)
        if (partes[0].length === 4) {
            return new Date(parseInt(partes[0]), parseInt(partes[1]) - 1, parseInt(partes[2]));
        }
        // Padrão BR (DD/MM/AAAA)
        return new Date(parseInt(partes[2]), parseInt(partes[1]) - 1, parseInt(partes[0]));
    }

    const dataObj = new Date(limpo);
    return isNaN(dataObj.getTime()) ? null : dataObj;
}

// NOVO: Função de classificação de Serviços vs Produtos
function identificarTipoServico(row) {
    const descricao = extrairValorColuna(row, ["DESCRICAO"]);

    const texto = String(descricao || "")
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .toUpperCase();

    if (texto.includes("SERVICO")) {
        return "SERVICOS";
    }

    return "PRODUTOS";
}

/* ============================================================
 * LEITURA DO ARQUIVO CSV / EXCEL
 * ============================================================ */
function processarArquivo(event) {
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
            const buffer = e.target.result;

            if (file.name.endsWith('.csv')) {
                const text = new TextDecoder('utf-8').decode(buffer);
                dadosBrutos = converterCSVParaArray(text);
            } else {
                const data = new Uint8Array(buffer);
                const workbook = XLSX.read(data, { type: "array" });
                const firstSheet = workbook.SheetNames[0];
                dadosBrutos = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet], { defval: "" });
            }

            if (!dadosBrutos || dadosBrutos.length === 0) {
                throw new Error("A planilha/CSV selecionado está vazio.");
            }

            if (statusIcon) statusIcon.innerText = "🟢";
            if (statusMensagem) statusMensagem.innerText = "Base de dados carregada!";
            if (statusDetalhes) statusDetalhes.innerText = `${dadosBrutos.length} linhas de itens importadas.`;

            popularFiltrosSelect(dadosBrutos);
            processarEAtualizar();

        } catch (error) {
            console.error("Erro ao processar arquivo:", error);
            if (statusIcon) statusIcon.innerText = "🔴";
            if (statusMensagem) statusMensagem.innerText = "Erro ao carregar o arquivo";
            if (statusDetalhes) statusDetalhes.innerText = "Verifique o formato e tente novamente.";

            if (painelErro) {
                painelErro.style.display = "block";
                document.getElementById("erroDetalhes").innerText = error.message || "Erro desconhecido ao ler o arquivo.";
            }
        }
    };

    reader.readAsArrayBuffer(file);
}

function converterCSVParaArray(strData) {
    const lines = strData.split(/\r\n|\n/);
    if (lines.length === 0) return [];

    const primeiraLinha = lines[0];
    const qtdPontoVirgula = (primeiraLinha.match(/;/g) || []).length;
    const qtdVirgula = (primeiraLinha.match(/,/g) || []).length;
    const strDelimiter = qtdPontoVirgula >= qtdVirgula ? ";" : ",";

    const headers = primeiraLinha.split(strDelimiter).map(h => h.replace(/^["\uFEFF]+|["\s]+$/g, ''));
    const result = [];

    for (let i = 1; i < lines.length; i++) {
        if (!lines[i].trim()) continue;

        const row = {};
        const currentline = lines[i].split(strDelimiter);

        headers.forEach((header, index) => {
            let val = currentline[index] || "";
            val = val.replace(/^["\s]+|["\s]+$/g, '');
            row[header] = val;
        });

        result.push(row);
    }

    return result;
}

/* ============================================================
 * POPULAR FILTROS SELECT DINAMICAMENTE
 * ============================================================ */
function popularFiltrosSelect(dados) {
    const statusFilter = document.getElementById("statusFilter");
    const categoryFilter = document.getElementById("categoryFilter");
    const buyerFilter = document.getElementById("buyerFilter");

    const statusSet = new Set();
    const catSet = new Set();
    const buyerSet = new Set();

    dados.forEach((row) => {
        const st = extrairValorColuna(row, ["STATUS", "SITUACAO", "STATUS DO PEDIDO"]);
        const cat = extrairValorColuna(row, ["CATEGORIA", "TIPO", "GRUPO"]);
        const buyer = extrairValorColuna(row, ["COMPRADOR", "RESPONSAVEL"]);

        if (st) statusSet.add(st);
        if (cat) catSet.add(cat);
        if (buyer) buyerSet.add(buyer);
    });

    if (statusFilter) {
        statusFilter.innerHTML = '<option value="TODOS">Todos os Status</option>';
        Array.from(statusSet).sort().forEach(st => statusFilter.innerHTML += `<option value="${st}">${st}</option>`);
    }

    if (categoryFilter) {
        categoryFilter.innerHTML = '<option value="TODAS">Todas as Categorias</option>';
        Array.from(catSet).sort().forEach(cat => categoryFilter.innerHTML += `<option value="${cat}">${cat}</option>`);
    }

    if (buyerFilter) {
        buyerFilter.innerHTML = '<option value="TODOS">Todos os Compradores</option>';
        Array.from(buyerSet).sort().forEach(b => buyerFilter.innerHTML += `<option value="${b}">${b}</option>`);
    }
}

/* ============================================================
 * FILTRAGEM E ATUALIZAÇÃO GERAL
 * ============================================================ */
function processarEAtualizar() {
    if (!dadosBrutos || dadosBrutos.length === 0) return;

    const statusSel = document.getElementById("statusFilter")?.value || "TODOS";
    const catSel = document.getElementById("categoryFilter")?.value || "TODAS";
    const buyerSel = document.getElementById("buyerFilter")?.value || "TODOS";
    const serviceSel = document.getElementById("serviceFilter")?.value || "TODOS"; // NOVO

    dadosFiltrados = dadosBrutos.filter((row) => {
        const st = extrairValorColuna(row, ["STATUS", "SITUACAO", "STATUS DO PEDIDO"]);
        if (statusSel !== "TODOS" && st !== statusSel) return false;

        const cat = extrairValorColuna(row, ["CATEGORIA", "TIPO", "GRUPO"]);
        if (catSel !== "TODAS" && cat !== catSel) return false;

        const buyer = extrairValorColuna(row, ["COMPRADOR", "RESPONSAVEL"]);
        if (buyerSel !== "TODOS" && buyer !== buyerSel) return false;

        // NOVO: Filtro de Tipo (Produtos vs Serviços)
        const tipoServico = identificarTipoServico(row);
        if (serviceSel !== "TODOS" && tipoServico !== serviceSel) {
            return false;
        }

        return true;
    });

    atualizarKPIs(dadosFiltrados);
    atualizarGraficos(dadosFiltrados);
    atualizarGraficoEvolucaoValor(dadosFiltrados); // NOVO
    atualizarTabela(dadosFiltrados);
}

/* ============================================================
 * ATUALIZAÇÃO DOS KPIS
 * ============================================================ */
function atualizarKPIs(dados) {
    const painelKPIs = document.getElementById("painelKPIs");
    if (!painelKPIs) return;

    const pedidosUnicosSet = new Set();
    let totalItens = dados.length;
    let valorTotalGeral = 0;
    let pedidosFinalizadosSet = new Set();

    dados.forEach((row) => {
        const numPedido = extrairValorColuna(row, ["PEDIDO_", "PEDIDO", "SOLICITACAO", "NUMERO DO PEDIDO"]);
        const status = extrairValorColuna(row, ["STATUS", "SITUACAO"]);
        const valorItem = converterMoedaParaNumero(extrairValorColuna(row, ["VLR.TOTAL", "VLRTOTAL", "VALOR TOTAL", "VALOR"]));

        valorTotalGeral += valorItem;

        if (numPedido) {
            pedidosUnicosSet.add(numPedido);
            if (status.toUpperCase().includes("FINALIZADO") || status.toUpperCase().includes("ENTREGUE")) {
                pedidosFinalizadosSet.add(numPedido);
            }
        }
    });

    const totalPedidosUnicos = pedidosUnicosSet.size;
    const totalFinalizados = pedidosFinalizadosSet.size;
    const ticketMedioPorPedido = totalPedidosUnicos > 0 ? (valorTotalGeral / totalPedidosUnicos) : 0;

    painelKPIs.innerHTML = `
        <div class="kpi-card">
            <h4>Pedidos Gerados (Únicos)</h4>
            <span class="kpi-value">${totalPedidosUnicos.toLocaleString('pt-BR')}</span>
        </div>
        <div class="kpi-card">
            <h4>Total de Itens / Linhas</h4>
            <span class="kpi-value">${totalItens.toLocaleString('pt-BR')}</span>
        </div>
        <div class="kpi-card">
            <h4>Pedidos Concluídos</h4>
            <span class="kpi-value">${totalFinalizados.toLocaleString('pt-BR')}</span>
        </div>
        <div class="kpi-card">
            <h4>Valor Total Acumulado</h4>
            <span class="kpi-value">${formatarMoedaBRL(valorTotalGeral)}</span>
        </div>
        <div class="kpi-card">
            <h4>Ticket Médio por Pedido</h4>
            <span class="kpi-value">${formatarMoedaBRL(ticketMedioPorPedido)}</span>
        </div>
    `;
}

/* ============================================================
 * ATUALIZAÇÃO DOS GRÁFICOS (CHART.JS)
 * ============================================================ */
function atualizarGraficos(dados) {
    const ctxStatus = document.getElementById("chartRegiao")?.getContext("2d");
    const ctxCategoria = document.getElementById("chartTipo")?.getContext("2d");

    const statusCounts = {};
    const catCounts = {};

    dados.forEach((row) => {
        const st = extrairValorColuna(row, ["STATUS", "SITUACAO"]) || "Indefinido";
        const cat = extrairValorColuna(row, ["CATEGORIA", "TIPO"]) || "Outros";

        statusCounts[st] = (statusCounts[st] || 0) + 1;
        catCounts[cat] = (catCounts[cat] || 0) + 1;
    });

    if (ctxStatus) {
        if (graficoStatusInstance) graficoStatusInstance.destroy();
        graficoStatusInstance = new Chart(ctxStatus, {
            type: "bar",
            data: {
                labels: Object.keys(statusCounts),
                datasets: [{
                    label: "Quantidade de Itens por Status",
                    data: Object.values(statusCounts),
                    backgroundColor: "#2563eb"
                }]
            },
            options: { responsive: true, maintainAspectRatio: false }
        });
    }

    if (ctxCategoria) {
        if (graficoCategoriaInstance) graficoCategoriaInstance.destroy();
        graficoCategoriaInstance = new Chart(ctxCategoria, {
            type: "doughnut",
            data: {
                labels: Object.keys(catCounts),
                datasets: [{
                    data: Object.values(catCounts),
                    backgroundColor: ["#16a34a", "#2563eb", "#eab308", "#dc2626", "#9333ea"]
                }]
            },
            options: { responsive: true, maintainAspectRatio: false }
        });
    }
}

/* ============================================================
 * NOVO GRÁFICO: EVOLUÇÃO DO VALOR FINALIZADO POR SEMANA
 * ============================================================ */
function atualizarGraficoEvolucaoValor(dados) {
    const canvas = document.getElementById("chartEvolucaoValor");
    if (!canvas) return;

    const ctx = canvas.getContext("2d");

    if (graficoEvolucaoValorInstance) {
        graficoEvolucaoValorInstance.destroy();
    }

    const valoresPorSemana = {};

    dados.forEach((row) => {
        const status = extrairValorColuna(row, ["STATUS", "SITUACAO"])
            .toUpperCase()
            .trim();

        // Considera finalizado ou entregue
        if (!status.includes("FINALIZADO") && !status.includes("ENTREGUE")) {
            return;
        }

        const dataAgendamento = extrairValorColuna(row, ["AGENDAMENTO", "DATA", "EMISSAO"]);
        const data = parseDataBR(dataAgendamento);

        if (!data) return;

        const valor = converterMoedaParaNumero(
            extrairValorColuna(row, ["VLR.TOTAL", "VLRTOTAL", "VALOR TOTAL", "VALOR"])
        );

        if (!valor) return;

        const ano = data.getFullYear();
        const mes = data.getMonth();
        const semana = Math.floor((data.getDate() - 1) / 7) + 1;

        const chave = `${ano}-${String(mes + 1).padStart(2, "0")}-S${semana}`;

        if (!valoresPorSemana[chave]) {
            valoresPorSemana[chave] = { ano, mes, semana, valor: 0 };
        }

        valoresPorSemana[chave].valor += valor;
    });

    const dadosGrafico = Object.values(valoresPorSemana).sort((a, b) => {
        if (a.ano !== b.ano) return a.ano - b.ano;
        if (a.mes !== b.mes) return a.mes - b.mes;
        return a.semana - b.semana;
    });

    const nomesMeses = [
        "Jan", "Fev", "Mar", "Abr", "Mai", "Jun",
        "Jul", "Ago", "Set", "Out", "Nov", "Dez"
    ];

    const labels = dadosGrafico.map(item => {
        const primeiroDia = ((item.semana - 1) * 7) + 1;
        const ultimoDia = Math.min(primeiroDia + 6, 31);
        return `${nomesMeses[item.mes]} S${item.semana} (${primeiroDia}-${ultimoDia})`;
    });

    const valores = dadosGrafico.map(item => item.valor);

    graficoEvolucaoValorInstance = new Chart(ctx, {
        type: "line",
        data: {
            labels: labels,
            datasets: [{
                label: "Valor Finalizado",
                data: valores,
                borderColor: "#2563eb",
                backgroundColor: "rgba(37, 99, 235, 0.10)",
                borderWidth: 3,
                tension: 0.3,
                fill: true,
                pointRadius: 5,
                pointHoverRadius: 7
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            interaction: {
                intersect: false,
                mode: "index"
            },
            plugins: {
                legend: { display: true },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            return " " + formatarMoedaBRL(context.raw);
                        }
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: {
                        callback: function(value) {
                            return formatarMoedaBRL(value);
                        }
                    }
                }
            }
        }
    });
}

/* ============================================================
 * ATUALIZAÇÃO DA TABELA DE DADOS
 * ============================================================ */
function atualizarTabela(dados) {
    const tbody = document.querySelector("#tabelaAgenda tbody");
    if (!tbody) return;

    tbody.innerHTML = "";

    if (dados.length === 0) {
        tbody.innerHTML = `<tr><td colspan="5" style="text-align:center;">Nenhum registro encontrado.</td></tr>`;
        return;
    }

    const limiteExibicao = dados.slice(0, 100);

    limiteExibicao.forEach((row) => {
        const tr = document.createElement("tr");
        const numPedido = extrairValorColuna(row, ["PEDIDO_", "PEDIDO", "SOLICITACAO"]);
        const item = extrairValorColuna(row, ["ITEM_", "ITEM"]);
        const fornecedor = extrairValorColuna(row, ["NOME", "FORNECEDOR"]);
        const categoria = extrairValorColuna(row, ["CATEGORIA", "TIPO"]);
        const valor = formatarMoedaBRL(converterMoedaParaNumero(extrairValorColuna(row, ["VLR.TOTAL", "VLRTOTAL", "VALOR TOTAL"])));
        const status = extrairValorColuna(row, ["STATUS", "SITUACAO"]);

        tr.innerHTML = `
            <td><strong>#${numPedido}</strong> (${item})</td>
            <td>${fornecedor}</td>
            <td>${categoria}</td>
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
    const categoryFilter = document.getElementById("categoryFilter");
    const buyerFilter = document.getElementById("buyerFilter");
    const serviceFilter = document.getElementById("serviceFilter"); // NOVO

    if (statusFilter) statusFilter.value = "TODOS";
    if (categoryFilter) categoryFilter.value = "TODAS";
    if (buyerFilter) buyerFilter.value = "TODOS";
    if (serviceFilter) serviceFilter.value = "TODOS"; // NOVO

    processarEAtualizar();
}