/* ============================================================
 * VARIÁVEIS GLOBAIS DE ESTADO
 * ============================================================ */
let dadosBrutos = [];
let dadosFiltrados = [];
let graficoStatusInstance = null;
let graficoCategoriaInstance = null;

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

    // Trata formato BRL ("23.969,74" -> 23969.74) ou formato americano ("23969.74")
    const valorStr = String(valor).trim();
    
    // Se tiver vírgula e ponto, assume padrão BR
    if (valorStr.includes(',') && valorStr.includes('.')) {
        const limpo = valorStr.replace(/\./g, '').replace(',', '.').replace(/[^0-9.-]/g, '');
        return parseFloat(limpo) || 0;
    }
    // Se tiver apenas vírgula
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

            // LOG DE DIAGNÓSTICO NO CONSOLE (F12)
            console.log("=== ARQUIVO CARREGADO COM SUCESSO ===");
            console.log("Total de linhas:", dadosBrutos.length);
            console.log("Colunas encontradas na primeira linha:", Object.keys(dadosBrutos[0]));
            console.log("Exemplo da primeira linha:", dadosBrutos[0]);

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

// Parser de CSV inteligente com detecção de delimitador (';' ou ',')
function converterCSVParaArray(strData) {
    const lines = strData.split(/\r\n|\n/);
    if (lines.length === 0) return [];

    // Detecta se usa ';' ou ',' na primeira linha
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

    dadosFiltrados = dadosBrutos.filter((row) => {
        const st = extrairValorColuna(row, ["STATUS", "SITUACAO", "STATUS DO PEDIDO"]);
        if (statusSel !== "TODOS" && st !== statusSel) return false;

        const cat = extrairValorColuna(row, ["CATEGORIA", "TIPO", "GRUPO"]);
        if (catSel !== "TODAS" && cat !== catSel) return false;

        const buyer = extrairValorColuna(row, ["COMPRADOR", "RESPONSAVEL"]);
        if (buyerSel !== "TODOS" && buyer !== buyerSel) return false;

        return true;
    });

    atualizarKPIs(dadosFiltrados);
    atualizarGraficos(dadosFiltrados);
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

    // Exibe os primeiros 100 registros para otimizar renderização DOM
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

    if (statusFilter) statusFilter.value = "TODOS";
    if (categoryFilter) categoryFilter.value = "TODAS";
    if (buyerFilter) buyerFilter.value = "TODOS";

    processarEAtualizar();
}