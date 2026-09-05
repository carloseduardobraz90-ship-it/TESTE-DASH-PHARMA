/* ============================================================
 * DASHBOARD OPERACIONAL
 * ============================================================
 *
 * FUNCIONALIDADES:
 *
 * - Leitura XLSX / XLS / CSV
 * - Filtro de Status
 * - Filtro de Categoria
 * - Filtro de Projeto
 * - Filtro de Tipo: Produto / Serviço
 * - Filtro de período de Agendamento
 * - KPIs
 * - Gráfico de Status
 * - Gráfico de Categoria
 * - Gráfico de evolução do valor finalizado por semana
 * - Tabela de pedidos
 *
 * ============================================================ */


/* ============================================================
 * VARIÁVEIS GLOBAIS
 * ============================================================ */

let dadosBrutos = [];

let dadosFiltrados = [];

let graficoStatusInstance = null;

let graficoCategoriaInstance = null;

let graficoEvolucaoValorInstance = null;


/* ============================================================
 * INICIALIZAÇÃO
 * ============================================================ */

document.addEventListener("DOMContentLoaded", () => {

    const excelFile = document.getElementById("excelFile");

    const btnUpdate = document.getElementById("btnUpdate");

    const btnClear = document.getElementById("btnClear");

    const statusFilter = document.getElementById("statusFilter");

    const categoryFilter = document.getElementById("categoryFilter");

    const serviceFilter = document.getElementById("serviceFilter");

    const projectFilter = document.getElementById("projectFilter");

    const dtInicio = document.getElementById("dtInicio");

    const dtFim = document.getElementById("dtFim");


    if (excelFile) {

        excelFile.addEventListener(
            "change",
            processarArquivo
        );

    }


    if (btnUpdate) {

        btnUpdate.addEventListener(
            "click",
            processarEAtualizar
        );

    }


    if (btnClear) {

        btnClear.addEventListener(
            "click",
            limparFiltros
        );

    }


    if (statusFilter) {

        statusFilter.addEventListener(
            "change",
            processarEAtualizar
        );

    }


    if (categoryFilter) {

        categoryFilter.addEventListener(
            "change",
            processarEAtualizar
        );

    }


    if (serviceFilter) {

        serviceFilter.addEventListener(
            "change",
            processarEAtualizar
        );

    }


    if (projectFilter) {

        projectFilter.addEventListener(
            "change",
            processarEAtualizar
        );

    }


    if (dtInicio) {

        dtInicio.addEventListener(
            "change",
            processarEAtualizar
        );

    }


    if (dtFim) {

        dtFim.addEventListener(
            "change",
            processarEAtualizar
        );

    }

});


/* ============================================================
 * NORMALIZAÇÃO DE CHAVES
 * ============================================================ */

function normalizarChave(str) {

    return String(str || "")
        .replace(/^\uFEFF/, "")
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .toLowerCase()
        .replace(/[^a-z0-9]/g, "");

}


/* ============================================================
 * EXTRAIR VALOR DE COLUNA
 * ============================================================ */

function extrairValorColuna(row, nomesPossiveis) {

    if (!row) return "";

    const chaves = Object.keys(row);


    for (const nome of nomesPossiveis) {

        const nomeAlvo =
            normalizarChave(nome);


        const chaveEncontrada =
            chaves.find(
                k => normalizarChave(k) === nomeAlvo
            );


        if (
            chaveEncontrada &&
            row[chaveEncontrada] !== undefined &&
            row[chaveEncontrada] !== null
        ) {

            return String(
                row[chaveEncontrada]
            )
                .replace(/^["\s]+|["\s]+$/g, "")
                .trim();

        }

    }


    return "";

}


/* ============================================================
 * CONVERSÃO DE MOEDA
 * ============================================================ */

function converterMoedaParaNumero(valor) {

    if (
        valor === null ||
        valor === undefined ||
        valor === ""
    ) {

        return 0;

    }


    if (typeof valor === "number") {

        return valor;

    }


    let valorStr =
        String(valor)
            .trim()
            .replace(/[R$\s]/g, "");


    /*
     * Formato brasileiro:
     *
     * 23.969,74
     *
     */

    if (
        valorStr.includes(".") &&
        valorStr.includes(",")
    ) {

        valorStr =
            valorStr
                .replace(/\./g, "")
                .replace(",", ".");

    }


    /*
     * Apenas vírgula:
     *
     * 23969,74
     *
     */

    else if (
        valorStr.includes(",")
    ) {

        valorStr =
            valorStr.replace(",", ".");

    }


    valorStr =
        valorStr.replace(
            /[^0-9.-]/g,
            ""
        );


    const numero =
        parseFloat(valorStr);


    return isNaN(numero)
        ? 0
        : numero;

}


/* ============================================================
 * FORMATAÇÃO BRL
 * ============================================================ */

function formatarMoedaBRL(valor) {

    return Number(valor || 0)
        .toLocaleString(
            "pt-BR",
            {
                style: "currency",
                currency: "BRL"
            }
        );

}


/* ============================================================
 * DATA BR
 * ============================================================ */

function parseDataBR(valor) {

    if (
        valor === null ||
        valor === undefined ||
        valor === ""
    ) {

        return null;

    }


    /*
     * Caso o XLSX entregue Date
     */

    if (valor instanceof Date) {

        if (!isNaN(valor.getTime())) {

            return new Date(
                valor.getFullYear(),
                valor.getMonth(),
                valor.getDate()
            );

        }

    }


    const str =
        String(valor)
            .trim()
            .replace(/^["']+|["']+$/g, "");


    /*
     * Formato DD/MM/YYYY
     */

    const matchBR =
        str.match(
            /^(\d{1,2})\/(\d{1,2})\/(\d{4})/
        );


    if (matchBR) {

        const dia =
            parseInt(matchBR[1], 10);

        const mes =
            parseInt(matchBR[2], 10) - 1;

        const ano =
            parseInt(matchBR[3], 10);


        const data =
            new Date(
                ano,
                mes,
                dia
            );


        if (
            data.getFullYear() === ano &&
            data.getMonth() === mes &&
            data.getDate() === dia
        ) {

            return data;

        }


        return null;

    }


    /*
     * Formato YYYY-MM-DD
     */

    const matchISO =
        str.match(
            /^(\d{4})-(\d{1,2})-(\d{1,2})/
        );


    if (matchISO) {

        const ano =
            parseInt(matchISO[1], 10);

        const mes =
            parseInt(matchISO[2], 10) - 1;

        const dia =
            parseInt(matchISO[3], 10);


        const data =
            new Date(
                ano,
                mes,
                dia
            );


        if (
            data.getFullYear() === ano &&
            data.getMonth() === mes &&
            data.getDate() === dia
        ) {

            return data;

        }

    }


    /*
     * Número serial do Excel
     */

    if (
        /^[0-9]+(\.[0-9]+)?$/.test(str)
    ) {

        const serial =
            Number(str);


        if (
            serial > 20000 &&
            serial < 100000
        ) {

            const data =
                new Date(
                    Date.UTC(
                        1899,
                        11,
                        30
                    ) +
                    serial *
                    24 *
                    60 *
                    60 *
                    1000
                );


            return new Date(
                data.getUTCFullYear(),
                data.getUTCMonth(),
                data.getUTCDate()
            );

        }

    }


    return null;

}


/* ============================================================
 * IDENTIFICAR PRODUTO / SERVIÇO
 *
 * REGRA:
 *
 * Se DESCRICAO contém SERVICO/SERVIÇO
 * => SERVICOS
 *
 * Caso contrário
 * => PRODUTOS
 *
 * ============================================================ */

function identificarTipoItem(row) {

    const descricao =
        extrairValorColuna(
            row,
            ["DESCRICAO"]
        );


    const texto =
        String(descricao || "")
            .normalize("NFD")
            .replace(/[\u0300-\u036f]/g, "")
            .toUpperCase();


    if (
        texto.includes("SERVICO")
    ) {

        return "SERVICOS";

    }


    return "PRODUTOS";

}


/* ============================================================
 * PROCESSAR ARQUIVO
 * ============================================================ */

function processarArquivo(event) {

    const file =
        event.target.files[0];


    if (!file) return;


    const statusBanner =
        document.getElementById(
            "statusBanner"
        );


    const statusIcon =
        document.getElementById(
            "statusIcon"
        );


    const statusMensagem =
        document.getElementById(
            "statusMensagem"
        );


    const statusDetalhes =
        document.getElementById(
            "statusDetalhes"
        );


    const painelErro =
        document.getElementById(
            "painelErro"
        );


    const erroDetalhes =
        document.getElementById(
            "erroDetalhes"
        );


    if (painelErro) {

        painelErro.style.display =
            "none";

    }


    const reader =
        new FileReader();


    reader.onload =
        function (e) {

            try {

                const buffer =
                    e.target.result;


                const nomeArquivo =
                    file.name.toLowerCase();


                /*
                 * CSV
                 */

                if (
                    nomeArquivo.endsWith(".csv")
                ) {

                    const text =
                        new TextDecoder(
                            "utf-8"
                        ).decode(buffer);


                    dadosBrutos =
                        converterCSVParaArray(
                            text
                        );

                }


                /*
                 * XLSX / XLS
                 */

                else {

                    const data =
                        new Uint8Array(
                            buffer
                        );


                    const workbook =
                        XLSX.read(
                            data,
                            {
                                type: "array"
                            }
                        );


                    const firstSheet =
                        workbook
                            .SheetNames[0];


                    dadosBrutos =
                        XLSX.utils.sheet_to_json(
                            workbook.Sheets[firstSheet],
                            {
                                defval: ""
                            }
                        );

                }


                if (
                    !dadosBrutos ||
                    dadosBrutos.length === 0
                ) {

                    throw new Error(
                        "A planilha/CSV selecionado está vazio."
                    );

                }


                console.log(
                    "================================="
                );

                console.log(
                    "ARQUIVO CARREGADO"
                );

                console.log(
                    "Total de linhas:",
                    dadosBrutos.length
                );

                console.log(
                    "Colunas:",
                    Object.keys(
                        dadosBrutos[0]
                    )
                );

                console.log(
                    "Primeira linha:",
                    dadosBrutos[0]
                );

                console.log(
                    "================================="
                );


                /*
                 * Status visual
                 */

                if (statusBanner) {

                    statusBanner.style.display =
                        "flex";

                    statusBanner.className =
                        "status-banner";

                }


                if (statusIcon) {

                    statusIcon.innerText =
                        "🟢";

                }


                if (statusMensagem) {

                    statusMensagem.innerText =
                        "Base de dados carregada com sucesso!";

                }


                if (statusDetalhes) {

                    statusDetalhes.innerText =
                        `${dadosBrutos.length.toLocaleString("pt-BR")} linhas carregadas com sucesso!`;

                }


                /*
                 * Popular filtros
                 */

                popularFiltrosSelect(
                    dadosBrutos
                );


                /*
                 * Atualizar dashboard
                 */

                processarEAtualizar();

            }


            catch (error) {

                console.error(
                    "Erro ao processar arquivo:",
                    error
                );


                if (statusBanner) {

                    statusBanner.style.display =
                        "flex";

                    statusBanner.className =
                        "status-banner erro";

                }


                if (statusIcon) {

                    statusIcon.innerText =
                        "🔴";

                }


                if (statusMensagem) {

                    statusMensagem.innerText =
                        "Erro ao carregar o arquivo";

                }


                if (statusDetalhes) {

                    statusDetalhes.innerText =
                        error.message ||
                        "Erro desconhecido.";

                }


                if (painelErro) {

                    painelErro.style.display =
                        "block";

                }


                if (erroDetalhes) {

                    erroDetalhes.innerHTML =
                        `<strong>${escapeHTML(
                            error.message ||
                            "Erro desconhecido ao ler o arquivo."
                        )}</strong>`;

                }

            }

        };


    reader.readAsArrayBuffer(file);

}


/* ============================================================
 * ESCAPAR HTML
 * ============================================================ */

function escapeHTML(valor) {

    return String(valor || "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");

}


/* ============================================================
 * PARSER CSV
 *
 * Suporta:
 * - ;
 * - ,
 * - campos entre aspas
 * - aspas duplas dentro de campos
 *
 * ============================================================ */

function converterCSVParaArray(texto) {

    const linhas =
        texto.split(/\r\n|\n/);


    if (
        linhas.length === 0
    ) {

        return [];

    }


    /*
     * Detectar delimitador
     */

    const primeiraLinha =
        linhas[0];


    const qtdPontoVirgula =
        (
            primeiraLinha.match(
                /;/g
            ) || []
        ).length;


    const qtdVirgula =
        (
            primeiraLinha.match(
                /,/g
            ) || []
        ).length;


    const delimitador =
        qtdPontoVirgula >= qtdVirgula
            ? ";"
            : ",";


    /*
     * Parser respeitando aspas
     */

    const separarLinha =
        function (linha) {

            const resultado = [];

            let atual = "";

            let dentroAspas = false;


            for (
                let i = 0;
                i < linha.length;
                i++
            ) {

                const caractere =
                    linha[i];


                if (
                    caractere === '"'
                ) {

                    /*
                     * Aspas duplas dentro
                     * de campo:
                     *
                     * ""
                     */

                    if (
                        dentroAspas &&
                        linha[i + 1] === '"'
                    ) {

                        atual += '"';

                        i++;

                        continue;

                    }


                    dentroAspas =
                        !dentroAspas;

                    continue;

                }


                if (
                    caractere === delimitador &&
                    !dentroAspas
                ) {

                    resultado.push(
                        atual
                    );

                    atual = "";

                }


                else {

                    atual +=
                        caractere;

                }

            }


            resultado.push(
                atual
            );


            return resultado;

        };


    const headers =
        separarLinha(
            linhas[0]
        ).map(
            h =>
                String(h || "")
                    .replace(/^\uFEFF/, "")
                    .trim()
        );


    const resultado = [];


    for (
        let i = 1;
        i < linhas.length;
        i++
    ) {

        const linha =
            linhas[i];


        if (
            !linha ||
            !linha.trim()
        ) {

            continue;

        }


        const valores =
            separarLinha(
                linha
            );


        const row = {};


        headers.forEach(
            (header, index) => {

                row[header] =
                    valores[index] !== undefined
                        ? valores[index].trim()
                        : "";

            }
        );


        resultado.push(
            row
        );

    }


    return resultado;

}


/* ============================================================
 * POPULAR FILTROS
 * ============================================================ */

function popularFiltrosSelect(dados) {

    const statusFilter =
        document.getElementById(
            "statusFilter"
        );


    const categoryFilter =
        document.getElementById(
            "categoryFilter"
        );


    const projectFilter =
        document.getElementById(
            "projectFilter"
        );


    const statusSet =
        new Set();


    const catSet =
        new Set();


    const projSet =
        new Set();


    dados.forEach(
        row => {

            const status =
                extrairValorColuna(
                    row,
                    [
                        "STATUS",
                        "SITUACAO",
                        "STATUS DO PEDIDO"
                    ]
                );


            const categoria =
                extrairValorColuna(
                    row,
                    [
                        "CATEGORIA"
                    ]
                );


            const projeto =
                extrairValorColuna(
                    row,
                    [
                        "PROJETO"
                    ]
                );


            if (status) {

                statusSet.add(
                    status
                );

            }


            if (categoria) {

                catSet.add(
                    categoria
                );

            }


            if (projeto) {

                projSet.add(
                    projeto
                );

            }

        }
    );


    /*
     * STATUS
     */

    if (statusFilter) {

        statusFilter.innerHTML =
            `<option value="TODOS">
                Todos os Status
            </option>`;


        Array.from(statusSet)
            .sort(
                (a, b) =>
                    a.localeCompare(
                        b,
                        undefined,
                        {
                            numeric: true
                        }
                    )
            )
            .forEach(
                status => {

                    const option =
                        document.createElement(
                            "option"
                        );


                    option.value =
                        status;

                    option.textContent =
                        status;


                    statusFilter.appendChild(
                        option
                    );

                }
            );

    }


    /*
     * CATEGORIA
     */

    if (categoryFilter) {

        categoryFilter.innerHTML =
            `<option value="TODAS">
                Todas as Categorias
            </option>`;


        Array.from(catSet)
            .sort(
                (a, b) =>
                    a.localeCompare(
                        b,
                        undefined,
                        {
                            numeric: true
                        }
                    )
            )
            .forEach(
                categoria => {

                    const option =
                        document.createElement(
                            "option"
                        );


                    option.value =
                        categoria;

                    option.textContent =
                        categoria;


                    categoryFilter.appendChild(
                        option
                    );

                }
            );

    }


    /*
     * PROJETO
     */

    if (projectFilter) {

        projectFilter.innerHTML =
            `<option value="TODOS">
                Todos os Projetos
            </option>`;


        Array.from(projSet)
            .sort(
                (a, b) =>
                    a.localeCompare(
                        b,
                        undefined,
                        {
                            numeric: true
                        }
                    )
            )
            .forEach(
                projeto => {

                    const option =
                        document.createElement(
                            "option"
                        );


                    option.value =
                        projeto;

                    option.textContent =
                        projeto;


                    projectFilter.appendChild(
                        option
                    );

                }
            );

    }

}


/* ============================================================
 * PROCESSAR FILTROS
 * ============================================================ */

function processarEAtualizar() {

    if (
        !dadosBrutos ||
        dadosBrutos.length === 0
    ) {

        return;

    }


    const statusSel =
        document.getElementById(
            "statusFilter"
        )?.value ||
        "TODOS";


    const catSel =
        document.getElementById(
            "categoryFilter"
        )?.value ||
        "TODAS";


    const serviceSel =
        document.getElementById(
            "serviceFilter"
        )?.value ||
        "TODOS";


    const projectSel =
        document.getElementById(
            "projectFilter"
        )?.value ||
        "TODOS";


    const dtInicioInput =
        document.getElementById(
            "dtInicio"
        )?.value;


    const dtFimInput =
        document.getElementById(
            "dtFim"
        )?.value;


    /*
     * Datas dos inputs
     */

    const dtInicio =
        dtInicioInput
            ? new Date(
                `${dtInicioInput}T00:00:00`
            )
            : null;


    const dtFim =
        dtFimInput
            ? new Date(
                `${dtFimInput}T23:59:59`
            )
            : null;


    /*
     * APLICA TODOS OS FILTROS
     */

    dadosFiltrados =
        dadosBrutos.filter(
            row => {

                /*
                 * STATUS
                 */

                const status =
                    extrairValorColuna(
                        row,
                        [
                            "STATUS",
                            "SITUACAO",
                            "STATUS DO PEDIDO"
                        ]
                    );


                if (
                    statusSel !== "TODOS" &&
                    status !== statusSel
                ) {

                    return false;

                }


                /*
                 * CATEGORIA
                 */

                const categoria =
                    extrairValorColuna(
                        row,
                        [
                            "CATEGORIA"
                        ]
                    );


                if (
                    catSel !== "TODAS" &&
                    categoria !== catSel
                ) {

                    return false;

                }


                /*
                 * SERVIÇO / PRODUTO
                 */

                const tipoItem =
                    identificarTipoItem(
                        row
                    );


                if (
                    serviceSel !== "TODOS" &&
                    tipoItem !== serviceSel
                ) {

                    return false;

                }


                /*
                 * PROJETO
                 */

                const projeto =
                    extrairValorColuna(
                        row,
                        [
                            "PROJETO"
                        ]
                    );


                if (
                    projectSel !== "TODOS" &&
                    projeto !== projectSel
                ) {

                    return false;

                }


                /*
                 * DATA DE AGENDAMENTO
                 */

                if (
                    dtInicio ||
                    dtFim
                ) {

                    const agendamento =
                        extrairValorColuna(
                            row,
                            [
                                "AGENDAMENTO"
                            ]
                        );


                    if (!agendamento) {

                        return false;

                    }


                    const dataAg =
                        parseDataBR(
                            agendamento
                        );


                    if (!dataAg) {

                        return false;

                    }


                    if (
                        dtInicio &&
                        dataAg < dtInicio
                    ) {

                        return false;

                    }


                    if (
                        dtFim &&
                        dataAg > dtFim
                    ) {

                        return false;

                    }

                }


                return true;

            }
        );


    /*
     * Atualizar tudo
     */

    atualizarKPIs(
        dadosFiltrados
    );


    atualizarGraficos(
        dadosFiltrados
    );


    atualizarGraficoEvolucaoValor(
        dadosFiltrados
    );


    atualizarTabela(
        dadosFiltrados
    );


    /*
     * Atualizar mensagem
     */

    atualizarStatusFiltro(
        dadosFiltrados.length
    );

}


/* ============================================================
 * STATUS DO FILTRO
 * ============================================================ */

function atualizarStatusFiltro(
    quantidade
) {

    const statusDetalhes =
        document.getElementById(
            "statusDetalhes"
        );


    if (!statusDetalhes) {

        return;

    }


    statusDetalhes.innerText =
        `${quantidade.toLocaleString(
            "pt-BR"
        )} linhas após aplicação dos filtros.`;

}


/* ============================================================
 * KPIs
 * ============================================================ */

function atualizarKPIs(dados) {

    const pedidosUnicos =
        new Set();


    let valorTotal =
        0;


    let somaLeadTime =
        0;


    let totalLeadTime =
        0;


    dados.forEach(
        row => {

            /*
             * PEDIDO
             */

            const pedido =
                extrairValorColuna(
                    row,
                    [
                        "PEDIDO_",
                        "PEDIDO",
                        "SOLICITACAO"
                    ]
                );


            if (pedido) {

                pedidosUnicos.add(
                    pedido
                );

            }


            /*
             * VALOR
             */

            const valor =
                converterMoedaParaNumero(
                    extrairValorColuna(
                        row,
                        [
                            "VLR.TOTAL",
                            "VLRTOTAL",
                            "VALOR TOTAL",
                            "VALOR"
                        ]
                    )
                );


            valorTotal +=
                valor;


            /*
             * LEAD TIME
             */

            const status =
                extrairValorColuna(
                    row,
                    [
                        "STATUS",
                        "SITUACAO"
                    ]
                )
                    .toUpperCase()
                    .trim();


            if (
                status === "FINALIZADO"
            ) {

                const emissao =
                    parseDataBR(
                        extrairValorColuna(
                            row,
                            [
                                "EMISSAO"
                            ]
                        )
                    );


                const agendamento =
                    parseDataBR(
                        extrairValorColuna(
                            row,
                            [
                                "AGENDAMENTO"
                            ]
                        )
                    );


                if (
                    emissao &&
                    agendamento &&
                    agendamento >= emissao
                ) {

                    const diferenca =
                        agendamento -
                        emissao;


                    const dias =
                        Math.ceil(
                            diferenca /
                            (
                                1000 *
                                60 *
                                60 *
                                24
                            )
                        );


                    somaLeadTime +=
                        dias;


                    totalLeadTime++;

                }

            }

        }
    );


    const totalPedidos =
        pedidosUnicos.size;


    const ticketMedio =
        totalPedidos > 0
            ? valorTotal / totalPedidos
            : 0;


    const leadTimeMedio =
        totalLeadTime > 0
            ? (
                somaLeadTime /
                totalLeadTime
            ).toFixed(1)
            : "0";


    /*
     * Atualizar tela
     */

    const kpiPedidos =
        document.getElementById(
            "kpiPedidos"
        );


    const kpiLinhas =
        document.getElementById(
            "kpiLinhas"
        );


    const kpiLeadTime =
        document.getElementById(
            "kpiLeadTime"
        );


    const kpiValorTotal =
        document.getElementById(
            "kpiValorTotal"
        );


    const kpiTicketMedio =
        document.getElementById(
            "kpiTicketMedio"
        );


    if (kpiPedidos) {

        kpiPedidos.innerText =
            totalPedidos.toLocaleString(
                "pt-BR"
            );

    }


    if (kpiLinhas) {

        kpiLinhas.innerText =
            dados.length.toLocaleString(
                "pt-BR"
            );

    }


    if (kpiLeadTime) {

        kpiLeadTime.innerText =
            `${leadTimeMedio} dias`;

    }


    if (kpiValorTotal) {

        kpiValorTotal.innerText =
            formatarMoedaBRL(
                valorTotal
            );

    }


    if (kpiTicketMedio) {

        kpiTicketMedio.innerText =
            formatarMoedaBRL(
                ticketMedio
            );

    }

}


/* ============================================================
 * GRÁFICOS
 * ============================================================ */

function atualizarGraficos(dados) {

    const ctxStatus =
        document
            .getElementById(
                "chartStatus"
            )
            ?.getContext("2d");


    const ctxCategoria =
        document
            .getElementById(
                "chartCategoria"
            )
            ?.getContext("2d");


    const statusCounts = {};

    const categoriaCounts = {};


    dados.forEach(
        row => {

            const status =
                extrairValorColuna(
                    row,
                    [
                        "STATUS",
                        "SITUACAO"
                    ]
                ) ||
                "Sem Status";


            const categoria =
                extrairValorColuna(
                    row,
                    [
                        "CATEGORIA"
                    ]
                ) ||
                "Outros";


            statusCounts[status] =
                (
                    statusCounts[status] ||
                    0
                ) + 1;


            categoriaCounts[categoria] =
                (
                    categoriaCounts[categoria] ||
                    0
                ) + 1;

        }
    );


    /*
     * GRÁFICO STATUS
     */

    if (ctxStatus) {

        if (
            graficoStatusInstance
        ) {

            graficoStatusInstance.destroy();

        }


        graficoStatusInstance =
            new Chart(
                ctxStatus,
                {

                    type: "bar",

                    data: {

                        labels:
                            Object.keys(
                                statusCounts
                            ),

                        datasets: [

                            {

                                label:
                                    "Itens",

                                data:
                                    Object.values(
                                        statusCounts
                                    ),

                                backgroundColor:
                                    "#2563eb"

                            }

                        ]

                    },


                    options: {

                        responsive: true,

                        maintainAspectRatio:
                            false,

                        plugins: {

                            legend: {

                                display:
                                    true

                            }

                        }

                    }

                }
            );

    }


    /*
     * GRÁFICO CATEGORIA
     */

    if (ctxCategoria) {

        if (
            graficoCategoriaInstance
        ) {

            graficoCategoriaInstance.destroy();

        }


        graficoCategoriaInstance =
            new Chart(
                ctxCategoria,
                {

                    type:
                        "doughnut",

                    data: {

                        labels:
                            Object.keys(
                                categoriaCounts
                            ),

                        datasets: [

                            {

                                data:
                                    Object.values(
                                        categoriaCounts
                                    ),

                                backgroundColor: [
                                    "#16a34a",
                                    "#2563eb",
                                    "#eab308",
                                    "#dc2626",
                                    "#9333ea",
                                    "#0891b2",
                                    "#ea580c"
                                ]

                            }

                        ]

                    },


                    options: {

                        responsive: true,

                        maintainAspectRatio:
                            false,

                        plugins: {

                            legend: {

                                position:
                                    "top"

                            }

                        }

                    }

                }
            );

    }

}


/* ============================================================
 * NOVO GRÁFICO
 *
 * EVOLUÇÃO DO VALOR FINALIZADO POR SEMANA
 *
 * REGRA:
 *
 * SOMENTE STATUS = FINALIZADO
 *
 * Semana 1 = 01 até 07
 * Semana 2 = 08 até 14
 * Semana 3 = 15 até 21
 * Semana 4 = 22 até 28
 * Semana 5 = 29 até fim do mês
 *
 * DATA UTILIZADA:
 *
 * AGENDAMENTO
 *
 * VALOR:
 *
 * VLR.TOTAL
 *
 * ============================================================ */

function atualizarGraficoEvolucaoValor(
    dados
) {

    const canvas =
        document.getElementById(
            "chartEvolucaoValor"
        );


    if (!canvas) {

        return;

    }


    const ctx =
        canvas.getContext("2d");


    /*
     * Destruir gráfico anterior
     */

    if (
        graficoEvolucaoValorInstance
    ) {

        graficoEvolucaoValorInstance.destroy();

        graficoEvolucaoValorInstance =
            null;

    }


    /*
     * Objeto de agrupamento
     */

    const valoresPorSemana = {};


    dados.forEach(
        row => {

            /*
             * SOMENTE FINALIZADO
             */

            const status =
                extrairValorColuna(
                    row,
                    [
                        "STATUS",
                        "SITUACAO"
                    ]
                )
                    .normalize("NFD")
                    .replace(
                        /[\u0300-\u036f]/g,
                        ""
                    )
                    .toUpperCase()
                    .trim();


            if (
                status !== "FINALIZADO"
            ) {

                return;

            }


            /*
             * DATA DE AGENDAMENTO
             */

            const valorData =
                extrairValorColuna(
                    row,
                    [
                        "AGENDAMENTO"
                    ]
                );


            const data =
                parseDataBR(
                    valorData
                );


            if (!data) {

                return;

            }


            /*
             * VALOR DO ITEM
             */

            const valor =
                converterMoedaParaNumero(
                    extrairValorColuna(
                        row,
                        [
                            "VLR.TOTAL",
                            "VLRTOTAL",
                            "VALOR TOTAL",
                            "VALOR"
                        ]
                    )
                );


            /*
             * Ignorar zero
             */

            if (!valor) {

                return;

            }


            const ano =
                data.getFullYear();


            const mes =
                data.getMonth();


            const dia =
                data.getDate();


            /*
             * Semana do mês
             */

            const semana =
                Math.floor(
                    (dia - 1) / 7
                ) + 1;


            /*
             * Quantidade de dias
             * do mês
             */

            const ultimoDiaMes =
                new Date(
                    ano,
                    mes + 1,
                    0
                ).getDate();


            const primeiroDiaSemana =
                (
                    (semana - 1) * 7
                ) + 1;


            const ultimoDiaSemana =
                Math.min(
                    semana * 7,
                    ultimoDiaMes
                );


            /*
             * Chave única
             */

            const chave =
                `${ano}-${String(
                    mes + 1
                ).padStart(
                    2,
                    "0"
                )}-S${semana}`;


            /*
             * Criar agrupamento
             */

            if (
                !valoresPorSemana[chave]
            ) {

                valoresPorSemana[chave] = {

                    ano:
                        ano,

                    mes:
                        mes,

                    semana:
                        semana,

                    primeiroDia:
                        primeiroDiaSemana,

                    ultimoDia:
                        ultimoDiaSemana,

                    valor:
                        0

                };

            }


            /*
             * Somar valor
             */

            valoresPorSemana[chave].valor +=
                valor;

        }
    );


    /*
     * Ordenar cronologicamente
     */

    const dadosGrafico =
        Object.values(
            valoresPorSemana
        )
            .sort(
                (a, b) => {

                    if (
                        a.ano !== b.ano
                    ) {

                        return (
                            a.ano -
                            b.ano
                        );

                    }


                    if (
                        a.mes !== b.mes
                    ) {

                        return (
                            a.mes -
                            b.mes
                        );

                    }


                    return (
                        a.semana -
                        b.semana
                    );

                }
            );


    /*
     * Se não houver dados
     */

    if (
        dadosGrafico.length === 0
    ) {

        graficoEvolucaoValorInstance =
            new Chart(
                ctx,
                {

                    type:
                        "line",

                    data: {

                        labels: [
                            "Sem dados"
                        ],

                        datasets: [

                            {

                                label:
                                    "Valor Finalizado",

                                data: [
                                    0
                                ],

                                borderWidth:
                                    3

                            }

                        ]

                    },


                    options: {

                        responsive:
                            true,

                        maintainAspectRatio:
                            false

                    }

                }
            );


        return;

    }


    const nomesMeses = [

        "Jan",
        "Fev",
        "Mar",
        "Abr",
        "Mai",
        "Jun",
        "Jul",
        "Ago",
        "Set",
        "Out",
        "Nov",
        "Dez"

    ];


    /*
     * Labels
     */

    const labels =
        dadosGrafico.map(
            item => {

                return (
                    `${nomesMeses[item.mes]}/${String(
                        item.ano
                    ).slice(-2)} - S${item.semana} (${item.primeiroDia}-${item.ultimoDia})`
                );

            }
        );


    /*
     * Valores
     */

    const valores =
        dadosGrafico.map(
            item =>
                item.valor
        );


    /*
     * Calcular variação percentual
     *
     * semana contra semana anterior
     */

    const variacoes = [];


    dadosGrafico.forEach(
        (item, index) => {

            if (index === 0) {

                variacoes.push(
                    null
                );

                return;

            }


            const anterior =
                dadosGrafico[
                    index - 1
                ].valor;


            if (!anterior) {

                variacoes.push(
                    null
                );

                return;

            }


            const variacao =
                (
                    (
                        item.valor -
                        anterior
                    ) /
                    anterior
                ) * 100;


            variacoes.push(
                variacao
            );

        }
    );


    /*
     * Criar gráfico
     */

    graficoEvolucaoValorInstance =
        new Chart(
            ctx,
            {

                type:
                    "line",


                data: {

                    labels:
                        labels,


                    datasets: [

                        {

                            label:
                                "Valor Finalizado",


                            data:
                                valores,


                            borderColor:
                                "#2563eb",


                            backgroundColor:
                                "rgba(37, 99, 235, 0.10)",


                            borderWidth:
                                3,


                            tension:
                                0.30,


                            fill:
                                true,


                            pointRadius:
                                5,


                            pointHoverRadius:
                                7

                        }

                    ]

                },


                options: {

                    responsive:
                        true,


                    maintainAspectRatio:
                        false,


                    interaction: {

                        intersect:
                            false,

                        mode:
                            "index"

                    },


                    plugins: {

                        legend: {

                            display:
                                true

                        },


                        tooltip: {

                            callbacks: {

                                label:
                                    function (
                                        context
                                    ) {

                                        const index =
                                            context.dataIndex;


                                        const valor =
                                            context.raw;


                                        const variacao =
                                            variacoes[
                                                index
                                            ];


                                        let texto =
                                            ` Valor: ${formatarMoedaBRL(
                                                valor
                                            )}`;


                                        if (
                                            variacao !== null &&
                                            isFinite(
                                                variacao
                                            )
                                        ) {

                                            const sinal =
                                                variacao >= 0
                                                    ? "+"
                                                    : "";


                                            texto +=
                                                ` | Variação: ${sinal}${variacao.toFixed(
                                                    1
                                                )}%`;

                                        }


                                        return texto;

                                    }

                            }

                        }

                    },


                    scales: {

                        y: {

                            beginAtZero:
                                true,


                            ticks: {

                                callback:
                                    function (
                                        value
                                    ) {

                                        return formatarMoedaCompacta(
                                            value
                                        );

                                    }

                            }

                        }

                    }

                }

            }
        );

}


/* ============================================================
 * FORMATAÇÃO COMPACTA
 *
 * Exemplo:
 *
 * R$ 100.000
 * R$ 1,2 mi
 *
 * ============================================================ */

function formatarMoedaCompacta(
    valor
) {

    const numero =
        Number(valor || 0);


    if (
        Math.abs(numero) >= 1000000
    ) {

        return (
            "R$ " +
            (
                numero / 1000000
            ).toFixed(1) +
            " mi"
        );

    }


    if (
        Math.abs(numero) >= 1000
    ) {

        return (
            "R$ " +
            (
                numero / 1000
            ).toFixed(0) +
            " mil"
        );

    }


    return formatarMoedaBRL(
        numero
    );

}


/* ============================================================
 * TABELA
 * ============================================================ */

function atualizarTabela(
    dados
) {

    const tbody =
        document.querySelector(
            "#tabelaAgenda tbody"
        );


    if (!tbody) {

        return;

    }


    tbody.innerHTML = "";


    if (
        dados.length === 0
    ) {

        tbody.innerHTML = `

            <tr>

                <td
                    colspan="7"
                    style="
                        text-align: center;
                        color: var(--text-muted);
                    ">

                    Nenhum registro encontrado.

                </td>

            </tr>

        `;


        return;

    }


    /*
     * Limite de 100 linhas
     */

    const registros =
        dados.slice(
            0,
            100
        );


    registros.forEach(
        row => {

            const tr =
                document.createElement(
                    "tr"
                );


            const numPedido =
                extrairValorColuna(
                    row,
                    [
                        "PEDIDO_",
                        "PEDIDO"
                    ]
                );


            const item =
                extrairValorColuna(
                    row,
                    [
                        "ITEM_",
                        "ITEM"
                    ]
                );


            const projeto =
                extrairValorColuna(
                    row,
                    [
                        "PROJETO"
                    ]
                ) ||
                "-";


            const dtEmissao =
                extrairValorColuna(
                    row,
                    [
                        "EMISSAO"
                    ]
                ) ||
                "-";


            const dtAgendamento =
                extrairValorColuna(
                    row,
                    [
                        "AGENDAMENTO"
                    ]
                ) ||
                "-";


            const valor =
                formatarMoedaBRL(
                    converterMoedaParaNumero(
                        extrairValorColuna(
                            row,
                            [
                                "VLR.TOTAL",
                                "VLRTOTAL",
                                "VALOR TOTAL"
                            ]
                        )
                    )
                );


            const status =
                extrairValorColuna(
                    row,
                    [
                        "STATUS",
                        "SITUACAO"
                    ]
                );


            /*
             * Lead Time
             */

            let leadTime =
                "-";


            if (
                status
                    .toUpperCase()
                    .trim() ===
                "FINALIZADO"
            ) {

                const emissao =
                    parseDataBR(
                        dtEmissao
                    );


                const agendamento =
                    parseDataBR(
                        dtAgendamento
                    );


                if (
                    emissao &&
                    agendamento &&
                    agendamento >= emissao
                ) {

                    const diferenca =
                        agendamento -
                        emissao;


                    const dias =
                        Math.ceil(
                            diferenca /
                            (
                                1000 *
                                60 *
                                60 *
                                24
                            )
                        );


                    leadTime =
                        `${dias} dias`;

                }

            }


            /*
             * Classe do badge
             */

            const classeStatus =
                normalizarChave(
                    status
                );


            tr.innerHTML = `

                <td>
                    <strong>
                        #${escapeHTML(
                            numPedido
                        )}
                    </strong>

                    (${escapeHTML(
                        item
                    )})
                </td>


                <td>
                    ${escapeHTML(
                        projeto
                    )}
                </td>


                <td>
                    ${escapeHTML(
                        dtEmissao
                    )}
                </td>


                <td>
                    ${escapeHTML(
                        dtAgendamento
                    )}
                </td>


                <td>

                    <span
                        class="badge badge-lt">

                        ${leadTime}

                    </span>

                </td>


                <td>
                    ${valor}
                </td>


                <td>

                    <span
                        class="badge badge-${classeStatus || "default"}">

                        ${escapeHTML(
                            status
                        )}

                    </span>

                </td>

            `;


            tbody.appendChild(
                tr
            );

        }
    );

}


/* ============================================================
 * LIMPAR FILTROS
 * ============================================================ */

function limparFiltros() {

    const statusFilter =
        document.getElementById(
            "statusFilter"
        );


    const categoryFilter =
        document.getElementById(
            "categoryFilter"
        );


    const serviceFilter =
        document.getElementById(
            "serviceFilter"
        );


    const projectFilter =
        document.getElementById(
            "projectFilter"
        );


    const dtInicio =
        document.getElementById(
            "dtInicio"
        );


    const dtFim =
        document.getElementById(
            "dtFim"
        );


    if (statusFilter) {

        statusFilter.value =
            "TODOS";

    }


    if (categoryFilter) {

        categoryFilter.value =
            "TODAS";

    }


    if (serviceFilter) {

        serviceFilter.value =
            "TODOS";

    }


    if (projectFilter) {

        projectFilter.value =
            "TODOS";

    }


    if (dtInicio) {

        dtInicio.value =
            "";

    }


    if (dtFim) {

        dtFim.value =
            "";

    }


    /*
     * Reprocessar
     */

    processarEAtualizar();

}