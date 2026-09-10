#target indesign

/**
 * Script: Atualizar Tabelas InDesign com Dados do Excel
 * Versão: macOS - Lê CSV diretamente
 * 
 * Funcionamento:
 * 1. Você seleciona um arquivo CSV
 * 2. Script lê o código na PRIMEIRA COLUNA da tabela do InDesign
 * 3. Busca o código no CSV
 * 4. Preenche as células com os dados correspondentes
 * 
 * IMPORTANTE: Salve seu Excel como CSV antes de usar:
 * 1. Abra o arquivo em Excel
 * 2. File > Save As
 * 3. Formato: CSV UTF-8 (.csv)
 * 4. Use este arquivo com o script
 */

// ============================================
// VERIFICAÇÕES INICIAIS
// ============================================

if (app.documents.length === 0) {
    alert("❌ Erro: Nenhum documento InDesign aberto!\nAbra um documento antes de rodar o script.");
    exit();
}

var doc = app.activeDocument;

// ============================================
// FUNÇÃO: Verificar se objeto está vazio
// ============================================

function objetoVazio(obj) {
    for (var chave in obj) {
        if (obj.hasOwnProperty(chave)) {
            return false;
        }
    }
    return true;
}

// ============================================
// FUNÇÃO: Selecionar arquivo CSV
// ============================================

function selecionarArquivoCSV() {
    var arquivo = File.openDialog("📁 Selecione o arquivo CSV com os dados", "CSV files:*.csv");
    
    if (!arquivo) {
        alert("❌ Nenhum arquivo selecionado. Script cancelado.");
        exit();
    }
    
    return arquivo;
}

// ============================================
// FUNÇÃO: Ler arquivo CSV
// ============================================

function lerCSV(caminhoCSV) {
    var dados = {};
    
    try {
        var arquivo = new File(caminhoCSV);
        
        if (!arquivo.exists) {
            alert("❌ Arquivo CSV não encontrado:\n" + caminhoCSV);
            return dados;
        }
        
        arquivo.open("r");
        
        // Ler primeira linha (cabeçalho)
        var cabecalho = [];
        var primeiraLinha = arquivo.readln();
        cabecalho = primeiraLinha.split(",");
        
        // Limpar cabeçalhos
        for (var i = 0; i < cabecalho.length; i++) {
            cabecalho[i] = cabecalho[i].trim().replace(/^"|"$/g, "");
        }
        
        // Ler dados
        var linhasLidas = 0;
        while (!arquivo.eof) {
            var linha = arquivo.readln();
            if (linha.trim() === "") continue;
            
            var valores = parseCSVLine(linha);
            
            if (valores.length === 0) continue;
            
            var codigo = valores[0].trim();
            if (codigo === "") continue;
            
            var linhaData = {};
            for (var j = 0; j < cabecalho.length; j++) {
                linhaData[cabecalho[j]] = valores[j] ? valores[j].trim() : "";
            }
            
            dados[codigo] = linhaData;
            linhasLidas++;
        }
        
        arquivo.close();
        
        if (linhasLidas === 0) {
            alert("⚠️ Arquivo CSV vazio ou sem dados válidos.");
        }
        
    } catch (e) {
        alert("❌ Erro ao ler CSV: " + e.message);
    }
    
    return dados;
}

// ============================================
// FUNÇÃO: Parse de linha CSV com tratamento de aspas
// ============================================

function parseCSVLine(linha) {
    var valores = [];
    var valorAtual = "";
    var dentro_aspas = false;
    
    for (var i = 0; i < linha.length; i++) {
        var caractere = linha[i];
        
        if (caractere === '"') {
            if (dentro_aspas && linha[i + 1] === '"') {
                valorAtual += '"';
                i++; // Pular próxima aspas
            } else {
                dentro_aspas = !dentro_aspas;
            }
        } else if (caractere === "," && !dentro_aspas) {
            valores.push(valorAtual);
            valorAtual = "";
        } else {
            valorAtual += caractere;
        }
    }
    
    valores.push(valorAtual);
    return valores;
}

// ============================================
// FUNÇÃO: Encontrar tabelas no documento
// ============================================

function encontrarTabelas(doc) {
    var tabelas = [];
    
    for (var p = 0; p < doc.pages.length; p++) {
        var page = doc.pages[p];
        
        for (var t = 0; t < page.tables.length; t++) {
            tabelas.push({
                pagina: p,
                indice: t,
                tabela: page.tables[t],
                objeto: page.tables[t]
            });
        }
    }
    
    return tabelas;
}

// ============================================
// FUNÇÃO: Extrair código da primeira coluna
// ============================================

function extrairCodigoDaTabela(tabela, linhaDados) {
    try {
        // Primeira célula da linha de dados (coluna 0)
        var celula = tabela.rows[linhaDados].cells[0];
        var codigo = celula.texts[0].contents.trim();
        return codigo;
    } catch (e) {
        return null;
    }
}

// ============================================
// FUNÇÃO: Obter cabeçalhos da tabela InDesign
// ============================================

function obterCabecalhosTabela(tabela) {
    var cabecalhos = [];
    
    try {
        // Primeira linha tem os cabeçalhos
        for (var c = 0; c < tabela.columns.length; c++) {
            var celula = tabela.rows[0].cells[c];
            var texto = celula.texts[0].contents.trim();
            cabecalhos.push(texto);
        }
    } catch (e) {
        alert("❌ Erro ao ler cabeçalhos da tabela: " + e.message);
    }
    
    return cabecalhos;
}

// ============================================
// FUNÇÃO: Atualizar célula
// ============================================

function atualizarCelula(tabela, linhaDados, colunaDados, valor) {
    try {
        var celula = tabela.rows[linhaDados].cells[colunaDados];
        
        // Remover conteúdo anterior
        while (celula.texts.length > 0) {
            celula.texts[0].remove();
        }
        
        // Adicionar novo conteúdo
        celula.texts.add(valor);
        
        return true;
    } catch (e) {
        return false;
    }
}

// ============================================
// FUNÇÃO: Processar uma tabela
// ============================================

function processarTabela(tabela, dadosCSV) {
    var relatorio = {
        linhasProcessadas: 0,
        linhasAtualizadas: 0,
        codigosNaoEncontrados: []
    };
    
    // Obter cabeçalhos da tabela InDesign
    var cabecalhosTabela = obterCabecalhosTabela(tabela);
    
    if (cabecalhosTabela.length === 0) {
        alert("⚠️ Tabela sem cabeçalhos identificáveis.");
        return relatorio;
    }
    
    // Processar cada linha (começando em 1, pulando cabeçalho)
    for (var linhaIdx = 1; linhaIdx < tabela.rows.length; linhaIdx++) {
        // Extrair código da primeira coluna
        var codigo = extrairCodigoDaTabela(tabela, linhaIdx);
        
        if (!codigo || codigo === "") {
            continue; // Pular linhas sem código
        }
        
        relatorio.linhasProcessadas++;
        
        // Buscar dados no CSV
        var dadosProduto = dadosCSV[codigo];
        
        if (!dadosProduto) {
            relatorio.codigosNaoEncontrados.push(codigo);
            continue;
        }
        
        // Atualizar cada coluna
        for (var colIdx = 0; colIdx < cabecalhosTabela.length; colIdx++) {
            var nomeCabecalho = cabecalhosTabela[colIdx];
            var valor = dadosProduto[nomeCabecalho] || "";
            
            if (atualizarCelula(tabela, linhaIdx, colIdx, valor)) {
                relatorio.linhasAtualizadas++;
            }
        }
    }
    
    return relatorio;
}

// ============================================
// FUNÇÃO: Mostrar relatório
// ============================================

function mostrarRelatorio(relatorios) {
    var msg = "✅ ATUALIZAÇÃO CONCLUÍDA!\n\n";
    var totalTabelas = relatorios.length;
    var totalLinhas = 0;
    var totalAtualizadas = 0;
    var todosCodigosNaoEncontrados = [];
    
    for (var t = 0; t < relatorios.length; t++) {
        var rel = relatorios[t];
        msg += "📄 Tabela " + (t + 1) + ":\n";
        msg += "   • Linhas processadas: " + rel.linhasProcessadas + "\n";
        msg += "   • Células atualizadas: " + rel.linhasAtualizadas + "\n";
        
        if (rel.codigosNaoEncontrados.length > 0) {
            msg += "   • ⚠️ Códigos não encontrados: " + rel.codigosNaoEncontrados.join(", ") + "\n";
            todosCodigosNaoEncontrados = todosCodigosNaoEncontrados.concat(rel.codigosNaoEncontrados);
        }
        msg += "\n";
        
        totalLinhas += rel.linhasProcessadas;
        totalAtualizadas += rel.linhasAtualizadas;
    }
    
    msg += "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n";
    msg += "📊 RESUMO GERAL:\n";
    msg += "   • Total de tabelas: " + totalTabelas + "\n";
    msg += "   • Total de linhas processadas: " + totalLinhas + "\n";
    msg += "   • Total de células atualizadas: " + totalAtualizadas + "\n";
    
    if (todosCodigosNaoEncontrados.length > 0) {
        msg += "\n⚠️ ATENÇÃO: Alguns códigos não foram encontrados no CSV:\n";
        msg += todosCodigosNaoEncontrados.join(", ");
    }
    
    alert(msg);
}

// ============================================
// FUNÇÃO PRINCIPAL
// ============================================

function main() {
    // Selecionar arquivo CSV
    var arquivoCSV = selecionarArquivoCSV();
    
    // Ler dados do CSV
    alert("📂 Lendo dados do CSV...");
    var dadosCSV = lerCSV(arquivoCSV.fsName);
    
    if (!dadosCSV || objetoVazio(dadosCSV)) {
        alert("❌ Nenhum dado foi carregado do CSV!");
        exit();
    }
    
    // Encontrar tabelas no documento
    var tabelas = encontrarTabelas(doc);
    
    if (tabelas.length === 0) {
        alert("❌ Nenhuma tabela encontrada no documento!");
        exit();
    }
    
    alert("📊 Processando " + tabelas.length + " tabela(s)...");
    
    // Processar cada tabela
    var relatorios = [];
    for (var t = 0; t < tabelas.length; t++) {
        var relatorio = processarTabela(tabelas[t].objeto, dadosCSV);
        relatorios.push(relatorio);
    }
    
    // Mostrar relatório
    mostrarRelatorio(relatorios);
}

// ============================================
// EXECUTAR
// ============================================

main();
