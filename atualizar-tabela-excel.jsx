#target indesign

/**
 * Script: Atualizar Tabelas InDesign com Dados do Excel
 * Versão: macOS - SEM DEPENDÊNCIA DE PYTHON
 * 
 * Funcionamento:
 * 1. Você seleciona o arquivo Excel
 * 2. Script converte para CSV automaticamente (usando AppleScript)
 * 3. Lê o código na PRIMEIRA COLUNA da tabela do InDesign
 * 4. Busca o código no Excel
 * 5. Preenche as células com os dados correspondentes
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
// FUNÇÃO: Selecionar arquivo Excel
// ============================================

function selecionarArquivoExcel() {
    var arquivo = File.openDialog("📁 Selecione o arquivo Excel com os dados", "Excel files:*.xlsx,*.xls");
    
    if (!arquivo) {
        alert("❌ Nenhum arquivo selecionado. Script cancelado.");
        exit();
    }
    
    return arquivo;
}

// ============================================
// FUNÇÃO: Converter Excel para CSV via AppleScript
// ============================================

function converterExcelParaCSV(caminhoExcel) {
    var caminhoCSV = Folder.temp.absoluteURI + "/temp_dados.csv";
    
    try {
        // AppleScript que abre Excel, exporta como CSV e fecha
        var applescript = 'tell application "Microsoft Excel"\n' +
            '    activate\n' +
            '    set docPath to "' + caminhoExcel + '"\n' +
            '    open file docPath\n' +
            '    tell active workbook\n' +
            '        tell active sheet\n' +
            '            save as it to "' + caminhoCSV + '" file format CSV file format\n' +
            '        end tell\n' +
            '        close without saving\n' +
            '    end tell\n' +
            'end tell';
        
        // Salvar AppleScript em arquivo temporário
        var scriptFile = new File(Folder.temp.absoluteURI + "/excel_to_csv.scpt");
        scriptFile.open("w");
        scriptFile.write(applescript);
        scriptFile.close();
        
        // Executar AppleScript
        system.callSystem("osascript '" + scriptFile.fsName + "'");
        
        // Pequeno delay para garantir que o arquivo foi criado
        $.sleep(2000);
        
        // Verificar se arquivo CSV foi criado
        var csvFile = new File(caminhoCSV);
        if (csvFile.exists) {
            scriptFile.remove();
            return caminhoCSV;
        } else {
            alert("❌ Erro ao converter Excel para CSV");
            scriptFile.remove();
            return null;
        }
        
    } catch (e) {
        alert("❌ Erro ao converter Excel: " + e.message + "\n\nCertifique-se de ter Microsoft Excel instalado.");
        return null;
    }
}

// ============================================
// FUNÇÃO: Ler arquivo CSV
// ============================================

function lerCSV(caminhoCSV) {
    var dados = {};
    
    try {
        var arquivo = new File(caminhoCSV);
        
        if (!arquivo.exists) {
            alert("❌ Arquivo CSV não encontrado");
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
        }
        
        arquivo.close();
        
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
// FUNÇÃO: Ler Excel (macOS via Excel nativo)
// ============================================

function lerDadosExcel(caminhoExcel) {
    var dados = null;
    
    try {
        // Converter para CSV usando Excel nativo
        var caminhoCSV = converterExcelParaCSV(caminhoExcel);
        
        if (!caminhoCSV) {
            alert("❌ Falha ao converter Excel para CSV");
            return null;
        }
        
        // Ler CSV
        dados = lerCSV(caminhoCSV);
        
        if (!dados || objetoVazio(dados)) {
            alert("⚠️ Nenhum dado foi lido do Excel.");
            return null;
        }
        
    } catch (e) {
        alert("❌ Erro ao processar Excel: " + e.message);
        return null;
    }
    
    return dados;
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

function processarTabela(tabela, dadosExcel) {
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
        
        // Buscar dados no Excel
        var dadosProduto = dadosExcel[codigo];
        
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
        msg += "\n⚠️ ATENÇÃO: Alguns códigos não foram encontrados no Excel:\n";
        msg += todosCodigosNaoEncontrados.join(", ");
    }
    
    alert(msg);
}

// ============================================
// FUNÇÃO PRINCIPAL
// ============================================

function main() {
    // Selecionar arquivo Excel
    var arquivoExcel = selecionarArquivoExcel();
    
    // Ler dados do Excel
    alert("📂 Convertendo Excel para CSV...\n(Aguarde, pode levar alguns segundos)");
    var dadosExcel = lerDadosExcel(arquivoExcel.fsName);
    
    if (!dadosExcel || objetoVazio(dadosExcel)) {
        alert("❌ Nenhum dado foi carregado do Excel!");
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
        var relatorio = processarTabela(tabelas[t].objeto, dadosExcel);
        relatorios.push(relatorio);
    }
    
    // Mostrar relatório
    mostrarRelatorio(relatorios);
}

// ============================================
// EXECUTAR
// ============================================

main();
