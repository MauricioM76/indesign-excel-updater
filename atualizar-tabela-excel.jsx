#target indesign

/**
 * Script: Atualizar Tabelas InDesign com Dados do Excel
 * Versão: macOS
 * 
 * Funcionamento:
 * 1. Você seleciona o arquivo Excel
 * 2. Script lê o código na PRIMEIRA COLUNA da tabela do InDesign
 * 3. Busca o código no Excel
 * 4. Preenche as células com os dados correspondentes
 * 5. Adapta-se a qualquer ordem de colunas
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
// FUNÇÃO: Ler Excel no macOS usando AppleScript
// ============================================

function lerExcelMac(caminhoExcel) {
    var dados = {};
    
    try {
        // Converter caminho para macOS
        var caminhoMac = caminhoExcel.replace(/\\/g, "/");
        
        // Script AppleScript para ler Excel
        var applescript = 'tell application "Microsoft Excel"\n' +
            'activate\n' +
            'open "' + caminhoMac + '"\n' +
            'tell active workbook\n' +
            'tell active sheet\n' +
            'set rowCount to count of rows whose value is not equal to ""\n' +
            'set colCount to count of columns\n' +
            'set allData to {}\n' +
            'repeat with i from 1 to rowCount\n' +
            'set rowData to {}\n' +
            'repeat with j from 1 to colCount\n' +
            'set cellValue to value of cell j of row i\n' +
            'set end of rowData to cellValue\n' +
            'end repeat\n' +
            'set end of allData to rowData\n' +
            'end repeat\n' +
            'return allData\n' +
            'end tell\n' +
            'end tell\n' +
            'close active workbook\n' +
            'end tell';
        
        // Executar AppleScript
        var processo = new File("/tmp/read_excel.scpt");
        processo.open("w");
        processo.write(applescript);
        processo.close();
        
        // Executar e capturar resultado
        var resultado = sistema.callSystem("osascript /tmp/read_excel.scpt");
        
        // Nota: Abordagem simplificada - usar método alternativo abaixo
        
    } catch (e) {
        // Fallback: Usar método via libreoffice/números ou CSV temporário
    }
    
    return dados;
}

// ============================================
// FUNÇÃO: Ler Excel usando conversão para CSV (macOS)
// ============================================

function lerExcelMacViaCSV(caminhoExcel) {
    var dados = {};
    
    try {
        // Criar arquivo CSV temporário
        var caminhoCSV = Folder.temp.absoluteURI + "/temp_dados.csv";
        
        // Usar Python para converter XLSX para CSV
        var pythonScript = 'import pandas as pd\n' +
            'import sys\n' +
            'try:\n' +
            '    df = pd.read_excel("' + caminhoExcel + '")\n' +
            '    df.to_csv("' + caminhoCSV + '", index=False, quoting=1)\n' +
            '    print("SUCCESS")\n' +
            'except:\n' +
            '    print("ERROR")\n';
        
        var scriptFile = new File(Folder.temp.absoluteURI + "/convert_excel.py");
        scriptFile.open("w");
        scriptFile.write(pythonScript);
        scriptFile.close();
        
        // Executar script Python
        var resultado = system.callSystem("python3 '" + scriptFile.fsName + "'");
        
        if (resultado.indexOf("SUCCESS") > -1) {
            // Ler CSV
            dados = lerCSV(caminhoCSV);
        }
        
        // Limpar arquivos temporários
        scriptFile.remove();
        
    } catch (e) {
        // Fallback para Excel via shell script
    }
    
    return dados;
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
        var linhaNum = 2;
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
        var char = linha[i];
        
        if (char === '"') {
            if (dentro_aspas && linha[i + 1] === '"') {
                valorAtual += '"';
                i++; // Pular próxima aspas
            } else {
                dentro_aspas = !dentro_aspas;
            }
        } else if (char === "," && !dentro_aspas) {
            valores.push(valorAtual);
            valorAtual = "";
        } else {
            valorAtual += char;
        }
    }
    
    valores.push(valorAtual);
    return valores;
}

// ============================================
// FUNÇÃO: Ler Excel (macOS genérica)
// ============================================

function lerDadosExcel(caminhoExcel) {
    var dados = null;
    
    try {
        // Tentar método via CSV
        dados = lerExcelMacViaCSV(caminhoExcel);
        
        if (!dados || Object.keys(dados).length === 0) {
            alert("⚠️ Não foi possível ler o Excel com Python.\nCertifique-se de ter pandas instalado:\npip3 install pandas openpyxl");
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
    alert("📂 Lendo dados do Excel...\n(Aguarde, pode levar alguns segundos)");
    var dadosExcel = lerDadosExcel(arquivoExcel.fsName);
    
    if (!dadosExcel || Object.keys(dadosExcel).length === 0) {
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
