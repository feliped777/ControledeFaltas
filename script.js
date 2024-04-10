function adicionarLinha() {
    var tabela = document.getElementById("tabelaProfissionais").getElementsByTagName('tbody')[0];
    var novaLinha = tabela.insertRow();
    var selecaoColuna = novaLinha.insertCell(0); // <---- Adicionando a coluna de seleção
    var matriculaColuna = novaLinha.insertCell(1);
    var nomeColuna = novaLinha.insertCell(2);
    var faltasColuna = novaLinha.insertCell(3);
    
    selecaoColuna.innerHTML = '<input type="checkbox" class="selecionarLinha">';
    matriculaColuna.innerHTML = '<input type="text" name="matricula">';
    nomeColuna.innerHTML = '<input type="text" name="nome">';
    faltasColuna.innerHTML = '<input type="number" name="faltas">';
}

function removerSelecionados() {
    var checkboxes = document.getElementsByClassName('selecionarLinha');
    var tabela = document.getElementById("tabelaProfissionais").getElementsByTagName('tbody')[0];
    
    for (var i = checkboxes.length - 1; i >= 0; i--) {
        if (checkboxes[i].checked) {
            tabela.deleteRow(i);
        }
    }
}

function exportarTabelaParaExcel() {
    console.log("Botão de exportação clicado.");

    var tabela = document.getElementById("tabelaProfissionais");
    var corpoTabela = tabela.getElementsByTagName('tbody')[0];
    var linhas = corpoTabela.getElementsByTagName('tr');
    var dados = [];

    if (linhas.length === 0) {
        alert("A tabela está vazia. Adicione dados antes de exportar para o Excel.");
        return;
    }

    for (var i = 0; i < linhas.length; i++) {
        var linha = linhas[i];
        var celulas = linha.getElementsByTagName('td');
        var linhaDados = [];

        for (var j = 1; j < celulas.length; j++) { // Começando do índice 1 para ignorar a primeira célula (checkbox)
            var input = celulas[j].querySelector('input'); // Selecionando o input dentro da célula
            linhaDados.push(input.value); // Obtendo o valor do input
        }

        dados.push(linhaDados);
    }

    var workbook = XLSX.utils.book_new();
    var worksheet = XLSX.utils.aoa_to_sheet(dados);
    XLSX.utils.book_append_sheet(workbook, worksheet, 'TabelaProfissionais');

    // Escrever o arquivo Excel diretamente
    XLSX.writeFile(workbook, 'tabela_profissionais.xlsx');
}

function importarPlanilhaDoExcel() {
    var input = document.createElement('input');
    input.type = 'file';
    input.accept = '.xlsx';
    input.onchange = function(event) {
        var file = event.target.files[0];
        var reader = new FileReader();
        reader.onload = function(e) {
            var data = new Uint8Array(e.target.result);
            var workbook = XLSX.read(data, {type: 'array'});
            var worksheet = workbook.Sheets[workbook.SheetNames[0]];
            var rows = XLSX.utils.sheet_to_json(worksheet, {header: 1});

            // Limpa a tabela atual
            var tabela = document.getElementById("tabelaProfissionais").getElementsByTagName('tbody')[0];
            tabela.innerHTML = '';

            // Adiciona as linhas da planilha
            for (var i = 0; i < rows.length; i++) {
                var row = rows[i];
                var newRow = tabela.insertRow();
                
                // Adiciona os checkboxes à primeira célula
                var selecaoColuna = newRow.insertCell(0);
                selecaoColuna.innerHTML = '<input type="checkbox" class="selecionarLinha">';
                
                // Adiciona as células restantes com os dados da planilha
                for (var j = 0; j < row.length; j++) {
                    var cell = newRow.insertCell(j + 1);
                    cell.innerHTML = '<input type="text" value="' + row[j] + '">';
                }
            }
        };
        reader.readAsArrayBuffer(file);
    };
    input.click();
}

function excluirTodasLinhas() {
    var tabela = document.getElementById("tabelaProfissionais").getElementsByTagName('tbody')[0];
    tabela.innerHTML = ''; // Limpa o conteúdo da tabela
}
