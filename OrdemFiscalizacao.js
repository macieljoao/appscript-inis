function onOpen() {
    SpreadsheetApp.getUi().createMenu("📖 Entrada de Dados 📖").addItem("🎫 Cadastrar Auto de Infracao 🎫", "showSidebar").addToUi();
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu('🤖 Preenchimento Automatico 🤖');
    menu.addItem('Criar Certidao - Atividade Licenciada 😳', 'criarNovaCertidao1')
    menu.addItem('Criar Certidao - Atividade Não-Licenciada 💣', 'criarNovaCertidao2')
    menu.addToUi();
}

function showSidebar() {
    SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutputFromFile("Sidebar.html").setTitle("Entre os Dados"));
}

function criarNovaCertidao1() {
    //ID do documento-padrão
    const googleDocTemplate = DriveApp.getFileById('1R3r4zJ2AXcwP4nw6y_UYHxgXg7q-wxXGClFMHfkdlZI');

    //ID da pasta de destino
    const destinationFolder = DriveApp.getFolderById('1sNydRwNcQcEmAUzh31q66w9lcBqBXEUN')

    //Definição da planilha como uma variável
    let sheet = SpreadsheetApp
        .getActiveSpreadsheet()
        .getSheetByName('Script')

    //Armazenamento de todos os valores como um array de duas dimensões
    let rows = sheet.getDataRange().getValues();

    //Processamento em cada linha
    rows.forEach(function (row, index) {
        //Se a linha é o cabeçalho, retorna
        if (index === 0) return;
        //Se um documento já foi gerado, verificando o link do documento, retorna
        if (row[5]) return;
        //Cópia do documento-padrão inserindo os dados da planilha
        const copy = googleDocTemplate.makeCopy(`Ordem Fiscalizacao ${row[1]}`, destinationFolder)
        //Once we have the copy, we then open it using the DocumentApp
        const doc = DocumentApp.openById(copy.getId())
        //Pegar o 'body' do documento para editar
        const body = doc.getBody();

        //Cada linha troca o parâmetro definido pelo valor da linha
        body.replaceText('{{SIPE}}', row[0]);
        body.replaceText('{{RelatOcorrencia}}', row[1]);
        body.replaceText('{{Licenca}}', row[2]);
        body.replaceText('{{DataHoje}}', row[3]);

        //Salva o documento e fecha
        doc.saveAndClose();
        //Armazena o link do documento em uma célula
        const url = doc.getUrl();
        sheet.getRange(index + 1, 5).setValue(url)

    })
}

function criarNovaCertidao2() {
    //ID do documento-padrão
    const googleDocTemplate = DriveApp.getFileById('15up_mjaFnfd_wudgC50OnMdkjvRQAR_EBU7XuAfkwBY');

    //ID da pasta de destino
    const destinationFolder = DriveApp.getFolderById('1sNydRwNcQcEmAUzh31q66w9lcBqBXEUN')

    //Definição da planilha como uma variável
    let sheet = SpreadsheetApp
        .getActiveSpreadsheet()
        .getSheetByName('Script')

    //Armazenamento de todos os valores como um array de duas dimensões
    let rows = sheet.getDataRange().getValues();

    //Processamento em cada linha
    rows.forEach(function (row, index) {
        //Se a linha é o cabeçalho, retorna
        if (index === 0) return;
        //Se um documento já foi gerado, verificando o link do documento, retorna
        if (row[4]) return;
        //Cópia do documento-padrão inserindo os dados da planilha
        const copy = googleDocTemplate.makeCopy(`Ordem Fiscalizacao ${row[1]}`, destinationFolder)
        //Once we have the copy, we then open it using the DocumentApp
        const doc = DocumentApp.openById(copy.getId())
        //Pegar o 'body' do documento para editar
        const body = doc.getBody();

        //Cada linha troca o parâmetro definido pelo valor da linha
        body.replaceText('{{SIPE}}', row[0]);
        body.replaceText('{{RelatOcorrencia}}', row[1]);
        body.replaceText('{{DataHoje}}', row[2]);

        //Salva o documento e fecha
        doc.saveAndClose();
        //Armazena o link do documento em uma célula
        const url = doc.getUrl();
        sheet.getRange(index + 1, 4).setValue(url)

    })
}