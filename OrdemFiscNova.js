function onOpen() {

    const ui = SpreadsheetApp.getUi();
    ui.createMenu("📄 Entrada de Dados 📄")
        .addItem("🤖 Adicionar Demanda de Fiscalizacao 🤖", "showSidebar")
        .addToUi();
}

function showSidebar() {
    SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutputFromFile("Sidebar.html").setTitle("Entre os Dados"));
}

function OrdemFiscalizacaoNova() {

    const googleDocTemplate = DriveApp.getFileById('15up_mjaFnfd_wudgC50OnMdkjvRQAR_EBU7XuAfkwBY');
    const destinationFolder = DriveApp.getFolderById('1sNydRwNcQcEmAUzh31q66w9lcBqBXEUN');

    let sheet = SpreadsheetApp
        .getActiveSpreadsheet()
        .getSheetByName('Ordem_Fiscalizacao_Nova')

    let rows = sheet.getDataRange().getValues();

    rows.forEach(function (row, index) {
        if (index === 0) return;
        if (row[5]) return;
        const copy = googleDocTemplate.makeCopy(`Ordem Fiscalizacao ${row[1]}`, destinationFolder)
        const doc = DocumentApp.openById(copy.getId())
        const body = doc.getBody();

        //Cada linha troca o parâmetro definido pelo valor da linha
        body.replaceText('{{RelatOcorrencia}}', row[0]);
        body.replaceText('{{SIPE}}', row[1]);
        body.replaceText('{{DataHoje}}', row[2]);

        //Salva o documento e fecha
        doc.saveAndClose();
        //Armazena o link do documento em uma célula
        const url = doc.getUrl();
        sheet.getRange(index + 1, 5).setValue(url)

    })
}
