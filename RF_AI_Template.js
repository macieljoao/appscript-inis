function onOpen() {
    SpreadsheetApp.getUi().createMenu("📄 Entrada de Dados 📄").addItem("Adicionar Relatorio de Fiscalizacao", "showSidebar").addToUi();
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu('🤖 Preenchimento Automatico 🤖');
    menu.addItem('Criar Novo Relatorio de Fiscalizacao', 'criarNovoRelatFisc')
    menu.addToUi();
}

function showSidebar() {
    SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutputFromFile("Sidebar.html").setTitle("Entre os Dados"));
}

function criarNovoRelatFisc() {
    //ID do documento-padrão
    const googleDocTemplate = DriveApp.getFileById('1PIJYbiTX69lxA3u-KoP-GIX3g5fKeRDuwQ-U0KUuW2c');

    //ID da pasta de destino
    const destinationFolder = DriveApp.getFolderById('1KSv73oHDO5cNe0GS9xcqaVoWziKOm9lw')

    //Definição da planilha como uma variável
    const sheet = SpreadsheetApp
        .getActiveSpreadsheet()
        .getSheetByName('DadosEntrada')

    //Armazenamento de todos os valores como um array de duas dimensões
    const rows = sheet.getDataRange().getValues();

    //Processamento em cada linha
    rows.forEach(function (row, index) {
        //Se a linha é o cabeçalho, retorna
        if (index === 0) return;
        //Se um documento já foi gerado, verificando o 'link do documento, retorna
        if (row[34]) return;
        //Cópia do documento-padrão inserindo os dados da planilha
        const copy = googleDocTemplate.makeCopy(`Relatorio de Fiscalizacao ${row[0]}`, destinationFolder)
        //Once we have the copy, we then open it using the DocumentApp
        const doc = DocumentApp.openById(copy.getId())
        //Pegar o 'body' do documento para editar
        const body = doc.getBody();

        //Cada linha troca o parâmetro definido pelo valor da linha
        body.replaceText('{{RelatFisc}}', row[0]);
        body.replaceText('{{IDInfrator}}', row[1]);
        body.replaceText('{{CNPJInfrator}}', row[2]);
        body.replaceText('{{EndInfrator}}', row[3]);
        body.replaceText('{{EndInfracao}}', row[4]);
        body.replaceText('{{CoordX}}', row[5]);
        body.replaceText('{{CoordY}}', row[6]);
        body.replaceText('{{Fiscal}}', row[7]);
        body.replaceText('{{DataConstat}}', row[8]);
        body.replaceText('{{InfraAdm}}', row[9]);
        body.replaceText('{{MedidaAdot}}', row[10]);
        body.replaceText('{{GrauLesividade}}', row[11]);
        body.replaceText('{{NivelGravidade}}', row[12]);
        body.replaceText('{{ValorMulta}}', row[13]);
        body.replaceText('{{CondEconomica}}', row[14]);
        body.replaceText('{{Agravante}}', row[15]);
        body.replaceText('{{Atenuante}}', row[16]);
        body.replaceText('{{Reincidencia}}', row[17]);
        body.replaceText('{{Licenca}}', row[18]);
        body.replaceText('{{DescricaoInfra}}', row[19]);
        body.replaceText('{{MotivCondut}}', row[20]);
        body.replaceText('{{EfeitoAmb}}', row[21]);
        body.replaceText('{{EfeitoSaude}}', row[22]);
        body.replaceText('{{MotivCondutValor}}', row[23]);
        body.replaceText('{{EfeitoAmbValor}}', row[24]);
        body.replaceText('{{EfeitoSaudeValor}}', row[25]);
        body.replaceText('{{MatriculaFiscal}}', row[26]);
        body.replaceText('{{FormacaoFiscal}}', row[27]);
        body.replaceText('{{PortariaFiscal}}', row[28]);
        body.replaceText('{{PorcenAten}}', row[29]);
        body.replaceText('{{PorcenAgrav}}', row[30]);
        body.replaceText('{{NatJur}}', row[31]);
        body.replaceText('{{AI}}', row[32]);


        //Salva o documento e fecha
        doc.saveAndClose();
        //Armazena o link do documento em uma célula
        const url = doc.getUrl();
        sheet.getRange(index + 1, 34).setValue(url)

    })

}
