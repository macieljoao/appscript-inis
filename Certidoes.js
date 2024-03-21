function onOpen() {
    SpreadsheetApp.getUi().createMenu("📖 Entrada de Dados 📖").addItem("🎫 Cadastrar Auto de Infracao 🎫", "showSidebar").addToUi();
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu('🤖 Preenchimento Automatico 🤖');
    menu.addItem('Criar Certidao - Apresentou Defesa 😳', 'criarNovaCertidao1')
    menu.addItem('Criar Certidao - Nao Apresentou Defesa 💣', 'criarNovaCertidao2')
    menu.addItem('Criar Certidao - Audiencia Marcada 😐', 'criarNovaCertidao3')
    menu.addItem('Criar Certidao - Audiencia Sucesso 😍', 'criarNovaCertidao4')
    menu.addItem('Criar Certidao - Audiencia Fracasso 😡', 'criarNovaCertidao5')
    menu.addItem('Criar Certidao - Contradita OK 👍', 'criarNovaCertidao6')
    menu.addItem('Criar Certidao - Aguardando Julgamento 🐶', 'criarNovaCertidao7')
    menu.addItem('Criar Certidao - Julgamento OK 🤔', 'criarNovaCertidao8')
    menu.addItem('Criar Certidao - Recurso COMDEMA 💩', 'criarNovaCertidao9')
    menu.addItem('Criar Certidao - PRAD OK 👩‍🌾', 'criarNovaCertidao10')
    menu.addItem('Criar Certidao - Negativa Debito 💸', 'criarNovaCertidao11')
    menu.addToUi();
}

function showSidebar() {
    SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutputFromFile("Sidebar.html").setTitle("Entre os Dados"));
}

function criarNovaCertidao1() {
    //ID do documento-padrão
    const googleDocTemplate = DriveApp.getFileById('14ixtNocUPEnD7HxfkLPxuDp2QZSBz22qw-mru4pjkyo');

    //ID da pasta de destino
    const destinationFolder = DriveApp.getFolderById('1jsHVNMdLWS1lUqPFhEv0wc239EhvdzDX')

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
        if (row[26]) return;
        //Cópia do documento-padrão inserindo os dados da planilha
        const copy = googleDocTemplate.makeCopy(`Certidao ${row[1]}`, destinationFolder)
        //Once we have the copy, we then open it using the DocumentApp
        const doc = DocumentApp.openById(copy.getId())
        //Pegar o 'body' do documento para editar
        const body = doc.getBody();

        //Cada linha troca o parâmetro definido pelo valor da linha
        body.replaceText('{{SIPE}}', row[0]);
        body.replaceText('{{AI}}', row[1]);
        body.replaceText('{{IDInfrator}}', row[2]);
        body.replaceText('{{DataCienciaAutuacao}}', row[3]);
        body.replaceText('{{DataPrazoDefesa}}', row[4]);
        body.replaceText('{{DataApresentacaoDefesa}}', row[5]);
        body.replaceText('{{DataHoje}}', row[6]);

        //Salva o documento e fecha
        doc.saveAndClose();
        //Armazena o link do documento em uma célula
        const url = doc.getUrl();
        sheet.getRange(index + 1, 26).setValue(url)

    })
}

function criarNovaCertidao2() {
    //ID do documento-padrão
    const googleDocTemplate = DriveApp.getFileById('18Nl4Y21ySXn4r3XfV8F_WlSRucocS2iIvKOUP1IKwZo');

    //ID da pasta de destino
    const destinationFolder = DriveApp.getFolderById('1jsHVNMdLWS1lUqPFhEv0wc239EhvdzDX')

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
        if (row[26]) return;
        //Cópia do documento-padrão inserindo os dados da planilha
        const copy = googleDocTemplate.makeCopy(`Certidao ${row[1]}`, destinationFolder)
        //Once we have the copy, we then open it using the DocumentApp
        const doc = DocumentApp.openById(copy.getId())
        //Pegar o 'body' do documento para editar
        const body = doc.getBody();

        //Cada linha troca o parâmetro definido pelo valor da linha
        body.replaceText('{{SIPE}}', row[0]);
        body.replaceText('{{AI}}', row[1]);
        body.replaceText('{{IDInfrator}}', row[2]);
        body.replaceText('{{DataCienciaAutuacao}}', row[3]);
        body.replaceText('{{DataPrazoDefesa}}', row[4]);
        body.replaceText('{{DataNotificacaoAlegFin}}', row[7]);
        body.replaceText('{{DataHoje}}', row[6]);

        //Salva o documento e fecha
        doc.saveAndClose();
        //Armazena o link do documento em uma célula
        const url = doc.getUrl();
        sheet.getRange(index + 1, 26).setValue(url)

    })
}

function criarNovaCertidao3() {
    //ID do documento-padrão
    const googleDocTemplate = DriveApp.getFileById('1Fl50vk5PGV7-uq8gMWw_DYeLosErfYy_gzQjFTCUniA');

    //ID da pasta de destino
    const destinationFolder = DriveApp.getFolderById('1jsHVNMdLWS1lUqPFhEv0wc239EhvdzDX')

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
        if (row[26]) return;
        //Cópia do documento-padrão inserindo os dados da planilha
        const copy = googleDocTemplate.makeCopy(`Certidao ${row[1]}`, destinationFolder)
        //Once we have the copy, we then open it using the DocumentApp
        const doc = DocumentApp.openById(copy.getId())
        //Pegar o 'body' do documento para editar
        const body = doc.getBody();

        //Cada linha troca o parâmetro definido pelo valor da linha
        body.replaceText('{{SIPE}}', row[0]);
        body.replaceText('{{AI}}', row[1]);
        body.replaceText('{{IDInfrator}}', row[2]);
        body.replaceText('{{DataHoje}}', row[6]);
        body.replaceText('{{DataAudiencia}}', row[8]);

        //Salva o documento e fecha
        doc.saveAndClose();
        //Armazena o link do documento em uma célula
        const url = doc.getUrl();
        sheet.getRange(index + 1, 26).setValue(url)

    })
}

function criarNovaCertidao4() {
    //ID do documento-padrão
    const googleDocTemplate = DriveApp.getFileById('1Dot0PvcwckdoOtglpTpSgG6NL0pCSVH--td3Lhu7_N4');

    //ID da pasta de destino
    const destinationFolder = DriveApp.getFolderById('1jsHVNMdLWS1lUqPFhEv0wc239EhvdzDX')

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
        if (row[26]) return;
        //Cópia do documento-padrão inserindo os dados da planilha
        const copy = googleDocTemplate.makeCopy(`Certidao ${row[1]}`, destinationFolder)
        //Once we have the copy, we then open it using the DocumentApp
        const doc = DocumentApp.openById(copy.getId())
        //Pegar o 'body' do documento para editar
        const body = doc.getBody();

        //Cada linha troca o parâmetro definido pelo valor da linha
        body.replaceText('{{SIPE}}', row[0]);
        body.replaceText('{{AI}}', row[1]);
        body.replaceText('{{IDInfrator}}', row[2]);
        body.replaceText('{{DataHoje}}', row[6]);
        body.replaceText('{{DataAudiencia}}', row[8]);

        //Salva o documento e fecha
        doc.saveAndClose();
        //Armazena o link do documento em uma célula
        const url = doc.getUrl();
        sheet.getRange(index + 1, 26).setValue(url)

    })
}

function criarNovaCertidao5() {
    //ID do documento-padrão
    const googleDocTemplate = DriveApp.getFileById('18rWvDtGFf5nk0XARRYTc9Ek-Aa4i7A33NSvsUF9LZuM');

    //ID da pasta de destino
    const destinationFolder = DriveApp.getFolderById('1jsHVNMdLWS1lUqPFhEv0wc239EhvdzDX')

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
        if (row[26]) return;
        //Cópia do documento-padrão inserindo os dados da planilha
        const copy = googleDocTemplate.makeCopy(`Certidao ${row[1]}`, destinationFolder)
        //Once we have the copy, we then open it using the DocumentApp
        const doc = DocumentApp.openById(copy.getId())
        //Pegar o 'body' do documento para editar
        const body = doc.getBody();

        //Cada linha troca o parâmetro definido pelo valor da linha
        body.replaceText('{{SIPE}}', row[0]);
        body.replaceText('{{AI}}', row[1]);
        body.replaceText('{{IDInfrator}}', row[2]);
        body.replaceText('{{DataHoje}}', row[6]);
        body.replaceText('{{DataAudiencia}}', row[8]);

        //Salva o documento e fecha
        doc.saveAndClose();
        //Armazena o link do documento em uma célula
        const url = doc.getUrl();
        sheet.getRange(index + 1, 26).setValue(url)

    })
}

function criarNovaCertidao6() {
    //ID do documento-padrão
    const googleDocTemplate = DriveApp.getFileById('1qf43RAqPStlmJ_JPinqHevhZH34iekwHh43aTKyLS04');

    //ID da pasta de destino
    const destinationFolder = DriveApp.getFolderById('1jsHVNMdLWS1lUqPFhEv0wc239EhvdzDX')

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
        if (row[26]) return;
        //Cópia do documento-padrão inserindo os dados da planilha
        const copy = googleDocTemplate.makeCopy(`Certidao ${row[1]}`, destinationFolder)
        //Once we have the copy, we then open it using the DocumentApp
        const doc = DocumentApp.openById(copy.getId())
        //Pegar o 'body' do documento para editar
        const body = doc.getBody();

        //Cada linha troca o parâmetro definido pelo valor da linha
        body.replaceText('{{SIPE}}', row[0]);
        body.replaceText('{{AI}}', row[1]);
        body.replaceText('{{IDInfrator}}', row[2]);
        body.replaceText('{{DataHoje}}', row[6]);
        body.replaceText('{{DataNotificacaoAlegFin}}', row[7]);

        //Salva o documento e fecha
        doc.saveAndClose();
        //Armazena o link do documento em uma célula
        const url = doc.getUrl();
        sheet.getRange(index + 1, 26).setValue(url)

    })
}

function criarNovaCertidao7() {
    //ID do documento-padrão
    const googleDocTemplate = DriveApp.getFileById('1qlzvpWws5rwqAjj9LehUKLFlTTNlgce0Ul5ma0sXAVY');

    //ID da pasta de destino
    const destinationFolder = DriveApp.getFolderById('1jsHVNMdLWS1lUqPFhEv0wc239EhvdzDX')

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
        if (row[26]) return;
        //Cópia do documento-padrão inserindo os dados da planilha
        const copy = googleDocTemplate.makeCopy(`Certidao ${row[1]}`, destinationFolder)
        //Once we have the copy, we then open it using the DocumentApp
        const doc = DocumentApp.openById(copy.getId())
        //Pegar o 'body' do documento para editar
        const body = doc.getBody();

        //Cada linha troca o parâmetro definido pelo valor da linha
        body.replaceText('{{SIPE}}', row[0]);
        body.replaceText('{{AI}}', row[1]);
        body.replaceText('{{IDInfrator}}', row[2]);
        body.replaceText('{{DataHoje}}', row[6]);
        body.replaceText('{{DataApresentacaoAlegFin}}', row[9]);

        //Salva o documento e fecha
        doc.saveAndClose();
        //Armazena o link do documento em uma célula
        const url = doc.getUrl();
        sheet.getRange(index + 1, 26).setValue(url)

    })
}

function criarNovaCertidao8() {
    //ID do documento-padrão
    const googleDocTemplate = DriveApp.getFileById('1DZFegzrYKIu076WRztXby021FnyTV2heVEnG3NfMfAs');

    //ID da pasta de destino
    const destinationFolder = DriveApp.getFolderById('1jsHVNMdLWS1lUqPFhEv0wc239EhvdzDX')

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
        if (row[26]) return;
        //Cópia do documento-padrão inserindo os dados da planilha
        const copy = googleDocTemplate.makeCopy(`Certidao ${row[1]}`, destinationFolder)
        //Once we have the copy, we then open it using the DocumentApp
        const doc = DocumentApp.openById(copy.getId())
        //Pegar o 'body' do documento para editar
        const body = doc.getBody();

        //Cada linha troca o parâmetro definido pelo valor da linha
        body.replaceText('{{SIPE}}', row[0]);
        body.replaceText('{{AI}}', row[1]);
        body.replaceText('{{IDInfrator}}', row[2]);
        body.replaceText('{{DataHoje}}', row[6]);
        body.replaceText('{{DataJulgamento}}', row[10]);
        body.replaceText('{{DataCienciaJulgamento}}', row[11]);

        //Salva o documento e fecha
        doc.saveAndClose();
        //Armazena o link do documento em uma célula
        const url = doc.getUrl();
        sheet.getRange(index + 1, 26).setValue(url)

    })
}

function criarNovaCertidao9() {
    //ID do documento-padrão
    const googleDocTemplate = DriveApp.getFileById('1FbjP5Q0Rg0JF8z6_cW-FY2qELAhzZGBH_xyFFeRKXx4');

    //ID da pasta de destino
    const destinationFolder = DriveApp.getFolderById('1jsHVNMdLWS1lUqPFhEv0wc239EhvdzDX')

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
        if (row[26]) return;
        //Cópia do documento-padrão inserindo os dados da planilha
        const copy = googleDocTemplate.makeCopy(`Certidao ${row[1]}`, destinationFolder)
        //Once we have the copy, we then open it using the DocumentApp
        const doc = DocumentApp.openById(copy.getId())
        //Pegar o 'body' do documento para editar
        const body = doc.getBody();

        //Cada linha troca o parâmetro definido pelo valor da linha
        body.replaceText('{{SIPE}}', row[0]);
        body.replaceText('{{AI}}', row[1]);
        body.replaceText('{{IDInfrator}}', row[2]);
        body.replaceText('{{DataHoje}}', row[6]);
        body.replaceText('{{DataApresentacaoRecurso}}', row[12]);
        body.replaceText('{{DataEncCOMDEMA}}', row[13]);

        //Salva o documento e fecha
        doc.saveAndClose();
        //Armazena o link do documento em uma célula
        const url = doc.getUrl();
        sheet.getRange(index + 1, 26).setValue(url)

    })
}

function criarNovaCertidao10() {
    //ID do documento-padrão
    const googleDocTemplate = DriveApp.getFileById('1Vio4jwQpxUznLbhthjQOxB-3b5vtzWELnf-H0-rJPUc');

    //ID da pasta de destino
    const destinationFolder = DriveApp.getFolderById('1jsHVNMdLWS1lUqPFhEv0wc239EhvdzDX')

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
        if (row[26]) return;
        //Cópia do documento-padrão inserindo os dados da planilha
        const copy = googleDocTemplate.makeCopy(`Certidao ${row[1]}`, destinationFolder)
        //Once we have the copy, we then open it using the DocumentApp
        const doc = DocumentApp.openById(copy.getId())
        //Pegar o 'body' do documento para editar
        const body = doc.getBody();

        //Cada linha troca o parâmetro definido pelo valor da linha
        body.replaceText('{{SIPE}}', row[0]);
        body.replaceText('{{AI}}', row[1]);
        body.replaceText('{{IDInfrator}}', row[2]);
        body.replaceText('{{DataHoje}}', row[6]);
        body.replaceText('{{DataApresentacaoPRAD}}', row[14]);
        body.replaceText('{{ProcessoPRAD}}', row[15]);

        //Salva o documento e fecha
        doc.saveAndClose();
        //Armazena o link do documento em uma célula
        const url = doc.getUrl();
        sheet.getRange(index + 1, 26).setValue(url)

    })
}

function criarNovaCertidao11() {
    //ID do documento-padrão
    const googleDocTemplate = DriveApp.getFileById('1Zenr27YSevOL4alBGGBQisAyUYNI9Q7SdsViPH2pgHg');

    //ID da pasta de destino
    const destinationFolder = DriveApp.getFolderById('1jsHVNMdLWS1lUqPFhEv0wc239EhvdzDX')

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
        if (row[26]) return;
        //Cópia do documento-padrão inserindo os dados da planilha
        const copy = googleDocTemplate.makeCopy(`Certidao ${row[1]}`, destinationFolder)
        //Once we have the copy, we then open it using the DocumentApp
        const doc = DocumentApp.openById(copy.getId())
        //Pegar o 'body' do documento para editar
        const body = doc.getBody();

        //Cada linha troca o parâmetro definido pelo valor da linha
        body.replaceText('{{SIPE}}', row[0]);
        body.replaceText('{{IDInfrator}}', row[2]);
        body.replaceText('{{DataHoje}}', row[6]);
        body.replaceText('{{CNPJInfrator}}', row[16]);


        //Salva o documento e fecha
        doc.saveAndClose();
        //Armazena o link do documento em uma célula
        const url = doc.getUrl();
        sheet.getRange(index + 1, 26).setValue(url)

    })
}