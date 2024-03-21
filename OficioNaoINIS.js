function OficioNaoINIS() {
    const templateOptions = ['Licenca_IMA',
        'Loteamento_Clandestino',
        'Loteamento_DescumprindoProj',
        'Invasao_ComCorte',
        'Invasao_SemCorte_APP',
        'SossegoEmpresa_DispLIC',
        'SossegoObra_DispLIC',
        'FumacaChamine',
        'EfluenteCasa_SEDUH'];

    const docTemplateID = ['1U4XhhN7C2nz7EYSsagrm5CMCM2cbR9Y8wFWc-tAkl9E',
        '1FQJiejtq0azKQTzMK2_SJS63jH6m0wvoafEbE_8wuFk',
        '1a_tdFDVTkyEs1pjtTtVM207rtY8uoeTiH1XT8pXje7Y',
        '1lO63NCZocF0IKYxr0NP6MePhxMj3NM5sc0rZoAvXl8g',
        '1bjK0JtaoZsoUeIew9XOH-2WnoP-PChnuyKsEqw0ydPM',
        '1rSLTGivffZ6azceO3iTKEEk1WccVPWMAZS3yHUFYGMs',
        '1vVssUTRGKZPu0LtFvWQsg0rk_uuTan7V1VG9gRVFAVo',
        '1TPDca21ShuU2eaXIv7lkzbPoFzAALVAM5JOo9YR8gXY',
        '1Jc9vzhi6WeKgHmP2Reob3agONYznmkcovX-GqUL8mQQ'];

    const driveFolderID = '1fTcNrJ6pDRuzzl4ZTcmclqNme9LKhWgf'

    const ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Oficio_Nao_Competencia_INIS");
    const data = ws.getDataRange().getValues()
    const valorPlanilha = data[1][3];
    let index = templateOptions.indexOf(valorPlanilha)

    const docTemplateIDselect = docTemplateID[index]
    const googleDocTemplate = DriveApp.getFileById(docTemplateIDselect);
    const destinationFolder = DriveApp.getFolderById(driveFolderID);

    const copy = googleDocTemplate.makeCopy(`Oficio ${data[1][0]}`, destinationFolder)
    const doc = DocumentApp.openById(copy.getId())
    const body = doc.getBody();

    body.replaceText('{{RelatOcorrencia}}', data[1][0]);
    body.replaceText('{{SIPE}}', data[1][1]);
    body.replaceText('{{DataHoje}}', data[1][2]);

    doc.saveAndClose();

    const url = doc.getUrl();
    ws.getRange(2, 5).setValue(url);
}