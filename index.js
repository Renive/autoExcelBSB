console.log("Uzupełniam arkusz pracy na dziś...");
const svnUltimate = require('node-svn-ultimate');
const Excel = require('exceljs');
const monthNames = ["Styczeń", "Luty", "Marzec", "Kwiecień", "Maj", "Czerwiec",
    "Lipiec", "Sierpień", "Wrzesień", "Październik", "Listopad", "Grudzień"
];
const svnAddress = process.env.npm_package_config_svnAddress;
const dzisiejszaData = new Date();
const fileName = monthNames[dzisiejszaData.getMonth()] + "_.xlsm";
console.log("Aktualny miesiąc to " + monthNames[dzisiejszaData.getMonth()] + ", używam zatem pliku " + fileName);
const configObjectForSvn = {	// optional options object - can be passed to any command not just update
    trustServerCert: true,	// same as --trust-server-cert
    //username: process.env.npm_package_config_svnLogin,	// same as --username
    //password: process.env.npm_package_config_svnPassword,	// same as --password
    //shell: "sh", 			// override shell used to execute command
    //cwd: process.cwd(),		// override working directory command is executed
    //quiet: true,			// provide --quiet to commands that accept it
    //force: true,			// provide --force to commands that accept it
    //revision: 33050,		// provide --revision to commands that accept it
    //depth: "empty",			// provide --depth to commands that accept it
    //ignoreExternals: true,	// provide --ignore-externals to commands that accept it
    params: ['--limit 50', `--search "${process.env.npm_package_config_svnLogin}"`] // extra parameters to pass
};
// svnUltimate.commands.log('https://hvbysvn00.bsb.com.pl/svn/aml_spert/trunk/src', configObjectForSvn, (test, error, test1) => {
//     //return null;
//     //console.log(test);
//     setTimeout(() => {
//        // console.log(error);
//        console.log(`Miejsce składowania SVN: ${svnAddress+"?r="+error.logentry[0].$.revision}`);
//        // console.warn(error.logentry[0].$.revision);
//        console.log(`Szczegółowy opis: ${error.logentry[0].msg}`);
//        console.log("Twój ostatni wyszukany commit, w celach sprawdznia poprawności");
//         console.log(error.logentry[0]);
//     }, 3000);
//     //console.log(test1);
// });
const backupPath = process.env.npm_package_config_sciezkaDoFolderuBackup;
const fileSystem = require('fs');
const backupFileName=backupPath+"\\"+monthNames[dzisiejszaData.getMonth()]+"_"+new Date().toLocaleDateString()+".xlsm";
fileSystem.copyFileSync(process.env.npm_package_config_adresDoPlikuExcel+"\\"+fileName,backupFileName);
console.log("Backup zrobiony do folderu - stworzyłem plik "+backupFileName);
//console.log(process.cwd());
let workbook = new Excel.Workbook();
workbook.xlsx.readFile(process.env.npm_package_config_adresDoPlikuExcel+"\\"+fileName)
    .then(function () {
        // use workbook
        let tmp = workbook.getWorksheet(1); //pierwszy, czyli "Rozliczenie"
        console.log(tmp);
        //console.log(workbook);
        // console.log(tmp);
        var nameCol = tmp.getColumn('E');
        //console.log(nameCol);
        var imieNazwisko = tmp.getCell('D11');
        console.log(imieNazwisko);
        // imieNazwisko.value = "Piotr Osiński";
        // var cell = tmp.getCell('E11');
        // var worksheet = workbook.getWorksheet(1);
        // var row = worksheet.getRow(5);
        // row.getCell(3).value = "OSIŃSKI PIOTR"; // C5's value set to 5
        // row.commit();
        // // imieNazwisko.commit();
        // // return workbook.xlsx.writeFile('new.xlsx');
        // console.log(cell.value);
        // nameCol.eachCell(function(cell, rowNumber) {
        //     console.log(cell);
        //     console.log(rowNumber);
        // });
    }
    )
    .then(function () {
        //  return workbook.xlsx.writeFile(process.env.npm_package_config_adresDoPlikuExcel) //zapisz zmiany
    });
    // .then(function () {
    //     console.log("Nowy wiersz dodany, wpisałem:");
    //     console.log(`Dzień miesiąca : ${dzisiejszaData.getDate()}`);
    //     console.log(`Typ prac: ${process.env.npm_package_config_typPrac}`);
    //     console.log(`Nazwa projektu: ${process.env.npm_package_config_nazwaProjektu}`);
    //     console.log(`Godziny poświecone na realizację prac: ${process.env.npm_package_config_godzinyPoswiecone}`);
    //     console.log(`Rodzaj pracy: ${process.env.npm_package_config_rodzajPracy}`);
    //     console.log(`Nazwa utworów: ${process.env.npm_package_config_nazwaUtworu}`);
    // });