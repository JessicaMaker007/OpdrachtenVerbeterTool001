 //as XLSX from 'xlsx';
//Deze variabele wordt in 2 functies gebruikt dus = globale variabele
import './styles.css';  // Als je een CSS bestand gebruikt (optioneel)
import * as XLSX from 'xlsx'; // Of een andere import zoals exceljs, afhankelijk van je code
import ExcelJS from 'exceljs';

 let foutAantal = 0; 
 let totaalPunten = 0;

 //Functie om excel-bestanden in te lezen
function readExcel(file, callback) {
    const reader = new FileReader();
    reader.readAsArrayBuffer(file);
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array"});

        let alleGegevens = {};
        let kolomNamen = [];

        if (workbook.SheetNames.length === 0){
            console.error("Fout: Geen sheets gevonden in het bestand.");
            alert("Het Excel-bestand lijkt leeg te zijn. Upload een geldig bestand.");
            return;
        }

        //Haal het aantal sheets op dat de gebruiker wil controleren door gebruiker zelf ingegeven in browser
        const maxSheetsInput = document.getElementById("aantalSheets").value;
        const maxSheets = parseInt(maxSheetsInput, 10) || 1; //Zorg ervoor dat het een geldig getal is, anders 1

        //const maxSheets = 3; //We beperken tot aantal bladen dit gebruiken indien gebruiker aantal sheets niet zelf kan ingeven in browser
        const sheetsToProcess = workbook.SheetNames.slice(0, maxSheets); //Neem elke eerste aantal bladen

        console.log("workbook.SheetNames:", workbook.SheetNames);
        console.log("maxSheets:", maxSheets);
        console.log("sheetNames:", workbook.SheetNames.slice(0, maxSheets));

        //Loop door aantal bladen
        sheetsToProcess.forEach(sheetName => {
            const sheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(sheet, {headers: 1,defval: ""});
            console.log("Volledige ruw data:", jsonData)
            
            if(jsonData.length === 0) { //Controleren of de sheet niet leeg is
                console.warn("Sheet" + sheetName + " is leeg." );
                return;
            }
                
                //Dynamisch de kolommen ophalen van de eerste rij
                kolomNamen = Object.keys(jsonData[0]);
                //Controle of de juiste kolommen opgehaald worden door deze te projecteren in console
                console.log("Opgehaalde kolommen:", kolomNamen);

                //Verwerk de gegevens per rij
                const meerdereKolomData = jsonData.map(row  => {
                    return kolomNamen.reduce((acc, kolom) => {
                        acc[kolom] = row[kolom] || "Leeg"; //Voorkom errors bij lege cellen
                        return acc;
                    }, {});   
                });

            alleGegevens[sheetName] = meerdereKolomData;
    });

            if (Object.keys(alleGegevens).length === 0){
                console.error("Fout: Geen data ingeladen.");
                return;
            }    

            callback(alleGegevens, workbook);
        };
    }

//Bestanden uploaden
document.addEventListener("DOMContentLoaded", function (){
    let userAnswers = null;
    let correctAnswers = null;
   
    //Haal bestanden op
    document.getElementById("uploadAntwoorden").addEventListener("change", function(event){
        readExcel(event.target.files[0], function(data) {
            userAnswers = data;
            console.log("Gebruikersantwoorden geladen:", userAnswers);
        });
    });

    document.getElementById("uploadCorrect").addEventListener("change", function(event){
        readExcel(event.target.files[0], function(data, workbook) {
            correctAnswers = data;
            console.log("Correcte antwoorden geladen:", correctAnswers);

            const maxSheetsInput = document.getElementById("aantalSheets").value;
            const maxSheets = parseInt(maxSheetsInput, 10) || 1;

            // Als maxSheets geen geldig getal is, gebruik fallback en geef waarschuwing
            if (isNaN(maxSheets) || maxSheets <= 0) {
                alert("Voer een geldig aantal sheets in om te controleren.");
                return;
            }

            const nietLegeCellen = countNumericCells(workbook, maxSheets);
            totaalPunten = nietLegeCellen.totalCount;

            console.log("Niet-lege cellen per sheet:", nietLegeCellen.countsPerSheet);
            console.log("Totaal aantal niet-lege cellen:", totaalPunten);

            //document.getElementById("nietLegeCellenOutput").textContent =`Totaal niet-lege cellen (correcte data): ${nietLegeCellen.totalCount}`;
        });
    });

    document.getElementById("vergelijkButton").addEventListener("click", function() {
        if(!userAnswers || Object.keys(userAnswers).length === 0){
            alert("Upload eerst een bestand met gebruikersantwoorden!");
            console.log("userAnswers:", userAnswers); // Debugging log
            return;
        } 
        if (!correctAnswers || Object.keys(correctAnswers).length === 0) {
            alert("Upload eerst een bestand met correcte antwoorden!");
            console.log("correctAnswers:", correctAnswers); // Debugging log
            return;
        }

        vergelijkAntwoorden(userAnswers, correctAnswers);

        //Indien berekenknop gebruikt wordt en de score handmatig ingevuld wordt
        /*document.getElementById("berekenScoreButton").addEventListener("click", function(){
            berekenScore();
    }); */

});
      
      //Functie om getallen af te ronden en als tring met 2 deciamelen weer te geven
      function formatNumber(value){
        if(!isNaN(value) && value !== "") {
            return parseFloat(value).toFixed(2); //Rond af naar 2 decimalen
        }
        return value;//Laat tekstwaarden ongemoeid
    }

      //Download verbeterd Excel-bestand - knop 
      document.getElementById("downloadVerbeterd").addEventListener("click", async function(){
        if (!userAnswers || !correctAnswers) {
            alert("Vergelijk eerst de bestanden!");
            return;
        }

        //Maak een nieuw werkboek
        const workbook = new ExcelJS.Workbook();
        
        //Loop door alle sheets
        for (const sheetName of Object.keys(userAnswers)) {
            if (!correctAnswers[sheetName]) continue;
        
            const userSheet = userAnswers[sheetName];
            const correctSheet = correctAnswers[sheetName];
            const sheet = workbook.addWorksheet(sheetName);

            //Kopregel toevoegen (eerste rij van het bestand)
            const headers = Object.keys(userSheet[0] || {}).concat("Opmerkingen");
            const headerRow = sheet.addRow(headers);

            headerRow.eachCell((cell) => {
            cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
            cell.fill = {
                type: "pattern",
                pattern: "solid",
                fgColor: { argb: "4472C4" } // Blauw
            };
            cell.alignment = { horizontal: "center" };
            cell.border = {
                bottom: { style: "thin" }
        };
    });

            //Loop door de rijen data rijen
            for (let r = 0; r < userSheet.length; r++) {
                let userRow = userSheet[r] || {};
                let correctRow = correctSheet[r] || {};
                let verbeterdeRij = [];
                let opmerkingen = [];
                
                //Loop door de kolommen
                for (const kolom of Object.keys(userRow)) {
                    let userValue = userRow[kolom]?.toString().trim() ?? "Leeg";
                    let correctValue = correctRow[kolom]?.toString().trim() ?? "Leeg";

                    //Controleer of het een getal is en formatteer naar 2 decimalen
                    const formattedUserValue = (!isNaN(userValue) && userValue !== "Leeg") ? parseFloat(userValue).toFixed(2) : userValue;
                    const formattedCorrectValue = (!isNaN(correctValue) && correctValue !== "Leeg") ? parseFloat(correctValue).toFixed(2) : correctValue;

                    if (!isNaN(userValue) && userValue !== "") userValue = parseFloat(userValue).toFixed(2);
                    if (!isNaN(correctValue) && correctValue !== "") correctValue = parseFloat(correctValue).toFixed(2);

                    if (userValue === correctValue) {
                    verbeterdeRij.push(userValue);
                    } else {
                    verbeterdeRij.push(userValue + "\u274C");
                    opmerkingen.push(`Fout in kolom ${kolom}: moet zijn ${correctValue}`);
                    }
                }

                //Voeg opmerkingen toe
                verbeterdeRij.push(opmerkingen.join("; "));
                const dataRow = sheet.addRow(verbeterdeRij);
            
                // Foutieve antwoorden rood kleuren
                dataRow.eachCell((cell) => {
                    if (typeof cell.value === "string" && cell.value.includes("\u274C")) {
                    cell.font = { color: { argb: "FF0000" } }; // Rood
                    }
                });
            }

                // Kolommen automatisch breder
                sheet.columns.forEach(column => {
                    let maxLength = 10;
                    column.eachCell({ includeEmpty: true }, cell => {
                    const len = cell.value ? cell.value.toString().length : 0;
                    if (len > maxLength) maxLength = len;
                    });
                    column.width = maxLength + 2;
                });
            }


            // Download bestand in browser
            const buffer = await workbook.xlsx.writeBuffer();
            const blob = new Blob([buffer], {
                type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            });

            const link = document.createElement("a");
            link.href = URL.createObjectURL(blob);
            link.download = "Verbeterd_Bestand.xlsx";
            link.click();
    });

//Functie om niet lege cellen te tellen
// Vereist SheetJS (xlsx.js)
// npm install xlsx als je een Node-omgeving gebruikt
  
//Functie om antwoorden te vergelijken en resultaat weer te geven
function vergelijkAntwoorden(userData, correctData) { 
    if (!userData || typeof userData !== 'object' || Object.keys(userData).length === 0) {
        console.error("Fout: userData is niet correct geladen.", userData);
        alert("Het bestand is niet correct geladen. Probeer opnieuw.");
        return;
    }

    if (!correctData || typeof correctData !== 'object' || Object.keys(correctData).length === 0) {
        console.error("Fout: correctData is niet correct geladen.", correctData);
        alert("Het bestand is niet correct geladen. Probeer opnieuw.");
        return;
    }

    let resultaatTekst = "";
    let correctAantal = 0;
    foutAantal = 0; //Reset fouten bij nieuwe vergelijking
    
    const sheet1 = Object.keys(userData); //Sheets in bestand 1
    const sheet2 = Object.keys(correctData); //Sheets in bestand 2

    const maxSheets = Math.min(sheet1.length, sheet2.length); //Door alle sheets lopen 

    //Loopt door de sheets
    for(let i = 0; i < maxSheets; i++) {
        const SheetName1 = sheet1[i];
        const SheetName2 = sheet2[i];

        //Controleer of de sheetnamen overeenkomen, anders sla over als de sheetnamen niet overeenkomen
        if (SheetName1 !== SheetName2) continue;

        const data1 = userData[SheetName1];
        const data2 = correctData[SheetName2];

        //Voeg een kopje toe voor elke sheet
        resultaatTekst += "<div class='sheet'>Vergelijking van sheet: " + SheetName1 + "</div>";

        //Controleer aantal rijen
        const maxRows = Math.min(data1.length, data2.length);

        //Vergelijk de rijen per sheet
        for(let r = 0; r < maxRows; r++) {
            const row1 = data1[r] || {};// Zorg dat row1 en row2 geen undefined is
            const row2 = data2[r] || {};

            Object.keys(row1).forEach(kolom => {
                const value1 = row1[kolom] ? row1[kolom].toString().trim() : "Leeg";
                const value2 = row2[kolom] ? row2[kolom].toString().trim() : "Leeg";

                //Vergelijken van de waarden
                if (value1 === value2) {
                    correctAantal++;
                    //resultaatTekst += '<div class="correct">\u{2705} Rij" + (r+1) + " , Kolom (" + kolom + "): Correct (" + value1 + ")\n"</div>';             
                } else {
                    foutAantal++;
                    resultaatTekst += "<div class='fout'> \u274C Rij " + (r+1) + " , Kolom (" + kolom + "): Fout (" + value1 + ") - Correct: (" + value2 + ")</div>";            
                } 
            });

            //Tussenlijn tussen rijen indien correctAantal resultaatTekst ook zichtbaar
            //resultaatTekst += '<hr>';
        }

        //Voeg extra regel toe tussen de sheets
        resultaatTekst += '<hr>';
        
    }
        //Berekenen van de resultaten en weergeven in browser
        const behaaldeScore = Math.max(0, totaalPunten - foutAantal);
        const percentage = totaalPunten > 0 ? ((behaaldeScore / totaalPunten) * 100).toFixed(1) : "0.0";
        const scoreOp20 = totaalPunten > 0 ? ((behaaldeScore / totaalPunten) * 20).toFixed(1) : "0.0";

        const samenvatting = "<div class='Samenvatting'>"
             + "Totaal te behalen punten: " + totaalPunten + "<br>" + "<br>"
             + "Totaal foutieve antwoorden: " + foutAantal + "<br>" + "<br>"
             + "<strong>Behaalde score: " + behaaldeScore + "</strong>" + "<br>" + "<br>"
             + "Percentage: <strong>" + percentage + "%</strong><br>"
             + "Score op 20: <strong>" + scoreOp20 + "/20</strong>"
             + "</div>";

        document.getElementById('samenvatting').innerHTML = samenvatting;

        //Update HTML met de resultaten en de samenvatting
        document.getElementById('resultaat').innerHTML = resultaatTekst;
        document.getElementById('samenvatting').innerHTML = samenvatting;

        //Om Score automatisch te laten berekenen bij invullen getal en druk op knop'vergelijken'
        //berekenScore();
        
        return resultaatTekst;
        
}     

//Functie om cellen met getallen te tellen per sheet en totaat sheets
function countNumericCells(workbook, maxSheets) {
    if (!workbook || !workbook.SheetNames || !Array.isArray(workbook.SheetNames)) {
        console.error("Ongeldig workbook-object:", workbook);
        return { countsPerSheet: {}, totalCount: 0 };
    }

    if (isNaN(maxSheets) || maxSheets <= 0) {
        console.warn("Ongeldig maxSheets-waarde:", maxSheets);
        return { countsPerSheet: {}, totalCount: 0 };
    }

    const sheetNames = workbook.SheetNames.slice(0, maxSheets);
    console.log("Verwerken van sheets:", sheetNames);

    let totalCount = 0;
    const countsPerSheet = {};

    sheetNames.forEach(sheetName => {
        const sheet = workbook.Sheets[sheetName];
        if (!sheet) {
            console.warn("Sheet niet gevonden:", sheetName);
            return;
        }

        let count = 0;

        for (const cellAddress in sheet) {
            if (cellAddress[0] === '!') continue;

            const cell = sheet[cellAddress];
            if (cell && typeof cell.v === 'number'){
            //if (cell && cell.v !== undefined && cell.v !== null && cell.v !== '') {
                count++;
            }
        }

        countsPerSheet[sheetName] = count;
        totalCount += count;
    });

    return {
        countsPerSheet,
        totalCount
    };
}

//Functie om score te berekenen enkel te gebruiken indien knop beschikbaar en totaal hantmatig wordt ingegeven
function berekenScore() {
    const startGetal = parseInt(document.getElementById("startGetal").value, 10);//Haal het ingevoerde getal op
    if (isNaN(startGetal)) {
        alert("Voer een geldig getal in");
        return
    }

    const score = startGetal - foutAantal; //Bereken de score

    //Resultaat weergeven op pagina
    const scoreResultaat = document.getElementById('scoreResultaat');
    scoreResultaat.textContent = "Je score is: " + score + " (Start getal: " + startGetal + " - aantal fouten: " + foutAantal + ")";  
}

{
    //Controle SheetJS ingeladen
    console.log(typeof XLSX !== "undefined" ? "SheetJS is geladen" : "Fout: SheetJS is niet geladen!");
}
})

