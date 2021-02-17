const { GoogleSpreadsheet } = require('google-spreadsheet');
const apiCredentials = require('./credentials.json');

const totalClasses = 60;

const editingDoc = async () => {
  const doc = new GoogleSpreadsheet(apiCredentials.fileId);

  await doc.useServiceAccountAuth({
    client_email: apiCredentials.client_email,
    private_key: apiCredentials.private_key.replace(/\\n/g, '\n')
  })
  await doc.loadInfo();

  console.log("You're connected with Google Sheet API!")

  const sheet = doc.sheetsByIndex[0];

  console.log("Loading cells to write the results...")
  await sheet.loadCells('A1:H27');

  try {

    console.log("Calculating students average and situation by pre-imposed test conditions...")
    for (let index = 3; index < 27; index++) {
      let situation = sheet.getCell(index, 6);
      let finalAverage = sheet.getCell(index, 7);
      let missedClasses = sheet.getCell(index, 2).value;

      let grade1 = sheet.getCell(index, 3).value;
      let grade2 = sheet.getCell(index, 4).value;
      let grade3 = sheet.getCell(index, 5).value;
      let average = (grade1 + grade2 + grade3) / 3;

      if (average < 50) {
        situation.value = "Reprovado por Nota";
        finalAverage.value = 0;
      } else if (average >= 50 && average < 70) {
        situation.value = "Exame Final";
        finalAverage.value = Math.abs((calcFinalAverage(average)));
      } else if (average >= 70) {
        situation.value = "Aprovado";
        finalAverage.value = 0;
      }

      const missedPercent = ((100 / totalClasses) * missedClasses);

      if (missedPercent > 25) {
        situation.value = "Reprovado por Falta";
        finalAverage.value = 0;
      }

    }
  } catch (error) {
    console.log(error);
  }
  await sheet.saveUpdatedCells();
  console.log("The final result is done! Check the doc updated: https://docs.google.com/spreadsheets/d/1FxQmS0PaZ0XOgMA-GtIifI6uEWY913kltzj-J42ZDpw/edit#gid=0")
}

const calcFinalAverage = (avg) => {
  return Math.round(10 - avg);
}

editingDoc();