
//Function for generating random numbers
function generateRandomNumbers(min, max) {
    return (Math.random() * (max - min) + min).toFixed(2);
}

// Generates weekly data 
function generateWeeklyData() {
  for(let i = 0; i < 7; i++) {
    generateData()
  }
}

// Generates random 24 hour data
function generateData() {

  let spreadSheet = SpreadsheetApp.getActiveSpreadsheet();

  let newData = spreadSheet.getSheets()[0];
  let averageSheet = spreadSheet.getSheets()[1];
  let chartSheet = spreadSheet.getSheets()[2];

  spreadSheet.getRange('I4').setValue(SpreadsheetApp.getActiveSheet().getRange('I4').getValue() + 1);

  // Number of days 
  let counter = spreadSheet.getRange('I4').getValue()
  
  spreadSheet.setActiveSheet(newData)


  let dateObj = new Date()
  dateObj.setDate(dateObj.getDate() + counter)

  spreadSheet.appendRow([`Day ${dateObj.toLocaleString().split(',')[0]}`
  ,generateRandomNumbers(90,120)
  ,generateRandomNumbers(20,50)
  ,generateRandomNumbers(10,40)
  ,generateRandomNumbers(5,30)
  ,generateRandomNumbers(50,80)
  ,generateRandomNumbers(130,160)
  ,generateRandomNumbers(115,140)
  ])

  Logger.log(counter)
}

// Deletes all the 24 hour data that was generated
function deleteAllDataGenerated() {
  
  let spreadSheet = SpreadsheetApp.getActiveSpreadsheet();

  let newData = spreadSheet.getSheets()[0];

  spreadSheet.setActiveSheet(newData)

  let numberofData = spreadSheet.getRange('I4').getValue()

  spreadSheet.deleteRows(11, spreadSheet.getLastRow() - 11 + 1)

  spreadSheet.getRange('I4').setValue(0)
}

// Compute weekly and monthly average and put data in 'Averages Sheet'
function computeAverage() {

  //Initialize spreadsheet and set 'New 24 Hour Data as current sheet'
  let spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  let newData = spreadSheet.getSheets()[0];
  let averageSheet = spreadSheet.getSheets()[1];
  let chartSheet = spreadSheet.getSheets()[2];
  spreadSheet.setActiveSheet(newData)  

  let lastRowNumber = spreadSheet.getLastRow()

  // Compute and set recent weekly average prices
  let weeklyData = []
  let weeklyAverage = []

  let alphabet = ['B','C','D','E','F','G','H']


  // Get the values from the 24 hour data sheet
  for (let i = 0; i < alphabet.length; i++) {
    for(let j = lastRowNumber - 6; j < lastRowNumber + 1; j++) {
      weeklyData.push(spreadSheet.getRange(`${alphabet[i]}${j}`).getValues())
    }
  }
  
  
  let chairave = 0
  let kettleave = 0
  let usbave = 0
  let keyave = 0
  let walletave = 0
  let bagave = 0
  let bedave = 0

  for (let i = 0; i < weeklyData.length; i++) {
    if (i <=6) {
      chairave = chairave + parseFloat(weeklyData[i][0])
    }
    else if (i <= 13) {
      kettleave = kettleave + parseFloat(weeklyData[i][0])
    }
    else if (i <= 20) {
      usbave = usbave + parseFloat(weeklyData[i][0])
    }
    else if (i <= 27) {
      keyave = keyave + parseFloat(weeklyData[i][0])
    }
    else if (i <= 34) {
      walletave = walletave + parseFloat(weeklyData[i][0])
    }
    else if (i <= 41) {
      bagave = bagave + parseFloat(weeklyData[i][0])
    }
    else if (i <= 48) {
      bedave = bedave + parseFloat(weeklyData[i][0])
    }    
  }

  weeklyAverage[0] = (chairave / 7).toFixed(2)
  weeklyAverage[1] = (kettleave / 7).toFixed(2)
  weeklyAverage[2] = (usbave / 7).toFixed(2)
  weeklyAverage[3] = (keyave / 7).toFixed(2)
  weeklyAverage[4] = (walletave / 7).toFixed(2)
  weeklyAverage[5] = (bagave / 7).toFixed(2)
  weeklyAverage[6] = (bedave / 7).toFixed(2)

  spreadSheet.setActiveSheet(averageSheet)

  //Append to weekly average prices table
  for (let i = 0; i < 7; i++) {
      spreadSheet.getRange(`B${i + 2}`).setValue(weeklyAverage[i])
  }


  // Compute and set montly average prices
  let monthlyData = []
  let monthlyAverage = []

  spreadSheet.setActiveSheet(newData)

  // Get the values from the 24 hour data sheet

  if (spreadSheet.getRange('I4').getValue() > 31) {
    for (let i = 0; i < alphabet.length; i++) {
      for(let j = lastRowNumber - 29; j < lastRowNumber + 1; j++) {
        monthlyData.push(spreadSheet.getRange(`${alphabet[i]}${j}`).getValues())
      }
    }
    let chairave2 = 0
    let kettleave2 = 0
    let usbave2= 0
    let keyave2 = 0
    let walletave2 = 0
    let bagave2 = 0
    let bedave2 = 0

    for (let i = 0; i < monthlyData.length; i++) {
      if (i <=29) {
        chairave2 = chairave2 + parseFloat(monthlyData[i][0])
      }
      else if (i <= 59) {
        kettleave2 = kettleave2 + parseFloat(monthlyData[i][0])
      }
      else if (i <= 89) {
        usbave2 = usbave2 + parseFloat(monthlyData[i][0])
      }
      else if (i <= 119) {
        keyave2 = keyave2 + parseFloat(monthlyData[i][0])
      }
      else if (i <= 149) {
        walletave2 = walletave2 + parseFloat(monthlyData[i][0])
      }
      else if (i <= 179) {
        bagave2 = bagave2 + parseFloat(monthlyData[i][0])
      }
      else if (i <= 209) {
        bedave2 = bedave2 + parseFloat(monthlyData[i][0])
      }    
    }

    Logger.log(monthlyData.length)

    monthlyAverage[0] = (chairave2 / 30).toFixed(2)
    monthlyAverage[1] = (kettleave2 / 30).toFixed(2)
    monthlyAverage[2] = (usbave2 / 30).toFixed(2)
    monthlyAverage[3] = (keyave2 / 30).toFixed(2)
    monthlyAverage[4] = (walletave2 / 30).toFixed(2)
    monthlyAverage[5] = (bagave2 / 30).toFixed(2)
    monthlyAverage[6] = (bedave2 / 30).toFixed(2)

    spreadSheet.setActiveSheet(averageSheet)

    //Append to weekly average prices table
    for (let i = 0; i < 7; i++) {
        spreadSheet.getRange(`C${i + 2}`).setValue(monthlyAverage[i])
    }
  }
  else {
    spreadSheet.setActiveSheet(averageSheet)

    for (let i = 0; i < 7; i++) {
      spreadSheet.getRange(`C${i + 2}`).setValue('Insufficient Data must be more than 31 days')
    }
  }
}

// This function will modify the charts to become up to date to the 24 hour data generated.
function modifyWeeklyChart() {

  let sheet = SpreadsheetApp.getActiveSheet();

  let numberofDays = sheet.getRange("'New 24 Hour Data'!I4").getValue()

  let weeklyRange = sheet.getRange(`'New 24 Hour Data'!A${numberofDays + 10 - 6}:H${numberofDays + 10}`)
  let weeklyChart = sheet.getCharts()[0]

  let ranges = weeklyChart.getRanges()
  weeklyChart = weeklyChart.modify();
  ranges.forEach(function(weeklyRange) {weeklyChart.removeRange(weeklyRange)});
  let modifiedChart = weeklyChart.addRange(weeklyRange).build();
  sheet.updateChart(modifiedChart);


  /*
  if (numberofDays > 31) {
    var montlyRange = sheet.getRange(`'New 24 Hour Data'!A${numberofDays + 10 - 30}:H${numberofDays + 10}`)
    var monthlyChart = sheet.getCharts()[1]

    monthlyChart = monthlyChart.modify()
        .addRange(montlyRange)
        .build()
    sheet.updateChart(monthlyChart);
  }
  else {}

  */

}

// This function will modify the charts to become up to date to the 24 hour data generated.
function modifyMonthlyChart() {

  let sheet = SpreadsheetApp.getActiveSheet();

  let numberofDays = sheet.getRange("'New 24 Hour Data'!I4").getValue()

  if (numberofDays > 31) {
    var montlyRange = sheet.getRange(`'New 24 Hour Data'!A${numberofDays + 10 - 29}:H${numberofDays + 10}`)
    var monthlyChart = sheet.getCharts()[1]

    monthlyChart = monthlyChart.modify()
        .addRange(montlyRange)
        .build()
    sheet.updateChart(monthlyChart);
  }
  else {}

}

