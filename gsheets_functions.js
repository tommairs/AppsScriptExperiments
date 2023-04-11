// Experiment one - GSheets

function prepopulate() {
  var d = new Date();
  var currentTime = d.toLocaleTimeString(); 
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
// Reset initial planned spend values
  const startValues = sheet.getRange(13,2,7).getValues();
  sheet.getRange(13,9,7).setValues(startValues);
// Reset % of spend values
  const startPercentages = sheet.getRange(13,14,7).getValues();
  sheet.getRange(13,15,7).setValues(startPercentages);

    // Clean up background
  sheet.getRange('B13:O19').setBackground("white");

  sheet.getRange('N3').setValue("Updated @ " + currentTime);
 
}  

function recalculate() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
 // var range = sheet.getDataRange();
  const exp_Percentages = sheet.getRange(13,15,7).getValues();
  const exp_Values = sheet.getRange(13,9,7).getValues();
  const conversionRates = sheet.getRange(13,5,7).getValues();
  const costPerS = sheet.getRange(13,7,7).getValues();
  const rateValidator = sheet.getRange(20,15,1).getValues();
  const totalSpend = sheet.getRange(2,1).getValues();

  // Clean up background
  sheet.getRange('B13:O19').setBackground("white");

// Find best conversion rate
  var maxConversion = Math.max(...conversionRates);
  for (r=0; r<7; r++) {
    if (conversionRates[r] == maxConversion) {
      sheet.getRange((13+r),5).setBackground("yellow");
    }
  }

// Find lowest S cost
  var minCostPerS = Math.min(...costPerS);
  for (r=0; r<7; r++) {
    if (costPerS[r] == minCostPerS) {
      sheet.getRange((13+r),7).setBackground("lightblue");
    }
  }

// Find largest exp_Percentage
  var maxPercent = Math.max(...exp_Percentages);
   for (r=0; r<7; r++) {
    if (exp_Percentages[r] == maxPercent) {
      sheet.getRange((13+r),15).setBackground("pink");
    }
  }

// Fix rate distribution if needed

// if the total is more than 100%
    if (rateValidator > 1) {
      var overage = rateValidator - 1;
      var overageRemaining = overage;
      var adjustment = overage/7;
      var accumulatedAdjustments = 0;

      for (r=1; r<8; r++) {
        r1=r+12;
        let expRate = sheet.getRange(r1,15).getValue();
        expRate1 = expRate - adjustment;
        overageRemaining = overageRemaining - adjustment;
        accumulatedAdjustments = accumulatedAdjustments + adjustment;
        if (expRate1 < 0) {
          overageRemaining = overageRemaining + adjustment;
          adjustment = overageRemaining/(7 - r); 
         }
        if (expRate1 >= 0) {
          sheet.getRange(r1,15).setValue(expRate1);
        }
      }
    }

// if the total is less than 100%
     if (rateValidator < 1) {
      var overage = 1 - rateValidator;
      var overageRemaining = overage;
      var adjustment = overage/7;
      var accumulatedAdjustments = 0;

      for (r=1; r<8; r++) {
        r1=r+12;
        let expRate = sheet.getRange(r1,15).getValue();
        expRate1 = expRate + adjustment;
        overageRemaining = overageRemaining - adjustment;
        accumulatedAdjustments = accumulatedAdjustments + adjustment;
      }
    }


  // Now extend the experiment to the spend
  
  Logger.log(exp_Values);
  
  var matrix = new Array();
  for (r=0; r<7; r++) {
    exp_Values[r]=[(exp_Percentages[r] * totalSpend)];
  }

    Logger.log(exp_Values);
    Logger.log(totalSpend);
    Logger.log(exp_Percentages);
    sheet.getRange(13,9,7).setValues(exp_Values);
}

function exporter(){
  var d = new Date();
  var currentTime = d.toLocaleTimeString(); 
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var range = sheet.getDataRange();
  var data = range.getValues();
  var name = "Export - " + currentTime + "";
  ss.insertSheet(name, {template: sheet});
  
}



