// Class Skill
function Skill(name, rate, commercialExperienceRate, overallExperienceRate) {
  
  function checkRateValue (rateValue, parameterName) {
    if (rateValue < 0 || rateValue > 5)
      throw "Bad " + parameterName + " value";
  };
  
  checkRateValue(rate, "rate");
  checkRateValue(commercialExperienceRate, "commercial experience rate");
  checkRateValue(overallExperienceRate, "overall experience rate");
  
  var this_name = name;
  var this_rate = rate;
  var this_commercialExperienceRate = commercialExperienceRate;
  var this_overallExperienceRate = overallExperienceRate;
  
  this.getName = function() {
    return this_name;
  };
  this.setName = function(name) {
    this_name = name;
  };
  this.getName = function() {
    return this_rate;
  };
  this.setRate = function(rate) {
    checkRateValue(rate);
    this_rate = rate;
  };
  this.getRate = function() {
    return this_rate;
  };
  this.setCommercialExperienceRate = function(commercialExperienceRate) {
    checkRateValue(commercialExperienceRate, "commercial experience rate");
    this_commercialExperienceRate = commercialExperienceRate;
  };
  this.getCommercialExperienceRate = function() {
    return this_commercialExperienceRate;
  };
  this.setOverallExperienceRate = function(overallExperienceRate) {
    checkRateValue(overallExperienceRate, "overall experience rate");
    this_overallExperienceRate = overallExperienceRate;
  };
  this.getOverallExperienceRate = function() {
    return this_overallExperienceRate;
  };
  
  this.toString = function() {
    return "Skill " + this_name + ", rate: " + this_rate + ", commercial experience rate: "
    + this_commercialExperienceRate + ", overall experience rate: " + this_overallExperienceRate;
  }
}

// Class SkillCathegory
function SkillCathegory() {
  
}

//// Class SkillTree
//function SkillTree() {
//  
//}

// Classes read utill section

function parseSkillFromSheetLine(sheet, line) {
  var data = sheet.getDataRange().getValues();
  return new Skill(data[line][0] + data[line][1], parseInt(data[line][2]), parseInt(data[line][3]), parseInt(data[line][4]));
}

function main() {
  var directory = DriveApp.getFoldersByName('Google Script examples').next();
  var skillsTableFile = directory.getFilesByName("Skills Tracker").next();
  var skillsTable = SpreadsheetApp.openById(skillsTableFile.getId());
  
  // Test section
  
  var sheets = skillsTable.getSheets();
  for (var i = 0; i < sheets.length; ++i) {
    Logger.log(sheets[i].getName());
  }
  
  Logger.log(parseSkillFromSheetLine(sheets[0], 1).toString());
  
  // TODO write
}
