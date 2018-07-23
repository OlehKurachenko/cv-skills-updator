// Class Skill
function Skill(name, rate, commercialExperienceRate, overallExperienceRate, interestRate, isImportant) {
  
  this.name = name;
  this.rate = rate;
  this.commercialExperienceRate = commercialExperienceRate;
  this.overallExperienceRate = overallExperienceRate;
  this.interestRate = interestRate;
  this.isImportant = isImportant;
  
  this.toString = function() {
    return "Skill " + this.name + ", rate: " + this.rate + ", commercial experience rate: "
    + this.commercialExperienceRate + ", overall experience rate: " + this.overallExperienceRate
    + ", interest rate: " + this.interestRate + ", important: " + this.isImportant;
  }
}

// Class SkillCathegory
function SkillCathegory(mainSkill) {
  this.mainSkill = mainSkill;
  this.subskills = [];
  
  this.toString = function() {
    var result =  "Skill cathegory " + this.mainSkill.name + "\n";
    result += this.mainSkill.toString() + "\n";
    for (var i = 0; i < this.subskills.length; ++i) {
      result += "  " + this.subskills[i].toString() + "\n";
    }
    return result;
  }
}

//// Class SkillTree
function SkillTree(name) {
  this.name = name;
  this.skills = [];
  
  this.toString = function() {
    var result = "Skill tree " + this.name + "\n";
    for (var i = 0; i < this.skills.length; ++i) {
      result += this.skills[i].toString() + "\n";
    }
    return result;
  }
}

// Classes read utill section

function parseSkill(sheet, line) {
  var data = sheet.getDataRange().getValues();
  return new Skill(
    data[line][0] + data[line][1], // Skill name
    parseInt(data[line][2]), // Skill rate
    parseInt(data[line][3]), // Skill commercial experince rate
    parseInt(data[line][4]), // Skill overall experience rate
    parseInt(data[line][5]), // Skill interest rate
    sheet.getDataRange().getCell(line + 1, 6).getNote().trim().split(/\s+/)[0] === "Important" // Skill is important
  );
}

function parseSkillTree(sheet) {
  var data = sheet.getDataRange().getValues();
  var skillTree = new SkillTree(sheet.getName());
  var skillCathegory = null;
  
  for (var i = 1; i < data.length; ++i) {
    if (data[i][0] != "") {
      if (skillCathegory !== null)
        skillTree.skills.push(skillCathegory);
      skillCathegory = new SkillCathegory(parseSkill(sheet, i));
    } else {
      skillCathegory.subskills.push(parseSkill(sheet, i));
    }
  }
  if (skillCathegory !== null)
    skillTree.skills.push(skillCathegory);
  return skillTree;
}

function main() {
  var directory = DriveApp.getFoldersByName('Google Script examples').next();
  var skillsTableFile = directory.getFilesByName("Skills Tracker").next();
  var skillsTable = SpreadsheetApp.openById(skillsTableFile.getId());
  
  var skillsSheets = skillsTable.getSheets();
  var skillTrees = [];
  
  // Test section
  
  var skillTree = parseSkillTree(skillsSheets[0]);
  
  Logger.log(skillTree.toString());
  
//  var sheets = skillsTable.getSheets();
//  for (var i = 0; i < sheets.length; ++i) {
//    Logger.log(sheets[i].getName());
//  }
  
  //Logger.log(parseSkillFromSheetLine(sheets[0], 1).toString()); // OK!
  
  // TODO write
  
  
  
//  for (var i = 0; i < SkillsSheets.length; ++i) {
//    
//  }
  
}
