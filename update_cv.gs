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
  
  this.sortByRate = function() {
    for (var i = 0; i < this.skills.length; ++i) {
      this.skills[i].subskills.sort(function (v1, v2) {
        return -1 * skillRateComparator(v1, v2);
      });
    }
    
    this.skills.sort(function(v1, v2) {
      return -1 * skillCathegoryRateComparator(v1, v2);
    });
  }
  
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

// Classes comparators

function skillInterestComparator(skill1, skill2) {
  if (skill1.isImportant != skill2.isImportant) {
    return (skill1.isImportant) ? 1 : -1;
  }
  if (skill1.interestRate != skill2.interestRate)
    return (skill1.interestRate < skill2.interestRate) ? -1 : 1;
  if (skill1.rate != skill2.rate)
    return (skill1.rate > skill2.rate) ? -1 : 1;
  if (skill1.overallExperienceRate != skill2.overallExperienceRate)
    return (skill1.overallExperienceRate > skill2.overallExperienceRate) ? -1 : 1;
  if (skill1.commercialExperienceRate != skill2.commercialExperienceRate)
    return (skill1.commercialExperienceRate > skill2.commercialExperienceRate) ? -1 : 1;
  return 0;
}

function skillRateComparator(skill1, skill2) {
  if (skill1.rate != skill2.rate)
    return (skill1.rate < skill2.rate) ? -1 : 1;
  if (skill1.commercialExperienceRate != skill2.commercialExperienceRate)
    return (skill1.commercialExperienceRate < skill2.commercialExperienceRate) ? -1 : 1;
  if (skill1.overallExperienceRate != skill2.overallExperienceRate)
    return (skill1.overallExperienceRate < skill2.overallExperienceRate) ? -1 : 1;
  if (skill1.interestRate != skill2.interestRate)
    return (skill1.interestRate < skill2.interestRate) ? -1 : 1;
  return (0);
}

function skillCathegoryRateComparator(skillCat1, skillCat2) {
  if (skillRateComparator(skillCat1.mainSkill, skillCat2.mainSkill) != 0)
    return skillRateComparator(skillCat1.mainSkill, skillCat2.mainSkill);
  for (var i = 0; i < skillCat1.subskills.length && i < skillCat2.subskills.length; ++i) {
    if (skillRateComparator(skillCat1.subskills[i], skillCat2.subskills[i]) != 0)
      return skillRateComparator(skillCat1.subskills[i], skillCat2.subskills[i]);
  }
  if (skillCat1.subskills.length != skillCat2.subskills.length)
    return (skillCat1.subskills.length < skillCat2.subskills.length) ? -1 : 1;
  return 0;
}

// API

function getSkillTrees() {
  var dataDirName = "Google Script examples";
  var dataFileName = "Skills Tracker";
  
  var directory = DriveApp.getFoldersByName(dataDirName).next();
  var skillsTableFile = directory.getFilesByName(dataFileName).next();
  var skillsTable = SpreadsheetApp.openById(skillsTableFile.getId());
  
  var sheets = skillsTable.getSheets();
  var skillTrees = [];
  
  for (var i = 0; i < sheets.length; ++i) {
    var skillTree = parseSkillTree(sheets[i]);
    skillTrees.push(skillTree);
  }
  
  return skillTrees;
}

function createSkillCard(skillTrees) {
//  function addFormatedSkill(tabulation, body, skill) {
//    body.append
//  }
  
  if (!skillTrees)
    skillTrees = getSkillTrees();
  
  var skillCardDirName = "Google Script examples";
  var skillCardFileName = "Test Skill Card";
  
  var skillCardDir = DriveApp.getFoldersByName(skillCardDirName).next();
  
  if (skillCardDir.getFilesByName(skillCardFileName).hasNext()) {
    skillCardDir.removeFile(skillCardDir.getFilesByName(skillCardFileName).next());
  }
  
  var skillCardDoc = DocumentApp.create(skillCardFileName);
  var skillCardFile = DriveApp.getFileById(skillCardDoc.getId());
  skillCardDir.addFile(skillCardFile);
  DriveApp.getRootFolder().removeFile(skillCardFile);
  
  // headings
  
  var skillCardBody = skillCardDoc.getBody();
  
  var header = skillCardBody.appendParagraph("Skill Card");
  header.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  header.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  
  var subheader = skillCardBody.appendParagraph("Oleh Kurachenko");
  subheader.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  subheader.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  subheader.setLinkUrl("https://drive.google.com/open?id=1aXUDQhL3jnsSBLi49Xcj7qs9nqXm2frz2ZFGdT9_siA"); // link to CV doc
  
  var rateMeasurmentDetails = skillCardBody.appendParagraph("A scale: Used once ➡️ Novice ➡️ Junior ➡️ Middle ➡️ Senior");
  rateMeasurmentDetails.editAsText().setForegroundColor("#999999");
  rateMeasurmentDetails.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  
  skillCardBody.getParagraphs()[0].removeFromParent();
  
  // body
  
  for (var i = 0; i < skillTrees.length; ++i) {
    var skillTree = skillTrees[i];
    
    var sectionHeading = skillCardBody.appendParagraph(skillTree.name);
    sectionHeading.setHeading(DocumentApp.ParagraphHeading.HEADING3);
    
    for (var j = 0; j < skillTree.skills.length; ++j) {
      var skill = skillTree.skills[j];
      var skillP = skillCardBody.appendParagraph("");
      var text;
      
      for (var k = 0; k < 5; ++k) {
        text = skillP.appendText((k < skill.mainSkill.rate) ? "⬛" : "⬜");
        text.setForegroundColor("#666666");
      }
      
      text = skillP.appendText("\t" + skill.mainSkill.name);
      text.setForegroundColor("#000000");
      text.setBold(true);
      text.setFontFamily("Cousine");
      
      for (var k = 0; k < 25 - skill.mainSkill.name.length; ++k) {
        text = skillP.appendText(" ");
      }
    }
  }
  
  // TODO write
}

function main() {
  var skillTrees = getSkillTrees();
  
  createSkillCard(skillTrees);
}
