//
//  This script is written to manage CV and skills by creating and
//  updating some documents.
//
//  Author: Oleh Kurachenko aka okurache
//  oleh.kurachenko@gmail.com
//

// Class Skill
var Skill = function (name, rate, commercialExperienceRate, overallExperienceRate, interestRate, isImportant) {
  
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
  
  function addFormatedSkill(tabulation, body, skill, main) {
    var skillP = body.appendParagraph(tabulation);
      
    for (var k = 0; k < 5; ++k) {
      var text = skillP.appendText((k < skill.rate) ? "⬛" : "⬜");
      text.setForegroundColor((main) ? "#666666" : "#999999");
    }
    
    var text = skillP.appendText("     " + skill.name);
    text.setForegroundColor("#000000");
    if (main)
      text.setBold(true);
  }
  
  if (!skillTrees)
    skillTrees = getSkillTrees();
  
  var skillCardDirName = "Google Script examples";
  var skillCardFileName = "Test Skill Card";
  
  var skillCardDir = DriveApp.getFoldersByName(skillCardDirName).next();
  
  if (skillCardDir.getFilesByName(skillCardFileName).hasNext()) {
    skillCardDir.getFilesByName(skillCardFileName).next().setTrashed(true);
  }
  
  var skillCardDoc = DocumentApp.create(skillCardFileName);
  var skillCardFile = DriveApp.getFileById(skillCardDoc.getId());
  skillCardDir.addFile(skillCardFile);
  DriveApp.getRootFolder().removeFile(skillCardFile);
  
  var skillCardBody = skillCardDoc.getBody();
  
  // headings
  
  var header = skillCardBody.appendParagraph("Skill Card");
  header.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  header.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  
  var subheader = skillCardBody.appendParagraph("Oleh Kurachenko");
  subheader.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  subheader.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  subheader.setLinkUrl("https://drive.google.com/open?id=1aXUDQhL3jnsSBLi49Xcj7qs9nqXm2frz2ZFGdT9_siA"); // link to CV doc
  
  var rateMeasurmentDetails = skillCardBody.appendParagraph("Scale: Used once | Novice |️ Junior | Middle |️ Senior");
  rateMeasurmentDetails.editAsText().setForegroundColor("#999999");
  rateMeasurmentDetails.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  
  skillCardBody.getParagraphs()[0].removeFromParent();
  
  // body
  
  for (var i = 0; i < skillTrees.length; ++i) {
    var skillTree = skillTrees[i];
    skillTree.sortByRate();
    
    var sectionHeading = skillCardBody.appendParagraph(skillTree.name);
    sectionHeading.setHeading(DocumentApp.ParagraphHeading.HEADING3);
    
    for (var j = 0; j < skillTree.skills.length; ++j) {
      var skill = skillTree.skills[j];
      if (skill.mainSkill.rate == 0)
        continue;
      addFormatedSkill("", skillCardBody, skill.mainSkill, true);
      for (var l = 0; l < skill.subskills.length; ++l) {
        if (skill.subskills[l].rate == 0)
          break;
        addFormatedSkill("    ", skillCardBody, skill.subskills[l], false);
      }
    }
  }
  
  // footer
  
  skillCardBody.appendHorizontalRule();
  var footer = skillCardBody.appendParagraph("Generated by ");
  var text = footer.appendText("a script");
  text.setLinkUrl("https://github.com/OlehKurachenko/cv-skills-updator/tree/test");
  var text = footer.appendText(" at " + Utilities.formatDate(new Date(), "GMT+3", "EEEEEEEEEE, MMMMMMMMMM dd yyyy, hh:mm:ss"));
  text.setLinkUrl("");
  
  footer.editAsText().setForegroundColor("#999999");
  footer.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
}

function createSkillImprovementTaskCard(skillTrees) {
  if (!skillTrees)
    skillTrees = getSkillTrees();
  
  var skillCardDirName = "Google Script examples";
  var skillCardFileName = "Test Skill Improvements Task Card " + Utilities.formatDate(new Date(), "GMT+3", "MMM dd yyyy");
  
  var skillCardDir = DriveApp.getFoldersByName(skillCardDirName).next();
  
  if (skillCardDir.getFilesByName(skillCardFileName).hasNext()) {
    skillCardDir.getFilesByName(skillCardFileName).next().setTrashed(true);
  }
  
  var skillCardDoc = DocumentApp.create(skillCardFileName);
  var skillCardFile = DriveApp.getFileById(skillCardDoc.getId());
  skillCardDir.addFile(skillCardFile);
  DriveApp.getRootFolder().removeFile(skillCardFile);
  
  var skillCardBody = skillCardDoc.getBody();
  
  // headings
  
  var header = skillCardBody.appendParagraph("Skill Improvement Task Card");
  header.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  header.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  
  var date = new Date();
  var lastDate = new Date();
  var days = 1;
  
  while (Utilities.formatDate(lastDate, "GMT+3", "EEE") != "Fri") {
    ++days;
    lastDate.setDate(lastDate.getDate() + 1);
  }
  
  Logger.log("days=" + days); 
  
  var dateSubheader = skillCardBody.appendParagraph(
    Utilities.formatDate(date, "GMT+3", "MMM dd yyyy") + " - "
    + Utilities.formatDate(lastDate, "GMT+3", "MMM dd yyyy")
  );
  dateSubheader.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  dateSubheader.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  dateSubheader.setSpacingBefore(0);
  dateSubheader.setSpacingAfter(0);
  
  var subheader = skillCardBody.appendParagraph("Oleh Kurachenko");
  subheader.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  subheader.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  subheader.setLinkUrl("https://drive.google.com/open?id=1aXUDQhL3jnsSBLi49Xcj7qs9nqXm2frz2ZFGdT9_siA"); // link to CV doc
  subheader.setSpacingBefore(0);
  
  var rateMeasurmentDetails = skillCardBody.appendParagraph(
    "Scale: Used once | Novice |️ Junior | Middle |️ Senior"
  );
  rateMeasurmentDetails.editAsText().setForegroundColor("#999999");
  rateMeasurmentDetails.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  
  skillCardBody.getParagraphs()[0].removeFromParent();
  
  // body
  
  skills = [];
  for (var i = 0; i < skillTrees.length; ++i) {
    var skillTree = skillTrees[i];
    for (var j = 0; j < skillTree.skills.length; ++j) {
      var skill = skillTree.skills[j];
      skills.push(new Skill(
        skill.mainSkill.name, skill.mainSkill.rate, skill.mainSkill.commercialExperienceRate,
        skill.mainSkill.overallExperienceRate, skill.mainSkill.interestRate, skill.mainSkill.isImportant
      ));
      Logger.log("Len: " + skill.subskills.length);
      for (var l = 0; l < skill.subskills.length; ++l) {
        Logger.log("Here!");
        var subskill = skill.subskills[l];
        skills.push(new Skill(skill.mainSkill.name + ": " + subskill.name,
                              subskill.rate,
                              subskill.commercialExperienceRate,
                              subskill.overallExperienceRate,
                              subskill.interestRate,
                              subskill.isImportant));
      }
    }
  }
  
  skills.sort(function(v1, v2) {
    return -1 * skillInterestComparator(v1, v2);
  });
  
  for (var i = 0; i < days; ++i) {
    var skill = skills[i];
    
    var skillP = skillCardBody.appendParagraph("");
    skillP.setSpacingBefore(4);
    
    for (var k = 0; k < 5; ++k) {
      var text = skillP.appendText((k < skill.rate) ? "⬛" : "⬜");
      text.setForegroundColor("#666666");
    }
    
    skillP.appendText("  ");
    var text = skillP.appendText(skill.name);
    text.setForegroundColor("#000000");
    text.setBold(true);
    
    for (var k = 0; k < 44; ++k) { // 44 is line length
      if (k > skill.name.length)
        skillP.appendText(" ");
    }
    
    for (var k = 0; k < 5; ++k) {
      var text = skillP.appendText("▶");
      if (k < skill.commercialExperienceRate) {
        text.setBold(true);
        text.setForegroundColor("#000000");
      } else if (k < skill.overallExperienceRate) {
        text.setBold(true);
        text.setForegroundColor("#666666");
      } else {
        text.setBold(true);
        text.setForegroundColor("#b7b7b7");
      }
    }
    var text = skillP.appendText("  ");
    
    for (var k = 0; k < 5; ++k) {
      var text = skillP.appendText((k < skill.interestRate) ? "★" : "☆");
      text.setBold(false);
      text.setForegroundColor("#000000");
    }
    if (skill.isImportant) {
      var text = skillP.appendText("!");
      text.setBold(true);
    }
    
    skillP.editAsText().setFontFamily("Cousine");
    
    var rawTable = [
      ["Main goal", ""],
      ["Resources", ""],
      ["Tasks", ""],
      ["Result", ""]
    ];
    
    var table = skillCardBody.appendTable(rawTable);
    table.editAsText().setFontFamily("Arial");
    table.setColumnWidth(0, 80);
    for (var d1 = 0; d1 < rawTable.length; ++d1) {
      for (var d2 = 0; d2 < rawTable[0].length; ++d2) {
        var cell = table.getCell(d1, d2);
        cell.setPaddingTop(0);
        cell.setPaddingBottom(0);
        cell.editAsText().setForegroundColor("#000000");
      }
    }
  }
  
  // footer
  
  // TODO fix code duplication, but don't forget about font
  skillCardBody.appendHorizontalRule();
  var footer = skillCardBody.appendParagraph("Generated by ");
  var text = footer.appendText("a script");
  text.setLinkUrl("https://github.com/OlehKurachenko/cv-skills-updator/tree/test");
  var text = footer.appendText(" at " + Utilities.formatDate(new Date(), "GMT+3", "EEEEEEEEEE, MMMMMMMMMM dd yyyy, hh:mm:ss aa"));
  text.setLinkUrl("");
  footer.editAsText().setFontFamily("Arial");
  
  footer.editAsText().setForegroundColor("#999999");
  footer.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  
  skillCardDoc.saveAndClose();
  GmailApp.sendEmail("oleh.kurachenko@gmail.com", "Skills improvement task card card updated", "", 
    {
      name: "CV Updator script",
      htmlBody: HtmlService.createHtmlOutputFromFile('skill_improvement_task_card_update_email').getContent(),
      attachments: [skillCardFile.getAs(MimeType.PDF)]
    });
}

function main() {
  var skillTrees = getSkillTrees();
  
  createSkillCard(skillTrees);
}
