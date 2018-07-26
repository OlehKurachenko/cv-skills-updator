function deleteFiles() {
  var files = DriveApp.getFilesByName("Test Skill Card");
  while (files.hasNext()) {
     files.next().setTrashed(true);
  }
}
