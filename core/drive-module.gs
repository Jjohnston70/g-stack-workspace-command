var DriveModule = {

  createFolder: function(name) {
    return DriveApp.createFolder(name).getId();
  },

  listFiles: function() {
    var files = DriveApp.getFiles();
    var results = [];

    while (files.hasNext()) {
      results.push(files.next().getName());
    }

    return results;
  }

};