var GmailModule = {

  createLabel: function(name) {
    GmailApp.createLabel(name);
    return { status: "created", name: name };
  },

  listLabels: function() {
    return GmailApp.getUserLabels().map(function(l) {
      return l.getName();
    });
  }

};