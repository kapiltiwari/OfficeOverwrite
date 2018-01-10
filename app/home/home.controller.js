(function () {
  'use strict';

  angular.module('officeAddin')
    .controller('homeController', ['dataService', homeController]);

  /**
   * Controller constructor
   */
  function homeController(dataService) {
    var vm = this;  // jshint ignore:line
    vm.title = 'home controller';
    vm.dataObject = {};

    vm.changeContent = changeContentFun;

    getDataFromService();

    function getDataFromService() {
      dataService.getData()
        .then(function (response) {
          vm.dataObject = response;
        });
    };

    function changeContentFun() {
      Word.run(function (context) {
        //app.showNotification("start");
        try {
          var thisDocument = context.document;
          thisDocument.body.clear();

          var str = "Hello World - this is a sample to show the issue";
          thisDocument.body.insertText(str, Word.InsertLocation.end);
          context.sync();
        }
        catch (ex) {
          //app.showNotification("error:" + ex.toString());
        }
      })
    };
  }
})();
