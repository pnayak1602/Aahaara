$("input").keypress(function(event){
    var x=event.key;
    console.log(x);
    function test() {
        if (!window['ActiveXObject']) {
          log('Error: ActiveX not supported');
          return;
        }
      
        try {
          var
            ExcelApp = new ActiveXObject("Excel.Application"),
            ExcelBook = ExcelApp.Workbooks.Add();
      
          ExcelBook.Application.Visible = true;
          log('Opened Excel');
      
          wait(2, enterData);
        }
        catch (ex) {
          log('An error occured while attempting to open Excel');
          console.log(ex);
        }
      
        function enterData() {
          try {
            ExcelBook.ActiveSheet.Cells(1, 1).Value = 'foo';
            ExcelBook.ActiveSheet.Cells(2, 1).Value = 'bar';
            log('Entered data');
          }
          catch (ex) {
            log('An error occured while attempting to enter data');
            console.log(ex);
          }
      
          wait(2, deleteRow);
        }
      
        function deleteRow () {
          try {
            ExcelBook.ActiveSheet.Rows(x).Delete();
            log('Deleted first row');
          }
          catch (ex) {
            log('An error occured while attempting to delete a row');
            console.log(ex);
          }
      
          wait(2, quitExcel);
        }
      
        function quitExcel () {
          try {
            // Allow excel to quit without prompting user to save.
            ExcelBook.Saved = true;
            ExcelBook.Application.Quit();
            log('Quit excel');
          }
          catch (ex) {
            log('An error occured while attempting to quit excel');
            console.log(ex);
          }
        }
      }
      
      function wait (time, action) {
        setTimeout(action, time * 1000);
      }
      
      function log (message) {
        var
          list = document.getElementById('log'),
          newLog = document.createElement('li');
        newLog.innerHTML = message;
        list.appendChild(newLog);
      }
      if (!window['console'] || !window.console['log']) { console = {log: log}; }
        }
});