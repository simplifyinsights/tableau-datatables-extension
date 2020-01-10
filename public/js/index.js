'use strict';

(function () {

  // Creates a global table reference for future use.
  let tableReference = null;

  // These variables will hold a reference to the unregister Event Listener functions.
  // https://tableau.github.io/extensions-api/docs/interfaces/dashboard.html#addeventlistener
  let unregisterSettingsEventListener = null;
  let unregisterFilterEventListener = null;
  let unregisterMarkSelectionEventListener = null;
  let unregisterParameterEventListener = null;

  $(document).ready(function () {
    // Add the configure option in the initialiseAsync to call the configure function
    // when we invoke this action on the user interface.
    tableau.extensions.initializeAsync({ 'configure': configure }).then(function () {
      // calls a function to show the table. There will be plenty of logic in this one.
      renderDataTable();

      // We add our Settings and Parameter listeners here  listener here.
      unregisterSettingsEventListener = tableau.extensions.settings.addEventListener(tableau.TableauEventType.SettingsChanged, (settingsEvent) => {
        renderDataTable();
      });
      tableau.extensions.dashboardContent.dashboard.getParametersAsync().then(function (parameters) {
        parameters.forEach(function (p) {
          p.addEventListener(tableau.TableauEventType.ParameterChanged, (filterEvent) => {
            renderDataTable();
          });
        });
      });

    }, function () { console.log('Error while Initializing: ' + err.toString()); });
  });

  // Here is where the meat of the Extension is.
  // In a nut shell, we will try to read values from Settings and have several
  // if statements to retrieve values and populate as appropriate. This will end
  // with a call to a datatable function.
  function renderDataTable() {

    const worksheets = tableau.extensions.dashboardContent.dashboard.worksheets;

    // Unregister Event Listeners for old Worksheet, if exists.
    if (unregisterFilterEventListener != null) {
      unregisterFilterEventListener();
    }
    if (unregisterMarkSelectionEventListener != null) {
      unregisterMarkSelectionEventListener();
    }

    // We will try to read the worksheet from the settings, if this exists we will show
    // the configuration screen, otherwise we will clear the table and destroy the
    // reference.
    var sheetName = tableau.extensions.settings.get("worksheet");
    if (sheetName == undefined || sheetName =="" || sheetName == null) {
      $("#configure").show();
      $("#datatable").text("");
      if( tableReference !== null) {
        tableReference.destroy();
      }
      // Exit the function if no worksheet name is present !!!
      return;
    } else {
      // If a worksheet is selected, then we hide the configuration screen.
      $("#configure").hide();
    }

    // Use the worksheet name saved in the Settings to find and return
    // the worksheet object.
    var worksheet = worksheets.find(function (sheet) {
      return sheet.name === sheetName;
    });

    // Retrieve values the other two values from the settings dialogue window.
    var underlying = tableau.extensions.settings.get("underlying");
    var max_no_records = tableau.extensions.settings.get("max_no_records");
    var includeTableName = (tableau.extensions.settings.get('include-table-name') == 'Y' ? true : false);

    // override default datatable lang variables
    var datatableLangObj = {
      oAria: {
        sSortAscending: ': activate to sort column ascending'+(includeTableName ? ' on '+sheetName+' table' : ''),
        sSortDescending: ': activate to sort column descending'+(includeTableName ? ' on '+sheetName+' table' : '')
      }
    };

    // Add an event listener to the worksheet.
    unregisterFilterEventListener = worksheet.addEventListener(tableau.TableauEventType.FilterChanged, (filterEvent) => {
      renderDataTable();
    });
    unregisterMarkSelectionEventListener = worksheet.addEventListener(tableau.TableauEventType.MarkSelectionChanged, (markSelectionEvent) => {
      renderDataTable();
    });

    // If underlying is 1 then get Underlying, else get Summary.
    if (underlying == 1) {
      worksheet.getUnderlyingDataAsync({ maxRows: max_no_records }).then(function (underlying) {
        // We will loop through our column names from our settings and save these into an array
        // We will use this later in our datatable function.
        // https://tableau.github.io/extensions-api/docs/interfaces/datatable.html#columns
        var data = [];
        var column_names = tableau.extensions.settings.get("column_names").split("|");
        for (i = 0; i < column_names.length; i++) {
          data.push({ title: column_names[i] });
        }

        // We have created an array to match the underlying data source and then
        // looped through to populate our array with the value data set. We also added
        // logic to read from the column names and column order from our configiration.
        const worksheetData = underlying.data;
        var column_order = tableau.extensions.settings.get("column_order").split("|");
        var tableData = makeArray(underlying.columns.length,underlying.totalRowCount);
        for (var i = 0; i < tableData.length; i++) {
          for (var j = 0; j < tableData[i].length; j++) {
            // you can get teh value or formatted value
            // https://tableau.github.io/extensions-api/docs/interfaces/datavalue.html
            tableData[i][j] = worksheetData[i][column_order[j]-1].formattedValue;
          }
        }

        // Destroy the old table.
        if (tableReference !== null) {
          tableReference.destroy();
          $("#datatable").text("");
        }

        // Read the Settings and get the single string for UI settings.
        var tableClass = tableau.extensions.settings.get("table-classes");
        $("#datatable").attr('class', '')
        $("#datatable").addClass(tableClass);

        // Read the Settings and create an array for the Buttons.
        var buttons = [];
        var clipboard = tableau.extensions.settings.get("export-clipboard");
        if (clipboard == "Y") {
          buttons.push('copy');
        }
        var csv = tableau.extensions.settings.get("export-csv");
        if (csv == "Y") {
          buttons.push('csv');
        }
        var excel = tableau.extensions.settings.get("export-excel");
        if (excel == "Y") {
          buttons.push('excel');
        }
        var pdf = tableau.extensions.settings.get("export-pdf");
        if (pdf == "Y") {
          buttons.push('pdf');
        }
        var print = tableau.extensions.settings.get("export-print");
        if (print == "Y") {
          buttons.push('print');
        }

        // If there are 1 or more Export options ticked, then we will add the dom: 'Bfrtip'
        // Else leave this out.
        if (buttons.length > 0) {
          tableReference = $('#datatable').DataTable({
            dom: 'Bfrtip',
            data: tableData,
            columns: data,
            responsive: true,
            buttons: buttons,
            bAutoWidth: false,
            initComplete: datatableInitCallback,
            drawCallback: datatableDrawCallback,
            oLanguage: datatableLangObj
          });
        } else {
          tableReference = $('#datatable').DataTable({
            data: tableData,
            columns: data,
            responsive: true,
            bAutoWidth: false,
            initComplete: datatableInitCallback,
            drawCallback: datatableDrawCallback,
            oLanguage: datatableLangObj
          });
        }
      })
    } else {
      worksheet.getSummaryDataAsync({ maxRows: max_no_records }).then(function (sumdata) {
        // We will loop through our column names from our settings and save these into an array
        // We will use this later in our datatable function.
        // https://tableau.github.io/extensions-api/docs/interfaces/datatable.html#columns
        var data = [];
        var column_names = tableau.extensions.settings.get("column_names").split("|");
        for (i = 0; i < column_names.length; i++) {
          data.push({ title: column_names[i] });
        }

        // We have created an array to match the underlying data source and then
        // looped through to populate our array with the value data set. We also added
        // logic to read from the column names and column order from our configiration.
        const worksheetData = sumdata.data;
        var column_order = tableau.extensions.settings.get("column_order").split("|");
        var tableData = makeArray(sumdata.columns.length,sumdata.totalRowCount);
        for (var i = 0; i < tableData.length; i++) {
          for (var j = 0; j < tableData[i].length; j++) {
            tableData[i][j] = worksheetData[i][column_order[j]-1].formattedValue;
          }
        }

        // Destroy the old table.
        if (tableReference !== null) {
          tableReference.destroy();
          $("#datatable").text("");
        }

        // Read the Settings and get the single string for UI settings.
        var tableClass = tableau.extensions.settings.get("table-classes");
        $("#datatable").attr('class', '')
        $("#datatable").addClass(tableClass);

        // Read the Settings and create an array for the Buttons.
        var buttons = [];
        var clipboard = tableau.extensions.settings.get("export-clipboard");
        if (clipboard == "Y") {
          buttons.push('copy');
        }
        var csv = tableau.extensions.settings.get("export-csv");
        if (csv == "Y") {
          buttons.push('csv');
        }
        var excel = tableau.extensions.settings.get("export-excel");
        if (excel == "Y") {
          buttons.push('excel');
        }
        var pdf = tableau.extensions.settings.get("export-pdf");
        if (pdf == "Y") {
          buttons.push('pdf');
        }
        var print = tableau.extensions.settings.get("export-print");
        if (print == "Y") {
          buttons.push('print');
        }

        // If there are 1 or more Export options ticked, then we will add the dom: 'Bfrtip'
        // Else leave this out.
        if (buttons.length > 0) {
          tableReference = $('#datatable').DataTable({
            dom: 'Bfrtip',
            data: tableData,
            columns: data,
            responsive: true,
            buttons: buttons,
            bAutoWidth: false,
            rowGroup: true,
            initComplete: datatableInitCallback,
            drawCallback: datatableDrawCallback,
            oLanguage: datatableLangObj
          });
        } else {
          tableReference = $('#datatable').DataTable({
            data: tableData,
            columns: data,
            responsive: true,
            bAutoWidth: false,
            initComplete: datatableInitCallback,
            drawCallback: datatableDrawCallback,
            oLanguage: datatableLangObj
          });
        }
      })
    }
  }

  function datatableInitCallback(settings, json) {
    // insert table caption
    var table = settings.oInstance.api();
    var $node = $(table.table().node());

    var sheetName = tableau.extensions.settings.get('worksheet');
    var includeTableName = (tableau.extensions.settings.get('include-table-name') == 'Y' ? true : false);

    // add screen reader only h2
    $('#datatable_wrapper').prepend('<h2 class="sr-only">'+sheetName+' | Data Table Extension | Tableau</h2>');


    // add screen readers only caption for table
    // make changes of caption announced by screen reader - used to update caption when sorting changed
    $node.prepend($('<caption id="datatable_caption" class="sr-only" role="alert" aria-live="polite">'+sheetName+'</caption>'));



    // update buttons aria-label to include information about table it is bound to
    table.buttons().each(function(item){
      var $buttonNode = $(item.node);

      var ariaLabel = '';

      if ($buttonNode.hasClass('buttons-copy')) {
        ariaLabel = 'Copy'+(includeTableName ? ' '+sheetName : '')+' table';
      }
      else if ($buttonNode.hasClass('buttons-csv')) {
        ariaLabel = 'CSV of'+(includeTableName ? ' '+sheetName : '')+' table';
      }
      else if ($buttonNode.hasClass('buttons-excel')) {
        ariaLabel = 'Excel of'+(includeTableName ? ' '+sheetName : '')+' table';
      }
      else if ($buttonNode.hasClass('buttons-pdf')) {
        ariaLabel = 'PDF of'+(includeTableName ? ' '+sheetName : '')+' table';
      }
      else if ($buttonNode.hasClass('buttons-print')) {
        ariaLabel = 'Print'+(includeTableName ? ' '+sheetName : '')+' table';
      }

      if (ariaLabel) {
        $buttonNode.attr('aria-label', ariaLabel);
      }
    });


    // update search input label
    var $searchEl = $('#datatable_filter input');
    $searchEl.attr('aria-label', 'Search'+(includeTableName ? ' '+sheetName : '')+' table');


    // set extension's iframe title
    if (window.frameElement) {
      window.frameElement.title = sheetName;
    }


    // set html lang attribute
    document.documentElement.setAttribute('lang', tableau.extensions.environment.language);
  }

  function datatableDrawCallback(settings) {

    var table = settings.oInstance.api();
    var $node = $(table.table().node());

    var $captionEl = $node.find('#datatable_caption');

    var sheetName = tableau.extensions.settings.get('worksheet');
    var includeTableName = (tableau.extensions.settings.get('include-table-name') == 'Y' ? true : false);
    var countOfColumnsForRowHeader = Number(tableau.extensions.settings.get('col-count-row-header'));

    // set row headers if setting is selected
    if (countOfColumnsForRowHeader > 0)
    {
      table.rows().every(function(){
        // for each row update needed number of cells to have role of row header
        $(this.node()).find('td').slice(0, countOfColumnsForRowHeader).attr('role', 'rowheader');
      });
    }

    // fix pagination buttons access by keyboard
    var $paginationNode = $('#datatable_paginate');

    if ($paginationNode.length) {

      // change role of element
      $paginationNode.attr('role', 'navigation');
      // set which element it controls
      $paginationNode.attr('aria-controls', $node.attr('id'));

      var paginateButEls = $paginationNode.find('.paginate_button');

      // if pagination button is disabled or current page (means no action when activated), remove from tab order
      paginateButEls.each(function(){
        var $item = $(this);

        // remove aria-controls set by default for each button (previously we set it for whole navigation element)
        $item.removeAttr('aria-controls');

        // disabled link, for example: prev or next button
        if ($item.hasClass('disabled')) {
          $item.attr('tabindex', -1);
        }
        // current page
        if ($item.hasClass('current')) {
          $item.addClass('disabled');
          $item.attr('tabindex', -1);
        }

        // prev page link text: add sr-only " page" text
        if ($item.attr('id') == 'datatable_previous')
        {
          $item.html('Previous <span class="sr-only">&nbsp;page of'+(includeTableName ? ' '+sheetName+' table' : '')+'</span>');
        }
        // link with number, for example "2" - add sr-only "page " text
        else if ($item.text().trim().match(/^\d+$/))
        {
            // page number
            var pageNum = Number($item.text().trim());
            // items per page
            var itemsPerPage = table.page.len();
            // total number of items in table
            var totalCount = table.data().length;

            // calculate number of first item on the page
            var firstItemNum = (pageNum-1)*itemsPerPage + 1;
            var lastItemNum = firstItemNum + itemsPerPage - 1;
            // correct last item num if last page is not full
            if (lastItemNum > totalCount)
            {
                lastItemNum = totalCount;
            }

            // set aria-label attribute
            $item.attr('aria-label', pageNum+' - entries '+firstItemNum+' to '+lastItemNum+' of '+totalCount+(includeTableName ? ' on '+sheetName+' table' : ''));
        }
        // next page link text: add sr-only " page" text
        else if ($item.attr('id') == 'datatable_next')
        {
          $item.html('Next <span class="sr-only">&nbsp;page of'+(includeTableName ? ' '+sheetName+' table' : '')+'</span>');
        }
      });
    }


    // fix sorting change announce by screen reader
    var order = table.order();

    if (order && order.length)
    {
      // remove aria-sort from any column set previously
      // as per spec it should be applied to only one element at a time: https://www.w3.org/WAI/PF/aria/states_and_properties#aria-sort
      $node.find('[aria-sort]').removeAttr('aria-sort');

      // get header element sorted by currently
      var $columnHeader = $(table.column(order[0][0]).header());

      // set aria-sort
      var ariaSortedByDirection = (order[0][1] == 'asc' ? 'ascending' : (order[0][1] == 'desc' ? 'descending' : 'other'));
      $columnHeader.attr('aria-sort', ariaSortedByDirection);

      // update table caption
      var sortedByDirectionText = ariaSortedByDirection;
      $captionEl.text(sheetName+' sorted by '+$columnHeader.text()+': '+sortedByDirectionText+' order');
    }
    // default for no sort
    else
    {
      $node.find('[aria-sort]').removeAttr('aria-sort');
      $captionEl.text(sheetName);
    }
  }

  // Creates an empty 2D array. we will use this to match the the data set returned
  // by Tableau and repopulate this with the values we want.
  function makeArray(d1, d2) {
    var arr = new Array(d2), i, l;
    for(i = 0, l = d2; i < l; i++) {
        arr[i] = new Array(d1);
    }
    return arr;
  }

  // This is called when you click on the Configure button.
  function configure() {

    const popupUrl = `./dialog.html`;

    let input = "";

    tableau.extensions.ui.displayDialogAsync(popupUrl, input, { height: 540, width: 800 }).then((closePayload) => {
      // The close payload is returned from the popup extension via the closeDialog method.
      $('#interval').text(closePayload);
    }).catch((error) => {
      // One expected error condition is when the popup is closed by the user (meaning the user
      // clicks the 'X' in the top right of the dialog).  This can be checked for like so:
      switch (error.errorCode) {
        case tableau.ErrorCodes.DialogClosedByUser:
          console.log("Dialog was closed by user");
          break;
        default:
          console.error(error.message);
      }
    });
  }
})();
