function onSubmit2(e) {
  var response = e.response;
  var responses = response.getItemResponses();

  // Fetching Calc sheet...
  var monthRecSheet = getMonthSheet().getSheets()[0];
  Logger.log(monthRecSheet);

  // Checking if this is an edit...
  if (resEdit) monthRecSheet.deleteRow(monthRecSheet.getLastRow());

  // Entering data into Calc sheet...
  var vals = [
    [
      new Date(new Date() - 25200000).toDateString(),
      responses[0].getResponse(),
    ],
  ];
  for (var i = 4; i < 10; i++) vals[0].push(responses[i].getResponse());
  vals[0].push(null);
  vals[0].push(null);
  vals[0].push(null);
  vals[0].push(responses[3].getResponse());
  vals[0].push(null);
  var submitRange = monthRecSheet
    .getRange(monthRecSheet.getLastRow() + 1, 1, 1, 13)
    .setValues(vals);
  submitRange.offset(0, 2, 1, 11).setNumberFormat("$0.00");
  submitRange.getCell(1, 1).setHorizontalAlignment("left");
  if (monthRecSheet.getLastRow() % 2 == 1) submitRange.setBackground("#F5F5F5");
  var endBalCell = submitRange.getCell(1, 13);
  endBalCell.setFormula(
    "SUM(" + submitRange.offset(0, 2, 1, 10).getA1Notation() + ")"
  );

  // Checking to see if beginning balance matches last end balance...
  var lr = monthRecSheet.getLastRow();
  if (lr > 2) {
    var lastEnd = monthRecSheet.getRange(lr - 1, 13).getValue();
    var begin = vals[0][11];
    if (Math.abs(begin - lastEnd) > 3) {
      // An attempt to correct the problem. Assuming that Cash I/O was performed before Daily Count
      Logger.log("A problem was caught");
      var prevCells = monthRecSheet.getRange(lr - 1, 9, 1, 3); // getting previous Misc. range
      var cellVals = prevCells.getValues()[0];
      var fixed = false; // was the problem fixed

      for (
        var i = 2;
        i >= 0;
        i-- // refining the cellVals to not contain empty cells at the end
      )
        if (cellVals[i] === "") cellVals.pop();

      var len = cellVals.length - 1;
      for (var i = 0; i < cellVals.length; i++) {
        lastEnd -= cellVals[len - i]; // editing lastEnd to be what it would have been without the last Cash I/O
        if (Math.abs(begin - lastEnd) <= 3) {
          // checking if the problem is fixed
          // performing corrections

          var copyRange = prevCells.offset(0, len - i, 1, i + 1);

          var entryRange = monthRecSheet.getRange(
            monthRecSheet.getLastRow(),
            9,
            1,
            i + 1
          );
          cellVals.splice(0, len - i);
          entryRange.setValues([cellVals]).setNotes(copyRange.getNotes());

          copyRange.clearContent().clearNote();

          fixed = true;
        }
      }

      if (!fixed) {
        submitRange.getCell(1, 12).setBackground("#FFEBEE");

        emailError(
          "Daily Count Sheet Alert!",
          "Employee " +
            responses[0].getResponse() +
            " input a beginning balance that does not match up with the previous end balance.\n",
          monthRecSheet.getParent().getUrl()
        );
      }
    }
  }
}

function getMonthSheet() {
  var date = new Date(new Date() - 25200000);
  var dir = DriveApp.getFolderById("0B88HlOjbQh4rNzNnSmp4YlFiV0k");
  var iter;
  var month = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
  ];

  iter = dir.searchFiles(
    "title contains " +
      "'" +
      month[date.getMonth()] +
      ", " +
      date.getFullYear() +
      "'"
  );
  if (iter.hasNext()) {
    // open existing file here
    var sheet = SpreadsheetApp.open(iter.next());
    return sheet;
  } else {
    // create new sheet here
    var templateSheet = DriveApp.getFileById(
      "1bTu3ARLGjXowjOzpMzMuTQgGvvQqZ0OnA7Ehh2LhdEE"
    );
    var newSheet = templateSheet.makeCopy(
      month[date.getMonth()] + ", " + date.getFullYear(),
      dir
    );
    return SpreadsheetApp.open(newSheet);
  }
}
