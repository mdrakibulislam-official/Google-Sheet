function doPost(e){

var action = e.parameter.action;

  if(action == 'addItem'){
    return addItem(e);
    }
}



function addItem(e){

var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheets = ss.getSheets();
var sheet = ss.getActiveSheet();

sheet.clear()

var firstRow = sheet.getRange("1:1").setHorizontalAlignment("center");
var firstColumn = sheet.getRange("A:A").setHorizontalAlignment("center");

dataRange = sheet.getRange(1, 1);
dataRange.setValue("Roll");
dataRange = sheet.getRange(1, 2);
dataRange.setValue("Name");
dataRange = sheet.getRange(1, 3);
dataRange.setValue("Email");

var xpath;
var courseCode = e.parameter.courseCode;
//var courseCode = "CSE112"

var firebaseUrl = "https://class-attendance-4e3f9-default-rtdb.firebaseio.com/Courses/"+courseCode+"/";
var baseName = FirebaseApp.getDatabaseByUrl(firebaseUrl);
var sheetName = [baseName.getData()];
var newName = sheetName[0].courseName;
var oldName = ss.getSheetName()
ss.getSheetByName(oldName).setName(newName);

var roll = [], roll1;

var firebaseUrl = "https://class-attendance-4e3f9-default-rtdb.firebaseio.com/Roll/";
var base = FirebaseApp.getDatabaseByUrl(firebaseUrl);
var dataSet1 = [base.getData()];
Logger.log(dataSet1)

for(var i=0; i<dataSet1.length; i++){
  roll1 = dataSet1[i]
  for(var j=0; j<roll1.length; j++){
    roll.push(roll1[j])
  }
}
Logger.log(roll)
//var roll = [ "1026", "1028","1029",1031,1032]

var rows = [], data;


for(var i in roll){
xpath=roll[i]
var firebaseUrl = "https://class-attendance-4e3f9-default-rtdb.firebaseio.com/Student/"+xpath+"/";
var base = FirebaseApp.getDatabaseByUrl(firebaseUrl);
var dataSet = [base.getData()];

  for (var i=0; i< dataSet.length; i++) {
    data = dataSet[i]
    //Logger.log(data)
    rows.push([data.examRoll, data.name, data.email]);
    }
}

dataRange = sheet.getRange(2, 1, rows.length, rows[0].length);  
dataRange.setValues(rows);

var attendance = []
var date = [], date1,singleDate;

var firebaseUrl = "https://class-attendance-4e3f9-default-rtdb.firebaseio.com/AttendanceDate/";
var base = FirebaseApp.getDatabaseByUrl(firebaseUrl);
var dataDate = [base.getData()];
Logger.log(dataDate)

for(var i=0; i<dataDate.length; i++){
  date1 = dataDate[i]
  for(var j=0; j<date1.length; j++){
    date.push(date1[j])
    singleDate = date1[j]

    var firebaseUrlNew = "https://class-attendance-4e3f9-default-rtdb.firebaseio.com/Courses/"+courseCode
    +"/Attendance/"+singleDate+"/";
    var base1 = FirebaseApp.getDatabaseByUrl(firebaseUrlNew);
    var dataFirst = [base1.getData()];

    present = []
    var count = []   

    for(var k=0; k<roll.length; k++){

      var base2 = FirebaseApp.getDatabaseByUrl(firebaseUrlNew+roll[k]+"/");
      var datasecond = [base2.getData()];
      if(datasecond ==""){
        present.push(["0"])
        count.push(0)
      } else{
          present.push(['1'])
          count.push(1)
        }
    }

    attendance.push(count)
    dataRange = sheet.getRange(1, 4+j);
    dataRange.setValue(singleDate);
    dataRange = sheet.getRange(2, 4+j, present.length, present[0].length).setHorizontalAlignment("center");
    dataRange.setValues(present);
    

    Logger.log(present)
    Logger.log(dataFirst)


  }
  
}
Logger.log(dataDate[0].length)
Logger.log(attendance)

dataRange = sheet.getRange(1, 4+dataDate[0].length);
dataRange.setValue("Total Class");
dataRange = sheet.getRange(2, 4+dataDate[0].length, roll.length, 1).setHorizontalAlignment("center");
dataRange.setValue(dataDate[0].length);


classAttendance = []

for(var i=0; i<attendance[0].length; i++){
  classAttendance.push(0)
}
Logger.log(classAttendance)
for(var i=0; i<attendance.length; i++){
  for(j=0; j<attendance[i].length; j++){
    var c = classAttendance[j]+attendance[i][j]
    
    classAttendance[j] = c
  }
}

attend = []

for(var i=0; i<classAttendance.length; i++){
  attend.push([classAttendance[i]])
}
Logger.log(attend)

dataRange = sheet.getRange(1, 5+dataDate[0].length);
dataRange.setValue("Total Attendance");
dataRange = sheet.getRange(2, 5+dataDate[0].length, classAttendance.length, 1).setHorizontalAlignment("center");
dataRange.setValues(attend);

percentage = []
marks = []
for(var i=0; i<attend.length; i++){
  var c = Math.round((attend[i] / dataDate[0].length)*100, 2)

  if(c>90){
    marks.push([10])
  } else if(c>80){
    marks.push([9])
  }else if(c>70){
    marks.push([8])
  }else if(c>60){
    marks.push([7])
  }else if(c>50){
    marks.push([6])
  }else{
    marks.push([0])
  }

  percentage.push([c])
}
Logger.log(percentage)

dataRange = sheet.getRange(1, 6+dataDate[0].length);
dataRange.setValue("Percentage");
dataRange = sheet.getRange(2, 6+dataDate[0].length, percentage.length, 1).setHorizontalAlignment("center");
dataRange.setValues(percentage);

return ContentService.createTextOutput("Success").setMimeType(ContentService.MimeType.TEXT)
}
