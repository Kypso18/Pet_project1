# Pet_project1
School enrollment tracker using google sheets and app script

https://docs.google.com/spreadsheets/d/11Ro9liRRRLLm9jSsVgF9WBANrgXsF1Cnfe-NqA9wAco/edit#gid=0

// Function to Clear the User Form

function clearForm() 
{
  var myGoogleSheet= SpreadsheetApp.getActiveSpreadsheet(); //declare a variable and set with active google sheet 
  var shUserForm    = myGoogleSheet.getSheetByName("User Form"); //declare a variable and set with the User Form worksheet

  //to create the instance of the user-interface environment to use the alert features
  var ui = SpreadsheetApp.getUi();

  // Display a dialog box with a title, message, and "Yes" and "No" buttons. The user can also
  // close the dialog by clicking the close button in its title bar.
  var response = ui.alert("Reset Confirmation", 'Do you want to reset this form?',ui.ButtonSet.YES_NO);

 // Checking the user response and proceed with clearing the form if user selects Yes
 if (response == ui.Button.YES) 
  {
     
  shUserForm.getRange("C5").clear(); //Search Field
  shUserForm.getRange("C8").clear();// Student Number
  shUserForm.getRange("C10").clear(); // Student Name
  shUserForm.getRange("C12").clear(); // Enrollment Status
  shUserForm.getRange("C14").clear(); // Previous grades
  shUserForm.getRange("C16").clear(); //Current year
  shUserForm.getRange("C18").clear();//Subject to Enroll

 //Assigning white as default background color

 shUserForm.getRange("C5").setBackground('#FFFFFF');
 shUserForm.getRange("C8").setBackground('#FFFFFF');
 shUserForm.getRange("C10").setBackground('#FFFFFF');
 shUserForm.getRange("C12").setBackground('#FFFFFF');
 shUserForm.getRange("C14").setBackground('#FFFFFF');
 shUserForm.getRange("C16").setBackground('#FFFFFF');
 shUserForm.getRange("C18").setBackground('#FFFFFF');

  return true ;
  
  }
}

//Declare a function to validate the entry made by user in UserForm

function validateEntry(){

  var myGooglSheet= SpreadsheetApp.getActiveSpreadsheet(); //declare a variable and set with active google sheet 
  var shUserForm    = myGooglSheet.getSheetByName("User Form"); //delcare a variable and set with the User Form worksheet

  //to create the instance of the user-interface environment to use the messagebox features
  var ui = SpreadsheetApp.getUi();

    //Assigning white as default background color

  shUserForm.getRange("C8").setBackground('#FFFFFF');
  shUserForm.getRange("C10").setBackground('#FFFFFF');
  shUserForm.getRange("C12").setBackground('#FFFFFF');
  shUserForm.getRange("C14").setBackground('#FFFFFF');
  shUserForm.getRange("C16").setBackground('#FFFFFF');
  shUserForm.getRange("C18").setBackground('#FFFFFF');
  
//Validating Student Number
  if(shUserForm.getRange("C8").isBlank()==true){
    ui.alert("Please enter Student Number.");
    shUserForm.getRange("C8").activate();
    shUserForm.getRange("C8").setBackground('#FF0000');
    return false;
  }

 //Validating Student Name
  else if(shUserForm.getRange("C10").isBlank()==true){
    ui.alert("Please enter Student Name.");
    shUserForm.getRange("C10").activate();
    shUserForm.getRange("C10").setBackground('#FF0000');
    return false;
  }
  //Validating Enrollment Status
  else if(shUserForm.getRange("C12").isBlank()==true){
    ui.alert("Please chose Enrollment Status in the drop-down.");
    shUserForm.getRange("C12").activate();
    shUserForm.getRange("C12").setBackground('#FF0000');
    return false;
  }
  //Validating Previous grades
  else if(shUserForm.getRange("C14").isBlank()==true){
    ui.alert("Please enter a valid Previous grades.");
    shUserForm.getRange("C14").activate();
    shUserForm.getRange("C14").setBackground('#FF0000');
    return false;
  }
  //Validating Current year
  else if(shUserForm.getRange("C16").isBlank()==true){
    ui.alert("Please select Advisor in the drop-down.");
    shUserForm.getRange("C16").activate();
    shUserForm.getRange("C16").setBackground('#FF0000');
    return false;
  }
  //Validating Subject to Enroll
  else if(shUserForm.getRange("C18").isBlank()==true){
    ui.alert("Please enter Subject to Enroll.");
    shUserForm.getRange("C18").activate();
    shUserForm.getRange("C18").setBackground('#FF0000');
    return false;
  }

  return true;
  
}

// Function to submit the data to Database sheet
function submitData() {
     
  var myGooglSheet= SpreadsheetApp.getActiveSpreadsheet(); //declare a variable and set with active google sheet 

  var shUserForm= myGooglSheet.getSheetByName("User Form"); //delcare a variable and set with the User Form worksheet

  var datasheet = myGooglSheet.getSheetByName("Database"); ////delcare a variable and set with the Database worksheet

  //to create the instance of the user-interface environment to use the messagebox features
  var ui = SpreadsheetApp.getUi();
  
  // Display a dialog box with a title, message, and "Yes" and "No" buttons. The user can also
  // close the dialog by clicking the close button in its title bar.
  var response = ui.alert("Submit", 'Do you want to submit the data?',ui.ButtonSet.YES_NO);

  // Checking the user response and proceed with clearing the form if user selects Yes
  if (response == ui.Button.NO) 
  {return;//exit from this function
  } 
 
  //Validating the entry. If validation is true then proceed with transferring the data to Database sheet
 if (validateEntry()==true) {
  
    var blankRow=datasheet.getLastRow()+1; //identify the next blank row

    datasheet.getRange(blankRow, 1).setValue(shUserForm.getRange("C8").getValue()); //Student Number
    datasheet.getRange(blankRow, 2).setValue(shUserForm.getRange("C10").getValue()); //Student Name
    datasheet.getRange(blankRow, 3).setValue(shUserForm.getRange("C12").getValue()); //Enrollment Status
    datasheet.getRange(blankRow, 4).setValue(shUserForm.getRange("C14").getValue()); // Previous grades
    datasheet.getRange(blankRow, 5).setValue(shUserForm.getRange("C16").getValue()); //Current year
    datasheet.getRange(blankRow, 6).setValue(shUserForm.getRange("C18").getValue());// Subject to Enroll
   
    // date function to update the current date and time as submittted on
    datasheet.getRange(blankRow, 7).setValue(new Date()).setNumberFormat('mm-dd-yyyy h:mm'); //Submitted On
    
    //get the email address of the person running the script and update as Submitted By
    datasheet.getRange(blankRow, 8).setValue(Session.getActiveUser().getEmail()); //Submitted By
    
    ui.alert(' "New Data Saved - Student Number' + shUserForm.getRange("C8").getValue() +' "');
  
  //Clearnign the data from the Data Entry Form

    shUserForm.getRange("C8").clear();
    shUserForm.getRange("C10").clear();
    shUserForm.getRange("C12").clear();
    shUserForm.getRange("C14").clear();
    shUserForm.getRange("C16").clear();
    shUserForm.getRange("C18").clear();
      
 }
}

//Function to Search the record

function searchRecord() {
  
  var myGooglSheet= SpreadsheetApp.getActiveSpreadsheet(); //declare a variable and set with active google sheet 
  var shUserForm= myGooglSheet.getSheetByName("User Form"); //delcare a variable and set with the User Form worksheet
  var datasheet = myGooglSheet.getSheetByName("Database"); ////delcare a variable and set with the Database worksheet
    
  var str       = shUserForm.getRange("C5").getValue();
  var values    = datasheet.getDataRange().getValues(); //getting the entire values from the used range and assigning it to values variable
  var valuesFound=false; //variable to store boolean value
  
 for (var i=0; i<values.length; i++)
    {
    var rowValue = values[i]; //declaraing a variable and storing the value
   
    //checking the first value of the record is equal to search item
    if (rowValue[0] == str) {
           
      shUserForm.getRange("C8").setValue(rowValue[0]) ;
      shUserForm.getRange("C10").setValue(rowValue[1]);
      shUserForm.getRange("C12").setValue(rowValue[2]);
      shUserForm.getRange("C14").setValue(rowValue[3]);
      shUserForm.getRange("C16").setValue(rowValue[4]);
      shUserForm.getRange("C18").setValue(rowValue[5]);
      return; //come out from the search function
      
      }
  }

if(valuesFound==false){
  //to create the instance of the user-interface environment to use the messagebox features
  var ui = SpreadsheetApp.getUi();
  ui.alert("No record found!");
 }

}

//Function to delete the record

function deleteRow() {
  
  var myGooglSheet= SpreadsheetApp.getActiveSpreadsheet(); //declare a variable and set with active google sheet 
  var shUserForm= myGooglSheet.getSheetByName("User Form"); //delcare a variable and set with the User Form worksheet
  var datasheet = myGooglSheet.getSheetByName("Database"); ////delcare a variable and set with the Database worksheet

  //to create the instance of the user-interface environment to use the messagebox features
  var ui = SpreadsheetApp.getUi();
  
  // Display a dialog box with a title, message, and "Yes" and "No" buttons. The user can also
  // close the dialog by clicking the close button in its title bar.
  var response = ui.alert("Submit", 'Do you want to delete the record?',ui.ButtonSet.YES_NO);

 // Checking the user response and proceed with clearing the form if user selects Yes
 if (response == ui.Button.NO) 
 {return;//exit from this function
 } 
    
  var str       = shUserForm.getRange("C5").getValue();
  var values    = datasheet.getDataRange().getValues(); //getting the entire values from the used range and assigning it to values variable
  
  var valuesFound=false; //variable to store boolean value to validate whether values found or not
  
  for (var i=0; i<values.length; i++) 
    {
    var rowValue = values[i]; //declaraing a variable and storing the value // SKIPPER MATTHEW PALOMA
   
    //checking the first value of the record is equal to search item
    if (rowValue[0] == str) {
      
      var  iRow = i+1; //identify the row number
      datasheet.deleteRow(iRow) ; //deleting the row

      //message to confirm the action
      ui.alert(' "Record deleted for Student Number' + shUserForm.getRange("C5").getValue() +' "');

      //Clearing the user form
      shUserForm.getRange("C5").clear() ;     
      shUserForm.getRange("C8").clear() ;
      shUserForm.getRange("C10").clear() ;
      shUserForm.getRange("C12").clear() ;
      shUserForm.getRange("C14").clear() ;
      shUserForm.getRange("C16").clear() ;
      shUserForm.getRange("C18").clear() ;

      valuesFound=true;
      return; //come out from the search function
      }
  }

if(valuesFound==false){
  //to create the instance of the user-interface environment to use the messagebox features
  var ui = SpreadsheetApp.getUi();
  ui.alert("No record found!");
 }

}

//Function to edit the record

function editRecord() {
  
  var myGooglSheet= SpreadsheetApp.getActiveSpreadsheet(); //declare a variable and set with active google sheet 
  var shUserForm= myGooglSheet.getSheetByName("User Form"); //delcare a variable and set with the User Form worksheet
  var datasheet = myGooglSheet.getSheetByName("Database"); ////delcare a variable and set with the Database worksheet

  //to create the instance of the user-interface environment to use the messagebox features
  var ui = SpreadsheetApp.getUi();
  
  // Display a dialog box with a title, message, and "Yes" and "No" buttons. The user can also
  // close the dialog by clicking the close button in its title bar.
  var response = ui.alert("Submit", 'Do you want to edit the data?',ui.ButtonSet.YES_NO);

 // Checking the user response and proceed with clearing the form if user selects Yes
 if (response == ui.Button.NO) 
 {return;//exit from this function
 } 
    
  var str       = shUserForm.getRange("C5").getValue();
  var values    = datasheet.getDataRange().getValues(); //getting the entire values from the used range and assigning it to values variable
  
  var valuesFound=false; //variable to store boolean value to validate whether values found or not
  
  for (var i=0; i<values.length; i++) 
    {
    var rowValue = values[i]; //declaraing a variable and storing the value
   
    //checking the first value of the record is equal to search item
    if (rowValue[0] == str) {
      
      var  iRow = i+1; //identify the row number

      datasheet.getRange(iRow, 1).setValue(shUserForm.getRange("C8").getValue()); //Student Number
      datasheet.getRange(iRow, 2).setValue(shUserForm.getRange("C10").getValue()); //Student Name
      datasheet.getRange(iRow, 3).setValue(shUserForm.getRange("C12").getValue()); //Enrollment Status
      datasheet.getRange(iRow, 4).setValue(shUserForm.getRange("C14").getValue()); // Previous grades
      datasheet.getRange(iRow, 5).setValue(shUserForm.getRange("C16").getValue()); //Current year
      datasheet.getRange(iRow, 6).setValue(shUserForm.getRange("C18").getValue());// Subject to Enroll
   
      // date function to update the current date and time as submittted on
      datasheet.getRange(iRow, 7).setValue(new Date()).setNumberFormat('mm-dd-yyyy h:mm'); //Submitted On
    
      //get the email address of the person running the script and update as Submitted By
      datasheet.getRange(iRow, 8).setValue(Session.getActiveUser().getEmail()); //Submitted By
    
      ui.alert(' "Data updated for - Student Number' + shUserForm.getRange("C8").getValue() +' "');
  
    //Clearnign the data from the Data Entry Form

      shUserForm.getRange("C5").clear();
      shUserForm.getRange("C8").clear();
      shUserForm.getRange("C10").clear();
      shUserForm.getRange("C12").clear();
      shUserForm.getRange("C14").clear();
      shUserForm.getRange("C16").clear();
      shUserForm.getRange("C18").clear();

      valuesFound=true;
      return; //come out from the search function
      }
  }

  if(valuesFound==false){
  //to create the instance of the user-interface environment to use the messagebox features
  var ui = SpreadsheetApp.getUi();
  ui.alert("No record found!");
 }

}




