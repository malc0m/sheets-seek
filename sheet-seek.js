// Must be run from a button in Google Sheets. Won't work from a custom function call - will return an error about setValue permissions.
// This is easy to do - [Insert>Drawing>Text Box>(Enter button text)>Save and close>RMB on new button>Click on three dots>Assign a script>enter SHEETSEEK>OK]
// Input cell should be initialised with a value that gives the output cell a value smaller than the tolerance set.
// Cell at (inputRow,inputCol) is the value that should be changed to try to achieve the value (targetGoal) in the cell at (targetRow,targetCell).
// "speed" is counterintuitive. It is a divisor, so the higher it is set the slower the script will run, but with less chance of the solution diverging.
// You'll need to play around with the speed value so that it's adding or subtracting a proportion of the output value that's smaller than the last decimal of the initial input value.
// "tolerance" is a stopping criteria. It is how close to the targetGoal you will accept a solution.
// "bound" is the value at which the script assumes it has failed with a divergent solution.
// "flops" and "maxFlops" deal with the case where a solution may not be found within the tolerance and the script is flipping between approaching from 
// the positive and negative sides trying to find a valid solution.

function SHEETSEEK(inputRow=1,inputCol=3,targetRow=10,targetCol=11,targetGoal=0,speed=10000000,tolerance=1,bound=100000,direction="negative",flops=0,maxFlops=4){
  // Access the active sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // grab the value of the cell to change - this is your input value
  var valueToChange = sheet.getRange(inputRow, inputCol).getValue();
  // grab the value of the target cell - this is the output value
  var target = sheet.getRange(targetRow, targetCol).getValue();
  // if the solution is diverging, return an error
  if(Math.abs(targetGoal-target)>bound){
    return "ERROR";
  }
  if(Math.abs(targetGoal-target)<tolerance || flops>maxFlops){
    // if the value in the target cell is within the tolerance to the target value, or it is flip-flopping positive-negative around your  set the value and end script
    sheet.getRange(inputRow, inputCol).setValue(valueToChange);
    return valueToChange;
  }
  else{
    // if approaching the target value from the positive side
    if(target-targetGoal>tolerance){
      // flops deals with the case where the colver can't get within the tolerance and ends up flopping back and forth between approaching from the negative and positive sides
      if(direction=="negative"){
        direction="positive";
        flops++;
      }
      //iterate by adding a proportion of the output value to the input value
      //lower "speed" values (somewhat counterintuitively) will converge faster but have a higher chance to diverge.
      valueToChange = valueToChange-Math.abs(target/speed);
      sheet.getRange(inputRow, inputCol).setValue(valueToChange);
      //loop
      SHEETSEEK(inputRow,inputCol,targetRow,targetCol,targetGoal,speed,tolerance,bound,direction,flops,maxFlops);
    }
    // if approaching the value from the negative side
    if(target-targetGoal<-tolerance){
      if(direction=="positive"){
        direction="negative";
        flops++;
      }
      //iterate by subtracting a proportion of the output value to the input value
      valueToChange = valueToChange+Math.abs(target/speed);
      sheet.getRange(inputRow, inputCol).setValue(valueToChange);
      SHEETSEEK(inputRow,inputCol,targetRow,targetCol,targetGoal,speed,tolerance,bound,direction,flops,maxFlops);
    }
  }
}
