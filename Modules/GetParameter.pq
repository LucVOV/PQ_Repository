/**
GetParameter:
    Get a parameter that can be used in a power query definition
**/


let
/* GetParameter
   Retrieve a parameter from the Tab_Parameters table 

   ParameterName = Name of the parameter 
   valueFieldName = (optional - default = "Value") Field that contains the value of the parameter   
*/
GetParameter = (ParameterName as text,optional valueFieldName as nullable text) =>

let
  realValueFieldName = if (valueFieldName = null) then "Value" else valueFieldName,
  ParamSource = Excel.CurrentWorkbook(){[Name="Tab_Parameters"]}[Content],
  ParamRow = Table.SelectRows(ParamSource, each ([Parameter] = ParameterName)),
  Value=
  if Table.IsEmpty(ParamRow)=true
    then " " & ParameterName & " not found>"
  else Record.Field(ParamRow{0},realValueFieldName)
in
  Value  
in
    GetParameter