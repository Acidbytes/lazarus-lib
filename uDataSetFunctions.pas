{
    uDataSetFunctions.pas

    Rotinas para facilitar o trabalho com DataSets.

    por Rubem Rocha - Manaus, AM - Outubro 2024
}

unit uDataSetFunctions;

interface

uses Db;

procedure SaveDataSetToCSV(dataSet: TDataSet; filePath: string;
  header: boolean = true; separator: char = ',');

implementation

uses Classes, SysUtils;

procedure SaveDataSetToCSV(dataSet: TDataSet; filePath: string;
  header: boolean = true; separator: char = ',');
const
  InvalidFieldTypes = [ftUnknown, ftBytes, ftVarBytes, ftBlob, ftGraphic,
    ftFmtMemo, ftParadoxOle, ftDBaseOle, ftTypedBinary, ftCursor, ftADT,
    ftArray, ftReference, ftDataSet, ftOraBlob, ftOraClob, ftVariant,
    ftInterface, ftIDispatch, ftParams, ftStream, ftObject];
  FloatingPointTypes = [ftFloat, ftCurrency, ftBCD, ftFMTBcd, ftSingle];
  StringTypes = [ftString, ftMemo, ftFixedChar, ftWideString];
var
  csvContent: TStringList;
  csvRow: string;
  fieldIndex: integer;
  bm: TBookmark;
  fieldType: TFieldType;

  function IsValidFieldType: boolean;
  begin
    fieldType := dataSet.Fields[fieldIndex].DataType;
    result := not (fieldType in InvalidFieldTypes);
  end;

  function FieldAsString: string;
  begin
    if fieldType = ftDateTime then
      result := FormatDateTime(
        FormatSettings.ShortDateFormat + ' ' + FormatSettings.ShortTimeFormat,
        dataSet.Fields[fieldIndex].Value)

    else if fieldType = ftDate then
      result := FormatDateTime(
        FormatSettings.ShortDateFormat, dataSet.Fields[fieldIndex].Value)

    else if fieldType = ftTime then
      result := FormatDateTime(
        FormatSettings.ShortTimeFormat, dataSet.Fields[fieldIndex].Value)

    else if fieldType in FloatingPointTypes then
    begin
      var oldDecimalSeparator: char := FormatSettings.DecimalSeparator;
      FormatSettings.DecimalSeparator := '.';
      result := FormatFloat('##0.00', dataSet.Fields[fieldIndex].Value);
      FormatSettings.DecimalSeparator := oldDecimalSeparator;
    end

    else if fieldType in StringTypes then
      result := '"' + dataSet.Fields[fieldIndex].AsString + '"'

    else
      result := dataSet.Fields[fieldIndex].AsString
  end;

begin
  csvContent := TStringList.Create;
  bm := dataSet.GetBookmark;
  while not dataSet.ControlsDisabled do
    dataSet.DisableControls;
  try
    try
      dataSet.First;

      if header then
      begin
        var headerString: string := '';
        for fieldIndex := 0 to dataSet.FieldCount - 1 do
        begin
          if not IsValidFieldType then
            continue;
          if Length(headerString) > 0 then
            headerString := headerString + separator;
          headerString := headerString + dataSet.Fields[fieldIndex].FieldName;
        end;
        if Length(headerString) > 0 then
          csvContent.Add(headerString);
      end;

      while not dataSet.Eof do
      begin
        csvRow := '';
        for fieldIndex := 0 to dataSet.FieldCount - 1 do
        begin
          if not IsValidFieldType then
            continue;
          if Length(csvRow) > 0 then
            csvRow := csvRow + separator;
          csvRow := csvRow + FieldAsString;
        end;
        if Length(csvRow) > 0 then
          csvContent.Add(csvRow);
        dataSet.Next;
      end;

    finally
      if csvContent.Count > 0 then
        csvContent.SaveToFile(filePath);
      csvContent.Free;
    end;
  finally
    if dataSet.BookmarkValid(bm) then
    begin
      dataSet.GotoBookmark(bm);
      dataSet.FreeBookmark(bm);
    end;
    while dataSet.ControlsDisabled do
      dataSet.EnableControls;
  end;
end;

end.
