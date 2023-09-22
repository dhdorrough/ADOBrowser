unit OpenADOTableDialog;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ADODB;

type
  TfrmOpenADOTable = class(TForm)
    btnOk: TButton;
    btnCancel: TButton;
    memoConnectionString: TMemo;
    Label1: TLabel;
    btnBuild: TButton;
    cbTableList: TComboBox;
    Label2: TLabel;
    procedure btnBuildClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure memoConnectionStringChange(Sender: TObject);
    procedure cbTableListClick(Sender: TObject);
    procedure cbTableListEnter(Sender: TObject);
  private
    function GetADOConnection: TADOConnection;
    procedure Enable_Buttons;
    function GetConnectionString: string;
    function GetTableName: string;
    procedure SetConnectionString(const Value: string);
    procedure SetTableName(const Value: string);
  private
    { Private declarations }
    fTableList: TStringList;
    fTableName: string;
    fADOConnection: TADOConnection;

    function GetTableList: TStringList;
    procedure MyGetTableNames(Connection: TADOConnection; List: TStrings;
      SystemTables: Boolean);
    property ADOConnection: TADOConnection
             read GetADOConnection;
    property TableList: TStringList
             read GetTableList;
  public
    { Public declarations }
    property ConnectionString: string
             read GetConnectionString
             write SetConnectionString;
    property TableName: string
             read GetTableName
             write SetTableName;
  end;

implementation

{$R *.dfm}

uses
  DB, MyUtils;

procedure TfrmOpenADOTable.btnBuildClick(Sender: TObject);
begin
  memoConnectionString.Text := PromptDataSource(handle, memoConnectionString.Text);
end;

procedure TfrmOpenADOTable.MyGetTableNames(Connection: TADOConnection; List: TStrings;
  SystemTables: Boolean);
var
  TypeField: TField;
  NameField: TField;
  SchemaField: TField;
  TableType: string;
  DataSet: TADODataSet;
  TableName: string;
begin
  DataSet := TADODataSet.Create(nil);
  try
    Connection.OpenSchema(siTables, EmptyParam, EmptyParam, DataSet);
    TypeField:= DataSet.FieldByName('TABLE_TYPE');
    NameField:= DataSet.FieldByName('TABLE_NAME');
    SchemaField:= DataSet.FieldByName('TABLE_SCHEMA');
    List.BeginUpdate;
    try
      List.Clear;
      while not DataSet.EOF do
      begin
        TableType:= TypeField.AsString;
        if (TableType='TABLE') or (TableType='VIEW') or (SystemTables and (TableType='SYSTEM TABLE')) then
        begin
          TableName := NameField.AsString;
          if SchemaField.AsString<>'' then
//          TableName := SchemaField.AsString+'.['+TableName+']';
            TableName := Format('%s.%s', [SchemaField.AsString, TableName]);
          List.Add(TableName);
        end;
        DataSet.Next;
      end;
    finally
      List.EndUpdate;
    end;
  finally
    DataSet.Free;
  end;
end;

function TfrmOpenADOTable.GetTableList: TStringList;

begin { TMW_ADO_Table.GetTableList }
  if not Assigned(fTableList) then
    begin
      fTableList := TStringList.Create;
      MyGetTableNames(ADOConnection, fTableList, false);   // returns names like "dbo.TableName" which cause SQL server problems
      if fTableList.Count=0 then
        MyGetTableNames(ADOConnection, fTableList, true);
    end;
  result := fTableList;
end;  { TMW_ADO_Table.GetTableList }

function TfrmOpenADOTable.GetADOConnection: TADOConnection;
begin
  if not Assigned(fADOConnection) then
    begin
      fADOConnection := TADOConnection.Create(self);
      fADOConnection.ConnectionString := memoConnectionString.Text;
      fADOConnection.LoginPrompt := false;
    end;
  result := fADOConnection;
end;

procedure TfrmOpenADOTable.FormCreate(Sender: TObject);
begin
  memoConnectionString.Clear;
end;

procedure TfrmOpenADOTable.memoConnectionStringChange(Sender: TObject);
begin
  FreeAndNil(fADOConnection);
  FreeAndNil(fTableList);
end;

procedure TfrmOpenADOTable.cbTableListClick(Sender: TObject);
begin
  with cbTableList do
    begin
      if ItemIndex >= 0  then
        fTableName := cbTableList.Items[ItemIndex]
      else
        fTableName := '';
    end;
  Enable_Buttons;
end;

procedure TfrmOpenADOTable.Enable_Buttons;
begin
  btnOk.Enabled := not Empty(fTableName);
end;


function TfrmOpenADOTable.GetConnectionString: string;
begin
  result := memoConnectionString.Text;
end;

function TfrmOpenADOTable.GetTableName: string;
begin
  result := fTableName;
end;

procedure TfrmOpenADOTable.SetConnectionString(const Value: string);
begin
  memoConnectionString.Text := Value;
end;

procedure TfrmOpenADOTable.SetTableName(const Value: string);
begin
  fTableName := Value;
end;

procedure TfrmOpenADOTable.cbTableListEnter(Sender: TObject);
begin
  cbTableList.Items.Assign(TableList);
end;

end.
