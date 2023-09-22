unit uBrowser;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Grids, DBGrids, Db, DBTables, ExtCtrls, Buttons, ComCtrls,
  DBCtrls, uBrowseMemo, DBInpReq, Menus, ADODB;

type
  TForm_Browser = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    DBGrid1: TDBGrid;
    OpenDialog1: TOpenDialog;
    Panel4: TPanel;
    DBNavigator1: TDBNavigator;
    Edit2: TLabel;
    Label1: TLabel;
    MainMenu1: TMainMenu;
    File1: TMenuItem;
    Exit1: TMenuItem;
    Open1: TMenuItem;
    Close1: TMenuItem;
    Navigate1: TMenuItem;
    FindRecord1: TMenuItem;
    AddRecord1: TMenuItem;
    TopRecord1: TMenuItem;
    PreviousRecord1: TMenuItem;
    NextRepord1: TMenuItem;
    BottomRecord1: TMenuItem;
    HideRecord1: TMenuItem;
    N1: TMenuItem;
    Order1: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    Label_DBFName: TLabel;
    FindDialog1: TFindDialog;
    PopupMenu1: TPopupMenu;
    DisplayMemoasText1: TMenuItem;
    DisplayMemoasObjects1: TMenuItem;
    MovetoFirstColumn1: TMenuItem;
    RecentTables1: TMenuItem;
    procedure Table1AfterScroll(DataSet: TDataSet);
    procedure FormCreate(Sender: TObject);
    procedure DBGrid1DblClick(Sender: TObject);
    procedure Table1AfterOpen(DataSet: TDataSet);
    procedure Open1Click(Sender: TObject);
    procedure Exit1Click(Sender: TObject);
    procedure Navigate1Click(Sender: TObject);
    procedure Close1Click(Sender: TObject);
    procedure AddRecord1Click(Sender: TObject);
    procedure TopRecord1Click(Sender: TObject);
    procedure PreviousRecord1Click(Sender: TObject);
    procedure NextRepord1Click(Sender: TObject);
    procedure BottomRecord1Click(Sender: TObject);
    procedure HideRecord1Click(Sender: TObject);
    procedure FindDialog1Find(Sender: TObject);
    procedure FindRecord1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure DisplayMemoasText1Click(Sender: TObject);
    procedure DisplayMemoasObjects1Click(Sender: TObject);
    procedure PopupMenu1Popup(Sender: TObject);
    procedure MovetoFirstColumn1Click(Sender: TObject);
  private
    { Private declarations }
    fConnectionString: string;
    fTableName: string;
    fTotalRecordCount: integer;
    RecentTables: TStringList;
    RecentConnections: TStringList;
    BrowseMemoList: TList;
    procedure OrderClick(Sender: TObject);
    procedure Enable_Stuff;
    procedure EnableNavigateMenu;
    procedure UpdateTagNames;
    procedure CloseDBF;
    procedure UpdateRecentTablesMenuItem;
    procedure DisplayMemoAsObjects(BrowseHow: TBrowseHow);
    procedure Table1AfterInsert(DataSet: TDataSet);
    function GetTotalRecordCount: integer;
    procedure LoadHistoryList;
    procedure UpdateOpenMemoFields;
  public
    { Public declarations }
    Table1: TADOTable;
    DataSource1: TDataSource;
    procedure OpenTable(aConnectionString: string; aTableName: string);
    procedure AddHistoryItem(aConnectionString, aTableName: string);
    procedure SaveHistoryList;
    procedure OpenRecentDBF(Sender: TObject);
    property TotalRecordCount: integer
             read GetTotalRecordCount;
  end;

var
  Form_Browser: TForm_Browser;

implementation

{$R *.DFM}

uses
  StStrL, StStrZ, StBase, MyUtils, OpenADOTableDialog, Registry;

CONST
  MAXHISTORY = 10;
  FILE_HISTORY = 'File History';
  FILE_PATHS   = 'File Paths';
  DEFAULT_PATH = 'Default Path';
  REG_ADOBROWSER  = 'SoftWare\BCC Software\ADOBrowser';

procedure TForm_Browser.SaveHistoryList;
  var
    i: integer;
    IniFile: TRegIniFile;
begin { TForm_Browser.SaveHistoryList }
  IniFile := TRegIniFile.Create;
  try
    with IniFile do
      begin
        WriteString(REG_ADOBROWSER, 'LastConnectionString', fConnectionString);
        with RecentTables do
          for i := 0 to Count-1 do
            begin
              WriteString(REG_ADOBROWSER, 'Table'+IntToStr(i), RecentTables[i]);
              WriteString(REG_ADOBROWSER, 'Connection'+IntToStr(i), RecentConnections[i]);
            end;
      end;
  finally
    IniFile.Free;
  end;
end;  { TForm_Browser.SaveHistoryList }

procedure TForm_Browser.LoadHistoryList;
  var
    IniFile: TRegIniFile;
    TableName, Connection: string;
    b: boolean;
begin
  IniFile := TRegIniFile.Create;
  try
    with IniFile do
      begin
        fConnectionString := ReadString(REG_ADOBROWSER, 'LastConnectionString', '');
        TableName  := ReadString(REG_ADOBROWSER, 'Table'+IntToStr(RecentTables.Count), '');
        Connection := ReadString(REG_ADOBROWSER, 'Connection'+IntToStr(RecentTables.Count), '');
        b   := TableName <> '';
        while b do
          begin
            RecentTables.Add(TableName);
            RecentConnections.Add(Connection);
            if RecentTables.Count < MAXHISTORY then
              begin
                TableName := ReadString(REG_ADOBROWSER, 'Table'+IntToStr(RecentTables.Count), '');
                Connection := ReadString(REG_ADOBROWSER, 'Connection'+IntToStr(RecentConnections.Count), '');
                b   := (TableName <> '') and (Connection <> '');
              end
            else
              b := false;
          end;
      end;
  finally
    IniFile.Free;
  end;
end;

procedure TForm_Browser.UpdateOpenMemoFields;
  var
    i: integer;
begin
  for i := 0 to BrowseMemoList.Count-1 do
    with TfrmBrowseMemo(BrowseMemoList[i]) do
      begin
        if BrowseHow = bhAsObject then
          UpdateObjectMemo;
      end;
end;


procedure TForm_Browser.Table1AfterScroll(DataSet: TDataSet);
begin
  with DataSet as TADOTable do
    if not ControlsDisabled then
      begin
        if TotalRecordCount > 0 then
          Edit2.Caption := IntToStr(RecNo) + '/' + IntToStr(TotalRecordCount)
        else
          Edit2.Caption := 'No records';
        EnableNavigateMenu;
        UpdateOpenMemoFields;
      end;
end;

procedure TForm_Browser.Table1AfterInsert(DataSet: TDataSet);
begin
  DBGrid1.SelectedField := DataSet.Fields[0];
end;


procedure TForm_Browser.CloseDBF;
  var
    i: integer;
begin
  with Table1 do
    if Active then
      begin
        if State in [dsEdit, dsInsert] then
          Post;
        Active := false;
        for i := 0 to BrowseMemoList.Count-1 do
          TForm(BrowseMemoList[i]).Free;
        BrowseMemoList.Clear;
        Label_DBFName.Caption := '';
        SaveHistoryList;
      end;
  Enable_Stuff;
end;

procedure TForm_Browser.OpenTable(aConnectionString: string; aTableName: string);
  var
    SavedCursor: TCursor;
begin
  SavedCursor   := Screen.Cursor;
  Screen.Cursor := crSQLWait;
  try
    fConnectionString := aConnectionString;
    fTableName        := aTableName;
    with Table1 do
      begin
        CloseDBF;
        ConnectionString := aConnectionString;
        TableName  := FixTableName(aTableName);
        IndexName  := '';
        try
          Active   := true;
          Label_DBFName.Caption := aTableName;
          AddHistoryItem(aConnectionString, aTableName);
        except
          on e:Exception do
            Error(Format('Unable to open %s [%s]', [aTableName, e.message]));
        end;
      end;
  finally
    Screen.Cursor := SavedCursor;
    Enable_Stuff;
  end;
end;

procedure TForm_Browser.Enable_Stuff;
begin
  with Table1 do
    begin
      Open1.Enabled             := not Active;
      Close1.Enabled            := Active;
      if not Active then
        begin
          Label1.Caption      := '';
          Edit2.Caption       := '';
        end;
    end;
  EnableNavigateMenu;
end;

procedure TForm_Browser.FormCreate(Sender: TObject);
begin { TForm_Browser.FormCreate }
  Label_DBFName.Caption := '';
  Table1           := TADOTable.Create(self);
  with Table1 do
    begin
      AfterOpen   := Table1AfterOpen;
      AfterScroll := Table1AfterScroll;
      AfterInsert := Table1AfterInsert;
    end;
  DataSource1         := TDataSource.Create(self);
  DataSource1.DataSet := Table1;
  DBGrid1.DataSource  := DataSource1;
  DBNavigator1.DataSource := DataSource1;

  RecentTables        := TStringList.Create;
  RecentConnections   := TStringList.Create;
  BrowseMemoList      := TList.Create;

  LoadHistoryList;
  UpdateRecentTablesMenuItem;

  Enable_Stuff;
end;  { TForm_Browser.FormCreate }

procedure TForm_Browser.AddHistoryItem(aConnectionString, aTableName: string);
  var
    i: integer;
begin { TForm_Browser.AddHistoryItem }
  i := RecentTables.IndexOf(aTableName);
  if i >= 0 then { already in list }
    begin
      RecentTables.Delete(i);
      RecentTables.Insert(0, aTableName);        { move to the top }
      RecentConnections.Delete(i);
      RecentConnections.Insert(0, aConnectionString);
    end else
  if RecentTables.Count = MAXHISTORY then { history list is full }
    begin
      RecentTables.Delete(MAXHISTORY-1);  { delete the oldest }
      RecentTables.Insert(0, aTableName); { add newest at the top }
      RecentConnections.Delete(MAXHISTORY-1);
      RecentConnections.Insert(0, aConnectionString);
    end
  else
    begin
      RecentTables.Insert(0, aTableName);          { add at the top }
      RecentConnections.Insert(0, aConnectionString);
    end;
  UpdateRecentTablesMenuItem;
end;  { TForm_Browser.AddHistoryItem }

procedure TForm_Browser.UpdateRecentTablesMenuItem;
  var
    i: integer;
    aMenuItem: TMenuItem;
begin
  with RecentTables1 do
    begin
      { empty previous 'Recent DBFs' sub-menu }
      for i := Count-1 downto 0 do
        Delete(i);

      for i := 0 to RecentTables.Count-1 do
        begin
          aMenuItem := TMenuItem.Create(self);
          with aMenuItem do
            begin
              Caption := RecentTables[i];
              OnClick := OpenRecentDBF;
              AutoHotkeys := maManual;
            end;
          Add(aMenuItem);
        end;
    end;
end;



procedure TForm_Browser.DBGrid1DblClick(Sender: TObject);
  var
    aBrowseMemo: TForm;
begin
  aBrowseMemo := TfrmBrowseMemo.Create(self, DataSource1, DBGrid1.SelectedField, bhAsText);
  BrowseMemoList.Add(aBrowseMemo);
  aBrowseMemo.Show;
end;

procedure TForm_Browser.Table1AfterOpen(DataSet: TDataSet);
begin
  with Table1 do
    begin

    end;
  UpdateTagNames;
end;

procedure TForm_Browser.UpdateTagNames;
  var
    aMenuItem: TMenuItem;
    i: integer;
    Items: TStringList;
begin
  with Table1 do
    begin
      Items := TStringList.Create;
      try
        Items.Clear;
        GetIndexNames(Items);
        Items.Add('(recno order)');

        { empty old 'Order' sub-menu }

        with Order1 do
          for i := Count-1 downto 0 do
            Delete(i);

        { Add items to 'Order1' sub-menu }
        for i := Items.Count-1 downto 0 do
          begin
            aMenuItem := TMenuItem.Create(self);
            aMenuItem.Caption := Items[i];
            aMenuItem.OnClick := OrderClick;
            Order1.Add(aMenuItem);
          end;
      finally
        Items.Free;
      end;
    end;
end;


procedure TForm_Browser.Open1Click(Sender: TObject);
begin
  with TfrmOpenADOTable.Create(self) do
    begin
      ConnectionString := fConnectionString;
      TableName        := fTableName;
      if ShowModal = mrOk then
        OpenTable(ConnectionString, TableName);
      Free;
    end;
end;

procedure TForm_Browser.Exit1Click(Sender: TObject);
begin
  Close;
end;

procedure TForm_Browser.Navigate1Click(Sender: TObject);
begin
  EnableNavigateMenu;
end;

procedure TForm_Browser.EnableNavigateMenu;
begin
  with Table1 do
    begin
      Open1.Enabled            := not Active;
      Close1.Enabled           := Active;
      Navigate1.Enabled        := Active;
      FindRecord1.Enabled      := not EOF;
      AddRecord1.Enabled       := Active;
      TopRecord1.Enabled       := not BOF;
      PreviousRecord1.Enabled  := not BOF;
      NextRepord1.Enabled      := not EOF;
      BottomRecord1.Enabled    := not EOF;
      Order1.Enabled           := Active;
    end;
end;


procedure TForm_Browser.Close1Click(Sender: TObject);
begin
  CloseDBF;
end;

procedure TForm_Browser.AddRecord1Click(Sender: TObject);
begin
  Table1.Append;
end;

procedure TForm_Browser.TopRecord1Click(Sender: TObject);
begin
  Table1.First;
end;

procedure TForm_Browser.PreviousRecord1Click(Sender: TObject);
begin
  Table1.Prior;
end;

procedure TForm_Browser.NextRepord1Click(Sender: TObject);
begin
  Table1.Next;
end;

procedure TForm_Browser.BottomRecord1Click(Sender: TObject);
begin
  Table1.Last;
end;

procedure TForm_Browser.HideRecord1Click(Sender: TObject);
begin
  Table1.Delete;
end;

procedure TForm_Browser.OrderClick(Sender: TObject);
  var
    idx: integer;
    aMenuItem: TMenuItem;
begin
  with Table1 do
    begin
      aMenuItem := TMenuItem(Sender);
      idx := Order1.IndexOf(aMenuItem);
      if idx > 0 then { index selected }
        IndexName := aMenuItem.Caption
      else
        IndexName := '';   // recno order selected
    end;
end;

procedure TForm_Browser.OpenRecentDBF(Sender: TObject);
  var
    idx: integer;
begin
  with Sender as TMenuItem do
    begin
      idx := RecentTables.IndexOf(Caption);
      if idx >= 0 then
        OpenTable(RecentConnections[idx], RecentTables[idx]);
    end;
end;

procedure TForm_Browser.FindDialog1Find(Sender: TObject);
  label
    FOUND_IT;
  var
    i: integer;
    Saved_Recno: integer;
    FoundIt: boolean;
    MatchString: string;
    FldNr: integer;
    temp2: string;
begin { TForm_Browser.FindDialog1Find }
  with FindDialog1 do
    begin
      MatchString := UpperCase(FindText);
      with Table1 do
        begin
          DisableControls;
          Saved_Recno := Recno;
          FoundIt     := false;
          FldNr       := -1;   // just to keep compiler happy
          Next;                // start searching at next record
          while not EOF do
            begin
              for i := 0 to Fields.Count-1 do
                with Fields[i] do
                  begin
                    Temp2 := UpperCase(AsString);
                    FoundIt := Pos(MatchString, Temp2) > 0;
                    if FoundIt then
                      begin
                        FldNr := i;
                        goto FOUND_IT;
                      end;
                  end;
              Next;
            end;
FOUND_IT:
          EnableControls;
          if FoundIt then
            begin
              DBGrid1.SelectedField := Fields[FldNr];
              FindDialog1.CloseDialog;
            end
          else
            begin
              Recno := Saved_Recno;
              Error(Format('Search string "%s" could not be found', [FindText]));
            end;
          Table1AfterScroll(Table1);  // do it with Controls Enabled
        end;
  end;
end;  { TForm_Browser.FindDialog1Find }

procedure TForm_Browser.FindRecord1Click(Sender: TObject);
begin
  FindDialog1.Execute;
end;

procedure TForm_Browser.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  SaveHistoryList;
end;

procedure TForm_Browser.DisplayMemoasText1Click(Sender: TObject);
  var
    aBrowseMemo: TForm;
begin
  aBrowseMemo := TfrmBrowseMemo.Create(self, DataSource1, DBGrid1.SelectedField, bhAsText);
  BrowseMemoList.Add(aBrowseMemo);
  aBrowseMemo.Show;
end;

procedure TForm_Browser.DisplayMemoAsObjects(BrowseHow: TBrowseHow);
  var
    aBrowseMemo: TForm;
begin
  aBrowseMemo := TfrmBrowseMemo.Create(self, DataSource1, DBGrid1.SelectedField, BrowseHow);
  BrowseMemoList.Add(aBrowseMemo);
  aBrowseMemo.Show;
end;

procedure TForm_Browser.DisplayMemoasObjects1Click(Sender: TObject);
begin
  DisplayMemoAsObjects(bhAsObject);
end;

procedure TForm_Browser.PopupMenu1Popup(Sender: TObject);
  var
    b1, b2: boolean;
begin
  b1 := DBGrid1.SelectedField is TBlobField;
  DisplayMemoasText1.Enabled := b1;
  if b1 then
    begin
      b2 := MemoContainsObject(DBGrid1.SelectedField);
      DisplayMemoasObjects1.Enabled := b2;
    end
  else
    DisplayMemoasObjects1.Enabled := false;
end;

procedure TForm_Browser.MovetoFirstColumn1Click(Sender: TObject);
begin
  DBGrid1.SelectedField := Table1.Fields[0];
end;

function TForm_Browser.GetTotalRecordCount: integer;
begin
  if fTotalRecordCount = 0 then
    fTotalRecordCount := Table1.RecordCount;
  result := fTotalRecordCount;
end;

end.
