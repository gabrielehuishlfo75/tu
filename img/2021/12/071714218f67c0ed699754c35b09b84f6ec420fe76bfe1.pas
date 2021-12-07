unit Main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Menus, cxLookAndFeelPainters, cxStyles, cxCustomData,
  cxGraphics, cxFilter, cxData, cxDataStorage, cxEdit, DB, cxDBData,
  cxGridLevel, cxClasses, cxControls, cxGridCustomView,
  cxGridCustomTableView, cxGridTableView, cxGridDBTableView, cxGrid,
  StdCtrls, cxButtons, ExtCtrls, ADODB, dxmdaset, RzPrgres, cxExportGrid4Link;

type
  TForm1 = class(TForm)
    Panel4: TPanel;
    btnAdd: TcxButton;
    btnView: TcxButton;
    cxgridMain: TcxGrid;
    cxgridMainDBTableView1: TcxGridDBTableView;
    cxgridMainLevel1: TcxGridLevel;
    ADOConnection1: TADOConnection;
    adoPage1: TADOQuery;
    dxM1: TDataSource;
    dxMemData1: TdxMemData;
    dxMemData1store: TStringField;
    dxMemData1amt: TFloatField;
    dxMemData1descr: TStringField;
    cxgridMainDBTableView1store: TcxGridDBColumn;
    cxgridMainDBTableView1amt: TcxGridDBColumn;
    cxgridMainDBTableView1descr: TcxGridDBColumn;
    RzProgressBar1: TRzProgressBar;
    dxMemData1times: TIntegerField;
    procedure btnAddClick(Sender: TObject);
    procedure btnViewClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
                  
    function ExportCxGrid(aGrid:TcxGrid;aFile:string=''):boolean;
    function ToExcel(sfilename:string; ADOQuery:TdxMemData):boolean;
implementation

uses ComObj,clipbrd;

{$R *.dfm}



function ToExcel(sfilename:string; ADOQuery:TdxMemData):boolean;

const

xlNormal=-4143;

var

    y      :   integer;

    tsList :   TStringList;

    s,filename   :string;

    aSheet   :Variant;

    excel :OleVariant;

    savedialog   :tsavedialog;

begin

    Result := true;

    try

excel:=CreateOleObject('Excel.Application');

excel.workbooks.add;

except

//screen.cursor:=crDefault;

showmessage('无法调用Excel！');

exit;

    end;

savedialog:=tsavedialog.Create(nil);

//savedialog.FileName:=sfilename;    //存入文件

savedialog.Filter:='Excel文件(*.xls)|*.xls';

    if    savedialog.Execute    then

    begin

if FileExists(savedialog.FileName)    then

try

if application.messagebox('该文件已经存在，要覆盖吗？','询问',mb_yesno+mb_iconquestion)=idyes    then

          DeleteFile(PChar(savedialog.FileName))

else

begin

Excel.Quit;

savedialog.free;

//screen.cursor:=crDefault;

Exit;

end;

except

Excel.Quit;

savedialog.free;

screen.cursor:=crDefault;

Exit;

end;

filename:=savedialog.FileName;

    end;

savedialog.free;

    if    filename=''    then

    begin

result:=true;

Excel.Quit;

//screen.cursor:=crDefault;

exit;

    end;

aSheet:=excel.Worksheets.Item[1];

tsList:=TStringList.Create;

    //tsList.Add('查询结果');    //加入标题



    s:='';    //加入字段名

  {  for y := 0 to adoquery.fieldCount - 1 do

    begin

s:=s+adoQuery.Fields.Fields[y].FieldName+#9 ;

Application.ProcessMessages;

    end;      }

    s:=s+'店名'+#9 ;
    s:=s+'扣罚金额'+#9 ;
    s:=s+'扣罚原因'+#9 ;
    tsList.Add(s);

    try

        try

ADOQuery.First;

While Not ADOQuery.Eof do

begin

s:='';

for y:=1 to ADOQuery.FieldCount-2 do

begin

s:=s+ADOQuery.Fields[y].AsString+#9;

Application.ProcessMessages;

end;

tsList.Add(s);



ADOQuery.next;

end;

Clipboard.AsText:=tsList.Text;

except

result:=false;

        end;

    finally

tsList.Free;

    end;

    aSheet.Paste;

MessageBox(Application.Handle,'数据导出完毕！','系统提示',MB_ICONINFORMATION or MB_OK);

    try

if copy(FileName,length(FileName)-3,4)<>'.xls'    then

FileName:=FileName+'.xls';

Excel.ActiveWorkbook.SaveAs(FileName,    xlNormal,    '',    '',    False,    False);

    except

Excel.Quit;

screen.cursor:=crDefault;

exit;

    end;

Excel.Visible    :=    false; //true会自动打开已经保存的excel

    Excel.Quit;

    Excel := UnAssigned;



end;

function ExportCxGrid(aGrid:TcxGrid;aFile:string=''):boolean;
var sFile :string;
begin
  result := false;
  sFile := aFile;
  if sFile='' then
  with TSaveDialog.Create(nil) do
  try
      filter := 'Excel files(xls;xlsx)|*.xls;*.xlsx';
    if execute then
    begin
      sFile :=fileName;
    end;
  finally
    free;
  end;  //try..
  if sFile='' then Exit;
  try
    ExportGrid4ToEXCEL(sFile,aGrid,True,True,True);
    result := True;
  except
  end;
end;

procedure TForm1.btnAddClick(Sender: TObject);
var
  sOpenFile, sSQL: string;
  ExlApp: Variant;
  ExlSht: Variant;
  iRow, iCloumn, sRowCount: integer;
  lstSQL:TStringList;      
  lst : TStrings;
  Row1, addversion, amt, descr :string;
begin
  sOpenFile := '';
  with TOpenDialog.Create(nil) do
  begin
    try
      DefaultExt := 'xls';
      Filter := 'Excel files(xls;xlsx)|*.xls;*.xlsx';
      Title := 'Load From File';
      if Execute then
        sOpenFile := FileName
      else
      Exit;
    finally
      Free;
    end;
  end;
  try
    ExlApp := CreateOleObject('Excel.Application');
  except
    begin
      ShowMessage('请确认本电脑已经安装了EXCEL');
      Exit;
    end;
  end;
  try
    RzProgressBar1.Visible := True;
    cxgridMain.Enabled := False;
    ExlApp.WorkBooks.Open(sOpenFile);
    ExlSht := ExlApp.workbooks[1].sheets[1];
    lstSQL :=TStringList.Create;
    iRow :=2;
    row1 := Trim(VarToStr(ExlSht.Cells[iRow,1].value));
    sRowCount := ExlSht.UsedRange.Rows.Count;
    RzProgressBar1.Percent:= 0;
    addversion := FormatDateTime('OO_YYYYMMDDHHMMSS',now);   
    dxMemData1.Close; 
    dxMemData1.Open;
    while row1 <> '' do
    begin
      sSQL := '';
      amt := Trim(VarToStr(ExlSht.Cells[iRow,2].value));
      descr := Trim(VarToStr(ExlSht.Cells[iRow,3].value));
      if dxMemData1.Locate('store', row1, []) then
      begin
        dxMemData1.Edit;
        dxMemData1.FieldByName('store').AsString := row1;
        dxMemData1.FieldByName('amt').AsFloat := dxMemData1.FieldByName('amt').AsFloat + StrToFloat(amt);  
        dxMemData1.FieldByName('times').AsInteger := dxMemData1.FieldByName('times').AsInteger + 1;  
        dxMemData1.Post;
        if Pos(descr,dxMemData1.FieldByName('descr').AsString)<=0 then
        begin
          dxMemData1.Edit;
          dxMemData1.FieldByName('descr').AsString := dxMemData1.FieldByName('descr').AsString + '，' + descr;
          dxMemData1.Post;
        end;  
      end else begin
        dxMemData1.Append;
        dxMemData1.FieldByName('store').AsString := row1;
        dxMemData1.FieldByName('amt').AsString := amt;
        dxMemData1.FieldByName('descr').AsString := descr;
        dxMemData1.FieldByName('times').AsInteger := 1;
        dxMemData1.Post;
      end;

      iRow := iRow + 1;
      row1 := Trim(VarToStr(ExlSht.Cells[iRow,1].value));
      RzProgressBar1.Percent:= Round(iRow /sRowCount * 100);
      application.ProcessMessages;
    end;
    RzProgressBar1 .Percent:= 100;
    ShowMessage('导入成功');
  finally
    ExlApp.Quit;
    ExlApp := Unassigned;
    lstSQL.Free;             
    RzProgressBar1.Visible := False;  
    cxgridMain.Enabled := true;
    Screen.Cursor := crDefault;
  end;
end;

procedure TForm1.btnViewClick(Sender: TObject);
begin
  if not dxMemData1.Active or dxMemData1.IsEmpty then
  begin
    ShowMessage('请先导入文件');
    exit;
  end;
  ToExcel('',dxMemData1);
  //	ExportCxGrid(cxgridMain);
end;

end.
