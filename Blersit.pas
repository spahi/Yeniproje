unit Blersit;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, DBCtrls, Grids, DBGrids, DB, ADODB, ComCtrls,
  RpRave, RpDefine, RpCon, RpConDS, Strutils,MmSystem, Mask, OleServer,
  ExcelXP, ComObj, Excel2000;

type
  TForm10 = class(TForm)
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    Bevel1: TBevel;
    DataSource1: TDataSource;
    ADOTable1: TADOTable;
    Bevel2: TBevel;
    BlersqarkqSource: TDataSource;
    Panel1: TPanel;
    Label1: TLabel;
    DBGrid1: TDBGrid;
    Panel2: TPanel;
    Button6: TButton;
    Panel7: TPanel;
    Label8: TLabel;
    DBGrid5: TDBGrid;
    Panel8: TPanel;
    Bevel3: TBevel;
    Panel9: TPanel;
    Panel14: TPanel;
    DBGrid4: TDBGrid;
    ADOQuery1: TADOQuery;
    DataSource5: TDataSource;
    TabSheet3: TTabSheet;
    ListBlersqSource: TDataSource;
    Panel15: TPanel;
    DBGrid3: TDBGrid;
    Label9: TLabel;
    Panel16: TPanel;
    Bevel5: TBevel;
    DBGrid6: TDBGrid;
    DataSource6: TDataSource;
    ADOTable5: TADOTable;
    Button2: TButton;
    Button3: TButton;
    Panel17: TPanel;
    Button4: TButton;
    Panel18: TPanel;
    Label18: TLabel;
    Label20: TLabel;
    Panel19: TPanel;
    Label17: TLabel;
    Panel21: TPanel;
    Bevel6: TBevel;
    Label10: TLabel;
    TabSheet4: TTabSheet;
    Panel20: TPanel;
    DBGrid7: TDBGrid;
    QuerySource2: TDataSource;
    Label19: TLabel;
    Panel22: TPanel;
    Bevel7: TBevel;
    DBGrid8: TDBGrid;
    ADOTable5Nr_Bleresit: TWideStringField;
    ADOTable5Bleresi: TWideStringField;
    ADOTable5Urdheresa_Fin: TWideStringField;
    ADOTable5Fat_Nr: TWideStringField;
    ADOTable5Data: TDateTimeField;
    ADOTable5Valuta: TDateTimeField;
    ADOTable5Debi: TFloatField;
    ADOTable5Kredi: TFloatField;
    ADOQuery1Nr_Bleresit: TWideStringField;
    ADOQuery1Bleresi: TWideStringField;
    ADOQuery1Debi: TFloatField;
    ADOQuery1Kredi: TFloatField;
    ADOQuery1Saldo: TFloatField;
    ADOTable1Nr_Bleresit: TWideStringField;
    ADOTable1Bleresi: TWideStringField;
    ADOTable1Adresa_Bleresit: TWideStringField;
    ADOTable1Numri_Regjistrimit: TWideStringField;
    ADOTable1Vendi: TWideStringField;
    ADOTable1Xhiro_Llogaria: TWideStringField;
    ADOTable1TelNo: TWideStringField;
    Button1: TButton;
    Button5: TButton;
    Button7: TButton;
    Edit1: TEdit;
    Label3: TLabel;
    Edit2: TEdit;
    LabeledEdit1: TLabeledEdit;
    LabeledEdit2: TLabeledEdit;
    Firmdat: TADOTable;
    FirmdatSource: TDataSource;
    SalBlerQuery: TADOQuery;
    SalBlerSource: TDataSource;
    SalBlerQueryDebi: TFloatField;
    SalBlerQueryKredi: TFloatField;
    SalBlerQuerySaldo: TFloatField;
    Button9: TButton;
    Button8: TButton;
    RvBlersit: TRvProject;
    RvBlersitSaldo: TRvDataSetConnection;
    RvBlersQarkullim: TRvDataSetConnection;
    RvBlersFirmdat: TRvDataSetConnection;
    UrdhBlerQark: TRvDataSetConnection;
    UrdhBlerLista: TRvDataSetConnection;
    TabSheet5: TTabSheet;
    Panel3: TPanel;
    Panel4: TPanel;
    Panel5: TPanel;
    DBEdit1: TDBEdit;
    DBEdit2: TDBEdit;
    Periudha: TADOTable;
    PeriudhaSource: TDataSource;
    Label2: TLabel;
    Label4: TLabel;
    PeriudhaData1: TDateTimeField;
    PeriudhaData2: TDateTimeField;
    Label5: TLabel;
    Label6: TLabel;
    DBGrid2: TDBGrid;
    DBGrid9: TDBGrid;
    SalBlerperq: TADOQuery;
    SalBlerperqSource: TDataSource;
    SalBlerperqBleresi: TWideStringField;
    SalBlerperqDebi: TFloatField;
    SalBlerperqKredi: TFloatField;
    SalBlerperqSaldo: TFloatField;
    SalBlerperqNrBleresit: TWideStringField;
    Qarkperq: TADOQuery;
    QarkperqSource: TDataSource;
    QarkperqNrBleresit: TWideStringField;
    QarkperqUrdheresa: TWideStringField;
    QarkperqFatno: TWideStringField;
    QarkperqPershkrimi: TWideStringField;
    QarkperqData: TDateTimeField;
    QarkperqDebi: TFloatField;
    QarkperqKredi: TFloatField;
    Label12: TLabel;
    DBText1: TDBText;
    Edit3: TEdit;
    Edit4: TEdit;
    Edit5: TEdit;
    Edit6: TEdit;
    Edit7: TEdit;
    Button10: TButton;
    RvProject1: TRvProject;
    QPQarkper: TRvDataSetConnection;
    QPPeriudha: TRvDataSetConnection;
    QPFirmdat: TRvDataSetConnection;
    QPSalblerper: TRvDataSetConnection;
    RvBlerPeriudha: TRvDataSetConnection;
    BlersitQarkullimq: TADOQuery;
    BlersitQarkullimqNr_bleresit: TWideStringField;
    BlersitQarkullimqBleresi: TWideStringField;
    BlersitQarkullimqPershkrimi: TWideStringField;
    BlersitQarkullimqUrdheresa_fin: TWideStringField;
    BlersitQarkullimqFat_nr: TWideStringField;
    BlersitQarkullimqData: TDateTimeField;
    BlersitQarkullimqValuta: TDateTimeField;
    BlersitQarkullimqDebi: TFloatField;
    BlersitQarkullimqKredi: TFloatField;
    BlersitQarkullimqUnik: TWideStringField;
    Button11: TButton;
    Listblerq: TADOQuery;
    ListblerqNr_bleresit: TWideStringField;
    ListblerqBleresi: TWideStringField;
    ListblerqVendi: TWideStringField;
    ListblerqAdresa: TWideStringField;
    Blersqark2source: TDataSource;
    Button12: TButton;
    BlersQarkq2: TADOQuery;
    BlersQarkq2Nr_bleresit: TWideStringField;
    BlersQarkq2Bleresi: TWideStringField;
    BlersQarkq2Urdheresa_fin: TWideStringField;
    BlersQarkq2Fat_nr: TWideStringField;
    BlersQarkq2Data: TDateTimeField;
    BlersQarkq2Valuta: TDateTimeField;
    BlersQarkq2Pershkrimi: TWideStringField;
    BlersQarkq2Debi: TFloatField;
    BlersQarkq2Kredi: TFloatField;
    ADOTable5Kulanici: TWideStringField;
    ADOQuery2: TADOQuery;
    ADOQuery2Urdheresa: TWideStringField;
    ADOQuery2Debi: TFloatField;
    ADOQuery2Kredi: TFloatField;
    ADOQuery2Saldo: TFloatField;
    ADOCommand1: TADOCommand;
    ADOCommand2: TADOCommand;
    ADOTable5Pershkrimi: TWideStringField;
    ADOTable1NumriFiskal: TWideStringField;
    Bevel8: TBevel;
    Label23: TLabel;
    Label21: TLabel;
    Label13: TLabel;
    Label15: TLabel;
    Label22: TLabel;
    Label24: TLabel;
    Label25: TLabel;
    Label26: TLabel;
    DBEdit3: TDBEdit;
    DBEdit4: TDBEdit;
    DBEdit5: TDBEdit;
    DBEdit6: TDBEdit;
    DBEdit7: TDBEdit;
    DBEdit8: TDBEdit;
    DBEdit9: TDBEdit;
    DBEdit10: TDBEdit;
    ADOTable1An: TAutoIncField;
    ADOTable1Nit: TWideStringField;
    Button13: TButton;
    TabSheet6: TTabSheet;
    Panel6: TPanel;
    DBGrid10: TDBGrid;
    Panel10: TPanel;
    DBGrid11: TDBGrid;
    SintBlersave: TADOQuery;
    SintBlersaveShifBler: TWideStringField;
    SintBlersaveDebi: TFloatField;
    SintBlersaveKredi: TFloatField;
    SintBlersaveSaldo: TFloatField;
    SintLersSource: TDataSource;
    KartelaSint: TADOQuery;
    KartelaSintNr_bleresit: TWideStringField;
    KartelaSintFat_nr: TWideStringField;
    KartelaSintData: TDateTimeField;
    KartelaSintUrdheresa_fin: TWideStringField;
    KartelaSintPershkrimi: TWideStringField;
    KartelaSintDebi: TFloatField;
    KartelaSintKredi: TFloatField;
    KartelaSintSource: TDataSource;
    KartelaSintBleresi: TWideStringField;
    SintBlersaveBleresi: TWideStringField;
    GrupatBlers: TADOQuery;
    DBGrid12: TDBGrid;
    GrupatBlersSource: TDataSource;
    GrupatBlersShifBler: TWideStringField;
    SintBlerSalldo: TADOQuery;
    SintBlerSalldoDebi: TFloatField;
    SintBlerSalldoKredi: TFloatField;
    SintBlerSalldoSaldo: TFloatField;
    SintBlerSalldoSource: TDataSource;
    DBText5: TDBText;
    DBText6: TDBText;
    DBText7: TDBText;
    Label27: TLabel;
    DBText8: TDBText;
    Button14: TButton;
    Label7: TLabel;
    Bevel4: TBevel;
    Label11: TLabel;
    Label14: TLabel;
    Label16: TLabel;
    Label28: TLabel;
    DBText2: TDBText;
    DBText3: TDBText;
    DBText4: TDBText;
    DBText9: TDBText;
    Button15: TButton;
    Button16: TButton;
    ADOQuery1Vendi: TWideStringField;
    BlersitQarkullimqKulanici: TWideStringField;
    Button17: TButton;
    BlersitQarkullim: TADOTable;
    procedure DBGrid1MouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure DBGrid1KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DBGrid1KeyPress(Sender: TObject; var Key: Char);
    procedure Button1Click(Sender: TObject);
    procedure DBGrid3KeyPress(Sender: TObject; var Key: Char);
    procedure DBGrid6KeyPress(Sender: TObject; var Key: Char);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure ADOTable5Urdhresa_FinChange(Sender: TField);
    procedure ADOTable5DataChange(Sender: TField);
    procedure ADOTable5ValutaChange(Sender: TField);
    procedure Button4Click(Sender: TObject);
    procedure TabSheet3Show(Sender: TObject);
    procedure TabSheet2Show(Sender: TObject);
    procedure TabSheet4Show(Sender: TObject);
    procedure ADOTable1BeforeInsert(DataSet: TDataSet);
    procedure ADOTable5DebiGetText(Sender: TField; var Text: String;
      DisplayText: Boolean);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Button5Click(Sender: TObject);
    procedure Button7Click(Sender: TObject);
    procedure Edit1Change(Sender: TObject);
    procedure Edit1KeyPress(Sender: TObject; var Key: Char);
    procedure Edit2KeyPress(Sender: TObject; var Key: Char);
    procedure Edit2Change(Sender: TObject);
    procedure ADOTable5NewRecord(DataSet: TDataSet);
    procedure DBGrid8KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DBGrid8KeyPress(Sender: TObject; var Key: Char);
    procedure FormCreate(Sender: TObject);
    procedure ADOTable5AfterPost(DataSet: TDataSet);
    procedure LabeledEdit1Change(Sender: TObject);
    procedure LabeledEdit1KeyPress(Sender: TObject; var Key: Char);
    procedure LabeledEdit2Change(Sender: TObject);
    procedure LabeledEdit2KeyPress(Sender: TObject; var Key: Char);
    procedure LabeledEdit2Enter(Sender: TObject);
    procedure LabeledEdit1Enter(Sender: TObject);
    procedure Button8Click(Sender: TObject);
    procedure ADOQuery1AfterScroll(DataSet: TDataSet);
    procedure Button9Click(Sender: TObject);
    procedure DBEdit1KeyPress(Sender: TObject; var Key: Char);
    procedure TabSheet5Show(Sender: TObject);
    procedure DBEdit2KeyPress(Sender: TObject; var Key: Char);
    procedure SalBlerperqAfterScroll(DataSet: TDataSet);
    procedure Edit2Enter(Sender: TObject);
    procedure Edit3Enter(Sender: TObject);
    procedure Edit3Change(Sender: TObject);
    procedure Edit3KeyPress(Sender: TObject; var Key: Char);
    procedure Panel15Enter(Sender: TObject);
    procedure Edit4Enter(Sender: TObject);
    procedure Edit5KeyPress(Sender: TObject; var Key: Char);
    procedure Edit4KeyPress(Sender: TObject; var Key: Char);
    procedure Edit5Change(Sender: TObject);
    procedure Edit4Change(Sender: TObject);
    procedure Edit5Enter(Sender: TObject);
    procedure Edit6Change(Sender: TObject);
    procedure Edit6Enter(Sender: TObject);
    procedure Edit7Enter(Sender: TObject);
    procedure Edit7Change(Sender: TObject);
    procedure Button10Click(Sender: TObject);
    procedure Button11Click(Sender: TObject);
    procedure Button12Click(Sender: TObject);
    procedure ADOQuery2AfterScroll(DataSet: TDataSet);
    procedure BlersQarkq2AfterInsert(DataSet: TDataSet);
    procedure DBGrid7Enter(Sender: TObject);
    procedure DBEdit3KeyPress(Sender: TObject; var Key: Char);
    procedure DBEdit4KeyPress(Sender: TObject; var Key: Char);
    procedure DBEdit8KeyPress(Sender: TObject; var Key: Char);
    procedure DBEdit5KeyPress(Sender: TObject; var Key: Char);
    procedure DBEdit6KeyPress(Sender: TObject; var Key: Char);
    procedure DBEdit7KeyPress(Sender: TObject; var Key: Char);
    procedure DBEdit9KeyPress(Sender: TObject; var Key: Char);
    procedure DBEdit10KeyPress(Sender: TObject; var Key: Char);
    procedure ADOTable1AfterScroll(DataSet: TDataSet);
    procedure ADOTable1AfterInsert(DataSet: TDataSet);
    procedure Button13Click(Sender: TObject);
    procedure TabSheet6Show(Sender: TObject);
    procedure SintBlersaveAfterScroll(DataSet: TDataSet);
    procedure GrupatBlersAfterScroll(DataSet: TDataSet);
    procedure Button14Click(Sender: TObject);
    procedure Button15Click(Sender: TObject);
    procedure Button16Click(Sender: TObject);
    procedure Button17Click(Sender: TObject);
    procedure TabSheet2Enter(Sender: TObject);
  private
     Urdheresa:String;
     Data,Valuta:TDate;
     SifBlers,Pershkrimi:string;
     SifBlersn:Real;
     Data1,Data2: Tdate;
     sData1,sData2: String;
     Yil,Ay,Gun:Word;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form10: TForm10;

implementation

uses Unit1;

{$R *.dfm}

Function Urdhesa(Tarih:Tdate):string;
var
 ay,yil,gun:word;
 sUrdhesa,ay1,gun1:string;
begin
  DecodeDate(Tarih,yil,ay,gun);
  ay1:=inttostr(ay);
  If length(ay1)=1 then
    begin
     ay1:='0'+ay1;
    end;
  gun1:=inttostr(gun);
  If length(gun1)=1 then
    begin
     gun1:='0'+gun1;
    end;
  sUrdhesa:=ay1+gun1;
  result:=sUrdhesa;
end;


Procedure NewBlers;
Var
 p:Integer;
 bulundu:Boolean;
 SifUzun,BoshStr:Integer;
Begin
 With form10 do
    begin
     begin
       If adotable1.RecordCount=0 then
         begin
             SifBlers:='000001';
         end
       else
       If Adotable1.RecordCount>0 then
         begin
             Adotable1.Last;
             //////////////////////
             Try
                begin
                  p:=Pos('/',Adotable1.FieldValues['Nr_Bleresit']);
                  If p=0 then
                     begin
                      SifBlers:=Adotable1.FieldValues['Nr_Bleresit'];
                     end
                  else
                     begin
                      SifBlers:=LeftStr(Adotable1.FieldValues['Nr_Bleresit'],p-1);
                     end;
                  SifUzun:=Length(SifBlers);
                  SifBlersn:=StrtoFloat(SifBlers)+1;
                  SifBlers:=FloatToStr(SifBlersN);
                  BoshStr:=SifUzun-Length(SifBlers);
                  SifBlers:=LeftStr('0000000000',BoshStr)+SifBlers;
                  Bulundu:=Adotable1.Locate('Nr_Bleresit',SifBlers,[]);
                  While Bulundu=True do
                     begin
                       SifBlersn:=StrtoFloat(SifBlers)+1;
                       SifBlers:=FloatToStr(SifBlersN);
                       BoshStr:=SifUzun-Length(SifBlers);
                       SifBlers:=LeftStr('0000000000',BoshStr)+SifBlers;
                       Bulundu:=Adotable1.Locate('Nr_Bleresit',SifBlers,[]);
                     end;
                end;
             Except
                begin
                  MessageDlg('Shifra e fundit papërshtatshëm, regjistroni shifrën me numër !',mtWarning,[mbOk],0);
                  SifBlers:='000000';
                end;
             end;
             ///////
         end;
      end;
   end;
end;


Procedure SumLabel;
Var
Sum1,sum2:Real;
BMark:TBookmark;
begin
      Begin
      Sum1:=0;
      Sum2:=0;
      BMark:=Form10.ADOTable5.GetBookmark;
      Form10.Adotable5.First;
      While not Form10.Adotable5.Eof do
      begin
        Sum1:=Sum1+(Form10.Adotable5.FieldValues['Debi']);
        Sum2:=Sum2+(Form10.Adotable5.FieldValues['Kredi']);
        Form10.AdoTable5.Next;
      end;
      Form10.ADOTable5.GotoBookmark(Bmark);
      Form10.ADOTable5.FreeBookmark(Bmark);
      Form10.Label10.Caption:=FormatFloat('0.00',Sum1)+'€';
      Form10.Label17.Caption:=FormatFloat('0.00',Sum2)+'€';
 end

end;


procedure TForm10.DBGrid1MouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
Var
BlersBul:Boolean;
Sum1,Sum2:Real;
begin
  Sum1:=0;
  Sum2:=0;

If Adotable1.RecordCount=0 then
Begin

End
else
Begin

  BlersBul:=AdoQuery1.Locate('Nr_Bleresit',(Adotable1.FieldValues['Nr_Bleresit']),[]);

End;

end;

procedure TForm10.DBGrid1KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
Var
BlersBul:Boolean;
Sum1, Sum2:Real;
begin
  Sum1:=0;
  Sum2:=0;

If Adotable1.RecordCount=0 then
Begin

End
else
Begin

  BlersBul:=AdoQuery1.Locate('Nr_Bleresit',(Adotable1.FieldValues['Nr_Bleresit']),[]);


End;

end;

procedure TForm10.DBGrid1KeyPress(Sender: TObject; var Key: Char);
begin
If key = #13 then
  Begin
   with DBGrid1 do
    If dbgrid1.SelectedIndex < (Dbgrid1.FieldCount-1) then
      Begin
        Adotable1.Edit;
        Adotable1.Post;
        SelectedIndex:=SelectedIndex+1;
      End
      else
          If ADOTable1.RecNo < ADOTable1.RecordCount then
      Begin
        Adotable1.Edit;
        Adotable1.Post;
        DBEdit3.SetFocus;
      end
      else
      If (Adotable1.RecNo = (Adotable1.RecordCount)) then
      Begin
         Adotable1.Edit;
         Adotable1.Post;
         DBEdit3.SetFocus;
      End;

end;

end;

procedure TForm10.Button1Click(Sender: TObject);
begin
  RvBlersit.ExecuteReport('BlersSaldo');
end;

procedure TForm10.DBGrid3KeyPress(Sender: TObject; var Key: Char);
begin
  If Key = #13 then
   Begin
     Adotable5.Append;
     Adotable5.FieldValues['Nr_Bleresit']:=Listblerq.FieldValues['Nr_Bleresit'];
     Adotable5.FieldValues['Bleresi']:=Listblerq.FieldValues['Bleresi'];
     Adotable5.FieldValues['Urdheresa_Fin']:=Urdheresa;
     Adotable5.FieldValues['Pershkrimi']:=Pershkrimi;
     Adotable5.FieldValues['Data']:=Data;
     Adotable5.FieldValues['Valuta']:=Valuta;
     Adotable5.FieldValues['Debi']:=0;
     Adotable5.FieldValues['Kredi']:=0;
     Adotable5.Post;
     Dbgrid6.SetFocus;
     Dbgrid6.SelectedIndex:=2;
   End;
end;

procedure TForm10.DBGrid6KeyPress(Sender: TObject; var Key: Char);
begin
If key = #13 then
  Begin
   with DBGrid6 do
    If dbgrid6.SelectedIndex < (Dbgrid6.FieldCount-1) then
      Begin
        Adotable5.Edit;
        Adotable5.Post;
        SelectedIndex:=SelectedIndex+1;
      End
      else
          If ADOTable5.RecNo < ADOTable5.RecordCount then
      Begin
        Adotable5.Edit;
        Adotable5.Post;
        Adotable5.Next;
        Dbgrid6.SelectedIndex:=0;
      end
      else
      If (Adotable5.RecNo = (Adotable5.RecordCount)) then
      Begin
        Urdheresa:=Adotable5.FieldValues['Urdheresa_Fin'];
        Pershkrimi:=Adotable5.FieldValues['Pershkrimi'];
        SumLabel;
        Edit5.SetFocus;
        //Dbgrid3.SetFocus;
      End;
end;

end;

procedure TForm10.Button2Click(Sender: TObject);
begin
If   Adotable5.RecordCount=0 then
  Begin
    ShowMessage('Nuk ka të dhëna!');
  End
  Else
   begin
    Repeat
     Adotable5.Delete;
    Until
     Adotable5.RecordCount=0;
   end;

end;

procedure TForm10.Button3Click(Sender: TObject);
begin
  If   Adotable5.RecordCount=0 then
    Begin
      ShowMessage('Nuk ka të dhëna!');
    End
  Else
    begin
      Adotable5.Delete;
    end;
end;

procedure TForm10.ADOTable5Urdhresa_FinChange(Sender: TField);
begin
  Urdheresa:=Adotable5.FieldValues['Urdheresa_Fin'];
end;

procedure TForm10.ADOTable5DataChange(Sender: TField);
begin
 Adotable5.FieldValues['Urdheresa_Fin']:=Urdhesa(Adotable5.FieldValues['Data']);
 Data:=Adotable5.FieldValues['Data'];
end;

procedure TForm10.ADOTable5ValutaChange(Sender: TField);
begin
Valuta:=Adotable5.FieldValues['Valuta'];
end;

procedure TForm10.Button4Click(Sender: TObject);
Var
  TransTamam,mError:Boolean;
  Defa:Integer;
begin
  If Adotable5.RecordCount=0 then
    Begin
      ShowMessage('Nuk ka të dhëna!');
    End
  else
   begin
     BlersitQarkullim.Open;
     /////////////////////////////////////////////////
     TransTamam:=False;
     Defa:=0;
     While not TransTamam=True do
        begin
          ///////////////////////////
          ADOCommand1.CommandText:='Begin transaction';
          ADOCommand1.Execute;
          //////////////////////
          Try
            begin
              Adotable5.First;
              While not Adotable5.Eof Do
                 Begin
                   BlersitQarkullim.Append;
                   BlersitQarkullim.FieldValues['Nr_Bleresit']:=ADOTable5.FieldValues['Nr_Bleresit'];
                   BlersitQarkullim.FieldValues['Bleresi']:= ADOTable5.FieldValues['Bleresi'];
                   BlersitQarkullim.FieldValues['Pershkrimi']:= ADOTable5.FieldValues['Pershkrimi'];
                   BlersitQarkullim.FieldValues['Urdheresa_Fin']:= ADOTable5.FieldValues['Urdheresa_Fin'];
                   BlersitQarkullim.FieldValues['Fat_Nr']:= ADOTable5.FieldValues['Fat_Nr'];
                   BlersitQarkullim.FieldValues['Data']:= ADOTable5.FieldValues['Data'];
                   BlersitQarkullim.FieldValues['Valuta']:= ADOTable5.FieldValues['Valuta'];
                   BlersitQarkullim.FieldValues['Debi']:= ADOTable5.FieldValues['Debi'];
                   BlersitQarkullim.FieldValues['Kredi']:= ADOTable5.FieldValues['Kredi'];
                   BlersitQarkullim.FieldValues['Kulanici']:=Form1.Kullanici;
                   BlersitQarkullim.Post;
                   Adotable5.Next;
                   SifBlers:=ADOTable5.FieldValues['Nr_Bleresit'];
                 End;
              ////////
              Adotable5.Close;
              ADOCommand2.CommandText:='Delete * From BlersitQarkullimTemp';
              ADOCommand2.Execute;
              Adotable5.Open;
              mError:=False;
            end;
          Except
            begin
              mError:=True;
              Defa:=Defa+1;
            end;
          end;
        //////////////
        If mError=False then
            begin
              ADOCommand1.CommandText:='Commit transaction';
              ADOCommand1.Execute;
              BlersitQarkullim.Close;
             // BlersitQarkullimq.Close;
             // BlersitQarkullimq.Parameters.ParamByName('ParShifbler').Value:=Listblerq.FieldValues['Nr_Bleresit'];
             // BlersitQarkullimq.Open;
              Adotable5.Close;
              Adotable5.Open;
              TransTamam:=True;
            end
        else
            begin
              ADOCommand1.CommandText:='Rollback transaction';
              ADOCommand1.Execute;
              BlersitQarkullim.Close;
             // BlersitQarkullimq.Close;
             // BlersitQarkullimq.Parameters.ParamByName('ParShifbler').Value:=Listblerq.FieldValues['Nr_Bleresit'];
             // BlersitQarkullimq.Open;
              Adotable5.Close;
              Adotable5.Open;
              If Defa>5 then
                begin
                 TransTamam:=True;
                 Showmessage('Gabim e transferit, kontrolloni listën the provoni edhe njiher !');
                end;
            end;
       end;
   end;
     ////////////////////////////////////////////
end;

procedure TForm10.TabSheet3Show(Sender: TObject);
begin
  Listblerq.Close;
  ListBlerq.Open;
end;

procedure TForm10.TabSheet2Show(Sender: TObject);
begin
  ADOTable1.Open;
  AdoQuery1.Close;
  AdoQuery1.Open;
  If ADOTable1.RecordCount>0 then
     begin
       If SifBlers>'0' then
          begin
            AdoQuery1.Locate('Nr_bleresit',SifBlers,[LopartialKey]);
          end;
     end;
  SalBlerQuery.Close;
  SalBlerQuery.Open;
end;

procedure TForm10.TabSheet4Show(Sender: TObject);
begin
  AdoQuery2.Close;
  AdoQuery2.Open;
  DBGrid8.ReadOnly:=True;
end;


procedure TForm10.ADOTable1BeforeInsert(DataSet: TDataSet);
begin
 If Adotable1.Eof and Adotable1.Bof then
   begin

   end
 else
   begin
      If Adotable1.FieldValues['Bleresi']=null then
      Begin
       Adotable1.Delete;
       Adotable1.Edit;
       Adotable1.Post;
      end;
   end;
   NewBlers;
end;

procedure TForm10.ADOTable5DebiGetText(Sender: TField; var Text: String;
  DisplayText: Boolean);
begin
   If Sender.IsNull then
      Text:='0.00'
   else
      Text:=FormatFloat('0.00',Sender.AsFloat);
end;

procedure TForm10.FormShow(Sender: TObject);
begin
  PageControl1.TabIndex:=1;
  Adotable1.Open;
  Adotable1.UpdateBatch();
  Adotable1.Close;
  Adotable1.Open;
  Adotable1.IndexFieldNames:='Nr_bleresit';
  If Adotable1.RecordCount>0 then
     begin
       SifBlers:=Adotable1.FieldValues['Nr_Bleresit'];
     end;
  //////////
  Listblerq.Close;
  Listblerq.Open;
  Adotable5.Open;
  Periudha.Open;
  SalBlerQuery.Close;
  SalBlerQuery.Open;
  AdoQuery1.Close;
  AdoQuery1.Open;
  AdoQuery1.Locate('Nr_Bleresit',SifBlers,[]);
 { AdoQuery2.Close;
  AdoQuery2.Open;}
end;

procedure TForm10.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Adotable1.Close;
  BlersitQarkullimq.Close;
  Listblerq.Close;
  Adotable5.Close;
  Blersqarkq2.Close;
  AdoQuery1.Close;
  AdoQuery2.Close;
  /////////
  SintBlersave.Close;
  GrupatBlers.Close;
  //////////
  SalBlerperq.Close;
  Periudha.Close;
  SalBlerQuery.Close;
end;

procedure TForm10.Button5Click(Sender: TObject);
begin
  Periudha.Edit;
  Periudha.FieldValues['data2']:=date;
  Periudha.Post;
  Periudha.UpdateBatch();
  RvBlersit.ExecuteReport('KartelaBlersit');
end;

procedure TForm10.Button7Click(Sender: TObject);
Var
C:Word;
Varmi:Boolean;
begin
C:=MessageDlg('A të fshihet Blerësi?',MtConfirmation,[MbYes,MbNo], 0);
If C=MrYes then
begin
 VarMi:=AdoQuery1.Locate('Nr_Bleresit',Adotable1.FieldValues['Nr_Bleresit'],[]);
 If varmi=true then
 begin
  ShowMessage('Nuk mund të fshihet Blerësi i cili ka qarkullim!');
 end
 else
 begin
  If Adotable1.RecordCount>0 then
  begin
    Adotable1.Delete;
  end;
 end;

end
else
begin
end;
end;


procedure TForm10.Edit1Change(Sender: TObject);
begin
Adotable1.Locate('Bleresi',Edit1.Text,[LopartialKey]);
end;

procedure TForm10.Edit1KeyPress(Sender: TObject; var Key: Char);
begin
If key=#13 then
begin
DBgrid1.SetFocus;
Edit1.Text:='';
end;
end;

procedure TForm10.Edit2KeyPress(Sender: TObject; var Key: Char);
begin
  If Key=#13 then
    begin
     DBGrid5.SetFocus;
     Edit2.Text:='';
    end;
end;

procedure TForm10.Edit2Change(Sender: TObject);
begin
   AdoQuery1.Locate('Bleresi',Edit2.Text,[LopartialKey]);
end;

procedure TForm10.ADOTable5NewRecord(DataSet: TDataSet);
begin
   Adotable5.FieldValues['Urdheresa_Fin']:=Urdheresa;
end;

procedure TForm10.DBGrid8KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  If (key=VK_F2) then
      Begin
        DBGrid8.ReadOnly:=False;
      end
  Else
   If (key=VK_DELETE) then
      Begin
         Button8.Click;
      end;
end;

procedure TForm10.DBGrid8KeyPress(Sender: TObject; var Key: Char);
begin
 If Key=#13 then
    begin
      If DBGrid8.ReadOnly=False then
        begin
          If dbgrid8.SelectedIndex < (Dbgrid8.FieldCount-1) then
             Begin
               Dbgrid8.SelectedIndex:=Dbgrid8.SelectedIndex+1;
             End
          else
             begin
               Blersqarkq2.Edit;
               Blersqarkq2.Post;
             end;
        end;
    end;
end;

procedure TForm10.FormCreate(Sender: TObject);
begin
   Urdheresa:='';
   Pershkrimi:='';
   Data:=Date;
   Valuta:=Date;
end;


procedure TForm10.ADOTable5AfterPost(DataSet: TDataSet);
begin
   Data:=Adotable5.FieldValues['Data'];
   Valuta:=Adotable5.FieldValues['valuta'];
   SumLabel;
end;

procedure TForm10.LabeledEdit1Change(Sender: TObject);
begin
  ListBlerq.Locate('Nr_bleresit',LabeledEdit1.Text,[LopartialKey]);
end;

procedure TForm10.LabeledEdit1KeyPress(Sender: TObject; var Key: Char);
begin
 If Key=#13 then
  begin
   DBGrid3.SetFocus;
   LabeledEdit1.Text:='';
  end;
end;
procedure TForm10.LabeledEdit2Change(Sender: TObject);
begin
  ListBlerq.Locate('Bleresi',LabeledEdit2.Text,[LopartialKey]);
end;

procedure TForm10.LabeledEdit2KeyPress(Sender: TObject; var Key: Char);
begin
If Key=#13 then
  begin
   DBGrid3.SetFocus;
   LabeledEdit2.Text:='';
  end;
end;

procedure TForm10.LabeledEdit2Enter(Sender: TObject);
begin
   ListBlerq.Sort:='Bleresi';
end;

procedure TForm10.LabeledEdit1Enter(Sender: TObject);
begin
  ListBlerq.Sort:='Nr_bleresit';
end;

procedure TForm10.Button8Click(Sender: TObject);
var
 RecPozicion,I:Integer;
 C:Word;
begin
  C:=MessageDlg('A jeni të sigurt ?',MtConfirmation,[MbYes,MbNo], 0);
  If c=mrYes then
    begin
      Blersqarkq2.Delete;
      Blersqarkq2.UpdateBatch();
      Recpozicion:=AdoQuery2.RecNo;
      AdoQuery2.Close;
      AdoQuery2.Open;
      For I:=1 to (RecPozicion-1) do
        begin
           AdoQuery2.Next;
        end;
    end;
end;


procedure TForm10.ADOQuery1AfterScroll(DataSet: TDataSet);
begin
  BlersitQarkullimq.Close;
  BlersitQarkullimq.Parameters.ParamByName('ParShifbler').Value:=AdoQuery1.FieldValues['Nr_Bleresit'];
  BlersitQarkullimq.Open;
  ADOTable1.Locate('Nr_bleresit',AdoQuery1.FieldValues['Nr_Bleresit'],[]);
end;

procedure TForm10.Button9Click(Sender: TObject);
begin
   RvBlersit.ExecuteReport('UrdhBlerQark');
end;

procedure TForm10.DBEdit1KeyPress(Sender: TObject; var Key: Char);
begin
   If Key = #13 Then
     begin
       Periudha.Edit;
       Periudha.Post;
       Periudha.UpdateBatch();
       DBedit2.SetFocus;
     end;
end;

procedure TForm10.TabSheet5Show(Sender: TObject);
Var
  Sgun,SAy,SYil:string;
  Yil2:Word;
begin
   ////////
   Periudha.Open;
   Data2:=date;
   Periudha.Edit;
   Periudha.FieldValues['data2']:=Data2;
   Decodedate(Data2,Yil2,Ay,Gun);
   Sgun:='01';
   Say:='01';
   DBEdit1.Text:= Sgun+'-'+Say+'-'+IntTostr(Yil2);
   Decodedate(Data2,Yil,Ay,Gun);
   If Length(IntTostr(Gun))=1 then
      Sgun:='0'+IntTostr(Gun)
   else
      Sgun:=IntTostr(Gun);
   If Length(IntTostr(Ay))=1 then
      Say:='0'+IntTostr(Ay)
   else
      Say:=IntTostr(Ay);
   DBEdit2.Text:= Sgun+'-'+Say+'-'+IntTostr(Yil);
   Periudha.Post;
   Periudha.Close;
   Periudha.Open;
   DBEdit1.SetFocus;
  //////////////////
end;

procedure TForm10.DBEdit2KeyPress(Sender: TObject; var Key: Char);
begin
 If Key = #13 Then
   begin
      Periudha.Edit;
      Periudha.Post;
      Periudha.UpdateBatch();
      Periudha.Close;
      Periudha.Open;
      SalBlerperq.Close;
      SalBlerperq.Open;
      DBGrid2.SetFocus;
   end;
end;

procedure TForm10.SalBlerperqAfterScroll(DataSet: TDataSet);
begin
  Qarkperq.Close;
  Qarkperq.Parameters.ParamByName('Parambleresi').Value:=SalBlerperq.FieldValues['Nrbleresit'];
  Qarkperq.Open;
  Qarkperq.Sort:='Data,Fatno';
end;

procedure TForm10.Edit2Enter(Sender: TObject);
begin
  ADOQuery1.Sort:='Bleresi';
end;

procedure TForm10.Edit3Enter(Sender: TObject);
begin
   ADOQuery1.Sort:='Nr_Bleresit';
end;

procedure TForm10.Edit3Change(Sender: TObject);
begin
  AdoQuery1.Locate('Nr_Bleresit',Edit3.Text,[LopartialKey]);
end;

procedure TForm10.Edit3KeyPress(Sender: TObject; var Key: Char);
begin
   If Key=#13 then
    begin
     DBGrid5.SetFocus;
     Edit3.Text:='';
    end;
end;

procedure TForm10.Panel15Enter(Sender: TObject);
begin
   ListBlerq.Sort:='Bleresi';
end;

procedure TForm10.Edit4Enter(Sender: TObject);
begin
 ListBlerq.Sort:='Nr_Bleresit';
end;

procedure TForm10.Edit5KeyPress(Sender: TObject; var Key: Char);
begin
  If Key=#13 then
    begin
     DBGrid3.SetFocus;
     Edit5.Text:='';
    end;
end;

procedure TForm10.Edit4KeyPress(Sender: TObject; var Key: Char);
begin
  If Key=#13 then
    begin
     DBGrid3.SetFocus;
     Edit4.Text:='';
    end;
end;

procedure TForm10.Edit5Change(Sender: TObject);
begin
  ListBlerq.Locate('Bleresi',Edit5.Text,[LopartialKey]);
end;

procedure TForm10.Edit4Change(Sender: TObject);
begin
  Listblerq.Locate('Nr_Bleresit',Edit4.Text,[LopartialKey]);
end;

procedure TForm10.Edit5Enter(Sender: TObject);
begin
  ListBlerq.Sort:='Bleresi';
end;



procedure TForm10.Edit6Change(Sender: TObject);
begin
  SalBlerperq.Locate('Bleresi',Edit6.Text,[LopartialKey]);
end;

procedure TForm10.Edit6Enter(Sender: TObject);
begin
  Edit6.Text:='';
  SalBlerperq.Sort:='Bleresi';
end;

procedure TForm10.Edit7Enter(Sender: TObject);
begin
  Edit7.Text:='';
  SalBlerperq.Sort:='NrBleresit';
end;

procedure TForm10.Edit7Change(Sender: TObject);
begin
  SalBlerperq.Locate('NrBleresit',Edit7.Text,[LopartialKey]);
end;

procedure TForm10.Button10Click(Sender: TObject);
begin
  RvProject1.ExecuteReport('KartBlerPer');
end;

procedure TForm10.Button11Click(Sender: TObject);
begin
    RvProject1.ExecuteReport('SaldoBlerPer');
end;

procedure TForm10.Button12Click(Sender: TObject);
begin
 RvBlersit.ExecuteReport('ListUrdhBlers');
end;


procedure TForm10.ADOQuery2AfterScroll(DataSet: TDataSet);
begin
   Blersqarkq2.Close;
   BlersQarkq2.Parameters.ParamByName('ParUrdhesa').Value:=AdoQuery2.FieldValues['Urdheresa'];
   Blersqarkq2.Open;
end;

procedure TForm10.BlersQarkq2AfterInsert(DataSet: TDataSet);
begin
 Blersqarkq2.Cancel;
end;


procedure TForm10.DBGrid7Enter(Sender: TObject);
begin
  Adoquery2.Close;
  Adoquery2.Open;
end;

procedure TForm10.DBEdit3KeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
   DBEdit4.SetFocus;
end;

procedure TForm10.DBEdit4KeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
   DBEdit5.SetFocus;
end;

procedure TForm10.DBEdit8KeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
   DBEdit9.SetFocus;
end;

procedure TForm10.DBEdit5KeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
   DBEdit6.SetFocus;
end;

procedure TForm10.DBEdit6KeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
   DBEdit7.SetFocus;
end;

procedure TForm10.DBEdit7KeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
   DBEdit8.SetFocus;
end;

procedure TForm10.DBEdit9KeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
   DBEdit10.SetFocus;
end;

procedure TForm10.DBEdit10KeyPress(Sender: TObject; var Key: Char);
begin
   If key = #13 then
    Begin
      AdoTable1.Edit;
      AdoTable1.Post;
      DBGrid1.SetFocus;
      Dbgrid1.SelectedIndex:=0;
    End
end;

procedure TForm10.ADOTable1AfterScroll(DataSet: TDataSet);
begin
  If not Adotable1.Eof then
     begin
       Dbgrid1.ReadOnly:=True;
     end
  Else
     begin
       Dbgrid1.ReadOnly:=False;
     end;
end;

procedure TForm10.ADOTable1AfterInsert(DataSet: TDataSet);
begin
 Adotable1.FieldValues['Nr_Bleresit']:=SifBlers;
end;

procedure TForm10.Button13Click(Sender: TObject);
begin
  Adotable1.IndexFieldNames:='Nr_Bleresit';
  Adotable1.First;
  Adotable1.Append;
  DBGrid1.SetFocus;
  DBGrid1.Fields[0].FocusControl;
end;

procedure TForm10.TabSheet6Show(Sender: TObject);
begin
  SintBlersave.Close;
  SintBlersave.Open;
  GrupatBlers.Close;
  GrupatBlers.Open;
end;

procedure TForm10.SintBlersaveAfterScroll(DataSet: TDataSet);
begin
  KartelaSint.Close;
  KartelaSint.Parameters.ParamByName('ParamShifBler').Value:=LeftStr(SintBlersave.FieldValues['ShifBler'],3);
  KartelaSint.Open;
end;

procedure TForm10.GrupatBlersAfterScroll(DataSet: TDataSet);
begin
  SintBlersave.Close;
  SintBlersave.Parameters.ParamByName('GrupParam').Value:=GrupatBlers.FieldValues['ShifBler'];
  SintBlersave.Open;
  SintBlerSalldo.Close;
  SintBlerSalldo.Parameters.ParamByName('GrupParam').Value:=GrupatBlers.FieldValues['ShifBler'];
  SintBlerSalldo.Open;
end;

procedure TForm10.Button14Click(Sender: TObject);
var
   XApp:Variant;
   sheet:Variant;
   r,c:Integer;
   row,col:Integer;
   filName:Integer;
   q:Integer;
begin
   XApp:=CreateOleObject('Excel.Application');
   XApp.Visible:=true;
   XApp.WorkBooks.Add(-4167);
   XApp.WorkBooks[1].WorkSheets[1].Name:='Sheet1';
   sheet:=XApp.WorkBooks[1].WorkSheets['Sheet1'];
   //For filName:=0 to KartelaSint.FieldCount-1 do
   For filName:=0 to 7 do
      begin
       q:=filName+1;
       sheet.Cells[1,q]:=KartelaSint.Fields[filName].FieldName;
      end;
   //////////
   For r:=0 to KartelaSint.RecordCount-1 do
     begin
       //for c:=0 to KartelaSint.FieldCount-1 do
      for c:=0 to 7 do
         begin
           row:=r+3;
           col:=c+1;
           sheet.Cells[row,col]:=KartelaSint.Fields[c].AsString;
         end;
       KartelaSint.Next;
     end;
  // XApp.WorkSheets['Sheet1'].Range['A1:AA1'].Font.Bold:=True;
  // XApp.WorkSheets['Sheet1'].Range['A1:F1'].Borders.LineStyle :=10;
  // XApp.WorkSheets['Sheet1'].Range['A3:F'+inttostr(KartelaSint.RecordCount+2)].Borders.LineStyle :=1;
   XApp.WorkSheets['Sheet1'].Columns[1].ColumnWidth:=10;
   XApp.WorkSheets['Sheet1'].Columns[2].ColumnWidth:=20;
   XApp.WorkSheets['Sheet1'].Columns[3].ColumnWidth:=9;
   XApp.WorkSheets['Sheet1'].Columns[4].ColumnWidth:=10;
   XApp.WorkSheets['Sheet1'].Columns[5].ColumnWidth:=10;
   XApp.WorkSheets['Sheet1'].Columns[6].ColumnWidth:=10;
   XApp.WorkSheets['Sheet1'].Columns[7].ColumnWidth:=10;
   XApp.WorkSheets['Sheet1'].Columns[8].ColumnWidth:=10;
   //////////////
end;

procedure TForm10.Button15Click(Sender: TObject);
var
   XApp:Variant;
   sheet:Variant;
   r,c:Integer;
   row,col:Integer;
   filName:Integer;
   q:Integer;
begin
   XApp:=CreateOleObject('Excel.Application');
   XApp.Visible:=true;
   XApp.WorkBooks.Add(-4167);
   XApp.WorkBooks[1].WorkSheets[1].Name:='Sheet1';
   sheet:=XApp.WorkBooks[1].WorkSheets['Sheet1'];
   For filName:=0 to ADOQuery1.FieldCount-1 do
      begin
       q:=filName+1;
       sheet.Cells[1,q]:=ADOQuery1.Fields[filName].FieldName;
      end;
   //////////
   For r:=0 to ADOQuery1.RecordCount-1 do
     begin
       for c:=0 to ADOQuery1.FieldCount-1 do
       //for c:=0 to 6 do
         begin
           row:=r+3;
           col:=c+1;
           sheet.Cells[row,col]:=ADOQuery1.Fields[c].AsString;
         end;
       ADOQuery1.Next;
     end;
  // XApp.WorkSheets['Sheet1'].Range['A1:AA1'].Font.Bold:=True;
  // XApp.WorkSheets['Sheet1'].Range['A1:F1'].Borders.LineStyle :=10;
  // XApp.WorkSheets['Sheet1'].Range['A3:F'+inttostr(ADOQuery1.RecordCount+2)].Borders.LineStyle :=1;
   XApp.WorkSheets['Sheet1'].Columns[1].ColumnWidth:=10;
   XApp.WorkSheets['Sheet1'].Columns[2].ColumnWidth:=25;
   XApp.WorkSheets['Sheet1'].Columns[3].ColumnWidth:=12;
   XApp.WorkSheets['Sheet1'].Columns[4].ColumnWidth:=12;
   XApp.WorkSheets['Sheet1'].Columns[5].ColumnWidth:=12;
   XApp.WorkSheets['Sheet1'].Columns[6].ColumnWidth:=12;
   //////////////
end;

procedure TForm10.Button16Click(Sender: TObject);
var
   XApp:Variant;
   sheet:Variant;
   r,c:Integer;
   row,col:Integer;
   filName:Integer;
   q:Integer;
begin
   XApp:=CreateOleObject('Excel.Application');
   XApp.Visible:=true;
   XApp.WorkBooks.Add(-4167);
   XApp.WorkBooks[1].WorkSheets[1].Name:='Sheet1';
   sheet:=XApp.WorkBooks[1].WorkSheets['Sheet1'];
   //For filName:=0 to BlersitQarkullimq.FieldCount-1 do
   For filName:=0 to 8 do
      begin
       q:=filName+1;
       sheet.Cells[1,q]:=BlersitQarkullimq.Fields[filName].FieldName;
      end;
   //////////
   For r:=0 to BlersitQarkullimq.RecordCount-1 do
     begin
       //for c:=0 to BlersitQarkullimq.FieldCount-1 do
      for c:=0 to 8 do
         begin
           row:=r+3;
           col:=c+1;
           sheet.Cells[row,col]:=BlersitQarkullimq.Fields[c].AsString;
         end;
       BlersitQarkullimq.Next;
     end;
  // XApp.WorkSheets['Sheet1'].Range['A1:AA1'].Font.Bold:=True;
  // XApp.WorkSheets['Sheet1'].Range['A1:F1'].Borders.LineStyle :=10;
  // XApp.WorkSheets['Sheet1'].Range['A3:F'+inttostr(ADOQuery1.RecordCount+2)].Borders.LineStyle :=1;
   XApp.WorkSheets['Sheet1'].Columns[1].ColumnWidth:=10;
   XApp.WorkSheets['Sheet1'].Columns[2].ColumnWidth:=20;
   XApp.WorkSheets['Sheet1'].Columns[3].ColumnWidth:=9;
   XApp.WorkSheets['Sheet1'].Columns[4].ColumnWidth:=10;
   XApp.WorkSheets['Sheet1'].Columns[5].ColumnWidth:=10;
   XApp.WorkSheets['Sheet1'].Columns[6].ColumnWidth:=10;
   XApp.WorkSheets['Sheet1'].Columns[7].ColumnWidth:=10;
   XApp.WorkSheets['Sheet1'].Columns[8].ColumnWidth:=10;
   XApp.WorkSheets['Sheet1'].Columns[9].ColumnWidth:=10;
   //////////////
end;

procedure TForm10.Button17Click(Sender: TObject);
var
   XApp:Variant;
   sheet:Variant;
   r,c:Integer;
   row,col:Integer;
   filName:Integer;
   q:Integer;
Begin
   ADOTable1.First;
   XApp:=CreateOleObject('Excel.Application');
   XApp.Visible:=true;
   XApp.WorkBooks.Add(-4167);
   XApp.WorkBooks[1].WorkSheets[1].Name:='Sheet1';
   sheet:=XApp.WorkBooks[1].WorkSheets['Sheet1'];
   For filName:=0 to ADOTable1.FieldCount-1 do
      begin
       q:=filName+1;
       sheet.Cells[1,q]:=ADOTable1.Fields[filName].FieldName;
      end;
   //////////
   For r:=0 to ADOTable1.RecordCount-1 do
     begin
       for c:=0 to ADOTable1.FieldCount-1 do
       //for c:=0 to 6 do
         begin
           row:=r+3;
           col:=c+1;
           sheet.Cells[row,col]:=ADOTable1.Fields[c].AsVariant;
         end;
       ADOTable1.Next;
     end;
  { XApp.WorkSheets['Sheet1'].Range['A1:AA1'].Font.Bold:=True;
   XApp.WorkSheets['Sheet1'].Range['A1:F1'].Borders.LineStyle :=10;
   XApp.WorkSheets['Sheet1'].Range['A3:F'+inttostr(ADOTable1.RecordCount+2)].Borders.LineStyle :=1;}
   XApp.WorkSheets['Sheet1'].Columns[1].ColumnWidth:=8;
   XApp.WorkSheets['Sheet1'].Columns[2].ColumnWidth:=25;
   XApp.WorkSheets['Sheet1'].Columns[3].ColumnWidth:=10;
   XApp.WorkSheets['Sheet1'].Columns[4].ColumnWidth:=15;
   XApp.WorkSheets['Sheet1'].Columns[5].ColumnWidth:=12;
   XApp.WorkSheets['Sheet1'].Columns[6].ColumnWidth:=12;
   XApp.WorkSheets['Sheet1'].Columns[7].ColumnWidth:=14;
   XApp.WorkSheets['Sheet1'].Columns[7].ColumnWidth:=0;
   XApp.WorkSheets['Sheet1'].Columns[8].ColumnWidth:=12;
   XApp.WorkSheets['Sheet1'].Columns[9].ColumnWidth:=0;
   XApp.WorkSheets['Sheet1'].Columns[10].ColumnWidth:=14;
   //////////////
   ADOTable1.First;
end;

procedure TForm10.TabSheet2Enter(Sender: TObject);
begin
  AdoQuery1.Close;
  AdoQuery1.Open;
  If ADOTable1.RecordCount>0 then
     begin
       If SifBlers>'0' then
          begin
            AdoQuery1.Locate('Nr_bleresit',SifBlers,[LopartialKey]);
          end;
     end;
  SalBlerQuery.Close;
  SalBlerQuery.Open;
end;

end.


