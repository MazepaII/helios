//���������� ��������� ���������
//����������� �����: Ltlvfpfqbpfzws@yandex.ru

unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, math, XPMaN, plmn, ComObj, Menus, Mask, jpeg, ExtCtrls,
  ActiveX, Printers, ComCtrls, ToolWin, scale ;
    const
/// ������� ����� � ������
  xlDiagonalDown = 5;
  xlDiagonalUp = 6;
  xlEdgeBottom = 9;
  xlEdgeLeft = 7;
  xlEdgeRight = 10;
  xlEdgeTop = 8;
  xlInsideHorizontal = 12;
  xlInsideVertical = 11;

  ScreenWidth: Integer = 1024; {� ������������ ���� ����� � ������ 800x600.}
  ScreenHeight: Integer = 1280;

type
  TForm1 = class(TForm)
    Timer4: TTimer;
    Timer3: TTimer;
    Timer2: TTimer;
    Timer1: TTimer;
    SaveDialog2: TSaveDialog;
    SaveDialog1: TSaveDialog;
    PrintDialog1: TPrintDialog;
    OpenDialog1: TOpenDialog;
    PanelTityl: TPanel;
    Image7: TImage;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    PanelOcnovn: TPanel;
    Image1: TImage;
    Label9: TLabel;
    Label8: TLabel;
    Label6: TLabel;
    Label5: TLabel;
    Panel5: TPanel;
    Label7: TLabel;
    CheckBox8: TCheckBox;
    Panel4: TPanel;
    Label4: TLabel;
    CheckBox1: TCheckBox;
    Panel3: TPanel;
    Label3: TLabel;
    CheckBox3: TCheckBox;
    Panel2: TPanel;
    Label1: TLabel;
    CheckBox2: TCheckBox;
    Panel1: TPanel;
    Label2: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Memo1: TMemo;
    EditTime: TEdit;
    Edit_tne: TEdit;
    Edit_tnb: TEdit;
    Edit_t_Bx: TEdit;
    Edit_t_Bix: TEdit;
    Edit_Stp: TEdit;
    Edit_Ssh: TEdit;
    Edit_S_shH_ak: TEdit;
    Edit_Pne: TEdit;
    Edit_P1: TEdit;
    Edit_ntp: TEdit;
    Edit_dsh: TEdit;
    Edit_Dpac: TEdit;
    Edit_d_shH_ak: TEdit;
    Edit_d_H: TEdit;
    Edit_btp: TEdit;
    Edit_b_shH_nn: TEdit;
    Edit_b_shH_ak: TEdit;
    ComboBox2: TComboBox;
    ComboBox1: TComboBox;
    CheckBox9: TCheckBox;
    CheckBox7: TCheckBox;
    CheckBox6: TCheckBox;
    CheckBox5: TCheckBox;
    CheckBox10: TCheckBox;
    Button2: TButton;
    MainMenu1: TMainMenu;
    N1: TMenuItem;
    N14: TMenuItem;
    N2: TMenuItem;
    N8: TMenuItem;
    N5: TMenuItem;
    N6: TMenuItem;
    N9: TMenuItem;
    N10: TMenuItem;
    N7: TMenuItem;
    N11: TMenuItem;
    N12: TMenuItem;
    N13: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    TabSheet2: TTabSheet;
    N15: TMenuItem;
    N16: TMenuItem;
    Timer5: TTimer;
    CheckBox4: TCheckBox;
    Image6: TImage;
    Image5: TImage;
    Image4: TImage;
    Image3: TImage;
    Image2: TImage;
    Image8: TImage;
    Image9: TImage;
    Image10: TImage;
    Shape2: TShape;
    N17: TMenuItem;
    Lite: TMenuItem;
    Full: TMenuItem;
    procedure Button2Click(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure N4Click(Sender: TObject);
    procedure ComboBox2Change(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure CheckBox4Click(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure Timer2Timer(Sender: TObject);
    procedure Timer3Timer(Sender: TObject);
    procedure CheckBox2Click(Sender: TObject);
    procedure CheckBox3Cick(Sender: TObject);
    procedure Image2Click(Sender: TObject);
    procedure Edit_DpacKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure Edit_PneKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure Edit_tneKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure Edit_tnbKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure Edit_t_BxKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure Edit_P1KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure Edit_t_BixKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure Edit_d_HKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure Edit_btpKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure Edit_StpKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure Edit_ntpKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure Edit_SshKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure Edit_dshKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure Edit_b_shH_nnKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure Edit_S_shH_akKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure Edit_d_shH_akKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure Edit_b_shH_akKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure Edit_PneKeyPress(Sender: TObject; var Key: Char);
    procedure Edit_tneKeyPress(Sender: TObject; var Key: Char);
    procedure Edit_tnbKeyPress(Sender: TObject; var Key: Char);
    procedure Edit_t_BxKeyPress(Sender: TObject; var Key: Char);
    procedure Edit_P1KeyPress(Sender: TObject; var Key: Char);
    procedure Edit_t_BixKeyPress(Sender: TObject; var Key: Char);
    procedure Edit_d_HKeyPress(Sender: TObject; var Key: Char);
    procedure Edit_btpKeyPress(Sender: TObject; var Key: Char);
    procedure Edit_StpKeyPress(Sender: TObject; var Key: Char);
    procedure Edit_ntpKeyPress(Sender: TObject; var Key: Char);
    procedure Edit_SshKeyPress(Sender: TObject; var Key: Char);
    procedure Edit_dshKeyPress(Sender: TObject; var Key: Char);
    procedure Edit_b_shH_nnKeyPress(Sender: TObject; var Key: Char);
    procedure Edit_S_shH_akKeyPress(Sender: TObject; var Key: Char);
    procedure Edit_d_shH_akKeyPress(Sender: TObject; var Key: Char);
    procedure Edit_b_shH_akKeyPress(Sender: TObject; var Key: Char);
    procedure Edit_DpacKeyPress(Sender: TObject; var Key: Char);
    procedure ComboBox1KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure ComboBox2KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure Label1Click(Sender: TObject);
    procedure Label2Click(Sender: TObject);
    procedure Label3Click(Sender: TObject);
    procedure Label4Click(Sender: TObject);
    procedure Timer4Timer(Sender: TObject);
    procedure Label5Click(Sender: TObject);
    procedure Label6Click(Sender: TObject);
    procedure CheckBox6Click(Sender: TObject);
    procedure CheckBox7Click(Sender: TObject);
    procedure Label7Click(Sender: TObject);
    procedure CheckBox9Click(Sender: TObject);
    procedure CheckBox10Click(Sender: TObject);
    procedure Label8Click(Sender: TObject);
    procedure Label9Click(Sender: TObject);
    procedure N9Click(Sender: TObject);
    procedure N11Click(Sender: TObject);
    procedure N12Click(Sender: TObject);
    procedure N10Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure N13Click(Sender: TObject);
    procedure N14Click(Sender: TObject);
    procedure Image7Click(Sender: TObject);
    procedure N15Click(Sender: TObject);
    procedure N16Click(Sender: TObject);
    procedure Timer5Timer(Sender: TObject);
    procedure FormMouseWheelDown(Sender: TObject; Shift: TShiftState;
      MousePos: TPoint; var Handled: Boolean);
    procedure FormMouseWheelUp(Sender: TObject; Shift: TShiftState;
      MousePos: TPoint; var Handled: Boolean);
    procedure LiteClick(Sender: TObject);
    procedure FullClick(Sender: TObject);

  private
  //���������
function RoundSignificant(num: Extended; col: integer): Extended;  //���������
function CheckExcelInstall: boolean;
function Night:boolean; //����
function Day:boolean;  /// ����
function UpDownSymbol(RangeH: String;StartUp, LengthUp, StartDown,LengthDown: integer ):boolean;
function SelLineEdit(Edit: TEdit;Index: integer): boolean;
function SelLineComboBox(ComboBox: TComboBox;Index: integer): boolean;
function SelLineCheckBox(CheckBox: TCheckBox;Index: integer): boolean;
function SelLineTime(Index: integer): boolean;
function SelLineInterfeis(Index: integer): boolean;
function SelSaveEdit(Edit: TEdit): boolean;
function SelSaveComboBox(ComboBox: TComboBox): boolean;
function SelSaveCheckBox(CheckBox: TCheckBox): boolean;
function ExcelPamka(Pamka: String;UP,Down,Left,Right: integer): integer ;  //����� Excel
function Save_and_Load:boolean;  /// ��������� ��������� (��� ������ ���������)

    { Private declarations }
Procedure CMDialogKey(Var Msg: TWMKey); message CM_DIALOGKEY;  //������ ������� [Tab � Enter]
Procedure WMDisplayChange(var message:TMessage); message WM_DISPLAYCHANGE;  //����������� ��������� �����������

  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  // Excel
    ExcelApp, Workbook, Range, Cell1, Cell2, ArrayData  : Variant;
    TemplateFile : String;
    BeginCol, BeginRow, i, j : integer;
    RowCount, ColCount : integer;
  // �������
    Schetchik: String;
  //��������� ������
    vreme: integer;
  //������ "�"
    Knopka_P: integer;
  //���������� ������
    h_Close, h_Open, h_Women: integer;
  // ������� �� ��������?
    interfeis: String;

implementation

uses Unit2, Unit3, Unit4, Unit5, Unit6, Unit7;


{$R *.dfm}

//������ ������� [Tab � Enter]
procedure TForm1.CMDialogKey(var Msg: TWMKey);
begin
 Msg.Result := 0
end;

//����������� ��������� �����������
procedure TForm1.WMDisplayChange(var message: TMessage);
var   FullProgPath: PChar;
begin
inherited;
Application.Messagebox('���������� ��������� �����������. ��������� ����� ������������','����������', MB_ICONASTERISK or mb_ok) ;
FullProgPath := PChar(Application.ExeName);
// ShowWindow(Form1.handle,SW_HIDE);
WinExec(FullProgPath, SW_SHOW); // Or better use the CreateProcess function
Application.Terminate; // or: Close;
end;

//���������� �����       ([([([(((���������)]))])])
function TForm1.RoundSignificant(num: Extended; col: integer): Extended;
var
  counter, MaxValue, MinValue, PreSign: integer;
  operand: Extended;
begin
  if (col <= 0) or (num = 0)
    then
      begin
        result := 0;
        Exit;
      end;
  try
    MaxValue := Trunc(IntPower(10, col));
  except
    result := num;
    Exit;
  end;
  MinValue := MaxValue div 10;
  counter := 0;
  PreSign := Sign(num);
  operand := Abs(num);
  while operand <= MinValue do
    begin
      operand := operand * 10;
      counter := counter + 1;
    end;
  while operand > MaxValue do
    begin
      operand := operand / 10;
      counter := counter - 1;
    end;
  result := Round(operand) / IntPower(10, counter) * PreSign;
end;

//���������� �� Microsoft Office (Excel)   ([([([(((���������)]))])])
function TForm1.CheckExcelInstall: boolean;
var
  ClassID: TCLSID;
  Rez : HRESULT;
begin
// ���� CLSID OLE-�������
  Rez := CLSIDFromProgID(PWideChar(WideString('Excel.Application')), ClassID);
  if Rez = S_OK then  // ������ ������
    Result := true
  else
     begin
    Result := false;
     end;
end;

//����            ([([([(((���������)]))])])
function TForm1.Night: boolean;
begin
   //�������
   EditTime.Font.Color:=clWhite ;
   EditTime.Color:=clWhite ;
Edit_Dpac.Color:=RGB(3,12,29) ;
Edit_Pne.Color:=RGB(3,12,29) ;
Edit_tne.Color:=RGB(3,12,29) ;
Edit_tnb.Color:=RGB(3,12,29) ;
Edit_t_Bx.Color:=RGB(3,12,29) ;
Edit_P1.Color:=RGB(3,12,29) ;
Edit_t_Bix.Color:=RGB(3,12,29) ;
ComboBox1.Color:=RGB(3,12,29) ;
Edit_d_H.Color:=RGB(9,25,25) ;
Edit_btp.Color:=RGB(1,8,27) ;
Edit_Stp.Color:=RGB(9,19,28) ;
Edit_ntp.Color:=RGB(1,8,27) ;
ComboBox2.Color:=RGB(2,9,28) ;
Edit_Ssh.Color:=RGB(0,0,28) ;
Edit_dsh.Color:=RGB(18,29,25) ;
Edit_b_shH_nn.Color:=RGB(4,11,29) ;
Edit_S_shH_ak.Color:=RGB(2,11,26) ;
Edit_d_shH_ak.Color:=RGB(4,16,28) ;
Edit_b_shH_ak.Color:=RGB(28,28,26) ;

Edit_Dpac.Font.Color:=RGB(219,219,217) ;
Edit_Pne.Font.Color:=Edit_Dpac.Font.Color ;
Edit_tne.Font.Color:=Edit_Dpac.Font.Color ;
Edit_tnb.Font.Color:=Edit_Dpac.Font.Color ;
Edit_t_Bx.Font.Color:=Edit_Dpac.Font.Color ;
Edit_P1.Font.Color:=Edit_Dpac.Font.Color ;
Edit_t_Bix.Font.Color:=Edit_Dpac.Font.Color ;
ComboBox1.Font.Color:=Edit_Dpac.Font.Color ;
Edit_d_H.Font.Color:=Edit_Dpac.Font.Color ;
Edit_btp.Font.Color:=Edit_Dpac.Font.Color ;
Edit_Stp.Font.Color:=Edit_Dpac.Font.Color ;
Edit_ntp.Font.Color:=Edit_Dpac.Font.Color ;
ComboBox2.Font.Color:=Edit_Dpac.Font.Color ;
Edit_Ssh.Font.Color:=Edit_Dpac.Font.Color ;
Edit_dsh.Font.Color:=Edit_Dpac.Font.Color ;
Edit_b_shH_nn.Font.Color:=Edit_Dpac.Font.Color ;
Edit_S_shH_ak.Font.Color:=Edit_Dpac.Font.Color ;
Edit_d_shH_ak.Font.Color:=Edit_Dpac.Font.Color ;
Edit_b_shH_ak.Font.Color:=Edit_Dpac.Font.Color ;
Panel1.Visible:=False ;
Panel2.Visible:=False ;
Panel3.Visible:=False ;
Panel4.Visible:=False ;
Panel5.Visible:=False ;
CheckBox1.Visible:=False ;
CheckBox3.Visible:=False ;
CheckBox4.Visible:=False ;
CheckBox6.Visible:=False ;
CheckBox7.Visible:=False ;
  Label1.Font.Color:=clBlack;
  Label2.Font.Color:=clBlack;
  Label7.Font.Color:=clBlack;
  Label10.Font.Color:=clBlack;
  Label11.Font.Color:=clBlack;
  Label3.Font.Color:=clWhite;
  Label4.Font.Color:=clWhite;
  Label5.Font.Color:=clWhite;
  Label6.Font.Color:=clWhite;
  Label8.Font.Color:=clWhite;
  Label9.Font.Color:=clWhite;
Panel1.Visible:=True ;
Panel2.Visible:=True ;
Panel3.Visible:=True ;
Panel4.Visible:=True ;
Panel5.Visible:=True ;
CheckBox1.Visible:=True  ;
CheckBox3.Visible:=True  ;
CheckBox4.Visible:=True  ;
CheckBox6.Visible:=True  ;
CheckBox7.Visible:=True  ;
end;

/// ����          ([([([(((���������)]))])])
function TForm1.Day: boolean;
begin
  Shape2.Visible:=False;
     //�������
     EditTime.Font.Color:=clBlack ;
     EditTime.Color:=clBlack ;
  Edit_Dpac.Color:=RGB(255,244,228) ;
  Edit_Pne.Color:=RGB(252,243,228) ;
  Edit_tne.Color:=RGB(255,255,243) ;
  Edit_tnb.Color:=RGB(250,237,228) ;
  Edit_t_Bx.Color:=RGB(252,253,239) ;
  Edit_P1.Color:=RGB(249,242,224) ;
  Edit_t_Bix.Color:=RGB(250,236,227) ;
  ComboBox1.Color:=RGB(254,245,228) ;
  Edit_d_H.Color:=RGB(254,242,228) ;
  Edit_btp.Color:=RGB(243,229,226) ;
  Edit_Stp.Color:=RGB(247,234,226) ;
  Edit_ntp.Color:=RGB(255,249,227) ;
  ComboBox2.Color:=RGB(254,245,228) ;
  Edit_Ssh.Color:=RGB(255,255,229) ;
  Edit_dsh.Color:=RGB(233,227,227) ;
  Edit_b_shH_nn.Color:=RGB(243,232,226) ;
  Edit_S_shH_ak.Color:=RGB(250,240,231) ;
  Edit_d_shH_ak.Color:=RGB(255,248,229) ;
  Edit_b_shH_ak.Color:=RGB(228,228,228) ;
Edit_Dpac.Font.Color:=RGB(0,0,0) ;
Edit_Pne.Font.Color:=RGB(0,0,0) ;
Edit_tne.Font.Color:=RGB(0,0,0) ;
Edit_tnb.Font.Color:=RGB(0,0,0) ;
Edit_t_Bx.Font.Color:=RGB(0,0,0) ;
Edit_P1.Font.Color:=RGB(0,0,0) ;
Edit_t_Bix.Font.Color:=RGB(0,0,0) ;
ComboBox1.Font.Color:=RGB(0,0,0) ;
Edit_d_H.Font.Color:=RGB(0,0,0) ;
Edit_btp.Font.Color:=RGB(0,0,0) ;
Edit_Stp.Font.Color:=RGB(0,0,0) ;
Edit_ntp.Font.Color:=RGB(0,0,0) ;
ComboBox2.Font.Color:=RGB(0,0,0) ;
Edit_Ssh.Font.Color:=RGB(0,0,0) ;
Edit_dsh.Font.Color:=RGB(0,0,0) ;
Edit_b_shH_nn.Font.Color:=RGB(0,0,0) ;
Edit_S_shH_ak.Font.Color:=RGB(0,0,0) ;
Edit_d_shH_ak.Font.Color:=RGB(0,0,0) ;
Edit_b_shH_ak.Font.Color:=RGB(0,0,0) ;
Panel1.Visible:=False ;
Panel2.Visible:=False ;
Panel3.Visible:=False ;
Panel4.Visible:=False ;
Panel5.Visible:=False ;
CheckBox1.Visible:=False ;
CheckBox3.Visible:=False ;
CheckBox4.Visible:=False ;
CheckBox6.Visible:=False ;
CheckBox7.Visible:=False ;
  Label1.Font.Color:=clWhite;
  Label2.Font.Color:=clWhite;
  Label7.Font.Color:=clWhite;
  Label10.Font.Color:=clWhite;
  Label11.Font.Color:=clWhite;
  Label3.Font.Color:=clBlack;
  Label4.Font.Color:=clBlack;
  Label5.Font.Color:=clBlack;
  Label6.Font.Color:=clBlack;
  Label8.Font.Color:=clBlack;
  Label9.Font.Color:=clBlack;
Panel1.Visible:=True ;
Panel2.Visible:=True ;
Panel3.Visible:=True ;
Panel4.Visible:=True ;
Panel5.Visible:=True ;
CheckBox1.Visible:=True  ;
CheckBox3.Visible:=True  ;
CheckBox4.Visible:=True  ;
CheckBox6.Visible:=True  ;
CheckBox7.Visible:=True  ;
   Form4.close;
   form4.ProgressBar1.position:=0 ;
   Shape2.Visible:=False;
end;

// ��������� ��������� (��� ������ ���������)  ([([([(((���������)]))])])
function TForm1.Save_and_Load:boolean;
var SR:TSearchRec; // ��������� ����������
    temp:Word;     // ������ � ���������� ������
begin
FindFirst('���������.ini',faAnyFile,SR); // ������� ������� ������ � ������ ������
if SR.name <> '' then   //���� ������
  begin                                     //���������
   FindClose(SR); // ��������� �����
   Memo1.Clear ;
   Memo1.Lines.LoadFromFile(ExtractFilePath( Application.ExeName )+'���������.ini');
   SelLineEdit(Edit_Dpac,2) ;       // Dpac
   SelLineEdit(Edit_Pne,5) ;        // Pne
   SelLineEdit(Edit_tne,8) ;        // tne
   SelLineEdit(Edit_tnb,11) ;       // tnb
   SelLineEdit(Edit_t_Bx,14) ;      // t_Bx
   SelLineEdit(Edit_P1,17) ;        // P1
   SelLineEdit(Edit_t_Bix,20) ;     // t_Bix
   SelLineComboBox(ComboBox1,23) ;  // �������� �����
   SelLineEdit(Edit_d_H,26) ;       // d_H
   SelLineEdit(Edit_btp,29) ;       // b_tp
   SelLineEdit(Edit_Stp,32) ;       // S_tp
   SelLineEdit(Edit_nTP,35) ;       // n_TP
   SelLineComboBox(ComboBox2,38) ;  // ������ �� �� ��������� ����?
   SelLineEdit(Edit_Ssh,41) ;       // S_sh_nn
   SelLineEdit(Edit_dsh,44) ;       // d_sh_nn
   SelLineEdit(Edit_b_shH_nn,47) ;  // b_shH_nn
   SelLineEdit(Edit_S_shH_ak,50) ;  // S_shH_ak
   SelLineEdit(Edit_d_shH_ak,53) ;  // d_shH_ak
   SelLineEdit(Edit_b_shH_ak,56) ;  // b_shH_ak
   //�����
   SelLineCheckBox(CheckBox8,60) ; //  ���������
   SelLineCheckBox(CheckBox2,63) ; //  ������
   SelLineCheckBox(CheckBox4,66) ; //  ������ >
   SelLineCheckBox(CheckBox3,69) ; //  �������
   SelLineCheckBox(CheckBox6,78) ; //  Excel
   SelLineCheckBox(CheckBox7,81) ; //  �������
   SelLineCheckBox(CheckBox9,72) ; //  �����������
   SelLineCheckBox(CheckBox10,75); //  ������
   SelLineCheckBox(CheckBox1,84) ; //  ������������
   //�������
   SelLineTime(86); //��������� � "Schetchik" �����
   Schetchik:=FloatToStr(StrToFloat(Schetchik)+1) ;
   Memo1.Lines[85]:=('�������');
   Memo1.Lines[86]:=(Schetchik);
   EditTime.Text:=Schetchik ;
   //���������
   SelLineInterfeis(88); //��������� � "Interfeis" �����
   Memo1.Lines[87]:=('"�������" ��� � ��� ���������?');
   Memo1.Lines[88]:=(Interfeis);
   //���������� "���������.ini"
   Memo1.Lines.SaveToFile(ExtractFilePath( Application.ExeName )+'���������.ini');
  end
 else  //���� �� ������
  begin                                      //���������
   FindClose(SR); // ��������� �����
    Memo1.Clear ;
    Memo1.Lines.Add('D���  (���������������������� ����� �������) [��/�]');
    SelSaveEdit(Edit_Dpac);
    Memo1.Lines.Add('���   (�������� ����������� ����)            [���]');
    SelSaveEdit(Edit_Pne);
    Memo1.Lines.Add('t��   (����������� ����������� ����)         [o�]');
    SelSaveEdit(Edit_tne);
    Memo1.Lines.Add('t��   (����������� ����������� ����)         [o�]');
    SelSaveEdit(Edit_tnb);
    Memo1.Lines.Add('t��   (����������� �� �� ����� � ��)         [o�]');
    SelSaveEdit(Edit_t_Bx);
    Memo1.Lines.Add('P1    (�������� ������� �������)             [���]');
    SelSaveEdit(Edit_P1);
    Memo1.Lines.Add('t���  (����������� �� �� ������ �� ��)       [o�]');
    SelSaveEdit(Edit_t_Bix);
    Memo1.Lines.Add('�������� �����');
    SelSaveComboBox(ComboBox1);
    Memo1.Lines.Add('d�    (�������� ������� �����)               [��]');
    SelSaveEdit(Edit_d_H);
    Memo1.Lines.Add('b��   (������� ������ ����)                  [��]');
    SelSaveEdit(Edit_btp);
    Memo1.Lines.Add('S��   (��� ������� �������)                  [��]');
    SelSaveEdit(Edit_Stp);
    Memo1.Lines.Add('n��   (����� ������������ ����)              [��]');
    SelSaveEdit(Edit_ntp);
    Memo1.Lines.Add('������ �� �� ��������� ����?');
    SelSaveComboBox(ComboBox2);
    Memo1.Lines.Add('S����  (��� ����� �� �� �������)             [��]');
    SelSaveEdit(Edit_Ssh);
    Memo1.Lines.Add('d����  (�������  ����� �� �� �������)        [��]');
    SelSaveEdit(Edit_dsh);
    Memo1.Lines.Add('b����  (������� ������� ����� �� �� �������) [��]');
    SelSaveEdit(Edit_b_shH_nn);
    Memo1.Lines.Add('S����  (��� ����� �� �� �������)             [��]');
    SelSaveEdit(Edit_S_shH_ak);
    Memo1.Lines.Add('d����  (�������  ����� ��  ��  �������)      [��]');
    SelSaveEdit(Edit_d_shH_ak);
    Memo1.Lines.Add('b����  (������� ������� ����� �� �� �������) [��]');
    SelSaveEdit(Edit_b_shH_ak);
    Memo1.Lines.Add('������');
    Memo1.Lines.Add('���������');
    SelSaveCheckBox(CheckBox8) ;
    Memo1.Lines.Add('������');
    SelSaveCheckBox(CheckBox2) ;
    Memo1.Lines.Add('������ >  < ������');
    SelSaveCheckBox(CheckBox4) ;
    Memo1.Lines.Add('�������');
    SelSaveCheckBox(CheckBox3) ;
    Memo1.Lines.Add('�����������');
    SelSaveCheckBox(CheckBox9) ;
    Memo1.Lines.Add('������');
    SelSaveCheckBox(CheckBox10) ;
    Memo1.Lines.Add('Excel');
    SelSaveCheckBox(CheckBox6) ;
    Memo1.Lines.Add('�������');
    SelSaveCheckBox(CheckBox7) ;
    Memo1.Lines.Add('������������');
    SelSaveCheckBox(CheckBox1) ;
    //�������
    Memo1.Lines.Add('�������');
    Schetchik:='1';
    Memo1.Lines.Add(Schetchik);
    EditTime.Text:=Schetchik ;
    //������ �����������
      Application.Messagebox('��� ������������ ��������� "������". ��������� ������� ��������������.','�����������', MB_ICONASTERISK or mb_ok) ;
    //���������
      // ������� �� ��������?
           temp:=Application.Messagebox('"�������" ��� � ��� ���������? (������ �� ������� ����������.)','���������', MB_ICONQUESTION+MB_YESNO+MB_DEFBUTTON1);
           case temp of
           idYES: interfeis:='�������' ;
           idNO: interfeis:='������' ;
           end;
    Memo1.Lines.Add('"�������" ��� � ��� ���������?');
    Memo1.Lines.Add(Interfeis);
    //���������� "���������.ini"
    Memo1.Lines.SaveToFile(ExtractFilePath( Application.ExeName )+'���������.ini');
  end;
/// ���������� ��� ������� (������� �� ��������)
//�����������
if Schetchik='10'      then    Application.Messagebox('����������� ��� � "���������" �������� ���������. ����� ��� ������ ��� ���������� �����.','�����������', MB_ICONASTERISK or mb_ok) ;
if Schetchik='25'      then    Application.Messagebox('����������� ��� � "����������" �������� ���������. ��� �� ��������� ������.','�����������', MB_ICONASTERISK or mb_ok) ;
if Schetchik='50'      then    Application.Messagebox('����������� ��� � "�������" �������� ���������. ������ ���, ���� ���� ����� ���� �������.','�����������', MB_ICONASTERISK or mb_ok) ;
if Schetchik='75'      then    Application.Messagebox('����������� ��� � "��������" �������� ���������. ������ ���� ���� �� ���� �� �� ����� ������.','�����������', MB_ICONASTERISK or mb_ok) ;
if Schetchik='100'     then    Application.Messagebox('����������� ��� � "�������" �������� ���������. �����, �������, �������, ����. ����� �������� ��� �� ����.','�����������', MB_ICONASTERISK or mb_ok) ;
if Schetchik='250'     then    Application.Messagebox('����������� ��� � 250 �������� ���������. ������ ��� ������� � �������� ����� ��� � �� ��������� �� ���, �� ���.','�����������', MB_ICONASTERISK or mb_ok) ;
if Schetchik='500'     then    Application.Messagebox('����������� ��� � 500 �������� ���������. ������ ������� � �����. ����� ����� ������� ������, ����� ����� ������� �����.','�����������', MB_ICONASTERISK or mb_ok) ;
if Schetchik='750'     then    Application.Messagebox('����������� ��� � 750 �������� ���������. ������ �������, ������� ������, ���� ��� ���������� ��� ������.','�����������', MB_ICONASTERISK or mb_ok) ;
if Schetchik='1000'    then    Application.Messagebox('����������� ��� � 1.000 �������� ���������. ������ ���� � ����� � ���� � ��������� ��������� ��������','�����������', MB_ICONASTERISK or mb_ok) ;
if Schetchik='5000'    then    Application.Messagebox('����������� ��� � 5.000 �������� ���������. ������, � ������ �������, ���������� ������� ���������!','�����������', MB_ICONASTERISK or mb_ok) ;
if Schetchik='10000'   then    Application.Messagebox('����������� ��� � 10.000 �������� ���������. ������ ������ � ������, ������ ���, ��� � ����� �����.','�����������', MB_ICONASTERISK or mb_ok) ;
if Schetchik='25000'   then    Application.Messagebox('����������� ��� � 25.000 �������� ���������. ��� ��������, ����� ����������, ��� ������� - ����� ����������.','�����������', MB_ICONASTERISK or mb_ok) ;
if Schetchik='50000'   then    Application.Messagebox('����������� ��� � 50.000 �������� ���������. ������ ��� ������� �����, ���� ���� ����� ���� ������.','�����������', MB_ICONASTERISK or mb_ok) ;
if Schetchik='75000'   then    Application.Messagebox('����������� ��� � 75.000 �������� ���������. ������ ������ � ������, ������ ���, ��� � ����� �����.','�����������', MB_ICONASTERISK or mb_ok) ;
if Schetchik='100000'  then    Application.Messagebox('����������� ��� � 100.000 �������� ���������. ������ ��� �������, ������ �����, ������ ������� � ������� � �������.','�����������', MB_ICONASTERISK or mb_ok) ;
if Schetchik='250000'  then    Application.Messagebox('����������� ��� � 250.000 �������� ���������. ������ ���� ���� �� ���� �� �� ����� ������, ���� �� ����, �� ���� ��� � ���� �� ������.','�����������', MB_ICONASTERISK or mb_ok) ;
if Schetchik='500000'  then    Application.Messagebox('����������� ��� � 500.000 �������� ���������. ������ ������ � ������, ���������, ���� ��� �����.','�����������', MB_ICONASTERISK or mb_ok) ;
if Schetchik='750000'  then    Application.Messagebox('����������� ��� � 750.000 �������� ���������. ���� ����� ������, � �������� ������, ����� ����� ����� �������� �����.','�����������', MB_ICONASTERISK or mb_ok) ;
if Schetchik='1000000' then
   if Application.Messagebox('����������� ��� � 1.000.000 �������� ���������. ����������� ��� ��������.  ������ ����� � ��������.','�����������', MB_ICONASTERISK or mb_ok) = mrOk    then
     begin
          if Application.Messagebox('��������� ��� ������  ������ ������, ��������.','�����������', MB_ICONASTERISK or mb_ok) = mrOk    then
          if Application.Messagebox('� ����� ��� ���� ��� ����  ������� ���� � �����.','�����������', MB_ICONASTERISK or mb_ok) = mrOk    then
          if Application.Messagebox('� ��� �� ������ �� �� ���� �  ��� ������ ���������� ����.','�����������', MB_ICONASTERISK or mb_ok) = mrOk    then
          if Application.Messagebox('��� ������ ��������� ������!  � ��� �� �������� ������','�����������', MB_ICONASTERISK or mb_ok) = mrOk    then
          if Application.Messagebox('����������� ������:  "������! ������!"','�����������', MB_ICONASTERISK or mb_ok) = mrOk    then
     end;
    //������ � �����
Edit_Dpac.setfocus  ;
end;


//�����������  ����������� �������       ([([([(((���������)]))])])
function TForm1.UpDownSymbol(RangeH: String; StartUp, LengthUp, StartDown,LengthDown: integer): boolean;
begin
// ����������� �������
if (StartDown<>777) and (LengthDown<>777) then
ExcelApp.Range[RangeH].Characters[StartDown, LengthDown].Font.Subscript:= True  ;
// ����������� �������
if (StartUp<>777) and (LengthUp<>777) then
ExcelApp.Range[RangeH].Characters[StartUp, LengthUp].Font.Superscript:= True  ;
end;

/// ���������   ([([([(((���������)]))])])

// Edit       (���������)                     ([([([(((���������)]))])])
function TForm1.SelSaveEdit(Edit: TEdit): boolean;
begin
with Edit do
   begin
        if Enabled=true then
           Memo1.Lines.Add('���')
       else
           Memo1.Lines.Add('����');
     Memo1.Lines.Add(Edit.Text);
       
  end;
end;
// ComboBox       (���������)                ([([([(((���������)]))])])
function TForm1.SelSaveComboBox(ComboBox: TComboBox): boolean;
begin
with ComboBox do
   begin
       if Enabled=true then
           Memo1.Lines.Add('���')
       else
           Memo1.Lines.Add('����');
       Memo1.Lines.Add(ComboBox.Text);
   end;
end;
// CheckBox       (���������)                ([([([(((���������)]))])])
function TForm1.SelSaveCheckBox(CheckBox : TCheckBox ): boolean;
begin
with CheckBox  do
   begin
       if Enabled=true then
           Memo1.Lines.Add('���')
       else
           Memo1.Lines.Add('����');
       if Checked=true then
           Memo1.Lines.Add('���')
       else
           Memo1.Lines.Add('���');
  end;
end;
// Edit       (��������)   (�����)             ([([([(((���������)]))])])
function TForm1.SelLineEdit(Edit: TEdit;Index: integer): boolean;
begin
  with Memo1 do
    begin
      // (�����)
      SelStart := Perform(EM_LINEINDEX, Index, 0);
      SelLength := Length(Lines[Index]);
      SetFocus;
      with Edit do
          Text:=Memo1.SelText ;
      ///(����������)
      Index:=Index-1 ;
      SelStart := Perform(EM_LINEINDEX, Index, 0);
      SelLength := Length(Lines[Index]);
      SetFocus;
      with Edit do
       if Memo1.SelText='���' then
           Enabled:=true
       else
           if Memo1.SelText='����' then
              Enabled:=False ;
    end;
end;
// ComboBox   (��������)   (�����)              ([([([(((���������)]))])])
function TForm1.SelLineComboBox(ComboBox: TComboBox;Index: integer): boolean;
begin
  with Memo1 do
    begin
      // (�����)
      SelStart := Perform(EM_LINEINDEX, Index, 0);
      SelLength := Length(Lines[Index]);
      SetFocus;
      with ComboBox do
        if  Memo1.SelText='�����'  then
             ItemIndex:=0
          else
               if  Memo1.SelText='�����' then
                  ItemIndex:=1
               else
                   if  Memo1.SelText='��' then
                     begin
                      ItemIndex:=0;
                     end
                  else
                      if  Memo1.SelText='���' then
                        begin
                         ItemIndex:=1;
                        end;
      ///(����������)
      Index:=Index-1 ;
      SelStart := Perform(EM_LINEINDEX, Index, 0);
      SelLength := Length(Lines[Index]);
      SetFocus;
      with ComboBox do
       if Memo1.SelText='���' then
           Enabled:=true
       else
           if Memo1.SelText='����' then
              Enabled:=False ;
    end;
end;
// CheckBox       (��������)   (�����)            ([([([(((���������)]))])])
function TForm1.SelLineCheckBox(CheckBox : TCheckBox ;Index: integer): boolean;
begin
  with Memo1 do
    begin
      ///(�������)
      SelStart := Perform(EM_LINEINDEX, Index, 0);
      SelLength := Length(Lines[Index]);
      SetFocus;
      with CheckBox do
       if Memo1.SelText='���' then
           Checked:=true
       else
           if Memo1.SelText='���' then
              Checked:=False ;
      ///(����������)
      Index:=Index-1 ;
      SelStart := Perform(EM_LINEINDEX, Index, 0);
      SelLength := Length(Lines[Index]);
      SetFocus;
      with CheckBox do
       if Memo1.SelText='���' then
           Enabled:=true
       else
           if Memo1.SelText='����' then
              Enabled:=False ;
    end;
end;
// �������      (��������)   (�����)            ([([([(((���������)]))])])
function TForm1.SelLineTime(Index: integer): boolean;
begin
  with Memo1 do
    begin
      // (�����)
      SelStart := Perform(EM_LINEINDEX, Index, 0);
      SelLength := Length(Lines[Index]);
      SetFocus;
      Schetchik:=Memo1.SelText ;
    end;
end;

// ���������    (��������)   (�����)            ([([([(((���������)]))])])
function TForm1.SelLineInterfeis(Index: integer): boolean;
begin
  with Memo1 do
    begin
      // (�����)
      SelStart := Perform(EM_LINEINDEX, Index, 0);
      SelLength := Length(Lines[Index]);
      SetFocus;
      interfeis:=Memo1.SelText ;
    end;
end;

//����� Excel   ([([([(((���������)]))])])         ([([([(((���������)]))])])
function TForm1.ExcelPamka(Pamka: String;UP,Down,Left,Right: integer): integer ;
begin
ExcelApp.Range[Pamka].Select;
ExcelApp.Selection.Borders[xlEdgeBottom].Weight := UP;
ExcelApp.Selection.Borders[xlEdgeTop].Weight := Down;
ExcelApp.Selection.Borders[xlEdgeLeft].Weight := Left;
ExcelApp.Selection.Borders[xlEdgeRight].Weight := Right;
end;

///   [[[����   �� ���]]]

procedure TForm1.ComboBox2Change(Sender: TObject);

begin

if ComboBox2.ItemIndex<>0 then
    begin
      Edit_Ssh.Enabled:=False ;
      Edit_dsh.Enabled:=False ;
      Edit_b_shH_nn.Enabled:=False ;
      Edit_S_shH_ak.Enabled:=False ;
      Edit_d_shH_ak.Enabled:=False ;
      Edit_b_shH_ak.Enabled:=False ;
    end
  else
    begin
      Edit_Ssh.Enabled:=true ;
      Edit_dsh.Enabled:=true ;
      Edit_b_shH_nn.Enabled:=true ;
      Edit_S_shH_ak.Enabled:=true ;
      Edit_d_shH_ak.Enabled:=true ;
      Edit_b_shH_ak.Enabled:=true ;
    end;
end;






//������ (���������)
procedure TForm1.Button2Click(Sender: TObject);


var
// ����
  Dpac,Pne,tne,tnb,t_Bx,P1,t_Bix,d_H,btp,Stp,n_TP,S_shH_nn,d_shH_nn,b_shH_nn,S_shH_ak,d_shH_ak,b_shH_ak: Extended;
  MAT_TP,XOT_ShH: String;
//��������� ���������
Rho_CTAL,Rho_MAT,V_MAT: Extended;
// �������
ftp_th,fnpox_th,d_BH,f_BH,fnpox_2,t_cp_TH,KINV_cp_TH,Pr_TH,P_icn,P_ak,P_BX,P_ak_cp,P_icn_cp,P_nn_cp,t_s_icn,t_s_ak,iYY,iY_s,i_nB,ine,t_cp_ak,t_cp_nn,Q_ak,Q_icn,Q_nn,Q_nr,G_TH,i_icn_TH,i_ak_TH,t_icn_TH,t_ak_TH,W_1,d_ek_1,Re_1,Nu_1,a_1,V_ak_cp,f_ak_TP,f_ak,W_ak,Mu_cp_ak,KINV_cp_ak,d_ek_ak,Re_ak,Lambda_ak,Pr_ak,Nu_ak,a_ak,K_ak,Lambda_MAT_nn,Delta_t_ak,F__ak,L_ak,V_nn_cp,Mu_cp_nn,KINV_cp_nn,f_nn_TP,f_nn,d_ek_nn,Re_nn,Lambda_nn,Pr_nn,Nu_nn,a_nn,Lambda_MAT_ak,K_nn,Delta_t_nn,F__nn,L_nn,W_nn,q_1,q_2,Prov_q,a_icn,Lambda_MAT_icn,K_icn,Delta_t_icn,t_cp_icn,F__icn,L_icn,d_a_ak,D_ak,Psi_ak,B_ak,m_ak,d_a_nn,D_nn,Psi_nn,B_nn,m_nn,L_nr,S_k_d,Zeta_1,c_ak,Zeta_ak,Re_icn,d_a_icn,D_icn,Psi_icn,B_icn,m_icn,c_icn,Zeta_icn,V_cp_YY,V_cp_Y,W_0,c_nn,Zeta_nn,Mu_cp_icn,KINV_cp_icn,d_ek_icn,Lambda_MAT,d_cp,d_cp_ak,d_cp_icn,d_cp_nn,t_cp_ct_icn,t_cp_ct_nn,t_cp_ct_ak,t_nn_TH,t_cp_TH_ne,t_cp_TH_icn,t_cp_TH_ak,Lambda_TH_ct,Mu_TH_ct,Pr_TH_ct,Mu_cp_ct_ak,Lambda_ct_ak,Pr_ct_ak,Mu_cp_ct_nn,Lambda_ct_nn,Pr_ct_nn: Extended;
//  ������
Delta_P_1_2,Delta_P_nn_1,Delta_P_icn_1,Delta_P_ak_1,Delta_P_nr_1,Prov_P_nr,Delta_P_nn_2,Delta_P_icn_2,Delta_P_ak_2,Delta_P_nr_2: Extended;
// ��������
V_TH,V_TH_ct,V_ak_cp_ct,V_nn_cp_ct,i_Bx_TH,i_Bix_TH,Lambda_TH,Mu_TH: Extended;
// �����
Start, Length: OleVariant;
// ����� ��������� ������
ClassID: TCLSID;
Rez : HRESULT;
Result  : Variant;
//������ �������
Line: TextFile;
P: integer;
//��������� ������
DirPath, FullFileName : String;


begin
Edit_Dpac .setfocus  ;  //������ �� ��������� � ���� �������� ������ ������� �����
///[[{{{��������� Excel}}}]]
if CheckBox6.Checked=true  then
   begin ///[[{{{��������� Excel}}}]]
          if interfeis='�������' then  //�������� ��������� ����������
           begin
             // ������� ������
              Shape2.Left:=0 ;
              Shape2.Top:=0 ;
              Shape2.Width:=1150 ;
              Shape2.Height:=726 ;
              Shape2.Visible:=true;
             //��������� ��������
                //��������
                Form4.Image1.Visible:=True;
                Form4.ProgressBar1.Top:=31;
                Form4.ProgressBar1.Left:=7;
                //������������
                Form4.Position:=poDefault ;
                Form4.Position:=poMainFormCenter ;
              form4.Show ;
              Night ;  //����
           end
         else
           begin
             //��������� ��������
                //��������
                Form4.Image1.Visible:=False;
                //������������
                Form4.Position:=poDefault ;
                Form4.Position:=poMainFormCenter ;
              form4.Show ;
           end
  end; ///[[{{{��������� Excel}}}]]

/// �������� ���������� �����
//������ ����
  if (Edit_Dpac.Text='') or (Edit_Pne.Text='') or (Edit_tne.Text='') or (Edit_tnb.Text='') or (Edit_t_Bx.Text='') or (Edit_P1.Text='') or (Edit_t_Bix.Text='') or (Edit_d_H.Text='') or (Edit_btp.Text='') or (Edit_ntp.Text='') or (Edit_Stp.Text='') or (Edit_Ssh.Text='') or (Edit_dsh.Text='') or (Edit_b_shH_nn.Text='') or (Edit_S_shH_ak.Text='') or (Edit_d_shH_ak.Text='') or (Edit_b_shH_ak.Text='') then
    begin
      Form4.close;
      Application.Messagebox('���� �� ��������� ����!','������ !!!', mb_iconerror or mb_ok);
      Day ; Exit; /// (���� �� ��������� ����)
    end ;
//�\���� � ���� ����
  if (Edit_Dpac.Text='0') or (Edit_Pne.Text='0') or (Edit_tne.Text='0') or (Edit_tnb.Text='0') or (Edit_t_Bx.Text='0') or (Edit_P1.Text='0') or (Edit_t_Bix.Text='0') or (Edit_d_H.Text='0') or (Edit_btp.Text='0') or (Edit_ntp.Text='0') or (Edit_Stp.Text='0') or (Edit_Ssh.Text='0') or (Edit_dsh.Text='0') or (Edit_b_shH_nn.Text='0') or (Edit_S_shH_ak.Text='0') or (Edit_d_shH_ak.Text='0') or (Edit_b_shH_ak.Text='0') then
    begin
      Form4.close;
      Application.Messagebox('����, �� � � ������ ����','������ !!!', mb_iconerror or mb_ok);
      Day ; Exit; /// (���� �� ��������� ����)
    end ;
form4.ProgressBar1.position:=form4.ProgressBar1.position+5 ;    //((((((())))))

// ����
//1) D���     [���������������������� ����� ������]
Dpac:=strtofloat(Edit_Dpac.Text);
//2) ���      [�������� ����������� ����]
Pne:=strtofloat(Edit_Pne.Text);
//3) t��      [����������� ����������� ����]
tne:=strtofloat(Edit_tne.Text);
//4) t��      [����������� ����������� ����]
tnb:=strtofloat(Edit_tnb.Text);
//5) t��      [����������� �� �� ����� � ��]
t_Bx:=strtofloat(Edit_t_Bx.Text);
//6)P1        [�������� ������� �������]
P1:=strtofloat(Edit_P1.Text);
//7) t���     [����������� �� �� ������ �� ��]
t_Bix:=strtofloat(Edit_t_Bix.Text);
//8) �������� �����
  Case ComboBox1.ItemIndex of
   0: MAT_TP:='�����' ;
   1: MAT_TP:='�����' ;
    end ;
//9) d�       [�������� ������� �����]
d_H:=strtofloat(Edit_d_H.Text);
//10) �.��    [������� ������ ����]
btp:=strtofloat(Edit_btp.Text);
//11) n.��    [���� ����������� ����]
n_TP:=strtofloat(Edit_ntp.Text);
//12) S��     [��� ������� �������]
Stp:=strtofloat(Edit_Stp.Text);

   if ComboBox2.ItemIndex=0 then
      begin
        XOT_ShH:='��' ;
        //14) S�� ��  [��� ����� �� �� �������]
        S_shH_nn:=strtofloat(Edit_Ssh.Text);
        //15) d�� ��  [������� ����� �� �� �������]
        d_shH_nn:=strtofloat(Edit_dsh.Text);
        //16) ��� ��  [������� ������� ����� �� �� �������]
        b_shH_nn:=strtofloat(Edit_b_shH_nn.Text);
        //17) S�� ��  [��� ����� �� �� �������]
        S_shH_ak:=strtofloat(Edit_S_shH_ak.Text);
        //18) d�� ��  [������� ����� �� �� �������]
       d_shH_ak:=strtofloat(Edit_d_shH_ak.Text);
       //19) ��� ��  [������� ������� ����� �� �� �������]
       b_shH_ak:=strtofloat(Edit_b_shH_ak.Text);
      end
   else
    begin
      XOT_ShH:='���' ;
      S_shH_nn:=0;
      d_shH_nn:=0;
      b_shH_nn:=0;
      S_shH_ak:=0;
      d_shH_ak:=0;
      b_shH_ak:=0;
      Edit_Ssh.Enabled:=False ;
      Edit_dsh.Enabled:=False ;
      Edit_b_shH_nn.Enabled:=False ;
      Edit_S_shH_ak.Enabled:=False ;
      Edit_d_shH_ak.Enabled:=False ;
      Edit_b_shH_ak.Enabled:=False ;
    end;



// [[�����]]

//0) ������� [��] � [�]
// ����
//[9] d�       [�������� ������� �����]
d_H:=d_H*power(10,-3);
//[10] �.��    [������� ������ ����]
btp:=btp*power(10,-3);
//[12] S��     [��� ������� �������]
Stp:=Stp*power(10,-3);
//[14] S�� ��  [��� ����� �� �� �������]
S_shH_nn:=S_shH_nn*power(10,-3);
//[15] d�� ��  [������� ����� �� �� �������]
d_shH_nn:=d_shH_nn*power(10,-3);
//[16] ��� ��  [������� ������� ����� �� �� �������]
b_shH_nn:=b_shH_nn*power(10,-3);
//[17] S�� ��  [��� ����� �� �� �������]
S_shH_ak:=S_shH_ak*power(10,-3);
//[18] d�� ��  [������� ����� �� �� �������]
d_shH_ak:=d_shH_ak*power(10,-3);
//[19] ��� ��  [������� ������� ����� �� �� �������]
b_shH_ak:=b_shH_ak*power(10,-3);


// 1)������� ���������� ������� ��� ����� ������ �� (1�) [�^2]
try  ftp_th :=(power(Stp,2)*cos(PI/6)-PI/4*power(d_H,2));     except   if Application.Messagebox('������ (1)   [������� ���������� ������� ��� ����� ������ ��].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
// 2)������� ���������� ������� (1�) [�^2]
try  fnpox_th:= ftp_th*n_TP;                                  except   if Application.Messagebox('������ (2)   [������� ���������� ������� �� ������� �������].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
// 3)���������� ������� ������ [�)
try  d_BH:=(d_H-2*btp);                                       except   if Application.Messagebox('������ (3)   [���������� ������� ������].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;

form4.ProgressBar1.position:=form4.ProgressBar1.position+5    ; //((((((())))))

 // 4)���������� ������� ������� ������ [�^2]
try  f_BH:=PI/4*power(d_BH,2);                                except   if Application.Messagebox('������ (4)   [���������� ������� ������� ������ �� ������� �������].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
 // 5)������� ����������� ������� ���� ������ (2�) [�^2]
try  fnpox_2:=n_TP*f_BH;                                      except   if Application.Messagebox('������ (5)   [������� ������� ���� ������ �� ������� �������].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
//��������� (1�)
 // 6)������� ����������� �� � �� [*�]
try  t_cp_TH:=(t_Bx+ t_Bix)/2;                                except   if Application.Messagebox('������ (6)   [������� ����������� �� � ��].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
 // 7) �������� ����� �� [�^3/��] (�������)
try  V_TH:= pul (3, P1, t_cp_TH);                             except   if Application.Messagebox('������ (7)   [�������� ����� �� (�������)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
 // 8)��������� �� �� ����� � �� [���/��]
try  i_Bx_TH:= pul (2, P1, t_Bx);                             except   if Application.Messagebox('������ (8)   [��������� �� �� ����� � ��  (�������)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
 // 9)��������� �� �� ������ �� �� [���/��]
try  i_Bix_TH:= pul (2, P1, t_Bix);                           except   if Application.Messagebox('������ (9)   [��������� �� �� ������ �� ��  (�������)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;

// 10)������� �������� �� �������� ��
//  ������������
Prov_P_nr:=5  ;
Delta_P_nn_2:=0 ;
Delta_P_icn_2:=0 ;
Delta_P_ak_2:=0  ;

form4.ProgressBar1.position:=form4.ProgressBar1.position+5    ; //((((((())))))
form4.height:=123 ;
form4.ProgressBar2.Visible:=True ;

 // �������  [%]
     while Prov_P_nr>=5 do begin  /// (�_2  � �_1 ������ 5%)
      //������������ �_2 � �_1
Delta_P_nn_1:=Delta_P_nn_2 ;
Delta_P_icn_1:=Delta_P_icn_2 ;
Delta_P_ak_1:=Delta_P_ak_2  ;
// 11)��������� ������� �������� � �� (2�)
Delta_P_nr_1:=Delta_P_nn_1+Delta_P_icn_1+Delta_P_ak_1 ;
// 12)�������� �� ������ �� ���
try  P_icn:=Pne+Delta_P_nn_1 ;                               except   if Application.Messagebox('������ (12)   [�������� �� ������ �� ���].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
// 13)�������� �� ������ �� ��
try  P_ak:=P_icn+Delta_P_icn_1 ;                             except   if Application.Messagebox('������ (13)   [�������� �� ������ �� ��].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
//14)�������� �� ����� � ��
try  P_BX:=P_ak+Delta_P_ak_1 ;                               except   if Application.Messagebox('������ (14)   [�������� �� ����� � ��].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
// 15)������� �������� �� �������� ��
try  P_ak_cp:=(P_ak+P_BX)/2 ;                                except   if Application.Messagebox('������ (15)   [������� �������� �: ��].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
try  P_icn_cp:=(P_ak+P_icn)/2 ;                              except   if Application.Messagebox('������ (15)   [������� �������� �: ���].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
try  P_nn_cp:=(P_icn+Pne)/2 ;                                except   if Application.Messagebox('������ (15)   [������� �������� �: ��].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
// 16)����������� �� ������ �� ��� � ��  (�������)
try  t_s_icn:=pul (10, P_icn, 0) ;                           except   if Application.Messagebox('������ (16)   [����������� �� ������ �� ��� (2�)  (�������)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
try  t_s_ak:=pul (10, P_ak, 0) ;                             except   if Application.Messagebox('������ (16)   [����������� �� ������ �� �� (2�)  (�������)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
// 17)��������� �������� ����  (2�) (�������)
   //��������� ����������� ����
     try  iYY:= pul (11, 0,  t_s_icn);                       except   if Application.Messagebox('������ (17)   [��������� ����������� ���� (2�)  (�������)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
   //��������� ���������� ����
     try  iY_s:=pul (6, 0,  t_s_ak) ;                        except   if Application.Messagebox('������ (17)   [��������� ���������� ���� (2�)  (�������)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
   //��������� ����������� ����
     try  i_nB:=pul (2, P_BX, tnb);                          except   if Application.Messagebox('������ (17)   [��������� ����������� ���� (2�)  (�������)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
   //��������� ����������� ����
     try  ine:=pul (15, Pne, tne) ;                          except   if Application.Messagebox('������ (17)   [��������� ����������� ���� (2�)  (�������)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
// 18)������� ����������� �������� ���� �� �������� ��
try  t_cp_ak:=(tnb+t_s_ak)/2 ;                               except   if Application.Messagebox('������ (18)   [������� ����������� �� �� ������� ��].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
try  t_cp_nn:=(t_s_icn+tne)/2 ;                              except   if Application.Messagebox('������ (18)   [������� ����������� �� �� ������� ��].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
try  t_cp_icn:=(t_s_ak+t_s_icn)/2 ;                          except   if Application.Messagebox('������ (18)   [������� ����������� �� �� ������� ���].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
// 19)�������� ��������, ������������ �� ������ �� ��������
try  Q_ak:=Dpac*(iY_s-i_nB) ;                                except   if Application.Messagebox('������ (19)   [�������� ��������, ������������  �� ������� ��].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
try  Q_icn:=Dpac*(iYY-iY_s) ;                                except   if Application.Messagebox('������ (19)   [�������� ��������, ������������  �� ������� ���].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
try  Q_nn:=Dpac*(ine-iYY) ;                                  except   if Application.Messagebox('������ (19)   [�������� ��������, ������������  �� ������� ��].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
// 20)�������� ��������� �������� � ��
try  Q_nr:=Q_ak+Q_icn+Q_nn;                                  except   if Application.Messagebox('������ (20)   [��������� �������� �������� � ��].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
// 21)������ ������������� ������� �������
try  G_TH:=Q_nr/(i_Bx_TH-i_Bix_TH) ;                         except   if Application.Messagebox('������ (21)   [������ �� (1�)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
// 22)�������� �� (1�) �� �������� �� � ���
  // 22.1)���������
    try  i_icn_TH:= i_Bx_TH-Q_nn/G_TH;                       except   if Application.Messagebox('������ (22.1)   [��������� �� (1�) �� ������� ��� (1�)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
    try  i_ak_TH:=i_icn_TH-Q_icn/G_TH ;                      except   if Application.Messagebox('������ (22.1)   [��������� �� (1�) �� ������� �� (1�)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
  // 22.2)����������� (�������)
    try  t_icn_TH:=pul (1, P1, i_icn_TH) ;                   except   if Application.Messagebox('������ (22.2)   [����������� �� (1�) �� ������� ��� (1�)  (�������)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
    try  t_ak_TH:=pul (1, P1, i_ak_TH) ;                     except   if Application.Messagebox('������ (22.2)   [����������� �� (1�) �� ������� �� (1�)  (�������)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
  // 22.3)������� ����������� �� �� ��������
    try  t_cp_TH_ne:=(t_Bx+t_icn_TH)/2   ;                   except   if Application.Messagebox('������ (22.3)   [������� ����. �� �� �� (1�)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
    try  t_cp_TH_icn:=(t_ak_TH+t_ak_TH)/2   ;                except   if Application.Messagebox('������ (22.3)   [������� ����. �� �� ������� ��� (1�)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
    try  t_cp_TH_ak:=(t_ak_TH+t_Bix)/2   ;                   except   if Application.Messagebox('������ (22.3)   [������� ����. �� �� ������� �� (1�)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
  // 22.4)������� ����������� ������ ������� ������ �� ��������
    try  t_cp_ct_nn:=(t_cp_TH_ne+t_cp_nn)/2   ;              except   if Application.Messagebox('������ (22.4)   [������� ����. ������ �� �� (1�)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
    try  t_cp_ct_icn:=(t_cp_TH_icn+t_cp_icn)/2   ;           except   if Application.Messagebox('������ (22.4)   [������� ����. ������ �� ������� ��� (1�)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
    try  t_cp_ct_ak:=(t_cp_TH_ak+t_cp_ak)/2   ;              except   if Application.Messagebox('������ (22.4)   [������� ����. ������ �� ������� �� (1�)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
// 23)�������� ����� �� � ������   (�������)
try  V_TH_ct:= pul (3, P1, t_cp_ct_icn);                     except   if Application.Messagebox('������ (23)   [�������� ����� �� � ������ (�������)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
// 24)���������������� �� � ������ � � ������ (�������)
       //�� � ������
       try  Lambda_TH:= pul (25, V_TH, t_cp_TH) ;            except   if Application.Messagebox('������ (24)   [���������������� �� � ������  (�������)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
       try  Lambda_TH:=Lambda_TH*power(10,-3) ;              except   if Application.Messagebox('������ (24)   [���������������� �� � ������ /1000].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
       //�� � ������
       try  Lambda_TH_ct:= pul (25, V_TH_ct, t_cp_ct_icn) ;  except   if Application.Messagebox('������ (24)   [���������������� �� � ������  (�������)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
       try  Lambda_TH_ct:=Lambda_TH_ct*power(10,-3) ;        except   if Application.Messagebox('������ (24)   [���������������� �� � ������ /1000)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
// 25)������������ �������� �� � ������ � � ������ (�������)
       //�� � ������
       try  Mu_TH:= pul (24, V_TH, t_cp_TH) ;                except   if Application.Messagebox('������ (25)   [������������ �������� �� � ������  (�������)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
       //�� � ������
       try  Mu_TH_ct:= pul (24, V_TH_ct, t_cp_ct_icn) ;      except   if Application.Messagebox('������ (25)   [������������ �������� �� � ������  (�������)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
// 26)�������������� �������� ��
try  KINV_cp_TH:=Mu_TH*V_TH*power(10,-6)  ;                  except   if Application.Messagebox('������ (26)   [��������������  �������� ��].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
//27) ������� �������� ��  � �� (1�)
try  W_1:=G_TH*V_TH/fnpox_th  ;                              except   if Application.Messagebox('������ (27)   [������� �������� ��  � �� (1�)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
// 28)������������� ������� ���������� � ����� ������ ��� ��
try  d_ek_1:=4*ftp_th/(PI*d_H)  ;                            except   if Application.Messagebox('������ (28)   [������������� ������� ����� ������ ��� ��].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
// 29)����� ���������� ��
try  Re_1:=(W_1 *d_ek_1)/KINV_cp_TH ;                        except   if Application.Messagebox('������ (29)   [����� ���������� ��].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
// 30)����� �������� �� � ������ � � ������
        // �� � ������
        try  Pr_TH:=(((i_Bx_TH-i_Bix_TH)*Mu_TH)/((t_Bx-t_Bix)*Lambda_TH))*power(10,-3);            except   if Application.Messagebox('������ (30)   [����� �������� �� � ������].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
        // �� � ������
        try  Pr_TH_ct:=(((i_Bx_TH-i_Bix_TH)*Mu_TH_ct)/((t_Bx-t_Bix)*Lambda_TH_ct))*power(10,-3);   except   if Application.Messagebox('������ (30)   [����� �������� �� � ������].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
// 31)����� ��������� ��
S_k_d:=Stp/d_H ;
  if  S_k_d>1.11 then
     begin
       try  Nu_1:=0.021*power(Re_1,0.8)*power(Pr_TH,0.43)*power(Pr_TH/Pr_TH_ct,0.25);              except   if Application.Messagebox('������ (31)   [����� ��������� ��� ��].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end
     end
  else
    begin
       try  Nu_1:=0.021*power(Re_1,0.8)*power(Pr_TH,0.43)*(1.2-0.724*exp(-11.7*(Stp/d_H-1)))*power(Pr_TH/Pr_TH_ct,0.25)  ;   except   if Application.Messagebox('������ (31)   [����� ��������� ��� ��].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
    end;
// 32)������������ ����������� (1�)
try  a_1 :=Nu_1*Lambda_TH/d_ek_1   ;                           except   if Application.Messagebox('������ (32)   [������������ ����������� (1�)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;


  //    [[�����������]]
// 33)�������� ����� �� �� �� ��. (�������)
        //�� � ������ (�������)
        try  V_ak_cp:=pul (3, P_ak_cp, t_cp_ak)  ;             except   if Application.Messagebox('������ (33)   [�������� ����� �� � �� (2�)  (�������)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
        //�� � ������ (�������)
        try  V_ak_cp_ct:=pul (3, P_ak_cp, t_cp_ct_ak)  ;       except   if Application.Messagebox('������ (33)   [�������� ����� � �� � ������ (2�)  (�������)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
// 34)������������ �������� �� (2�) �� �� ��. (�������)
   //�� [��/(��)]   [������������ ��������] (�������)
        //�� � ������ (�������)
        try  Mu_cp_ak:=pul (24, V_ak_cp, t_cp_ak)  ;           except   if Application.Messagebox('������ (34)   [������������ �������� �� � �� (2�)  (�������)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
        //�� � ������ (�������)
        try  Mu_cp_ct_ak:=pul (24, V_ak_cp_ct, t_cp_ct_ak)  ;  except   if Application.Messagebox('������ (34)   [������������ �������� �� � ������ � �� (2�)  (�������)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
// 35)�������������� �������� �� (2�) �� �� ��.
   try  KINV_cp_ak:=Mu_cp_ak*V_ak_cp*power(10,-6)  ;           except   if Application.Messagebox('������ (35)   [�������������� �������� �� � �� (2�)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
// 36)���������������� �� (2�) �� �� ��. (�������)
        //�� � ������ (�������)
       try  Lambda_ak:=pul (25, V_ak_cp, t_cp_ak)  ;           except   if Application.Messagebox('������ (36)   [���������������� �� � �� (2�)  (�������)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
       try  Lambda_ak:=Lambda_ak*power(10,-3) ;                except   if Application.Messagebox('������ (36)   [���������������� �� � �� /1000 (2�)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
        //�� � ������ (�������)
       try  Lambda_ct_ak:= pul (25, V_ak_cp_ct, t_cp_ct_ak) ;  except   if Application.Messagebox('������ (36)   [���������������� �� � ������ � �� (2�)  (�������)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
       try  Lambda_ct_ak:=Lambda_ct_ak*power(10,-3) ;          except   if Application.Messagebox('������ (36)   [���������������� �� � ������ � �� /1000 (2�)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
// 37)����� �������� ��� �� (2�) �� �� ��.
        //�� � ������
        try  Pr_ak:=((iY_s-i_nB)*Mu_cp_ak)/((t_s_ak-tnb)*Lambda_ak)*power(10,-3)  ;          except   if Application.Messagebox('������ (37)   [����� �������� ��� � �� � �� (2�)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
        //�� � ������
        try  Pr_ct_ak:=((iY_s-i_nB)*Mu_cp_ct_ak)/((t_s_ak-tnb)*Lambda_ct_ak)*power(10,-3);   except   if Application.Messagebox('������ (37)   [����� �������� ���  �� � �� (2�)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;

//���������������� ��������� ������ �� �� (2�) �� �� ��.
   try  Case ComboBox1.ItemIndex of
{38) �����}   0:  Lambda_MAT_ak:=(62.1-0.0515*t_cp_ct_ak) ;                 //�����
{39) �����}   1:  Lambda_MAT_ak:=(0.009+0.00003*t_cp_ct_ak)*power(10,3) ;   //�����
        end;
   except   if Application.Messagebox('������ (38,39)   [���������������� ��������� ������ � �� (2�)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
// 40)������� ���������� ������� (2�) �� �� ��.
try  f_ak_TP:=(S_shH_ak-b_shH_ak)*(d_BH-d_shH_ak)/2; {//���� ((��))}  except   if Application.Messagebox('������ (40)   [������� ������ � �� (2�)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
try  f_ak:=f_ak_TP*n_TP  ;                           {//���� ((��))}  except   if Application.Messagebox('������ (40)   [������� ���� ����  � �� (2�)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
//�������� �� �� (2�) �� �� ��.

   try  Case ComboBox2.ItemIndex of
{41) ����}   0:  W_ak:=Dpac*V_ak_cp/f_ak  ;      //���� ((��))
{42) ����}   1:  W_ak:=Dpac*V_ak_cp/fnpox_2  ;   //���� ((���))
        end;
   except   if Application.Messagebox('������ (41,42)   [�������� � �� (2�)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;

// 43)������������� ������� ��� �� (2�) �� �� ��.
  try  d_a_ak:=d_BH-d_shH_ak ;                                        except   if Application.Messagebox('������ (43)   [������������� ������� ��������� ������ �����  � �� (2�)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
// 44)����� ���������� ��� �� (2�) �� �� ��.

   try  Case ComboBox2.ItemIndex of
           0:  Re_ak:=W_ak*d_a_ak/KINV_cp_ak  ;  //���� ((��))
           1:  Re_ak:=W_ak*d_BH/KINV_cp_ak  ;    //���� ((���))
      end;
   except   if Application.Messagebox('������ (44)   [����� ���������� � �� (2�)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;

// 45)������� �������� ��������� ������ ��� �� (2�) �� �� ��.
  try  D_ak:=0.5*(d_BH+d_shH_ak)*(1+power(2*S_shH_ak/(PI*(d_BH+d_shH_ak)),2)) ; {���� ((��))}   except   if Application.Messagebox('������ (45)   [������� �������� ��������� ������ � �� (2�)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
    //������������� ��� (2�) �� �� ��.
  try  Psi_ak:=1+0.4*power(S_shH_ak/(d_BH+d_shH_ak),2) ; {���� ((��))}                          except   if Application.Messagebox('������ (45)   [������������� � �� (2�)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
    //����� ��������� ��� �� (2�) �� �� ��.

   try  Case ComboBox2.ItemIndex of
           0:  Nu_ak:=0.65*power(Re_ak,0.5)*power(d_a_ak/D_ak,0.45)*power(Psi_ak,0.2)*power(Pr_ak,0.43)*power(Pr_ak/Pr_ct_ak,0.25); //���� ((��))
           1:  Nu_ak:=0.021*power(Re_ak,0.8)*power(Pr_ak,0.43)*power(Pr_ak/Pr_ct_ak,0.25) ; //���� ((���))
          end;
   except   if Application.Messagebox('������ (45)   [����� ��������� ��� �� (2�)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;

// 46)����������� ����������� �� (2�) �� �� ��.

   try  Case ComboBox2.ItemIndex of
           0:  a_ak :=Nu_ak*Lambda_ak/d_a_ak ;  //���� ((��))
           1:  a_ak :=Nu_ak*Lambda_ak/d_BH ;   //���� ((���))
          end;
   except   if Application.Messagebox('������ (46)   [����������� ����������� ��� �� (2�)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;

//47)����������� ������������� �� �� ��.
try  K_ak:=1/((1/a_1+btp /Lambda_MAT_ak+1/a_ak))*power(10,-3)  ;                    except   if Application.Messagebox('������ (47)   [����������� ������������� ��� �� (2�)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
//48)������������� ����� ��� (2�) �� �� ��.
try  Delta_t_ak:=((t_Bix-tnb)-(t_ak_TH-t_s_ak))/ln((t_Bix-tnb)/(t_ak_TH-t_s_ak)) ;  except   if Application.Messagebox('������ (48)   [������������� ����� ��� �� (2�)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
// 49)������� ����������� �� ��.
try  F__ak:=Q_ak/(K_ak*Delta_t_ak) ;                           except   if Application.Messagebox('������ (49)   [������� ����������� ��� �� ].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
// 50)������� ������� ������� ������ �� �� ��.
try  d_cp_ak:=(a_1*d_BH+a_ak*d_H)/(a_1+a_ak) ;                 except   if Application.Messagebox('������ (50)   [������� ������� �� ].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
// 51)������ �� ��. ��� (2�)
try  L_ak:=F__ak/(PI*d_cp_ak*n_TP) ;                           except   if Application.Messagebox('������ (51)   [������ �� ��. [�] ].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;

  //    [[�����������������]]
// 52)�������� ����� �� �� �� ��. (�������)
        //�� � ������ (�������)
        try  V_nn_cp:=pul (16, P_nn_cp, t_cp_nn)  ;            except   if Application.Messagebox('������ (52)   [�������� ����� �� � ��  (�������)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
        //�� � ������ (�������)
        try  V_nn_cp_ct:=pul (16, P_nn_cp, t_cp_ct_nn);        except   if Application.Messagebox('������ (52)   [�������� ����� �� � ������ � ��  (�������)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
// 53)������������ �������� �� (2�) �� �� ��. (�������)
        //�� � ������ (�������)
        try  Mu_cp_nn:=pul (24, V_nn_cp, t_cp_nn)  ;           except   if Application.Messagebox('������ (53)   [������������ �������� ��  (�������)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
        //�� � ������ (�������)
        try  Mu_cp_ct_nn:=pul (24, V_nn_cp_ct, t_cp_ct_nn);    except   if Application.Messagebox('������ (53)   [������������ �������� �� � ������  (�������)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
// 54)�������������� �������� �� (2�) �� �� ��.
try  KINV_cp_nn:=Mu_cp_nn*V_nn_cp*power(10,-6)  ;              except   if Application.Messagebox('������ (54)   [�������������� �������� ��].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
// 55)���������������� �� (2�) �� �� ��. (�������)
        //�� � ������ (�������)
        try  Lambda_nn:=pul (25, V_nn_cp, t_cp_nn)  ;          except   if Application.Messagebox('������ (55)   [���������������� ��  (�������)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
        try  Lambda_nn:=Lambda_nn*power(10,-3) ;               except   if Application.Messagebox('������ (55)   [���������������� �� /1000].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
        //�� � ������ (�������)
        try  Lambda_ct_nn:= pul (25, V_nn_cp_ct, t_cp_ct_nn) ; except   if Application.Messagebox('������ (55)   [���������������� �� � ������  (�������)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
        try  Lambda_ct_nn:=Lambda_ct_nn*power(10,-3) ;         except   if Application.Messagebox('������ (55)   [���������������� �� � ������ /1000].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
// 56)����� �������� ��� �� (2�) �� �� ��.
        //�� � ������ (�������)
        try  Pr_nn:=((ine-iYY)*Mu_cp_nn)/((tne-t_s_icn)*Lambda_nn)*power(10,-3)  ;          except   if Application.Messagebox('������ (56)   [����� �������� ��� ��].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
        //�� � ������ (�������)
        try  Pr_ct_nn:=((ine-iYY)*Mu_cp_ct_nn)/((tne-t_s_icn)*Lambda_ct_nn)*power(10,-3);   except   if Application.Messagebox('������ (56)   [����� �������� ��� �� � ������].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
//���������������� ��������� ������ �� �� (2�) �� �� ��.

    try  Case ComboBox1.ItemIndex of
{57) �����}   0:  Lambda_MAT_nn:=(62.1-0.0515*t_cp_ct_nn) ;                 {�����}
{58) �����}   1:  Lambda_MAT_nn:=(0.009+0.00003*t_cp_ct_nn)*power(10,3) ;   {�����}
         end;
    except   if Application.Messagebox('������ (57,58)   [���������������� ��������� ������ ��].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
// 59)������� ���������� ������� (2�) �� �� ��.
try  f_nn_TP:=(S_shH_nn-b_shH_nn)*(d_BH-d_shH_nn)/2 ;     except   if Application.Messagebox('������ (59)   [������� ������ ��].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
try  f_nn:=f_nn_TP*n_TP  ;                                except   if Application.Messagebox('������ (59)   [������� ���� ���� � ��].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
//������� �������� �� �� (2�) �� �� ��.

   try  Case ComboBox2.ItemIndex of
{60) ����}   0:  W_nn:=Dpac*V_nn_cp/f_nn  ;   //���� ((��))
{61) ����}   1:  W_nn:=Dpac*V_nn_cp/fnpox_2 ; //���� ((���))
        end;
    except   if Application.Messagebox('������ (60,61)   [�������� � ��].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
//������������� ������� ��� �� (2�) �� �� ��.
   try  Case ComboBox2.ItemIndex of
{62) ����}   0:  d_a_nn:= d_BH-d_shH_nn ;   //���� ((��))
{63) ����}   1:  d_a_nn:= d_BH          ; //���� ((���))
        end;
    except   if Application.Messagebox('������ (62;63)   [������������� ������� ��].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
// 64)����� ���������� ��� �� (2�) �� �� ��.
  try  Case ComboBox2.ItemIndex of
          0:  Re_nn:=W_nn*d_a_nn/KINV_cp_nn ; //���� ((��))
          1:  Re_nn:=W_nn*d_BH/KINV_cp_nn ; //���� ((���))
       end;
    except   if Application.Messagebox('������ (64)   [����� ���������� ��� ��].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
// 65)
  //������� �������� ��������� ������ ��� �� (2�) �� �� ��.
   try  D_nn:=0.5*(d_BH+d_shH_nn)*(1+power(2*S_shH_nn/(PI*(d_BH+d_shH_nn)),2)) ; {���� ((��))} except   if Application.Messagebox('������ (65)   [������� �������� ��������� ������ ��].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
  //������������� ��� (2�) �� �� ��.
   try  Psi_nn:=1+0.4*power(S_shH_nn/(d_BH+d_shH_nn),2) ;  {���� ((��))}                       except   if Application.Messagebox('������ (65)   [������������� ��].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
  //����� ��������� ��� �� (2�) �� �� ��.
   try  Case ComboBox2.ItemIndex of
           0:  Nu_nn:=0.65*power(Re_nn,0.5)*power(d_a_nn/D_nn,0.45)*power(Psi_nn,0.2)*power(Pr_nn,0.43)*power(Pr_nn/Pr_ct_nn,0.25); //���� ((��))
           1:  Nu_nn:=0.021*power(Re_nn,0.8)*power(Pr_nn,0.43)*power(Pr_nn/Pr_ct_nn,0.25) ; //���� ((���))
        end;
   except   if Application.Messagebox('������ (65)   [����� ��������� ��� ��].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
// 66)����������� ����������� �� (2�) �� �� ��.

   try  Case ComboBox2.ItemIndex of
           0:  a_nn :=Nu_nn*Lambda_nn/d_a_nn ; //���� ((��))
           1:  a_nn :=Nu_nn*Lambda_nn/d_BH ;   //���� ((���))
        end;
   except   if Application.Messagebox('������ (66)   [����������� ����������� ��].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
// 67)����������� ������������� �� �� ��.
try  K_nn:=1/(1/a_1+btp/Lambda_MAT_nn+1/a_nn)*power(10,-3)  ;                                except   if Application.Messagebox('������ (67)   [����������� ������������� ��].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
// 68)������������� ����� ��� (2�) �� �� ��.
try  Delta_t_nn:=((t_Bx-tne)-(t_icn_TH-t_s_icn))/ln((t_Bx-tne)/(t_icn_TH-t_s_icn)) ;         except   if Application.Messagebox('������ (68)   [������������� ����� ��].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
// 69) ������� ����������� �� ��.
try  F__nn:=Q_nn/(K_nn*Delta_t_nn) ;             except   if Application.Messagebox('������ (69)   [������� ����������� ��].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
// 70)������� ������� ������� ������ �� �� ��.
try  d_cp_nn:=(a_1*d_BH+a_nn*d_H)/(a_1+a_nn) ;   except   if Application.Messagebox('������ (70)   [������� ������� ��].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
// 71)������ �� ��. ��� (2�)
try  L_nn:=F__nn/(PI*d_cp_nn*n_TP)  ;            except   if Application.Messagebox('������ (71)   [������ ��].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;


  //    [[����������]]
// 72)�������� �������� �����
try  q_1:=500*power(10,3) ;                      except   if Application.Messagebox('������ (72)   [q_1].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
// ��� �����
//// ���� ������ ������ ����� 5%
Prov_q:=5  ;
q_2:=q_1   ;
 // �������  [%]
     while Prov_q>=5 do begin  /// (q1 � q2 ������ 5%)
      // ������������ q2 � q1
      q_1:=q_2;
      //73)����������� ����������� �� (2�) �� ��� ��.
        try  a_icn:=power(10,6)*power(q_1/power(10,6),0.7)/(39-0.1*t_s_icn) ;   except   if Application.Messagebox('������ (73)   [����������� ����������� ���].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
      // 74)������� ����������� ������ ������� ������ �� (2�) �� ��� ��.
       try  t_cp_ct_icn:=((t_icn_TH+t_ak_TH)/2+t_cp_icn)/2  ;                   except   if Application.Messagebox('������ (74)   [������� ����������� ������ ���].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
      // 75)���������������� ��������� ������ ������� ������ �� (2�) �� ��� ��.
       try  Case ComboBox1.ItemIndex of
              0:  Lambda_MAT_icn:=(62.1-0.0515*t_cp_ct_icn) ;               //�����
              1:  Lambda_MAT_icn:=(0.009+0.00003*t_cp_ct_icn)*power(10,3) ; //�����
            end;
       except   if Application.Messagebox('������ (75)   [���������������� ��������� ������ ���].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;

     // 76)����������� ������������� �� (2�) �� ��� ��.
       try  K_icn:=1/(1/a_1+btp/Lambda_MAT_icn+1/a_icn)*power(10,-3)   ;        except   if Application.Messagebox('������ (76)   [����������� ������������� ���].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
     // 77)������������� ����� ��� (2�) �� ��� ��.
       try  Delta_t_icn:=((t_icn_TH-t_s_icn)-(t_ak_TH-t_s_ak))/ln((t_icn_TH-t_s_icn)/(t_ak_TH-t_s_ak))  ;   except   if Application.Messagebox('������ (77)   [������������� ����� ���].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
     // 78)�������� �������� �����
       try  q_2:=K_icn*Delta_t_icn*power(10,3)   ;     except   if Application.Messagebox('������ (78)   [q_2].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
     // 79)��������
       try  Prov_q:= abs((q_2-q_1)/q_2)*100  ;         except   if Application.Messagebox('������ (79)   [�������� �� q].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
     // 80)������� ����������� ��� ��. ��� (2�)
     try  F__icn:=Q_icn/(K_icn*Delta_t_icn) ;          except   if Application.Messagebox('������ (80)   [������� ����������� ���].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
     // 81)������� ������� ������� ������ �� ��� ��.
     try  d_cp_icn:=(a_1*d_BH+a_icn*d_H)/(a_1+a_icn) ; except   if Application.Messagebox('������ (81)   [������� ������� ���].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
     // 82)������ ��� ��.
     try  L_icn:=F__icn/(PI*d_cp_icn*n_TP) ;           except   if Application.Messagebox('������ (82)   [������ ���].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;

                   end; /// (q1 � q2 ������ 5%)


  //    [[[������ ��������]]]

// 83)������ ��
try  L_nr:=L_ak+L_nn+L_icn  ;                          except   if Application.Messagebox('������ (83)   [������ ��].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
// 84)����������� ������ �� �� (1�)
S_k_d:=Stp/d_H ;
  if  S_k_d>1.11 then
     begin
        try  Zeta_1:=0.3164*power(Re_1,-0.25);         except   if Application.Messagebox('������ (84)   [����. ������ (1�)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
     end
  else
     begin
        try  Zeta_1:=0.3164*power(Re_1,-0.25)*(1.1-0.42*exp(-13*(Stp/d_H-1)));  except   if Application.Messagebox('������ (84)   [����. ������ (1�)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
     end;
// 85)������ �������� � �� �� (1�) [���]
try  Delta_P_1_2:=(1.5+Zeta_1*L_nr/d_ek_1)*power(W_1,2)/(2*V_TH)*power(10,-6);  except   if Application.Messagebox('������ (85)   [������� �������� (1�)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;

//    ((�����������))
// ��� �������� ������
   c_ak:= 5.55 ;
// 86)����������� ��������������� ������������� �� �� ��. �� (2�)
    try  Case ComboBox2.ItemIndex of
            0:  Zeta_ak:=c_ak*power(Re_ak,-0.42)*power(d_a_ak/D_ak,0.65)*power(Psi_ak,0.3)  ; //���� ((��))
            1:  Zeta_ak:=0.316*power(Re_ak,-0.25) ;   //���� ((���))
         end;
    except   if Application.Messagebox('������ (86)   [����������� ��������������� ������������� ��].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
// 87)������ �������� � �� �� �� ��. �� (2�) [���]
    try  Case ComboBox2.ItemIndex of
            0:  Delta_P_ak_2:=power(10,-6)*Zeta_ak*L_ak*W_ak/(d_a_ak*2*V_ak_cp)  ; //���� ((��))
            1:  Delta_P_ak_2:=power(10,-6)*Zeta_ak*L_ak*W_ak/(d_BH*2*V_ak_cp) ;   //���� ((���))
         end;
    except   if Application.Messagebox('������ (87)   [������� �������� ��].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;

//   ((����������))
// 88)�������� ����� ���� � ���� �� ����� ��������� (�������)
    try  V_cp_YY:= pul (12, 0,  t_s_icn);              except   if Application.Messagebox('������ (88)   [�������� ����� ���� �� ����� ���������  (�������)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
    try  V_cp_Y:=pul (7, 0,  t_s_ak);                  except   if Application.Messagebox('������ (88)   [�������� ����� ���� �� ����� ���������  (�������)].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
// 89)������� �������� ���������� �� (2�) �� �� ��.
    try  W_0:=Dpac*V_cp_Y/fnpox_2  ;                   except   if Application.Messagebox('������ (89)   [�������� � ���].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
// 90)������������ �������� �� (2�) �� ��� ��. (�������)
    try  Mu_cp_icn:=pul (24, V_cp_Y, t_cp_icn)  ;      except   if Application.Messagebox('������ (90)   [�������� ������������ � ���].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
// 91)�������������� �������� �� (2�) �� ��� ��.
    try  KINV_cp_icn:=Mu_cp_icn*V_cp_Y*power(10,-6);   except   if Application.Messagebox('������ (91)   [�������� �������������� � ���].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
// 92)������������� ������� ��� �� (2�) �� ��� ��.
    d_ek_icn:=d_BH  ;   //����������������
// 93)����� ���������� ��� �� (2�) �� ��� ��.
    try  Re_icn:=W_0*d_ek_icn/KINV_cp_icn  ;           except   if Application.Messagebox('������ (93)   [����� ���������� ��� ���].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
// 94)����������� ��������������� ������������� �� (2�) �� ��� ��.
    try  Zeta_icn:=0.3164*power(Re_icn,-0.25) ;        except   if Application.Messagebox('������ (94)   [����������� ��������������� ������������� ���].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
// 95)������ �������� � �� �� ��� ��. �� (2�)
    try  Delta_P_icn_2:=power(10,-6)*Zeta_icn*L_icn*power(W_0,2)/(d_ek_icn*V_cp_Y)*(1+0.5*(V_cp_YY/V_cp_Y-1))  ;   except   if Application.Messagebox('������ (95)   [������ �������� � �� �� ��� �������].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;

//   ((����������������� ))
// ��� �������� ������
   c_nn:= 5.55 ;
// 96)����������� ��������������� ������������� �� �� ��. �� (2�)
    try  Case ComboBox2.ItemIndex of
            0:  Zeta_nn:=c_nn*power(Re_nn,-0.42)*power(d_a_nn/D_nn,0.65)*power(Psi_nn,0.3)  ; //���� ((��))
            1:  Zeta_nn:=c_nn*power(Re_nn,-0.42)*power(d_BH/D_nn,0.65)*power(Psi_nn,0.3)  ;   //���� ((���))
         end;
    except   if Application.Messagebox('������ (96)   [����������� ��������������� ������������� ��].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
// 97)������ �������� � �� �� �� ��. �� (2�)
    try  Case ComboBox2.ItemIndex of
            0:  Delta_P_nn_2:=power(10,-6)*Zeta_nn*L_nn*power(W_nn,2)/(d_a_nn*2*V_nn_cp)  ; //���� ((��))
            1:  Delta_P_nn_2:=power(10,-6)*Zeta_nn*L_nn*power(W_nn,2)/(d_BH*2*V_nn_cp)  ;   //���� ((���))
         end;
    except   if Application.Messagebox('������ (97)   [������ �������� � �� �� �� �������].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;

// 98)��������� ������ �������� �� (2�)
try  Delta_P_nr_2:=Delta_P_ak_2+Delta_P_icn_2+Delta_P_nn_2 ;              except   if Application.Messagebox('������ (98)   [����� ������ �������� � ��].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
// 99)��������
try  Prov_P_nr:= abs((Delta_P_nr_2-Delta_P_nr_1)/Delta_P_nr_2)*100  ;     except   if Application.Messagebox('������ (99)   [�������� Delta_P].','������ !!!', mb_iconerror or mb_ok) = mrOk    then   begin Day ; Exit;  end;end;
    end ;  /// (�_2  � �_1 ������ 5%)

//������� �� [��] [�] (��� ������)
d_H:=d_H*1000 ;
btp:=btp*1000 ;
Stp:=Stp*1000 ;
S_shH_nn:=S_shH_nn*1000 ;
d_shH_nn:=d_shH_nn*1000 ;
b_shH_nn:=b_shH_nn*1000 ;
S_shH_ak:=S_shH_ak*1000 ;
d_shH_ak:=d_shH_ak*1000 ;
b_shH_ak:=b_shH_ak*1000 ;

/// ��������
form4.ProgressBar2.Visible:=False ;

form4.ProgressBar1.position:=form4.ProgressBar1.position+5 ;    //((((((())))))

/// ((((����� ������))))

///  Edit_npo.Text:=FloatToStr(Rho_MAT);
///  Label96.caption:='  Rho_MAT  ';




///  [[[����� ������]]]





//       [[{{{��������� Excel}}}]]
    //           (((�����������)))

// �������� �� �� ������� ������������� Excel
      // ���� CLSID OLE-�������
  Rez := CLSIDFromProgID(PWideChar(WideString('Excel.Application')), ClassID);
  if Rez = S_OK then  // ������ ������
    Result := true
  else
     begin
    Result := false;
    ShowMessage('� ������ Excel ��������, � ����� �� ������� �����.') ;
    Day ; Exit;
     end;


if (CheckBox6.Checked=true) and (CheckBox9.Checked=true) then
  BeGiN ///[[{{{��������� Excel}}}]] (((�����������)))

// ���������� ������ �������� ���� �������, � ������� ����� �������� ������
  BeginCol := 4;
  BeginRow := 3;

// ������� ���������� ������� ������
  // �������
  RowCount := 300;
  // �������
  ColCount := 3;

// �������� Excel
  ExcelApp := CreateOleObject('Excel.Application');

form4.ProgressBar1.position:=form4.ProgressBar1.position+5 ;    //((((((())))))

// ��������� ������� Excel �� �������, ����� �������� ����� ����������
  ExcelApp.Application.EnableEvents := false;



 //  ������� ����� (Workbook)
 //  ���� ��������� ������, ��
 //  Workbook := ExcelApp.WorkBooks.Add('C:\MyTemplate.xls');
//OpenDialog1.filter:= ' ���� Excel (*.xls)  |*.xls| ��� ����� (*.*) |*.*| ';
//OpenDialog1.Title:='����� �������' ;
//if OpenDialog1.Execute then
//  try  //(/�������� ������)
//    Workbook := ExcelApp.WorkBooks.Add(OpenDialog1.filename);

 Workbook := ExcelApp.WorkBooks.Add();

// ������� ���������� ������, ������� �������� ��������� �������
ArrayData := VarArrayCreate([1, RowCount, 1, ColCount], varVariant);

form4.ProgressBar1.position:=form4.ProgressBar1.position+5 ;    //((((((())))))

//����������

t_icn_TH:=RoundSignificant(t_icn_TH,5);  // t_icn_TH
t_ak_TH:=RoundSignificant(t_ak_TH,5);   // t_ak_TH
t_s_ak:=RoundSignificant(t_s_ak,5);    // t_s_ak
W_1:=RoundSignificant(W_1,5);       // W_1
a_1 :=RoundSignificant(a_1 ,6);    // a_1
L_nr:=RoundSignificant(L_nr,5);   // L_nr
G_TH:=RoundSignificant(G_TH,5);   // G_TH
Delta_P_1_2:=RoundSignificant(Delta_P_1_2,5); // Delta_P_1_2
Delta_P_nr_2:=RoundSignificant(Delta_P_nr_2,5);  // Delta_P_nr_2

form4.ProgressBar1.position:=form4.ProgressBar1.position+5 ;    //((((((())))))

a_nn:=RoundSignificant(a_nn,7);    // a_nn
Delta_t_nn:=RoundSignificant(Delta_t_nn,5);  // Delta_t_nn
L_nn:=RoundSignificant(L_nn,5);   // L_nn
K_nn:=RoundSignificant(K_nn,5); // K_nn
Delta_P_nn_2:=RoundSignificant(Delta_P_nn_2,5); // Delta_P_nn_2
F__nn:=RoundSignificant(F__nn,5);    // F__nn
W_nn:=RoundSignificant(W_nn,5); // W_nn
Q_nn:=RoundSignificant(Q_nn,5); // Q_nn

form4.ProgressBar1.position:=form4.ProgressBar1.position+5 ;    //((((((())))))

a_icn:=RoundSignificant(a_icn,7);   // a_icn
Delta_t_icn:=RoundSignificant(Delta_t_icn,5); // Delta_t_icn
L_icn:=RoundSignificant(L_icn,5);  // L_icn
K_icn:=RoundSignificant(K_icn,5);   // K_icn
Delta_P_icn_2:=RoundSignificant(Delta_P_icn_2,5); // Delta_P_icn_2
F__icn:=RoundSignificant(F__icn,5);  // F__icn
W_0:=RoundSignificant(W_0,5);   // W_0
Q_icn:=RoundSignificant(Q_icn,5); // Q_icn

form4.ProgressBar1.position:=form4.ProgressBar1.position+5 ;    //((((((())))))

a_ak:=RoundSignificant(a_ak,7);  // a_ak
Delta_t_ak:=RoundSignificant(Delta_t_ak,5);  // Delta_t_ak
L_ak:=RoundSignificant(L_ak,5);    // L_ak
K_ak:=RoundSignificant(K_ak,5);    // K_ak
Delta_P_ak_2:=RoundSignificant(Delta_P_ak_2,5);  // Delta_P_ak_2
F__ak:=RoundSignificant(F__ak,5);  // F__ak
W_ak:=RoundSignificant(W_ak,5);   // W_ak
Q_ak:=RoundSignificant(Q_ak,5);  // Q_ak



 // ��������� ������
// �������� ������


     ArrayData[1,3] := Dpac;
     ArrayData[2,3] := Pne;
     ArrayData[3,3] := tne;
     ArrayData[4,3] := tnb;
     ArrayData[5,3] := t_Bx;
     ArrayData[6,3] := P1;
     ArrayData[7,3] := t_Bix;
     ArrayData[8,3] := MAT_TP;
     ArrayData[9,3] := d_H;
     ArrayData[10,3] := btp;
     ArrayData[11,3] := Stp;
     ArrayData[12,3] := n_TP;
     ArrayData[13,3] := XOT_ShH;
 If ComboBox2.ItemIndex=0 then  //���� ((��))
       begin
         ArrayData[14,3] := S_shH_nn;
         ArrayData[15,3] := d_shH_nn;
         ArrayData[16,3] := b_shH_nn;
         ArrayData[17,3] := S_shH_ak;
         ArrayData[18,3] := d_shH_ak;
         ArrayData[19,3] := b_shH_ak;
       end;


// ������

ArrayData[22,3] := t_icn_TH;
ArrayData[23,3] := t_ak_TH;
ArrayData[24,3] := t_s_ak;
ArrayData[25,3] := W_1;
ArrayData[26,3] := a_1 ;
ArrayData[27,3] := L_nr;
ArrayData[28,3] := G_TH;
ArrayData[29,3] := Delta_P_1_2;
ArrayData[30,3] := Delta_P_nr_2;

form4.ProgressBar1.position:=form4.ProgressBar1.position+5 ;    //((((((())))))

ArrayData[33,1] := a_nn;
ArrayData[34,1] := Delta_t_nn;
ArrayData[35,1] := L_nn;
ArrayData[36,1] := K_nn;
ArrayData[37,1] := Delta_P_nn_2;
ArrayData[38,1] := F__nn;
ArrayData[39,1] := W_nn;
ArrayData[40,1] := Q_nn;

ArrayData[33,2] := a_icn;
ArrayData[34,2] := Delta_t_icn;
ArrayData[35,2] := L_icn;
ArrayData[36,2] := K_icn;
ArrayData[37,2] := Delta_P_icn_2;
ArrayData[38,2] := F__icn;
ArrayData[39,2] := W_0;
ArrayData[40,2] := Q_icn;


ArrayData[33,3] := a_ak;
ArrayData[34,3] := Delta_t_ak;
ArrayData[35,3] := L_ak;
ArrayData[36,3] := K_ak;
ArrayData[37,3] := Delta_P_ak_2;
ArrayData[38,3] := F__ak;
ArrayData[39,3] := W_ak;
ArrayData[40,3] := Q_ak;


form4.ProgressBar1.position:=form4.ProgressBar1.position+5 ;    //((((((())))))


// ����� ������� ������ �������, � ������� ����� �������� ������
  Cell1 := WorkBook.WorkSheets[1].Cells[BeginRow, BeginCol];

// ������ ������ ������ �������, � ������� ����� �������� ������
  Cell2 := WorkBook.WorkSheets[1].Cells[BeginRow  + RowCount - 1, BeginCol +
ColCount - 1];

// �������, � ������� ����� �������� ������
  Range := WorkBook.WorkSheets[1].Range[Cell1, Cell2];

// � ��� � ��� ����� ������
  // ������� ������� ����������� ����������
  Range.Value := ArrayData;

form4.ProgressBar1.position:=form4.ProgressBar1.position+5 ;    //((((((())))))

//������ ������
  //����
ExcelPamka('A2:F2',2,2,-4138,-4138) ;
ExcelPamka('A3:F3',2,2,-4138,-4138) ;
ExcelPamka('A4:F4',2,2,-4138,-4138) ;
ExcelPamka('A5:F5',2,2,-4138,-4138) ;
ExcelPamka('A6:F6',2,2,-4138,-4138) ;
ExcelPamka('A7:F7',2,2,-4138,-4138) ;
ExcelPamka('A8:F8',2,2,-4138,-4138) ;
ExcelPamka('A9:F9',2,2,-4138,-4138) ;
ExcelPamka('A10:F10',2,2,-4138,-4138) ;
ExcelPamka('A11:F11',2,2,-4138,-4138) ;
ExcelPamka('A12:F12',2,2,-4138,-4138) ;
ExcelPamka('A13:F13',2,2,-4138,-4138) ;
ExcelPamka('A14:F14',2,2,-4138,-4138) ;
ExcelPamka('A15:F15',2,2,-4138,-4138) ;
ExcelPamka('A16:F16',2,2,-4138,-4138) ;
ExcelPamka('A17:F17',2,2,-4138,-4138) ;
ExcelPamka('A18:F18',2,2,-4138,-4138) ;
ExcelPamka('A19:F19',2,2,-4138,-4138) ;
ExcelPamka('A20:F20',2,2,-4138,-4138) ;
ExcelPamka('A21:F21',2,2,-4138,-4138) ;
form4.ProgressBar1.position:=form4.ProgressBar1.position+5 ;    //((((((())))))
ExcelPamka('A24:F24',2,2,-4138,-4138) ;
ExcelPamka('A25:F25',2,2,-4138,-4138) ;
ExcelPamka('A26:F26',2,2,-4138,-4138) ;
ExcelPamka('A27:F27',2,2,-4138,-4138) ;
ExcelPamka('A28:F28',2,2,-4138,-4138) ;
ExcelPamka('A29:F29',2,2,-4138,-4138) ;
ExcelPamka('A30:F30',2,2,-4138,-4138) ;
ExcelPamka('A31:F31',2,2,-4138,-4138) ;
ExcelPamka('A32:F32',2,2,-4138,-4138) ;
ExcelPamka('A24:F24',2,2,-4138,-4138) ;
ExcelPamka('A25:F25',2,2,-4138,-4138) ;
ExcelPamka('A26:F26',2,2,-4138,-4138) ;
ExcelPamka('A27:F27',2,2,-4138,-4138) ;
ExcelPamka('A28:F28',2,2,-4138,-4138) ;
ExcelPamka('A29:F29',2,2,-4138,-4138) ;
ExcelPamka('A30:F30',2,2,-4138,-4138) ;
ExcelPamka('A31:F31',2,2,-4138,-4138) ;
ExcelPamka('A32:F32',2,2,-4138,-4138) ;
form4.ProgressBar1.position:=form4.ProgressBar1.position+5 ;    //((((((())))))
ExcelPamka('A34:F34',2,2,-4138,-4138) ;
ExcelPamka('A35:F35',2,2,-4138,-4138) ;
ExcelPamka('A36:F36',2,2,-4138,-4138) ;
ExcelPamka('A37:F37',2,2,-4138,-4138) ;
ExcelPamka('A38:F38',2,2,-4138,-4138) ;
ExcelPamka('A39:F39',2,2,-4138,-4138) ;
ExcelPamka('A40:F40',2,2,-4138,-4138) ;
ExcelPamka('A41:F41',2,2,-4138,-4138) ;
ExcelPamka('A42:F42',2,2,-4138,-4138) ;

form4.ProgressBar1.position:=form4.ProgressBar1.position+5 ;    //((((((())))))

//������� �������
ExcelPamka('A2:F21',-4138,-4138,-4138,-4138) ;
ExcelPamka('A24:F32',-4138,-4138,-4138,-4138) ;
ExcelPamka('A34:F42',-4138,-4138,-4138,-4138) ;
form4.ProgressBar1.position:=form4.ProgressBar1.position+5 ;    //((((((())))))
ExcelPamka('A2:A21',-4138,-4138,-4138,-4138) ;
ExcelPamka('B2:B21',-4138,-4138,-4138,-4138) ;
ExcelPamka('C2:C21',-4138,-4138,-4138,-4138) ;
ExcelPamka('D2:D21',-4138,-4138,-4138,-4138) ;
ExcelPamka('E2:E21',-4138,-4138,-4138,-4138) ;
ExcelPamka('F2:F21',-4138,-4138,-4138,-4138) ;
form4.ProgressBar1.position:=form4.ProgressBar1.position+5 ;    //((((((())))))
ExcelPamka('A24:A32',-4138,-4138,-4138,-4138) ;
ExcelPamka('B24:B32',-4138,-4138,-4138,-4138) ;
ExcelPamka('C24:C32',-4138,-4138,-4138,-4138) ;
ExcelPamka('D24:D32',-4138,-4138,-4138,-4138) ;
ExcelPamka('E24:E32',-4138,-4138,-4138,-4138) ;
ExcelPamka('F24:F32',-4138,-4138,-4138,-4138) ;

ExcelPamka('A34:A42',-4138,-4138,-4138,-4138) ;
ExcelPamka('B34:B42',-4138,-4138,-4138,-4138) ;
ExcelPamka('C34:C42',-4138,-4138,-4138,-4138) ;
ExcelPamka('D34:D42',-4138,-4138,-4138,-4138) ;
ExcelPamka('E34:E42',-4138,-4138,-4138,-4138) ;
ExcelPamka('F34:F42',-4138,-4138,-4138,-4138) ;
form4.ProgressBar1.position:=form4.ProgressBar1.position+5 ;    //((((((())))))

  // ������� �� �����
ExcelApp.Range['B3:B21'].Select;
ExcelApp.Selection.WrapText:=True;
ExcelApp.Range['B24:B32'].Select;
ExcelApp.Selection.WrapText:=True;
ExcelApp.Range['B35:B42'].Select;
ExcelApp.Selection.WrapText:=True;

// ������������ � �����
ExcelApp.Range['A3:A32'].Select;
ExcelApp.Selection.HorizontalAlignment:=3;
ExcelApp.Selection.VerticalAlignment:=2 ;
ExcelApp.Selection.Font.Size := 11;
ExcelApp.Selection.Font.Name := 'Calibri';
ExcelApp.Range['B3:B32'].Select;
ExcelApp.Selection.HorizontalAlignment:=2;
ExcelApp.Selection.VerticalAlignment:=2 ;
ExcelApp.Selection.Font.Size := 11;
ExcelApp.Selection.Font.Name := 'Times New Roman';
ExcelApp.Range['D3:D32'].Select;
ExcelApp.Selection.HorizontalAlignment:=3;
ExcelApp.Selection.VerticalAlignment:=2 ;
ExcelApp.Selection.Font.Size := 13;
ExcelApp.Selection.Font.Name := 'Calibri';
ExcelApp.Range['E3:E32'].Select;
ExcelApp.Selection.HorizontalAlignment:=3;
ExcelApp.Selection.VerticalAlignment:=2 ;
ExcelApp.Selection.Font.Size := 11;
ExcelApp.Selection.Font.Name := 'Cambria';
ExcelApp.Range['F3:F32'].Select;
ExcelApp.Selection.HorizontalAlignment:=2;
ExcelApp.Selection.VerticalAlignment:=2 ;
ExcelApp.Selection.Font.Size := 11;
ExcelApp.Selection.Font.Name := 'Times New Roman';
///���������� �� ��������:
ExcelApp.Range['A34:A42'].Select;
ExcelApp.Selection.HorizontalAlignment:=3;
ExcelApp.Selection.VerticalAlignment:=2 ;
ExcelApp.Selection.Font.Size := 11;
ExcelApp.Selection.Font.Name := 'Calibri';
ExcelApp.Range['B34:B42'].Select;
ExcelApp.Selection.HorizontalAlignment:=2;
ExcelApp.Selection.VerticalAlignment:=2 ;
ExcelApp.Selection.Font.Size := 11;
ExcelApp.Selection.Font.Name := 'Times New Roman';
ExcelApp.Range['C34:C42'].Select;
ExcelApp.Selection.HorizontalAlignment:=3;
ExcelApp.Selection.VerticalAlignment:=2 ;
ExcelApp.Selection.Font.Size := 11;
ExcelApp.Selection.Font.Name := 'Cambria';
ExcelApp.Range['D34:D42'].Select;
ExcelApp.Selection.HorizontalAlignment:=2;
ExcelApp.Selection.VerticalAlignment:=2 ;
ExcelApp.Selection.Font.Size := 11;
ExcelApp.Selection.Font.Name := 'Times New Roman';
ExcelApp.Range['E34:E42'].Select;
ExcelApp.Selection.HorizontalAlignment:=2;
ExcelApp.Selection.VerticalAlignment:=2 ;
ExcelApp.Selection.Font.Size := 11;
ExcelApp.Selection.Font.Name := 'Times New Roman';
ExcelApp.Range['F34:F42'].Select;
ExcelApp.Selection.HorizontalAlignment:=2;
ExcelApp.Selection.VerticalAlignment:=2 ;
ExcelApp.Selection.Font.Size := 11;
ExcelApp.Selection.Font.Name := 'Times New Roman';
ExcelApp.Range['F34:F42'].Select;
ExcelApp.Selection.HorizontalAlignment:=2;
ExcelApp.Selection.VerticalAlignment:=2 ;
ExcelApp.Selection.Font.Size := 11;
ExcelApp.Selection.Font.Name := 'Times New Roman';



form4.ProgressBar1.position:=form4.ProgressBar1.position+5 ;    //((((((())))))

// ����� ��������

// ����������� ����� � �� ����� � �����
ExcelApp.Range['A1:F1'].Select;
ExcelApp.Selection.MergeCells:=True ;
ExcelApp.Selection.HorizontalAlignment:=3;
ExcelApp.Selection.VerticalAlignment:=2 ;
ExcelApp.Selection.Font.Size := 14;
ExcelApp.Selection.Font.Name := 'Times New Roman';
ExcelApp.Selection:='�������� ������'  ;

ExcelApp.Range['B2:C2'].Select;
ExcelApp.Selection.MergeCells:=True ;
ExcelApp.Range['B3:C3'].Select;
ExcelApp.Selection.MergeCells:=True ;
ExcelApp.Range['B4:C4'].Select;
ExcelApp.Selection.MergeCells:=True ;
ExcelApp.Range['B5:C5'].Select;
ExcelApp.Selection.MergeCells:=True ;
ExcelApp.Range['B6:C6'].Select;
ExcelApp.Selection.MergeCells:=True ;
ExcelApp.Range['B7:C7'].Select;
ExcelApp.Selection.MergeCells:=True ;
ExcelApp.Range['B8:C8'].Select;
ExcelApp.Selection.MergeCells:=True ;
ExcelApp.Range['B9:C9'].Select;
ExcelApp.Selection.MergeCells:=True ;
ExcelApp.Range['B10:C10'].Select;
ExcelApp.Selection.MergeCells:=True ;
ExcelApp.Range['B11:C11'].Select;
ExcelApp.Selection.MergeCells:=True ;
ExcelApp.Range['B12:C12'].Select;
ExcelApp.Selection.MergeCells:=True ;
ExcelApp.Range['B13:C13'].Select;
ExcelApp.Selection.MergeCells:=True ;
ExcelApp.Range['B14:C14'].Select;
ExcelApp.Selection.MergeCells:=True ;
ExcelApp.Range['B15:C15'].Select;
ExcelApp.Selection.MergeCells:=True ;
ExcelApp.Range['B16:C16'].Select;
ExcelApp.Selection.MergeCells:=True ;
ExcelApp.Range['B17:C17'].Select;
ExcelApp.Selection.MergeCells:=True ;
ExcelApp.Range['B18:C18'].Select;
ExcelApp.Selection.MergeCells:=True ;
ExcelApp.Range['B19:C19'].Select;
ExcelApp.Selection.MergeCells:=True ;
ExcelApp.Range['B20:C20'].Select;
ExcelApp.Selection.MergeCells:=True ;
ExcelApp.Range['B21:C21'].Select;
ExcelApp.Selection.MergeCells:=True ;

ExcelApp.Range['B24:C24'].Select;
ExcelApp.Selection.MergeCells:=True ;
ExcelApp.Range['B25:C25'].Select;
ExcelApp.Selection.MergeCells:=True ;
ExcelApp.Range['B26:C26'].Select;
ExcelApp.Selection.MergeCells:=True ;
ExcelApp.Range['B27:C27'].Select;
ExcelApp.Selection.MergeCells:=True ;
ExcelApp.Range['B28:C28'].Select;
ExcelApp.Selection.MergeCells:=True ;
ExcelApp.Range['B29:C29'].Select;
ExcelApp.Selection.MergeCells:=True ;
ExcelApp.Range['B30:C30'].Select;
ExcelApp.Selection.MergeCells:=True ;
ExcelApp.Range['B31:C31'].Select;
ExcelApp.Selection.MergeCells:=True ;
ExcelApp.Range['B32:C32'].Select;
ExcelApp.Selection.MergeCells:=True ;

ExcelApp.Range['B34:C34'].Select;
ExcelApp.Selection.MergeCells:=True ;

ExcelApp.Range['A22:F23'].Select;
ExcelApp.Selection.MergeCells:=True ;
ExcelApp.Selection.HorizontalAlignment:=3;
ExcelApp.Selection.VerticalAlignment:=2 ;
ExcelApp.Selection.Font.Size := 14;
ExcelApp.Selection.Font.Name := 'Times New Roman';
ExcelApp.Selection:='���������� �������';

 //������ �������
  // ���������
  ExcelApp.Range['A2:F2'].Select;
  ExcelApp.Selection.WrapText:=True;
  ExcelApp.Selection.HorizontalAlignment:=3;
  ExcelApp.Selection.VerticalAlignment:=2 ;
  ExcelApp.Selection.Font.Size := 11;
  ExcelApp.Selection.Font.Name := 'Times New Roman';
  ExcelApp.Range['A2']:='�';
  ExcelApp.Range['B2']:='��������';
  ExcelApp.Range['D2']:='�����������';
  ExcelApp.Range['E2']:='������� ���������';
  ExcelApp.Range['E2'].Select;
  ExcelApp.Selection.Font.Size := 10.5;
  ExcelApp.Range['F2']:='��������';

form4.ProgressBar1.position:=form4.ProgressBar1.position+5 ;    //((((((())))))

ExcelApp.Range['A3']:='1' ;
ExcelApp.Range['A4']:='2' ;
ExcelApp.Range['A5']:='3' ;
ExcelApp.Range['A6']:='4' ;
ExcelApp.Range['A7']:='5' ;
ExcelApp.Range['A8']:='6' ;
ExcelApp.Range['A9']:='7' ;
ExcelApp.Range['A10']:='8' ;
ExcelApp.Range['A11']:='9' ;
ExcelApp.Range['A12']:='10' ;
ExcelApp.Range['A13']:='11' ;
ExcelApp.Range['A14']:='12' ;
ExcelApp.Range['A15']:='13' ;
ExcelApp.Range['A16']:='14' ;
ExcelApp.Range['A17']:='15' ;
ExcelApp.Range['A18']:='16' ;
ExcelApp.Range['A19']:='17' ;
ExcelApp.Range['A20']:='18' ;
ExcelApp.Range['A21']:='19' ;

form4.ProgressBar1.position:=form4.ProgressBar1.position+5 ;    //((((((())))))

ExcelApp.Range['B3']:='���������������������� ����� �������'     ;
ExcelApp.Range['B4']:='�������� ����������� ����'                ;
ExcelApp.Range['B5']:='����������� ����������� ����'             ;
ExcelApp.Range['B6']:='����������� ����������� ����'             ;
ExcelApp.Range['B7']:='����������� �� �� ����� � ��'             ;
ExcelApp.Range['B8']:='�������� ������� �������'                 ;
ExcelApp.Range['B9']:='����������� �� �� ������ �� ��'           ;
ExcelApp.Range['B10']:='�������� �����'                          ;
ExcelApp.Range['B11']:='�������� ������� ������� ������'         ;
ExcelApp.Range['B12']:='������� ������ ������'                   ;
ExcelApp.Range['B13']:='��� ������� �������'                     ;
ExcelApp.Range['B14']:='����� ������������ ����'                 ;
ExcelApp.Range['B15']:='������ �� �� ��������� �����'            ;
ExcelApp.Range['B16']:='��� ����� �� �� �������'                 ;
ExcelApp.Range['B17']:='�������  ����� �� �� �������'            ;
ExcelApp.Range['B18']:='������� �������  ����� �� �� �������'    ;
ExcelApp.Range['B19']:='��� ����� �� �� �������'                 ;
ExcelApp.Range['B20']:='�������  ����� ��  ��  �������'          ;
ExcelApp.Range['B21']:='������� �������  ����� ��  ��  �������'  ;

form4.ProgressBar1.position:=form4.ProgressBar1.position+5 ;    //((((((())))))

ExcelApp.Range['D3']:='D���' ;
ExcelApp.Range['D4']:='���' ;
ExcelApp.Range['D5']:='t��' ;
ExcelApp.Range['D6']:='t��' ;
ExcelApp.Range['D7']:='t����' ;
ExcelApp.Range['D8']:='P1' ;
ExcelApp.Range['D9']:='t�����' ;
ExcelApp.Range['D10']:='' ;
ExcelApp.Range['D11']:='d�' ;
ExcelApp.Range['D12']:=Chr(100) + '��' ;
Start := 1;
Length := 1;
ExcelApp.Range['D12'].Characters[Start, Length].Font.Name := 'Symbol';
ExcelApp.Range['D13']:='S��' ;
ExcelApp.Range['D14']:='n��' ;
ExcelApp.Range['D15']:='' ;
ExcelApp.Range['D16']:='S����' ;
ExcelApp.Range['D17']:='d����' ;
ExcelApp.Range['D18']:=Chr(100) + '����' ;
Start := 1;
Length := 1;
ExcelApp.Range['D18'].Characters[Start, Length].Font.Name := 'Symbol';
ExcelApp.Range['D19']:='S����' ;
ExcelApp.Range['D20']:='d����' ;
ExcelApp.Range['D21']:=Chr(100) + '����' ;
Start := 1;
Length := 1;
ExcelApp.Range['D21'].Characters[Start, Length].Font.Name := 'Symbol';

form4.ProgressBar1.position:=form4.ProgressBar1.position+5 ;    //((((((())))))

ExcelApp.Range['E3']:='��/�' ;
ExcelApp.Range['E4']:='���' ;
ExcelApp.Range['E5']:=Chr(176) + '�' ;
Start := 1;
Length := 1;
ExcelApp.Range['E5'].Characters[Start, Length].Font.Name := 'Symbol';
ExcelApp.Range['E6']:=Chr(176) + '�' ;
Start := 1;
Length := 1;
ExcelApp.Range['E6'].Characters[Start, Length].Font.Name := 'Symbol';
ExcelApp.Range['E7']:=Chr(176) + '�' ;
Start := 1;
Length := 1;
ExcelApp.Range['E7'].Characters[Start, Length].Font.Name := 'Symbol';
ExcelApp.Range['E8']:='���' ;
ExcelApp.Range['E9']:=Chr(176) + '�' ;
Start := 1;
Length := 1;
ExcelApp.Range['E9'].Characters[Start, Length].Font.Name := 'Symbol';
ExcelApp.Range['E10']:=''   ;
ExcelApp.Range['E11']:='��' ;
ExcelApp.Range['E12']:='��' ;
ExcelApp.Range['E13']:='��' ;
ExcelApp.Range['E14']:=''   ;
ExcelApp.Range['E15']:=''   ;
ExcelApp.Range['E16']:='��' ;
ExcelApp.Range['E17']:='��' ;
ExcelApp.Range['E18']:='��' ;
ExcelApp.Range['E19']:='��' ;
ExcelApp.Range['E20']:='��' ;
ExcelApp.Range['E21']:='��' ;


 //������ �������
  //�����
ExcelApp.Range['A24']:='1' ;
ExcelApp.Range['A25']:='2' ;
ExcelApp.Range['A26']:='3' ;
ExcelApp.Range['A27']:='4' ;
ExcelApp.Range['A28']:='5' ;
ExcelApp.Range['A29']:='6' ;
ExcelApp.Range['A30']:='7' ;
ExcelApp.Range['A31']:='8' ;
ExcelApp.Range['A32']:='9' ;
ExcelApp.Range['A35']:='' ;
ExcelApp.Range['A34']:='' ;
ExcelApp.Range['A35']:='10' ;
ExcelApp.Range['A36']:='11' ;
ExcelApp.Range['A37']:='12' ;
ExcelApp.Range['A38']:='13' ;
ExcelApp.Range['A39']:='14' ;
ExcelApp.Range['A40']:='15' ;
ExcelApp.Range['A41']:='16' ;
ExcelApp.Range['A42']:='17' ;

ExcelApp.Range['B24']:='����������� �� �� ����� � �� ��-��' ;
ExcelApp.Range['B25']:='����������� �� �� ����� � �� ��-��' ;
ExcelApp.Range['B26']:='����������� 2� �� ����� � �� ��-��' ;
ExcelApp.Range['B27']:='������� �������� ��' ;
ExcelApp.Range['B28']:='������� ����������� ����������� ��' ;
ExcelApp.Range['B29']:='����� ������ ��������������' ;
ExcelApp.Range['B30']:='������ �� (1�)' ;
ExcelApp.Range['B31']:='������� �������� (1�)' ;
ExcelApp.Range['B32']:='������� �������� (2�)' ;
ExcelApp.Range['B33']:='' ;
ExcelApp.Range['B34']:='���������� �� ��������:' ;
ExcelApp.Range['B35']:='����������� ����������� (2�)' ;
ExcelApp.Range['B36']:='������������� �����' ;
ExcelApp.Range['B37']:='������ �������' ;
ExcelApp.Range['B38']:='����������� ������������� (2�)' ;
ExcelApp.Range['B39']:='������� �������� (2�)' ;
ExcelApp.Range['B40']:='������� �������' ;
ExcelApp.Range['B41']:='�������� (2�)' ;
ExcelApp.Range['B42']:='������������ ���������� ������� �� �������' ;

ExcelApp.Range['D24']:='t�����' ;
ExcelApp.Range['D25']:='t����' ;
ExcelApp.Range['D26']:='ts ��' ;
ExcelApp.Range['D27']:='W1' ;
ExcelApp.Range['D28']:=Chr(97) + '1' ;
Start := 1;
Length := 1;
ExcelApp.Range['D28'].Characters[Start, Length].Font.Name := 'Symbol';
ExcelApp.Range['D29']:='L��' ;
ExcelApp.Range['D30']:='G��' ;
ExcelApp.Range['D31']:=Chr(68) + '�1' ;
Start := 1;
Length := 1;
ExcelApp.Range['D31'].Characters[Start, Length].Font.Name := 'Symbol';
ExcelApp.Range['D32']:=Chr(68) + '���' ;
Start := 1;
Length := 1;
ExcelApp.Range['D32'].Characters[Start, Length].Font.Name := 'Symbol';

ExcelApp.Range['E24']:='o�' ;
ExcelApp.Range['E25']:='o�' ;
ExcelApp.Range['E26']:='o�' ;
ExcelApp.Range['E27']:='�2/�' ;
ExcelApp.Range['E28']:='��/(�2  o�)' ;
ExcelApp.Range['E29']:='�' ;
ExcelApp.Range['E30']:='��/�' ;
ExcelApp.Range['E31']:='���' ;
ExcelApp.Range['E32']:='���' ;

ExcelApp.Range['C35']:='��/(�2 ��)' ;
ExcelApp.Range['C36']:='��' ;
ExcelApp.Range['C37']:='�' ;
ExcelApp.Range['C38']:='���/(�2 ��)' ;
ExcelApp.Range['C39']:='���' ;
ExcelApp.Range['C40']:='�2' ;
ExcelApp.Range['C41']:='�/�' ;
ExcelApp.Range['C42']:='���' ;



 //������ �������
  // ���������
  ExcelApp.Range['A34:F34'].Select;
  ExcelApp.Selection.WrapText:=True;
  ExcelApp.Selection.HorizontalAlignment:=3;
  ExcelApp.Selection.VerticalAlignment:=2 ;
  ExcelApp.Selection.Font.Size := 11;
  ExcelApp.Selection.Font.Name := 'Times New Roman';
  ExcelApp.Range['D34']:='��';
  ExcelApp.Range['E34']:='��';
  ExcelApp.Range['F34']:='��';


form4.ProgressBar1.position:=form4.ProgressBar1.position+5 ;    //((((((())))))

// ������ � ������ �������� � �����
    //������
  ExcelApp.Range['A:A', EmptyParam].EntireColumn.ColumnWidth :=5.71 ;
  ExcelApp.Range['B:B', EmptyParam].EntireColumn.ColumnWidth :=30.71 ;
  ExcelApp.Range['C:C', EmptyParam].EntireColumn.ColumnWidth :=12 ;
  ExcelApp.Range['D:D', EmptyParam].EntireColumn.ColumnWidth :=12 ;
  ExcelApp.Range['E:E', EmptyParam].EntireColumn.ColumnWidth :=11.14 ;
  ExcelApp.Range['F:F', EmptyParam].EntireColumn.ColumnWidth :=10 ;
  //������   (������ ����!!!!)
 ExcelApp.Range['A1:A50', EmptyParam].EntireRow.autofit ;
     /// ����������
      ExcelApp.Range['1:1', EmptyParam].EntireRow.RowHeight :=25.5 ;
      ExcelApp.Range['2:2', EmptyParam].EntireRow.RowHeight :=27.75 ;
      ExcelApp.Range['3:21', EmptyParam].EntireRow.RowHeight :=17.25 ;
      ExcelApp.Range['22:22', EmptyParam].EntireRow.RowHeight :=8.25 ;
      ExcelApp.Range['23:23', EmptyParam].EntireRow.RowHeight :=17.25 ;
      ExcelApp.Range['24:32', EmptyParam].EntireRow.RowHeight :=17.25 ;
      ExcelApp.Range['33:33', EmptyParam].EntireRow.RowHeight :=16.5 ;
      ExcelApp.Range['34:34', EmptyParam].EntireRow.RowHeight :=27 ;
      ExcelApp.Range['35:42', EmptyParam].EntireRow.RowHeight :=17.25 ;
      ExcelApp.Range['42:42', EmptyParam].EntireRow.RowHeight :=30.75 ;

// ����������� �������
UpDownSymbol('D3:D21',777,777,2,6) ;
UpDownSymbol('D24:D30',777,777,2,6) ;
UpDownSymbol('D31:D32',777,777,3,6) ;

// ����������� �������
UpDownSymbol('D7:D7',4,6,777,777) ;
UpDownSymbol('D9:D9',5,6,777,777) ;
UpDownSymbol('D16:D21',4,6,777,777) ;
UpDownSymbol('D24:D24',5,7,777,777) ;
UpDownSymbol('D25:D25',4,7,777,777) ;
UpDownSymbol('D16:D21',4,6,777,777) ;
UpDownSymbol('E24:E26',0,1,777,777) ;
UpDownSymbol('E27:E27',2,1,777,777) ;
UpDownSymbol('E28:E28',6,9,777,777) ;
ExcelApp.Range['E28:E28'].Characters[10, 13].Font.Superscript:= False  ;
UpDownSymbol('C35:C35',6,9,777,777) ;
ExcelApp.Range['C35:C35'].Characters[9, 14].Font.Superscript:= False  ;
UpDownSymbol('C36:C36',0,1,777,777) ;
UpDownSymbol('C38:C38',7,10,777,777) ;
ExcelApp.Range['C38:C38'].Characters[10, 14].Font.Superscript:= False  ;
UpDownSymbol('C40:C40',2,5,777,777) ;

//������ ������� ��� �� ������
// ������ �� ������ ���� ��������
ExcelApp.Workbooks[1].WorkSheets[1].PageSetup.LeftMargin := ExcelApp.Workbooks[1].WorkSheets[1].Application.CentimetersToPoints(1.5); //(1.5 �� )
// ������ �� ������� ���� ��������
ExcelApp.Workbooks[1].WorkSheets[1].PageSetup.RightMargin:= ExcelApp.Workbooks[1].WorkSheets[1].Application.CentimetersToPoints(1.5); //(1.5 �� )
// ������ �� �������� ���� ��������.
ExcelApp.Workbooks[1].WorkSheets[1].PageSetup.TopMargin :=  ExcelApp.Workbooks[1].WorkSheets[1].Application.CentimetersToPoints(1.5); //(1.5 �� )
// ������ �� ������� ���� ��������.
ExcelApp.Workbooks[1].WorkSheets[1].PageSetup.BottomMargin:=ExcelApp.Workbooks[1].WorkSheets[1].Application.CentimetersToPoints(1.5); //(1.5 �� )
// ������ �������� ����������� ��������.
ExcelApp.Workbooks[1].WorkSheets[1].PageSetup.HeaderMargin:=ExcelApp.Workbooks[1].WorkSheets[1].Application.CentimetersToPoints(0); //(0 �� )
// ������ ������� ����������� ��������.
ExcelApp.Workbooks[1].WorkSheets[1].PageSetup.FooterMargin:=ExcelApp.Workbooks[1].WorkSheets[1].Application.CentimetersToPoints(0); //(0 ��)



form4.ProgressBar1.position:=form4.ProgressBar1.position+5 ;    //((((((())))))





// ��������� �������� ������
ExcelApp.Range['A2'].Select;  //(����� ������ ���������� �������)

form4.ProgressBar1.position:=form4.ProgressBar1.position+5 ;    //((((((())))))

  // ��������� (������� ���)
  if (CheckBox8.Checked=True) and (CheckBox3.Checked=True) then
    begin
       SaveDialog2.Title:='���������� �����������' ;
       SaveDialog2.filename := ExtractFilePath( Application.ExeName )+'�����';
       if SaveDialog2.Execute then
         try  //(/�������� ������)
           ExcelApp.ActiveSheet.SaveAs(SaveDialog2.filename);
         except
         end;
    end;
// ������ Excel ������� (�������)
 if CheckBox3.Checked=False  then
 begin
   ExcelApp.Visible := False ;
   CheckBox1.Checked:=False  ;
   end
 else
   ExcelApp.Visible := True ;

  // ��������� (������� ���� � ������ ����)
  if (CheckBox8.Checked=True) and (CheckBox3.Checked=False) and (CheckBox2.Checked=False) then
    begin
       SaveDialog2.Title:='���������� �����������' ;
       SaveDialog2.filename := ExtractFilePath( Application.ExeName )+'�����';
       if SaveDialog2.Execute then
         begin
             try  //(/�������� ������)
                ExcelApp.ActiveSheet.SaveAs(SaveDialog2.filename);
             except
             end;
        end;
             ExcelApp.DisplayAlerts := False;
             ExcelApp.Quit; // ������� Excel
             ExcelApp := Unassigned;
    end;

  // ��������� (������� ���� � ������ ���)
  if (CheckBox8.Checked=True) and (CheckBox3.Checked=False) and (CheckBox2.Checked=True)  then
    begin
       SaveDialog2.Title:='���������� �����������' ;
       SaveDialog2.filename := ExtractFilePath( Application.ExeName )+'�����';
       if SaveDialog2.Execute then
         begin
             try  //(/�������� ������)
                ExcelApp.ActiveSheet.SaveAs(SaveDialog2.filename);
             except
             end;
        end;
    end;
  // ������
   if (CheckBox2.Checked=True) and (CheckBox3.Checked=True) then
      CheckBox1.Enabled:=True  ;
   if (CheckBox2.Checked=False) and (CheckBox3.Checked=False) then
      CheckBox1.Enabled:=False  ;
   if (CheckBox2.Checked=True) and (CheckBox1.Checked=False) then
     begin
       CheckBox1.Enabled:=True;
       ExcelApp.ActiveSheet.PrintOut  ;
       ExcelApp.DisplayAlerts := False;
       ExcelApp.Quit; // ������� Excel
       ExcelApp := Unassigned;
     end;
   if  (CheckBox2.Checked=true)  and (CheckBox1.Checked=true)  then
           begin
      Form4.close;
    form4.ProgressBar1.position:=0 ;
       ExcelApp.Dialogs[222].Show ;
          end;


   Form4.close;
   form4.ProgressBar1.position:=0 ;
 Day ; ///����
   EnD ; ///[[{{{��������� Excel}}}]]  (((�����������)))






//       [[{{{��������� Excel}}}]]
    //           (((������)))

// �������� �� �� ������� ������������� Excel
      // ���� CLSID OLE-�������
  Rez := CLSIDFromProgID(PWideChar(WideString('Excel.Application')), ClassID);
  if Rez = S_OK then  // ������ ������
    Result := true
  else
     begin
    Result := false;
    ShowMessage('� ������ Excel ��������, � ����� �� ������� �����.') ;
    Day ; Exit;
     end;


if (CheckBox6.Checked=true) and (CheckBox10.Checked=true) then
  BeGiN ///[[{{{��������� Excel}}}]] (((�����������)))
// ���������� ������ �������� ���� �������, � ������� ����� �������� ������
  BeginCol := 1;
  BeginRow := 4;

// ������� ���������� ������� ������
  // �������
  RowCount := 300;
  // �������
  ColCount := 5;

// �������� Excel
  ExcelApp := CreateOleObject('Excel.Application');

form4.ProgressBar1.position:=form4.ProgressBar1.position+5 ;    //((((((())))))

// ��������� ������� Excel �� �������, ����� �������� ����� ����������
  ExcelApp.Application.EnableEvents := false;



 //  ������� ����� (Workbook)
 //  ���� ��������� ������, ��
 //  Workbook := ExcelApp.WorkBooks.Add('C:\MyTemplate.xls');
OpenDialog1.filter:= ' ���� Excel (*.xls)  |*.xls| ��� ����� (*.*) |*.*| ';
OpenDialog1.Title:='����� �������' ;
//if OpenDialog1.Execute then
//  try  //(/�������� ������)
//    Workbook := ExcelApp.WorkBooks.Add(OpenDialog1.filename);
 Workbook := ExcelApp.WorkBooks.Add();
// ������� ���������� ������, ������� �������� ��������� �������
ArrayData := VarArrayCreate([1, RowCount, 1, ColCount], varVariant);

form4.ProgressBar1.position:=form4.ProgressBar1.position+5 ;    //((((((())))))

//����������
ftp_th :=RoundSignificant(ftp_th ,5);              ////            ftp_th 
fnpox_th:=RoundSignificant(fnpox_th,5);            ////            fnpox_th
d_BH:=RoundSignificant(d_BH,5);                    ////            d_BH
f_BH:=RoundSignificant(f_BH,5);                    ////            f_BH
fnpox_2:=RoundSignificant(fnpox_2,5);              ////            fnpox_2
t_cp_TH:=RoundSignificant(t_cp_TH,5);              ////            t_cp_TH
V_TH:=RoundSignificant(V_TH,5);                    ////            V_TH
i_Bx_TH:=RoundSignificant(i_Bx_TH,5);              ////            i_Bx_TH
i_Bix_TH:=RoundSignificant(i_Bix_TH,5);            ////            i_Bix_TH
Delta_P_nn_1:=RoundSignificant(Delta_P_nn_1,5);        ////            Delta_P_nn
Delta_P_icn_1:=RoundSignificant(Delta_P_icn_1,5);      ////            Delta_P_icp
Delta_P_ak_1:=RoundSignificant(Delta_P_ak_1,5);        ////            Delta_P_ak
P_icn:=RoundSignificant(P_icn,5);                  ////            P_icp
P_ak:=RoundSignificant(P_ak,5);                    ////            P_ak
P_BX:=RoundSignificant(P_BX,5);                    ////            P_BX
P_ak_cp:=RoundSignificant(P_ak_cp,5);              ////            P_ak_cp
P_icn_cp:=RoundSignificant(P_icn_cp,5);            ////            P_icp_cp
P_nn_cp:=RoundSignificant(P_nn_cp,5);              ////            P_nn_cp
t_s_icn:=RoundSignificant(t_s_icn,5);              ////            t_s_icp
t_s_ak:=RoundSignificant(t_s_ak,5);                ////            t_s_ak
iYY:=RoundSignificant(iYY,5);                      ////            iYY
iY_s:=RoundSignificant(iY_s,5);                    ////            iY_s
i_nB:=RoundSignificant(i_nB,5);                    ////            i_nB
ine:=RoundSignificant(ine,5);                      ////            ine
t_cp_ak:=RoundSignificant(t_cp_ak,5);              ////            t_cp_ak
t_cp_nn:=RoundSignificant(t_cp_nn,5);              ////            t_cp_nn
t_cp_icn:=RoundSignificant(t_cp_icn,5);            ////            t_cp_icp
Q_ak:=RoundSignificant(Q_ak,5);                    ////            Q_ak
Q_icn:=RoundSignificant(Q_icn,5);                ////            Q_icp:
form4.ProgressBar1.position:=form4.ProgressBar1.position+5 ;    //((((((())))))
Q_nn:=RoundSignificant(Q_nn,5);                    ////            Q_nn
Q_nr:=RoundSignificant(Q_nr,5);                    ////            Q_nr
G_TH:=RoundSignificant(G_TH,5);                    ////            G_TH
i_icn_TH:=RoundSignificant(i_icn_TH,5);            ////            i_icp_TH
i_ak_TH:=RoundSignificant(i_ak_TH,5);              ////            i_ak_TH
t_icn_TH:=RoundSignificant(t_icn_TH,5);            ////            t_icp_TH
t_ak_TH:=RoundSignificant(t_ak_TH,5);              ////            t_ak_TH
t_cp_TH_ne:=RoundSignificant(t_cp_TH_ne,5);          ////            t_cp_TH_n
t_cp_TH_icn:=RoundSignificant(t_cp_TH_icn,5);      ////            t_cp_TH_icp
t_cp_ct_nn:=RoundSignificant(t_cp_ct_nn,5);        ////            t_cp_ct_nn
t_cp_ct_icn:=RoundSignificant(t_cp_ct_icn,5);      ////            t_cp_ct_icp
t_cp_ct_ak:=RoundSignificant(t_cp_ct_ak,5);        ////            t_cp_ct_ak
W_1:=RoundSignificant(W_1,5);                      ////            W_1
d_ek_1:=RoundSignificant(d_ek_1,5);                ////            d_ek_1
Lambda_TH:=RoundSignificant(Lambda_TH,5);          ////            Lambda_TH
Lambda_TH_ct:=RoundSignificant(Lambda_TH_ct,5);    ////            Lambda_TH_ct
Mu_TH:=RoundSignificant(Mu_TH,5);                  ////            Mu_TH
Mu_TH_ct:=RoundSignificant(Mu_TH_ct,5);            ////            Mu_TH_ct
Re_1:=RoundSignificant(Re_1,5);                    ////            Re_1
Pr_TH:=RoundSignificant(Pr_TH,5);                  ////            Pr_TH
Pr_TH_ct:=RoundSignificant(Pr_TH_ct,5);              ////            Pr_TH_c
Nu_1:=RoundSignificant(Nu_1,5);                    ////            Nu_1
V_ak_cp:=RoundSignificant(V_ak_cp,5);              ////            V_ak_cp
f_ak_TP:=RoundSignificant(f_ak_TP,5);              ////            f_ak_TP
f_ak:=RoundSignificant(f_ak,5);                    ////            f_ak
W_ak:=RoundSignificant(W_ak,5);                    ////            W_ak
Mu_cp_ak:=RoundSignificant(Mu_cp_ak,5);            ////            Mu_cp_ak
Mu_cp_ct_ak:=RoundSignificant(Mu_cp_ct_ak,5);      ////            Mu_cp_ct_ak
KINV_cp_ak:=RoundSignificant(KINV_cp_ak,5);        ////            KINV_cp_ak
d_a_ak:=RoundSignificant(d_a_ak,5);                ////            d_a_ak
Re_ak:=RoundSignificant(Re_ak,5);                  ////            Re_ak
Lambda_ak:=RoundSignificant(Lambda_ak,5);          ////            Lambda_ak
Lambda_ct_ak:=RoundSignificant(Lambda_ct_ak,5);      ////            Lambda_ct_a
Pr_ak:=RoundSignificant(Pr_ak,5);                  ////            Pr_ak
Pr_ct_ak:=RoundSignificant(Pr_ct_ak,5);            ////            Pr_ct_ak
D_ak:=RoundSignificant(D_ak,5);                    ////            D_ak
Psi_ak:=RoundSignificant(  Psi_ak,5);              ////              Psi_ak
B_ak:=RoundSignificant(  B_ak,5);                  ////              B_ak
m_ak:=RoundSignificant(  m_ak,5);                  ////              m_ak
Nu_ak:=RoundSignificant(Nu_ak,5);                    ////            Nu_a
a_ak :=RoundSignificant(a_ak ,5);                  ////            a_ak
t_cp_ct_ak:=RoundSignificant(t_cp_ct_ak,5);        ////            t_cp_ct_ak
Lambda_MAT_ak:=RoundSignificant(Lambda_MAT_ak,5);  ////            Lambda_MAT_ak
Lambda_MAT_ak:=RoundSignificant(Lambda_MAT_ak,5);  ////            Lambda_MAT_ak
K_ak:=RoundSignificant(K_ak,5);                    ////            K_ak
Delta_t_ak:=RoundSignificant(Delta_t_ak,5);        ////            Delta_t_ak
F__ak:=RoundSignificant(F__ak,5);                  ////            F__ak
L_ak:=RoundSignificant(L_ak,5);                    ////            L_ak
V_nn_cp:=RoundSignificant(V_nn_cp,5);              ////            V_nn_cp
Mu_cp_nn:=RoundSignificant(Mu_cp_nn,5);            ////            Mu_cp_nn
Mu_cp_ct_nn:=RoundSignificant(Mu_cp_ct_nn,5);      ////            Mu_cp_ct_nn
KINV_cp_nn:=RoundSignificant(KINV_cp_nn,5);        ////            KINV_cp_nn
f_nn_TP:=RoundSignificant(f_nn_TP,5);              ////            f_nn_TP
f_nn:=RoundSignificant(f_nn,5);                    ////            f_nn
W_nn:=RoundSignificant(W_nn,5);                    ////            W_nn
d_a_nn:=RoundSignificant(d_a_nn,5);                ////            d_a_nn
Re_nn:=RoundSignificant(Re_nn,5);                  ////            Re_nn
Lambda_nn:=RoundSignificant(Lambda_nn,5);          ////            Lambda_nn
Lambda_ct_nn:=RoundSignificant(Lambda_ct_nn,5);    ////            Lambda_ct_nn
Pr_nn:=RoundSignificant(Pr_nn,5);                  ////            Pr_nn
Pr_ct_nn:=RoundSignificant(Pr_ct_nn,5);            ////            Pr_ct_nn
D_nn:=RoundSignificant(D_nn,5);                    ////            D_nn
Psi_nn:=RoundSignificant(Psi_nn,5);                ////            Psi_nn
B_nn:=RoundSignificant(  B_nn,5);                  ////              B_nn
m_nn:=RoundSignificant(  m_nn,5);                  ////              m_nn
Nu_nn:=RoundSignificant(Nu_nn,5);                  ////            Nu_nn
a_nn :=RoundSignificant(a_nn ,5);                  ////            a_nn
Lambda_MAT_nn:=RoundSignificant(Lambda_MAT_nn,5);  ////            Lambda_MAT_nn
t_nn_TH:=RoundSignificant(t_nn_TH,5);              ////            t_nn_TH
Lambda_MAT_nn:=RoundSignificant(Lambda_MAT_nn,5);  ////            Lambda_MAT_nn
Lambda_MAT_nn:=RoundSignificant(Lambda_MAT_nn,5);  ////            Lambda_MAT_nn
K_nn:=RoundSignificant(K_nn,5);                    ////            K_nn
Delta_t_nn:=RoundSignificant(Delta_t_nn,5);        ////            Delta_t_nn
F__nn:=RoundSignificant(F__nn,5);                  ////            F__nn
L_nn:=RoundSignificant(L_nn,5);                    ////            L_nn
q_1:=RoundSignificant(q_1,5);                      ////            q_1
a_icn:=RoundSignificant(a_icn,5);                  ////            a_icp
t_cp_ct_icn:=RoundSignificant(t_cp_ct_icn,5);      ////            t_cp_ct_icp
Lambda_MAT_icn:=RoundSignificant(Lambda_MAT_icn,5);////            Lambda_MAT_icp
Lambda_MAT_icn:=RoundSignificant(Lambda_MAT_icn,5);////            Lambda_MAT_icp
K_icn:=RoundSignificant(K_icn,5);                  ////            K_icp
Delta_t_icn:=RoundSignificant(Delta_t_icn,5);      ////            Delta_t_icp
q_2:=RoundSignificant(q_2,5);                      ////            q_2
Prov_q:=RoundSignificant(Prov_q,5);                    ////            Prov
F__icn:=RoundSignificant(F__icn,5);                ////            F__icp
L_nn:=RoundSignificant(L_nn,5);                    ////            L_nn
L_nr:=RoundSignificant(L_nr,5);                    ////            L_nr
Zeta_1:=RoundSignificant(Zeta_1,5);                ////            Zeta_1
Delta_P_1_2:=RoundSignificant(Delta_P_1_2,5);      ////            Delta_P_1_2
c_ak:=RoundSignificant(c_ak,5);                    ////            c_ak
Zeta_ak:=RoundSignificant(Zeta_ak,5);              ////            Zeta_ak
Delta_P_ak_2:=RoundSignificant(Delta_P_ak_2,5);    ////            Delta_P_ak_2
V_cp_YY:=RoundSignificant(V_cp_YY,5);              ////            V_cp_YY
V_cp_Y:=RoundSignificant(V_cp_Y,5);                ////            V_cp_Y
W_0:=RoundSignificant(W_0,5);                      ////            W_0
Mu_cp_icn:=RoundSignificant(Mu_cp_icn,5);          ////            Mu_cp_icp
KINV_cp_icn:=RoundSignificant(KINV_cp_icn,5);      ////            KINV_cp_icp
d_ek_icn:=RoundSignificant(d_ek_icn,5);            ////            d_ek_icp
Re_icn:=RoundSignificant(Re_icn,5);                ////            Re_icp
Zeta_icn:=RoundSignificant(Zeta_icn,5);            ////            Zeta_icp
Delta_P_icn_2:=RoundSignificant(Delta_P_icn_2,5);  ////            Delta_P_icp_2
c_nn:=RoundSignificant(c_nn,5);                    ////            c_nn
Zeta_nn:=RoundSignificant(Zeta_nn,5);              ////            Zeta_nn
Delta_P_nn_2:=RoundSignificant(Delta_P_nn_2,5);    ////            Delta_P_nn_2
Delta_P_nr_2:=RoundSignificant(Delta_P_nr_2,5);    ////            Delta_P_nr_2
Prov_P_nr:=RoundSignificant(Prov_P_nr,5);          ////            Prov_P_nr






form4.ProgressBar1.position:=form4.ProgressBar1.position+5 ;    //((((((())))))

 // ��������� ������
// �������� ������
      ArrayData[1, 5] := Dpac;
      ArrayData[2, 5] := Pne;
      ArrayData[3, 5] := tne;
      ArrayData[4, 5] := tnb;
      ArrayData[5, 5] := t_Bx;
      ArrayData[6, 5] := P1;
      ArrayData[7, 5] := t_Bix;
      ArrayData[8, 5] := MAT_TP;
      ArrayData[9, 5] := d_H;
      ArrayData[10, 5] := btp;
      ArrayData[11, 5] := Stp;
      ArrayData[12, 5] := n_TP;
      ArrayData[13, 5] := XOT_ShH;
      ArrayData[14, 5] := S_shH_nn;
      ArrayData[15, 5] := d_shH_nn;
      ArrayData[16, 5] := b_shH_nn;
      ArrayData[17, 5] := S_shH_ak;
      ArrayData[18, 5] := d_shH_ak;
      ArrayData[19, 5] := b_shH_ak;

//������� ������
  //������ �������

ArrayData[23,1] :=  '1' ;
ArrayData[24,1] :=  '2' ;
ArrayData[25,1] :=  '3' ;
ArrayData[26,1] :=  '4' ;
ArrayData[27,1] :=  '5' ;
ArrayData[28,1] :=  '6' ;
ArrayData[29,1] :=  '7' ;
ArrayData[31,1] :=  '8' ;
ArrayData[32,1] :=  '9' ;
ArrayData[33,1] :=  '10' ;
ArrayData[45,1] :=  '11' ;
ArrayData[53,1] :=  '12' ;
ArrayData[57,1] :=  '13' ;
ArrayData[62,1] :=  '14' ;
ArrayData[77,1] :=  '15' ;
ArrayData[78,1] :=  '16' ;
ArrayData[79,1] :=  '17' ;
ArrayData[88,1] :=  '18' ;
ArrayData[91,1] :=  '19' ;
ArrayData[92,1] :=  '20' ;
ArrayData[94,1] :=  '21' ;
ArrayData[95,1] :=  '22' ;
ArrayData[96,1] :=  '23' ;
ArrayData[97,1] :=  '24' ;
ArrayData[98,1] :=  '25' ;
ArrayData[102,1] :=  '26' ;
ArrayData[103,1] :=  '27' ;
ArrayData[107,1] :=  '28' ;
ArrayData[110,1] :=  '29' ;
ArrayData[116,1] :=  '30' ;
ArrayData[118,1] :=  '31' ;
ArrayData[124,1] :=  '32' ;
ArrayData[125,1] :=  '33' ;
ArrayData[126,1] :=  '34' ;
ArrayData[128,1] :=  '35' ;
ArrayData[129,1] :=  '36' ;
ArrayData[133,1] :=  '37' ;
ArrayData[134,1] :=  '38' ;
ArrayData[135,1] :=  '39' ;
ArrayData[136,1] :=  '40' ;
ArrayData[137,1] :=  '41' ;
ArrayData[138,1] :=  '42' ;
ArrayData[146,1] :=  '43' ;
ArrayData[152,1] :=  '44' ;
ArrayData[154,1] :=  '45' ;
ArrayData[160,1] :=  '46' ;
ArrayData[161,1] :=  '47' ;
ArrayData[162,1] :=  '48' ;
ArrayData[164,1] :=  '49' ;
ArrayData[165,1] :=  '50' ;
ArrayData[166,1] :=  '51' ;
ArrayData[172,1] :=  '52' ;
ArrayData[173,1] :=  '53' ;
ArrayData[174,1] :=  '54' ;
ArrayData[175,1] :=  '55' ;
ArrayData[176,1] :=  '56' ;
ArrayData[178,1] :=  '57' ;
ArrayData[180,1] :=  '59' ;
ArrayData[181,1] :=  '60' ;
ArrayData[184,1] :=  '61' ;
ArrayData[185,1] :=  '62' ;
ArrayData[195,1] :=  '63' ;
ArrayData[197,1] :=  '64' ;
ArrayData[200,1] :=  '65' ;
ArrayData[201,1] :=  '66' ;
ArrayData[202,1] :=  '67' ;

form4.ProgressBar1.position:=form4.ProgressBar1.position+5 ;    //((((((())))))

  //������ �������
ArrayData[23,2] :=  ' ������� ������� ������� ��� �����  ������' ;
ArrayData[24,2] :=  ' ������� ���������� ������� �� ������� �������' ;
ArrayData[25,2] :=  '���������� ������� ������' ;
ArrayData[26,2] :=  ' ���������� ������� ������� ������' ;
ArrayData[27,2] :=  ' ������� ������� ���� ������ �� ������� �������' ;
ArrayData[28,2] :=  '������� ����������� ��' ;
ArrayData[29,2] :=  '�������� ����� ��' ;
ArrayData[30,2] :=  '���������������� ��' ;
ArrayData[31,2] :=  '��������� �� �� ����� � ��' ;
ArrayData[32,2] :=  ' ��������� �� �� ������ �� ��' ;
ArrayData[33,2] :=  '��������' ;
ArrayData[38,2] :=  '�) �� ������ �� ����������' ;
ArrayData[39,2] :=  '�) �� ������ �� ��' ;
ArrayData[40,2] :=  '�)�� �����' ;
ArrayData[41,2] :=  '�) ������� �������� �:' ;
ArrayData[42,2] :=  '�      ��' ;
ArrayData[43,2] :=  '�      ���' ;
ArrayData[44,2] :=  '�      ��' ;
ArrayData[45,2] :=  '����������� � ���������' ;
ArrayData[46,2] :=  '�) �����������' ;
ArrayData[48,2] :=  '�) ���������' ;
ArrayData[49,2] :=  '�         ' ;
ArrayData[50,2] :=  '�         ' ;
ArrayData[51,2] :=  '�         ' ;
ArrayData[52,2] :=  '�         ' ;
ArrayData[53,2] :=  '������� ����������� �� �������� �� � ��' ;
ArrayData[54,2] :=  '�) ' ;
ArrayData[55,2] :=  '�) ' ;
ArrayData[56,2] :=  '�) ' ;
ArrayData[57,2] :=  '���������� ����� �� �������� ��������������' ;
ArrayData[62,2] :=  '����������� � ��������� ��' ;
ArrayData[63,2] :=  '������ �� ����� ��' ;
ArrayData[64,2] :=  '�) ���������' ;
ArrayData[65,2] :=  '�           ' ;
ArrayData[66,2] :=  '�           ' ;
ArrayData[67,2] :=  '�) �����������' ;
ArrayData[68,2] :=  '�           ' ;
ArrayData[69,2] :=  '�           ' ;
ArrayData[70,2] :=  '�) ������� ����������� ��' ;
ArrayData[71,2] :=  '�           ' ;
ArrayData[72,2] :=  '�           ' ;
ArrayData[73,2] :=  '�) ������� ����������� ������' ;
ArrayData[74,2] :=  '�           ' ;
ArrayData[75,2] :=  '�           ' ;
ArrayData[76,2] :=  '�           ' ;
ArrayData[77,2] :=  '�������� �� ������� ������� ����� �������������' ;
ArrayData[78,2] :=  '������������� ������� ����� ������ ��� ��' ;
ArrayData[80,2] :=  '���������������� �� � ������' ;
ArrayData[81,2] :=  '�         ' ;
ArrayData[82,2] :=  '�         ' ;
ArrayData[83,2] :=  '������������ �������� �� � ������' ;
ArrayData[84,2] :=  '�         ' ;
ArrayData[85,2] :=  '�         ' ;
ArrayData[86,2] :=  '�������������� �������� ��' ;
ArrayData[87,2] :=  '����� ���������� ��� ��' ;
ArrayData[88,2] :=  ' ����� �������� �� � ������' ;
ArrayData[89,2] :=  '�         ' ;
ArrayData[90,2] :=  '�         ' ;
ArrayData[91,2] :=  '����� ��������� ��� ��' ;
ArrayData[92,2] :=  '������������ �����������' ;
ArrayData[93,2] :=  '�����������' ;
ArrayData[94,2] :=  '�������� ����� � ��' ;
ArrayData[95,2] :=  ' ������� ���� ��' ;
ArrayData[96,2] :=  ' ������� ���� ���� � ��' ;
ArrayData[97,2] :=  '�������� � ��' ;
ArrayData[98,2] :=  '������������ �������� �� � ������' ;
ArrayData[99,2] :=  '�         ' ;
ArrayData[100,2] :=  '�         ' ;
ArrayData[101,2] :=  '�������������� �������� ��' ;
ArrayData[102,2] :=  '������������� ������� ��' ;
ArrayData[103,2] :=  ' ����� ���������� ��� ��' ;
ArrayData[104,2] :=  '���������������� ��' ;
ArrayData[105,2] :=  '�         ' ;
ArrayData[106,2] :=  '�         ' ;
ArrayData[107,2] :=  ' ����� �������� ��� �� � ������' ;
ArrayData[108,2] :=  '�         ' ;
ArrayData[109,2] :=  '�         ' ;
ArrayData[111,2] :=  '������� �������� ��������� ������' ;
ArrayData[112,2] :=  '�������������' ;
ArrayData[113,2] :=  '����������� ����������� ������' ;
ArrayData[114,2] :=  '����� ������� ������� �������� ���������' ;
ArrayData[115,2] :=  '����� ��������� ��� ��' ;
ArrayData[117,2] :=  '����������� ����������� ��' ;
ArrayData[119,2] :=  '������� ����������� ������ ��' ;
ArrayData[120,2] :=  '���������������� ��������� ������ �� (�����,�����)' ;
ArrayData[123,2] :=  '����������� �������������' ;
ArrayData[124,2] :=  '��������������� ��������� ���������� � ��' ;
ArrayData[125,2] :=  ' ������� ����������� ��' ;
ArrayData[126,2] :=  '������ ��' ;
ArrayData[127,2] :=  '�����������������' ;
ArrayData[128,2] :=  ' �������� ����� � ��' ;
ArrayData[129,2] :=  ' ������������ �������� �� � ������' ;
ArrayData[130,2] :=  '�         ' ;
ArrayData[131,2] :=  '�         ' ;
ArrayData[132,2] :=  '�������������� �������� ��' ;
ArrayData[133,2] :=  ' ������� ���� ��' ;
ArrayData[134,2] :=  ' ������� ���� ���� � ��' ;
ArrayData[135,2] :=  ' �������� � ��' ;
ArrayData[136,2] :=  ' ������������� ������� ��' ;
ArrayData[137,2] :=  ' ����� ���������� ��� ��' ;
ArrayData[139,2] :=  '���������������� �� � ������' ;
ArrayData[140,2] :=  '�         ' ;
ArrayData[142,2] :=  '�         ' ;
ArrayData[143,2] :=  '����� �������� ��� �� � ������' ;
ArrayData[144,2] :=  '�         ' ;
ArrayData[145,2] :=  '�         ' ;
ArrayData[147,2] :=  '������� �������� ��������� ������' ;
ArrayData[148,2] :=  '�������������' ;
ArrayData[149,2] :=  '����������� ����������� ������' ;
ArrayData[150,2] :=  '����� ������� ������� �������� ���������' ;
ArrayData[151,2] :=  '����� ��������� ��� ��' ;
ArrayData[153,2] :=  '����������� ����������� ��' ;
ArrayData[155,2] :=  '���������������� ��������� ������ �� ' ;
ArrayData[156,2] :=  '����������� �� � ��' ;
ArrayData[157,2] :=  '���������������� ��������� ������ �� (�����,�����)' ;
ArrayData[159,2] :=  '����������� ������������� ��' ;
ArrayData[160,2] :=  '��������������� ��������� ���������� � ��' ;
ArrayData[161,2] :=  ' ������� ����������� ��' ;
ArrayData[162,2] :=  '������ ��' ;
ArrayData[163,2] :=  '����������' ;
ArrayData[164,2] :=  '�������' ;
ArrayData[165,2] :=  '50) ����������� ����������� ���' ;
ArrayData[167,2] :=  '������� ����������� ������ ��' ;
ArrayData[168,2] :=  '���������������� ��������� ������ �� (�����,�����)' ;
ArrayData[171,2] :=  '����������� ������������� ���' ;
ArrayData[172,2] :=  ' ��������������� ��������� ���������� � ���' ;
ArrayData[174,2] :=  '��������' ;
ArrayData[175,2] :=  ' ������� ����������� ���' ;
ArrayData[176,2] :=  '������ ���' ;
ArrayData[177,2] :=  '������ ��������' ;
ArrayData[178,2] :=  ' ������ ��' ;
ArrayData[180,2] :=  '������ �������� � �� � 1�� �������' ;
ArrayData[182,2] :=  '��� �������� ������' ;
ArrayData[183,2] :=  '����������� ��������������� ������������� ��' ;
ArrayData[184,2] :=  '������ �������� � �� �� �� �������' ;
ArrayData[186,2] :=  '�������� ����� ���� � ���� �� ����� ���������' ;
ArrayData[189,2] :=  '�������� � ���' ;
ArrayData[190,2] :=  '������������ �������� ���' ;
ArrayData[191,2] :=  '�������������� �������� ���' ;
ArrayData[192,2] :=  '������������� ������� ��������� ������ �����' ;
ArrayData[193,2] :=  '����� ���������� ��� ���' ;
ArrayData[194,2] :=  '����������� ��������������� ������������� ���' ;
ArrayData[196,2] :=  '������ �������� � �� �� ��� �������' ;
ArrayData[198,2] :=  '��� �������� ������' ;
ArrayData[199,2] :=  '����������� ��������������� ������������� ��' ;
ArrayData[200,2] :=  '������ �������� � �� �� �� �������' ;
ArrayData[201,2] :=  ' ��������� ������ ��������' ;
ArrayData[202,2] :=  '��������' ;
form4.ProgressBar1.position:=form4.ProgressBar1.position+5 ;    //((((((())))))

//������ �������

ArrayData[23,3] := 'f����' ;
ArrayData[24,3] := 'f������' ;
ArrayData[25,3] := 'd��' ;
ArrayData[26,3] := 'f��' ;
ArrayData[27,3] := 'f���� 2' ;
ArrayData[28,3] := 'tcp��' ;
ArrayData[29,3] :=  Chr(110)+ '��' ;
ArrayData[31,3] := 'i����' ;
ArrayData[32,3] := 'i�����' ;
ArrayData[37,3] :=  Chr(68)+ '��� 1' ;
ArrayData[38,3] := '����' ;
ArrayData[39,3] := '���' ;
ArrayData[40,3] := '���' ;
ArrayData[42,3] := '���' ;
ArrayData[43,3] := '����' ;
ArrayData[44,3] := '���' ;
ArrayData[46,3] := 'ts ���' ;
ArrayData[47,3] := 'ts ��' ;
ArrayData[49,3] := 'i''''' ;
ArrayData[50,3] := 'is''' ;
ArrayData[51,3] := 'i��' ;
ArrayData[52,3] := 'i��' ;
ArrayData[54,3] := 't��' ;
ArrayData[55,3] := 't��' ;
ArrayData[56,3] := 't���' ;
ArrayData[58,3] := 'Q��' ;
ArrayData[59,3] := 'Q���' ;
ArrayData[60,3] := 'Q��' ;
ArrayData[61,3] := 'Q��' ;
ArrayData[63,3] := 'G��' ;
ArrayData[65,3] := 'i�����' ;
ArrayData[66,3] := 'i����' ;
ArrayData[68,3] := 't�����' ;
ArrayData[69,3] := 't����' ;
ArrayData[71,3] := 't����' ;
ArrayData[72,3] := 't�����' ;
ArrayData[74,3] := 't����' ;
ArrayData[75,3] := 't�����' ;
ArrayData[76,3] := 't����' ;
ArrayData[77,3] := 'W1' ;
ArrayData[78,3] := 'd�� 1' ;
ArrayData[81,3] :=  Chr(108)+ '��' ;
ArrayData[82,3] :=  Chr(108)+ '����' ;
ArrayData[84,3] :=  Chr(109)+ '��' ;
ArrayData[85,3] :=  Chr(109)+ '����' ;
ArrayData[86,3] :=  Chr(74)+ '��' ;
ArrayData[87,3] := 'Re1' ;
ArrayData[89,3] := 'Pr��' ;
ArrayData[90,3] := 'Pr����' ;
ArrayData[91,3] := 'Nu1' ;
ArrayData[94,3] :=  Chr(110)+ '��' ;
ArrayData[95,3] := 'f����' ;
ArrayData[96,3] := 'f��' ;
ArrayData[97,3] := 'W��' ;
ArrayData[99,3] :=  Chr(109)+ '��' ;
ArrayData[100,3] :=  Chr(109)+ '����' ;
ArrayData[101,3] :=  Chr(74)+ '��' ;
ArrayData[102,3] := 'd� ��' ;
ArrayData[103,3] := 'Re ��' ;
ArrayData[105,3] :=  Chr(108)+ '��' ;
ArrayData[106,3] :=  Chr(108)+ '����' ;
ArrayData[108,3] := 'Pr��' ;
ArrayData[109,3] := 'Pr����' ;
ArrayData[111,3] := 'D��' ;
ArrayData[112,3] :=  Chr(89)+ '��' ;
ArrayData[113,3] := 'B��' ;
ArrayData[114,3] := 'm��' ;
ArrayData[115,3] := 'Nu��' ;
ArrayData[117,3] := 'a��' ;
ArrayData[119,3] := 't����' ;
ArrayData[121,3] :=  Chr(108)+ '����' ;
ArrayData[122,3] :=  Chr(108)+ '����' ;
ArrayData[123,3] := '���' ;
ArrayData[124,3] :=  Chr(68)+ 't��' ;
ArrayData[125,3] := 'F��' ;
ArrayData[126,3] := 'L��' ;
ArrayData[128,3] :=  Chr(110)+ '��' ;
ArrayData[130,3] :=  Chr(109)+ '��' ;
ArrayData[131,3] :=  Chr(109)+ '����' ;
ArrayData[132,3] :=  Chr(74)+ '��' ;
ArrayData[133,3] := 'f����' ;
ArrayData[134,3] := 'f��' ;
ArrayData[135,3] := 'W��' ;
ArrayData[136,3] := 'd� ��' ;
ArrayData[137,3] := 'Re��' ;
ArrayData[140,3] :=  Chr(108)+ '��' ;
ArrayData[142,3] :=  Chr(108)+ '����' ;
ArrayData[144,3] := 'Pr��' ;
ArrayData[145,3] := 'Pr����' ;
ArrayData[147,3] := 'D��' ;
ArrayData[148,3] :=  Chr(89)+ '��' ;
ArrayData[149,3] := 'B��' ;
ArrayData[150,3] := 'm��' ;
ArrayData[151,3] := 'Nu��' ;
ArrayData[153,3] := 'a��' ;
ArrayData[155,3] :=  Chr(108)+ '����' ;
ArrayData[156,3] := 't����' ;
ArrayData[157,3] :=  Chr(108)+ '����' ;
ArrayData[158,3] :=  Chr(108)+ '����' ;
ArrayData[159,3] := 'K��' ;
ArrayData[160,3] :=  Chr(68)+ 't��' ;
ArrayData[161,3] := 'F��' ;
ArrayData[162,3] := 'L��' ;
ArrayData[164,3] := 'q1' ;
ArrayData[165,3] := 'a���' ;
ArrayData[167,3] := 't�����' ;
ArrayData[169,3] :=  Chr(108)+ '�����' ;
ArrayData[170,3] :=  Chr(108)+ '�����' ;
ArrayData[171,3] := 'K���' ;
ArrayData[172,3] :=  Chr(68)+ 't���' ;
ArrayData[173,3] := 'q2' ;
ArrayData[175,3] := 'F���' ;
ArrayData[176,3] := 'L���' ;
ArrayData[178,3] := 'L��' ;
ArrayData[179,3] :=  Chr(122)+ '1' ;
ArrayData[180,3] :=  Chr(68)+ 'P1' ;
ArrayData[182,3] := '���' ;
ArrayData[183,3] :=  Chr(122)+ '��' ;
ArrayData[184,3] :=  Chr(68)+ 'P��' ;
ArrayData[187,3] :=  Chr(110)+ '''''' ;
ArrayData[188,3] :=  Chr(110)+ '''' ;
ArrayData[189,3] := 'W0' ;
ArrayData[190,3] :=  Chr(109)+ '���' ;
ArrayData[191,3] :=  Chr(74)+ '���' ;
ArrayData[192,3] := 'd� ���' ;
ArrayData[193,3] := 'Re���' ;
ArrayData[194,3] :=  Chr(122)+ '���' ;
ArrayData[196,3] :=  Chr(68)+ 'P���' ;
ArrayData[198,3] := '���' ;
ArrayData[199,3] :=  Chr(122)+ '��' ;
ArrayData[200,3] :=  Chr(68)+ 'P��' ;
ArrayData[201,3] :=  Chr(68)+ 'P�� 2' ;
form4.ProgressBar1.position:=form4.ProgressBar1.position+5 ;    //((((((())))))

//��������� �������
ArrayData[23,4] :=  '�2' ;
ArrayData[24,4] :=  '�2' ;
ArrayData[25,4] :=  '��' ;
ArrayData[26,4] :=  '�2' ;
ArrayData[27,4] :=  '�2' ;
ArrayData[28,4] :=  '��' ;
ArrayData[29,4] :=  '�3/��' ;
ArrayData[37,4] :=  '���' ;
ArrayData[38,4] :=  '���' ;
ArrayData[39,4] :=  '���' ;
ArrayData[40,4] :=  '���' ;
ArrayData[42,4] :=  '���' ;
ArrayData[43,4] :=  '���' ;
ArrayData[44,4] :=  '���' ;
ArrayData[46,4] :=  '��' ;
ArrayData[47,4] :=  '��' ;
ArrayData[54,4] :=  '��' ;
ArrayData[55,4] :=  '��' ;
ArrayData[56,4] :=  '��' ;
ArrayData[58,4] :=  '���' ;
ArrayData[59,4] :=  '���' ;
ArrayData[60,4] :=  '���' ;
ArrayData[61,4] :=  '���' ;
ArrayData[63,4] :=  '�3/��' ;
ArrayData[68,4] :=  '��' ;
ArrayData[69,4] :=  '��' ;
ArrayData[71,4] :=  '��' ;
ArrayData[72,4] :=  '��' ;
ArrayData[74,4] :=  '��' ;
ArrayData[75,4] :=  '��' ;
ArrayData[76,4] :=  '��' ;
ArrayData[77,4] :=  '�/�' ;
ArrayData[78,4] :=  '�' ;
ArrayData[81,4] :=  '���/� �' ;
ArrayData[82,4] :=  '���/� �' ;
ArrayData[84,4] :=  '����*�' ;
ArrayData[85,4] :=  '����*�' ;
ArrayData[86,4] :=  '�2/�' ;
ArrayData[94,4] :=  '�3/��' ;
ArrayData[95,4] :=  '�2' ;
ArrayData[96,4] :=  '�2' ;
ArrayData[97,4] :=  '�/�' ;
ArrayData[99,4] :=  '����*�' ;
ArrayData[100,4] :=  '����*�' ;
ArrayData[101,4] :=  '�2/�' ;
ArrayData[102,4] :=  '�' ;
ArrayData[105,4] :=  '���/� �' ;
ArrayData[106,4] :=  '���/� �' ;
ArrayData[111,4] :=  '�' ;
ArrayData[117,4] :=  '���/�2*��' ;
ArrayData[119,4] :=  '��' ;
ArrayData[121,4] :=  '���/� �' ;
ArrayData[122,4] :=  '���/� �' ;
ArrayData[123,4] :=  '���/�2*�' ;
ArrayData[124,4] :=  '��' ;
ArrayData[125,4] :=  '�2' ;
ArrayData[126,4] :=  '�' ;
ArrayData[128,4] :=  '�3/��' ;
ArrayData[130,4] :=  '����*�' ;
ArrayData[131,4] :=  '����*�' ;
ArrayData[132,4] :=  '�2/�' ;
ArrayData[133,4] :=  '�2' ;
ArrayData[134,4] :=  '�2' ;
ArrayData[135,4] :=  '�/�' ;
ArrayData[140,4] :=  '���/� �' ;
ArrayData[142,4] :=  '���/� �' ;
ArrayData[147,4] :=  '�' ;
ArrayData[153,4] :=  '���/�2*��' ;
ArrayData[155,4] :=  '���/� �' ;
ArrayData[156,4] :=  '��' ;
ArrayData[157,4] :=  '���/� �' ;
ArrayData[158,4] :=  '���/� �' ;
ArrayData[159,4] :=  '���/�2*�' ;
ArrayData[160,4] :=  '��' ;
ArrayData[161,4] :=  '�2' ;
ArrayData[162,4] :=  '�' ;
ArrayData[164,4] :=  '��/�2' ;
ArrayData[165,4] :=  '���/�2*��' ;
ArrayData[167,4] :=  '��' ;
ArrayData[169,4] :=  '���/� �' ;
ArrayData[170,4] :=  '���/� �' ;
ArrayData[171,4] :=  '���/�2*�' ;
ArrayData[172,4] :=  '��' ;
ArrayData[173,4] :=  '��/�2' ;
ArrayData[175,4] :=  '�2' ;
ArrayData[176,4] :=  '�' ;
ArrayData[178,4] :=  '�' ;
ArrayData[180,4] :=  '���' ;
ArrayData[184,4] :=  '���' ;
ArrayData[187,4] :=  '�3 /��' ;
ArrayData[188,4] :=  '�3 /��' ;
ArrayData[189,4] :=  '�/�' ;
ArrayData[191,4] :=  '�2/�' ;
ArrayData[192,4] :=  '�' ;
ArrayData[196,4] :=  '���' ;
ArrayData[200,4] :=  '���' ;
ArrayData[201,4] :=  '���' ;
ArrayData[202,4] :=  '%' ;
form4.ProgressBar1.position:=form4.ProgressBar1.position+5 ;    //((((((())))))

//����� �������
ArrayData[23,5] :=  ftp_th  ;
ArrayData[24,5] :=  fnpox_th ;
ArrayData[25,5] :=  d_BH ;
ArrayData[26,5] :=  f_BH ;
ArrayData[27,5] :=  fnpox_2 ;
ArrayData[28,5] :=  t_cp_TH ;
ArrayData[29,5] :=  V_TH ;
ArrayData[31,5] :=  i_Bx_TH ;
ArrayData[32,5] :=  i_Bix_TH ;
ArrayData[34,5] :=  Delta_P_nn_1 ;
ArrayData[35,5] :=  Delta_P_icn_1 ;
ArrayData[36,5] :=  Delta_P_ak_1 ;
ArrayData[38,5] :=  P_icn ;
ArrayData[39,5] :=  P_ak ;
ArrayData[40,5] :=  P_BX ;
ArrayData[42,5] :=  P_ak_cp ;
ArrayData[43,5] :=  P_icn_cp ;
ArrayData[44,5] :=  P_nn_cp ;
ArrayData[46,5] :=  t_s_icn ;
ArrayData[47,5] :=  t_s_ak ;
ArrayData[49,5] :=  iYY ;
ArrayData[50,5] :=  iY_s ;
ArrayData[51,5] :=  i_nB ;
ArrayData[52,5] :=  ine ;
ArrayData[54,5] :=  t_cp_ak ;
ArrayData[55,5] :=  t_cp_nn ;
ArrayData[56,5] :=  t_cp_icn ;
ArrayData[58,5] :=  Q_ak ;
ArrayData[59,5] :=  Q_icn ;
ArrayData[60,5] :=  Q_nn ;
ArrayData[61,5] :=  Q_nr ;
ArrayData[63,5] :=  G_TH ;
ArrayData[65,5] :=  i_icn_TH ;
ArrayData[66,5] :=  i_ak_TH ;
ArrayData[68,5] :=  t_icn_TH ;
ArrayData[69,5] :=  t_ak_TH ;
ArrayData[71,5] :=  t_cp_TH_ne ;
ArrayData[72,5] :=  t_cp_TH_icn ;
ArrayData[74,5] :=  t_cp_ct_nn ;
ArrayData[75,5] :=  t_cp_ct_icn ;
ArrayData[76,5] :=  t_cp_ct_ak ;
ArrayData[77,5] :=  W_1 ;
ArrayData[78,5] :=  d_ek_1 ;
ArrayData[81,5] :=  Lambda_TH ;
ArrayData[82,5] :=  Lambda_TH_ct ;
ArrayData[84,5] :=  Mu_TH ;
ArrayData[85,5] :=  Mu_TH_ct ;
ArrayData[87,5] :=  Re_1 ;
ArrayData[89,5] :=  Pr_TH ;
ArrayData[90,5] :=  Pr_TH_ct ;
ArrayData[91,5] :=  Nu_1 ;
ArrayData[94,5] :=  V_ak_cp ;
ArrayData[95,5] :=  f_ak_TP ;
ArrayData[96,5] :=  f_ak ;
ArrayData[97,5] :=  W_ak ;
ArrayData[99,5] :=  Mu_cp_ak ;
ArrayData[100,5] :=  Mu_cp_ct_ak ;
ArrayData[101,5] :=  KINV_cp_ak ;
ArrayData[102,5] :=  d_a_ak ;
ArrayData[103,5] :=  Re_ak ;
ArrayData[105,5] :=  Lambda_ak ;
ArrayData[106,5] :=  Lambda_ct_ak ;
ArrayData[108,5] :=  Pr_ak ;
ArrayData[109,5] :=  Pr_ct_ak ;
ArrayData[111,5] :=  D_ak ;
ArrayData[112,5] :=    Psi_ak ;
ArrayData[113,5] :=    B_ak ;
ArrayData[114,5] :=    m_ak ;
ArrayData[115,5] :=  Nu_ak ;
ArrayData[117,5] :=  a_ak  ;
ArrayData[119,5] :=  t_cp_ct_ak ;
ArrayData[121,5] :=  Lambda_MAT_ak ;
ArrayData[122,5] :=  Lambda_MAT_ak ;
ArrayData[123,5] :=  K_ak ;
ArrayData[124,5] :=  Delta_t_ak ;
ArrayData[125,5] :=  F__ak ;
ArrayData[126,5] :=  L_ak ;
ArrayData[128,5] :=  V_nn_cp ;
ArrayData[130,5] :=  Mu_cp_nn ;
ArrayData[131,5] :=  Mu_cp_ct_nn ;
ArrayData[132,5] :=  KINV_cp_nn ;
ArrayData[133,5] :=  f_nn_TP ;
ArrayData[134,5] :=  f_nn ;
ArrayData[135,5] :=  W_nn ;
ArrayData[136,5] :=  d_a_nn ;
ArrayData[137,5] :=  Re_nn ;
ArrayData[140,5] :=  Lambda_nn ;
ArrayData[142,5] :=  Lambda_ct_nn ;
ArrayData[144,5] :=  Pr_nn ;
ArrayData[145,5] :=  Pr_ct_nn ;
ArrayData[147,5] :=  D_nn ;
ArrayData[148,5] :=  Psi_nn ;
ArrayData[149,5] :=    B_nn ;
ArrayData[150,5] :=    m_nn ;
ArrayData[151,5] :=  Nu_nn ;
ArrayData[153,5] :=  a_nn  ;
ArrayData[155,5] :=  Lambda_MAT_nn ;
ArrayData[156,5] :=  t_nn_TH ;
ArrayData[157,5] :=  Lambda_MAT_nn ;
ArrayData[158,5] :=  Lambda_MAT_nn ;
ArrayData[159,5] :=  K_nn ;
ArrayData[160,5] :=  Delta_t_nn ;
ArrayData[161,5] :=  F__nn ;
ArrayData[162,5] :=  L_nn ;
ArrayData[164,5] :=  q_1 ;
ArrayData[165,5] :=  a_icn ;
ArrayData[167,5] :=  t_cp_ct_icn ;
ArrayData[169,5] :=  Lambda_MAT_icn ;
ArrayData[170,5] :=  Lambda_MAT_icn ;
ArrayData[171,5] :=  K_icn ;
ArrayData[172,5] :=  Delta_t_icn ;
ArrayData[173,5] :=  q_2 ;
ArrayData[174,5] :=  Prov_q ;
ArrayData[175,5] :=  F__icn ;
ArrayData[176,5] :=  L_nn ;
ArrayData[178,5] :=  L_nr ;
ArrayData[179,5] :=  Zeta_1 ;
ArrayData[180,5] :=  Delta_P_1_2 ;
ArrayData[182,5] :=  c_ak ;
ArrayData[183,5] :=  Zeta_ak ;
ArrayData[184,5] :=  Delta_P_ak_2 ;
ArrayData[187,5] :=  V_cp_YY ;
ArrayData[188,5] :=  V_cp_Y ;
ArrayData[189,5] :=  W_0 ;
ArrayData[190,5] :=  Mu_cp_icn ;
ArrayData[191,5] :=  KINV_cp_icn ;
ArrayData[192,5] :=  d_ek_icn ;
ArrayData[193,5] :=  Re_icn ;
ArrayData[194,5] :=  Zeta_icn ;
ArrayData[196,5] :=  Delta_P_icn_2 ;
ArrayData[198,5] :=  c_nn ;
ArrayData[199,5] :=  Zeta_nn ;
ArrayData[200,5] :=  Delta_P_nn_2 ;
ArrayData[201,5] :=  Delta_P_nr_2 ;
ArrayData[202,5] :=  Prov_P_nr ;
form4.ProgressBar1.position:=form4.ProgressBar1.position+5 ;    //((((((())))))


// ����� ������� ������ �������, � ������� ����� �������� ������
  Cell1 := WorkBook.WorkSheets[1].Cells[BeginRow, BeginCol];

// ������ ������ ������ �������, � ������� ����� �������� ������
  Cell2 := WorkBook.WorkSheets[1].Cells[BeginRow  + RowCount - 1, BeginCol +
ColCount - 1];

// �������, � ������� ����� �������� ������
  Range := WorkBook.WorkSheets[1].Range[Cell1, Cell2];

// � ��� � ��� ����� ������
  // ������� ������� ����������� ����������
  Range.Value := ArrayData;

form4.ProgressBar1.position:=form4.ProgressBar1.position+5 ;    //((((((())))))

//������ ������
  //����
ExcelPamka('A3:E3',2,2,-4138,-4138) ;
ExcelPamka('A4:E4',2,2,-4138,-4138) ;
ExcelPamka('A5:E5',2,2,-4138,-4138) ;
ExcelPamka('A6:E6',2,2,-4138,-4138) ;
ExcelPamka('A7:E7',2,2,-4138,-4138) ;
ExcelPamka('A8:E8',2,2,-4138,-4138) ;
ExcelPamka('A9:E9',2,2,-4138,-4138) ;
ExcelPamka('A10:E10',2,2,-4138,-4138) ;
ExcelPamka('A11:E11',2,2,-4138,-4138) ;
ExcelPamka('A12:E12',2,2,-4138,-4138) ;
ExcelPamka('A13:E13',2,2,-4138,-4138) ;
ExcelPamka('A14:E14',2,2,-4138,-4138) ;
ExcelPamka('A15:E15',2,2,-4138,-4138) ;
ExcelPamka('A16:E16',2,2,-4138,-4138) ;
ExcelPamka('A17:E17',2,2,-4138,-4138) ;
ExcelPamka('A18:E18',2,2,-4138,-4138) ;
ExcelPamka('A19:E19',2,2,-4138,-4138) ;
ExcelPamka('A20:E20',2,2,-4138,-4138) ;
ExcelPamka('A21:E21',2,2,-4138,-4138) ;
ExcelPamka('A22:E22',2,2,-4138,-4138) ;
  //������
ExcelPamka('A25:E25',2,2,-4138,-4138) ;
ExcelPamka('A27:E27',2,2,-4138,-4138) ;
ExcelPamka('A29:E29',2,2,-4138,-4138) ;
ExcelPamka('A31:E31',2,2,-4138,-4138) ;
ExcelPamka('A33:E33',2,2,-4138,-4138) ;
ExcelPamka('A35:E35',2,2,-4138,-4138) ;
ExcelPamka('A37:E37',2,2,-4138,-4138) ;
ExcelPamka('A39:E39',2,2,-4138,-4138) ;
ExcelPamka('A41:E41',2,2,-4138,-4138) ;
ExcelPamka('A43:E43',2,2,-4138,-4138) ;
ExcelPamka('A45:E45',2,2,-4138,-4138) ;
ExcelPamka('A47:E47',2,2,-4138,-4138) ;
ExcelPamka('A49:E49',2,2,-4138,-4138) ;
ExcelPamka('A51:E51',2,2,-4138,-4138) ;
ExcelPamka('A53:E53',2,2,-4138,-4138) ;
ExcelPamka('A55:E55',2,2,-4138,-4138) ;
ExcelPamka('A57:E57',2,2,-4138,-4138) ;
ExcelPamka('A59:E59',2,2,-4138,-4138) ;
ExcelPamka('A61:E61',2,2,-4138,-4138) ;
ExcelPamka('A63:E63',2,2,-4138,-4138) ;
ExcelPamka('A65:E65',2,2,-4138,-4138) ;
ExcelPamka('A67:E67',2,2,-4138,-4138) ;
ExcelPamka('A69:E69',2,2,-4138,-4138) ;
ExcelPamka('A71:E71',2,2,-4138,-4138) ;
ExcelPamka('A73:E73',2,2,-4138,-4138) ;
ExcelPamka('A75:E75',2,2,-4138,-4138) ;
ExcelPamka('A77:E77',2,2,-4138,-4138) ;
ExcelPamka('A79:E79',2,2,-4138,-4138) ;
ExcelPamka('A81:E81',2,2,-4138,-4138) ;
ExcelPamka('A83:E83',2,2,-4138,-4138) ;
ExcelPamka('A85:E85',2,2,-4138,-4138) ;
ExcelPamka('A87:E87',2,2,-4138,-4138) ;
ExcelPamka('A89:E89',2,2,-4138,-4138) ;
ExcelPamka('A91:E91',2,2,-4138,-4138) ;
ExcelPamka('A93:E93',2,2,-4138,-4138) ;
ExcelPamka('A95:E95',2,2,-4138,-4138) ;
ExcelPamka('A97:E97',2,2,-4138,-4138) ;
ExcelPamka('A99:E99',2,2,-4138,-4138) ;
form4.ProgressBar1.position:=form4.ProgressBar1.position+5 ;    //((((((())))))
ExcelPamka('A101:E101',2,2,-4138,-4138) ;
ExcelPamka('A103:E103',2,2,-4138,-4138) ;
ExcelPamka('A105:E105',2,2,-4138,-4138) ;
ExcelPamka('A107:E107',2,2,-4138,-4138) ;
ExcelPamka('A109:E109',2,2,-4138,-4138) ;
ExcelPamka('A111:E111',2,2,-4138,-4138) ;
ExcelPamka('A113:E113',2,2,-4138,-4138) ;
ExcelPamka('A115:E115',2,2,-4138,-4138) ;
ExcelPamka('A117:E117',2,2,-4138,-4138) ;
ExcelPamka('A119:E119',2,2,-4138,-4138) ;
ExcelPamka('A121:E121',2,2,-4138,-4138) ;
ExcelPamka('A123:E123',2,2,-4138,-4138) ;
ExcelPamka('A125:E125',2,2,-4138,-4138) ;
ExcelPamka('A127:E127',2,2,-4138,-4138) ;
ExcelPamka('A129:E129',2,2,-4138,-4138) ;
ExcelPamka('A131:E131',2,2,-4138,-4138) ;
ExcelPamka('A133:E133',2,2,-4138,-4138) ;
ExcelPamka('A135:E135',2,2,-4138,-4138) ;
ExcelPamka('A137:E137',2,2,-4138,-4138) ;
ExcelPamka('A139:E139',2,2,-4138,-4138) ;
ExcelPamka('A141:E141',2,2,-4138,-4138) ;
ExcelPamka('A143:E143',2,2,-4138,-4138) ;
ExcelPamka('A145:E145',2,2,-4138,-4138) ;
ExcelPamka('A147:E147',2,2,-4138,-4138) ;
ExcelPamka('A149:E149',2,2,-4138,-4138) ;
ExcelPamka('A151:E151',2,2,-4138,-4138) ;
ExcelPamka('A153:E153',2,2,-4138,-4138) ;
ExcelPamka('A155:E155',2,2,-4138,-4138) ;
ExcelPamka('A157:E157',2,2,-4138,-4138) ;
ExcelPamka('A159:E159',2,2,-4138,-4138) ;
ExcelPamka('A161:E161',2,2,-4138,-4138) ;
ExcelPamka('A163:E163',2,2,-4138,-4138) ;
ExcelPamka('A165:E165',2,2,-4138,-4138) ;
ExcelPamka('A167:E167',2,2,-4138,-4138) ;
ExcelPamka('A169:E169',2,2,-4138,-4138) ;
ExcelPamka('A171:E171',2,2,-4138,-4138) ;
ExcelPamka('A173:E173',2,2,-4138,-4138) ;
ExcelPamka('A175:E175',2,2,-4138,-4138) ;
ExcelPamka('A177:E177',2,2,-4138,-4138) ;
ExcelPamka('A179:E179',2,2,-4138,-4138) ;
ExcelPamka('A181:E181',2,2,-4138,-4138) ;
ExcelPamka('A183:E183',2,2,-4138,-4138) ;
ExcelPamka('A185:E185',2,2,-4138,-4138) ;
ExcelPamka('A187:E187',2,2,-4138,-4138) ;
ExcelPamka('A189:E189',2,2,-4138,-4138) ;
ExcelPamka('A191:E191',2,2,-4138,-4138) ;
ExcelPamka('A193:E193',2,2,-4138,-4138) ;
ExcelPamka('A195:E195',2,2,-4138,-4138) ;
ExcelPamka('A197:E197',2,2,-4138,-4138) ;
ExcelPamka('A199:E199',2,2,-4138,-4138) ;
ExcelPamka('A201:E201',2,2,-4138,-4138) ;
ExcelPamka('A203:E203',2,2,-4138,-4138) ;
ExcelPamka('A205:E205',2,2,-4138,-4138) ;


//������� �������
ExcelPamka('A3:A22',-4138,-4138,-4138,-4138) ;
ExcelPamka('B3:B22',-4138,-4138,-4138,-4138) ;
ExcelPamka('C3:C22',-4138,-4138,-4138,-4138) ;
ExcelPamka('D3:D22',-4138,-4138,-4138,-4138) ;
ExcelPamka('E3:E22',-4138,-4138,-4138,-4138) ;
ExcelPamka('A25:A205',-4138,-4138,-4138,-4138) ;
ExcelPamka('B25:B205',-4138,-4138,-4138,-4138) ;
ExcelPamka('C25:C205',-4138,-4138,-4138,-4138) ;
ExcelPamka('D25:D205',-4138,-4138,-4138,-4138) ;
ExcelPamka('E25:E205',-4138,-4138,-4138,-4138) ;


  // ������� �� �����
ExcelApp.Range['B4:B23'].Select;
ExcelApp.Selection.WrapText:=True;
ExcelApp.Range['B25:B205'].Select;
ExcelApp.Selection.WrapText:=True;

// ������������ � �����
ExcelApp.Range['A4:A205'].Select;
ExcelApp.Selection.HorizontalAlignment:=2;
ExcelApp.Selection.VerticalAlignment:=2 ;
ExcelApp.Selection.Font.Size := 11;
ExcelApp.Selection.Font.Name := 'Calibri';
ExcelApp.Range['B4:B205'].Select;
ExcelApp.Selection.HorizontalAlignment:=2;
ExcelApp.Selection.VerticalAlignment:=2 ;
ExcelApp.Selection.Font.Size := 11;
ExcelApp.Selection.Font.Name := 'Cambria';
ExcelApp.Range['C4:C205'].Select;
ExcelApp.Selection.HorizontalAlignment:=3;
ExcelApp.Selection.VerticalAlignment:=2 ;
ExcelApp.Selection.Font.Size := 13;
ExcelApp.Selection.Font.Name := 'Calibri';
ExcelApp.Range['D4:D205'].Select;
ExcelApp.Selection.HorizontalAlignment:=3;
ExcelApp.Selection.VerticalAlignment:=2 ;
ExcelApp.Selection.Font.Size := 11;
ExcelApp.Selection.Font.Name := 'Cambria';
ExcelApp.Range['E4:E205'].Select;
ExcelApp.Selection.HorizontalAlignment:=2;
ExcelApp.Selection.VerticalAlignment:=2 ;
ExcelApp.Selection.Font.Size := 11;
ExcelApp.Selection.Font.Name := 'Times New Roman';

form4.ProgressBar1.position:=form4.ProgressBar1.position+5 ;    //((((((())))))

// ����� ��������
 //������ �������
  // ���������
  ExcelApp.Range['A3:E3'].Select;
  ExcelApp.Selection.WrapText:=True;
  ExcelApp.Selection.HorizontalAlignment:=3;
  ExcelApp.Selection.VerticalAlignment:=2 ;
  ExcelApp.Selection.Font.Size := 12;
  ExcelApp.Selection.Font.Name := 'Times New Roman';
  ExcelApp.Range['A3']:='�';
  ExcelApp.Range['B3']:='��������';
  ExcelApp.Range['C3']:='�����������';
  ExcelApp.Range['D3']:='������� ���������';
  ExcelApp.Range['E3']:='��������';

form4.ProgressBar1.position:=form4.ProgressBar1.position+5 ;    //((((((())))))

ExcelApp.Range['A4']:='1' ;
ExcelApp.Range['A5']:='2' ;
ExcelApp.Range['A6']:='3' ;
ExcelApp.Range['A7']:='4' ;
ExcelApp.Range['A8']:='5' ;
ExcelApp.Range['A9']:='6' ;
ExcelApp.Range['A10']:='7' ;
ExcelApp.Range['A11']:='8' ;
ExcelApp.Range['A12']:='9' ;
ExcelApp.Range['A13']:='10' ;
ExcelApp.Range['A14']:='11' ;
ExcelApp.Range['A15']:='12' ;
ExcelApp.Range['A16']:='13' ;
ExcelApp.Range['A17']:='14' ;
ExcelApp.Range['A18']:='15' ;
ExcelApp.Range['A19']:='16' ;
ExcelApp.Range['A20']:='17' ;
ExcelApp.Range['A21']:='18' ;
ExcelApp.Range['A22']:='19' ;

form4.ProgressBar1.position:=form4.ProgressBar1.position+5 ;    //((((((())))))

ExcelApp.Range['B4']:='���������������������� ����� �������'   ;
ExcelApp.Range['B5']:='�������� ����������� ����'           ;
ExcelApp.Range['B6']:='����������� ����������� ����'       ;
ExcelApp.Range['B7']:='����������� ����������� ����'        ;
ExcelApp.Range['B8']:='����������� �� �� ����� � ��'        ;
ExcelApp.Range['B9']:='�������� ������� �������'            ;
ExcelApp.Range['B10']:='����������� �� �� ������ �� ��'     ;
ExcelApp.Range['B11']:='�������� �����'                    ;
ExcelApp.Range['B12']:='�������� ������� ������� ������'          ;
ExcelApp.Range['B13']:='������� ������ ������'             ;
ExcelApp.Range['B14']:='��� ������� �������'             ;
ExcelApp.Range['B15']:='����� ������������ ����'          ;
ExcelApp.Range['B16']:='������ �� �� ��������� �����'      ;
ExcelApp.Range['B17']:='��� ����� �� �� �������'              ;
ExcelApp.Range['B18']:='�������  ����� �� �� �������'          ;
ExcelApp.Range['B19']:='������� �������  ����� �� �� �������'   ;
ExcelApp.Range['B20']:='��� ����� �� �� �������'                 ;
ExcelApp.Range['B21']:='�������  ����� ��  ��  �������'           ;
ExcelApp.Range['B22']:='������� �������  ����� ��  ��  �������'    ;



ExcelApp.Range['C4']:='D���' ;
ExcelApp.Range['C5']:='���' ;
ExcelApp.Range['C6']:='t��' ;
ExcelApp.Range['C7']:='t��' ;
ExcelApp.Range['C8']:='t��' ;
ExcelApp.Range['C9']:='P1' ;
ExcelApp.Range['C10']:='t���' ;
ExcelApp.Range['C11']:='' ;
ExcelApp.Range['C12']:='d�' ;
ExcelApp.Range['C13']:=Chr(100) + '��' ;
Start := 1;
Length := 1;
ExcelApp.Range['C13'].Characters[Start, Length].Font.Name := 'Symbol';
ExcelApp.Range['C14']:='S��' ;
ExcelApp.Range['C15']:='n��' ;
ExcelApp.Range['C16']:='' ;
ExcelApp.Range['C17']:='S����' ;
ExcelApp.Range['C18']:='d����' ;
ExcelApp.Range['C19']:=Chr(100) + '����' ;
Start := 1;
Length := 1;
ExcelApp.Range['C19'].Characters[Start, Length].Font.Name := 'Symbol';
ExcelApp.Range['C20']:='S����' ;
ExcelApp.Range['C21']:='d����' ;
ExcelApp.Range['C22']:=Chr(100) + '����' ;
Start := 1;
Length := 1;
ExcelApp.Range['C19'].Characters[Start, Length].Font.Name := 'Symbol';

form4.ProgressBar1.position:=form4.ProgressBar1.position+5 ;    //((((((())))))

ExcelApp.Range['D4']:='��/�' ;
ExcelApp.Range['D5']:='���' ;
ExcelApp.Range['D6']:=Chr(176) + '�' ;
Start := 1;
Length := 1;
ExcelApp.Range['D6'].Characters[Start, Length].Font.Name := 'Symbol';
ExcelApp.Range['D7']:=Chr(176) + '�' ;
Start := 1;
Length := 1;
ExcelApp.Range['D7'].Characters[Start, Length].Font.Name := 'Symbol';
ExcelApp.Range['D8']:=Chr(176) + '�' ;
Start := 1;
Length := 1;
ExcelApp.Range['D8'].Characters[Start, Length].Font.Name := 'Symbol';
ExcelApp.Range['D9']:='���' ;
ExcelApp.Range['D10']:=Chr(176) + '�' ;
Start := 1;
Length := 1;
ExcelApp.Range['D10'].Characters[Start, Length].Font.Name := 'Symbol';
ExcelApp.Range['D11']:=''   ;
ExcelApp.Range['D12']:='��' ;
ExcelApp.Range['D13']:='��' ;
ExcelApp.Range['D14']:='��' ;
ExcelApp.Range['D15']:=''   ;
ExcelApp.Range['D16']:=''   ;
ExcelApp.Range['D17']:='��' ;
ExcelApp.Range['D18']:='��' ;
ExcelApp.Range['D19']:='��' ;
ExcelApp.Range['D20']:='��' ;
ExcelApp.Range['D21']:='��' ;
ExcelApp.Range['D22']:='��' ;




form4.ProgressBar1.position:=form4.ProgressBar1.position+5 ;    //((((((())))))

// ������ � ������ �������� � �����
    //������
  ExcelApp.Range['A:A', EmptyParam].EntireColumn.ColumnWidth :=4.71 ;
  ExcelApp.Range['B:B', EmptyParam].EntireColumn.ColumnWidth :=37.57 ;
  ExcelApp.Range['C:C', EmptyParam].EntireColumn.ColumnWidth :=13.14 ;
  ExcelApp.Range['D:D', EmptyParam].EntireColumn.ColumnWidth :=12.43 ;
  ExcelApp.Range['E:E', EmptyParam].EntireColumn.ColumnWidth :=15.29 ;
  //������   (������ ����!!!!)
   ExcelApp.Range['A1:A22', EmptyParam].EntireRow.autofit ;
   ExcelApp.Range['A25:A205', EmptyParam].EntireRow.autofit ;
     /// ����������
     //  ExcelApp.Range['30:30', EmptyParam].EntireRow.RowHeight :=33.75 ;

      
//��������� ��������

ExcelApp.Range['C32'].Characters[0, 1].Font.Name := 'Symbol';
ExcelApp.Range['C32'].Characters[0, 1].Font.Underline := True;
ExcelApp.Range['C40'].Characters[0, 1].Font.Name := 'Symbol';
ExcelApp.Range['C44'].Characters[0, 1].Font.Underline := True;
ExcelApp.Range['C45'].Characters[0, 1].Font.Underline := True;
ExcelApp.Range['C46'].Characters[0, 1].Font.Underline := True;
ExcelApp.Range['C57'].Characters[0, 1].Font.Underline := True;
ExcelApp.Range['C58'].Characters[0, 1].Font.Underline := True;
ExcelApp.Range['C59'].Characters[0, 1].Font.Underline := True;
ExcelApp.Range['C74'].Characters[0, 1].Font.Underline := True;
ExcelApp.Range['C75'].Characters[0, 1].Font.Underline := True;
ExcelApp.Range['C77'].Characters[0, 1].Font.Underline := True;
ExcelApp.Range['C78'].Characters[0, 1].Font.Underline := True;
ExcelApp.Range['C79'].Characters[0, 1].Font.Underline := True;
ExcelApp.Range['C84'].Characters[0, 1].Font.Name := 'Symbol';
ExcelApp.Range['C85'].Characters[0, 1].Font.Name := 'Symbol';
ExcelApp.Range['C87'].Characters[0, 1].Font.Name := 'Symbol';
ExcelApp.Range['C88'].Characters[0, 1].Font.Name := 'Symbol';
ExcelApp.Range['C89'].Characters[0, 1].Font.Name := 'Symbol';
ExcelApp.Range['C97'].Characters[0, 1].Font.Name := 'Symbol';
ExcelApp.Range['C97'].Characters[0, 1].Font.Underline := True;
ExcelApp.Range['C102'].Characters[0, 1].Font.Name := 'Symbol';
ExcelApp.Range['C102'].Characters[0, 1].Font.Underline := True;
ExcelApp.Range['C103'].Characters[0, 1].Font.Name := 'Symbol';
ExcelApp.Range['C103'].Characters[0, 1].Font.Underline := True;
ExcelApp.Range['C104'].Characters[0, 1].Font.Name := 'Symbol';
ExcelApp.Range['C104'].Characters[0, 1].Font.Underline := True;
ExcelApp.Range['C108'].Characters[0, 1].Font.Name := 'Symbol';
ExcelApp.Range['C109'].Characters[0, 1].Font.Name := 'Symbol';
ExcelApp.Range['C115'].Characters[0, 1].Font.Name := 'Symbol';
ExcelApp.Range['C122'].Characters[0, 1].Font.Underline := True;
ExcelApp.Range['C124'].Characters[0, 1].Font.Name := 'Symbol';
ExcelApp.Range['C125'].Characters[0, 1].Font.Name := 'Symbol';
ExcelApp.Range['C127'].Characters[0, 1].Font.Name := 'Symbol';
ExcelApp.Range['C131'].Characters[0, 1].Font.Name := 'Symbol';
ExcelApp.Range['C131'].Characters[0, 1].Font.Underline := True;
ExcelApp.Range['C133'].Characters[0, 1].Font.Name := 'Symbol';
ExcelApp.Range['C133'].Characters[0, 1].Font.Underline := True;
ExcelApp.Range['C134'].Characters[0, 1].Font.Name := 'Symbol';
ExcelApp.Range['C134'].Characters[0, 1].Font.Underline := True;
ExcelApp.Range['C135'].Characters[0, 1].Font.Name := 'Symbol';
ExcelApp.Range['C135'].Characters[0, 1].Font.Underline := True;
ExcelApp.Range['C143'].Characters[0, 1].Font.Name := 'Symbol';
ExcelApp.Range['C145'].Characters[0, 1].Font.Name := 'Symbol';
ExcelApp.Range['C151'].Characters[0, 1].Font.Name := 'Symbol';
ExcelApp.Range['C158'].Characters[0, 1].Font.Name := 'Symbol';
ExcelApp.Range['C160'].Characters[0, 1].Font.Name := 'Symbol';
ExcelApp.Range['C161'].Characters[0, 1].Font.Name := 'Symbol';
ExcelApp.Range['C163'].Characters[0, 1].Font.Name := 'Symbol';
ExcelApp.Range['C172'].Characters[0, 1].Font.Name := 'Symbol';
ExcelApp.Range['C173'].Characters[0, 1].Font.Name := 'Symbol';
ExcelApp.Range['C175'].Characters[0, 1].Font.Name := 'Symbol';
ExcelApp.Range['C182'].Characters[0, 1].Font.Name := 'Symbol';
ExcelApp.Range['C183'].Characters[0, 1].Font.Name := 'Symbol';
ExcelApp.Range['C186'].Characters[0, 1].Font.Name := 'Symbol';
ExcelApp.Range['C187'].Characters[0, 1].Font.Name := 'Symbol';
ExcelApp.Range['C190'].Characters[0, 1].Font.Name := 'Symbol';
ExcelApp.Range['C190'].Characters[0, 1].Font.Underline := True;
ExcelApp.Range['C191'].Characters[0, 1].Font.Name := 'Symbol';
ExcelApp.Range['C191'].Characters[0, 1].Font.Underline := True;
ExcelApp.Range['C193'].Characters[0, 1].Font.Name := 'Symbol';
ExcelApp.Range['C193'].Characters[0, 1].Font.Underline := True;
ExcelApp.Range['C194'].Characters[0, 1].Font.Name := 'Symbol';
ExcelApp.Range['C194'].Characters[0, 1].Font.Underline := True;
ExcelApp.Range['C197'].Characters[0, 1].Font.Name := 'Symbol';
ExcelApp.Range['C199'].Characters[0, 1].Font.Name := 'Symbol';
ExcelApp.Range['C202'].Characters[0, 1].Font.Name := 'Symbol';
ExcelApp.Range['C203'].Characters[0, 1].Font.Name := 'Symbol';
ExcelApp.Range['C204'].Characters[0, 1].Font.Name := 'Symbol';




// ����������� �������
UpDownSymbol('C4:C22',777,777,2,8) ;
UpDownSymbol('C26:C26',4,10,2,10) ;
UpDownSymbol('C27:C27',6,10,2,10) ;
UpDownSymbol('C28:C28',777,777,2,10) ;
UpDownSymbol('C29:C29',777,777,2,10) ;
UpDownSymbol('C30:C30',777,777,2,10) ;
UpDownSymbol('C31:C31',4,10,2,10) ;
UpDownSymbol('C32:C32',777,777,2,10) ;
UpDownSymbol('C34:C34',4,10,2,10) ;
UpDownSymbol('C35:C35',5,10,2,10) ;
UpDownSymbol('C40:C40',777,777,3,10) ;
UpDownSymbol('C41:C41',777,777,2,10) ;
UpDownSymbol('C42:C42',777,777,2,10) ;
UpDownSymbol('C43:C43',777,777,2,10) ;
UpDownSymbol('C45:C45',777,777,2,10) ;
UpDownSymbol('C46:C46',777,777,2,10) ;
UpDownSymbol('C47:C47',777,777,2,10) ;
UpDownSymbol('C49:C49',777,777,2,10) ;
UpDownSymbol('C50:C50',777,777,2,10) ;
UpDownSymbol('C52:C52',777,777,777,777) ;
UpDownSymbol('C53:C53',3,10,2,10) ;
UpDownSymbol('C54:C54',777,777,2,10) ;
UpDownSymbol('C55:C55',777,777,2,10) ;
UpDownSymbol('C57:C57',777,777,2,10) ;
UpDownSymbol('C58:C58',777,777,2,10) ;
UpDownSymbol('C59:C59',777,777,2,10) ;
UpDownSymbol('C61:C61',777,777,2,10) ;
UpDownSymbol('C62:C62',777,777,2,10) ;
UpDownSymbol('C63:C63',777,777,2,10) ;
UpDownSymbol('C64:C64',777,777,2,10) ;
UpDownSymbol('C66:C66',777,777,2,10) ;
UpDownSymbol('C68:C68',5,10,2,10) ;
UpDownSymbol('C69:C69',4,10,2,10) ;
UpDownSymbol('C71:C71',5,10,2,10) ;
UpDownSymbol('C72:C72',4,10,2,10) ;
UpDownSymbol('C74:C74',4,10,2,10) ;
UpDownSymbol('C75:C75',5,10,2,10) ;
UpDownSymbol('C77:C77',4,10,2,10) ;
UpDownSymbol('C78:C78',5,10,2,10) ;
UpDownSymbol('C79:C79',4,10,2,10) ;
UpDownSymbol('C80:C80',777,777,2,10) ;
UpDownSymbol('C81:C81',777,777,2,10) ;
UpDownSymbol('C84:C84',777,777,2,10) ;
UpDownSymbol('C85:C85',4,10,2,10) ;
UpDownSymbol('C87:C87',777,777,2,10) ;
UpDownSymbol('C88:C88',4,10,2,10) ;
UpDownSymbol('C89:C89',777,777,2,10) ;
UpDownSymbol('C90:C90',777,777,3,10) ;
UpDownSymbol('C92:C92',777,777,3,10) ;
UpDownSymbol('C93:C93',5,10,3,10) ;
UpDownSymbol('C94:C94',777,777,3,10) ;
UpDownSymbol('C97:C97',777,777,2,10) ;
UpDownSymbol('C98:C98',4,10,2,10) ;
UpDownSymbol('C99:C99',777,777,2,10) ;
UpDownSymbol('C100:C100',777,777,2,10) ;
UpDownSymbol('C102:C102',777,777,2,10) ;
UpDownSymbol('C103:C103',4,10,2,10) ;
UpDownSymbol('C104:C104',777,777,2,10) ;
UpDownSymbol('C105:C105',777,777,2,10) ;
UpDownSymbol('C106:C106',777,777,3,10) ;
UpDownSymbol('C108:C108',777,777,2,10) ;
UpDownSymbol('C109:C109',4,10,2,10) ;
UpDownSymbol('C111:C111',777,777,3,10) ;
UpDownSymbol('C112:C112',5,10,3,10) ;
UpDownSymbol('C114:C114',777,777,2,10) ;
UpDownSymbol('C115:C115',777,777,2,10) ;
UpDownSymbol('C116:C116',777,777,2,10) ;
UpDownSymbol('C117:C117',777,777,2,10) ;
UpDownSymbol('C118:C118',777,777,3,10) ;
UpDownSymbol('C120:C120',777,777,2,10) ;
UpDownSymbol('C122:C122',4,10,2,10) ;
UpDownSymbol('C124:C124',4,10,2,10) ;
UpDownSymbol('C125:C125',4,10,2,10) ;
UpDownSymbol('C126:C126',777,777,2,10) ;
UpDownSymbol('C127:C127',777,777,3,10) ;
UpDownSymbol('C128:C128',777,777,2,10) ;
UpDownSymbol('C129:C129',777,777,2,10) ;
UpDownSymbol('C131:C131',777,777,2,10) ;
UpDownSymbol('C133:C133',777,777,2,10) ;
UpDownSymbol('C134:C134',4,10,2,10) ;
UpDownSymbol('C135:C135',777,777,2,10) ;
UpDownSymbol('C136:C136',4,10,2,10) ;
UpDownSymbol('C137:C137',777,777,2,10) ;
UpDownSymbol('C138:C138',777,777,2,10) ;
UpDownSymbol('C139:C139',777,777,2,10) ;
UpDownSymbol('C140:C140',777,777,3,10) ;
UpDownSymbol('C143:C143',777,777,2,10) ;
UpDownSymbol('C145:C145',4,10,2,10) ;
UpDownSymbol('C147:C147',777,777,3,10) ;
UpDownSymbol('C148:C148',5,10,3,10) ;
UpDownSymbol('C150:C150',777,777,2,10) ;
UpDownSymbol('C151:C151',777,777,2,10) ;
UpDownSymbol('C152:C152',777,777,2,10) ;
UpDownSymbol('C153:C153',777,777,2,10) ;
UpDownSymbol('C154:C154',777,777,3,10) ;
UpDownSymbol('C156:C156',777,777,2,10) ;
UpDownSymbol('C158:C158',4,10,2,10) ;
UpDownSymbol('C159:C159',4,10,2,10) ;
UpDownSymbol('C160:C160',777,777,2,10) ;
UpDownSymbol('C161:C161',4,10,2,10) ;
UpDownSymbol('C162:C162',777,777,2,10) ;
UpDownSymbol('C163:C163',777,777,3,10) ;
UpDownSymbol('C164:C164',777,777,2,10) ;
UpDownSymbol('C165:C165',777,777,2,10) ;
UpDownSymbol('C167:C167',777,777,2,10) ;
UpDownSymbol('C168:C168',777,777,2,10) ;
UpDownSymbol('C170:C170',5,10,2,10) ;
UpDownSymbol('C172:C172',777,777,2,10) ;
UpDownSymbol('C173:C173',5,10,2,10) ;
UpDownSymbol('C174:C174',777,777,2,10) ;
UpDownSymbol('C175:C175',777,777,3,10) ;
UpDownSymbol('C176:C176',777,777,2,10) ;
UpDownSymbol('C178:C178',777,777,2,10) ;
UpDownSymbol('C179:C179',777,777,2,10) ;
UpDownSymbol('C181:C181',777,777,2,10) ;
UpDownSymbol('C182:C182',777,777,2,10) ;
UpDownSymbol('C183:C183',777,777,3,10) ;
UpDownSymbol('C185:C185',777,777,2,10) ;
UpDownSymbol('C186:C186',777,777,2,10) ;
UpDownSymbol('C187:C187',777,777,3,10) ;
UpDownSymbol('C190:C190',777,777,777,777) ;
UpDownSymbol('C191:C191',777,777,777,777) ;
UpDownSymbol('C192:C192',777,777,2,10) ;
UpDownSymbol('C193:C193',777,777,2,10) ;
UpDownSymbol('C194:C194',777,777,2,10) ;
UpDownSymbol('C195:C195',777,777,2,10) ;
UpDownSymbol('C196:C196',777,777,3,10) ;
UpDownSymbol('C197:C197',777,777,2,10) ;
UpDownSymbol('C199:C199',777,777,3,10) ;
UpDownSymbol('C201:C201',777,777,2,10) ;
UpDownSymbol('C202:C202',777,777,2,10) ;
UpDownSymbol('C203:C203',777,777,3,10) ;
UpDownSymbol('C204:C204',777,777,3,10) ;



// ����������� �������
UpDownSymbol('C17:C22',4,6,777,777) ;

form4.ProgressBar1.position:=form4.ProgressBar1.position+5 ;    //((((((())))))

       /// Superscript  �����������  ExcelApp.Selection.Font.Subscript:= True
      ///    Subscript  �����������


// ����������� ����� � �� ����� � �����
ExcelApp.Range['A1:E1'].Select;
ExcelApp.Selection.MergeCells:=True ;
ExcelApp.Selection.HorizontalAlignment:=3;
ExcelApp.Selection.VerticalAlignment:=2 ;
ExcelApp.Selection.Font.Size := 15;
ExcelApp.Selection.Font.Name := 'Times New Roman';
ExcelApp.Selection:='�������� ������'  ;
ExcelApp.Range['A25:E25'].Select;
ExcelApp.Selection.MergeCells:=True ;
ExcelApp.Selection.HorizontalAlignment:=3;
ExcelApp.Selection.VerticalAlignment:=2 ;
ExcelApp.Selection.Font.Size := 15;
ExcelApp.Selection.Font.Name := 'Times New Roman';
ExcelApp.Selection:='������';

// ��������� �������� ������
ExcelApp.Range['A2'].Select;

form4.ProgressBar1.position:=form4.ProgressBar1.position+5 ;    //((((((())))))

  // ��������� (������� ���)
  if (CheckBox8.Checked=True) and (CheckBox3.Checked=True) then
    begin
       SaveDialog2.Title:='���������� �����������' ;
       SaveDialog2.filename := ExtractFilePath( Application.ExeName )+'����� (������)';
       if SaveDialog2.Execute then
         try  //(/�������� ������)
           ExcelApp.ActiveSheet.SaveAs(SaveDialog2.filename);
         except
         end;
    end;
// ������ Excel ������� (�������)
 if CheckBox3.Checked=False  then
 begin
   ExcelApp.Visible := False ;
   CheckBox1.Checked:=False  ;
   end
 else
   ExcelApp.Visible := True ;

  // ��������� (������� ���� � ������ ����)
  if (CheckBox8.Checked=True) and (CheckBox3.Checked=False) and (CheckBox2.Checked=False) then
    begin
       SaveDialog2.Title:='���������� �����������' ;
       SaveDialog2.filename := ExtractFilePath( Application.ExeName )+'����� (������)';
       if SaveDialog2.Execute then
         begin
             try  //(/�������� ������)
                ExcelApp.ActiveSheet.SaveAs(SaveDialog2.filename);
             except
             end;
        end;
             ExcelApp.DisplayAlerts := False;
             ExcelApp.Quit; // ������� Excel
             ExcelApp := Unassigned;
    end;

  // ��������� (������� ���� � ������ ���)
  if (CheckBox8.Checked=True) and (CheckBox3.Checked=False) and (CheckBox2.Checked=True)  then
    begin
       SaveDialog2.Title:='���������� �����������' ;
       SaveDialog2.filename := ExtractFilePath( Application.ExeName )+'����� (������)';
       if SaveDialog2.Execute then
         begin
             try  //(/�������� ������)
                ExcelApp.ActiveSheet.SaveAs(SaveDialog2.filename);
             except
             end;
        end;
    end;

  // ������
   if (CheckBox2.Checked=True) and (CheckBox3.Checked=True) then
      CheckBox1.Enabled:=True  ;
   if (CheckBox2.Checked=False) and (CheckBox3.Checked=False) then
      CheckBox1.Enabled:=False  ;
   if (CheckBox2.Checked=True) and (CheckBox1.Checked=False) then
     begin
       CheckBox1.Enabled:=True;
       ExcelApp.ActiveSheet.PrintOut  ;
       ExcelApp.DisplayAlerts := False;
       ExcelApp.Quit; // ������� Excel
       ExcelApp := Unassigned;
     end;
   if  (CheckBox2.Checked=true)  and (CheckBox1.Checked=true)  then
          begin
       Form4.close;
   form4.ProgressBar1.position:=0 ;
       ExcelApp.Dialogs[222].Show ;
          end;


form4.ProgressBar1.position:=100 ;    //((((((())))))


//     Form4.close;
//     form4.ProgressBar1.position:=0 ;
//   except  //(�������� ������)


//   end;
 Day ; ///����

   EnD ; ///[[{{{��������� Excel}}}]]






//       [[{{{��������� �������}}}]]


if CheckBox7.Checked=true  then
  BeGiN ///[[{{{��������� �������}}}]]

  //����������
t_icn_TH:=RoundSignificant(t_icn_TH,5);  // t_icn_TH
t_ak_TH:=RoundSignificant(t_ak_TH,5);   // t_ak_TH
t_s_ak:=RoundSignificant(t_s_ak,5);    // t_s_ak
W_1:=RoundSignificant(W_1,5);       // W_1
a_1 :=RoundSignificant(a_1 ,6);    // a_1
L_nr:=RoundSignificant(L_nr,5);   // L_nr
G_TH:=RoundSignificant(G_TH,5);   // G_TH
Delta_P_1_2:=RoundSignificant(Delta_P_1_2,5); // Delta_P_1_2
Delta_P_nr_2:=RoundSignificant(Delta_P_nr_2,5);  // Delta_P_nr_2

a_nn:=RoundSignificant(a_nn,7);    // a_nn
Delta_t_nn:=RoundSignificant(Delta_t_nn,5);  // Delta_t_nn
L_nn:=RoundSignificant(L_nn,5);   // L_nn
K_nn:=RoundSignificant(K_nn,5); // K_nn
Delta_P_nn_2:=RoundSignificant(Delta_P_nn_2,5); // Delta_P_nn_2
F__nn:=RoundSignificant(F__nn,5);    // F__nn
W_nn:=RoundSignificant(W_nn,5); // W_nn
Q_nn:=RoundSignificant(Q_nn,5); // Q_nn

a_icn:=RoundSignificant(a_icn,7);   // a_icn
Delta_t_icn:=RoundSignificant(Delta_t_icn,5); // Delta_t_icn
L_icn:=RoundSignificant(L_icn,5);  // L_icn
K_icn:=RoundSignificant(K_icn,5);   // K_icn
Delta_P_icn_2:=RoundSignificant(Delta_P_icn_2,5); // Delta_P_icn_2
F__icn:=RoundSignificant(F__icn,5);  // F__icn
W_0:=RoundSignificant(W_0,5);   // W_0
Q_icn:=RoundSignificant(Q_icn,5); // Q_icn

a_ak:=RoundSignificant(a_ak,7);  // a_ak
Delta_t_ak:=RoundSignificant(Delta_t_ak,5);  // Delta_t_ak
L_ak:=RoundSignificant(L_ak,5);    // L_ak
K_ak:=RoundSignificant(K_ak,5);    // K_ak
Delta_P_ak_2:=RoundSignificant(Delta_P_ak_2,5);  // Delta_P_ak_2
F__ak:=RoundSignificant(F__ak,5);  // F__ak
W_ak:=RoundSignificant(W_ak,5);   // W_ak
Q_ak:=RoundSignificant(Q_ak,5);  // Q_ak

     Form2.Memo1.Clear ;
Form2.Memo1.Lines.Add('              �������� ������');
Form2.Memo1.Lines.Add('');
Form2.Memo1.Lines.Add('���������������������� ����� �������     ��/�       '+FloatToStr(Dpac));
Form2.Memo1.Lines.Add('�������� ����������� ����                ���        '+FloatToStr(Pne));
Form2.Memo1.Lines.Add('����������� ����������� ����             ����\�     '+FloatToStr(tne));
Form2.Memo1.Lines.Add('����������� ����������� ����             ����\�     '+FloatToStr(tnb));
Form2.Memo1.Lines.Add('����������� �� �� ����� � ��             ����\�     '+FloatToStr(t_Bx));
Form2.Memo1.Lines.Add('�������� ������� �������                 ���        '+FloatToStr(P1));
Form2.Memo1.Lines.Add('����������� �� �� ������ �� ��           ����\�     '+FloatToStr(t_Bix));
Form2.Memo1.Lines.Add('�������� �����                                      '+(MAT_TP)) ;
Form2.Memo1.Lines.Add('�������� ������� �����                   ��         '+FloatToStr(d_H));
Form2.Memo1.Lines.Add('������� ������ ����                      ��         '+FloatToStr(btp));
Form2.Memo1.Lines.Add('��� ������� �������                      ��         '+FloatToStr(Stp));
Form2.Memo1.Lines.Add('����� ������������ ����                             '+FloatToStr(n_TP));
Form2.Memo1.Lines.Add('������ �� �� ��������� ����?                        '+(XOT_ShH)) ;
Form2.Memo1.Lines.Add('��� ����� �� �� �������                  ��         '+FloatToStr(S_shH_nn));
Form2.Memo1.Lines.Add('�������  ����� �� �� �������             ��         '+FloatToStr(d_shH_nn));
Form2.Memo1.Lines.Add('������� �������  ����� �� �� �������     ��         '+FloatToStr(b_shH_nn));
Form2.Memo1.Lines.Add('��� ����� �� �� �������                  ��         '+FloatToStr(S_shH_ak));
Form2.Memo1.Lines.Add('�������  ����� ��  ��  �������           ��         '+FloatToStr(d_shH_ak));
Form2.Memo1.Lines.Add('������� �������  ����� ��  ��  �������   ��         '+FloatToStr(b_shH_ak));
Form2.Memo1.Lines.Add('');
Form2.Memo1.Lines.Add('              ���������� �������');
Form2.Memo1.Lines.Add('');
Form2.Memo1.Lines.Add('����������� �� �� ������ �� �� ��-��     ����\�     '+FloatToStr(t_icn_TH) );
Form2.Memo1.Lines.Add('����������� �� �� ������ �� �� ��-��     ����\�     '+FloatToStr(t_ak_TH));
Form2.Memo1.Lines.Add('����������� 2� �� ����� � �� ��-��       ����\�     '+FloatToStr(t_s_ak));
Form2.Memo1.Lines.Add('������� �������� ��                      �2/�       '+FloatToStr(W_1));
Form2.Memo1.Lines.Add('����������� ����������� �� ��� ��-��     ��/(�2 ��) '+FloatToStr(a_1 ));
Form2.Memo1.Lines.Add('����� ������ �������                     �          '+FloatToStr(L_nr));
Form2.Memo1.Lines.Add('������ �� (1�)                           ��/�       '+FloatToStr(G_TH));
Form2.Memo1.Lines.Add('������� �������� (1�)                    ���        '+FloatToStr(Delta_P_1_2));
Form2.Memo1.Lines.Add('������� �������� (2�)                    ���        '+FloatToStr(Delta_P_nr_2));
Form2.Memo1.Lines.Add('');

Form2.Memo1.Lines.Add('���������� �� ��������:');
Form2.Memo1.Lines.Add('');
Form2.Memo1.Lines.Add('������������                      �����������   ��         ��         ��');
Form2.Memo1.Lines.Add('����������� ����������� (2�)      ��/(�2 ��)    '+FloatToStr(a_nn) +'               ' );
Form2.Memo1.Lines.Add('������������� �����               ��            '+FloatToStr(Delta_t_nn) +'               ' );
Form2.Memo1.Lines.Add('������ �������                    �             '+FloatToStr(L_nn) +'               ' );
Form2.Memo1.Lines.Add('����������� ������������� (2�)    ���/(�2 ��)   '+FloatToStr(K_nn) +'               ' );
Form2.Memo1.Lines.Add('������� �������� (2�)             ���           '+FloatToStr(Delta_P_nn_2) +'               ' );
Form2.Memo1.Lines.Add('������� �������                   �2            '+FloatToStr(F__nn) +'               ' );
Form2.Memo1.Lines.Add('�������� (2�)                     �/�           '+FloatToStr(W_nn) +'               ' );
Form2.Memo1.Lines.Add('������������ ���������� �������   ���           '+FloatToStr(Q_nn)+'               ' );

Form2.Memo1.CaretPos:=point(59,37);
Form2.memo1.seltext := FloatToStr(a_icn);
Form2.Memo1.CaretPos:=point(70,37);
Form2.memo1.seltext := FloatToStr(a_ak);
Form2.Memo1.CaretPos:=point(59,38);
Form2.memo1.seltext := FloatToStr(Delta_t_icn);
Form2.Memo1.CaretPos:=point(70,38);
Form2.memo1.seltext := FloatToStr(Delta_t_ak);
Form2.Memo1.CaretPos:=point(59,39);
Form2.memo1.seltext := FloatToStr(L_icn);
Form2.Memo1.CaretPos:=point(70,39);
Form2.memo1.seltext := FloatToStr(L_ak);
Form2.Memo1.CaretPos:=point(59,40);
Form2.memo1.seltext := FloatToStr(K_icn);
Form2.Memo1.CaretPos:=point(70,40);
Form2.memo1.seltext := FloatToStr(K_ak);
Form2.Memo1.CaretPos:=point(59,41);
Form2.memo1.seltext := FloatToStr(Delta_P_icn_2);
Form2.Memo1.CaretPos:=point(70,41);
Form2.memo1.seltext := FloatToStr(Delta_P_ak_2);
Form2.Memo1.CaretPos:=point(59,42);
Form2.memo1.seltext := FloatToStr(F__icn);
Form2.Memo1.CaretPos:=point(70,42);
Form2.memo1.seltext := FloatToStr(F__ak);
Form2.Memo1.CaretPos:=point(59,43);
Form2.memo1.seltext := FloatToStr(W_0);
Form2.Memo1.CaretPos:=point(70,43);
Form2.memo1.seltext := FloatToStr(W_ak);
Form2.Memo1.CaretPos:=point(59,44);
Form2.memo1.seltext := FloatToStr(Q_icn);
Form2.Memo1.CaretPos:=point(70,44);
Form2.memo1.seltext := FloatToStr(Q_ak);
//���������� ������� � �� ���� ������
Form2.Memo1.CaretPos:=point(0,0);

        //���������
     if CheckBox8.Checked=true  then  //C��������
        Begin
         Form2.Memo1.Lines.SaveToFile(ExtractFilePath( Application.ExeName )+'�����.txt');
        End;    //C��������
        //������
     if CheckBox2.Checked=true  then  //������
        If PrintDialog1.Execute then
          begin
               AssignPrn(Line);
               ReWrite(Line);
               Printer.Canvas.Font := Form2.Memo1.Font;
             for P := 0 to Form2.Memo1.Lines.Count -1 do Writeln (Line, Form2.Memo1.Lines[P]);
               System.CloseFile(Line);
          end; //������
        //�������
     if CheckBox3.Checked=true  then  //�������
        Begin
          Form2.ShowModal ;
        End;   //�������

  EnD ///[[{{{��������� �������}}}]]

end;



///    [[[����]]]

//           ((((����))))
// �����               (F12)
procedure TForm1.N2Click(Sender: TObject);
begin
Close;
end;
// ���������          (������)
procedure TForm1.N14Click(Sender: TObject);
begin
if PageControl1.ActivePage=TabSheet1 then
Button2.Click ;
end;
// ������� ���������  (Esc)
procedure TForm1.N15Click(Sender: TObject);
begin
if Knopka_P=0 then   //(������)
  begin
    Knopka_P:=1 ;
    Timer5.Enabled:= False ;
    PageControl1.ActivePage:=TabSheet1 ;
    Save_and_Load; // ��������� ��������� (��� ������ ���������)
       if interfeis='�������' then  Full.Checked:=true    //  Full ��� Lite ������
          else Lite.Checked:=true ;
    Exit;
  end;
if Knopka_P=1 then    //(���������)
  begin
    Knopka_P:=2;
    Form1.Width:=h_Close ;
    PageControl1.ActivePage:=TabSheet2 ;
    Exit;
  end;
if Knopka_P=2 then    //(������� ����)
  begin
    if CheckBox4.Checked=False  then  // ������  >
     begin
       Knopka_P:=1;
       PageControl1.ActivePage:=TabSheet1 ;
       Exit;
     end;
    if CheckBox4.Checked=True  then  //  ������  <
     begin
       Knopka_P:=1;
       Form1.Width:=h_Open ;
       PageControl1.ActivePage:=TabSheet1 ;
       Exit;
     end;
  end;
end;
// ������� ���������  (P)
procedure TForm1.N16Click(Sender: TObject);
begin
if Knopka_P=0 then    //(������)
    Exit;
if Knopka_P=1 then    //(���������)
  begin
    Knopka_P:=2;
    if (Timer1.Enabled=True) or (Timer2.Enabled=True) or (Timer3.Enabled=True) or (Timer4.Enabled=True) then
      begin
        Exit;
      end
    else
      begin
        Form1.Width:=h_Close ;
        PageControl1.ActivePage:=TabSheet2 ;
        Exit;
      end
  end;
if Knopka_P=2 then    //(������� ����)
  begin
    if CheckBox4.Checked=False  then  // ������  >
     begin
       Knopka_P:=1;
       PageControl1.ActivePage:=TabSheet1 ;
       Exit;
     end;
    if CheckBox4.Checked=True  then  //  ������  <
     begin
       Knopka_P:=1;
       Form1.Width:=h_Open ;
       PageControl1.ActivePage:=TabSheet1 ;
       Exit;
     end;
  end;
end;

//���������
  // ���������
    //  Lite
procedure TForm1.LiteClick(Sender: TObject);
begin
Lite.Checked:=true  ;
interfeis:='������' ;
//���������� ��������
if (Timer1.Enabled= True) and (CheckBox4.Checked=true) then   // ������ ��������� ������
   begin
       Form1.Width:=h_Open    ;
       Timer1.Enabled:= False ;   //()
       Timer2.Enabled:= False  ;
       Timer3.Enabled:= False  ;
       Timer4.Enabled:= False  ;
       Exit ;
   end;
if (Timer1.Enabled= True) and (CheckBox4.Checked=False) then  // ������ �������� ������
   begin
       Form1.Width:=h_Close    ;
       Timer1.Enabled:= False  ;   //()
       Timer2.Enabled:= False  ;
       Timer3.Enabled:= False  ;
       Timer4.Enabled:= False  ;
       Exit ;
   end;
if  Timer2.Enabled= True  then   // ������ ������  (���������)
   begin
       Form1.Width:=h_Women    ;
       Timer1.Enabled:= False  ;
       Timer2.Enabled:= False  ;   //()
       Timer3.Enabled:= False  ;
       Timer4.Enabled:= False  ;
       Exit ;
   end;
if  Timer3.Enabled= True  then   // ������ ������  (������)
   begin
       Form1.Width:=h_Open     ;
       Timer1.Enabled:= False  ;
       Timer2.Enabled:= False  ;
       Timer3.Enabled:= False  ;   //()
       Timer4.Enabled:= False  ;
       Exit ;
   end;
if  Timer4.Enabled= True  then   // ������ �������� �� �����
   begin
       Form1.Width:=h_Close    ;
       Timer1.Enabled:= False  ;
       Timer2.Enabled:= False  ;
       Timer3.Enabled:= False  ;   
       Timer4.Enabled:= False  ;   //()
       Exit ;
   end;
end;
    //  Full
procedure TForm1.FullClick(Sender: TObject);
begin
Full.Checked:=true   ;
interfeis:='�������' ;
end;

 // ��������
     // �������       (F6)
procedure TForm1.N9Click(Sender: TObject);
var SR:TSearchRec; // ��������� ����������
begin
if PageControl1.ActivePage=TabSheet2 then
  begin
      Application.Messagebox('�� ��������. ������� ���� �������� ������','������ !!!', mb_iconerror or mb_ok);
      Day ; Exit;
  end;
FindFirst(ExtractFilePath( Application.ExeName )+'���������.ini',faAnyFile,SR); // ������� ������� ������ � ������ ������
  if SR.name <> '' then   //���� ������
    begin
     FindClose(SR); // ��������� �����
     Memo1.Clear ;
     Memo1.Lines.LoadFromFile(ExtractFilePath( Application.ExeName )+'���������.ini');
     SelLineTime(86); //��������� � "SelLineTime" �����
    end
   else  //���� �� ������
    begin
      FindClose(SR); // ��������� �����
      Schetchik:='0' ;
    end;
Memo1.Clear ;
Memo1.Lines.Add('D���  (���������������������� ����� �������) [��/�]');
SelSaveEdit(Edit_Dpac);
Memo1.Lines.Add('���   (�������� ����������� ����)            [���]');
SelSaveEdit(Edit_Pne);
Memo1.Lines.Add('t��   (����������� ����������� ����)         [o�]');
SelSaveEdit(Edit_tne);
Memo1.Lines.Add('t��   (����������� ����������� ����)         [o�]');
SelSaveEdit(Edit_tnb);
Memo1.Lines.Add('t��   (����������� �� �� ����� � ��)         [o�]');
SelSaveEdit(Edit_t_Bx);
Memo1.Lines.Add('P1    (�������� ������� �������)             [���]');
SelSaveEdit(Edit_P1);
Memo1.Lines.Add('t���  (����������� �� �� ������ �� ��)       [o�]');
SelSaveEdit(Edit_t_Bix);
Memo1.Lines.Add('�������� �����');
SelSaveComboBox(ComboBox1);
Memo1.Lines.Add('d�    (�������� ������� �����)               [��]');
SelSaveEdit(Edit_d_H);
Memo1.Lines.Add('b��   (������� ������ ����)                  [��]');
SelSaveEdit(Edit_btp);
Memo1.Lines.Add('S��   (��� ������� �������)                  [��]');
SelSaveEdit(Edit_Stp);
Memo1.Lines.Add('n��   (����� ������������ ����)              [��]');
SelSaveEdit(Edit_ntp);
Memo1.Lines.Add('������ �� �� ��������� ����?');
SelSaveComboBox(ComboBox2);
Memo1.Lines.Add('S����  (��� ����� �� �� �������)             [��]');
SelSaveEdit(Edit_Ssh);
Memo1.Lines.Add('d����  (�������  ����� �� �� �������)        [��]');
SelSaveEdit(Edit_dsh);
Memo1.Lines.Add('b����  (������� ������� ����� �� �� �������) [��]');
SelSaveEdit(Edit_b_shH_nn);
Memo1.Lines.Add('S����  (��� ����� �� �� �������)             [��]');
SelSaveEdit(Edit_S_shH_ak);
Memo1.Lines.Add('d����  (�������  ����� ��  ��  �������)      [��]');
SelSaveEdit(Edit_d_shH_ak);
Memo1.Lines.Add('b����  (������� ������� ����� �� �� �������) [��]');
SelSaveEdit(Edit_b_shH_ak);
Memo1.Lines.Add('������');
Memo1.Lines.Add('���������');
SelSaveCheckBox(CheckBox8) ;
Memo1.Lines.Add('������');
SelSaveCheckBox(CheckBox2) ;
Memo1.Lines.Add('������ >  < ������');
SelSaveCheckBox(CheckBox4) ;
Memo1.Lines.Add('�������');
SelSaveCheckBox(CheckBox3) ;
Memo1.Lines.Add('�����������');
SelSaveCheckBox(CheckBox9) ;
Memo1.Lines.Add('������');
SelSaveCheckBox(CheckBox10) ;
Memo1.Lines.Add('Excel');
SelSaveCheckBox(CheckBox6) ;
Memo1.Lines.Add('�������');
SelSaveCheckBox(CheckBox7) ;
Memo1.Lines.Add('������������');
SelSaveCheckBox(CheckBox1) ;
Memo1.Lines.Add('�������');
Schetchik:=EditTime.Text; //��������� ������� �� Edit
Memo1.Lines.Add(Schetchik);
Memo1.Lines.Add('"�������" ��� � ��� ���������?');
Memo1.Lines.Add(interfeis);
Memo1.Lines.SaveToFile(ExtractFilePath( Application.ExeName )+'���������.ini');
Edit_Dpac .setfocus  ;
end;
    // ������
procedure TForm1.N10Click(Sender: TObject);
begin
if PageControl1.ActivePage=TabSheet2 then
  begin
      Application.Messagebox('�� ��������. ������� ���� �������� ������','������ !!!', mb_iconerror or mb_ok);
      Day ; Exit;
  end;
SaveDialog1.Title:='���������� �������� ������' ;
SaveDialog1.filename := ExtractFilePath( Application.ExeName )+'���������';
if SaveDialog1.Execute then
  try  //
    Memo1.Clear ;
    Memo1.Lines.Add('D���  (���������������������� ����� �������) [��/�]');
    SelSaveEdit(Edit_Dpac);
    Memo1.Lines.Add('���   (�������� ����������� ����)            [���]');
    SelSaveEdit(Edit_Pne);
    Memo1.Lines.Add('t��   (����������� ����������� ����)         [o�]');
    SelSaveEdit(Edit_tne);
    Memo1.Lines.Add('t��   (����������� ����������� ����)         [o�]');
    SelSaveEdit(Edit_tnb);
    Memo1.Lines.Add('t��   (����������� �� �� ����� � ��)         [o�]');
    SelSaveEdit(Edit_t_Bx);
    Memo1.Lines.Add('P1    (�������� ������� �������)             [���]');
    SelSaveEdit(Edit_P1);
    Memo1.Lines.Add('t���  (����������� �� �� ������ �� ��)       [o�]');
    SelSaveEdit(Edit_t_Bix);
    Memo1.Lines.Add('�������� �����');
    SelSaveComboBox(ComboBox1);
    Memo1.Lines.Add('d�    (�������� ������� �����)               [��]');
    SelSaveEdit(Edit_d_H);
    Memo1.Lines.Add('b��   (������� ������ ����)                  [��]');
    SelSaveEdit(Edit_btp);
    Memo1.Lines.Add('S��   (��� ������� �������)                  [��]');
    SelSaveEdit(Edit_Stp);
    Memo1.Lines.Add('n��   (����� ������������ ����)              [��]');
    SelSaveEdit(Edit_ntp);
    Memo1.Lines.Add('������ �� �� ��������� ����?');
    SelSaveComboBox(ComboBox2);
    Memo1.Lines.Add('S����  (��� ����� �� �� �������)             [��]');
    SelSaveEdit(Edit_Ssh);
    Memo1.Lines.Add('d����  (�������  ����� �� �� �������)        [��]');
    SelSaveEdit(Edit_dsh);
    Memo1.Lines.Add('b����  (������� ������� ����� �� �� �������) [��]');
    SelSaveEdit(Edit_b_shH_nn);
    Memo1.Lines.Add('S����  (��� ����� �� �� �������)             [��]');
    SelSaveEdit(Edit_S_shH_ak);
    Memo1.Lines.Add('d����  (�������  ����� ��  ��  �������)      [��]');
    SelSaveEdit(Edit_d_shH_ak);
    Memo1.Lines.Add('b����  (������� ������� ����� �� �� �������) [��]');
    SelSaveEdit(Edit_b_shH_ak);
    Memo1.Lines.Add('������');
    Memo1.Lines.Add('���������');
    SelSaveCheckBox(CheckBox8) ;
    Memo1.Lines.Add('������');
    SelSaveCheckBox(CheckBox2) ;
    Memo1.Lines.Add('������ >  < ������');
    SelSaveCheckBox(CheckBox4) ;
    Memo1.Lines.Add('�������');
    SelSaveCheckBox(CheckBox3) ;
    Memo1.Lines.Add('�����������');
    SelSaveCheckBox(CheckBox9) ;
    Memo1.Lines.Add('������');
    SelSaveCheckBox(CheckBox10) ;
    Memo1.Lines.Add('Excel');
    SelSaveCheckBox(CheckBox6) ;
    Memo1.Lines.Add('�������');
    SelSaveCheckBox(CheckBox7) ;
    Memo1.Lines.Add('������������');
    SelSaveCheckBox(CheckBox1) ;
    Memo1.Lines.Add('�������');
    Schetchik:=EditTime.Text; //��������� ������� �� Edit
    Memo1.Lines.Add(Schetchik);
    Memo1.Lines.Add('"�������" ��� � ��� ���������?');
    Memo1.Lines.Add(interfeis);
      if FileExists(SaveDialog1.FileName) then   // ���� ����� ���� ��� ���� �� ��������������
         begin
          DeleteFile(SaveDialog1.FileName);
          Memo1.Lines.SaveToFile(SaveDialog1.FileName)
         end
      else //���� ��� ��� �� ���������� � �������� .txt
     Memo1.Lines.SaveToFile(SaveDialog1.FileName+'.txt')
  except
  end;
end;

  // ��������
       // �������      (F8)
procedure TForm1.N11Click(Sender: TObject);
Var SR:TSearchRec; // ��������� ����������
begin
if PageControl1.ActivePage=TabSheet2 then
  begin
      Application.Messagebox('�� ��������. ������� ���� �������� ������','������ !!!', mb_iconerror or mb_ok);
      Day ; Exit;
  end;
FindFirst(ExtractFilePath( Application.ExeName )+'���������.ini',faAnyFile,SR); // ������� ������� ������ � ������ ������
if SR.name <> '' then   //���� ������
  begin
   FindClose(SR); // ��������� �����
   Memo1.Clear ;
   Memo1.Lines.LoadFromFile(ExtractFilePath( Application.ExeName )+'���������.ini');
   SelLineEdit(Edit_Dpac,2) ;       // Dpac
   SelLineEdit(Edit_Pne,5) ;        // Pne
   SelLineEdit(Edit_tne,8) ;        // tne
   SelLineEdit(Edit_tnb,11) ;       // tnb
   SelLineEdit(Edit_t_Bx,14) ;      // t_Bx
   SelLineEdit(Edit_P1,17) ;        // P1
   SelLineEdit(Edit_t_Bix,20) ;     // t_Bix
   SelLineComboBox(ComboBox1,23) ;  // �������� �����
   SelLineEdit(Edit_d_H,26) ;       // d_H
   SelLineEdit(Edit_btp,29) ;       // b_tp
   SelLineEdit(Edit_Stp,32) ;       // S_tp
   SelLineEdit(Edit_nTP,35) ;       // n_TP
   SelLineComboBox(ComboBox2,38) ;  // ������ �� �� ��������� ����?
   SelLineEdit(Edit_Ssh,41) ;       // S_sh_nn
   SelLineEdit(Edit_dsh,44) ;       // d_sh_nn
   SelLineEdit(Edit_b_shH_nn,47) ;  // b_shH_nn
   SelLineEdit(Edit_S_shH_ak,50) ;  // S_shH_ak
   SelLineEdit(Edit_d_shH_ak,53) ;  // d_shH_ak
   SelLineEdit(Edit_b_shH_ak,56) ;  // b_shH_ak
   //�����
   SelLineCheckBox(CheckBox8,60) ; //  ���������
   SelLineCheckBox(CheckBox2,63) ; //  ������
   SelLineCheckBox(CheckBox4,66) ; //  ������ >
   SelLineCheckBox(CheckBox3,69) ; //  �������
   SelLineCheckBox(CheckBox6,78) ; //  Excel
   SelLineCheckBox(CheckBox7,81) ; //  �������
   SelLineCheckBox(CheckBox9,72) ; //  �����������
   SelLineCheckBox(CheckBox10,75); //  ������
   SelLineCheckBox(CheckBox1,84) ; //  ������������
   SelLineInterfeis(88); //��������� � "Interfeis" �����
      if interfeis='�������' then  Full.Checked:=true    //  Full ��� Lite ������
         else Lite.Checked:=true ;
   Edit_Dpac .setfocus  ;
  end
 else  //���� �� ������
  begin
    FindClose(SR); // ��������� �����
    Application.Messagebox('����� ������ ��� ��? ������ ������ ���� "���������" �� �����!','������ !!!', mb_iconerror or mb_ok);
    Edit_Dpac .setfocus  ;
  end;
end;
    // ������
procedure TForm1.N12Click(Sender: TObject);
begin
if PageControl1.ActivePage=TabSheet2 then
  begin
      Application.Messagebox('�� ��������. ������� ���� �������� ������','������ !!!', mb_iconerror or mb_ok);
      Day ; Exit;
  end;
OpenDialog1.Title:='����� �������� ������' ;
OpenDialog1.filename := ExtractFilePath( Application.ExeName )+'���������.txt';
if OpenDialog1.Execute then
  try
   Memo1.Clear ;
   Memo1.Lines.LoadFromFile(OpenDialog1.FileName);
   SelLineEdit(Edit_Dpac,2) ;       // Dpac
   SelLineEdit(Edit_Pne,5) ;        // Pne
   SelLineEdit(Edit_tne,8) ;        // tne
   SelLineEdit(Edit_tnb,11) ;       // tnb
   SelLineEdit(Edit_t_Bx,14) ;      // t_Bx
   SelLineEdit(Edit_P1,17) ;        // P1
   SelLineEdit(Edit_t_Bix,20) ;     // t_Bix
   SelLineComboBox(ComboBox1,23) ;  // �������� �����
   SelLineEdit(Edit_d_H,26) ;       // d_H
   SelLineEdit(Edit_btp,29) ;       // b_tp
   SelLineEdit(Edit_Stp,32) ;       // S_tp
   SelLineEdit(Edit_nTP,35) ;       // n_TP
   SelLineComboBox(ComboBox2,38) ;  // ������ �� �� ��������� ����?
   SelLineEdit(Edit_Ssh,41) ;       // S_sh_nn
   SelLineEdit(Edit_dsh,44) ;       // d_sh_nn
   SelLineEdit(Edit_b_shH_nn,47) ;  // b_shH_nn
   SelLineEdit(Edit_S_shH_ak,50) ;  // S_shH_ak
   SelLineEdit(Edit_d_shH_ak,53) ;  // d_shH_ak
   SelLineEdit(Edit_b_shH_ak,56) ;  // b_shH_ak
   //�����
   SelLineCheckBox(CheckBox8,60) ; //  ���������
   SelLineCheckBox(CheckBox2,63) ; //  ������
   SelLineCheckBox(CheckBox4,66) ; //  ������ >
   SelLineCheckBox(CheckBox3,69) ; //  �������
   SelLineCheckBox(CheckBox6,78) ; //  Excel
   SelLineCheckBox(CheckBox7,81) ; //  �������
   SelLineCheckBox(CheckBox9,72) ; //  �����������
   SelLineCheckBox(CheckBox10,75); //  ������
   SelLineCheckBox(CheckBox1,84) ; //  ������������
   SelLineInterfeis(88); //��������� � "Interfeis" �����
      if interfeis='�������' then  Full.Checked:=true    //  Full ��� Lite ������
         else Lite.Checked:=true ;
   Edit_Dpac .setfocus  ;
  except
  end;
end;

//����������
procedure TForm1.N13Click(Sender: TObject);
begin
Form6.Vkl.Visible:=true ;   //����� �� ������� ������
Form6.VbIkl.Visible:=false ;//�� ����� ������� ������
Form6.ShowModal;
end;


// � ��������� (������� � Form5)     (F1)
procedure TForm1.N4Click(Sender: TObject);
var Forma5: integer;
begin
Forma5:=Form5.ShowModal;
end;


///   [[����� � ������� � ������ �����]]

//1) Dpac
procedure TForm1.Edit_DpacKeyPress(Sender: TObject; var Key: Char);
Var k:integer;
begin
//������ ����� �� �������
if  Key='.' then
Key:=',';
//������� ������� [Esc]
if  Key=Chr(VK_ESCAPE) then
Key:=Chr( VK_BACK );
//���� ����
if Edit_Dpac.Text='0' then          //���� ���������, �� ����� ����
 if not(key in [',','.',#8]) then   //������ ������ �������
       begin                        //���� ������ ������� ���
        key:=#0;
        beep ;
       end;
if key in['0'..'9',',','.',#8] then //��������� ������� ������ �����
  begin
  if (key=',') or (key='.') then //�������� ��� ������ ����� �������
    begin
    if Edit_Dpac.Text='' then key:=#0;
      For k:=1 to Length(Edit_Dpac.Text) do
      begin
       if (Edit_Dpac.Text[k]=',') or (Edit_Dpac.Text[k]=',')then
                begin
                 key:=#0;
                 beep ;
                end;
      end;
    end;
  end else key:=#0;
end;
//2)Pne
procedure TForm1.Edit_PneKeyPress(Sender: TObject; var Key: Char);
Var k:integer;
begin
//������ ����� �� �������
if  Key='.' then
Key:=',';
//������� ������� [Esc]
if  Key=Chr(VK_ESCAPE) then
Key:=Chr( VK_BACK );
//���� ����
if Edit_Pne.Text='0' then          //���� ���������, �� ����� ����
 if not(key in [',','.',#8]) then   //������ ������ �������
       begin                        //���� ������ ������� ���
        key:=#0;
        beep ;
       end;
if key in['0'..'9',',','.',#8] then //��������� ������� ������ �����
  begin
  if (key=',') or (key='.') then //�������� ��� ������ ����� �������
    begin
    if Edit_Pne.Text='' then key:=#0;
      For k:=1 to Length(Edit_Pne.Text) do
      begin
       if (Edit_Pne.Text[k]=',') or (Edit_Pne.Text[k]=',')then
                begin
                 key:=#0;
                 beep ;
                end;
      end;
    end;
  end else key:=#0;
end;
//3)tne
procedure TForm1.Edit_tneKeyPress(Sender: TObject; var Key: Char);
Var k:integer;
begin
//������ ����� �� �������
if  Key='.' then
Key:=',';
//������� ������� [Esc]
if  Key=Chr(VK_ESCAPE) then
Key:=Chr( VK_BACK );
//���� ����
if Edit_tne.Text='0' then          //���� ���������, �� ����� ����
 if not(key in [',','.',#8]) then   //������ ������ �������
       begin                        //���� ������ ������� ���
        key:=#0;
        beep ;
       end;
if key in['0'..'9',',','.',#8] then //��������� ������� ������ �����
  begin
  if (key=',') or (key='.') then //�������� ��� ������ ����� �������
    begin
    if Edit_tne.Text='' then key:=#0;
      For k:=1 to Length(Edit_tne.Text) do
      begin
       if (Edit_tne.Text[k]=',') or (Edit_tne.Text[k]=',')then
                begin
                 key:=#0;
                 beep ;
                end;
      end;
    end;
  end else key:=#0;
end;
//4)tnb
procedure TForm1.Edit_tnbKeyPress(Sender: TObject; var Key: Char);
Var k:integer;
begin
//������ ����� �� �������
if  Key='.' then
Key:=',';
//������� ������� [Esc]
if  Key=Chr(VK_ESCAPE) then
Key:=Chr( VK_BACK );
//���� ����
if Edit_tnb.Text='0' then          //���� ���������, �� ����� ����
 if not(key in [',','.',#8]) then   //������ ������ �������
       begin                        //���� ������ ������� ���
        key:=#0;
        beep ;
       end;
if key in['0'..'9',',','.',#8] then //��������� ������� ������ �����
  begin
  if (key=',') or (key='.') then //�������� ��� ������ ����� �������
    begin
    if Edit_tnb.Text='' then key:=#0;
      For k:=1 to Length(Edit_tnb.Text) do
      begin
       if (Edit_tnb.Text[k]=',') or (Edit_tnb.Text[k]=',')then
                begin
                 key:=#0;
                 beep ;
                end;
      end;
    end;
  end else key:=#0;
end;
//5)t_Bx
procedure TForm1.Edit_t_BxKeyPress(Sender: TObject; var Key: Char);
Var k:integer;
begin
//������ ����� �� �������
if  Key='.' then
Key:=',';
//������� ������� [Esc]
if  Key=Chr(VK_ESCAPE) then
Key:=Chr( VK_BACK );
//���� ����
if Edit_t_Bx.Text='0' then          //���� ���������, �� ����� ����
 if not(key in [',','.',#8]) then   //������ ������ �������
       begin                        //���� ������ ������� ���
        key:=#0;
        beep ;
       end;
if key in['0'..'9',',','.',#8] then //��������� ������� ������ �����
  begin
  if (key=',') or (key='.') then //�������� ��� ������ ����� �������
    begin
    if Edit_t_Bx.Text='' then key:=#0;
      For k:=1 to Length(Edit_t_Bx.Text) do
      begin
       if (Edit_t_Bx.Text[k]=',') or (Edit_t_Bx.Text[k]=',')then
                begin
                 key:=#0;
                 beep ;
                end;
      end;
    end;
  end else key:=#0;
end;
//6)P1
procedure TForm1.Edit_P1KeyPress(Sender: TObject; var Key: Char);
Var k:integer;
begin
//������ ����� �� �������
if  Key='.' then
Key:=',';
//������� ������� [Esc]
if  Key=Chr(VK_ESCAPE) then
Key:=Chr( VK_BACK );
//���� ����
if Edit_P1.Text='0' then          //���� ���������, �� ����� ����
 if not(key in [',','.',#8]) then   //������ ������ �������
       begin                        //���� ������ ������� ���
        key:=#0;
        beep ;
       end;
if key in['0'..'9',',','.',#8] then //��������� ������� ������ �����
  begin
  if (key=',') or (key='.') then //�������� ��� ������ ����� �������
    begin
    if Edit_P1.Text='' then key:=#0;
      For k:=1 to Length(Edit_P1.Text) do
      begin
       if (Edit_P1.Text[k]=',') or (Edit_P1.Text[k]=',')then
                begin
                 key:=#0;
                 beep ;
                end;
      end;
    end;
  end else key:=#0;
end;
//7)t_Bix
procedure TForm1.Edit_t_BixKeyPress(Sender: TObject; var Key: Char);
Var k:integer;
begin
//������ ����� �� �������
if  Key='.' then
Key:=',';
//������� ������� [Esc]
if  Key=Chr(VK_ESCAPE) then
Key:=Chr( VK_BACK );
//���� ����
if Edit_t_Bix.Text='0' then          //���� ���������, �� ����� ����
 if not(key in [',','.',#8]) then   //������ ������ �������
       begin                        //���� ������ ������� ���
        key:=#0;
        beep ;
       end;
if key in['0'..'9',',','.',#8] then //��������� ������� ������ �����
  begin
  if (key=',') or (key='.') then //�������� ��� ������ ����� �������
    begin
    if Edit_t_Bix.Text='' then key:=#0;
      For k:=1 to Length(Edit_t_Bix.Text) do
      begin
       if (Edit_t_Bix.Text[k]=',') or (Edit_t_Bix.Text[k]=',')then
                begin
                 key:=#0;
                 beep ;
                end;
      end;
    end;
  end else key:=#0;
end;
//9)d_H
procedure TForm1.Edit_d_HKeyPress(Sender: TObject; var Key: Char);
Var k:integer;
begin
//������ ����� �� �������
if  Key='.' then
Key:=',';
//������� ������� [Esc]
if  Key=Chr(VK_ESCAPE) then
Key:=Chr( VK_BACK );
//���� ����
if Edit_d_H.Text='0' then          //���� ���������, �� ����� ����
 if not(key in [',','.',#8]) then   //������ ������ �������
       begin                        //���� ������ ������� ���
        key:=#0;
        beep ;
       end;
if key in['0'..'9',',','.',#8] then //��������� ������� ������ �����
  begin
  if (key=',') or (key='.') then //�������� ��� ������ ����� �������
    begin
    if Edit_d_H.Text='' then key:=#0;
      For k:=1 to Length(Edit_d_H.Text) do
      begin
       if (Edit_d_H.Text[k]=',') or (Edit_d_H.Text[k]=',')then
                begin
                 key:=#0;
                 beep ;
                end;
      end;
    end;
  end else key:=#0;
end;
//10)b_tp
procedure TForm1.Edit_btpKeyPress(Sender: TObject; var Key: Char);
Var k:integer;
begin
//������ ����� �� �������
if  Key='.' then
Key:=',';
//������� ������� [Esc]
if  Key=Chr(VK_ESCAPE) then
Key:=Chr( VK_BACK );
//���� ����
if Edit_btp.Text='0' then          //���� ���������, �� ����� ����
 if not(key in [',','.',#8]) then   //������ ������ �������
       begin                        //���� ������ ������� ���
        key:=#0;
        beep ;
       end;
if key in['0'..'9',',','.',#8] then //��������� ������� ������ �����
  begin
  if (key=',') or (key='.') then //�������� ��� ������ ����� �������
    begin
    if Edit_btp.Text='' then key:=#0;
      For k:=1 to Length(Edit_btp.Text) do
      begin
       if (Edit_btp.Text[k]=',') or (Edit_btp.Text[k]=',')then
                begin
                 key:=#0;
                 beep ;
                end;
      end;
    end;
  end else key:=#0;
end;
//11)Stp
procedure TForm1.Edit_StpKeyPress(Sender: TObject; var Key: Char);
Var k:integer;
begin
//������ ����� �� �������
if  Key='.' then
Key:=',';
//������� ������� [Esc]
if  Key=Chr(VK_ESCAPE) then
Key:=Chr( VK_BACK );
//���� ����
if Edit_Stp.Text='0' then          //���� ���������, �� ����� ����
 if not(key in [',','.',#8]) then   //������ ������ �������
       begin                        //���� ������ ������� ���
        key:=#0;
        beep ;
       end;
if key in['0'..'9',',','.',#8] then //��������� ������� ������ �����
  begin
  if (key=',') or (key='.') then //�������� ��� ������ ����� �������
    begin
    if Edit_Stp.Text='' then key:=#0;
      For k:=1 to Length(Edit_Stp.Text) do
      begin
       if (Edit_Stp.Text[k]=',') or (Edit_Stp.Text[k]=',')then
                begin
                 key:=#0;
                 beep ;
                end;
      end;
    end;
  end else key:=#0;
end;
//12)ntp
procedure TForm1.Edit_ntpKeyPress(Sender: TObject; var Key: Char);
Var k:integer;
begin
//������ ����� �� �������
if  Key='.' then
Key:=',';
//������� ������� [Esc]
if  Key=Chr(VK_ESCAPE) then
Key:=Chr( VK_BACK );
//���� ����
if Edit_ntp.Text='0' then          //���� ���������, �� ����� ����
 if not(key in [',','.',#8]) then   //������ ������ �������
       begin                        //���� ������ ������� ���
        key:=#0;
        beep ;
       end;
if key in['0'..'9',',','.',#8] then //��������� ������� ������ �����
  begin
  if (key=',') or (key='.') then //�������� ��� ������ ����� �������
    begin
    if Edit_ntp.Text='' then key:=#0;
      For k:=1 to Length(Edit_ntp.Text) do
      begin
       if (Edit_ntp.Text[k]=',') or (Edit_ntp.Text[k]=',')then
                begin
                 key:=#0;
                 beep ;
                end;
      end;
    end;
  end else key:=#0;
end;
//14)Ssh
procedure TForm1.Edit_SshKeyPress(Sender: TObject; var Key: Char);
Var k:integer;
begin
//������ ����� �� �������
if  Key='.' then
Key:=',';
//������� ������� [Esc]
if  Key=Chr(VK_ESCAPE) then
Key:=Chr( VK_BACK );
//���� ����
if Edit_Ssh.Text='0' then          //���� ���������, �� ����� ����
 if not(key in [',','.',#8]) then   //������ ������ �������
       begin                        //���� ������ ������� ���
        key:=#0;
        beep ;
       end;
if key in['0'..'9',',','.',#8] then //��������� ������� ������ �����
  begin
  if (key=',') or (key='.') then //�������� ��� ������ ����� �������
    begin
    if Edit_Ssh.Text='' then key:=#0;
      For k:=1 to Length(Edit_Ssh.Text) do
      begin
       if (Edit_Ssh.Text[k]=',') or (Edit_Ssh.Text[k]=',')then
                begin
                 key:=#0;
                 beep ;
                end;
      end;
    end;
  end else key:=#0;
end;
//15)dsh
procedure TForm1.Edit_dshKeyPress(Sender: TObject; var Key: Char);
Var k:integer;
begin
//������ ����� �� �������
if  Key='.' then
Key:=',';
//������� ������� [Esc]
if  Key=Chr(VK_ESCAPE) then
Key:=Chr( VK_BACK );
//���� ����
if Edit_dsh.Text='0' then          //���� ���������, �� ����� ����
 if not(key in [',','.',#8]) then   //������ ������ �������
       begin                        //���� ������ ������� ���
        key:=#0;
        beep ;
       end;
if key in['0'..'9',',','.',#8] then //��������� ������� ������ �����
  begin
  if (key=',') or (key='.') then //�������� ��� ������ ����� �������
    begin
    if Edit_dsh.Text='' then key:=#0;
      For k:=1 to Length(Edit_Dpac.Text) do
      begin
       if (Edit_dsh.Text[k]=',') or (Edit_dsh.Text[k]=',')then
                begin
                 key:=#0;
                 beep ;
                end;
      end;
    end;
  end else key:=#0;
end;
//16)b_shH_nn
procedure TForm1.Edit_b_shH_nnKeyPress(Sender: TObject; var Key: Char);
Var k:integer;
begin
//������ ����� �� �������
if  Key='.' then
Key:=',';
//������� ������� [Esc]
if  Key=Chr(VK_ESCAPE) then
Key:=Chr( VK_BACK );
//���� ����
if Edit_b_shH_nn.Text='0' then          //���� ���������, �� ����� ����
 if not(key in [',','.',#8]) then   //������ ������ �������
       begin                        //���� ������ ������� ���
        key:=#0;
        beep ;
       end;
if key in['0'..'9',',','.',#8] then //��������� ������� ������ �����
  begin
  if (key=',') or (key='.') then //�������� ��� ������ ����� �������
    begin
    if Edit_b_shH_nn.Text='' then key:=#0;
      For k:=1 to Length(Edit_b_shH_nn.Text) do
      begin
       if (Edit_b_shH_nn.Text[k]=',') or (Edit_b_shH_nn.Text[k]=',')then
                begin
                 key:=#0;
                 beep ;
                end;
      end;
    end;
  end else key:=#0;
end;
//17)S_shH_ak
procedure TForm1.Edit_S_shH_akKeyPress(Sender: TObject; var Key: Char);
Var k:integer;
begin
//������ ����� �� �������
if  Key='.' then
Key:=',';
//������� ������� [Esc]
if  Key=Chr(VK_ESCAPE) then
Key:=Chr( VK_BACK );
//���� ����
if Edit_S_shH_ak.Text='0' then          //���� ���������, �� ����� ����
 if not(key in [',','.',#8]) then   //������ ������ �������
       begin                        //���� ������ ������� ���
        key:=#0;
        beep ;
       end;
if key in['0'..'9',',','.',#8] then //��������� ������� ������ �����
  begin
  if (key=',') or (key='.') then //�������� ��� ������ ����� �������
    begin
    if Edit_S_shH_ak.Text='' then key:=#0;
      For k:=1 to Length(Edit_S_shH_ak.Text) do
      begin
       if (Edit_S_shH_ak.Text[k]=',') or (Edit_S_shH_ak.Text[k]=',')then
                begin
                 key:=#0;
                 beep ;
                end;
      end;
    end;
  end else key:=#0;
end;
//18)d_shH_ak
procedure TForm1.Edit_d_shH_akKeyPress(Sender: TObject; var Key: Char);
Var k:integer;
begin
//������ ����� �� �������
if  Key='.' then
Key:=',';
//������� ������� [Esc]
if  Key=Chr(VK_ESCAPE) then
Key:=Chr( VK_BACK );
//���� ����
if Edit_d_shH_ak.Text='0' then          //���� ���������, �� ����� ����
 if not(key in [',','.',#8]) then   //������ ������ �������
       begin                        //���� ������ ������� ���
        key:=#0;
        beep ;
       end;
if key in['0'..'9',',','.',#8] then //��������� ������� ������ �����
  begin
  if (key=',') or (key='.') then //�������� ��� ������ ����� �������
    begin
    if Edit_d_shH_ak.Text='' then key:=#0;
      For k:=1 to Length(Edit_d_shH_ak.Text) do
      begin
       if (Edit_d_shH_ak.Text[k]=',') or (Edit_d_shH_ak.Text[k]=',')then
                begin
                 key:=#0;
                 beep ;
                end;
      end;
    end;
  end else key:=#0;
end;
//18)b_shH_ak
procedure TForm1.Edit_b_shH_akKeyPress(Sender: TObject; var Key: Char);
Var k:integer;
begin
//������ ����� �� �������
if  Key='.' then
Key:=',';
//������� ������� [Esc]
if  Key=Chr(VK_ESCAPE) then
Key:=Chr( VK_BACK );
//���� ����
if Edit_b_shH_ak.Text='0' then          //���� ���������, �� ����� ����
 if not(key in [',','.',#8]) then   //������ ������ �������
       begin                        //���� ������ ������� ���
        key:=#0;
        beep ;
       end;
if key in['0'..'9',',','.',#8] then //��������� ������� ������ �����
  begin
  if (key=',') or (key='.') then //�������� ��� ������ ����� �������
    begin
    if Edit_b_shH_ak.Text='' then key:=#0;
      For k:=1 to Length(Edit_b_shH_ak.Text) do
      begin
       if (Edit_b_shH_ak.Text[k]=',') or (Edit_b_shH_ak.Text[k]=',')then
                begin
                 key:=#0;
                 beep ;
                end;
      end;
    end;
  end else key:=#0;
end;



///  ������� �� ��������  Edit

procedure TForm1.Edit_DpacKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
                                     //1 �� 2
begin
if  Key=Word( VK_RETURN) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
 Edit_Pne .setfocus  ;
 end;
if  Key=Word( VK_DOWN) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_Pne .setfocus  ;
 end;
if  Key=Word(VK_NEXT) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_Pne .setfocus  ;
 end;
if  ComboBox2.ItemIndex=0  then
  begin
         if  Key=Word( VK_UP) then
            begin
             /// MessageBeep(MB_ICONASTERISK) ;
               Edit_b_shH_ak .setfocus  ;
           end;
         if  Key=Word( VK_PRIOR) then
            begin
             /// MessageBeep(MB_ICONASTERISK) ;
               Edit_b_shH_ak .setfocus  ;
           end;
         if  Key=Word(VK_CONTROL) then
            begin
             /// MessageBeep(MB_ICONASTERISK) ;
               Edit_b_shH_ak .setfocus  ;
           end;
  end
else
    begin
         if  Key=Word( VK_UP) then
                begin
                    /// MessageBeep(MB_ICONASTERISK) ;
                     ComboBox2 .setfocus  ;
              end;
         if  Key=Word( VK_PRIOR) then
                begin
                    /// MessageBeep(MB_ICONASTERISK) ;
                     ComboBox2 .setfocus  ;
                end;
         if  Key=Word(VK_CONTROL) then
                begin
                    /// MessageBeep(MB_ICONASTERISK) ;
                     ComboBox2 .setfocus  ;
                end;
    end;
end; 
procedure TForm1.Edit_PneKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
                                     //2 �� 3
begin
if  Key=Word( VK_RETURN) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
 Edit_tne .setfocus  ;
 end;
if  Key=Word(VK_CONTROL) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
 Edit_Dpac .setfocus  ;
 end;
if  Key=Word( VK_DOWN) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_tne .setfocus  ;
 end;
if  Key=Word( VK_NEXT) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_tne .setfocus  ;
 end;
if  Key=Word( VK_UP) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_Dpac .setfocus  ;
 end;
if  Key=Word( VK_PRIOR) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_Dpac .setfocus  ;
 end;
end;
procedure TForm1.Edit_tneKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
                                     //3 �� 4
begin
if  Key=Word( VK_RETURN) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
 Edit_tnb .setfocus  ;
 end;
if  Key=Word(VK_CONTROL) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
 Edit_Pne .setfocus  ;
 end;
if  Key=Word( VK_DOWN) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_tnb .setfocus  ;
 end;
if  Key=Word( VK_NEXT) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_tnb .setfocus  ;
 end;
if  Key=Word( VK_UP) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_Pne .setfocus  ;
 end;
if  Key=Word( VK_PRIOR) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_Pne .setfocus  ;
 end;
end;
procedure TForm1.Edit_tnbKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
                                     //4 �� 5
begin
if  Key=Word( VK_RETURN) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
 Edit_t_Bx .setfocus  ;
 end;
if  Key=Word(VK_CONTROL) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
 Edit_tne .setfocus  ;
 end;
if  Key=Word( VK_DOWN) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_t_Bx .setfocus  ;
 end;
if  Key=Word( VK_NEXT) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_t_Bx .setfocus  ;
 end;
if  Key=Word( VK_UP) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_tne .setfocus  ;
 end;
if  Key=Word( VK_PRIOR) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_tne .setfocus  ;
 end;
end;
procedure TForm1.Edit_t_BxKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
                                     //5 �� 6
begin
if  Key=Word( VK_RETURN) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
 Edit_P1 .setfocus  ;
 end;
if  Key=Word(VK_CONTROL) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
 Edit_tnb .setfocus  ;
 end;
if  Key=Word( VK_DOWN) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_P1 .setfocus  ;
 end;
if  Key=Word( VK_NEXT) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_P1 .setfocus  ;
 end;
if  Key=Word( VK_UP) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_tnb .setfocus  ;
 end;
if  Key=Word( VK_PRIOR) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_tnb .setfocus  ;
 end;
end;
procedure TForm1.Edit_P1KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
                                     //6 �� 7
begin
if  Key=Word( VK_RETURN) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
 Edit_t_Bix .setfocus  ;
 end;
if  Key=Word(VK_CONTROL) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
 Edit_t_Bx .setfocus  ;
 end;
if  Key=Word( VK_DOWN) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_t_Bix .setfocus  ;
 end;
if  Key=Word( VK_NEXT) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_t_Bix .setfocus  ;
 end;
if  Key=Word( VK_UP) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_t_Bx .setfocus  ;
 end;
if  Key=Word( VK_PRIOR) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_t_Bx .setfocus  ;
 end;
end;
procedure TForm1.Edit_t_BixKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
                                     //7 �� 8
begin
if  Key=Word( VK_RETURN) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
 ComboBox1 .setfocus  ;
 end;
if  Key=Word(VK_CONTROL) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
 Edit_P1 .setfocus  ;
 end;
if  Key=Word( VK_DOWN) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
ComboBox1 .setfocus  ;
 end;
if  Key=Word( VK_NEXT) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
ComboBox1 .setfocus  ;
 end;
if  Key=Word( VK_UP) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_P1 .setfocus  ;
 end;
if  Key=Word( VK_PRIOR) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_P1 .setfocus  ;
 end;
end;
procedure TForm1.ComboBox1KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
                                     //8 �� 9
begin
if  Key=Word( VK_RETURN) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
 Edit_d_H .setfocus  ;
 end;
if  Key=Word(VK_CONTROL) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
 Edit_t_Bix .setfocus  ;
 end;
end;
procedure TForm1.Edit_d_HKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
                                     //9 �� 10
begin
if  Key=Word( VK_RETURN) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
 Edit_btp .setfocus  ;
 end;
if  Key=Word(VK_CONTROL) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
 ComboBox1 .setfocus  ;
 end;
if  Key=Word( VK_DOWN) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_btp .setfocus  ;
 end;
if  Key=Word( VK_NEXT) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_btp .setfocus  ;
 end;
if  Key=Word( VK_UP) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
ComboBox1 .setfocus  ;
 end;
if  Key=Word( VK_PRIOR) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
ComboBox1 .setfocus  ;
 end;
end;
procedure TForm1.Edit_btpKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
                                     //10 �� 11
begin
if  Key=Word( VK_RETURN) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
 Edit_Stp .setfocus  ;
 end;
if  Key=Word(VK_CONTROL) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
 Edit_d_H .setfocus  ;
 end;
if  Key=Word( VK_DOWN) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_Stp .setfocus  ;
 end;
if  Key=Word( VK_NEXT) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_Stp .setfocus  ;
 end;
if  Key=Word( VK_UP) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_d_H .setfocus  ;
 end;
if  Key=Word( VK_PRIOR) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_d_H .setfocus  ;
 end;
end;
procedure TForm1.Edit_StpKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
                                     //11 �� 12
begin
if  Key=Word( VK_RETURN) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
 Edit_ntp .setfocus  ;
 end;
if  Key=Word(VK_CONTROL) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
 Edit_btp .setfocus  ;
 end;
if  Key=Word( VK_DOWN) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_ntp .setfocus  ;
 end;
if  Key=Word( VK_NEXT) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_ntp .setfocus  ;
 end;
if  Key=Word( VK_UP) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_btp .setfocus  ;
 end;
if  Key=Word( VK_PRIOR) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_btp .setfocus  ;
 end;
end;
procedure TForm1.Edit_ntpKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
                                      //12 �� 13
begin
if  Key=Word( VK_RETURN) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
 ComboBox2 .setfocus  ;
 end;
if  Key=Word(VK_CONTROL) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
 Edit_Stp .setfocus  ;
 end;
if  Key=Word( VK_DOWN) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
ComboBox2 .setfocus  ;
 end;
if  Key=Word( VK_NEXT) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
ComboBox2 .setfocus  ;
 end;
if  Key=Word( VK_UP) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_Stp .setfocus  ;
 end;
if  Key=Word( VK_PRIOR) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_Stp .setfocus  ;
 end;
end;
procedure TForm1.ComboBox2KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
                                    //13 �� 14
begin
if  ComboBox2.ItemIndex=0  then
  begin
         if  Key=Word( VK_RETURN) then
                begin
                    /// MessageBeep(MB_ICONASTERISK) ;
                     Edit_Ssh .setfocus  ;
              end;
         if  Key=Word( VK_CONTROL) then
                begin
                    /// MessageBeep(MB_ICONASTERISK) ;
                     Edit_ntp .setfocus  ;
              end;
  end
else
    begin
         if  Key=Word( VK_RETURN) then
                begin
                    /// MessageBeep(MB_ICONASTERISK) ;
                     Edit_Dpac .setfocus  ;
              end;
         if  Key=Word( VK_CONTROL) then
              begin
                  /// MessageBeep(MB_ICONASTERISK) ;
                 Edit_ntp .setfocus  ;
              end;
    end;
end;
procedure TForm1.Edit_SshKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
                                    //14 �� 15
begin
if  Key=Word( VK_RETURN) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
 Edit_dsh .setfocus  ;
 end;
if  Key=Word(VK_CONTROL) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
 ComboBox2 .setfocus  ;
 end;
if  Key=Word( VK_DOWN) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_dsh .setfocus  ;
 end;
if  Key=Word( VK_NEXT) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_dsh .setfocus  ;
 end;
if  Key=Word( VK_UP) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
ComboBox2 .setfocus  ;
 end;
if  Key=Word( VK_PRIOR) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
ComboBox2 .setfocus  ;
 end;
end;
procedure TForm1.Edit_dshKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
                                 //15 �� 16
begin
if  Key=Word( VK_RETURN) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
 Edit_b_shH_nn .setfocus  ;
 end;
if  Key=Word(VK_CONTROL) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
 Edit_Ssh .setfocus  ;
 end;
if  Key=Word( VK_DOWN) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_b_shH_nn .setfocus  ;
 end;
if  Key=Word( VK_NEXT) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_b_shH_nn .setfocus  ;
 end;
if  Key=Word( VK_UP) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_Ssh .setfocus  ;
 end;
if  Key=Word( VK_PRIOR) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_Ssh .setfocus  ;
 end;
end;
procedure TForm1.Edit_b_shH_nnKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
                                  //16 �� 17
begin
if  Key=Word( VK_RETURN) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
 Edit_S_shH_ak .setfocus  ;
 end;
if  Key=Word(VK_CONTROL) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
 Edit_dsh .setfocus  ;
 end;
if  Key=Word( VK_DOWN) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_S_shH_ak .setfocus  ;
 end;
if  Key=Word( VK_NEXT) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_S_shH_ak .setfocus  ;
 end;
if  Key=Word( VK_UP) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_dsh .setfocus  ;
 end;
if  Key=Word( VK_PRIOR) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_dsh .setfocus  ;
 end;
end;
procedure TForm1.Edit_S_shH_akKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
                                 //17 �� 18
begin
if  Key=Word( VK_RETURN) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
 Edit_d_shH_ak .setfocus  ;
 end;
if  Key=Word(VK_CONTROL) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
 Edit_b_shH_nn .setfocus  ;
 end;
if  Key=Word( VK_DOWN) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_d_shH_ak .setfocus  ;
 end;
if  Key=Word( VK_NEXT) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_d_shH_ak .setfocus  ;
 end;
if  Key=Word( VK_UP) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_b_shH_nn .setfocus  ;
 end;
if  Key=Word( VK_PRIOR) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_b_shH_nn .setfocus  ;
 end;
end;
procedure TForm1.Edit_d_shH_akKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
                                 //18 �� 19
begin
if  Key=Word( VK_RETURN) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
 Edit_b_shH_ak .setfocus  ;
 end;
if  Key=Word(VK_CONTROL) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
 Edit_S_shH_ak .setfocus  ;
 end;
if  Key=Word( VK_DOWN) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_b_shH_ak .setfocus  ;
 end;
if  Key=Word( VK_NEXT) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_b_shH_ak .setfocus  ;
 end;
if  Key=Word( VK_UP) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_S_shH_ak .setfocus  ;
 end;
if  Key=Word( VK_PRIOR) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_S_shH_ak .setfocus  ;
 end;
end;
procedure TForm1.Edit_b_shH_akKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
                              //19 �� 1
begin
if  Key=Word( VK_RETURN) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
 Edit_Dpac .setfocus  ;
 end;
if  Key=Word(VK_CONTROL) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
 Edit_d_shH_ak .setfocus  ;
 end;
if  Key=Word( VK_DOWN) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_Dpac .setfocus  ;
 end;
if  Key=Word( VK_NEXT) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_Dpac .setfocus  ;
 end;
if  Key=Word( VK_UP) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_d_shH_ak .setfocus  ;
 end;
if  Key=Word( VK_PRIOR) then
 begin
  /// MessageBeep(MB_ICONASTERISK) ;
Edit_d_shH_ak .setfocus  ;
 end;
end;


/// �������� ���������

procedure TForm1.FormCreate(Sender: TObject);
var
  Height, Width: integer;
  HH, WW: real;
begin
// ��������� ���������� �����
if ComboBox2.ItemIndex<>0 then
       begin
      Edit_Ssh.Enabled:=False ;
      Edit_dsh.Enabled:=False ;
      Edit_b_shH_nn.Enabled:=False ;
      Edit_S_shH_ak.Enabled:=False ;
      Edit_d_shH_ak.Enabled:=False ;
      Edit_b_shH_ak.Enabled:=False ;
       end
    else
       begin
      Edit_Ssh.Enabled:=true ;
      Edit_dsh.Enabled:=true ;
      Edit_b_shH_nn.Enabled:=true ;
      Edit_S_shH_ak.Enabled:=true ;
      Edit_d_shH_ak.Enabled:=true ;
      Edit_b_shH_ak.Enabled:=true ;
       end;
  // ������ ������
   if CheckBox2.Checked=False  then
      begin
        CheckBox1.Enabled:=false  ;
        Label4.Enabled:=false  ;
      end;
   if CheckBox2.Checked=true  then
       begin
         CheckBox1.Enabled:=true;
         Label4.Enabled:=true  ;
       end;
   //  ������  >  ������  <
  CheckBox4.Checked:=False ;
  Label2.Caption:='��������� >' ;
  Label10.Caption:='' ;
  Label11.Caption:='' ;
  // ������
  Timer1.Enabled:= False ;
  Timer2.Enabled:= False ;
  Timer3.Enabled:= False ;
  // ������ � ���� � � �����
///  form1.Left:=108;
///  form1.Top:=112;
  //���� �������
  Label1.Font.Color:=clWhite;
  Label2.Font.Color:=clWhite;
  Label7.Font.Color:=clWhite;
  Label3.Font.Color:=clBlack;
  Label4.Font.Color:=clBlack;
  //����� ���������
          CheckBox8.Enabled:= true ;
          Label7.Enabled:= true    ;
//�������������� �����                  (!!!!!!)                (!!!!!!)
ScaleForm(form1) ;
//������� �� ���� � �������� ������� [[������������ ������]]
PageControl1.ActivePage:=TabSheet1 ;
//��������� ������
vreme:=0 ;
Knopka_P:=0 ;
end;


/// ��������� ���� ���������
  // ��������
       // �������
procedure TForm1.FormShow(Sender: TObject);
begin
Day ; ///����
PageControl1.ActivePage:=TabSheet2 ; //���������
Timer5.Enabled:= True ;              //���������

//�������������� �����           (!!!!!!)           (!!!!!!)
   if screen.width=1280 then            //[1280]
     begin
       //���������� ������
        h_Close:=670 ;
        h_Open:=798 ;
        h_Women:=1136 ;
       //���������
       if screen.Height<1024 then
          begin
           VertScrollBar.Range:=PanelOcnovn.Height-47;  {������������� �������� ������������ ���������.}
           VertScrollBar.Visible:=True;                 {���������� ������������ ������ ���������.}
          end;
       //����� ���������
       Form1.Width:=h_Close ;

     end;
   if screen.width=1152 then            //[1152]
     begin
     //���������� ������
       h_Close:=598 ;
       h_Open:=718 ;
       h_Women:=1016 ;
     //���������
       if screen.Height<864 then
          begin
           VertScrollBar.Range:=PanelOcnovn.Height-46;  {������������� �������� ������������ ���������.}
           VertScrollBar.Visible:=True;                 {���������� ������������ ������ ���������.}
          end;
       //����� ���������
       Form1.Width:=h_Close ;
       Image1.Picture:=Form7.Image1152.Picture ;
       PageControl1.Top:=PageControl1.Top-2 ;
       Form1.Height:=679;
       ComboBox1.Left:=ComboBox1.Left+1 ;
       ComboBox2.Left:=ComboBox1.Left+1 ;
       Panel3.Height:=Panel3.Height+1;
     end;
   if screen.width=1024 then            //[1024]
     begin
     //���������� ������
       h_Close:=534 ;
       h_Open:=638 ;
       h_Women:=911 ;
     //���������
       if screen.Height<768 then
          begin
           VertScrollBar.Range:=PanelOcnovn.Height-41;  {������������� �������� ������������ ���������.}
           VertScrollBar.Visible:=True;                 {���������� ������������ ������ ���������.}
          end;
       //����� ���������
       Form1.Width:=h_Close ;
       Image1.Picture:=Form7.Image1024.Picture ;
       PageControl1.Top:=PageControl1.Top-4 ;
       Form1.Height:=607;
       ComboBox1.Left:=ComboBox1.Left+1 ;
       ComboBox2.Left:=ComboBox1.Left+1 ;
       Panel3.Height:=Panel3.Height+3;
       Panel3.Width:=Panel3.Width+5;
       Panel5.Height:=Panel3.Height+1;
       Panel5.Width:=Panel3.Width+10;
       Panel2.Height:=Panel3.Height+1;
       Label3.Width:=Label7.Width+5;
       Label7.Width:=Label7.Width+5;
     end;
   if screen.width=800 then            //[800]
     begin
     //���������
     //���������� ������
       h_Close:=414 ;
       h_Open:=486 ;
       h_Women:=701 ;
     //���������
       if screen.Height<600 then
          begin
           VertScrollBar.Range:=PanelOcnovn.Height-31;  {������������� �������� ������������ ���������.}
           VertScrollBar.Visible:=True;                 {���������� ������������ ������ ���������.}
          end;
       //����� ���������
       Form1.Width:=h_Close ;
       Image1.Picture:=Form7.Image800.Picture ;
       PageControl1.Top:=PageControl1.Top-9 ;
       PageControl1.Left:=PageControl1.Left-1;
       PageControl1.Height:=480;
       Form1.Height:=483;
       ComboBox1.Style:=csDropDownList;
       ComboBox2.Style:=csDropDownList;
       ComboBox1.Font.Height:=ComboBox1.Height-12 ;
       ComboBox2.Font.Height:=ComboBox1.Height-12 ;
       ComboBox1.Left:=ComboBox1.Left+1 ;
       ComboBox2.Left:=ComboBox1.Left+1 ;
       Panel3.Height:=Panel3.Height+3;
       Panel3.Width:=Panel3.Width+5;
       Panel5.Height:=Panel3.Height+1;
       Panel5.Width:=Panel3.Width+10;
       Panel2.Height:=Panel3.Height+1;
       Label3.Width:=Label7.Width+5;
       Label7.Width:=Label7.Width+5;
     end;
end;

//������ ��������� �����
//����
procedure TForm1.FormMouseWheelDown(Sender: TObject; Shift: TShiftState;
  MousePos: TPoint; var Handled: Boolean);
begin
 VertScrollBar.Position:=VertScrollBar.Position+10;
end;
//�����
procedure TForm1.FormMouseWheelUp(Sender: TObject; Shift: TShiftState;
  MousePos: TPoint; var Handled: Boolean);
begin
 VertScrollBar.Position:=VertScrollBar.Position-10;
end;


//��������� ��������
procedure TForm1.Image7Click(Sender: TObject);

begin
if Knopka_P=0 then   //(������)
  begin
    Knopka_P:=1 ;
    Timer5.Enabled:= False ;
    PageControl1.ActivePage:=TabSheet1 ;
    Save_and_Load; // ��������� ��������� (��� ������ ���������)
       if interfeis='�������' then  Full.Checked:=true    //  Full ��� Lite ������
          else Lite.Checked:=true ;
    Exit;
  end;
if Knopka_P=1 then    //(���������)
  begin
    Knopka_P:=2;
    Form1.Width:=h_Close ;
    PageControl1.ActivePage:=TabSheet2 ;
    Exit;
  end;
if Knopka_P=2 then    //(������� ����)
  begin
    if CheckBox4.Checked=False  then  // ������  >
     begin
       Knopka_P:=1;
       PageControl1.ActivePage:=TabSheet1 ;
       Exit;
     end;
    if CheckBox4.Checked=True  then  //  ������  <
     begin
       Knopka_P:=1;
       Form1.Width:=h_Open ;
       PageControl1.ActivePage:=TabSheet1 ;
       Exit;
     end;
  end;
end;

//��������� ������
procedure TForm1.Timer5Timer(Sender: TObject);

begin
  vreme:=vreme+1 ;
  if vreme=8 then
    begin
      Timer5.Enabled:= False ;
      PageControl1.ActivePage:=TabSheet1 ;
      Save_and_Load; // ��������� ��������� (��� ������ ���������)
          if interfeis='�������' then  Full.Checked:=true    //  Full ��� Lite ������
             else Lite.Checked:=true ;
      Knopka_P:=1;
    end;
end;


// ����� ������

procedure TForm1.CheckBox2Click(Sender: TObject);
begin
   if (CheckBox2.Checked=False) or (CheckBox3.Checked=False) then
      begin
        CheckBox1.Enabled:=False  ;
        Label4.Enabled:=False  ;
      end;
   if (CheckBox2.Checked=True) and (CheckBox3.Checked=True) then
      begin
         CheckBox1.Enabled:=True;
         Label4.Enabled:=True  ;
       end;
end;

// ����� �������

procedure TForm1.CheckBox3Cick(Sender: TObject);
begin
if CheckBox3.Checked=False  then
  begin
     CheckBox1.Checked:=False;
     CheckBox1.Enabled:=False;
     Label4.Enabled:=False  ;
  end;
if CheckBox3.Checked=true  then
  begin
     CheckBox1.Checked:=true;
     CheckBox1.Enabled:=true;
     Label4.Enabled:=True  ;
  end;
if (CheckBox3.Checked=true)  and (CheckBox2.Checked=False) then
   begin
     CheckBox1.Checked:=False;
     CheckBox1.Enabled:=False;
     Label4.Enabled:=False  ;
   end;
if (CheckBox2.Checked=False) or (CheckBox3.Checked=False) then
   begin
     CheckBox1.Enabled:=False  ;
     Label4.Enabled:=False  ;
   end;
end;


// ������������ ������ (�� ������)
//1) ������
procedure TForm1.Label1Click(Sender: TObject);
begin
if CheckBox2.Checked=true then
   CheckBox2.Checked:=False
  else CheckBox2.Checked:=true ;
end;
//2)������  >   ������  <
procedure TForm1.Label2Click(Sender: TObject);
begin
if CheckBox4.Checked=true then
   CheckBox4.Checked:=False
  else CheckBox4.Checked:=true ;
end;
//3)�������
procedure TForm1.Label3Click(Sender: TObject);
begin
if CheckBox3.Checked=true then
   CheckBox3.Checked:=False
  else CheckBox3.Checked:=true ;
end;
//4)������������
procedure TForm1.Label4Click(Sender: TObject);
begin
if CheckBox1.Checked=true then
   CheckBox1.Checked:=False
  else CheckBox1.Checked:=true ;
end;
//5) ���������
procedure TForm1.Label7Click(Sender: TObject);
begin
if CheckBox8.Checked=true then
   CheckBox8.Checked:=False
  else CheckBox8.Checked:=true ;
end;

// �����   (('������  >'    '������  <'))
procedure TForm1.CheckBox4Click(Sender: TObject);
begin
Timer2.Enabled:= False ;
Timer3.Enabled:= False ;
if Form1.Width<=(h_Open+15) then              //�������
     if CheckBox4.Checked=False  then
      begin
       CheckBox5.Checked:=False  ;
       Label2.Caption:='��������� >' ;
       Label10.Caption:='' ;
       Label11.Caption:='' ;
          if interfeis='�������' then //�������� ��������� ����������
           begin
              Timer1.Interval:=1  ;
              Timer1.Enabled:= true ;
           end
         else Form1.Width:=h_Close ;
      end;
    if CheckBox4.Checked=true  then           //�������
      begin
       CheckBox5.Checked:=true  ;
       Label2.Caption:='������';
       Label10.Caption:='��������' ;
       Label11.Caption:='<' ;
          if interfeis='�������' then  //�������� ��������� ����������
           begin
              Timer1.Interval:=1   ;
              Timer1.Enabled:= true ;
           end
         else Form1.Width:=h_Open ;
      end;
   if Form1.Width>=(h_Women-15)  then           //�������� �� �����
     if CheckBox4.Checked=False  then
      begin
        CheckBox5.Checked:=true  ;
        Label2.Caption:='��������� >' ;
        Label10.Caption:='' ;
        Label11.Caption:='' ;
          if interfeis='�������' then   //�������� ��������� ����������
           begin
              Timer4.Interval:=1  ;
              Timer4.Enabled:= true ;
           end
         else Form1.Width:=h_Close ;
      end;
end;

// ������ ���������� ������
procedure TForm1.Timer1Timer(Sender: TObject);
begin
if PageControl1.ActivePage=TabSheet1 then    //���� ������� ���� �������� ������
  begin
   if  CheckBox4.Checked=true then
    begin
      Form1.Width:=Form1.Width+8 ;
      Timer1.Interval:= Timer1.Interval + 1  ;
      if (Form1.Width>=h_Open) and  (CheckBox4.Checked=true) then
        begin
          Timer1.Enabled:= False  ;
          Timer1.Interval:=1 ;
               if Knopka_P=2 then  //����� ��������� ������� ���� ���������
                 begin
                  Form1.Width:=h_Close ;
                  PageControl1.ActivePage:=TabSheet2 ;Exit;
                 end
               else Exit;
        end;
    end;
   if  CheckBox4.Checked=False then
     begin
      Form1.Width:=Form1.Width-8 ;
      Timer1.Interval:= Timer1.Interval + 1  ;
        if (Form1.Width<=h_Close) and  (CheckBox4.Checked=False) then
          begin
            Timer1.Enabled:= False  ;
            Timer1.Interval:=1 ;
               if Knopka_P=2 then  //����� ��������� �������� ���� ���������
                 begin
                  Form1.Width:=h_Close ;
                  PageControl1.ActivePage:=TabSheet2 ;
                 end
               else Exit;
          end;
     end;
  end;   //���� ������� ���� �������� ������
end;
// ������ �������� �� �����
procedure TForm1.Timer4Timer(Sender: TObject);
begin
  if Form1.Width>=(h_Close+1) then
  begin
   Form1.Width:=Form1.Width-15 ;
   Timer4.Interval:= Timer4.Interval + 1 ;
   if Form1.Width<=(h_Close+1)  then
     begin
      Timer4.Enabled:= False  ;
      Timer4.Interval:=1   ;
         if Knopka_P=2 then  //����� ��������� ������� ���� ���������
            begin
              Form1.Width:=h_Close ;
              PageControl1.ActivePage:=TabSheet2 ;
            end;
     end;
  end;
end;

// ������ ������� (���� ����)
procedure TForm1.Image2Click(Sender: TObject);
begin
// ������������ ������
if  CheckBox5.Checked=true  then
  CheckBox5.Checked:=False
  else
   begin
    if  CheckBox5.Checked=False then
      CheckBox5.Checked:=true ;
   end;
 if CheckBox5.Checked=False then
  begin
   if interfeis='�������' then    //�������� ��������� ����������
      begin
        Timer3.Enabled:= False  ;
        Timer2.Enabled:= true ;
        Timer2.Interval:=1  ;
      end
      else Form1.Width:=h_Women ;
  end;
 if CheckBox5.Checked=true then
  begin
   if interfeis='�������' then    //�������� ��������� ����������
      begin
        Timer2.Enabled:= False ;
        Timer3.Enabled:= true ;
        Timer3.Interval:=1  ;
      end
      else Form1.Width:=h_Open ;
  end;
end;

// ������ ������  (���������)
procedure TForm1.Timer2Timer(Sender: TObject);
begin
  if Form1.Width<=(h_Women-15) then
  begin
   Form1.Width:=Form1.Width+13 ;
   Timer2.Interval:= Timer2.Interval + 1 ;
   if Form1.Width>=h_Women  then
     begin
      Timer2.Enabled:= False  ;
      Timer2.Interval:=1   ;
       if Knopka_P=2 then  //����� ��������� ������� ���� ���������
          begin
            Form1.Width:=h_Close ;
            CheckBox5.Checked:=true;
            PageControl1.ActivePage:=TabSheet2 ;
          end;
     end;
  end;
end;

// ������ ������  (������)
procedure TForm1.Timer3Timer(Sender: TObject);
begin
  if Form1.Width>=h_Open then
  begin
   Form1.Width:=Form1.Width-13 ;
   Timer3.Interval:= Timer3.Interval + 1 ;
   if Form1.Width<=h_Open  then
     begin
      Timer3.Enabled:= False  ;
      Timer3.Interval:=1   ;
       if Knopka_P=2 then  //����� ��������� ������� ���� ���������
          begin
            Form1.Width:=h_Close ;
            PageControl1.ActivePage:=TabSheet2 ;
          end;
     end;
  end;
end;

//��������� Excel [�����]  (Excel)
procedure TForm1.Label5Click(Sender: TObject);
begin
    CheckBox6.Checked:=true;
end;
//��������� ������� [�����]  (�������)
procedure TForm1.Label6Click(Sender: TObject);
begin
     CheckBox7.Checked:=true;
     CheckBox10.Enabled:=true;
     CheckBox10.Checked:=False;
     CheckBox9.Checked:=true;
end;
//��������� Excel
procedure TForm1.CheckBox6Click(Sender: TObject);
begin
if CheckBox6.Checked=true then
    CheckBox7.Checked:=False  ;
if CheckBox6.Checked=False then
    CheckBox7.Checked:=true  ;
end;
//��������� �������
procedure TForm1.CheckBox7Click(Sender: TObject);
begin
  if CheckBox7.Checked=true  then
        begin
          CheckBox6.Checked:=False  ;
          CheckBox10.Enabled:=False;
          Label9.Enabled:= False   ;
          CheckBox9.Enabled:=False;
          Label8.Enabled:= False   ;
          CheckBox10.Checked:=False;
          CheckBox9.Checked:=true;
        end
    else
        begin
          CheckBox6.Checked:=true  ;
          CheckBox10.Enabled:=true;
          Label9.Enabled:= true   ;
          CheckBox9.Enabled:=true;
          Label8.Enabled:= true   ;
          CheckBox10.Checked:=False;
          CheckBox9.Checked:=true;
        end;
end;

//��������� �����������  [�����]  (����������� )
procedure TForm1.Label8Click(Sender: TObject);
begin
    CheckBox9.Checked:=true;
end;
//��������� ������ [�����]  (������)
procedure TForm1.Label9Click(Sender: TObject);
begin
     CheckBox10.Checked:=true;
end;
//��������� �����������
procedure TForm1.CheckBox9Click(Sender: TObject);
begin
if CheckBox9.Checked=true then
    CheckBox10.Checked:=False  ;
if CheckBox9.Checked=False then
    CheckBox10.Checked:=true  ;
end;
//��������� ������
procedure TForm1.CheckBox10Click(Sender: TObject);
begin
  if CheckBox10.Checked=true  then
          CheckBox9.Checked:=False
    else
          CheckBox9.Checked:=true  ;
end;


END.
 
