//���������� ��������� ���������
//����������� �����: Ltlvfpfqbpfzws@yandex.ru

unit Unit5;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, jpeg, ExtCtrls, StdCtrls, ComCtrls;

type
  TForm5 = class(TForm)
    Button1: TButton;
    Image1: TImage;
    Memo1: TMemo;
    Timer1: TTimer;
    Timer2: TTimer;
    Timer3: TTimer;
    Timer4: TTimer;
    Timer5: TTimer;
    Timer6: TTimer;
    Timer7: TTimer;
    Timer8: TTimer;
    Timer9: TTimer;
    Timer10: TTimer;
    Timer11: TTimer;
    Timer12: TTimer;
    TimerKoordinatbI: TTimer;
    procedure Button1Click(Sender: TObject);
    procedure Image1Click(Sender: TObject);
    procedure Memo1KeyPress(Sender: TObject; var Key: Char);
    procedure Timer1Timer(Sender: TObject);
    procedure Timer2Timer(Sender: TObject);
    procedure Timer3Timer(Sender: TObject);
    procedure Timer4Timer(Sender: TObject);
    procedure Timer6Timer(Sender: TObject);
    procedure Timer7Timer(Sender: TObject);
    procedure Timer8Timer(Sender: TObject);
    procedure Timer9Timer(Sender: TObject);
    procedure Timer5Timer(Sender: TObject);
    procedure Timer10Timer(Sender: TObject);
    procedure Timer11Timer(Sender: TObject);
    procedure Timer12Timer(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure Button1KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure TimerKoordinatbITimer(Sender: TObject);
  private
  //���������
function NachaloTekct: boolean;  //������� �����      ([])
    { Private declarations }
Procedure CMDialogKey(Var Msg: TWMKey); message CM_DIALOGKEY;  //������ ������� [Tab � Enter]

  public
    { Public declarations }
  end;

var
  Form5: TForm5;
  ///������ (�����)
  i_i: integer ;

implementation

uses Unit3, Unit1;

{$R *.dfm}

//������ ������� [Tab � Enter]
procedure TForm5.CMDialogKey(var Msg: TWMKey);
begin
 Msg.Result := 0
end;

function TForm5.NachaloTekct: boolean;
begin
Form5.Memo1.Lines.Clear;
Form5.Memo1.Lines.Add('��������� ��� ������� ');
Form5.Memo1.Lines.Add('������������� ��������������');
Form5.Memo1.Lines.Add('');
Form5.Memo1.Lines.Add('Version 1.12  build 36  (2011)');
Form5.Memo1.Lines.Add('');
Form5.Memo1.Lines.Add('�����-������������� ��������������');
Form5.Memo1.Lines.Add('������� ����������� �����������');
Form5.Memo1.Lines.Add('');
Form5.Memo1.Lines.Add('�����: ����������� �.�.');
Form5.Memo1.Lines.Add('            ������ �.�.');
end;

procedure TForm5.Button1Click(Sender: TObject);
begin
Form5.close;
NachaloTekct ; //������� �����
Timer1.Enabled:= False   ;
Timer2.Enabled:= False   ;
Timer3.Enabled:= False   ;
Timer4.Enabled:= False   ;
Timer5.Enabled:= False   ;
Timer6.Enabled:= False   ;
Timer7.Enabled:= False   ;
Timer8.Enabled:= False   ;
Timer9.Enabled:= False   ;
Timer10.Enabled:= False   ;
Timer11.Enabled:= False   ;
Timer12.Enabled:= False   ;
TimerKoordinatbI.Enabled:=False ;  //������  ���������� ��������� ���������
Form5.Memo1.Enabled:= true  ;

end;

procedure TForm5.Image1Click(Sender: TObject);
begin
Form5.close;
Form3.ShowModal;
form1.caption:='��� �����' ;
Timer1.Enabled:= False   ;
Timer2.Enabled:= False   ;
Timer3.Enabled:= False   ;
Timer4.Enabled:= False   ;
Timer5.Enabled:= False   ;
Timer6.Enabled:= False   ;
Timer7.Enabled:= False   ;
Timer8.Enabled:= False   ;
Timer9.Enabled:= False   ;
Timer10.Enabled:= False   ;
Timer11.Enabled:= False   ;
Timer12.Enabled:= False   ;
TimerKoordinatbI.Enabled:=False ;  //������  ���������� ��������� ���������
Form5.Memo1.Enabled:= true  ;
NachaloTekct ; //������� �����                        ([])
end;

procedure TForm5.Memo1KeyPress(Sender: TObject; var Key: Char);
begin
if not( Key in [Chr( VK_SPACE),Chr( VK_RETURN)]) then
     begin
      beep;
      Key:=#0;
     end
  else
    begin
     Form5.Memo1.Lines.Clear;
     Key:=#0;
     Form5.Memo1.Lines.Add('�����: ��� ����� � �����');
     Form5.Memo1.Lines.Add('-------------------------------------------------------');
     Form5.Memo1.Lines.Add('���� ������ �� ��������');
     Timer1.Enabled:= False   ;
     Timer2.Enabled:= False   ;
     Timer3.Enabled:= False   ;
     Timer4.Enabled:= False   ;
     Timer5.Enabled:= False   ;
     Timer6.Enabled:= False   ;
     Timer7.Enabled:= False   ;
     Timer8.Enabled:= False   ;
     Timer9.Enabled:= False   ;
     Timer10.Enabled:= False   ;
     Timer11.Enabled:= False   ;
     Timer12.Enabled:= False   ;
     Timer1.Enabled:= true  ;
    end;
end;
/// �������

//1) ����������
procedure TForm5.Timer1Timer(Sender: TObject);
begin
     Form5.Memo1.Lines.Add('');
     Form5.Memo1.Lines.Add('����������');
     Form5.Memo1.Lines.Add('');
     Timer1.Enabled:= False   ;
     Timer2.Enabled:= true   ;
     
end;
///2) (1) ������ ����� � ������ ����)
procedure TForm5.Timer2Timer(Sender: TObject);
begin
     Form5.Memo1.Lines.Add('1) ������ ����� � ������ ����');
     Timer2.Enabled:= False  ;
     Timer3.Enabled:= true   ;
end;
///3) (2) �������� ������ �� ������� ��������������)
procedure TForm5.Timer3Timer(Sender: TObject);
begin
     Form5.Memo1.Lines.Add('2) �������� ������ �� ������� ��������������');
     Timer3.Enabled:= False  ;
     Timer4.Enabled:= true   ;
end;
///4) ((�� ����� ������))
procedure TForm5.Timer4Timer(Sender: TObject);
begin
     Form5.Memo1.Lines.Add('   (�� ����� ������)');
     Timer4.Enabled:= False  ;
     Timer5.Enabled:= true   ;
end;
///5) (3) ��������� ����� ������ ����)
procedure TForm5.Timer5Timer(Sender: TObject);
begin
     Form5.Memo1.Lines.Add('3) ��������� ����� ������ ����');
     Timer5.Enabled:= False  ;
     Timer6.Enabled:= true   ;
end;
///6) (��, ��� ������.)
procedure TForm5.Timer6Timer(Sender: TObject);
begin
     Form5.Memo1.Lines.Add('   ��, ��� ������.');
     Timer6.Enabled:= False  ;
     Timer7.Enabled:= true   ;
end;
///7) (�����, ���� ������!)
procedure TForm5.Timer7Timer(Sender: TObject);
begin
     Form5.Memo1.Lines.Add('   �����, ���� ������!');
     Timer7.Enabled:= False  ;
     Timer8.Enabled:= true   ;
end;
///8)  (��������, ��������!!!)
procedure TForm5.Timer8Timer(Sender: TObject);
begin
Form5.Memo1.Lines.Add('   ��������, ��������!!!');
     Timer8.Enabled:= False  ;
     Timer9.Enabled:= true   ;
end;
///9)  (��� � ������ ����.)
procedure TForm5.Timer9Timer(Sender: TObject);
begin
     Form5.Memo1.Lines.Add('   ��� � ������ ����.');
     Timer9.Enabled:= False  ;
     Timer10.Enabled:= true   ;
end;

/// 10 (������������� ����)
procedure TForm5.Timer10Timer(Sender: TObject);
begin
NachaloTekct ;            //������� �����                 ([])
Form5.Memo1.Lines.Add('');
     Timer10.Enabled:= False  ;
     Timer11.Enabled:= true   ;
end;

/// 11)  (��� � �� ��� �����...)
procedure TForm5.Timer11Timer(Sender: TObject);
begin
Form5.Memo1.Lines.Add('��� � �� ��� �����...');
     Timer11.Enabled:= False  ;
     Timer12.Enabled:= true   ;
end;
// 12) (Enabled)
procedure TForm5.Timer12Timer(Sender: TObject);
begin
 Form5.Memo1.Enabled:= False;
     Timer11.Enabled:= False  ;
end;

procedure TForm5.FormClose(Sender: TObject; var Action: TCloseAction);
begin
Form5.close;
Timer1.Enabled:= False   ;
Timer2.Enabled:= False   ;
Timer3.Enabled:= False   ;
Timer4.Enabled:= False   ;
Timer5.Enabled:= False   ;
Timer6.Enabled:= False   ;
Timer7.Enabled:= False   ;
Timer8.Enabled:= False   ;
Timer9.Enabled:= False   ;
Timer10.Enabled:= False   ;
Timer11.Enabled:= False   ;
Timer12.Enabled:= False   ;
TimerKoordinatbI.Enabled:=False ;  //������  ���������� ��������� ���������
Form5.Memo1.Enabled:= true  ;
NachaloTekct ;               //������� �����             ([])

end;

// ��������  � ���������
procedure TForm5.FormCreate(Sender: TObject);
//const
//CS_DROPSHADOW=$00020000;
begin
//���� �� �����
//SetClassLong(Handle,GCL_STYLE,GetWindowLong(Handle, GCL_STYLE) or CS_DROPSHADOW);
//����1 ������ ��� ������
Memo1.ReadOnly:= true ;
TimerKoordinatbI.Enabled:=true ;  //������  ���������� ��������� ���������
end;

//������� ������ "��"
procedure TForm5.Button1KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
if  Key=Word( VK_RETURN) then
 Button1.Click
end;

//������  ���������� ��������� ���������
procedure TForm5.TimerKoordinatbITimer(Sender: TObject);
begin
  Form5.Memo1.Lines.Clear;
  Form5.Memo1.Lines.Add('���������� ��������� ���������') ;
  Form5.Memo1.Lines.Add('') ;
  Form5.Memo1.Lines.Add('�.�.�.:') ;
  Form5.Memo1.Lines.Add('����������� ���� ��������') ;
  Form5.Memo1.Lines.Add('') ;
  Form5.Memo1.Lines.Add('��������� ������� �����:') ;
  Form5.Memo1.Lines.Add('+7(921)434-75-16') ;
  Form5.Memo1.Lines.Add('') ;
  Form5.Memo1.Lines.Add('����������� �����:') ;
  Form5.Memo1.Lines.Add('Ltlvfpfqbpfzws@yandex.ru') ;
  TimerKoordinatbI.Enabled:=False ;
end;

end.
