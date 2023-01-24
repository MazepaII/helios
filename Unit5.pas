//КООРДИНАТЫ СОЗДАТЕЛЯ ПРОГРАММЫ
//Электронная почта: Ltlvfpfqbpfzws@yandex.ru

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
  //Процидуры
function NachaloTekct: boolean;  //Вернуть текст      ([])
    { Private declarations }
Procedure CMDialogKey(Var Msg: TWMKey); message CM_DIALOGKEY;  //Запрет клавишы [Tab и Enter]

  public
    { Public declarations }
  end;

var
  Form5: TForm5;
  ///Таймер (текст)
  i_i: integer ;

implementation

uses Unit3, Unit1;

{$R *.dfm}

//Запрет клавишы [Tab и Enter]
procedure TForm5.CMDialogKey(var Msg: TWMKey);
begin
 Msg.Result := 0
end;

function TForm5.NachaloTekct: boolean;
begin
Form5.Memo1.Lines.Clear;
Form5.Memo1.Lines.Add('Программа для расчета ');
Form5.Memo1.Lines.Add('прямотрубного Парогенератора');
Form5.Memo1.Lines.Add('');
Form5.Memo1.Lines.Add('Version 1.12  build 36  (2011)');
Form5.Memo1.Lines.Add('');
Form5.Memo1.Lines.Add('Санкт-Петербургский государственый');
Form5.Memo1.Lines.Add('морской технический университет');
Form5.Memo1.Lines.Add('');
Form5.Memo1.Lines.Add('Автор: Мазилевский И.И.');
Form5.Memo1.Lines.Add('            Ревков М.В.');
end;

procedure TForm5.Button1Click(Sender: TObject);
begin
Form5.close;
NachaloTekct ; //Вернуть текст
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
TimerKoordinatbI.Enabled:=False ;  //ТАЙМЕР  Координаты создателя программы
Form5.Memo1.Enabled:= true  ;

end;

procedure TForm5.Image1Click(Sender: TObject);
begin
Form5.close;
Form3.ShowModal;
form1.caption:='Дед Мазай' ;
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
TimerKoordinatbI.Enabled:=False ;  //ТАЙМЕР  Координаты создателя программы
Form5.Memo1.Enabled:= true  ;
NachaloTekct ; //Вернуть текст                        ([])
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
     Form5.Memo1.Lines.Add('Автор: ДЕД МАЗАЙ И ЗАЯЦЫ');
     Form5.Memo1.Lines.Add('-------------------------------------------------------');
     Form5.Memo1.Lines.Add('Пора нажать на картинку');
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
/// ТАЙМЕРЫ

//1) Инструкция
procedure TForm5.Timer1Timer(Sender: TObject);
begin
     Form5.Memo1.Lines.Add('');
     Form5.Memo1.Lines.Add('Инструкция');
     Form5.Memo1.Lines.Add('');
     Timer1.Enabled:= False   ;
     Timer2.Enabled:= true   ;
     
end;
///2) (1) Берете мышку в правую руку)
procedure TForm5.Timer2Timer(Sender: TObject);
begin
     Form5.Memo1.Lines.Add('1) Берете мышку в правую руку');
     Timer2.Enabled:= False  ;
     Timer3.Enabled:= true   ;
end;
///3) (2) Наводите курсор на рисунок парогенератора)
procedure TForm5.Timer3Timer(Sender: TObject);
begin
     Form5.Memo1.Lines.Add('2) Наводите курсор на рисунок парогенератора');
     Timer3.Enabled:= False  ;
     Timer4.Enabled:= true   ;
end;
///4) ((он левее текста))
procedure TForm5.Timer4Timer(Sender: TObject);
begin
     Form5.Memo1.Lines.Add('   (он левее текста)');
     Timer4.Enabled:= False  ;
     Timer5.Enabled:= true   ;
end;
///5) (3) Нажимаете левею кнопку мыши)
procedure TForm5.Timer5Timer(Sender: TObject);
begin
     Form5.Memo1.Lines.Add('3) Нажимаете левею кнопку мыши');
     Timer5.Enabled:= False  ;
     Timer6.Enabled:= true   ;
end;
///6) (Ну, что тупишь.)
procedure TForm5.Timer6Timer(Sender: TObject);
begin
     Form5.Memo1.Lines.Add('   Ну, что тупишь.');
     Timer6.Enabled:= False  ;
     Timer7.Enabled:= true   ;
end;
///7) (Тыкай, кому говорю!)
procedure TForm5.Timer7Timer(Sender: TObject);
begin
     Form5.Memo1.Lines.Add('   Тыкай, кому говорю!');
     Timer7.Enabled:= False  ;
     Timer8.Enabled:= true   ;
end;
///8)  (Арбайтан, АРБАЙТОН!!!)
procedure TForm5.Timer8Timer(Sender: TObject);
begin
Form5.Memo1.Lines.Add('   Арбайтан, АРБАЙТОН!!!');
     Timer8.Enabled:= False  ;
     Timer9.Enabled:= true   ;
end;
///9)  (Все я умываю руки.)
procedure TForm5.Timer9Timer(Sender: TObject);
begin
     Form5.Memo1.Lines.Add('   Все я умываю руки.');
     Timer9.Enabled:= False  ;
     Timer10.Enabled:= true   ;
end;

/// 10 (Возабновление поля)
procedure TForm5.Timer10Timer(Sender: TObject);
begin
NachaloTekct ;            //Вернуть текст                 ([])
Form5.Memo1.Lines.Add('');
     Timer10.Enabled:= False  ;
     Timer11.Enabled:= true   ;
end;

/// 11)  (Как я от вас устал...)
procedure TForm5.Timer11Timer(Sender: TObject);
begin
Form5.Memo1.Lines.Add('Как я от вас устал...');
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
TimerKoordinatbI.Enabled:=False ;  //ТАЙМЕР  Координаты создателя программы
Form5.Memo1.Enabled:= true  ;
NachaloTekct ;               //Вернуть текст             ([])

end;

// ОТКРЫТИЕ  О программе
procedure TForm5.FormCreate(Sender: TObject);
//const
//CS_DROPSHADOW=$00020000;
begin
//Тень от формы
//SetClassLong(Handle,GCL_STYLE,GetWindowLong(Handle, GCL_STYLE) or CS_DROPSHADOW);
//Мемо1 только для чтения
Memo1.ReadOnly:= true ;
TimerKoordinatbI.Enabled:=true ;  //ТАЙМЕР  Координаты создателя программы
end;

//Нажатие кнопки "ОК"
procedure TForm5.Button1KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
if  Key=Word( VK_RETURN) then
 Button1.Click
end;

//ТАЙМЕР  Координаты создателя программы
procedure TForm5.TimerKoordinatbITimer(Sender: TObject);
begin
  Form5.Memo1.Lines.Clear;
  Form5.Memo1.Lines.Add('КООРДИНАТЫ СОЗДАТЕЛЯ ПРОГРАММЫ') ;
  Form5.Memo1.Lines.Add('') ;
  Form5.Memo1.Lines.Add('Ф.И.О.:') ;
  Form5.Memo1.Lines.Add('Мазилевский Илья Игоревич') ;
  Form5.Memo1.Lines.Add('') ;
  Form5.Memo1.Lines.Add('Мобильный телефон номер:') ;
  Form5.Memo1.Lines.Add('+7(921)434-75-16') ;
  Form5.Memo1.Lines.Add('') ;
  Form5.Memo1.Lines.Add('Электронная почта:') ;
  Form5.Memo1.Lines.Add('Ltlvfpfqbpfzws@yandex.ru') ;
  TimerKoordinatbI.Enabled:=False ;
end;

end.
