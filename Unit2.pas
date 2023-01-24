//КООРДИНАТЫ СОЗДАТЕЛЯ ПРОГРАММЫ
//Электронная почта: Ltlvfpfqbpfzws@yandex.ru

unit Unit2;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, jpeg, ExtCtrls,ComObj, Menus, Printers, scale ;

type
  TForm2 = class(TForm)
    Memo1: TMemo;
    MainMenu1: TMainMenu;
    A1: TMenuItem;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    SaveDialog1: TSaveDialog;
    N5: TMenuItem;
    PrintDialog1: TPrintDialog;
    N6: TMenuItem;
    procedure FormResize(Sender: TObject);
    procedure N1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure N3Click(Sender: TObject);
    procedure N4Click(Sender: TObject);
    procedure N5Click(Sender: TObject);
    procedure N6Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form2: TForm2;

implementation

uses Unit5;


{$R *.dfm}

// Размер мемо
procedure TForm2.FormResize(Sender: TObject);
begin
Memo1.Width:=Form2.Width - 16 ;
Memo1.Height:=Form2.Height - 58 ;
end;

/// ОТКРЫТИЕ программы
procedure TForm2.FormCreate(Sender: TObject);
begin
Memo1.ReadOnly:= true ;
//МАШТАБИРОВАНИЕ ФОРМЫ                  (!!!!!!)                (!!!!!!)
ScaleForm(form2) ;
end;

/// ПОЯВЛЕНИИ ОКНА программы
procedure TForm2.FormShow(Sender: TObject);
begin
//МАШТАБИРОВАНИЕ ФОРМЫ           (!!!!!!)           (!!!!!!)
   if screen.width=1152 then            //[1152]
      Memo1.Font.Size:=Memo1.Font.Size+2 ;
   if screen.width=1024 then            //[1024]
      Memo1.Font.Size:=Memo1.Font.Size+2 ;
   if screen.width=800 then            //[800]
      Memo1.Font.Size:=Memo1.Font.Size+3 ;

end;


///    [[[[МЕНЮ]]]]

//Выход
procedure TForm2.N1Click(Sender: TObject);
begin
Form2.Close ;
end;
//Сохранить
procedure TForm2.N3Click(Sender: TObject);
begin
Form2.Memo1.Lines.SaveToFile(ExtractFilePath( Application.ExeName )+'Отчет.txt');
end;
//Сохранить как...
procedure TForm2.N4Click(Sender: TObject);
begin
SaveDialog1.Title:='Соранить как...' ;
SaveDialog1.filename := ExtractFilePath( Application.ExeName )+'Отчет';
if SaveDialog1.Execute then
   try
      if FileExists(SaveDialog1.FileName) then   // если такой файл уже есть то перезаписываем
         begin
          DeleteFile(SaveDialog1.FileName);
          Memo1.Lines.SaveToFile(SaveDialog1.FileName)
         end
      else //если его нет то прибовляем к названию .txt
     Memo1.Lines.SaveToFile(SaveDialog1.FileName+'.txt')
  except
  end;
end;
//Печать...
procedure TForm2.N5Click(Sender: TObject);
var
Line: TextFile;
P: integer;
begin {кнопка "Печать (поле текст)"}
If PrintDialog1.Execute then
     begin
       AssignPrn(Line);
       ReWrite(Line);
       Printer.Canvas.Font := Memo1.Font;
       for P := 0 to Memo1.Lines.Count -1 do Writeln (Line, Memo1.Lines[P]);
       System.CloseFile(Line);
     end;
end;
// О программе
procedure TForm2.N6Click(Sender: TObject);
begin
Form5.ShowModal;

end;


END.

 