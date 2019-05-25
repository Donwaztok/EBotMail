program EBotMail;

uses
  Forms,
  Controls,
  Main in 'Main.pas' {Form1},
  Email in 'Email.pas' {Form2},
  FormBG in 'FormBG.pas' {Form3},
  EASendMailObjLib_TLB in 'EASendMailObjLib_TLB.pas';

{$R *.RES}

begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.CreateForm(TForm3, Form3);
  Application.Run;
end.
