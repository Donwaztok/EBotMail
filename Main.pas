unit Main;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.Imaging.pngimage, Vcl.ExtCtrls,
  Vcl.StdCtrls, Vcl.ComCtrls, Vcl.Imaging.jpeg, System.Types, EASendMailObjLib_TLB;

type
  TForm1 = class(TForm)
    Footer: TImage;
    Header: TImage;
    Background: TImage;
    Menu: TImage;
    Maximizar: TImage;
    Fechar: TImage;
    Minimizar: TImage;
    Version: TLabel;
    BTN_Enviar: TButton;
    Corpo: TMemo;
    Assunto: TEdit;
    Anexo: TEdit;
    Emails: TMemo;
    Progresso: TLabel;
    ProgressBar1: TProgressBar;
    Label1: TLabel;
    Label2: TLabel;
    BTN_Anexo: TButton;
    OpenDialog1: TOpenDialog;
    Label3: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure MaximizarClick(Sender: TObject);
    procedure FecharClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure BTN_EnviarClick(Sender: TObject);
    procedure BTN_AnexoClick(Sender: TObject);
  private
    { Private declarations }
    procedure WMNCHitTest(var Msg: TWMNCHitTest);
      message WM_NCHITTEST;
    procedure WMGetMinmaxInfo(var Msg: TWMGetMinmaxInfo);
      message WM_GETMINMAXINFO;
  public
    { Public declarations }
    FAllowSize: Boolean;
    Email,Senha:string;
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

uses Email, FormBG;


//== Função da Versão do Aplicativo ============================================
Function VersaoExe: String;
type
   PFFI = ^vs_FixedFileInfo;
var
   F       : PFFI;
   Handle  : Dword;
   Len     : Longint;
   Data    : Pchar;
   Buffer  : Pointer;
   Tamanho : Dword;
   Parquivo: Pchar;
   Arquivo : String;
begin
   Arquivo  := Application.ExeName;
   Parquivo := StrAlloc(Length(Arquivo) + 1);
   StrPcopy(Parquivo, Arquivo);
   Len := GetFileVersionInfoSize(Parquivo, Handle);
   Result := '';
   if Len > 0 then
   begin
      Data:=StrAlloc(Len+1);
      if GetFileVersionInfo(Parquivo,Handle,Len,Data) then
      begin
         VerQueryValue(Data, '',Buffer,Tamanho);
         F := PFFI(Buffer);
         Result := Format('%d.%d.%d.%d',
                          [HiWord(F^.dwFileVersionMs),
                           LoWord(F^.dwFileVersionMs),
                           HiWord(F^.dwFileVersionLs),
                           Loword(F^.dwFileVersionLs)]
                         );
      end;
      StrDispose(Data);
   end;
   StrDispose(Parquivo);
end;
//==============================================================================


//=== Correção do Maximizar ====================================================
procedure TForm1.WMGETMINMAXINFO(var Msg: TWMGetMinmaxInfo);
var
  R: TRect;
begin
  inherited;

  // Obtem o retangulo com a area livre do desktop
  SystemParametersInfo(SPI_GETWORKAREA, SizeOf(R), @R, 0);

  Msg.MinMaxInfo^.ptMaxPosition := R.TopLeft;
  OffsetRect(R, -R.Left, -R.Top);
  Msg.MinMaxInfo^.ptMaxSize := R.BottomRight;
end;

//=== Permitir alterar o tamanho do Form sem borda =============================
procedure TForm1.WMNCHitTest(var Msg: TWMNCHitTest);
var
  ScreenPt : TPoint;
  MoveArea : TRect;
  HANDLE_WIDTH: Integer;
  SIZEGRIP: Integer;
begin
HANDLE_WIDTH := 3;
Sizegrip := 19;
inherited;
  if not (csDesigning in ComponentState) then
    begin
      ScreenPt := ScreenToClient(Point(Msg.Xpos, Msg.Ypos));
      MoveArea := Rect(HANDLE_WIDTH,
      HANDLE_WIDTH,
      Width - HANDLE_WIDTH,
      Height - HANDLE_WIDTH);
  if FAllowSize then
    begin
      // left side
      if (ScreenPt.x < HANDLE_WIDTH) then Msg.Result := HTLEFT
      // top side
      else if (ScreenPt.y < HANDLE_WIDTH) then Msg.Result := HTTOP
      // right side
      else if (ScreenPt.x >= Width - HANDLE_WIDTH) then Msg.Result := HTRIGHT
      // bottom side
      else if (ScreenPt.y >= Height - HANDLE_WIDTH) then Msg.Result := HTBOTTOM
      // top left corner
      //else if (ScreenPt.x < Sizegrip) and (ScreenPt.y < Sizegrip) then
      //  Msg.Result := HTTOPLEFT
      // bottom left corner
      else if (ScreenPt.x < Sizegrip) and (ScreenPt.y >= Height - Sizegrip) then
        Msg.Result := HTBOTTOMLEFT
      // top right corner
      //else if (ScreenPt.x >= Width - Sizegrip) and (ScreenPt.y < Sizegrip) then
      //  Msg.Result := HTTOPRIGHT
      // bottom right corner
      else if (ScreenPt.x >= Width - Sizegrip) and (ScreenPt.y >= Height - Sizegrip) then
        Msg.Result := HTBOTTOMRIGHT
    end;
  end;
{
// IF you want to allow moving the form, add an FAllowMove variable and
// set it to true, then uncomment this code.
if FAllowMove then
begin
// no sides or corners, this will do the dragging
if PtInRect(MoveArea, ScreenPt) then
Msg.Result := HTCAPTION;
end;
}
end;
//=== Adicionar Anexo ==========================================================
procedure TForm1.BTN_AnexoClick(Sender: TObject);
begin
if OpenDialog1.Execute then
  if OpenDialog1.FileName <> '' then Anexo.Text:=OpenDialog1.FileName;
end;
//=== Enviar Email =============================================================
procedure TForm1.BTN_EnviarClick(Sender: TObject);
var
  oSmtp : TMail;
  I,concluido:Integer;
begin
concluido:=0;
//Visuaisinhos \o
BTN_Enviar.Enabled:=False;
ProgressBar1.Min:=0;
ProgressBar1.Max:=Emails.Lines.Count;
Progresso.Caption:='Progresso: '+IntToStr(concluido)+' de '+IntToStr(Emails.Lines.Count);
ProgressBar1.Position:=concluido+1;
//enviar varios emails
for I := 0 to Emails.Lines.Count-1 do
  begin
    oSmtp := TMail.Create(Application);
    oSmtp.LicenseCode := 'TryIt';
    oSmtp.FromAddr := Email;              //Remetente
    oSmtp.Subject := Assunto.Text;          //Titulo do Email
    oSmtp.BodyText := Corpo.Text;         //Corpo do Email

    if Anexo.Text<>'' then
    if oSmtp.AddAttachment(Anexo.Text) <> 0 then
      begin
        ShowMessage('Erro com o Anexo: '+oSmtp.GetLastErrDescription());
        Break;
      end;

    oSmtp.ServerAddr := 'smtp.live.com';  //Server
    oSmtp.ServerPort := 587;              //Porta

    oSmtp.UserName := Email;              //seu login
    oSmtp.Password := Senha;              //sua senha

    oSmtp.SSL_init();

    //Enviar email
    oSmtp.AddRecipientEx(Emails.Lines[I], 0);
      if oSmtp.SendMail() = 0 then
        begin
          concluido:=concluido+1;
          ProgressBar1.Position:=concluido;
          Progresso.Caption:='Progresso: '+IntToStr(concluido)+' de '+IntToStr(Emails.Lines.Count);
          Invalidate; Application.ProcessMessages;
        end
      else
        begin
          ShowMessage( 'Falha ao enviar o E-mail com o seguinte erro: '
            + oSmtp.GetLastErrDescription());
          Break;
        end;
  end;
oSmtp.LogFileName:=
    ExtractFileDir(ParamStr(0))+'\Logs\Log-'+FormatDateTime('dd-mm-yyyy',Date)+'.txt';
oSmtp.Free;
BTN_Enviar.Enabled:=True;
end;
//=== Botão Fechar =============================================================
procedure TForm1.FecharClick(Sender: TObject);
begin
Application.Terminate;
end;
//=== Ativação do Form =========================================================
procedure TForm1.FormActivate(Sender: TObject);
begin
if Form2=nil then
  begin
    Form2:=TForm2.Create(Application);
    try
      if Form2.ShowModal = mrOK then
        begin
          Email:=Form2.Email.Text;
          Senha:=Form2.Senha.text;
        end
       else
        begin
          Application.Terminate;
        end;
    finally
      Form2.Free;
      Form2:=nil;
      end;
  end else Application.Terminate;
end;

//=== Criação do Form ==========================================================
procedure TForm1.FormCreate(Sender: TObject);
begin
//True Permite o reajuste do tamanho do Form sem borda
FAllowSize := True;
//Background ajustado com o tamanho do Form
Background.Top:=0;
Background.Left:=0;
Background.Width:=Form1.Width;
Background.Height:=Form1.Height;
//Menu Lateral ajustado com o tamanho do form
Menu.Height:=Form1.Height-129;
//Versão do Aplicativo
Version.Caption:='|  ['+VersaoExe+' ]';
end;
//=== Atualizar informações com a alteração do tamanho do Form =================
procedure TForm1.FormResize(Sender: TObject);
begin
//Background
Background.Width:=Form1.Width;
BackGround.Height:=Form1.Height;
//Menu Lateral
Menu.Height:=Form1.Height-129;
end;

//=== Maximizar ================================================================
procedure TForm1.MaximizarClick(Sender: TObject);
begin
if FAllowSize=True then //Maximixar e bloquear a alteração do tamanho do Form
  begin
    SendMessage(Handle, WM_SYSCOMMAND, SC_MAXIMIZE, 0);
    FAllowSize := False;
  end
else if FAllowSize=False then//Restaurar e permitir a alteração do tamanho do Form
  begin
    SendMessage(Handle, WM_SYSCOMMAND, SC_RESTORE, 0);
    FAllowSize := True;
  end;
end;

//==============================================================================


end.
