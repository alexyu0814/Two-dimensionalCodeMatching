program QRCode;

uses
  Vcl.Forms,
  formQRCode in 'formQRCode.pas' {frmQRCode};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TfrmQRCode, frmQRCode);
  Application.Run;
end.
