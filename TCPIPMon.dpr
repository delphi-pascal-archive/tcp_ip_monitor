program TCPIPMon;

uses
  Forms,
  uTCPIP in 'uTCPIP.pas' {IPForm},
  IPHelper in 'IPHelper.pas',
  IPHLPAPI in 'IPHLPAPI.pas';

{$R *.RES}

begin
  Application.Initialize;
  Application.Title:='TCP/IP/ARP monitor';
  Application.CreateForm(TIPForm, IPForm);
  Application.Run;
end.
