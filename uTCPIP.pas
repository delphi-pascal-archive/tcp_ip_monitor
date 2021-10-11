unit uTCPIP;

(*
  ======================
  Delphi  TCP/IP MONITOR
  ======================

  Developed on:  D6 PE
  Tested on   :  WIN-NT4/SP6, WIN98se, WIN95/OSR1
              :  WIN98, W2K-SP2, 3, 4

  ================================================================
                    This software is FREEWARE
                    -------------------------
  If this software works, it was surely written by Dirk Claessens
                    <dirkcl@pandora.be>
  (If it doesn't, I don't know anything about it.)
  ================================================================

v1.3 - 18th September 2001
----
  Angus Robertson, Magenta Systems Ltd, England
     delphi@magsys.co.uk, http://www.magsys.co.uk/delphi/
  Dynamic load DLL, show error if not available
  Re-arranged windows slightly

v1.4 - 28th February 2002 - Angus
  Re-arranged windows again so usable on 800x600 screen

v1.5 - 5th October 2003
----
  Jean-Pierre Turchi "From South of France" <jpturchi@mageos.com>
  Cosmetic for IPForm :
  - Removed the Lucida Console font, using Courier New
  - Re-re-arranged windows (no more need of scrollbars
    on memos of 'statistics' tab. Still usable on 800x600)
  - Added 3 labels for information purpose

*)



interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ExtCtrls, IPHelper, StdCtrls, ComCtrls, Buttons, IpHlpApi, jpeg;

const
   version = 'TCP/IP/ARP monitor 1.5';
type
  TIPForm = class( TForm )
    Timer1: TTimer;
    PageControl1: TPageControl;
    ARPSheet: TTabSheet;
    ConnSheet: TTabSheet;
    IP1Sheet: TTabSheet;
    ARPMemo: TMemo;
    StaticText1: TStaticText;
    TCPMemo: TMemo;
    StaticText2: TStaticText;
    IPAddrMemo: TMemo;
    StaticText6: TStaticText;
    IP2Sheet: TTabSheet;
    IPForwMemo: TMemo;
    StaticText8: TStaticText;
    AdaptSheet: TTabSheet;
    AdaptMemo: TMemo;
    SpeedButton1: TSpeedButton;
    cbTimer: TCheckBox;
    StaticText9: TStaticText;
    NwMemo: TMemo;
    StaticText10: TStaticText;
    TabSheet1: TTabSheet;
    ICMPInMemo: TMemo;
    ICMPOutMemo: TMemo;
    StaticText12: TStaticText;
    btRTTI: TSpeedButton;
    cbRecentIPs: TComboBox;
    edtRTTI: TEdit;
    StaticText14: TStaticText;
    IfMemo: TMemo;
    StaticText15: TStaticText;
    TCPStatMemo: TMemo;
    StaticText7: TStaticText;
    UDPStatsMemo: TMemo;
    StaticText4: TStaticText;
    IPStatsMemo: TMemo;
    StaticText5: TStaticText;
    UDPMemo: TMemo;
    StaticText3: TStaticText;
    Panel1: TPanel;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    procedure Timer1Timer( Sender: TObject );
    procedure FormCreate( Sender: TObject );
    procedure SpeedButton1Click( Sender: TObject );
    procedure cbTimerClick( Sender: TObject );
    procedure btRTTIClick( Sender: TObject );
    procedure cbRecentIPsClick( Sender: TObject );
  private
    { Private declarations }
    procedure DOIpStuff;
  public
    { Public declarations }
  end;

var
  IPForm        : TIPForm;

implementation

{$R *.DFM}

//------------------------------------------------------------------------------
procedure TIPForm.FormCreate( Sender: TObject );
begin
  PageControl1.ActivePage := ConnSheet;
  Caption :=  Version;
  if LoadIpHlp then
  begin
      DOIpStuff;
      Timer1.Enabled := true;
  end
  else
      ShowMessage( 'Internet Helper DLL Not Available or Not Supported') ;
end;

//------------------------------------------------------------------------------
procedure TIPForm.DOIpStuff;
begin
  Get_NetworkParams( NwMemo.Lines );
  Get_ARPTable( ARPMemo.Lines );
  Get_TCPTable( TCPMemo.Lines );
  Get_TCPStatistics( TCPStatMemo.Lines );
  Get_UDPTable( UDPMemo.Lines );
  Get_IPStatistics( IPStatsMemo.Lines );
  Get_IPAddrTable( IPAddrMemo.Lines );
  Get_IPForwardTable( IPForwMemo.Lines );
  Get_UDPStatistics( UDPStatsMemo.Lines );
  Get_AdaptersInfo( AdaptMemo.Lines );
  Get_ICMPStats( ICMPInMemo.Lines, ICMPOutMemo.Lines );
  Get_IfTable( IfMemo.Lines );
  Get_RecentDestIPs( cbRecentIPs.Items );
end;

//------------------------------------------------------------------------------
procedure TIPForm.cbTimerClick( Sender: TObject );
begin
  Timer1.Enabled := (cbTimer.State = cbCHECKED) ;
end;

//------------------------------------------------------------------------------
procedure TIPForm.Timer1Timer( Sender: TObject );
begin
  if cbTimer.State = cbCHECKED then
  begin
    Timer1.Enabled := false;
    DoIPStuff;
    Timer1.Enabled := true;
  end;
end;

//------------------------------------------------------------------------------
procedure TIPForm.SpeedButton1Click( Sender: TObject );
begin
  Speedbutton1.Enabled := false;
  DoIPStuff;
  Speedbutton1.Enabled := true;
end;

//------------------------------------------------------------------------------
procedure TIPForm.btRTTIClick( Sender: TObject );
var
  IPadr         : dword;
  Rtt, HopCount : longint;
  Res           : integer;
begin
  btRTTI.Enabled := false;
  Screen.Cursor := crHOURGLASS;
  IPadr := Str2IPAddr( edtRTTI.Text );
  Res := Get_RTTAndHopCount( IPadr, 128, RTT, HopCount );
  if Res = NO_ERROR then
    ShowMessage( ' Round Trip Time '
      + inttostr( rtt ) + ' ms, '
      + inttostr( HopCount )
      + ' hops to : ' + edtRTTI.Text
      )
  else
    ShowMessage( 'Error occurred:' + #13
                 + ICMPErr2Str( Res ) ) ;
  btRTTI.Enabled := true;
  Screen.Cursor := crDEFAULT;
end;

//------------------------------------------------------------------------------
procedure TIPForm.cbRecentIPsClick( Sender: TObject );
begin
  edtRTTI.Text := cbRecentIPs.Items[cbRecentIPs.ItemIndex];
end;

end.
