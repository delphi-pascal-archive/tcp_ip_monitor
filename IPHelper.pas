unit IPHelper;

(*
  ==========================
  Delphi IPHelper functions
  ==========================
  Required OS : NT4/SP4 or higher, WIN98/WIN98se
  Developed on:  D6 Ent. & Prof.
  Tested on   :  WIN-NT4/SP6, WIN98se, WIN95/OSR1
              :  WIN98, W2K-SP2, 3, 4
              :  W2K, W2K prof, W2K server

  Warning - currently only supports Delphi 5 and later unless int64 is removed
  (Int64 is only used to force Format to show unsigned 32-bit numbers)

  ================================================================
                    This software is FREEWARE
                    -------------------------
  If this software works, it was surely written by Dirk Claessens
                    dirkcl@pandora.be
        (If it doesn't, I don't know anything about it.)
  ================================================================

{ List of Fixes & Additions

v1.1  dirkcl
-----
Fix :  wrong errorcode reported in GetNetworkParams()
Fix :  RTTI MaxHops 20 > 128
Add :  ICMP -statistics
Add :  Well-Known port numbers
Add :  RecentIP list
Add :  Timer update

v1.2   dirkcl
----
Fix :  Recent IP's correct update
ADD :  ICMP-error codes translated

v1.3 - 18th September 2001
----
  Angus Robertson, Magenta Systems Ltd, England
     delphi@magsys.co.uk, http://www.magsys.co.uk/delphi/
  Slowly converting procs into functions that can be used by other programs,
     ie Get_ becomes IpHlp
  Primary improvements are that current DNS server is now shown, also
     in/out bytes for each interface (aka adaptor)
  All functions are dynamically loaded so program can be used on W95/NT4
  Tested with Delphi 6 on Windows 2000 and XP

v1.4 - 28th February 2002 - Angus
----
  Fixed major memory leak in IpHlpIfTable (except instead of finally)
  Fixed major memory leak in Get_AdaptersInfo (incremented buffer pointer)
  Created IpHlpAdaptersInfo which returns TAdaptorRows


  Note: IpHlpNetworkParams returns dynamic DNS address (and other stuff)
  Note: IpHlpIfEntry returns bytes in/out for a network adaptor

v1.5 - 5th October 2003
----
  Jean-Pierre Turchi "From South of France" <jpturchi@mageos.com>
  Cosmetic (more readable) and add-in's from iana.org in "WellKnownPorts"

}
*)

interface

uses
  Windows, Messages, SysUtils, Classes, Dialogs, IpHlpApi;

const
  NULL_IP       = '  0.  0.  0.  0';

//------conversion of well-known port numbers to service names----------------

type
  TWellKnownPort = record
    Prt: DWORD;
    Srv: string[20];
  end;


const
    // only most "popular" services...
  WellKnownPorts: array[1..32] of TWellKnownPort
  = (
//    ( Prt: 0; Srv:   'RESRVED' ),     {Reserved}
    ( Prt: 7; Srv:   'ECHO   ' ),     {Ping    }
    ( Prt: 9; Srv:   'DISCARD' ),
    ( Prt: 13; Srv:  'DAYTIME' ),
    ( Prt: 17; Srv:  'QOTD   ' ),     {Quote Of The Day}
    ( Prt: 19; Srv:  'CHARGEN' ),     {Character Generator}
    ( Prt: 20; Srv:  'FTPDATA' ),     { File Transfer Protocol - datas}
    ( Prt: 21; Srv:  'FTPCTRL' ),     { File Transfer Protocol - Control}
    ( Prt: 22; Srv:  'SSH    ' ),
    ( Prt: 23; Srv:  'TELNET ' ),
    ( Prt: 25; Srv:  'SMTP   ' ),     { Simple Mail Transfer Protocol}
    ( Prt: 37; Srv:  'TIME   ' ),     { Time Protocol }
    ( Prt: 43; Srv:  'WHOIS  ' ),     { WHO IS service  }
    ( Prt: 53; Srv:  'DNS    ' ),     { Domain Name Service }
    ( Prt: 67; Srv:  'BOOTPS ' ),     { BOOTP Server }
    ( Prt: 68; Srv:  'BOOTPC ' ),     { BOOTP Client }
    ( Prt: 69; Srv:  'TFTP   ' ),     { Trivial FTP  }
    ( Prt: 70; Srv:  'GOPHER ' ),     { Gopher       }
    ( Prt: 79; Srv:  'FINGER ' ),     { Finger       }
    ( Prt: 80; Srv:  'HTTP   ' ),     { HTTP         }
    ( Prt: 88; Srv:  'KERBROS' ),     { Kerberos     }
    ( Prt: 109; Srv: 'POP2   ' ),     { Post Office Protocol Version 2 }
    ( Prt: 110; Srv: 'POP3   ' ),     { Post Office Protocol Version 3 }
    ( Prt: 111; Srv: 'SUN_RPC' ),     { SUN Remote Procedure Call }
    ( Prt: 119; Srv: 'NNTP   ' ),     { Network News Transfer Protocol }
    ( Prt: 123; Srv: 'NTP    ' ),     { Network Time protocol          }
    ( Prt: 135; Srv: 'DCOMRPC' ),     { Location Service              }
    ( Prt: 137; Srv: 'NBNAME ' ),     { NETBIOS Name service          }
    ( Prt: 138; Srv: 'NBDGRAM' ),     { NETBIOS Datagram Service     }
    ( Prt: 139; Srv: 'NBSESS ' ),     { NETBIOS Session Service        }
    ( Prt: 143; Srv: 'IMAP   ' ),     { Internet Message Access Protocol }
    ( Prt: 161; Srv: 'SNMP   ' ),     { Simple Netw. Management Protocol }
    ( Prt: 169; Srv: 'SEND   ' )
    );


//-----------conversion of ICMP error codes to strings--------------------------
             {taken from www.sockets.com/ms_icmp.c }

const
  ICMP_ERROR_BASE = 11000;
  IcmpErr : array[1..22] of string =
  (
   'IP_BUFFER_TOO_SMALL','IP_DEST_NET_UNREACHABLE', 'IP_DEST_HOST_UNREACHABLE',
   'IP_PROTOCOL_UNREACHABLE', 'IP_DEST_PORT_UNREACHABLE', 'IP_NO_RESOURCES',
   'IP_BAD_OPTION','IP_HARDWARE_ERROR', 'IP_PACKET_TOO_BIG', 'IP_REQUEST_TIMED_OUT',
   'IP_BAD_REQUEST','IP_BAD_ROUTE', 'IP_TTL_EXPIRED_TRANSIT',
   'IP_TTL_EXPIRED_REASSEM','IP_PARAMETER_PROBLEM', 'IP_SOURCE_QUENCH',
   'IP_OPTION_TOO_BIG', 'IP_BAD_DESTINATION','IP_ADDRESS_DELETED',
   'IP_SPEC_MTU_CHANGE', 'IP_MTU_CHANGE', 'IP_UNLOAD'
  );


//----------conversion of diverse enumerated values to strings------------------

  ARPEntryType  : array[1..4] of string = ( 'Other', 'Invalid',
    'Dynamic', 'Static'
    );
  TCPConnState  :
    array[1..12] of string =
    ( 'closed', 'listening', 'syn_sent',
    'syn_rcvd', 'established', 'fin_wait1',
    'fin_wait2', 'close_wait', 'closing',
    'last_ack', 'time_wait', 'delete_tcb'
    );

  TCPToAlgo     : array[1..4] of string =
    ( 'Const.Timeout', 'MIL-STD-1778',
    'Van Jacobson', 'Other' );

  IPForwTypes   : array[1..4] of string =
    ( 'other', 'invalid', 'local', 'remote' );

  IPForwProtos  : array[1..18] of string =
    ( 'OTHER', 'LOCAL', 'NETMGMT', 'ICMP', 'EGP',
    'GGP', 'HELO', 'RIP', 'IS_IS', 'ES_IS',
    'CISCO', 'BBN', 'OSPF', 'BGP', 'BOOTP',
    'AUTO_STAT', 'STATIC', 'NOT_DOD' );

type

// for IpHlpNetworkParams
  TNetworkParams = record
    HostName: string ;
    DomainName: string ;
    CurrentDnsServer: string ;
    DnsServerTot: integer ;
    DnsServerNames: array [0..9] of string ;
    NodeType: UINT;
    ScopeID: string ;
    EnableRouting: UINT;
    EnableProxy: UINT;
    EnableDNS: UINT;
  end;

  TIfRows = array of TMibIfRow ; // dynamic array of rows

// for IpHlpAdaptersInfo
  TAdaptorInfo = record
    AdapterName: string ;
    Description: string ;
    MacAddress: string ;
    Index: DWORD;
    aType: UINT;
    DHCPEnabled: UINT;
    CurrIPAddress: string ;
    CurrIPMask: string ;
    IPAddressTot: integer ;
    IPAddressList: array of string ;
    IPMaskList: array of string ;
    GatewayTot: integer ;
    GatewayList: array of string ;
    DHCPTot: integer ;
    DHCPServer: array of string ;
    HaveWINS: BOOL;
    PrimWINSTot: integer ;
    PrimWINSServer: array of string ;
    SecWINSTot: integer ;
    SecWINSServer: array of string ;
    LeaseObtained: LongInt ; // UNIX time, seconds since 1970
    LeaseExpires: LongInt;   // UNIX time, seconds since 1970
  end ;

  TAdaptorRows = array of TAdaptorInfo ;


//---------------exported stuff-----------------------------------------------

function IpHlpAdaptersInfo(var AdpTot: integer;var AdpRows: TAdaptorRows): integer ;
procedure Get_AdaptersInfo( List: TStrings );
function IpHlpNetworkParams (var NetworkParams: TNetworkParams): integer ;
procedure Get_NetworkParams( List: TStrings );
procedure Get_ARPTable( List: TStrings );
procedure Get_TCPTable( List: TStrings );
procedure Get_TCPStatistics( List: TStrings );
function IpHlpTCPStatistics (var TCPStats: TMibTCPStats): integer ;
procedure Get_UDPTable( List: TStrings );
procedure Get_UDPStatistics( List: TStrings );
function IpHlpUdpStatistics (UdpStats: TMibUDPStats): integer ;
procedure Get_IPAddrTable( List: TStrings );
procedure Get_IPForwardTable( List: TStrings );
procedure Get_IPStatistics( List: TStrings );
function IpHlpIPStatistics (var IPStats: TMibIPStats): integer ;
function Get_RTTAndHopCount( IPAddr: DWORD; MaxHops: Longint;
  var RTT: longint; var HopCount: longint ): integer;
procedure Get_ICMPStats( ICMPIn, ICMPOut: TStrings );
function IpHlpIfTable(var IfTot: integer; var IfRows: TIfRows): integer ;
procedure Get_IfTable( List: TStrings );
function IpHlpIfEntry(Index: integer; var IfRow: TMibIfRow): integer ;
procedure Get_RecentDestIPs( List: TStrings );

// conversion utils
function MacAddr2Str( MacAddr: TMacAddress; size: integer ): string;
function IpAddr2Str( IPAddr: DWORD ): string;
function Str2IpAddr( IPStr: string ): DWORD;
function Port2Str( nwoPort: DWORD ): string;
function Port2Wrd( nwoPort: DWORD ): DWORD;
function Port2Svc( Port: DWORD ): string;
function ICMPErr2Str( ICMPErrCode: DWORD) : string;

implementation

var
  RecentIPs     : TStringList;

//--------------General utilities-----------------------------------------------

{ extracts next "token" from string, then eats string }
function NextToken( var s: string; Separator: char ): string;
var
  Sep_Pos       : byte;
begin
  Result := '';
  if length( s ) > 0 then begin
    Sep_Pos := pos( Separator, s );
    if Sep_Pos > 0 then begin
      Result := copy( s, 1, Pred( Sep_Pos ) );
      Delete( s, 1, Sep_Pos );
    end
    else begin
      Result := s;
      s := '';
    end;
  end;
end;

//------------------------------------------------------------------------------
{ concerts numerical MAC-address to ww-xx-yy-zz string }
function MacAddr2Str( MacAddr: TMacAddress; size: integer ): string;
var
  i             : integer;
begin
  if Size = 0 then
  begin
    Result := '00-00-00-00-00-00';
    EXIT;
  end
  else Result := '';
  //
  for i := 1 to Size do
    Result := Result + IntToHex( MacAddr[i], 2 ) + '-';
  Delete( Result, Length( Result ), 1 );
end;

//------------------------------------------------------------------------------
{ converts IP-address in network byte order DWORD to dotted decimal string}
function IpAddr2Str( IPAddr: DWORD ): string;
var
  i             : integer;
begin
  Result := '';
  for i := 1 to 4 do
  begin
    Result := Result + Format( '%3d.', [IPAddr and $FF] );
    IPAddr := IPAddr shr 8;
  end;
  Delete( Result, Length( Result ), 1 );
end;

//------------------------------------------------------------------------------
{ converts dotted decimal IP-address to network byte order DWORD}
function Str2IpAddr( IPStr: string ): DWORD;
var
  i             : integer;
  Num           : DWORD;
begin
  Result := 0;
  for i := 1 to 4 do
  try
    Num := ( StrToInt( NextToken( IPStr, '.' ) ) ) shl 24;
    Result := ( Result shr 8 ) or Num;
  except
    Result := 0;
  end;

end;

//------------------------------------------------------------------------------
{ converts port number in network byte order to DWORD }
function Port2Wrd( nwoPort: DWORD ): DWORD;
begin
  Result := Swap( WORD( nwoPort ) );
end;

//------------------------------------------------------------------------------
{ converts port number in network byte order to string }
function Port2Str( nwoPort: DWORD ): string;
begin
  Result := IntToStr( Port2Wrd( nwoPort ) );
end;

//------------------------------------------------------------------------------
{ converts well-known port numbers to service ID }
function Port2Svc( Port: DWORD ): string;
var
  i             : integer;
begin
  Result := Format( '%4d', [Port] ); // in case port not found
  for i := Low( WellKnownPorts ) to High( WellKnownPorts ) do
    if Port = WellKnownPorts[i].Prt then
    begin
      Result := WellKnownPorts[i].Srv;
      BREAK;
    end;
end;

//-----------------------------------------------------------------------------
{ general,  fixed network parameters }

procedure Get_NetworkParams( List: TStrings );
var
    NetworkParams: TNetworkParams ;
    I, ErrorCode: integer ;
begin
    if not Assigned( List ) then EXIT;
    List.Clear;
    ErrorCode := IpHlpNetworkParams (NetworkParams) ;
    if ErrorCode <> 0 then
    begin
        List.Add (SysErrorMessage (ErrorCode));
        exit;
    end ;
    with NetworkParams do
    begin
        List.Add( 'HOSTNAME          : ' + HostName );
        List.Add( 'DOMAIN            : ' + DomainName );
        List.Add( 'NETBIOS NODE TYPE : ' + NETBIOSTypes[NodeType] );
        List.Add( 'DHCP SCOPE        : ' + ScopeID );
        List.Add( 'ROUTING ENABLED   : ' + IntToStr( EnableRouting ) );
        List.Add( 'PROXY   ENABLED   : ' + IntToStr( EnableProxy ) );
        List.Add( 'DNS     ENABLED   : ' + IntToStr( EnableDNS ) );
        if DnsServerTot <> 0 then
        begin
            for I := 0 to Pred (DnsServerTot) do
                List.Add( 'DNS SERVER ADDR   : ' + DnsServerNames [I] ) ;
        end ;
    end ;
end ;

//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *//
function IpHlpNetworkParams (var NetworkParams: TNetworkParams): integer ;
var
  FixedInfo     : PTFixedInfo;         // Angus
  InfoSize      : Longint;
  PDnsServer    : PTIP_ADDR_STRING ;   // Angus
begin
    InfoSize := 0 ;   // Angus
    result := ERROR_NOT_SUPPORTED ;
    if NOT LoadIpHlp then exit ;
    result := GetNetworkParams( Nil, @InfoSize );  // Angus
    if result <> ERROR_BUFFER_OVERFLOW then exit ; // Angus
    GetMem (FixedInfo, InfoSize) ;                    // Angus
    try
    result := GetNetworkParams( FixedInfo, @InfoSize );   // Angus
    if result <> ERROR_SUCCESS then exit ;
    NetworkParams.DnsServerTot := 0 ;
    with FixedInfo^ do
    begin
        NetworkParams.HostName := trim (HostName) ;
        NetworkParams.DomainName := trim (DomainName) ;
        NetworkParams.ScopeId := trim (ScopeID) ;
        NetworkParams.NodeType := NodeType ;
        NetworkParams.EnableRouting := EnableRouting ;
        NetworkParams.EnableProxy := EnableProxy ;
        NetworkParams.EnableDNS := EnableDNS ;
        NetworkParams.DnsServerNames [0] := DNSServerList.IPAddress ;  // Angus
        if NetworkParams.DnsServerNames [0] <> '' then
                                        NetworkParams.DnsServerTot := 1 ;
        PDnsServer := DnsServerList.Next;
        while PDnsServer <> Nil do
        begin
            NetworkParams.DnsServerNames [NetworkParams.DnsServerTot] :=
                                        PDnsServer^.IPAddress ;  // Angus
            inc (NetworkParams.DnsServerTot) ;
            if NetworkParams.DnsServerTot >=
                            Length (NetworkParams.DnsServerNames) then exit ;
            PDnsServer := PDnsServer.Next ;
        end;
    end ;
    finally
       FreeMem (FixedInfo) ;                     // Angus
    end ;
end;

//------------------------------------------------------------------------------

function ICMPErr2Str( ICMPErrCode: DWORD) : string;
begin
   Result := 'UnknownError : ' + IntToStr( ICMPErrCode );
   dec( ICMPErrCode, ICMP_ERROR_BASE );
   if ICMPErrCode in [Low(ICMpErr)..High(ICMPErr)] then
     Result := ICMPErr[ ICMPErrCode];
end;


//------------------------------------------------------------------------------

// include bytes in/out for each adaptor

function IpHlpIfTable(var IfTot: integer; var IfRows: TIfRows): integer ;
var
  I,
  TableSize   : integer;
  pBuf, pNext : PChar;
begin
  result := ERROR_NOT_SUPPORTED ;
  if NOT LoadIpHlp then exit ;
  SetLength (IfRows, 0) ;
  IfTot := 0 ; // Angus
  TableSize := 0;
   // first call: get memsize needed
  result := GetIfTable (Nil, @TableSize, false) ;  // Angus
  if result <> ERROR_INSUFFICIENT_BUFFER then exit ;
  GetMem( pBuf, TableSize );
  try
      FillChar (pBuf^, TableSize, #0);  // clear buffer, since W98 does not

   // get table pointer
      result := GetIfTable (PTMibIfTable (pBuf), @TableSize, false) ;
      if result <> NO_ERROR then exit ;
      IfTot := PTMibIfTable (pBuf)^.dwNumEntries ;
      if IfTot = 0 then exit ;
      SetLength (IfRows, IfTot) ;
      pNext := pBuf + SizeOf(IfTot) ;
      for i := 0 to Pred (IfTot) do
      begin
         IfRows [i] := PTMibIfRow (pNext )^ ;
         inc (pNext, SizeOf (TMibIfRow)) ;
      end;
  finally
      FreeMem (pBuf) ;
  end ;
end;

procedure Get_IfTable( List: TStrings );
var
  IfRows        : TIfRows ;
  Error, I      : integer;
  NumEntries    : integer;
  sDescr, sIfName: string ;
begin
  if not Assigned( List ) then EXIT;
  List.Clear;
  SetLength (IfRows, 0) ;
  Error := IpHlpIfTable (NumEntries, IfRows) ;
  if (Error <> 0) then
      List.Add( SysErrorMessage( GetLastError ) )
  else if NumEntries = 0 then
      List.Add( 'no entries.' )
  else
  begin
      for I := 0 to Pred (NumEntries) do
      begin
          with IfRows [I] do
          begin
             if wszName [1] = #0 then
                 sIfName := ''
             else
                 sIfName := WideCharToString (@wszName) ;  // convert Unicode to string
             sIfName := trim (sIfName) ;
             sDescr := bDescr ;
             sDescr := trim (sDescr);
             List.Add (Format (
               '%0.8x |%3d | %16s |%8d |%12d |%2d |%2d |%10d |%10d | %-s| %-s',
               [dwIndex, dwType, MacAddr2Str( TMacAddress( bPhysAddr ),
               dwPhysAddrLen ), dwMTU, dwSpeed, dwAdminStatus,
               dwOPerStatus, Int64 (dwInOctets), Int64 (dwOutOctets),  // counters are 32-bit
               sIfName, sDescr] )  // Angus, added in/out
               );
          end;
      end ;
  end ;
  SetLength (IfRows, 0) ;  // free memory
end ;

function IpHlpIfEntry(Index: integer; var IfRow: TMibIfRow): integer ;
begin
  result := ERROR_NOT_SUPPORTED ;
  if NOT LoadIpHlp then exit ;
  FillChar (IfRow, SizeOf (TMibIfRow), #0);  // clear buffer, since W98 does not
  IfRow.dwIndex := Index ;
  result := GetIfEntry (@IfRow) ;
end ;

//-----------------------------------------------------------------------------
{ Info on installed adapters }

function IpHlpAdaptersInfo(var AdpTot: integer; var AdpRows: TAdaptorRows): integer ;
var
  BufLen        : DWORD;
  AdapterInfo   : PTIP_ADAPTER_INFO;
  PIpAddr       : PTIP_ADDR_STRING;
  PBuf          : PCHAR ;
  I             : integer ;
begin
  SetLength (AdpRows, 4) ;
  AdpTot := 0 ;
  BufLen := 0 ;
  result := GetAdaptersInfo( Nil, @BufLen );
  if (result <> ERROR_INSUFFICIENT_BUFFER) and (result = NO_ERROR) then exit ;
  GetMem( pBuf, BufLen );
  try
      FillChar (pBuf^, BufLen, #0);  // clear buffer
      result := GetAdaptersInfo( PTIP_ADAPTER_INFO (PBuf), @BufLen );
      if result = NO_ERROR then
      begin
         AdapterInfo := PTIP_ADAPTER_INFO (PBuf) ;
         while ( AdapterInfo <> nil ) do
         begin
            AdpRows [AdpTot].IPAddressTot := 0 ;
            SetLength (AdpRows [AdpTot].IPAddressList, 2) ;
            SetLength (AdpRows [AdpTot].IPMaskList, 2) ;
            AdpRows [AdpTot].GatewayTot := 0 ;
            SetLength (AdpRows [AdpTot].GatewayList, 2) ;
            AdpRows [AdpTot].DHCPTot := 0 ;
            SetLength (AdpRows [AdpTot].DHCPServer, 2) ;
            AdpRows [AdpTot].PrimWINSTot := 0 ;
            SetLength (AdpRows [AdpTot].PrimWINSServer, 2) ;
            AdpRows [AdpTot].SecWINSTot := 0 ;
            SetLength (AdpRows [AdpTot].SecWINSServer, 2) ;
            AdpRows [AdpTot].CurrIPAddress := NULL_IP;
            AdpRows [AdpTot].CurrIPMask := NULL_IP;
            AdpRows [AdpTot].AdapterName := Trim( string( AdapterInfo^.AdapterName ) );
            AdpRows [AdpTot].Description := Trim( string( AdapterInfo^.Description ) );
            AdpRows [AdpTot].MacAddress := MacAddr2Str( TMacAddress(
                                 AdapterInfo^.Address ), AdapterInfo^.AddressLength ) ;
            AdpRows [AdpTot].Index := AdapterInfo^.Index ;
            AdpRows [AdpTot].aType := AdapterInfo^.aType ;
            AdpRows [AdpTot].DHCPEnabled := AdapterInfo^.DHCPEnabled ;
            if AdapterInfo^.CurrentIPAddress <> Nil then
            begin
                AdpRows [AdpTot].CurrIPAddress := AdapterInfo^.CurrentIPAddress.IpAddress ;
                AdpRows [AdpTot].CurrIPMask := AdapterInfo^.CurrentIPAddress.IpMask ;
            end ;

        // get list of IP addresses and masks for IPAddressList
            I := 0 ;
            PIpAddr := @AdapterInfo^.IPAddressList ;
            while (PIpAddr <> Nil) do
            begin
                AdpRows [AdpTot].IPAddressList [I] := PIpAddr.IpAddress ;
                AdpRows [AdpTot].IPMaskList [I] := PIpAddr.IpMask ;
                PIpAddr := PIpAddr.Next ;
                inc (I) ;
                if Length (AdpRows [AdpTot].IPAddressList) <= I then
                begin
                     SetLength (AdpRows [AdpTot].IPAddressList, I * 2) ;
                     SetLength (AdpRows [AdpTot].IPMaskList, I * 2) ;
                end ;
            end ;
            AdpRows [AdpTot].IPAddressTot := I ;

        // get list of IP addresses for GatewayList
            I := 0 ;
            PIpAddr := @AdapterInfo^.GatewayList ;
            while (PIpAddr <> Nil) do
            begin
                AdpRows [AdpTot].GatewayList [I] := PIpAddr.IpAddress ;
                PIpAddr := PIpAddr.Next ;
                inc (I) ;
                if Length (AdpRows [AdpTot].GatewayList) <= I then
                             SetLength (AdpRows [AdpTot].GatewayList, I * 2) ;
            end ;
            AdpRows [AdpTot].GatewayTot := I ;

        // get list of IP addresses for GatewayList
            I := 0 ;
            PIpAddr := @AdapterInfo^.DHCPServer ;
            while (PIpAddr <> Nil) do
            begin
                AdpRows [AdpTot].DHCPServer [I] := PIpAddr.IpAddress ;
                PIpAddr := PIpAddr.Next ;
                inc (I) ;
                if Length (AdpRows [AdpTot].DHCPServer) <= I then
                             SetLength (AdpRows [AdpTot].DHCPServer, I * 2) ;
            end ;
            AdpRows [AdpTot].DHCPTot := I ;

        // get list of IP addresses for PrimaryWINSServer
            I := 0 ;
            PIpAddr := @AdapterInfo^.PrimaryWINSServer ;
            while (PIpAddr <> Nil) do
            begin
                AdpRows [AdpTot].PrimWINSServer [I] := PIpAddr.IpAddress ;
                PIpAddr := PIpAddr.Next ;
                inc (I) ;
                if Length (AdpRows [AdpTot].PrimWINSServer) <= I then
                             SetLength (AdpRows [AdpTot].PrimWINSServer, I * 2) ;
            end ;
            AdpRows [AdpTot].PrimWINSTot := I ;

       // get list of IP addresses for SecondaryWINSServer
            I := 0 ;
            PIpAddr := @AdapterInfo^.SecondaryWINSServer ;
            while (PIpAddr <> Nil) do
            begin
                AdpRows [AdpTot].SecWINSServer [I] := PIpAddr.IpAddress ;
                PIpAddr := PIpAddr.Next ;
                inc (I) ;
                if Length (AdpRows [AdpTot].SecWINSServer) <= I then
                             SetLength (AdpRows [AdpTot].SecWINSServer, I * 2) ;
            end ;
            AdpRows [AdpTot].SecWINSTot := I ;

            AdpRows [AdpTot].LeaseObtained := AdapterInfo^.LeaseObtained ;
            AdpRows [AdpTot].LeaseExpires := AdapterInfo^.LeaseExpires ;

            inc (AdpTot) ;
            if Length (AdpRows) <= AdpTot then
                            SetLength (AdpRows, AdpTot * 2) ;  // more memory
            AdapterInfo := AdapterInfo^.Next;
         end ;
         SetLength (AdpRows, AdpTot) ;
      end ;
  finally
      FreeMem( pBuf );
  end ;
end ;

procedure Get_AdaptersInfo( List: TStrings );
var
  AdpTot: integer;
  AdpRows: TAdaptorRows ;
  Error: DWORD ;
  I: integer ;
  //J: integer ;  jpt - see below
  //S: string ;        id.
begin
  if not Assigned( List ) then EXIT;
  List.Clear;
  SetLength (AdpRows, 0) ;
  AdpTot := 0 ;
  Error := IpHlpAdaptersInfo(AdpTot, AdpRows) ;
  if (Error <> 0) then
      List.Add( SysErrorMessage( GetLastError ) )
  else if AdpTot = 0 then
      List.Add( 'no entries.' )
  else
  begin
      for I := 0 to Pred (AdpTot) do
      begin
        with AdpRows [I] do
        begin
            //List.Add(AdapterName + '|' + Description ); // jpt : not useful
            List.Add( Format('%8.8x | %6s | %16s | %2d | %16s | %16s | %16s',
                [Index, AdaptTypes[aType], MacAddress, DHCPEnabled,
                GatewayList [0], DHCPServer [0], PrimWINSServer [0]] ) );
            {if IPAddressTot <> 0 then    // jpt : not useful
            begin
                S := '' ;
                for J := 0 to Pred (IPAddressTot) do
                        S := S + IPAddressList [J] + '/' + IPMaskList [J] + ' | ' ;
                List.Add(IntToStr (IPAddressTot) + ' IP Addresse(s): ' + S);
            end ;
            List.Add( '  ' ); }
        end ;
      end ;
  end ;
  SetLength (AdpRows, 0) ;
end ;

//-----------------------------------------------------------------------------
{ get round trip time and hopcount to indicated IP }
function Get_RTTAndHopCount( IPAddr: DWORD; MaxHops: Longint; var RTT: Longint;
  var HopCount: Longint ): integer;
begin
  if not GetRTTAndHopCount( IPAddr, @HopCount, MaxHops, @RTT ) then
  begin
    Result := GetLastError;
    RTT := -1; // Destination unreachable, BAD_HOST_NAME,etc...
    HopCount := -1;
  end
  else
    Result := NO_ERROR;
end;

//-----------------------------------------------------------------------------
{ ARP-table lists relations between remote IP and remote MAC-address.
 NOTE: these are cached entries ;when there is no more network traffic to a
 node, entry is deleted after a few minutes.
}
procedure Get_ARPTable( List: TStrings );
var
  IPNetRow      : TMibIPNetRow;
  TableSize     : DWORD;
  NumEntries    : DWORD;
  ErrorCode     : DWORD;
  i             : integer;
  pBuf          : PChar;
begin
  if not Assigned( List ) then EXIT;
  List.Clear;
  // first call: get table length
  TableSize := 0;
  ErrorCode := GetIPNetTable( Nil, @TableSize, false );   // Angus
  //
  if ErrorCode = ERROR_NO_DATA then
  begin
    List.Add( ' ARP-cache empty.' );
    EXIT;
  end;
  // get table
  GetMem( pBuf, TableSize );
  NumEntries := 0 ;
  try
  ErrorCode := GetIpNetTable( PTMIBIPNetTable( pBuf ), @TableSize, false );
  if ErrorCode = NO_ERROR then
  begin
    NumEntries := PTMIBIPNetTable( pBuf )^.dwNumEntries;
    if NumEntries > 0 then // paranoia striking, but you never know...
    begin
      inc( pBuf, SizeOf( DWORD ) ); // get past table size
      for i := 1 to NumEntries do
      begin
        IPNetRow := PTMIBIPNetRow( PBuf )^;
        with IPNetRow do
          List.Add( Format( '%8x | %12s | %16s | %10s',
                           [dwIndex, MacAddr2Str( bPhysAddr, dwPhysAddrLen ),
                           IPAddr2Str( dwAddr ), ARPEntryType[dwType]
                           ]));
        inc( pBuf, SizeOf( IPNetRow ) );
      end;
    end
    else
      List.Add( ' ARP-cache empty.' );
  end
  else
    List.Add( SysErrorMessage( ErrorCode ) );

  // we _must_ restore pointer!
  finally
      dec( pBuf, SizeOf( DWORD ) + NumEntries * SizeOf( IPNetRow ) );
      FreeMem( pBuf );
  end ;
end;


//------------------------------------------------------------------------------
procedure Get_TCPTable( List: TStrings );
var
  TCPRow        : TMIBTCPRow;
  i,
    NumEntries  : integer;
  TableSize     : DWORD;
  ErrorCode     : DWORD;
  DestIP        : string;
  pBuf          : PChar;
begin
  if not Assigned( List ) then EXIT;
  List.Clear;
  RecentIPs.Clear;
  // first call : get size of table
  TableSize := 0;
  NumEntries := 0 ;
  ErrorCode := GetTCPTable(Nil, @TableSize, false );  // Angus
  if Errorcode <> ERROR_INSUFFICIENT_BUFFER then
    EXIT;

  // get required memory size, call again
  GetMem( pBuf, TableSize );
  // get table
  ErrorCode := GetTCPTable( PTMIBTCPTable( pBuf ), @TableSize, false );
  if ErrorCode = NO_ERROR then
  begin

    NumEntries := PTMIBTCPTable( pBuf )^.dwNumEntries;
    if NumEntries > 0 then
    begin
      inc( pBuf, SizeOf( DWORD ) ); // get past table size
      for i := 1 to NumEntries do
      begin
        TCPRow := PTMIBTCPRow( pBuf )^; // get next record
        with TCPRow do
        begin
          if dwRemoteAddr = 0 then
            dwRemotePort := 0;
          DestIP := IPAddr2Str( dwRemoteAddr );
          List.Add(
            Format( '%15s : %-7s | %15s : %-7s | %-16s',
            [IpAddr2Str( dwLocalAddr ),
            Port2Svc( Port2Wrd( dwLocalPort ) ),
              DestIP,
              Port2Svc( Port2Wrd( dwRemotePort ) ),
              TCPConnState[dwState]
              ] ) );
         //
            if (not ( dwRemoteAddr = 0 ))
            and ( RecentIps.IndexOf(DestIP) = -1 ) then
               RecentIPs.Add( DestIP );
        end;
        inc( pBuf, SizeOf( TMIBTCPRow ) );
      end;
    end;
  end
  else
    List.Add( SyserrorMessage( ErrorCode ) );
  dec( pBuf, SizeOf( DWORD ) + NumEntries * SizeOf( TMibTCPRow ) );
  FreeMem( pBuf );
end;

//------------------------------------------------------------------------------
procedure Get_TCPStatistics( List: TStrings );
var
  TCPStats      : TMibTCPStats;
  ErrorCode     : DWORD;
begin
  if not Assigned( List ) then EXIT;
  List.Clear;
  if NOT LoadIpHlp then exit ;
  ErrorCode := GetTCPStatistics( @TCPStats );
  if ErrorCode = NO_ERROR then
    with TCPStats do
    begin
      List.Add( 'Retransmission algorithm : ' + TCPToAlgo[dwRTOAlgorithm] );
      List.Add( 'Minimum Time-Out         : ' + IntToStr( dwRTOMin ) + ' ms' );
      List.Add( 'Maximum Time-Out         : ' + IntToStr( dwRTOMax ) + ' ms' );
      List.Add( 'Maximum Pend.Connections : ' + IntToStr( dwRTOAlgorithm ) );
      List.Add( 'Active Opens             : ' + IntToStr( dwActiveOpens ) );
      List.Add( 'Passive Opens            : ' + IntToStr( dwPassiveOpens ) );
      List.Add( 'Failed Open Attempts     : ' + IntToStr( dwAttemptFails ) );
      List.Add( 'Established conn. Reset  : ' + IntToStr( dwEstabResets ) );
      List.Add( 'Current Established Conn.: ' + IntToStr( dwCurrEstab ) );
      List.Add( 'Segments Received        : ' + IntToStr( dwInSegs ) );
      List.Add( 'Segments Sent            : ' + IntToStr( dwOutSegs ) );
      List.Add( 'Segments Retransmitted   : ' + IntToStr( dwReTransSegs ) );
      List.Add( 'Incoming Errors          : ' + IntToStr( dwInErrs ) );
      List.Add( 'Outgoing Resets          : ' + IntToStr( dwOutRsts ) );
      List.Add( 'Cumulative Connections   : ' + IntToStr( dwNumConns ) );
    end
  else
    List.Add( SyserrorMessage( ErrorCode ) );
end;

function IpHlpTCPStatistics (var TCPStats: TMibTCPStats): integer ;
begin
    result := ERROR_NOT_SUPPORTED ;
    if NOT LoadIpHlp then exit ;
    result := GetTCPStatistics( @TCPStats );
end;

//------------------------------------------------------------------------------
procedure Get_UDPTable( List: TStrings );
var
  UDPRow        : TMIBUDPRow;
  i,
    NumEntries  : integer;
  TableSize     : DWORD;
  ErrorCode     : DWORD;
  pBuf          : PChar;
begin
  if not Assigned( List ) then EXIT;
  List.Clear;

  // first call : get size of table
  TableSize := 0;
  NumEntries := 0 ;
  ErrorCode := GetUDPTable(Nil, @TableSize, false );
  if Errorcode <> ERROR_INSUFFICIENT_BUFFER then
    EXIT;

  // get required size of memory, call again
  GetMem( pBuf, TableSize );

  // get table
  ErrorCode := GetUDPTable( PTMIBUDPTable( pBuf ), @TableSize, false );
  if ErrorCode = NO_ERROR then
  begin
    NumEntries := PTMIBUDPTable( pBuf )^.dwNumEntries;
    if NumEntries > 0 then
    begin
      inc( pBuf, SizeOf( DWORD ) ); // get past table size
      for i := 1 to NumEntries do
      begin
        UDPRow := PTMIBUDPRow( pBuf )^; // get next record
        with UDPRow do
          List.Add( Format( '%15s : %-6s',
            [IpAddr2Str( dwLocalAddr ),
            Port2Svc( Port2Wrd( dwLocalPort ) )
              ] ) );
        inc( pBuf, SizeOf( TMIBUDPRow ) );
      end;
    end
    else
      List.Add( 'no entries.' );
  end
  else
    List.Add( SyserrorMessage( ErrorCode ) );
  dec( pBuf, SizeOf( DWORD ) + NumEntries * SizeOf( TMibUDPRow ) );
  FreeMem( pBuf );
end;

//------------------------------------------------------------------------------
procedure Get_IPAddrTable( List: TStrings );
var
  IPAddrRow     : TMibIPAddrRow;
  TableSize     : DWORD;
  ErrorCode     : DWORD;
  i             : integer;
  pBuf          : PChar;
  NumEntries    : DWORD;
begin
  if not Assigned( List ) then EXIT;
  List.Clear;
  TableSize := 0; ;
  NumEntries := 0 ;
  // first call: get table length
  ErrorCode := GetIpAddrTable(Nil, @TableSize, true );  // Angus
  if Errorcode <> ERROR_INSUFFICIENT_BUFFER then
    EXIT;

  GetMem( pBuf, TableSize );
  // get table
  ErrorCode := GetIpAddrTable( PTMibIPAddrTable( pBuf ), @TableSize, true );
  if ErrorCode = NO_ERROR then
  begin
    NumEntries := PTMibIPAddrTable( pBuf )^.dwNumEntries;
    if NumEntries > 0 then
    begin
      inc( pBuf, SizeOf( DWORD ) );
      for i := 1 to NumEntries do
      begin
        IPAddrRow := PTMIBIPAddrRow( pBuf )^;
        with IPAddrRow do
          List.Add( Format( '%8.8x | %15s | %15s | %15s | %8.8d',
            [dwIndex,
            IPAddr2Str( dwAddr ),
              IPAddr2Str( dwMask ),
              IPAddr2Str( dwBCastAddr ),
              dwReasmSize
              ] ) );
        inc( pBuf, SizeOf( TMIBIPAddrRow ) );
      end;
    end
    else
      List.Add( 'no entries.' );
  end
  else
    List.Add( SysErrorMessage( ErrorCode ) );

  // we must restore pointer!
  dec( pBuf, SizeOf( DWORD ) + NumEntries * SizeOf( IPAddrRow ) );
  FreeMem( pBuf );
end;

//-----------------------------------------------------------------------------
{ gets entries in routing table; equivalent to "Route Print" }
procedure Get_IPForwardTable( List: TStrings );
var
  IPForwRow     : TMibIPForwardRow;
  TableSize     : DWORD;
  ErrorCode     : DWORD;
  i             : integer;
  pBuf          : PChar;
  NumEntries    : DWORD;
begin

  if not Assigned( List ) then EXIT;
  List.Clear;
  TableSize := 0;

  // first call: get table length
  NumEntries := 0 ;
  ErrorCode := GetIpForwardTable(Nil, @TableSize, true);
  if Errorcode <> ERROR_INSUFFICIENT_BUFFER then
    EXIT;

  // get table
  GetMem( pBuf, TableSize );
  ErrorCode := GetIpForwardTable( PTMibIPForwardTable( pBuf ), @TableSize, true);
  if ErrorCode = NO_ERROR then
  begin
    NumEntries := PTMibIPForwardTable( pBuf )^.dwNumEntries;
    if NumEntries > 0 then
    begin
      inc( pBuf, SizeOf( DWORD ) );
      for i := 1 to NumEntries do
      begin
        IPForwRow := PTMibIPForwardRow( pBuf )^;
        with IPForwRow do
        begin
          if (dwForwardType < 1)
          or (dwForwardType > 4) then
                   dwForwardType := 1 ;   // Angus, allow for bad value
          List.Add( Format(
            '%15s | %15s | %15s | %8.8x | %7s |   %5.5d |  %7s |   %2.2d',
            [IPAddr2Str( dwForwardDest ),
            IPAddr2Str( dwForwardMask ),
              IPAddr2Str( dwForwardNextHop ),
              dwForwardIFIndex,
              IPForwTypes[dwForwardType],
              dwForwardNextHopAS,
              IPForwProtos[dwForwardProto],
              dwForwardMetric1
              ] ) );
        end ;       
        inc( pBuf, SizeOf( TMibIPForwardRow ) );
      end;
    end
    else
      List.Add( 'no entries.' );
  end
  else
    List.Add( SysErrorMessage( ErrorCode ) );
  dec( pBuf, SizeOf( DWORD ) + NumEntries * SizeOf( TMibIPForwardRow ) );
  FreeMem( pBuf );
end;

//------------------------------------------------------------------------------
procedure Get_IPStatistics( List: TStrings );
var
  IPStats       : TMibIPStats;
  ErrorCode     : integer;
begin
  if not Assigned( List ) then EXIT;
  if NOT LoadIpHlp then exit ;
  ErrorCode := GetIPStatistics( @IPStats );
  if ErrorCode = NO_ERROR then
  begin
    List.Clear;
    with IPStats do
    begin
      if dwForwarding = 1 then
        List.add( 'Forwarding Enabled      : ' + 'Yes' )
      else
        List.add( 'Forwarding Enabled      : ' + 'No' );
      List.add( 'Default TTL             : ' + inttostr( dwDefaultTTL ) );
      List.add( 'Datagrams Received      : ' + inttostr( dwInReceives ) );
      List.add( 'Header Errors     (In)  : ' + inttostr( dwInHdrErrors ) );
      List.add( 'Address Errors    (In)  : ' + inttostr( dwInAddrErrors ) );
      List.add( 'Datagrams Forwarded     : ' + inttostr( dwForwDatagrams ) );   // Angus
      List.add( 'Unknown Protocols (In)  : ' + inttostr( dwInUnknownProtos ) );
      List.add( 'Datagrams Discarded     : ' + inttostr( dwInDiscards ) );
      List.add( 'Datagrams Delivered     : ' + inttostr( dwInDelivers ) );
      List.add( 'Requests Out            : ' + inttostr( dwOutRequests ) );
      List.add( 'Routings Discarded      : ' + inttostr( dwRoutingDiscards ) );
      List.add( 'No Routes        (Out)  : ' + inttostr( dwOutNoRoutes ) );
      List.add( 'Reassemble TimeOuts     : ' + inttostr( dwReasmTimeOut ) );
      List.add( 'Reassemble Requests     : ' + inttostr( dwReasmReqds ) );
      List.add( 'Succesfull Reassemblies : ' + inttostr( dwReasmOKs ) );
      List.add( 'Failed Reassemblies     : ' + inttostr( dwReasmFails ) );
      List.add( 'Succesful Fragmentations: ' + inttostr( dwFragOKs ) );
      List.add( 'Failed Fragmentations   : ' + inttostr( dwFragFails ) );
      List.add( 'Datagrams Fragmented    : ' + inttostr( dwFRagCreates ) );
      List.add( 'Number of Interfaces    : ' + inttostr( dwNumIf ) );
      List.add( 'Number of IP-addresses  : ' + inttostr( dwNumAddr ) );
      List.add( 'Routes in RoutingTable  : ' + inttostr( dwNumRoutes ) );
    end;
  end
  else
    List.Add( SysErrorMessage( ErrorCode ) );
end;

function IpHlpIPStatistics (var IPStats: TMibIPStats): integer ;      // Angus
begin
    result := ERROR_NOT_SUPPORTED ;
    if NOT LoadIpHlp then exit ;
    result := GetIPStatistics( @IPStats );
end ;

//------------------------------------------------------------------------------
procedure Get_UdpStatistics( List: TStrings );
var
  UdpStats      : TMibUDPStats;
  ErrorCode     : integer;
begin
  if not Assigned( List ) then EXIT;
  ErrorCode := GetUDPStatistics( @UdpStats );
  if ErrorCode = NO_ERROR then
  begin
    List.Clear;
    with UDPStats do
    begin
      List.add( 'Datagrams (In)    : ' + inttostr( dwInDatagrams ) );
      List.add( 'Datagrams (Out)   : ' + inttostr( dwOutDatagrams ) );
      List.add( 'No Ports          : ' + inttostr( dwNoPorts ) );
      List.add( 'Errors    (In)    : ' + inttostr( dwInErrors ) );
      List.add( 'UDP Listen Ports  : ' + inttostr( dwNumAddrs ) );
    end;
  end
  else
    List.Add( SysErrorMessage( ErrorCode ) );
end;

//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *//
function IpHlpUdpStatistics (UdpStats: TMibUDPStats): integer ;     // Angus
begin
    result := ERROR_NOT_SUPPORTED ;
    if NOT LoadIpHlp then exit ;
    result := GetUDPStatistics (@UdpStats) ;
end ;

//------------------------------------------------------------------------------
procedure Get_ICMPStats( ICMPIn, ICMPOut: TStrings );
var
  ErrorCode     : DWORD;
  ICMPStats     : PTMibICMPInfo;
begin
  if ( ICMPIn = nil ) or ( ICMPOut = nil ) then EXIT;
  ICMPIn.Clear;
  ICMPOut.Clear;
  New( ICMPStats );
  ErrorCode := GetICMPStatistics( ICMPStats );
  if ErrorCode = NO_ERROR then
  begin
    with ICMPStats.InStats do
    begin
      ICMPIn.Add( 'Messages received    : ' + IntToStr( dwMsgs ) );
      ICMPIn.Add( 'Errors               : ' + IntToStr( dwErrors ) );
      ICMPIn.Add( 'Dest. Unreachable    : ' + IntToStr( dwDestUnreachs ) );
      ICMPIn.Add( 'Time Exceeded        : ' + IntToStr( dwTimeEcxcds ) );
      ICMPIn.Add( 'Param. Problems      : ' + IntToStr( dwParmProbs ) );
      ICMPIn.Add( 'Source Quench        : ' + IntToStr( dwSrcQuenchs ) );
      ICMPIn.Add( 'Redirects            : ' + IntToStr( dwRedirects ) );
      ICMPIn.Add( 'Echo Requests        : ' + IntToStr( dwEchos ) );
      ICMPIn.Add( 'Echo Replies         : ' + IntToStr( dwEchoReps ) );
      ICMPIn.Add( 'Timestamp Requests   : ' + IntToStr( dwTimeStamps ) );
      ICMPIn.Add( 'Timestamp Replies    : ' + IntToStr( dwTimeStampReps ) );
      ICMPIn.Add( 'Addr. Masks Requests : ' + IntToStr( dwAddrMasks ) );
      ICMPIn.Add( 'Addr. Mask Replies   : ' + IntToStr( dwAddrReps ) );
    end;
     //
    with ICMPStats.OutStats do
    begin
      ICMPOut.Add( 'Messages sent        : ' + IntToStr( dwMsgs ) );
      ICMPOut.Add( 'Errors               : ' + IntToStr( dwErrors ) );
      ICMPOut.Add( 'Dest. Unreachable    : ' + IntToStr( dwDestUnreachs ) );
      ICMPOut.Add( 'Time Exceeded        : ' + IntToStr( dwTimeEcxcds ) );
      ICMPOut.Add( 'Param. Problems      : ' + IntToStr( dwParmProbs ) );
      ICMPOut.Add( 'Source Quench        : ' + IntToStr( dwSrcQuenchs ) );
      ICMPOut.Add( 'Redirects            : ' + IntToStr( dwRedirects ) );
      ICMPOut.Add( 'Echo Requests        : ' + IntToStr( dwEchos ) );
      ICMPOut.Add( 'Echo Replies         : ' + IntToStr( dwEchoReps ) );
      ICMPOut.Add( 'Timestamp Requests   : ' + IntToStr( dwTimeStamps ) );
      ICMPOut.Add( 'Timestamp Replies    : ' + IntToStr( dwTimeStampReps ) );
      ICMPOut.Add( 'Addr. Masks Requests : ' + IntToStr( dwAddrMasks ) );
      ICMPOut.Add( 'Addr. Mask Replies   : ' + IntToStr( dwAddrReps ) );
    end;
  end
  else
    IcmpIn.Add( SysErrorMessage( ErrorCode ) );
  Dispose( ICMPStats );
end;

//------------------------------------------------------------------------------
procedure Get_RecentDestIPs( List: TStrings );
begin
  if Assigned( List ) then
    List.Assign( RecentIPs )
end;

initialization

  RecentIPs := TStringList.Create;

finalization

  RecentIPs.Free;

end.


