unit EASendMailObjLib_TLB;

// ************************************************************************ //
// WARNING                                                                    
// -------                                                                    
// The types declared in this file were generated from data read from a       
// Type Library. If this type library is explicitly or indirectly (via        
// another type library referring to this type library) re-imported, or the   
// 'Refresh' command of the Type Library Editor activated while editing the   
// Type Library, the contents of this file will be regenerated and all        
// manual modifications will be lost.                                         
// ************************************************************************ //

// $Rev: 52393 $
// File generated on 05/11/2015 14:45:48 from Type Library described below.

// ************************************************************************  //
// Type Lib: C:\Program Files (x86)\EASendMail\EASendMailObj.dll (1)
// LIBID: {8B5A2BD0-5638-4CCA-A7FF-91B9E6768AC4}
// LCID: 0
// Helpfile: 
// HelpString: EASendMailObj ActiveX Object 1.0 Type Library
// DepndLst: 
//   (1) v2.0 stdole, (C:\Windows\SysWOW64\stdole2.tlb)
// SYS_KIND: SYS_WIN32
// Errors:
//   Error creating palette bitmap of (TMail) : Server C:\Program Files (x86)\EASendMail\EASendMailObj.dll contains no icons
//   Error creating palette bitmap of (TFastSender) : Server C:\Program Files (x86)\EASendMail\EASendMailObj.dll contains no icons
//   Error creating palette bitmap of (TCertificate) : Server C:\Program Files (x86)\EASendMail\EASendMailObj.dll contains no icons
//   Error creating palette bitmap of (TCertificateCollection) : Server C:\Program Files (x86)\EASendMail\EASendMailObj.dll contains no icons
// ************************************************************************ //
{$TYPEDADDRESS OFF} // Unit must be compiled without type-checked pointers. 
{$WARN SYMBOL_PLATFORM OFF}
{$WRITEABLECONST ON}
{$VARPROPSETTER ON}
{$ALIGN 4}

interface

uses Winapi.Windows, System.Classes, System.Variants, System.Win.StdVCL, Vcl.Graphics, Vcl.OleServer, Winapi.ActiveX;
  

// *********************************************************************//
// GUIDS declared in the TypeLibrary. Following prefixes are used:        
//   Type Libraries     : LIBID_xxxx                                      
//   CoClasses          : CLASS_xxxx                                      
//   DISPInterfaces     : DIID_xxxx                                       
//   Non-DISP interfaces: IID_xxxx                                        
// *********************************************************************//
const
  // TypeLibrary Major and minor versions
  EASendMailObjLibMajorVersion = 1;
  EASendMailObjLibMinorVersion = 0;

  LIBID_EASendMailObjLib: TGUID = '{8B5A2BD0-5638-4CCA-A7FF-91B9E6768AC4}';

  IID_IMail: TGUID = '{1AD28FC9-0C71-4E89-85C9-CAECDE8BE3AB}';
  IID_ICertificate: TGUID = '{A2809780-C98E-4C6D-A552-DAB146D4AD12}';
  IID_ICertificateCollection: TGUID = '{DC8D5635-B8E7-441E-B550-CE1BF3BA5C55}';
  IID_IFastSender: TGUID = '{92298BE3-ADEC-438F-800C-CF6311A7DF1D}';
  DIID__IMailEvents: TGUID = '{68CB8B02-D4AA-4A16-97A0-6B9488F98189}';
  CLASS_Mail: TGUID = '{DF8A4FE2-221A-4504-987A-3FD4720DB929}';
  DIID__IFastSenderEvents: TGUID = '{A1B45F08-67E7-4276-A7CA-7664C08F9EF7}';
  CLASS_FastSender: TGUID = '{FF80631D-E750-4C67-AFC3-5170AB72518B}';
  CLASS_Certificate: TGUID = '{EAFC4EAA-9390-492A-8E53-E179527780F6}';
  CLASS_CertificateCollection: TGUID = '{036C2F8C-8D3C-4F4B-9B36-3B6F1D29C0B4}';
type

// *********************************************************************//
// Forward declaration of types defined in TypeLibrary                    
// *********************************************************************//
  IMail = interface;
  IMailDisp = dispinterface;
  ICertificate = interface;
  ICertificateDisp = dispinterface;
  ICertificateCollection = interface;
  ICertificateCollectionDisp = dispinterface;
  IFastSender = interface;
  IFastSenderDisp = dispinterface;
  _IMailEvents = dispinterface;
  _IFastSenderEvents = dispinterface;

// *********************************************************************//
// Declaration of CoClasses defined in Type Library                       
// (NOTE: Here we map each CoClass to its Default Interface)              
// *********************************************************************//
  Mail = IMail;
  FastSender = IFastSender;
  Certificate = ICertificate;
  CertificateCollection = ICertificateCollection;


// *********************************************************************//
// Interface: IMail
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {1AD28FC9-0C71-4E89-85C9-CAECDE8BE3AB}
// *********************************************************************//
  IMail = interface(IDispatch)
    ['{1AD28FC9-0C71-4E89-85C9-CAECDE8BE3AB}']
    function Get_BodyFormat: Integer; safecall;
    procedure Set_BodyFormat(pVal: Integer); safecall;
    function Get_BodyText: WideString; safecall;
    procedure Set_BodyText(const pVal: WideString); safecall;
    function Get_Charset: WideString; safecall;
    procedure Set_Charset(const pVal: WideString); safecall;
    function Get_From: WideString; safecall;
    procedure Set_From(const pVal: WideString); safecall;
    function Get_FromAddr: WideString; safecall;
    procedure Set_FromAddr(const pVal: WideString); safecall;
    function Get_LogFileName: WideString; safecall;
    procedure Set_LogFileName(const pVal: WideString); safecall;
    function Get_LicenseCode: WideString; safecall;
    procedure Set_LicenseCode(const pVal: WideString); safecall;
    function Get_ServerAddr: WideString; safecall;
    procedure Set_ServerAddr(const pVal: WideString); safecall;
    function Get_ServerPort: Integer; safecall;
    procedure Set_ServerPort(pVal: Integer); safecall;
    function Get_Subject: WideString; safecall;
    procedure Set_Subject(const pVal: WideString); safecall;
    function Get_ReplyTo: WideString; safecall;
    procedure Set_ReplyTo(const pVal: WideString); safecall;
    function Get_Priority: Integer; safecall;
    procedure Set_Priority(pVal: Integer); safecall;
    function Get_Timeout: Integer; safecall;
    procedure Set_Timeout(pVal: Integer); safecall;
    function Get_UserName: WideString; safecall;
    procedure Set_UserName(const pVal: WideString); safecall;
    function Get_Password: WideString; safecall;
    procedure Set_Password(const pVal: WideString); safecall;
    function Get_Version: WideString; safecall;
    function Get_Asynchronous: Integer; safecall;
    procedure Set_Asynchronous(pVal: Integer); safecall;
    function Get_AltBody: WideString; safecall;
    procedure Set_AltBody(const pVal: WideString); safecall;
    function AddAttachment(const strFile: WideString): Integer; safecall;
    function AddRecipient(const strName: WideString; const strAddress: WideString; Flags: Integer): Integer; safecall;
    procedure ClearAttachment; safecall;
    procedure ClearRecipient; safecall;
    procedure ConvertHTML(Flags: Integer); safecall;
    function ImportMail(const strFile: WideString): Integer; safecall;
    procedure Reset; safecall;
    function SendMail: Integer; safecall;
    function AddAttachmentEx(const strFile: WideString; const strAlt: WideString): Integer; safecall;
    function AddInline(const strFile: WideString): WideString; safecall;
    function AddInlineEx(const strFile: WideString; const strAlt: WideString): WideString; safecall;
    procedure ClearInline; safecall;
    function SaveMail(const strFile: WideString): Integer; safecall;
    function AddHeader(const strHeader: WideString; const strValue: WideString): Integer; safecall;
    procedure ClearHeader; safecall;
    procedure Terminate; safecall;
    function GetLastError: Integer; safecall;
    function GetLastErrDescription: WideString; safecall;
    function Get_Anonymous: Integer; safecall;
    procedure Set_Anonymous(pVal: Integer); safecall;
    procedure SetMailer(const Mailer: WideString); safecall;
    function Get_KeepConnection: Integer; safecall;
    procedure Set_KeepConnection(pVal: Integer); safecall;
    function ImportMailEx(const strFile: WideString): Integer; safecall;
    function Get_TransferEncoding: Integer; safecall;
    procedure Set_TransferEncoding(pVal: Integer); safecall;
    function GetEmailServer(const EmailAddr: WideString): WideString; safecall;
    function AddRecipientEx(const AddressList: WideString; Flags: Integer): Integer; safecall;
    function AddAttachments(const sPath: WideString): Integer; safecall;
    function Get_ComputerName: WideString; safecall;
    procedure Set_ComputerName(const pVal: WideString); safecall;
    function Get_BodyFormatEx: WideString; safecall;
    procedure Set_BodyFormatEx(const pVal: WideString); safecall;
    function Get_HeaderEncoding: Integer; safecall;
    procedure Set_HeaderEncoding(pVal: Integer); safecall;
    function SaveMailEx(const PickupPath: WideString): Integer; safecall;
    function TestEmailAddr: Integer; safecall;
    function GetAllEmailServers(const EmailAddr: WideString): WideString; safecall;
    function GetEmailContent: WideString; safecall;
    function GetEmailHeaders: WideString; safecall;
    function GetAllRecipients: WideString; safecall;
    function GetSenderAddr: WideString; safecall;
    function Get_TryAllSmtpServers: Integer; safecall;
    procedure Set_TryAllSmtpServers(pVal: Integer); safecall;
    function CreateFolder(const FolderName: WideString): Integer; safecall;
    function DeleteFile(const FileName: WideString): Integer; safecall;
    function Get_RawModeEnable: Integer; safecall;
    procedure Set_RawModeEnable(pVal: Integer); safecall;
    function Get_WrapEmailAddr: Integer; safecall;
    procedure Set_WrapEmailAddr(pVal: Integer); safecall;
    function Get_DeliveryNotification: Integer; safecall;
    procedure Set_DeliveryNotification(pVal: Integer); safecall;
    function Get__Idle: Integer; safecall;
    function SSL_init: Integer; safecall;
    function Get_SSL_ignorecerterror: Integer; safecall;
    procedure Set_SSL_ignorecerterror(pVal: Integer); safecall;
    function Get_SSL_starttls: Integer; safecall;
    procedure Set_SSL_starttls(pVal: Integer); safecall;
    procedure SSL_uninit; safecall;
    function Get_SSL_enabled: Integer; safecall;
    function Get_raw_Content: WideString; safecall;
    procedure Set_raw_Content(const pVal: WideString); safecall;
    function Get_LogLevel: Integer; safecall;
    procedure Set_LogLevel(pVal: Integer); safecall;
    function Get_SignerCert: ICertificate; safecall;
    function Get_RecipientsCerts: ICertificateCollection; safecall;
    procedure WriteLog(const LogContent: WideString); safecall;
    function Get_ReturnPath: WideString; safecall;
    procedure Set_ReturnPath(const pVal: WideString); safecall;
    function Get_LocalIP: WideString; safecall;
    procedure Set_LocalIP(const pVal: WideString); safecall;
    function ImportHtml(const html: WideString; const BasePath: WideString): Integer; safecall;
    function AddAttachment1(const FileName: WideString; Stream: OleVariant): Integer; safecall;
    function Get_AuthType: Integer; safecall;
    procedure Set_AuthType(pVal: Integer); safecall;
    function Get_SpecialFlags: Integer; safecall;
    procedure Set_SpecialFlags(pVal: Integer); safecall;
    function Get_DisplayTo: WideString; safecall;
    procedure Set_DisplayTo(const pVal: WideString); safecall;
    function Get_Date: TDateTime; safecall;
    procedure Set_Date(pVal: TDateTime); safecall;
    function Get_MessageID: WideString; safecall;
    procedure Set_MessageID(const pVal: WideString); safecall;
    procedure AppendBody(const BodyText: WideString; bAlt: Integer); safecall;
    function AddInline1(const FileName: WideString; Stream: OleVariant): WideString; safecall;
    function SendMailToQueue: Integer; safecall;
    function Get_NoWrapBody: Integer; safecall;
    procedure Set_NoWrapBody(pVal: Integer); safecall;
    function Get_EncryptionAlgorithm: Integer; safecall;
    procedure Set_EncryptionAlgorithm(pVal: Integer); safecall;
    procedure ClearHeaderEx(const HeaderName: WideString); safecall;
    function GetEmailChunk: OleVariant; safecall;
    function AddAttachmentCT(const FileName: WideString; const ContentType: WideString): Integer; safecall;
    function Get_SocksProxyServer: WideString; safecall;
    procedure Set_SocksProxyServer(const pVal: WideString); safecall;
    function Get_SocksProxyUser: WideString; safecall;
    procedure Set_SocksProxyUser(const pVal: WideString); safecall;
    function Get_SocksProxyPassword: WideString; safecall;
    procedure Set_SocksProxyPassword(const pVal: WideString); safecall;
    function Get_SocksProxyPort: Integer; safecall;
    procedure Set_SocksProxyPort(pVal: Integer); safecall;
    function Get_ProxyProtocol: Integer; safecall;
    procedure Set_ProxyProtocol(pVal: Integer); safecall;
    function Get_DK_PublicKey: WideString; safecall;
    function LoadMessage(const FileName: WideString): Integer; safecall;
    function Get_ReadReceipt: WordBool; safecall;
    procedure Set_ReadReceipt(pVal: WordBool); safecall;
    function LoadMessageChunk(newVal: OleVariant): Integer; safecall;
    function Get_Recipients: OleVariant; safecall;
    function Get_Style: Integer; safecall;
    procedure Set_Style(pVal: Integer); safecall;
    procedure SetAttHeader(Index: Integer; const HeaderKey: WideString; 
                           const HeaderValue: WideString); safecall;
    function Get_AutoCalendar: Integer; safecall;
    procedure Set_AutoCalendar(pVal: Integer); safecall;
    function Get_AttachmentCount: Integer; safecall;
    function Get_DnsServerIP: WideString; safecall;
    procedure Set_DnsServerIP(const pVal: WideString); safecall;
    function SendMailToQueueEx(const Instant: WideString): Integer; safecall;
    function LoadRawMessage(const FileName: WideString; Flag: Integer): Integer; safecall;
    function Get_Protocol: Integer; safecall;
    procedure Set_Protocol(pVal: Integer); safecall;
    function Get_Alias: WideString; safecall;
    procedure Set_Alias(const pVal: WideString); safecall;
    function Get_Drafts: WideString; safecall;
    procedure Set_Drafts(const pVal: WideString); safecall;
    function Get_Sender: WideString; safecall;
    procedure Set_Sender(const pVal: WideString); safecall;
    procedure Quit; safecall;
    procedure Close; safecall;
    function Get_HttpProxyAuthType: Integer; safecall;
    procedure Set_HttpProxyAuthType(pVal: Integer); safecall;
    function Get_SMIMERFCCompatibility: WordBool; safecall;
    procedure Set_SMIMERFCCompatibility(pVal: WordBool); safecall;
    function Get_PIPELINING: Integer; safecall;
    procedure Set_PIPELINING(pVal: Integer); safecall;
    function Get_IgnoreDeliveryNotificationError: Integer; safecall;
    procedure Set_IgnoreDeliveryNotificationError(pVal: Integer); safecall;
    function Get_IPv6Policy: Integer; safecall;
    procedure Set_IPv6Policy(pVal: Integer); safecall;
    function Get_LocalIP6: WideString; safecall;
    procedure Set_LocalIP6(const pVal: WideString); safecall;
    function PostToRemoteQueue(const Instance: WideString; const URL: WideString; 
                               const User: WideString; const Password: WideString): Integer; safecall;
    property BodyFormat: Integer read Get_BodyFormat write Set_BodyFormat;
    property BodyText: WideString read Get_BodyText write Set_BodyText;
    property Charset: WideString read Get_Charset write Set_Charset;
    property From: WideString read Get_From write Set_From;
    property FromAddr: WideString read Get_FromAddr write Set_FromAddr;
    property LogFileName: WideString read Get_LogFileName write Set_LogFileName;
    property LicenseCode: WideString read Get_LicenseCode write Set_LicenseCode;
    property ServerAddr: WideString read Get_ServerAddr write Set_ServerAddr;
    property ServerPort: Integer read Get_ServerPort write Set_ServerPort;
    property Subject: WideString read Get_Subject write Set_Subject;
    property ReplyTo: WideString read Get_ReplyTo write Set_ReplyTo;
    property Priority: Integer read Get_Priority write Set_Priority;
    property Timeout: Integer read Get_Timeout write Set_Timeout;
    property UserName: WideString read Get_UserName write Set_UserName;
    property Password: WideString read Get_Password write Set_Password;
    property Version: WideString read Get_Version;
    property Asynchronous: Integer read Get_Asynchronous write Set_Asynchronous;
    property AltBody: WideString read Get_AltBody write Set_AltBody;
    property Anonymous: Integer read Get_Anonymous write Set_Anonymous;
    property KeepConnection: Integer read Get_KeepConnection write Set_KeepConnection;
    property TransferEncoding: Integer read Get_TransferEncoding write Set_TransferEncoding;
    property ComputerName: WideString read Get_ComputerName write Set_ComputerName;
    property BodyFormatEx: WideString read Get_BodyFormatEx write Set_BodyFormatEx;
    property HeaderEncoding: Integer read Get_HeaderEncoding write Set_HeaderEncoding;
    property TryAllSmtpServers: Integer read Get_TryAllSmtpServers write Set_TryAllSmtpServers;
    property RawModeEnable: Integer read Get_RawModeEnable write Set_RawModeEnable;
    property WrapEmailAddr: Integer read Get_WrapEmailAddr write Set_WrapEmailAddr;
    property DeliveryNotification: Integer read Get_DeliveryNotification write Set_DeliveryNotification;
    property _Idle: Integer read Get__Idle;
    property SSL_ignorecerterror: Integer read Get_SSL_ignorecerterror write Set_SSL_ignorecerterror;
    property SSL_starttls: Integer read Get_SSL_starttls write Set_SSL_starttls;
    property SSL_enabled: Integer read Get_SSL_enabled;
    property raw_Content: WideString read Get_raw_Content write Set_raw_Content;
    property LogLevel: Integer read Get_LogLevel write Set_LogLevel;
    property SignerCert: ICertificate read Get_SignerCert;
    property RecipientsCerts: ICertificateCollection read Get_RecipientsCerts;
    property ReturnPath: WideString read Get_ReturnPath write Set_ReturnPath;
    property LocalIP: WideString read Get_LocalIP write Set_LocalIP;
    property AuthType: Integer read Get_AuthType write Set_AuthType;
    property SpecialFlags: Integer read Get_SpecialFlags write Set_SpecialFlags;
    property DisplayTo: WideString read Get_DisplayTo write Set_DisplayTo;
    property Date: TDateTime read Get_Date write Set_Date;
    property MessageID: WideString read Get_MessageID write Set_MessageID;
    property NoWrapBody: Integer read Get_NoWrapBody write Set_NoWrapBody;
    property EncryptionAlgorithm: Integer read Get_EncryptionAlgorithm write Set_EncryptionAlgorithm;
    property SocksProxyServer: WideString read Get_SocksProxyServer write Set_SocksProxyServer;
    property SocksProxyUser: WideString read Get_SocksProxyUser write Set_SocksProxyUser;
    property SocksProxyPassword: WideString read Get_SocksProxyPassword write Set_SocksProxyPassword;
    property SocksProxyPort: Integer read Get_SocksProxyPort write Set_SocksProxyPort;
    property ProxyProtocol: Integer read Get_ProxyProtocol write Set_ProxyProtocol;
    property DK_PublicKey: WideString read Get_DK_PublicKey;
    property ReadReceipt: WordBool read Get_ReadReceipt write Set_ReadReceipt;
    property Recipients: OleVariant read Get_Recipients;
    property Style: Integer read Get_Style write Set_Style;
    property AutoCalendar: Integer read Get_AutoCalendar write Set_AutoCalendar;
    property AttachmentCount: Integer read Get_AttachmentCount;
    property DnsServerIP: WideString read Get_DnsServerIP write Set_DnsServerIP;
    property Protocol: Integer read Get_Protocol write Set_Protocol;
    property Alias: WideString read Get_Alias write Set_Alias;
    property Drafts: WideString read Get_Drafts write Set_Drafts;
    property Sender: WideString read Get_Sender write Set_Sender;
    property HttpProxyAuthType: Integer read Get_HttpProxyAuthType write Set_HttpProxyAuthType;
    property SMIMERFCCompatibility: WordBool read Get_SMIMERFCCompatibility write Set_SMIMERFCCompatibility;
    property PIPELINING: Integer read Get_PIPELINING write Set_PIPELINING;
    property IgnoreDeliveryNotificationError: Integer read Get_IgnoreDeliveryNotificationError write Set_IgnoreDeliveryNotificationError;
    property IPv6Policy: Integer read Get_IPv6Policy write Set_IPv6Policy;
    property LocalIP6: WideString read Get_LocalIP6 write Set_LocalIP6;
  end;

// *********************************************************************//
// DispIntf:  IMailDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {1AD28FC9-0C71-4E89-85C9-CAECDE8BE3AB}
// *********************************************************************//
  IMailDisp = dispinterface
    ['{1AD28FC9-0C71-4E89-85C9-CAECDE8BE3AB}']
    property BodyFormat: Integer dispid 1;
    property BodyText: WideString dispid 2;
    property Charset: WideString dispid 3;
    property From: WideString dispid 4;
    property FromAddr: WideString dispid 5;
    property LogFileName: WideString dispid 6;
    property LicenseCode: WideString dispid 7;
    property ServerAddr: WideString dispid 8;
    property ServerPort: Integer dispid 9;
    property Subject: WideString dispid 10;
    property ReplyTo: WideString dispid 11;
    property Priority: Integer dispid 12;
    property Timeout: Integer dispid 13;
    property UserName: WideString dispid 14;
    property Password: WideString dispid 15;
    property Version: WideString readonly dispid 16;
    property Asynchronous: Integer dispid 17;
    property AltBody: WideString dispid 18;
    function AddAttachment(const strFile: WideString): Integer; dispid 19;
    function AddRecipient(const strName: WideString; const strAddress: WideString; Flags: Integer): Integer; dispid 20;
    procedure ClearAttachment; dispid 21;
    procedure ClearRecipient; dispid 22;
    procedure ConvertHTML(Flags: Integer); dispid 23;
    function ImportMail(const strFile: WideString): Integer; dispid 24;
    procedure Reset; dispid 25;
    function SendMail: Integer; dispid 26;
    function AddAttachmentEx(const strFile: WideString; const strAlt: WideString): Integer; dispid 27;
    function AddInline(const strFile: WideString): WideString; dispid 28;
    function AddInlineEx(const strFile: WideString; const strAlt: WideString): WideString; dispid 29;
    procedure ClearInline; dispid 30;
    function SaveMail(const strFile: WideString): Integer; dispid 31;
    function AddHeader(const strHeader: WideString; const strValue: WideString): Integer; dispid 32;
    procedure ClearHeader; dispid 33;
    procedure Terminate; dispid 34;
    function GetLastError: Integer; dispid 35;
    function GetLastErrDescription: WideString; dispid 36;
    property Anonymous: Integer dispid 37;
    procedure SetMailer(const Mailer: WideString); dispid 38;
    property KeepConnection: Integer dispid 39;
    function ImportMailEx(const strFile: WideString): Integer; dispid 40;
    property TransferEncoding: Integer dispid 41;
    function GetEmailServer(const EmailAddr: WideString): WideString; dispid 42;
    function AddRecipientEx(const AddressList: WideString; Flags: Integer): Integer; dispid 43;
    function AddAttachments(const sPath: WideString): Integer; dispid 44;
    property ComputerName: WideString dispid 45;
    property BodyFormatEx: WideString dispid 46;
    property HeaderEncoding: Integer dispid 47;
    function SaveMailEx(const PickupPath: WideString): Integer; dispid 48;
    function TestEmailAddr: Integer; dispid 49;
    function GetAllEmailServers(const EmailAddr: WideString): WideString; dispid 50;
    function GetEmailContent: WideString; dispid 51;
    function GetEmailHeaders: WideString; dispid 52;
    function GetAllRecipients: WideString; dispid 53;
    function GetSenderAddr: WideString; dispid 54;
    property TryAllSmtpServers: Integer dispid 55;
    function CreateFolder(const FolderName: WideString): Integer; dispid 56;
    function DeleteFile(const FileName: WideString): Integer; dispid 57;
    property RawModeEnable: Integer dispid 58;
    property WrapEmailAddr: Integer dispid 59;
    property DeliveryNotification: Integer dispid 60;
    property _Idle: Integer readonly dispid 61;
    function SSL_init: Integer; dispid 62;
    property SSL_ignorecerterror: Integer dispid 63;
    property SSL_starttls: Integer dispid 64;
    procedure SSL_uninit; dispid 65;
    property SSL_enabled: Integer readonly dispid 66;
    property raw_Content: WideString dispid 67;
    property LogLevel: Integer dispid 68;
    property SignerCert: ICertificate readonly dispid 69;
    property RecipientsCerts: ICertificateCollection readonly dispid 70;
    procedure WriteLog(const LogContent: WideString); dispid 71;
    property ReturnPath: WideString dispid 72;
    property LocalIP: WideString dispid 73;
    function ImportHtml(const html: WideString; const BasePath: WideString): Integer; dispid 74;
    function AddAttachment1(const FileName: WideString; Stream: OleVariant): Integer; dispid 75;
    property AuthType: Integer dispid 76;
    property SpecialFlags: Integer dispid 77;
    property DisplayTo: WideString dispid 78;
    property Date: TDateTime dispid 79;
    property MessageID: WideString dispid 80;
    procedure AppendBody(const BodyText: WideString; bAlt: Integer); dispid 81;
    function AddInline1(const FileName: WideString; Stream: OleVariant): WideString; dispid 82;
    function SendMailToQueue: Integer; dispid 83;
    property NoWrapBody: Integer dispid 84;
    property EncryptionAlgorithm: Integer dispid 85;
    procedure ClearHeaderEx(const HeaderName: WideString); dispid 86;
    function GetEmailChunk: OleVariant; dispid 87;
    function AddAttachmentCT(const FileName: WideString; const ContentType: WideString): Integer; dispid 88;
    property SocksProxyServer: WideString dispid 89;
    property SocksProxyUser: WideString dispid 90;
    property SocksProxyPassword: WideString dispid 91;
    property SocksProxyPort: Integer dispid 92;
    property ProxyProtocol: Integer dispid 93;
    property DK_PublicKey: WideString readonly dispid 94;
    function LoadMessage(const FileName: WideString): Integer; dispid 95;
    property ReadReceipt: WordBool dispid 96;
    function LoadMessageChunk(newVal: OleVariant): Integer; dispid 97;
    property Recipients: OleVariant readonly dispid 98;
    property Style: Integer dispid 99;
    procedure SetAttHeader(Index: Integer; const HeaderKey: WideString; 
                           const HeaderValue: WideString); dispid 100;
    property AutoCalendar: Integer dispid 101;
    property AttachmentCount: Integer readonly dispid 102;
    property DnsServerIP: WideString dispid 103;
    function SendMailToQueueEx(const Instant: WideString): Integer; dispid 104;
    function LoadRawMessage(const FileName: WideString; Flag: Integer): Integer; dispid 105;
    property Protocol: Integer dispid 106;
    property Alias: WideString dispid 107;
    property Drafts: WideString dispid 108;
    property Sender: WideString dispid 109;
    procedure Quit; dispid 110;
    procedure Close; dispid 111;
    property HttpProxyAuthType: Integer dispid 112;
    property SMIMERFCCompatibility: WordBool dispid 113;
    property PIPELINING: Integer dispid 114;
    property IgnoreDeliveryNotificationError: Integer dispid 115;
    property IPv6Policy: Integer dispid 116;
    property LocalIP6: WideString dispid 117;
    function PostToRemoteQueue(const Instance: WideString; const URL: WideString; 
                               const User: WideString; const Password: WideString): Integer; dispid 118;
  end;

// *********************************************************************//
// Interface: ICertificate
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {A2809780-C98E-4C6D-A552-DAB146D4AD12}
// *********************************************************************//
  ICertificate = interface(IDispatch)
    ['{A2809780-C98E-4C6D-A552-DAB146D4AD12}']
    function FindSubject(const FindKey: WideString; StoreLocation: Integer; 
                         const StoreName: WideString): WordBool; safecall;
    function LoadPFX(PFXContent: OleVariant; const Password: WideString; KeyLocation: Integer): WordBool; safecall;
    function LoadPFXFromFile(const PFXFile: WideString; const Password: WideString; 
                             KeyLocation: Integer): WordBool; safecall;
    function LoadCert(CERTContent: OleVariant): WordBool; safecall;
    function LoadCertFromFile(const CERTFile: WideString): WordBool; safecall;
    procedure Unload; safecall;
    function Get_HasCertificate: WordBool; safecall;
    function Get_Store: Largeuint; safecall;
    function Get_Handle: Largeuint; safecall;
    procedure Set_Handle(pVal: Largeuint); safecall;
    function SignMessage(Content: OleVariant): OleVariant; safecall;
    function Get_HasPrivateKey: WordBool; safecall;
    function GetLastError: WideString; safecall;
    property HasCertificate: WordBool read Get_HasCertificate;
    property Store: Largeuint read Get_Store;
    property Handle: Largeuint read Get_Handle write Set_Handle;
    property HasPrivateKey: WordBool read Get_HasPrivateKey;
  end;

// *********************************************************************//
// DispIntf:  ICertificateDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {A2809780-C98E-4C6D-A552-DAB146D4AD12}
// *********************************************************************//
  ICertificateDisp = dispinterface
    ['{A2809780-C98E-4C6D-A552-DAB146D4AD12}']
    function FindSubject(const FindKey: WideString; StoreLocation: Integer; 
                         const StoreName: WideString): WordBool; dispid 1;
    function LoadPFX(PFXContent: OleVariant; const Password: WideString; KeyLocation: Integer): WordBool; dispid 2;
    function LoadPFXFromFile(const PFXFile: WideString; const Password: WideString; 
                             KeyLocation: Integer): WordBool; dispid 3;
    function LoadCert(CERTContent: OleVariant): WordBool; dispid 4;
    function LoadCertFromFile(const CERTFile: WideString): WordBool; dispid 5;
    procedure Unload; dispid 6;
    property HasCertificate: WordBool readonly dispid 7;
    property Store: Largeuint readonly dispid 8;
    property Handle: Largeuint dispid 9;
    function SignMessage(Content: OleVariant): OleVariant; dispid 10;
    property HasPrivateKey: WordBool readonly dispid 11;
    function GetLastError: WideString; dispid 12;
  end;

// *********************************************************************//
// Interface: ICertificateCollection
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {DC8D5635-B8E7-441E-B550-CE1BF3BA5C55}
// *********************************************************************//
  ICertificateCollection = interface(IDispatch)
    ['{DC8D5635-B8E7-441E-B550-CE1BF3BA5C55}']
    function Get_Count: Integer; safecall;
    function Item(Index: Integer): ICertificate; safecall;
    procedure Add(const oCert: ICertificate); safecall;
    procedure Insert(Index: Integer; const oCert: ICertificate); safecall;
    procedure Clear; safecall;
    procedure RemoveAt(Index: Integer); safecall;
    function EncryptMessage(EncryptionAlgorithm: Integer; Content: OleVariant): OleVariant; safecall;
    function Get_HasEncryptCert: WordBool; safecall;
    function GetLastError: WideString; safecall;
    property Count: Integer read Get_Count;
    property HasEncryptCert: WordBool read Get_HasEncryptCert;
  end;

// *********************************************************************//
// DispIntf:  ICertificateCollectionDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {DC8D5635-B8E7-441E-B550-CE1BF3BA5C55}
// *********************************************************************//
  ICertificateCollectionDisp = dispinterface
    ['{DC8D5635-B8E7-441E-B550-CE1BF3BA5C55}']
    property Count: Integer readonly dispid 1;
    function Item(Index: Integer): ICertificate; dispid 2;
    procedure Add(const oCert: ICertificate); dispid 3;
    procedure Insert(Index: Integer; const oCert: ICertificate); dispid 4;
    procedure Clear; dispid 5;
    procedure RemoveAt(Index: Integer); dispid 6;
    function EncryptMessage(EncryptionAlgorithm: Integer; Content: OleVariant): OleVariant; dispid 7;
    property HasEncryptCert: WordBool readonly dispid 8;
    function GetLastError: WideString; dispid 9;
  end;

// *********************************************************************//
// Interface: IFastSender
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {92298BE3-ADEC-438F-800C-CF6311A7DF1D}
// *********************************************************************//
  IFastSender = interface(IDispatch)
    ['{92298BE3-ADEC-438F-800C-CF6311A7DF1D}']
    function Send(const pSmtp: IMail; nKey: Integer; const tParam: WideString): Integer; safecall;
    function Test(const pSmtp: IMail; nKey: Integer; const tParam: WideString): Integer; safecall;
    function Get_MaxThreads: Integer; safecall;
    procedure Set_MaxThreads(pVal: Integer); safecall;
    function GetCurrentThreads: Integer; safecall;
    function GetQueuedCount: Integer; safecall;
    procedure ClearQueuedMails; safecall;
    procedure StopAllThreads; safecall;
    function GetIdleThreads: Integer; safecall;
    function SendByPickup(const PickupPath: WideString; const pSmtp: IMail; nKey: Integer; 
                          const tParam: WideString): Integer; safecall;
    function SendEmlFile(const FileName: WideString; const senderAddr: WideString; 
                         const recipientAddrs: WideString; nKey: Integer; const tParam: WideString; 
                         const RegisterKey: WideString): Integer; safecall;
    function Get_ComputerName: WideString; safecall;
    procedure Set_ComputerName(const pVal: WideString); safecall;
    procedure LockEvent; safecall;
    procedure UnlockEvent; safecall;
    procedure ClearAllMails; safecall;
    procedure Pause; safecall;
    procedure Resume; safecall;
    function Get_KeepConnection: Integer; safecall;
    procedure Set_KeepConnection(pVal: Integer); safecall;
    property MaxThreads: Integer read Get_MaxThreads write Set_MaxThreads;
    property ComputerName: WideString read Get_ComputerName write Set_ComputerName;
    property KeepConnection: Integer read Get_KeepConnection write Set_KeepConnection;
  end;

// *********************************************************************//
// DispIntf:  IFastSenderDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {92298BE3-ADEC-438F-800C-CF6311A7DF1D}
// *********************************************************************//
  IFastSenderDisp = dispinterface
    ['{92298BE3-ADEC-438F-800C-CF6311A7DF1D}']
    function Send(const pSmtp: IMail; nKey: Integer; const tParam: WideString): Integer; dispid 1;
    function Test(const pSmtp: IMail; nKey: Integer; const tParam: WideString): Integer; dispid 2;
    property MaxThreads: Integer dispid 3;
    function GetCurrentThreads: Integer; dispid 4;
    function GetQueuedCount: Integer; dispid 5;
    procedure ClearQueuedMails; dispid 6;
    procedure StopAllThreads; dispid 7;
    function GetIdleThreads: Integer; dispid 8;
    function SendByPickup(const PickupPath: WideString; const pSmtp: IMail; nKey: Integer; 
                          const tParam: WideString): Integer; dispid 9;
    function SendEmlFile(const FileName: WideString; const senderAddr: WideString; 
                         const recipientAddrs: WideString; nKey: Integer; const tParam: WideString; 
                         const RegisterKey: WideString): Integer; dispid 10;
    property ComputerName: WideString dispid 11;
    procedure LockEvent; dispid 12;
    procedure UnlockEvent; dispid 13;
    procedure ClearAllMails; dispid 14;
    procedure Pause; dispid 15;
    procedure Resume; dispid 16;
    property KeepConnection: Integer dispid 17;
  end;

// *********************************************************************//
// DispIntf:  _IMailEvents
// Flags:     (4096) Dispatchable
// GUID:      {68CB8B02-D4AA-4A16-97A0-6B9488F98189}
// *********************************************************************//
  _IMailEvents = dispinterface
    ['{68CB8B02-D4AA-4A16-97A0-6B9488F98189}']
    function OnClosed: HResult; dispid 1;
    function OnSending(lSent: Integer; lTotal: Integer): HResult; dispid 2;
    function OnError(lError: Integer; const ErrDescription: WideString): HResult; dispid 3;
    function OnConnected: HResult; dispid 4;
    function OnAuthenticated: HResult; dispid 5;
    function OnSendCommand(var Command: WideString): HResult; dispid 6;
    function OnServerRespond(var Response: WideString): HResult; dispid 7;
  end;

// *********************************************************************//
// DispIntf:  _IFastSenderEvents
// Flags:     (4096) Dispatchable
// GUID:      {A1B45F08-67E7-4276-A7CA-7664C08F9EF7}
// *********************************************************************//
  _IFastSenderEvents = dispinterface
    ['{A1B45F08-67E7-4276-A7CA-7664C08F9EF7}']
    function OnSent(lRet: Integer; const ErrDesc: WideString; nKey: Integer; 
                    const tParam: WideString; const senderAddr: WideString; 
                    const Recipients: WideString): HResult; dispid 1;
    function OnConnected(nKey: Integer; const tParam: WideString): HResult; dispid 2;
    function OnAuthenticated(nKey: Integer; const tParam: WideString): HResult; dispid 3;
    function OnSending(lSent: Integer; lTotal: Integer; nKey: Integer; const tParam: WideString): HResult; dispid 4;
  end;

// *********************************************************************//
// The Class CoMail provides a Create and CreateRemote method to          
// create instances of the default interface IMail exposed by              
// the CoClass Mail. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoMail = class
    class function Create: IMail;
    class function CreateRemote(const MachineName: string): IMail;
  end;

  TMailOnSending = procedure(ASender: TObject; lSent: Integer; lTotal: Integer) of object;
  TMailOnError = procedure(ASender: TObject; lError: Integer; const ErrDescription: WideString) of object;
  TMailOnSendCommand = procedure(ASender: TObject; var Command: WideString) of object;
  TMailOnServerRespond = procedure(ASender: TObject; var Response: WideString) of object;


// *********************************************************************//
// OLE Server Proxy class declaration
// Server Object    : TMail
// Help String      : Mail Class
// Default Interface: IMail
// Def. Intf. DISP? : No
// Event   Interface: _IMailEvents
// TypeFlags        : (2) CanCreate
// *********************************************************************//
  TMail = class(TOleServer)
  private
    FOnClosed: TNotifyEvent;
    FOnSending: TMailOnSending;
    FOnError: TMailOnError;
    FOnConnected: TNotifyEvent;
    FOnAuthenticated: TNotifyEvent;
    FOnSendCommand: TMailOnSendCommand;
    FOnServerRespond: TMailOnServerRespond;
    FAutoQuit: Boolean;
    FIntf: IMail;
    function GetDefaultInterface: IMail;
  protected
    procedure InitServerData; override;
    procedure InvokeEvent(DispID: TDispID; var Params: TVariantArray); override;
    function Get_BodyFormat: Integer;
    procedure Set_BodyFormat(pVal: Integer);
    function Get_BodyText: WideString;
    procedure Set_BodyText(const pVal: WideString);
    function Get_Charset: WideString;
    procedure Set_Charset(const pVal: WideString);
    function Get_From: WideString;
    procedure Set_From(const pVal: WideString);
    function Get_FromAddr: WideString;
    procedure Set_FromAddr(const pVal: WideString);
    function Get_LogFileName: WideString;
    procedure Set_LogFileName(const pVal: WideString);
    function Get_LicenseCode: WideString;
    procedure Set_LicenseCode(const pVal: WideString);
    function Get_ServerAddr: WideString;
    procedure Set_ServerAddr(const pVal: WideString);
    function Get_ServerPort: Integer;
    procedure Set_ServerPort(pVal: Integer);
    function Get_Subject: WideString;
    procedure Set_Subject(const pVal: WideString);
    function Get_ReplyTo: WideString;
    procedure Set_ReplyTo(const pVal: WideString);
    function Get_Priority: Integer;
    procedure Set_Priority(pVal: Integer);
    function Get_Timeout: Integer;
    procedure Set_Timeout(pVal: Integer);
    function Get_UserName: WideString;
    procedure Set_UserName(const pVal: WideString);
    function Get_Password: WideString;
    procedure Set_Password(const pVal: WideString);
    function Get_Version: WideString;
    function Get_Asynchronous: Integer;
    procedure Set_Asynchronous(pVal: Integer);
    function Get_AltBody: WideString;
    procedure Set_AltBody(const pVal: WideString);
    function Get_Anonymous: Integer;
    procedure Set_Anonymous(pVal: Integer);
    function Get_KeepConnection: Integer;
    procedure Set_KeepConnection(pVal: Integer);
    function Get_TransferEncoding: Integer;
    procedure Set_TransferEncoding(pVal: Integer);
    function Get_ComputerName: WideString;
    procedure Set_ComputerName(const pVal: WideString);
    function Get_BodyFormatEx: WideString;
    procedure Set_BodyFormatEx(const pVal: WideString);
    function Get_HeaderEncoding: Integer;
    procedure Set_HeaderEncoding(pVal: Integer);
    function Get_TryAllSmtpServers: Integer;
    procedure Set_TryAllSmtpServers(pVal: Integer);
    function Get_RawModeEnable: Integer;
    procedure Set_RawModeEnable(pVal: Integer);
    function Get_WrapEmailAddr: Integer;
    procedure Set_WrapEmailAddr(pVal: Integer);
    function Get_DeliveryNotification: Integer;
    procedure Set_DeliveryNotification(pVal: Integer);
    function Get__Idle: Integer;
    function Get_SSL_ignorecerterror: Integer;
    procedure Set_SSL_ignorecerterror(pVal: Integer);
    function Get_SSL_starttls: Integer;
    procedure Set_SSL_starttls(pVal: Integer);
    function Get_SSL_enabled: Integer;
    function Get_raw_Content: WideString;
    procedure Set_raw_Content(const pVal: WideString);
    function Get_LogLevel: Integer;
    procedure Set_LogLevel(pVal: Integer);
    function Get_SignerCert: ICertificate;
    function Get_RecipientsCerts: ICertificateCollection;
    function Get_ReturnPath: WideString;
    procedure Set_ReturnPath(const pVal: WideString);
    function Get_LocalIP: WideString;
    procedure Set_LocalIP(const pVal: WideString);
    function Get_AuthType: Integer;
    procedure Set_AuthType(pVal: Integer);
    function Get_SpecialFlags: Integer;
    procedure Set_SpecialFlags(pVal: Integer);
    function Get_DisplayTo: WideString;
    procedure Set_DisplayTo(const pVal: WideString);
    function Get_Date: TDateTime;
    procedure Set_Date(pVal: TDateTime);
    function Get_MessageID: WideString;
    procedure Set_MessageID(const pVal: WideString);
    function Get_NoWrapBody: Integer;
    procedure Set_NoWrapBody(pVal: Integer);
    function Get_EncryptionAlgorithm: Integer;
    procedure Set_EncryptionAlgorithm(pVal: Integer);
    function Get_SocksProxyServer: WideString;
    procedure Set_SocksProxyServer(const pVal: WideString);
    function Get_SocksProxyUser: WideString;
    procedure Set_SocksProxyUser(const pVal: WideString);
    function Get_SocksProxyPassword: WideString;
    procedure Set_SocksProxyPassword(const pVal: WideString);
    function Get_SocksProxyPort: Integer;
    procedure Set_SocksProxyPort(pVal: Integer);
    function Get_ProxyProtocol: Integer;
    procedure Set_ProxyProtocol(pVal: Integer);
    function Get_DK_PublicKey: WideString;
    function Get_ReadReceipt: WordBool;
    procedure Set_ReadReceipt(pVal: WordBool);
    function Get_Recipients: OleVariant;
    function Get_Style: Integer;
    procedure Set_Style(pVal: Integer);
    function Get_AutoCalendar: Integer;
    procedure Set_AutoCalendar(pVal: Integer);
    function Get_AttachmentCount: Integer;
    function Get_DnsServerIP: WideString;
    procedure Set_DnsServerIP(const pVal: WideString);
    function Get_Protocol: Integer;
    procedure Set_Protocol(pVal: Integer);
    function Get_Alias: WideString;
    procedure Set_Alias(const pVal: WideString);
    function Get_Drafts: WideString;
    procedure Set_Drafts(const pVal: WideString);
    function Get_Sender: WideString;
    procedure Set_Sender(const pVal: WideString);
    function Get_HttpProxyAuthType: Integer;
    procedure Set_HttpProxyAuthType(pVal: Integer);
    function Get_SMIMERFCCompatibility: WordBool;
    procedure Set_SMIMERFCCompatibility(pVal: WordBool);
    function Get_PIPELINING: Integer;
    procedure Set_PIPELINING(pVal: Integer);
    function Get_IgnoreDeliveryNotificationError: Integer;
    procedure Set_IgnoreDeliveryNotificationError(pVal: Integer);
    function Get_IPv6Policy: Integer;
    procedure Set_IPv6Policy(pVal: Integer);
    function Get_LocalIP6: WideString;
    procedure Set_LocalIP6(const pVal: WideString);
  public
    constructor Create(AOwner: TComponent); override;
    destructor  Destroy; override;
    procedure Connect; override;
    procedure ConnectTo(svrIntf: IMail);
    procedure Disconnect; override;
    function AddAttachment(const strFile: WideString): Integer;
    function AddRecipient(const strName: WideString; const strAddress: WideString; Flags: Integer): Integer;
    procedure ClearAttachment;
    procedure ClearRecipient;
    procedure ConvertHTML(Flags: Integer);
    function ImportMail(const strFile: WideString): Integer;
    procedure Reset;
    function SendMail: Integer;
    function AddAttachmentEx(const strFile: WideString; const strAlt: WideString): Integer;
    function AddInline(const strFile: WideString): WideString;
    function AddInlineEx(const strFile: WideString; const strAlt: WideString): WideString;
    procedure ClearInline;
    function SaveMail(const strFile: WideString): Integer;
    function AddHeader(const strHeader: WideString; const strValue: WideString): Integer;
    procedure ClearHeader;
    procedure Terminate;
    function GetLastError: Integer;
    function GetLastErrDescription: WideString;
    procedure SetMailer(const Mailer: WideString);
    function ImportMailEx(const strFile: WideString): Integer;
    function GetEmailServer(const EmailAddr: WideString): WideString;
    function AddRecipientEx(const AddressList: WideString; Flags: Integer): Integer;
    function AddAttachments(const sPath: WideString): Integer;
    function SaveMailEx(const PickupPath: WideString): Integer;
    function TestEmailAddr: Integer;
    function GetAllEmailServers(const EmailAddr: WideString): WideString;
    function GetEmailContent: WideString;
    function GetEmailHeaders: WideString;
    function GetAllRecipients: WideString;
    function GetSenderAddr: WideString;
    function CreateFolder(const FolderName: WideString): Integer;
    function DeleteFile(const FileName: WideString): Integer;
    function SSL_init: Integer;
    procedure SSL_uninit;
    procedure WriteLog(const LogContent: WideString);
    function ImportHtml(const html: WideString; const BasePath: WideString): Integer;
    function AddAttachment1(const FileName: WideString; Stream: OleVariant): Integer;
    procedure AppendBody(const BodyText: WideString; bAlt: Integer);
    function AddInline1(const FileName: WideString; Stream: OleVariant): WideString;
    function SendMailToQueue: Integer;
    procedure ClearHeaderEx(const HeaderName: WideString);
    function GetEmailChunk: OleVariant;
    function AddAttachmentCT(const FileName: WideString; const ContentType: WideString): Integer;
    function LoadMessage(const FileName: WideString): Integer;
    function LoadMessageChunk(newVal: OleVariant): Integer;
    procedure SetAttHeader(Index: Integer; const HeaderKey: WideString; 
                           const HeaderValue: WideString);
    function SendMailToQueueEx(const Instant: WideString): Integer;
    function LoadRawMessage(const FileName: WideString; Flag: Integer): Integer;
    procedure Quit;
    procedure Close;
    function PostToRemoteQueue(const Instance: WideString; const URL: WideString; 
                               const User: WideString; const Password: WideString): Integer;
    property DefaultInterface: IMail read GetDefaultInterface;
    property Version: WideString read Get_Version;
    property _Idle: Integer read Get__Idle;
    property SSL_enabled: Integer read Get_SSL_enabled;
    property SignerCert: ICertificate read Get_SignerCert;
    property RecipientsCerts: ICertificateCollection read Get_RecipientsCerts;
    property DK_PublicKey: WideString read Get_DK_PublicKey;
    property Recipients: OleVariant read Get_Recipients;
    property AttachmentCount: Integer read Get_AttachmentCount;
    property BodyFormat: Integer read Get_BodyFormat write Set_BodyFormat;
    property BodyText: WideString read Get_BodyText write Set_BodyText;
    property Charset: WideString read Get_Charset write Set_Charset;
    property From: WideString read Get_From write Set_From;
    property FromAddr: WideString read Get_FromAddr write Set_FromAddr;
    property LogFileName: WideString read Get_LogFileName write Set_LogFileName;
    property LicenseCode: WideString read Get_LicenseCode write Set_LicenseCode;
    property ServerAddr: WideString read Get_ServerAddr write Set_ServerAddr;
    property ServerPort: Integer read Get_ServerPort write Set_ServerPort;
    property Subject: WideString read Get_Subject write Set_Subject;
    property ReplyTo: WideString read Get_ReplyTo write Set_ReplyTo;
    property Priority: Integer read Get_Priority write Set_Priority;
    property Timeout: Integer read Get_Timeout write Set_Timeout;
    property UserName: WideString read Get_UserName write Set_UserName;
    property Password: WideString read Get_Password write Set_Password;
    property Asynchronous: Integer read Get_Asynchronous write Set_Asynchronous;
    property AltBody: WideString read Get_AltBody write Set_AltBody;
    property Anonymous: Integer read Get_Anonymous write Set_Anonymous;
    property KeepConnection: Integer read Get_KeepConnection write Set_KeepConnection;
    property TransferEncoding: Integer read Get_TransferEncoding write Set_TransferEncoding;
    property ComputerName: WideString read Get_ComputerName write Set_ComputerName;
    property BodyFormatEx: WideString read Get_BodyFormatEx write Set_BodyFormatEx;
    property HeaderEncoding: Integer read Get_HeaderEncoding write Set_HeaderEncoding;
    property TryAllSmtpServers: Integer read Get_TryAllSmtpServers write Set_TryAllSmtpServers;
    property RawModeEnable: Integer read Get_RawModeEnable write Set_RawModeEnable;
    property WrapEmailAddr: Integer read Get_WrapEmailAddr write Set_WrapEmailAddr;
    property DeliveryNotification: Integer read Get_DeliveryNotification write Set_DeliveryNotification;
    property SSL_ignorecerterror: Integer read Get_SSL_ignorecerterror write Set_SSL_ignorecerterror;
    property SSL_starttls: Integer read Get_SSL_starttls write Set_SSL_starttls;
    property raw_Content: WideString read Get_raw_Content write Set_raw_Content;
    property LogLevel: Integer read Get_LogLevel write Set_LogLevel;
    property ReturnPath: WideString read Get_ReturnPath write Set_ReturnPath;
    property LocalIP: WideString read Get_LocalIP write Set_LocalIP;
    property AuthType: Integer read Get_AuthType write Set_AuthType;
    property SpecialFlags: Integer read Get_SpecialFlags write Set_SpecialFlags;
    property DisplayTo: WideString read Get_DisplayTo write Set_DisplayTo;
    property Date: TDateTime read Get_Date write Set_Date;
    property MessageID: WideString read Get_MessageID write Set_MessageID;
    property NoWrapBody: Integer read Get_NoWrapBody write Set_NoWrapBody;
    property EncryptionAlgorithm: Integer read Get_EncryptionAlgorithm write Set_EncryptionAlgorithm;
    property SocksProxyServer: WideString read Get_SocksProxyServer write Set_SocksProxyServer;
    property SocksProxyUser: WideString read Get_SocksProxyUser write Set_SocksProxyUser;
    property SocksProxyPassword: WideString read Get_SocksProxyPassword write Set_SocksProxyPassword;
    property SocksProxyPort: Integer read Get_SocksProxyPort write Set_SocksProxyPort;
    property ProxyProtocol: Integer read Get_ProxyProtocol write Set_ProxyProtocol;
    property ReadReceipt: WordBool read Get_ReadReceipt write Set_ReadReceipt;
    property Style: Integer read Get_Style write Set_Style;
    property AutoCalendar: Integer read Get_AutoCalendar write Set_AutoCalendar;
    property DnsServerIP: WideString read Get_DnsServerIP write Set_DnsServerIP;
    property Protocol: Integer read Get_Protocol write Set_Protocol;
    property Alias: WideString read Get_Alias write Set_Alias;
    property Drafts: WideString read Get_Drafts write Set_Drafts;
    property Sender: WideString read Get_Sender write Set_Sender;
    property HttpProxyAuthType: Integer read Get_HttpProxyAuthType write Set_HttpProxyAuthType;
    property SMIMERFCCompatibility: WordBool read Get_SMIMERFCCompatibility write Set_SMIMERFCCompatibility;
    property PIPELINING: Integer read Get_PIPELINING write Set_PIPELINING;
    property IgnoreDeliveryNotificationError: Integer read Get_IgnoreDeliveryNotificationError write Set_IgnoreDeliveryNotificationError;
    property IPv6Policy: Integer read Get_IPv6Policy write Set_IPv6Policy;
    property LocalIP6: WideString read Get_LocalIP6 write Set_LocalIP6;
  published
    property AutoQuit: Boolean read FAutoQuit write FAutoQuit; 
    property OnClosed: TNotifyEvent read FOnClosed write FOnClosed;
    property OnSending: TMailOnSending read FOnSending write FOnSending;
    property OnError: TMailOnError read FOnError write FOnError;
    property OnConnected: TNotifyEvent read FOnConnected write FOnConnected;
    property OnAuthenticated: TNotifyEvent read FOnAuthenticated write FOnAuthenticated;
    property OnSendCommand: TMailOnSendCommand read FOnSendCommand write FOnSendCommand;
    property OnServerRespond: TMailOnServerRespond read FOnServerRespond write FOnServerRespond;
  end;

// *********************************************************************//
// The Class CoFastSender provides a Create and CreateRemote method to          
// create instances of the default interface IFastSender exposed by              
// the CoClass FastSender. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoFastSender = class
    class function Create: IFastSender;
    class function CreateRemote(const MachineName: string): IFastSender;
  end;

  TFastSenderOnSent = procedure(ASender: TObject; lRet: Integer; const ErrDesc: WideString; 
                                                  nKey: Integer; const tParam: WideString; 
                                                  const senderAddr: WideString; 
                                                  const Recipients: WideString) of object;
  TFastSenderOnConnected = procedure(ASender: TObject; nKey: Integer; const tParam: WideString) of object;
  TFastSenderOnAuthenticated = procedure(ASender: TObject; nKey: Integer; const tParam: WideString) of object;
  TFastSenderOnSending = procedure(ASender: TObject; lSent: Integer; lTotal: Integer; 
                                                     nKey: Integer; const tParam: WideString) of object;


// *********************************************************************//
// OLE Server Proxy class declaration
// Server Object    : TFastSender
// Help String      : FastSender Class
// Default Interface: IFastSender
// Def. Intf. DISP? : No
// Event   Interface: _IFastSenderEvents
// TypeFlags        : (2) CanCreate
// *********************************************************************//
  TFastSender = class(TOleServer)
  private
    FOnSent: TFastSenderOnSent;
    FOnConnected: TFastSenderOnConnected;
    FOnAuthenticated: TFastSenderOnAuthenticated;
    FOnSending: TFastSenderOnSending;
    FIntf: IFastSender;
    function GetDefaultInterface: IFastSender;
  protected
    procedure InitServerData; override;
    procedure InvokeEvent(DispID: TDispID; var Params: TVariantArray); override;
    function Get_MaxThreads: Integer;
    procedure Set_MaxThreads(pVal: Integer);
    function Get_ComputerName: WideString;
    procedure Set_ComputerName(const pVal: WideString);
    function Get_KeepConnection: Integer;
    procedure Set_KeepConnection(pVal: Integer);
  public
    constructor Create(AOwner: TComponent); override;
    destructor  Destroy; override;
    procedure Connect; override;
    procedure ConnectTo(svrIntf: IFastSender);
    procedure Disconnect; override;
    function Send(const pSmtp: IMail; nKey: Integer; const tParam: WideString): Integer;
    function Test(const pSmtp: IMail; nKey: Integer; const tParam: WideString): Integer;
    function GetCurrentThreads: Integer;
    function GetQueuedCount: Integer;
    procedure ClearQueuedMails;
    procedure StopAllThreads;
    function GetIdleThreads: Integer;
    function SendByPickup(const PickupPath: WideString; const pSmtp: IMail; nKey: Integer; 
                          const tParam: WideString): Integer;
    function SendEmlFile(const FileName: WideString; const senderAddr: WideString; 
                         const recipientAddrs: WideString; nKey: Integer; const tParam: WideString; 
                         const RegisterKey: WideString): Integer;
    procedure LockEvent;
    procedure UnlockEvent;
    procedure ClearAllMails;
    procedure Pause;
    procedure Resume;
    property DefaultInterface: IFastSender read GetDefaultInterface;
    property MaxThreads: Integer read Get_MaxThreads write Set_MaxThreads;
    property ComputerName: WideString read Get_ComputerName write Set_ComputerName;
    property KeepConnection: Integer read Get_KeepConnection write Set_KeepConnection;
  published
    property OnSent: TFastSenderOnSent read FOnSent write FOnSent;
    property OnConnected: TFastSenderOnConnected read FOnConnected write FOnConnected;
    property OnAuthenticated: TFastSenderOnAuthenticated read FOnAuthenticated write FOnAuthenticated;
    property OnSending: TFastSenderOnSending read FOnSending write FOnSending;
  end;

// *********************************************************************//
// The Class CoCertificate provides a Create and CreateRemote method to          
// create instances of the default interface ICertificate exposed by              
// the CoClass Certificate. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoCertificate = class
    class function Create: ICertificate;
    class function CreateRemote(const MachineName: string): ICertificate;
  end;


// *********************************************************************//
// OLE Server Proxy class declaration
// Server Object    : TCertificate
// Help String      : Certificate Class
// Default Interface: ICertificate
// Def. Intf. DISP? : No
// Event   Interface: 
// TypeFlags        : (2) CanCreate
// *********************************************************************//
  TCertificate = class(TOleServer)
  private
    FIntf: ICertificate;
    function GetDefaultInterface: ICertificate;
  protected
    procedure InitServerData; override;
    function Get_HasCertificate: WordBool;
    function Get_HasPrivateKey: WordBool;
  public
    constructor Create(AOwner: TComponent); override;
    destructor  Destroy; override;
    procedure Connect; override;
    procedure ConnectTo(svrIntf: ICertificate);
    procedure Disconnect; override;
    function FindSubject(const FindKey: WideString; StoreLocation: Integer; 
                         const StoreName: WideString): WordBool;
    function LoadPFX(PFXContent: OleVariant; const Password: WideString; KeyLocation: Integer): WordBool;
    function LoadPFXFromFile(const PFXFile: WideString; const Password: WideString; 
                             KeyLocation: Integer): WordBool;
    function LoadCert(CERTContent: OleVariant): WordBool;
    function LoadCertFromFile(const CERTFile: WideString): WordBool;
    procedure Unload;
    function GetLastError: WideString;
    property DefaultInterface: ICertificate read GetDefaultInterface;
    property HasCertificate: WordBool read Get_HasCertificate;
    property HasPrivateKey: WordBool read Get_HasPrivateKey;
  published
  end;

// *********************************************************************//
// The Class CoCertificateCollection provides a Create and CreateRemote method to          
// create instances of the default interface ICertificateCollection exposed by              
// the CoClass CertificateCollection. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoCertificateCollection = class
    class function Create: ICertificateCollection;
    class function CreateRemote(const MachineName: string): ICertificateCollection;
  end;


// *********************************************************************//
// OLE Server Proxy class declaration
// Server Object    : TCertificateCollection
// Help String      : CertificateCollection Class
// Default Interface: ICertificateCollection
// Def. Intf. DISP? : No
// Event   Interface: 
// TypeFlags        : (2) CanCreate
// *********************************************************************//
  TCertificateCollection = class(TOleServer)
  private
    FIntf: ICertificateCollection;
    function GetDefaultInterface: ICertificateCollection;
  protected
    procedure InitServerData; override;
    function Get_Count: Integer;
    function Get_HasEncryptCert: WordBool;
  public
    constructor Create(AOwner: TComponent); override;
    destructor  Destroy; override;
    procedure Connect; override;
    procedure ConnectTo(svrIntf: ICertificateCollection);
    procedure Disconnect; override;
    function Item(Index: Integer): ICertificate;
    procedure Add(const oCert: ICertificate);
    procedure Insert(Index: Integer; const oCert: ICertificate);
    procedure Clear;
    procedure RemoveAt(Index: Integer);
    function GetLastError: WideString;
    property DefaultInterface: ICertificateCollection read GetDefaultInterface;
    property Count: Integer read Get_Count;
    property HasEncryptCert: WordBool read Get_HasEncryptCert;
  published
  end;

procedure Register;

resourcestring
  dtlServerPage = 'ActiveX';

  dtlOcxPage = 'ActiveX';

implementation

uses System.Win.ComObj;

class function CoMail.Create: IMail;
begin
  Result := CreateComObject(CLASS_Mail) as IMail;
end;

class function CoMail.CreateRemote(const MachineName: string): IMail;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_Mail) as IMail;
end;

procedure TMail.InitServerData;
const
  CServerData: TServerData = (
    ClassID:   '{DF8A4FE2-221A-4504-987A-3FD4720DB929}';
    IntfIID:   '{1AD28FC9-0C71-4E89-85C9-CAECDE8BE3AB}';
    EventIID:  '{68CB8B02-D4AA-4A16-97A0-6B9488F98189}';
    LicenseKey: nil;
    Version: 500);
begin
  ServerData := @CServerData;
end;

procedure TMail.Connect;
var
  punk: IUnknown;
begin
  if FIntf = nil then
  begin
    punk := GetServer;
    ConnectEvents(punk);
    Fintf:= punk as IMail;
  end;
end;

procedure TMail.ConnectTo(svrIntf: IMail);
begin
  Disconnect;
  FIntf := svrIntf;
  ConnectEvents(FIntf);
end;

procedure TMail.DisConnect;
begin
  if Fintf <> nil then
  begin
    DisconnectEvents(FIntf);
    if FAutoQuit then
      Quit();
    FIntf := nil;
  end;
end;

function TMail.GetDefaultInterface: IMail;
begin
  if FIntf = nil then
    Connect;
  Assert(FIntf <> nil, 'DefaultInterface is NULL. Component is not connected to Server. You must call "Connect" or "ConnectTo" before this operation');
  Result := FIntf;
end;

constructor TMail.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
end;

destructor TMail.Destroy;
begin
  inherited Destroy;
end;

procedure TMail.InvokeEvent(DispID: TDispID; var Params: TVariantArray);
begin
  case DispID of
    -1: Exit;  // DISPID_UNKNOWN
    1: if Assigned(FOnClosed) then
         FOnClosed(Self);
    2: if Assigned(FOnSending) then
         FOnSending(Self,
                    Params[0] {Integer},
                    Params[1] {Integer});
    3: if Assigned(FOnError) then
         FOnError(Self,
                  Params[0] {Integer},
                  Params[1] {const WideString});
    4: if Assigned(FOnConnected) then
         FOnConnected(Self);
    5: if Assigned(FOnAuthenticated) then
         FOnAuthenticated(Self);
    6: if Assigned(FOnSendCommand) then
         FOnSendCommand(Self, WideString((TVarData(Params[0]).VPointer)^) {var WideString});
    7: if Assigned(FOnServerRespond) then
         FOnServerRespond(Self, WideString((TVarData(Params[0]).VPointer)^) {var WideString});
  end; {case DispID}
end;

function TMail.Get_BodyFormat: Integer;
begin
  Result := DefaultInterface.BodyFormat;
end;

procedure TMail.Set_BodyFormat(pVal: Integer);
begin
  DefaultInterface.BodyFormat := pVal;
end;

function TMail.Get_BodyText: WideString;
begin
  Result := DefaultInterface.BodyText;
end;

procedure TMail.Set_BodyText(const pVal: WideString);
begin
  DefaultInterface.BodyText := pVal;
end;

function TMail.Get_Charset: WideString;
begin
  Result := DefaultInterface.Charset;
end;

procedure TMail.Set_Charset(const pVal: WideString);
begin
  DefaultInterface.Charset := pVal;
end;

function TMail.Get_From: WideString;
begin
  Result := DefaultInterface.From;
end;

procedure TMail.Set_From(const pVal: WideString);
begin
  DefaultInterface.From := pVal;
end;

function TMail.Get_FromAddr: WideString;
begin
  Result := DefaultInterface.FromAddr;
end;

procedure TMail.Set_FromAddr(const pVal: WideString);
begin
  DefaultInterface.FromAddr := pVal;
end;

function TMail.Get_LogFileName: WideString;
begin
  Result := DefaultInterface.LogFileName;
end;

procedure TMail.Set_LogFileName(const pVal: WideString);
begin
  DefaultInterface.LogFileName := pVal;
end;

function TMail.Get_LicenseCode: WideString;
begin
  Result := DefaultInterface.LicenseCode;
end;

procedure TMail.Set_LicenseCode(const pVal: WideString);
begin
  DefaultInterface.LicenseCode := pVal;
end;

function TMail.Get_ServerAddr: WideString;
begin
  Result := DefaultInterface.ServerAddr;
end;

procedure TMail.Set_ServerAddr(const pVal: WideString);
begin
  DefaultInterface.ServerAddr := pVal;
end;

function TMail.Get_ServerPort: Integer;
begin
  Result := DefaultInterface.ServerPort;
end;

procedure TMail.Set_ServerPort(pVal: Integer);
begin
  DefaultInterface.ServerPort := pVal;
end;

function TMail.Get_Subject: WideString;
begin
  Result := DefaultInterface.Subject;
end;

procedure TMail.Set_Subject(const pVal: WideString);
begin
  DefaultInterface.Subject := pVal;
end;

function TMail.Get_ReplyTo: WideString;
begin
  Result := DefaultInterface.ReplyTo;
end;

procedure TMail.Set_ReplyTo(const pVal: WideString);
begin
  DefaultInterface.ReplyTo := pVal;
end;

function TMail.Get_Priority: Integer;
begin
  Result := DefaultInterface.Priority;
end;

procedure TMail.Set_Priority(pVal: Integer);
begin
  DefaultInterface.Priority := pVal;
end;

function TMail.Get_Timeout: Integer;
begin
  Result := DefaultInterface.Timeout;
end;

procedure TMail.Set_Timeout(pVal: Integer);
begin
  DefaultInterface.Timeout := pVal;
end;

function TMail.Get_UserName: WideString;
begin
  Result := DefaultInterface.UserName;
end;

procedure TMail.Set_UserName(const pVal: WideString);
begin
  DefaultInterface.UserName := pVal;
end;

function TMail.Get_Password: WideString;
begin
  Result := DefaultInterface.Password;
end;

procedure TMail.Set_Password(const pVal: WideString);
begin
  DefaultInterface.Password := pVal;
end;

function TMail.Get_Version: WideString;
begin
  Result := DefaultInterface.Version;
end;

function TMail.Get_Asynchronous: Integer;
begin
  Result := DefaultInterface.Asynchronous;
end;

procedure TMail.Set_Asynchronous(pVal: Integer);
begin
  DefaultInterface.Asynchronous := pVal;
end;

function TMail.Get_AltBody: WideString;
begin
  Result := DefaultInterface.AltBody;
end;

procedure TMail.Set_AltBody(const pVal: WideString);
begin
  DefaultInterface.AltBody := pVal;
end;

function TMail.Get_Anonymous: Integer;
begin
  Result := DefaultInterface.Anonymous;
end;

procedure TMail.Set_Anonymous(pVal: Integer);
begin
  DefaultInterface.Anonymous := pVal;
end;

function TMail.Get_KeepConnection: Integer;
begin
  Result := DefaultInterface.KeepConnection;
end;

procedure TMail.Set_KeepConnection(pVal: Integer);
begin
  DefaultInterface.KeepConnection := pVal;
end;

function TMail.Get_TransferEncoding: Integer;
begin
  Result := DefaultInterface.TransferEncoding;
end;

procedure TMail.Set_TransferEncoding(pVal: Integer);
begin
  DefaultInterface.TransferEncoding := pVal;
end;

function TMail.Get_ComputerName: WideString;
begin
  Result := DefaultInterface.ComputerName;
end;

procedure TMail.Set_ComputerName(const pVal: WideString);
begin
  DefaultInterface.ComputerName := pVal;
end;

function TMail.Get_BodyFormatEx: WideString;
begin
  Result := DefaultInterface.BodyFormatEx;
end;

procedure TMail.Set_BodyFormatEx(const pVal: WideString);
begin
  DefaultInterface.BodyFormatEx := pVal;
end;

function TMail.Get_HeaderEncoding: Integer;
begin
  Result := DefaultInterface.HeaderEncoding;
end;

procedure TMail.Set_HeaderEncoding(pVal: Integer);
begin
  DefaultInterface.HeaderEncoding := pVal;
end;

function TMail.Get_TryAllSmtpServers: Integer;
begin
  Result := DefaultInterface.TryAllSmtpServers;
end;

procedure TMail.Set_TryAllSmtpServers(pVal: Integer);
begin
  DefaultInterface.TryAllSmtpServers := pVal;
end;

function TMail.Get_RawModeEnable: Integer;
begin
  Result := DefaultInterface.RawModeEnable;
end;

procedure TMail.Set_RawModeEnable(pVal: Integer);
begin
  DefaultInterface.RawModeEnable := pVal;
end;

function TMail.Get_WrapEmailAddr: Integer;
begin
  Result := DefaultInterface.WrapEmailAddr;
end;

procedure TMail.Set_WrapEmailAddr(pVal: Integer);
begin
  DefaultInterface.WrapEmailAddr := pVal;
end;

function TMail.Get_DeliveryNotification: Integer;
begin
  Result := DefaultInterface.DeliveryNotification;
end;

procedure TMail.Set_DeliveryNotification(pVal: Integer);
begin
  DefaultInterface.DeliveryNotification := pVal;
end;

function TMail.Get__Idle: Integer;
begin
  Result := DefaultInterface._Idle;
end;

function TMail.Get_SSL_ignorecerterror: Integer;
begin
  Result := DefaultInterface.SSL_ignorecerterror;
end;

procedure TMail.Set_SSL_ignorecerterror(pVal: Integer);
begin
  DefaultInterface.SSL_ignorecerterror := pVal;
end;

function TMail.Get_SSL_starttls: Integer;
begin
  Result := DefaultInterface.SSL_starttls;
end;

procedure TMail.Set_SSL_starttls(pVal: Integer);
begin
  DefaultInterface.SSL_starttls := pVal;
end;

function TMail.Get_SSL_enabled: Integer;
begin
  Result := DefaultInterface.SSL_enabled;
end;

function TMail.Get_raw_Content: WideString;
begin
  Result := DefaultInterface.raw_Content;
end;

procedure TMail.Set_raw_Content(const pVal: WideString);
begin
  DefaultInterface.raw_Content := pVal;
end;

function TMail.Get_LogLevel: Integer;
begin
  Result := DefaultInterface.LogLevel;
end;

procedure TMail.Set_LogLevel(pVal: Integer);
begin
  DefaultInterface.LogLevel := pVal;
end;

function TMail.Get_SignerCert: ICertificate;
begin
  Result := DefaultInterface.SignerCert;
end;

function TMail.Get_RecipientsCerts: ICertificateCollection;
begin
  Result := DefaultInterface.RecipientsCerts;
end;

function TMail.Get_ReturnPath: WideString;
begin
  Result := DefaultInterface.ReturnPath;
end;

procedure TMail.Set_ReturnPath(const pVal: WideString);
begin
  DefaultInterface.ReturnPath := pVal;
end;

function TMail.Get_LocalIP: WideString;
begin
  Result := DefaultInterface.LocalIP;
end;

procedure TMail.Set_LocalIP(const pVal: WideString);
begin
  DefaultInterface.LocalIP := pVal;
end;

function TMail.Get_AuthType: Integer;
begin
  Result := DefaultInterface.AuthType;
end;

procedure TMail.Set_AuthType(pVal: Integer);
begin
  DefaultInterface.AuthType := pVal;
end;

function TMail.Get_SpecialFlags: Integer;
begin
  Result := DefaultInterface.SpecialFlags;
end;

procedure TMail.Set_SpecialFlags(pVal: Integer);
begin
  DefaultInterface.SpecialFlags := pVal;
end;

function TMail.Get_DisplayTo: WideString;
begin
  Result := DefaultInterface.DisplayTo;
end;

procedure TMail.Set_DisplayTo(const pVal: WideString);
begin
  DefaultInterface.DisplayTo := pVal;
end;

function TMail.Get_Date: TDateTime;
begin
  Result := DefaultInterface.Date;
end;

procedure TMail.Set_Date(pVal: TDateTime);
begin
  DefaultInterface.Date := pVal;
end;

function TMail.Get_MessageID: WideString;
begin
  Result := DefaultInterface.MessageID;
end;

procedure TMail.Set_MessageID(const pVal: WideString);
begin
  DefaultInterface.MessageID := pVal;
end;

function TMail.Get_NoWrapBody: Integer;
begin
  Result := DefaultInterface.NoWrapBody;
end;

procedure TMail.Set_NoWrapBody(pVal: Integer);
begin
  DefaultInterface.NoWrapBody := pVal;
end;

function TMail.Get_EncryptionAlgorithm: Integer;
begin
  Result := DefaultInterface.EncryptionAlgorithm;
end;

procedure TMail.Set_EncryptionAlgorithm(pVal: Integer);
begin
  DefaultInterface.EncryptionAlgorithm := pVal;
end;

function TMail.Get_SocksProxyServer: WideString;
begin
  Result := DefaultInterface.SocksProxyServer;
end;

procedure TMail.Set_SocksProxyServer(const pVal: WideString);
begin
  DefaultInterface.SocksProxyServer := pVal;
end;

function TMail.Get_SocksProxyUser: WideString;
begin
  Result := DefaultInterface.SocksProxyUser;
end;

procedure TMail.Set_SocksProxyUser(const pVal: WideString);
begin
  DefaultInterface.SocksProxyUser := pVal;
end;

function TMail.Get_SocksProxyPassword: WideString;
begin
  Result := DefaultInterface.SocksProxyPassword;
end;

procedure TMail.Set_SocksProxyPassword(const pVal: WideString);
begin
  DefaultInterface.SocksProxyPassword := pVal;
end;

function TMail.Get_SocksProxyPort: Integer;
begin
  Result := DefaultInterface.SocksProxyPort;
end;

procedure TMail.Set_SocksProxyPort(pVal: Integer);
begin
  DefaultInterface.SocksProxyPort := pVal;
end;

function TMail.Get_ProxyProtocol: Integer;
begin
  Result := DefaultInterface.ProxyProtocol;
end;

procedure TMail.Set_ProxyProtocol(pVal: Integer);
begin
  DefaultInterface.ProxyProtocol := pVal;
end;

function TMail.Get_DK_PublicKey: WideString;
begin
  Result := DefaultInterface.DK_PublicKey;
end;

function TMail.Get_ReadReceipt: WordBool;
begin
  Result := DefaultInterface.ReadReceipt;
end;

procedure TMail.Set_ReadReceipt(pVal: WordBool);
begin
  DefaultInterface.ReadReceipt := pVal;
end;

function TMail.Get_Recipients: OleVariant;
begin
  Result := DefaultInterface.Recipients;
end;

function TMail.Get_Style: Integer;
begin
  Result := DefaultInterface.Style;
end;

procedure TMail.Set_Style(pVal: Integer);
begin
  DefaultInterface.Style := pVal;
end;

function TMail.Get_AutoCalendar: Integer;
begin
  Result := DefaultInterface.AutoCalendar;
end;

procedure TMail.Set_AutoCalendar(pVal: Integer);
begin
  DefaultInterface.AutoCalendar := pVal;
end;

function TMail.Get_AttachmentCount: Integer;
begin
  Result := DefaultInterface.AttachmentCount;
end;

function TMail.Get_DnsServerIP: WideString;
begin
  Result := DefaultInterface.DnsServerIP;
end;

procedure TMail.Set_DnsServerIP(const pVal: WideString);
begin
  DefaultInterface.DnsServerIP := pVal;
end;

function TMail.Get_Protocol: Integer;
begin
  Result := DefaultInterface.Protocol;
end;

procedure TMail.Set_Protocol(pVal: Integer);
begin
  DefaultInterface.Protocol := pVal;
end;

function TMail.Get_Alias: WideString;
begin
  Result := DefaultInterface.Alias;
end;

procedure TMail.Set_Alias(const pVal: WideString);
begin
  DefaultInterface.Alias := pVal;
end;

function TMail.Get_Drafts: WideString;
begin
  Result := DefaultInterface.Drafts;
end;

procedure TMail.Set_Drafts(const pVal: WideString);
begin
  DefaultInterface.Drafts := pVal;
end;

function TMail.Get_Sender: WideString;
begin
  Result := DefaultInterface.Sender;
end;

procedure TMail.Set_Sender(const pVal: WideString);
begin
  DefaultInterface.Sender := pVal;
end;

function TMail.Get_HttpProxyAuthType: Integer;
begin
  Result := DefaultInterface.HttpProxyAuthType;
end;

procedure TMail.Set_HttpProxyAuthType(pVal: Integer);
begin
  DefaultInterface.HttpProxyAuthType := pVal;
end;

function TMail.Get_SMIMERFCCompatibility: WordBool;
begin
  Result := DefaultInterface.SMIMERFCCompatibility;
end;

procedure TMail.Set_SMIMERFCCompatibility(pVal: WordBool);
begin
  DefaultInterface.SMIMERFCCompatibility := pVal;
end;

function TMail.Get_PIPELINING: Integer;
begin
  Result := DefaultInterface.PIPELINING;
end;

procedure TMail.Set_PIPELINING(pVal: Integer);
begin
  DefaultInterface.PIPELINING := pVal;
end;

function TMail.Get_IgnoreDeliveryNotificationError: Integer;
begin
  Result := DefaultInterface.IgnoreDeliveryNotificationError;
end;

procedure TMail.Set_IgnoreDeliveryNotificationError(pVal: Integer);
begin
  DefaultInterface.IgnoreDeliveryNotificationError := pVal;
end;

function TMail.Get_IPv6Policy: Integer;
begin
  Result := DefaultInterface.IPv6Policy;
end;

procedure TMail.Set_IPv6Policy(pVal: Integer);
begin
  DefaultInterface.IPv6Policy := pVal;
end;

function TMail.Get_LocalIP6: WideString;
begin
  Result := DefaultInterface.LocalIP6;
end;

procedure TMail.Set_LocalIP6(const pVal: WideString);
begin
  DefaultInterface.LocalIP6 := pVal;
end;

function TMail.AddAttachment(const strFile: WideString): Integer;
begin
  Result := DefaultInterface.AddAttachment(strFile);
end;

function TMail.AddRecipient(const strName: WideString; const strAddress: WideString; Flags: Integer): Integer;
begin
  Result := DefaultInterface.AddRecipient(strName, strAddress, Flags);
end;

procedure TMail.ClearAttachment;
begin
  DefaultInterface.ClearAttachment;
end;

procedure TMail.ClearRecipient;
begin
  DefaultInterface.ClearRecipient;
end;

procedure TMail.ConvertHTML(Flags: Integer);
begin
  DefaultInterface.ConvertHTML(Flags);
end;

function TMail.ImportMail(const strFile: WideString): Integer;
begin
  Result := DefaultInterface.ImportMail(strFile);
end;

procedure TMail.Reset;
begin
  DefaultInterface.Reset;
end;

function TMail.SendMail: Integer;
begin
  Result := DefaultInterface.SendMail;
end;

function TMail.AddAttachmentEx(const strFile: WideString; const strAlt: WideString): Integer;
begin
  Result := DefaultInterface.AddAttachmentEx(strFile, strAlt);
end;

function TMail.AddInline(const strFile: WideString): WideString;
begin
  Result := DefaultInterface.AddInline(strFile);
end;

function TMail.AddInlineEx(const strFile: WideString; const strAlt: WideString): WideString;
begin
  Result := DefaultInterface.AddInlineEx(strFile, strAlt);
end;

procedure TMail.ClearInline;
begin
  DefaultInterface.ClearInline;
end;

function TMail.SaveMail(const strFile: WideString): Integer;
begin
  Result := DefaultInterface.SaveMail(strFile);
end;

function TMail.AddHeader(const strHeader: WideString; const strValue: WideString): Integer;
begin
  Result := DefaultInterface.AddHeader(strHeader, strValue);
end;

procedure TMail.ClearHeader;
begin
  DefaultInterface.ClearHeader;
end;

procedure TMail.Terminate;
begin
  DefaultInterface.Terminate;
end;

function TMail.GetLastError: Integer;
begin
  Result := DefaultInterface.GetLastError;
end;

function TMail.GetLastErrDescription: WideString;
begin
  Result := DefaultInterface.GetLastErrDescription;
end;

procedure TMail.SetMailer(const Mailer: WideString);
begin
  DefaultInterface.SetMailer(Mailer);
end;

function TMail.ImportMailEx(const strFile: WideString): Integer;
begin
  Result := DefaultInterface.ImportMailEx(strFile);
end;

function TMail.GetEmailServer(const EmailAddr: WideString): WideString;
begin
  Result := DefaultInterface.GetEmailServer(EmailAddr);
end;

function TMail.AddRecipientEx(const AddressList: WideString; Flags: Integer): Integer;
begin
  Result := DefaultInterface.AddRecipientEx(AddressList, Flags);
end;

function TMail.AddAttachments(const sPath: WideString): Integer;
begin
  Result := DefaultInterface.AddAttachments(sPath);
end;

function TMail.SaveMailEx(const PickupPath: WideString): Integer;
begin
  Result := DefaultInterface.SaveMailEx(PickupPath);
end;

function TMail.TestEmailAddr: Integer;
begin
  Result := DefaultInterface.TestEmailAddr;
end;

function TMail.GetAllEmailServers(const EmailAddr: WideString): WideString;
begin
  Result := DefaultInterface.GetAllEmailServers(EmailAddr);
end;

function TMail.GetEmailContent: WideString;
begin
  Result := DefaultInterface.GetEmailContent;
end;

function TMail.GetEmailHeaders: WideString;
begin
  Result := DefaultInterface.GetEmailHeaders;
end;

function TMail.GetAllRecipients: WideString;
begin
  Result := DefaultInterface.GetAllRecipients;
end;

function TMail.GetSenderAddr: WideString;
begin
  Result := DefaultInterface.GetSenderAddr;
end;

function TMail.CreateFolder(const FolderName: WideString): Integer;
begin
  Result := DefaultInterface.CreateFolder(FolderName);
end;

function TMail.DeleteFile(const FileName: WideString): Integer;
begin
  Result := DefaultInterface.DeleteFile(FileName);
end;

function TMail.SSL_init: Integer;
begin
  Result := DefaultInterface.SSL_init;
end;

procedure TMail.SSL_uninit;
begin
  DefaultInterface.SSL_uninit;
end;

procedure TMail.WriteLog(const LogContent: WideString);
begin
  DefaultInterface.WriteLog(LogContent);
end;

function TMail.ImportHtml(const html: WideString; const BasePath: WideString): Integer;
begin
  Result := DefaultInterface.ImportHtml(html, BasePath);
end;

function TMail.AddAttachment1(const FileName: WideString; Stream: OleVariant): Integer;
begin
  Result := DefaultInterface.AddAttachment1(FileName, Stream);
end;

procedure TMail.AppendBody(const BodyText: WideString; bAlt: Integer);
begin
  DefaultInterface.AppendBody(BodyText, bAlt);
end;

function TMail.AddInline1(const FileName: WideString; Stream: OleVariant): WideString;
begin
  Result := DefaultInterface.AddInline1(FileName, Stream);
end;

function TMail.SendMailToQueue: Integer;
begin
  Result := DefaultInterface.SendMailToQueue;
end;

procedure TMail.ClearHeaderEx(const HeaderName: WideString);
begin
  DefaultInterface.ClearHeaderEx(HeaderName);
end;

function TMail.GetEmailChunk: OleVariant;
begin
  Result := DefaultInterface.GetEmailChunk;
end;

function TMail.AddAttachmentCT(const FileName: WideString; const ContentType: WideString): Integer;
begin
  Result := DefaultInterface.AddAttachmentCT(FileName, ContentType);
end;

function TMail.LoadMessage(const FileName: WideString): Integer;
begin
  Result := DefaultInterface.LoadMessage(FileName);
end;

function TMail.LoadMessageChunk(newVal: OleVariant): Integer;
begin
  Result := DefaultInterface.LoadMessageChunk(newVal);
end;

procedure TMail.SetAttHeader(Index: Integer; const HeaderKey: WideString; 
                             const HeaderValue: WideString);
begin
  DefaultInterface.SetAttHeader(Index, HeaderKey, HeaderValue);
end;

function TMail.SendMailToQueueEx(const Instant: WideString): Integer;
begin
  Result := DefaultInterface.SendMailToQueueEx(Instant);
end;

function TMail.LoadRawMessage(const FileName: WideString; Flag: Integer): Integer;
begin
  Result := DefaultInterface.LoadRawMessage(FileName, Flag);
end;

procedure TMail.Quit;
begin
  DefaultInterface.Quit;
end;

procedure TMail.Close;
begin
  DefaultInterface.Close;
end;

function TMail.PostToRemoteQueue(const Instance: WideString; const URL: WideString; 
                                 const User: WideString; const Password: WideString): Integer;
begin
  Result := DefaultInterface.PostToRemoteQueue(Instance, URL, User, Password);
end;

class function CoFastSender.Create: IFastSender;
begin
  Result := CreateComObject(CLASS_FastSender) as IFastSender;
end;

class function CoFastSender.CreateRemote(const MachineName: string): IFastSender;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_FastSender) as IFastSender;
end;

procedure TFastSender.InitServerData;
const
  CServerData: TServerData = (
    ClassID:   '{FF80631D-E750-4C67-AFC3-5170AB72518B}';
    IntfIID:   '{92298BE3-ADEC-438F-800C-CF6311A7DF1D}';
    EventIID:  '{A1B45F08-67E7-4276-A7CA-7664C08F9EF7}';
    LicenseKey: nil;
    Version: 500);
begin
  ServerData := @CServerData;
end;

procedure TFastSender.Connect;
var
  punk: IUnknown;
begin
  if FIntf = nil then
  begin
    punk := GetServer;
    ConnectEvents(punk);
    Fintf:= punk as IFastSender;
  end;
end;

procedure TFastSender.ConnectTo(svrIntf: IFastSender);
begin
  Disconnect;
  FIntf := svrIntf;
  ConnectEvents(FIntf);
end;

procedure TFastSender.DisConnect;
begin
  if Fintf <> nil then
  begin
    DisconnectEvents(FIntf);
    FIntf := nil;
  end;
end;

function TFastSender.GetDefaultInterface: IFastSender;
begin
  if FIntf = nil then
    Connect;
  Assert(FIntf <> nil, 'DefaultInterface is NULL. Component is not connected to Server. You must call "Connect" or "ConnectTo" before this operation');
  Result := FIntf;
end;

constructor TFastSender.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
end;

destructor TFastSender.Destroy;
begin
  inherited Destroy;
end;

procedure TFastSender.InvokeEvent(DispID: TDispID; var Params: TVariantArray);
begin
  case DispID of
    -1: Exit;  // DISPID_UNKNOWN
    1: if Assigned(FOnSent) then
         FOnSent(Self,
                 Params[0] {Integer},
                 Params[1] {const WideString},
                 Params[2] {Integer},
                 Params[3] {const WideString},
                 Params[4] {const WideString},
                 Params[5] {const WideString});
    2: if Assigned(FOnConnected) then
         FOnConnected(Self,
                      Params[0] {Integer},
                      Params[1] {const WideString});
    3: if Assigned(FOnAuthenticated) then
         FOnAuthenticated(Self,
                          Params[0] {Integer},
                          Params[1] {const WideString});
    4: if Assigned(FOnSending) then
         FOnSending(Self,
                    Params[0] {Integer},
                    Params[1] {Integer},
                    Params[2] {Integer},
                    Params[3] {const WideString});
  end; {case DispID}
end;

function TFastSender.Get_MaxThreads: Integer;
begin
  Result := DefaultInterface.MaxThreads;
end;

procedure TFastSender.Set_MaxThreads(pVal: Integer);
begin
  DefaultInterface.MaxThreads := pVal;
end;

function TFastSender.Get_ComputerName: WideString;
begin
  Result := DefaultInterface.ComputerName;
end;

procedure TFastSender.Set_ComputerName(const pVal: WideString);
begin
  DefaultInterface.ComputerName := pVal;
end;

function TFastSender.Get_KeepConnection: Integer;
begin
  Result := DefaultInterface.KeepConnection;
end;

procedure TFastSender.Set_KeepConnection(pVal: Integer);
begin
  DefaultInterface.KeepConnection := pVal;
end;

function TFastSender.Send(const pSmtp: IMail; nKey: Integer; const tParam: WideString): Integer;
begin
  Result := DefaultInterface.Send(pSmtp, nKey, tParam);
end;

function TFastSender.Test(const pSmtp: IMail; nKey: Integer; const tParam: WideString): Integer;
begin
  Result := DefaultInterface.Test(pSmtp, nKey, tParam);
end;

function TFastSender.GetCurrentThreads: Integer;
begin
  Result := DefaultInterface.GetCurrentThreads;
end;

function TFastSender.GetQueuedCount: Integer;
begin
  Result := DefaultInterface.GetQueuedCount;
end;

procedure TFastSender.ClearQueuedMails;
begin
  DefaultInterface.ClearQueuedMails;
end;

procedure TFastSender.StopAllThreads;
begin
  DefaultInterface.StopAllThreads;
end;

function TFastSender.GetIdleThreads: Integer;
begin
  Result := DefaultInterface.GetIdleThreads;
end;

function TFastSender.SendByPickup(const PickupPath: WideString; const pSmtp: IMail; nKey: Integer; 
                                  const tParam: WideString): Integer;
begin
  Result := DefaultInterface.SendByPickup(PickupPath, pSmtp, nKey, tParam);
end;

function TFastSender.SendEmlFile(const FileName: WideString; const senderAddr: WideString; 
                                 const recipientAddrs: WideString; nKey: Integer; 
                                 const tParam: WideString; const RegisterKey: WideString): Integer;
begin
  Result := DefaultInterface.SendEmlFile(FileName, senderAddr, recipientAddrs, nKey, tParam, 
                                         RegisterKey);
end;

procedure TFastSender.LockEvent;
begin
  DefaultInterface.LockEvent;
end;

procedure TFastSender.UnlockEvent;
begin
  DefaultInterface.UnlockEvent;
end;

procedure TFastSender.ClearAllMails;
begin
  DefaultInterface.ClearAllMails;
end;

procedure TFastSender.Pause;
begin
  DefaultInterface.Pause;
end;

procedure TFastSender.Resume;
begin
  DefaultInterface.Resume;
end;

class function CoCertificate.Create: ICertificate;
begin
  Result := CreateComObject(CLASS_Certificate) as ICertificate;
end;

class function CoCertificate.CreateRemote(const MachineName: string): ICertificate;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_Certificate) as ICertificate;
end;

procedure TCertificate.InitServerData;
const
  CServerData: TServerData = (
    ClassID:   '{EAFC4EAA-9390-492A-8E53-E179527780F6}';
    IntfIID:   '{A2809780-C98E-4C6D-A552-DAB146D4AD12}';
    EventIID:  '';
    LicenseKey: nil;
    Version: 500);
begin
  ServerData := @CServerData;
end;

procedure TCertificate.Connect;
var
  punk: IUnknown;
begin
  if FIntf = nil then
  begin
    punk := GetServer;
    Fintf:= punk as ICertificate;
  end;
end;

procedure TCertificate.ConnectTo(svrIntf: ICertificate);
begin
  Disconnect;
  FIntf := svrIntf;
end;

procedure TCertificate.DisConnect;
begin
  if Fintf <> nil then
  begin
    FIntf := nil;
  end;
end;

function TCertificate.GetDefaultInterface: ICertificate;
begin
  if FIntf = nil then
    Connect;
  Assert(FIntf <> nil, 'DefaultInterface is NULL. Component is not connected to Server. You must call "Connect" or "ConnectTo" before this operation');
  Result := FIntf;
end;

constructor TCertificate.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
end;

destructor TCertificate.Destroy;
begin
  inherited Destroy;
end;

function TCertificate.Get_HasCertificate: WordBool;
begin
  Result := DefaultInterface.HasCertificate;
end;

function TCertificate.Get_HasPrivateKey: WordBool;
begin
  Result := DefaultInterface.HasPrivateKey;
end;

function TCertificate.FindSubject(const FindKey: WideString; StoreLocation: Integer; 
                                  const StoreName: WideString): WordBool;
begin
  Result := DefaultInterface.FindSubject(FindKey, StoreLocation, StoreName);
end;

function TCertificate.LoadPFX(PFXContent: OleVariant; const Password: WideString; 
                              KeyLocation: Integer): WordBool;
begin
  Result := DefaultInterface.LoadPFX(PFXContent, Password, KeyLocation);
end;

function TCertificate.LoadPFXFromFile(const PFXFile: WideString; const Password: WideString; 
                                      KeyLocation: Integer): WordBool;
begin
  Result := DefaultInterface.LoadPFXFromFile(PFXFile, Password, KeyLocation);
end;

function TCertificate.LoadCert(CERTContent: OleVariant): WordBool;
begin
  Result := DefaultInterface.LoadCert(CERTContent);
end;

function TCertificate.LoadCertFromFile(const CERTFile: WideString): WordBool;
begin
  Result := DefaultInterface.LoadCertFromFile(CERTFile);
end;

procedure TCertificate.Unload;
begin
  DefaultInterface.Unload;
end;

function TCertificate.GetLastError: WideString;
begin
  Result := DefaultInterface.GetLastError;
end;

class function CoCertificateCollection.Create: ICertificateCollection;
begin
  Result := CreateComObject(CLASS_CertificateCollection) as ICertificateCollection;
end;

class function CoCertificateCollection.CreateRemote(const MachineName: string): ICertificateCollection;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_CertificateCollection) as ICertificateCollection;
end;

procedure TCertificateCollection.InitServerData;
const
  CServerData: TServerData = (
    ClassID:   '{036C2F8C-8D3C-4F4B-9B36-3B6F1D29C0B4}';
    IntfIID:   '{DC8D5635-B8E7-441E-B550-CE1BF3BA5C55}';
    EventIID:  '';
    LicenseKey: nil;
    Version: 500);
begin
  ServerData := @CServerData;
end;

procedure TCertificateCollection.Connect;
var
  punk: IUnknown;
begin
  if FIntf = nil then
  begin
    punk := GetServer;
    Fintf:= punk as ICertificateCollection;
  end;
end;

procedure TCertificateCollection.ConnectTo(svrIntf: ICertificateCollection);
begin
  Disconnect;
  FIntf := svrIntf;
end;

procedure TCertificateCollection.DisConnect;
begin
  if Fintf <> nil then
  begin
    FIntf := nil;
  end;
end;

function TCertificateCollection.GetDefaultInterface: ICertificateCollection;
begin
  if FIntf = nil then
    Connect;
  Assert(FIntf <> nil, 'DefaultInterface is NULL. Component is not connected to Server. You must call "Connect" or "ConnectTo" before this operation');
  Result := FIntf;
end;

constructor TCertificateCollection.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
end;

destructor TCertificateCollection.Destroy;
begin
  inherited Destroy;
end;

function TCertificateCollection.Get_Count: Integer;
begin
  Result := DefaultInterface.Count;
end;

function TCertificateCollection.Get_HasEncryptCert: WordBool;
begin
  Result := DefaultInterface.HasEncryptCert;
end;

function TCertificateCollection.Item(Index: Integer): ICertificate;
begin
  Result := DefaultInterface.Item(Index);
end;

procedure TCertificateCollection.Add(const oCert: ICertificate);
begin
  DefaultInterface.Add(oCert);
end;

procedure TCertificateCollection.Insert(Index: Integer; const oCert: ICertificate);
begin
  DefaultInterface.Insert(Index, oCert);
end;

procedure TCertificateCollection.Clear;
begin
  DefaultInterface.Clear;
end;

procedure TCertificateCollection.RemoveAt(Index: Integer);
begin
  DefaultInterface.RemoveAt(Index);
end;

function TCertificateCollection.GetLastError: WideString;
begin
  Result := DefaultInterface.GetLastError;
end;

procedure Register;
begin
  RegisterComponents(dtlServerPage, [TMail, TFastSender, TCertificate, TCertificateCollection]);
end;

end.
